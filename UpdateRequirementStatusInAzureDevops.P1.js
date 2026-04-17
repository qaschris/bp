const { Webhooks } = require("@qasymphony/pulse-sdk");
const axios = require("axios");

exports.handler = async function ({ event, constants, triggers }, context, callback) {
  const emittedMessageKeys = new Set();

  function emitEvent(name, payload) {
    const trigger = triggers.find(item => item.name === name);
    return trigger
      ? new Webhooks().invoke(trigger, payload)
      : console.error(`[ERROR]: (emitEvent) Webhook named '${name}' not found.`);
  }

  function emitFriendlyFailure(details = {}) {
    const platform = details.platform || "Unknown";
    const objectType = details.objectType || "Object";
    const objectId = details.objectId != null ? details.objectId : "Unknown";
    const fieldName = details.fieldName ? ` Field: ${details.fieldName}.` : "";
    const fieldValue = details.fieldValue != null && details.fieldValue !== ""
      ? ` Value: ${details.fieldValue}.`
      : "";
    const detail = details.detail || "Sync failed.";
    const dedupKey = details.dedupKey || `failure|${platform}|${objectType}|${objectId}|${fieldName}|${fieldValue}|${detail}`;

    const message =
      `Sync failed. Platform: ${platform}. Object Type: ${objectType}. Object ID: ${objectId}.${fieldName}${fieldValue} Detail: ${detail}`;

    if (emittedMessageKeys.has(dedupKey)) return false;
    emittedMessageKeys.add(dedupKey);
    console.error(`[Error] ${message}`);
    emitEvent("ChatOpsEvent", { message });
    return true;
  }

  function normalizeText(value) {
    return value == null
      ? ""
      : String(value)
        .normalize("NFKC")
        .replace(/[\u200B-\u200D\uFEFF]/g, "")
        .replace(/<0x(?:200b|200c|200d|feff)>/gi, "")
        .trim();
  }

  function normalizeBaseUrl(value) {
    const raw = normalizeText(value).replace(/\/+$/, "");
    if (!raw) throw new Error("A qTest base URL is required.");
    return raw.startsWith("http://") || raw.startsWith("https://") ? raw : `https://${raw}`;
  }

  function validateRequiredConfiguration() {
    const missingConstants = [
      "ProjectID",
      "RequirementStatusFieldID",
      "AzDoTestingStatusFieldRef",
      "AZDO_TOKEN",
      "ManagerURL",
      "AzDoProjectURL",
      "QTEST_TOKEN",
    ].filter(name => !normalizeText(constants[name]));

    if (!missingConstants.length) return true;

    emitFriendlyFailure({
      platform: "Pulse",
      objectType: "Configuration",
      objectId: "Unknown",
      fieldName: missingConstants.join(", "),
      detail: "Required requirement status sync constants are missing in Pulse.",
      dedupKey: `requirement-status-config:${missingConstants.join("|")}`
    });
    return false;
  }

  function getLabelById(id) {
    const map = {
      11163: "SIT 1 In Progress",
      11164: "SIT 1 Complete",
      11165: "UAT In Progress",
      11166: "UAT Complete",
      11219: "SIT Dry Run In Progress",
      11220: "SIT Dry Run Complete",
      11236: "SIT 2 In Progress",
      11235: "SIT 2 Complete"
    };
    return map[id];
  }

  async function getAdoWorkItem(workItemId, token, baseUrl) {
    const url = `${baseUrl}/_apis/wit/workitems/${workItemId}?api-version=6.0`;
    const encodedToken = Buffer.from(`:${token}`).toString("base64");
    const resp = await axios.get(url, {
      headers: { Authorization: `Basic ${encodedToken}` }
    });
    return resp.data;
  }

  function getEventFieldIds(eventData) {
    const fieldIds = [];
    const pushIfPresent = value => {
      if (value !== undefined && value !== null && value !== "") {
        fieldIds.push(String(value));
      }
    };

    pushIfPresent(eventData?.field_id);
    pushIfPresent(eventData?.fieldId);
    pushIfPresent(eventData?.property?.field_id);
    pushIfPresent(eventData?.property?.fieldId);
    pushIfPresent(eventData?.change?.field_id);
    pushIfPresent(eventData?.change?.fieldId);
    pushIfPresent(eventData?.change?.property_id);

    const collections = [
      eventData?.fields,
      eventData?.properties,
      eventData?.changes,
    ];

    for (const collection of collections) {
      if (!Array.isArray(collection)) continue;

      for (const item of collection) {
        pushIfPresent(item?.field_id);
        pushIfPresent(item?.fieldId);
        pushIfPresent(item?.property_id);
      }
    }

    return [...new Set(fieldIds)];
  }

  function shouldProcessStatusEvent(eventData, statusFieldId) {
    const eventFieldIds = getEventFieldIds(eventData);
    if (!eventFieldIds.length) {
      console.log("[Info] qTest event does not expose changed field ids. Continuing with status sync.");
      return true;
    }

    console.log(`[Debug] qTest event field ids: ${JSON.stringify(eventFieldIds)}`);

    if (!eventFieldIds.includes(String(statusFieldId))) {
      console.log(`[Info] qTest event does not reference status field '${statusFieldId}'. Skipping to prevent loop.`);
      return false;
    }

    return true;
  }

  try {
    console.log(`[Info] Incoming event: ${JSON.stringify(event, null, 2)}`);
    if (!validateRequiredConfiguration()) return;

    const requirementId = event.requirement && event.requirement.id;
    const projectId = event.requirement && event.requirement.project_id;

    if (!requirementId || !projectId) {
      console.error(`[Error] Missing requirement ID or project ID.`);
      emitFriendlyFailure({
        platform: "qTest",
        objectType: "Requirement",
        objectId: requirementId || "Unknown",
        detail: "Event did not include the required requirement or project identifier."
      });
      return;
    }

    if (projectId != constants.ProjectID) {
      console.log(`[Info] Project not matching '${projectId}' != '${constants.ProjectID}', exiting.`);
      return;
    }

    console.log(`[Info] Requirement update event received for ID '${requirementId}'`);

    const statusFieldId = constants.RequirementStatusFieldID;
    const adoToken = constants.AZDO_TOKEN;
    const managerUrl = normalizeBaseUrl(constants.ManagerURL);
    const adoProjectUrl = constants.AzDoProjectURL;
    const adoTestingStatusFieldRef = constants.AzDoTestingStatusFieldRef;

    if (!shouldProcessStatusEvent(event, statusFieldId)) {
      return;
    }

    const reqUrl = `${managerUrl}/api/v3/projects/${projectId}/requirements/${requirementId}`;
    let requirement;

    try {
      const response = await axios.get(reqUrl, {
        headers: { Authorization: `bearer ${constants.QTEST_TOKEN}` }
      });
      requirement = response.data;
    } catch (error) {
      console.error(`[Error] Failed to fetch requirement ${requirementId}:`, error.response?.data || error.message);
      emitFriendlyFailure({
        platform: "qTest",
        objectType: "Requirement",
        objectId: requirementId,
        detail: "Unable to read requirement details from qTest."
      });
      return;
    }

    if (constants.SyncUserRegex) {
      const updaterName =
        requirement?.updated_by?.name ||
        requirement?.last_modified_by?.name ||
        requirement?.modified_by?.name ||
        "";

      if (new RegExp(constants.SyncUserRegex, "i").test(updaterName)) {
        console.log(`[Info] Update appears from sync user '${updaterName}'; skipping to avoid loop.`);
        return;
      }
    }

    const props = Array.isArray(requirement.properties) ? requirement.properties : [];
    const statusProp = props.find((p) => p.field_id == statusFieldId);

    if (!statusProp) {
      console.error(`[Error] Could not find status field (ID ${statusFieldId}) on requirement.`);
      emitFriendlyFailure({
        platform: "qTest",
        objectType: "Requirement",
        objectId: requirementId,
        fieldName: "Status",
        detail: "Required status field was not found on the qTest requirement."
      });
      return;
    }

    const qtestStatusId = statusProp.field_value;
    const qtestStatusLabel = statusProp.field_value_name || statusProp.field_label || getLabelById(qtestStatusId);

    console.log(`[Info] qTest Status ID: ${qtestStatusId}, Label: '${qtestStatusLabel || "Unknown"}'`);

    const adoStatus = normalizeText(qtestStatusLabel || getLabelById(qtestStatusId));
    if (!adoStatus) {
      console.log(`[Info] No ADO mapping found for qTest status ID '${qtestStatusId}', skipping update.`);
      return;
    }

    if (adoStatus.toLowerCase() === "new") {
      console.log(`[Info] qTest status '${qtestStatusLabel || qtestStatusId}' maps to ADO status 'New'. Skipping ADO update by design.`);
      return;
    }

    const workItemName = requirement.name || "";
    const match = workItemName.match(/^WI(\d+):/);

    if (!match) {
      console.error(`[Error] Work item ID not found in requirement name '${workItemName}'.`);
      emitFriendlyFailure({
        platform: "ADO",
        objectType: "Requirement",
        objectId: requirementId,
        fieldName: "Work Item ID",
        detail: "Unable to determine the linked Azure DevOps work item from the qTest requirement name."
      });
      return;
    }

    const workItemId = match[1];
    console.log(`[Info] Mapped ADO Work Item ID: '${workItemId}'`);

    let adoCurrent;
    try {
      adoCurrent = await getAdoWorkItem(workItemId, adoToken, adoProjectUrl);
    } catch (error) {
      console.error(`[Error] Failed to read ADO Work Item '${workItemId}':`, error.response?.data || error.message);
      emitFriendlyFailure({
        platform: "ADO",
        objectType: "Requirement",
        objectId: workItemId,
        detail: "Unable to read the Azure DevOps work item."
      });
      return;
    }

    const curAdoStatus = adoCurrent?.fields?.[adoTestingStatusFieldRef] || "";
    console.log(`[Info] Current ADO Testing Status: '${curAdoStatus || "(empty)"}'`);
    console.log(`[Info] Desired ADO Testing Status: '${adoStatus}'`);

    if (String(curAdoStatus).trim() === String(adoStatus).trim()) {
      console.log(`[Info] ADO Testing Status already matches qTest status. Skipping patch to avoid loop.`);
      return;
    }

    try {
      const adoUrl = `${adoProjectUrl}/_apis/wit/workitems/${workItemId}?api-version=6.0`;
      const requestBody = [
        {
          op: "add",
          path: `/fields/${adoTestingStatusFieldRef}`,
          value: adoStatus
        }
      ];

      console.log(`[Info] Sending ADO Testing Status patch: ${JSON.stringify(requestBody, null, 2)}`);

      await axios.patch(adoUrl, requestBody, {
        headers: {
          "Content-Type": "application/json-patch+json",
          Authorization: `basic ${Buffer.from(`:${adoToken}`).toString("base64")}`
        }
      });

      console.log(
        `[Info] Successfully updated ADO Work Item '${workItemId}' to Testing Status '${adoStatus}' ` +
        `(from qTest: '${qtestStatusLabel || qtestStatusId}')`
      );
    } catch (error) {
      console.error(`[Error] Failed to update ADO Work Item '${workItemId}':`, error.response?.data || error.message);
      emitFriendlyFailure({
        platform: "ADO",
        objectType: "Requirement",
        objectId: workItemId,
        fieldName: "Testing Status",
        fieldValue: adoStatus,
        detail: "Unable to update Azure DevOps testing status from qTest.",
        dedupKey: `requirement-status-patch:${workItemId}:${adoStatus}`
      });
    }

  } catch (fatal) {
    console.error(`[Fatal] Unexpected error:`, fatal.response?.data || fatal.message);
    emitFriendlyFailure({
      platform: "ADO",
      objectType: "Requirement",
      objectId: event?.requirement?.id || "Unknown",
      detail: "Unexpected error occurred during requirement status sync."
    });
  }
};

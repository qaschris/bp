const { Webhooks } = require('@qasymphony/pulse-sdk');
const axios = require('axios');

exports.handler = async function ({ event, constants, triggers }, context, callback) {
  function emitEvent(name, payload) {
    return (t = triggers.find(t => t.name === name))
      ? new Webhooks().invoke(t, payload)
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

    const message =
      `Sync failed. Platform: ${platform}. Object Type: ${objectType}. Object ID: ${objectId}.${fieldName}${fieldValue} Detail: ${detail}`;

    console.error(`[Error] ${message}`);
    emitEvent('ChatOpsEvent', { message });
  }

  function normalizeAreaPathLabel(value) {
    return typeof value === "string" ? value.trim() : "";
  }

  function extractAdoIdentityEmail(v) {
    if (!v) return "";
    if (typeof v === "string") return v.trim();

    if (typeof v === "object") {
      return (v.uniqueName || v.mail || v.email || v.displayName || "").toString().trim();
    }
    return "";
  }

  async function getAdoWorkItem(workItemId, token, baseUrl) {
    const url = `${baseUrl}/_apis/wit/workitems/${workItemId}?api-version=6.0&$expand=Relations`;
    const encodedToken = Buffer.from(`:${token}`).toString("base64");
    const resp = await axios.get(url, {
      headers: { Authorization: `Basic ${encodedToken}` }
    });
    return resp.data;
  }

  try {
    console.log("[Info] Defect update event received.");

    const iteration = event.iteration !== undefined ? event.iteration : 1;
    console.log("[Info] Iteration:", iteration);

    const DEFAULT_AREA_PATH = constants.AreaPath;
    const DEFAULT_ASSIGNED_TO_TEAM_VALUE = 1189;

    const AREA_PATH_TO_QTEST_TEAM_VALUE = {
      "bp_Quantum\\Process\\Asset Management\\Squad 1 - Data": 1,
      "bp_Quantum\\Process\\Asset Management\\AM - Work Mgmt Core": 2,
      "bp_Quantum\\Process\\Finance\\Finance - Capex Squad": 3,
      "bp_Quantum\\Process\\Finance\\Finance - CB Hub 2.0 Squad": 4,
      "bp_Quantum\\Process\\Finance\\Finance - R2R Squad": 6,
      "bp_Quantum\\Process\\Finance\\Finance - Tax Squad": 7,
      "bp_Quantum\\Process\\Procurement\\Invoice to Pay": 8,
      "bp_Quantum\\Process\\Procurement\\Quality and Logistics": 9,
      "bp_Quantum\\Process\\Procurement\\Services Procurement": 10,
      "bp_Quantum\\Process\\Procurement\\Materials Procurement": 11,
      "bp_Quantum\\Process\\Procurement\\Source to Contract": 12,
      "bp_Quantum\\Process\\Procurement\\Materials and Inventory": 13,
      "bp_Quantum\\Process\\MDG\\Asset Management": 14,
      "bp_Quantum\\Process\\MDG\\Procurement": 15,
      "bp_Quantum\\Process\\MDG\\Finance": 16,
      "bp_Quantum\\Process\\MDG\\Material": 17,
      "bp_Quantum\\Data and Analytics\\ETL\\Asset Management": 18,
      "bp_Quantum\\Data and Analytics\\ETL\\Customer Management": 19,
      "bp_Quantum\\Data and Analytics\\ETL\\Finance": 20,
      "bp_Quantum\\Data and Analytics\\ETL\\Material Master": 1320,
      "bp_Quantum\\Data and Analytics\\ETL\\Procurement": 21,
      "bp_Quantum\\Data and Analytics\\Reporting and Analytics\\Asset Management (R and A)": 89,
      "bp_Quantum\\Data and Analytics\\Reporting and Analytics\\Finance (R and A)": 90,
      "bp_Quantum\\Technical\\Dev and Integration\\Asset Management": 1185,
      "bp_Quantum\\Data and Analytics\\Reporting and Analytics\\Procurement (R and A)": 1184,
      "bp_Quantum\\Technical\\Dev and Integration\\Finance": 1186,
      "bp_Quantum\\Technical\\Dev and Integration\\Integration": 1187,
      "bp_Quantum\\Technical\\Dev and Integration\\Procurement": 1188,
      "bp_Quantum\\Technical\\Testing": 1189,
      "bp_Quantum\\Technical\\Cutover and Release Management": 1193,
      "bp_Quantum\\Technical\\Architecture": 1195,
      "bp_Quantum\\Technical\\Digital Security and Compliance\\Asset Management": 1194,
      "bp_Quantum\\Technical\\Digital Security and Compliance\\Finance": 1211,
      "bp_Quantum\\Technical\\Digital Security and Compliance\\Procurement": 1212,
      "bp_Quantum\\Technical\\Digital Security and Compliance\\MDG": 1213,
      "bp_Quantum\\Technical\\Digital Security and Compliance\\Cross - Entity": 1214,
      "bp_Quantum\\Technical\\Identity and Access Management\\Asset Management": 1215,
      "bp_Quantum\\Technical\\Identity and Access Management\\Finance": 1216,
      "bp_Quantum\\Technical\\Identity and Access Management\\Procurement": 1217,
      "bp_Quantum\\Technical\\Identity and Access Management\\MDG": 1218,
      "bp_Quantum\\Technical\\Identity and Access Management\\Cross - Entity": 1219,
      "bp_Quantum\\Technical\\Platforms\\BW4": 1220,
      "bp_Quantum\\Technical\\Platforms\\C and P": 1221,
      "bp_Quantum\\Technical\\Platforms\\CB": 1222,
      "bp_Quantum\\Technical\\Platforms\\CFIN or Core": 1223,
      "bp_Quantum\\Technical\\Platforms\\MDG": 1224,
      "bp_Quantum\\Technical\\Platforms\\P and O": 1225,
      "bp_Quantum\\Process\\Finance\\Finance - ARAPINTERCO Squad": 1314,
      "bp_Quantum\\Data and Analytics\\ETL\\Site Castellon": 1332,
      "bp_Quantum\\Data and Analytics\\ETL\\Site Kwinana": 1334,
      "bp_Quantum\\Data and Analytics\\ETL\\Site Whiting": 1335,
      "bp_Quantum\\Data and Analytics\\ETL\\Site Global": 1348,
      "bp_Quantum\\Data and Analytics\\ETL\\Data office": 1349,
      "bp_Quantum\\Technical\\Platforms\\Shared\\BTP or SaaS": 1354,
      "bp_Quantum\\Technical\\Platforms\\Shared\\GRC": 1355,
      "bp_Quantum\\Technical\\Platforms\\Shared\\OpenText": 1356,
      "bp_Quantum\\Technical\\Platforms\\Shared\\TL or Architecture or GRC": 1357
    };

    const QTEST_TEAM_VALUE_TO_AREA_PATH = Object.fromEntries(
      Object.entries(AREA_PATH_TO_QTEST_TEAM_VALUE).map(([label, value]) => [String(value), label])
    );

    function mapQtestTeamValueToAreaPath(valueId) {
      return QTEST_TEAM_VALUE_TO_AREA_PATH[String(valueId)] || DEFAULT_AREA_PATH;
    }

    const defectId = event.defect?.id || event.entityId;
    const projectId = event.defect?.project_id || event.projectId;

    if (!defectId || !projectId) {
      console.error("[Error] Missing defect or project ID in event.");
      emitFriendlyFailure({
        platform: "qTest",
        objectType: "Defect",
        objectId: defectId || "Unknown",
        detail: "Event did not include the required defect or project identifier."
      });
      return;
    }

    if (projectId != constants.ProjectID) {
      console.log(`[Info] Project not matching '${projectId}' != '${constants.ProjectID}', exiting.`);
      return;
    }

    const qTestDefectUrl = `https://${constants.ManagerURL}/api/v3/projects/${projectId}/defects/${defectId}`;
    console.log("[Info] Fetching defect details:", qTestDefectUrl);

    let defect;
    try {
      const qTestResponse = await axios.get(qTestDefectUrl, {
        headers: { Authorization: `Bearer ${constants.QTEST_TOKEN}` }
      });
      defect = qTestResponse.data;
    } catch (err) {
      console.error("[Error] Failed to fetch defect:", err.message);
      emitFriendlyFailure({
        platform: "qTest",
        objectType: "Defect",
        objectId: defectId,
        detail: "Unable to read defect details from qTest."
      });
      return;
    }

    if (constants.SyncUserRegex) {
      const updaterName = defect?.updated_by?.name || defect?.last_modified_by?.name || "";
      if (new RegExp(constants.SyncUserRegex, "i").test(updaterName)) {
        console.log("[Info] Update appears from ADO sync user; skipping to avoid loop.");
        return;
      }
    }

    const props = Array.isArray(defect.properties) ? defect.properties : [];
    const getPropById = (fid) => props.find(p => p.field_id == fid);
    const firstNonEmpty = (...vals) => vals.find(v => v && String(v).trim().length) || "";
    const norm = (s) => (typeof s === "string" ? s.trim() : s);

    const summary = firstNonEmpty(getPropById(constants.DefectSummaryFieldID)?.field_value, defect.name);
    const description = firstNonEmpty(getPropById(constants.DefectDescriptionFieldID)?.field_value, defect.description);

    const wiRegex = /WI[-\s:]?(\d+)/i;
    let wiMatch = wiRegex.exec(summary) || wiRegex.exec(description || "");

    if (!wiMatch && props.length) {
      for (const p of props) {
        const v = firstNonEmpty(p.field_value, p.field_value_name);
        const m = v ? wiRegex.exec(v) : null;
        if (m) {
          wiMatch = m;
          break;
        }
      }
    }

    if (!wiMatch) {
      console.error("[Error] Could not extract Azure Work Item ID.");
      emitFriendlyFailure({
        platform: "ADO",
        objectType: "Defect",
        objectId: defectId,
        fieldName: "Work Item ID",
        detail: "Unable to determine the linked Azure DevOps work item from the qTest defect."
      });
      return;
    }

    const workItemId = wiMatch[1];
    console.log("[Info] Found Azure Work Item ID:", workItemId);

    const applicationProp = getPropById(constants.DefectApplicationFieldID);
    const sourceTeamProp = getPropById(constants.DefectSourceTeamFieldID);
    const siteNameProp = getPropById(constants.DefectSiteNameFieldID);
    const assignedToProp = getPropById(constants.DefectAssignedToFieldID);
    const assignedToTeamProp = getPropById(constants.DefectAssignedToTeamFieldID);
    const targetDateProp = getPropById(constants.DefectTargetDateFieldID);
    const targetDate = norm(firstNonEmpty(targetDateProp?.field_value));

    const appLabel = norm(firstNonEmpty(applicationProp?.field_value_name));
    const srcLabel = norm(firstNonEmpty(sourceTeamProp?.field_value_name));
    const siteLabel = norm(firstNonEmpty(siteNameProp?.field_value_name));
    const assignedLabel = norm(firstNonEmpty(assignedToProp?.field_value));
    const assignedToLabel = norm(firstNonEmpty(assignedToProp?.field_value_name));

    let userEmail = "";
    let userName = "";

    if (assignedLabel) {
      const userApiUrl = `https://${constants.ManagerURL}/api/v3/users/${assignedLabel}`;
      console.log("[Info] Fetching qTest user details:", userApiUrl);

      try {
        const userResp = await axios.get(userApiUrl, {
          headers: { Authorization: `Bearer ${constants.QTEST_TOKEN}` }
        });

        const u = userResp?.data || {};
        const identity = (u.username || "").trim() || (u.ldap_username || "").trim() || (u.external_user_name || "").trim();
        userEmail = (u.email || "").trim();
        userName = identity;
      } catch (e) {
        console.error("[Error] Failed to fetch qTest user details:", e.response?.data || e.message);
        // Log-only: do not emit ChatOps for non-fatal assigned user lookup issues
      }
    }

    const assignedToTeamLabel = norm(
      firstNonEmpty(
        assignedToTeamProp?.field_value_name,
        mapQtestTeamValueToAreaPath(assignedToTeamProp?.field_value)
      )
    );

    let isoDate;
    if (targetDate) {
      const d = new Date(targetDate);
      if (!isNaN(d.getTime())) isoDate = d.toISOString().replace(".000Z", "+00:00");
    }

    console.log("[Info] Assigned To Username:", userName);
    console.log("[Info] Assigned To Email:", userEmail);
    console.log("[Info] Assigned To Label:", assignedToLabel);

    let adoCurrent;
    try {
      adoCurrent = await getAdoWorkItem(workItemId, constants.AZDO_TOKEN, constants.AzDoProjectURL);
    } catch (e) {
      console.error("[Error] Failed to read ADO work item:", e.response?.data || e.message);
      emitFriendlyFailure({
        platform: "ADO",
        objectType: "Defect",
        objectId: workItemId,
        detail: "Unable to read the Azure DevOps work item."
      });
      return;
    }

    const cur = adoCurrent?.fields || {};

    const curApp = norm(constants.AzDoApplicationFieldRef ? cur[constants.AzDoApplicationFieldRef] : "");
    const curSrc = norm(constants.AzDoSourceTeamFieldRef ? cur[constants.AzDoSourceTeamFieldRef] : "");
    const curSite = norm(constants.AzDoSiteNameFieldRef ? cur[constants.AzDoSiteNameFieldRef] : "");
    const curAreaPath = norm(cur[constants.AzDoAreaPathFieldRef || "System.AreaPath"]);
    const curAssignedToRaw = cur[constants.AzDoAssignedToFieldRef];
    const curAssignedTo = norm(extractAdoIdentityEmail(curAssignedToRaw));
    const curTargetDateRaw = cur[constants.AzDoTargetDateFieldRef];

    const normalizeDate = (d) => {
      if (!d) return "";
      const dt = new Date(d);
      return isNaN(dt.getTime()) ? "" : dt.toISOString().split("T")[0];
    };

    const curTargetDate = normalizeDate(curTargetDateRaw);
    const newTargetDate = normalizeDate(isoDate);

    console.log("curTargetDate", curTargetDate);
    console.log("newTargetDate", newTargetDate, isoDate);

    const patchData = [];

    if (constants.AzDoApplicationFieldRef && appLabel && curApp !== appLabel) {
      console.log("[Info] Updating Application:", { from: curApp || "(empty)", to: appLabel });
      patchData.push({ op: "add", path: `/fields/${constants.AzDoApplicationFieldRef}`, value: appLabel });
    }

    if (constants.AzDoSourceTeamFieldRef && srcLabel && curSrc !== srcLabel) {
      console.log("[Info] Updating SourceTeam:", { from: curSrc || "(empty)", to: srcLabel });
      patchData.push({ op: "add", path: `/fields/${constants.AzDoSourceTeamFieldRef}`, value: srcLabel });
    }

    if (constants.AzDoSiteNameFieldRef && siteLabel && curSite !== siteLabel) {
      console.log("[Info] Updating SubEntity:", { from: curSite || "(empty)", to: siteLabel });
      patchData.push({ op: "add", path: `/fields/${constants.AzDoSiteNameFieldRef}`, value: siteLabel });
    }

    if (assignedToTeamLabel && curAreaPath !== assignedToTeamLabel) {
      console.log("[Info] Updating AreaPath from qTest Assigned to Team:", {
        from: curAreaPath || "(empty)",
        to: assignedToTeamLabel
      });
      patchData.push({
        op: "add",
        path: `/fields/${constants.AzDoAreaPathFieldRef || "System.AreaPath"}`,
        value: assignedToTeamLabel
      });
    }

    if (constants.AzDoAssignedToFieldRef) {
      const desiredAssignedTo = norm(userEmail || userName);

      if (desiredAssignedTo) {
        if (curAssignedTo !== desiredAssignedTo) {
          console.log("[Info] Updating AssignedTo:", { from: curAssignedTo || "(unassigned)", to: desiredAssignedTo });
          patchData.push({ op: "add", path: `/fields/${constants.AzDoAssignedToFieldRef}`, value: desiredAssignedTo });
        }
      } else if (curAssignedTo) {
        console.log("[Info] Clearing AssignedTo (ADO Unassigned):", { from: curAssignedTo });
        patchData.push({ op: "remove", path: `/fields/${constants.AzDoAssignedToFieldRef}` });
      }
    }

    if (constants.AzDoTargetDateFieldRef && isoDate && curTargetDate !== newTargetDate) {
      console.log("[Info] Updating TargetDate:", { from: curTargetDate || "(empty)", to: isoDate });
      patchData.push({ op: "add", path: `/fields/${constants.AzDoTargetDateFieldRef}`, value: isoDate });
    }

    const backlink = defect.web_url || qTestDefectUrl;
    const hasLink = (adoCurrent?.relations || []).some(
      r => r.rel === "Hyperlink" && (r.url || "").toLowerCase() === (backlink || "").toLowerCase()
    );
    if (backlink && !hasLink) {
      patchData.push({ op: "add", path: "/relations/-", value: { rel: "Hyperlink", url: backlink } });
    }

    if (patchData.length === 0) {
      console.log("[Info] No ADO changes detected; skipping patch (prevents loops).");
      return;
    }

    const adoPatchUrl = `${constants.AzDoProjectURL}/_apis/wit/workitems/${workItemId}?api-version=6.0`;
    const encodedToken = Buffer.from(`:${constants.AZDO_TOKEN}`).toString("base64");

    console.log("[Info] Sending update to ADO:", JSON.stringify(patchData, null, 2));

    try {
      await axios.patch(adoPatchUrl, patchData, {
        headers: {
          Authorization: `Basic ${encodedToken}`,
          "Content-Type": "application/json-patch+json"
        }
      });
      console.log("[Info] Successfully updated Azure DevOps work item.");
    } catch (err) {
      console.error("[Error] Azure update failed:", err.response?.data || err.message);
      emitFriendlyFailure({
        platform: "ADO",
        objectType: "Defect",
        objectId: workItemId,
        detail: "Unable to update the Azure DevOps work item from qTest."
      });
    }

  } catch (fatal) {
    console.error("[Fatal] Unexpected error:", fatal.message);
    emitFriendlyFailure({
      platform: "ADO",
      objectType: "Defect",
      objectId: event?.defect?.id || event?.entityId || "Unknown",
      detail: "Unexpected error occurred during defect sync."
    });
  }
};
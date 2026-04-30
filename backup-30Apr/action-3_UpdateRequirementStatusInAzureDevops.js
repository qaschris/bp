const { Webhooks } = require('@qasymphony/pulse-sdk');
const axios = require("axios");

exports.handler = async function ({ event, constants, triggers }, context, callback) {
  console.log(`[Info] Incoming event: ${JSON.stringify(event, null, 2)}`);

  const requirementId = event.requirement && event.requirement.id;
  const projectId = event.requirement && event.requirement.project_id;

  if (!requirementId || !projectId) {
    console.error(`[Error] Missing requirement ID or project ID.`);
    return;
  }

  console.log(`[Info] Requirement update event received for ID '${requirementId}'`);

  const statusFieldId = constants.RequirementStatusFieldID; //1455
  const adoToken = constants.AZDO_TOKEN;
  const managerUrl = constants.ManagerURL;
  const adoProjectUrl = constants.AzDoProjectURL;

  // qTest Status ID ➜ ADO State mapping
  const statusMap = {
    10903: "SIT 1 In Progress",   // SIT 1 In Progress
    10904: "SIT 1 Complete",     // SIT 1 Complete
    914: "UAT In Progress",     // UAT In Progress
    10897: "UAT Complete",       // UAT Complete
    912: "SIT Dry Run In Progress", //SIT Dry Run In Progress
    913: "SIT Dry Run Complete", //SIT Dry Run Complete
    10901: "SIT 2 In Progress",    // SIT 2 In Progress
    10902: "SIT 2 Complete"        // SIT 2 Complete
  };

  // Get qTest requirement details to get Status + Name
  const reqUrl = `https://${managerUrl}/api/v3/projects/${projectId}/requirements/${requirementId}`;
  let requirement;
  try {
    const response = await axios.get(reqUrl, {
      headers: {
        Authorization: `bearer ${constants.QTEST_TOKEN}`
      }
    });
    requirement = response.data;
  } catch (error) {
    console.error(`[Error] Failed to fetch requirement ${requirementId}: ${error}`);
    return;
  }

  // Extract qTest status
  const statusProp = requirement.properties.find((p) => p.field_id == statusFieldId);
  if (!statusProp) {
    console.error(`[Error] Could not find status field on requirement.`);
    return;
  }

  const qtestStatusId = statusProp.field_value;
  const qtestStatusLabel = statusProp.field_label || getLabelById(qtestStatusId);
  console.log(`[Info] qTest Status ID: ${qtestStatusId}, Label: '${qtestStatusLabel || "Unknown"}'`);

  const adoStatus = statusMap[qtestStatusId];
  if (!adoStatus) {
    console.log(`[Info] No ADO mapping found for qTest status ID '${qtestStatusId}', skipping update.`);
    return;
  }

  // Extract Work Item ID from name (format: WI<id>: Title)
  const workItemName = requirement.name;
  const match = workItemName.match(/^WI(\d+):/);
  if (!match) {
    console.error(`[Error] Work item ID not found in requirement name.`);
    return;
  }

  const workItemId = match[1];
  console.log(`[Info] Mapped ADO Work Item ID: '${workItemId}'`);

  // Update ADO work item status using PATCH
  try {
    const adoUrl = `${adoProjectUrl}/_apis/wit/workitems/${workItemId}?api-version=6.0`;
    const requestBody = [
      {
        op: "add",
        path: "/fields/Custom.TestingStatus", // <== your ADO field for testing status
        value: adoStatus
      }
    ];

    await axios.patch(adoUrl, requestBody, {
      headers: {
        "Content-Type": "application/json-patch+json",
        Authorization: `basic ${Buffer.from(`:${adoToken}`).toString("base64")}`
      }
    });

    console.log(`[Info] Successfully updated ADO Work Item '${workItemId}' to state '${adoStatus}'`);
  } catch (error) {
    console.error(`[Error] Failed to update ADO Work Item '${workItemId}': ${error}`);
    if (error.response) {
      console.error(`[Error] Status: ${error.response.status}`);
      console.error(`[Error] Data: ${JSON.stringify(error.response.data, null, 2)}`);
    }
  }

    // Helper: manually resolve label for known IDs
  function getLabelById(id) {
    const map = {
      10903: "SIT 1 In Progress",   // SIT 1 In Progress
      10904: "SIT 1 Complete",     // SIT 1 Complete
      914: "UAT In Progress",     // UAT In Progress
      10897: "UAT Complete",       // UAT Complete
      912: "SIT Dry Run In Progress", //SIT Dry Run In Progress
      913: "SIT Dry Run Complete", //SIT Dry Run Complete
      10901: "SIT 2 In Progress",    // SIT 2 In Progress
      10902: "SIT 2 Complete"        // SIT 2 Complete
    };
    return map[id];
  }
};

const { Webhooks } = require('@qasymphony/pulse-sdk');
const axios = require('axios');

exports.handler = async function ({ event, constants, triggers }, context, callback) {
    let iteration;
    if (event.iteration != undefined) {
        iteration = event.iteration;
    } else {
        iteration = 1;
    }
    const maxIterations = 21;
    const defectId = event.defect.id;
    const projectId = event.defect.project_id;
    console.log(`[Info] Create defect event received for defect '${defectId}' in project '${projectId}'`);

    if (projectId != constants.ProjectID) {
        console.log(`[Info] Project not matching '${projectId}' != '${constants.ProjectID}', exiting.`);
        return;
    }

    const defectDetails = await getDefectDetailsByIdWithRetry(defectId);
    if (!defectDetails) return;

    const bug = await createAzDoBug(
        defectId,
        defectDetails.summary,
        defectDetails.description,
        defectDetails.link,
        defectDetails.severity,
        defectDetails.priority,
        defectDetails.defectType,
        defectDetails.status,
        defectDetails.affectedRelease,
        defectDetails.createdBy,
        defectDetails.externalReference
    );

    if (!bug) return;

    const workItemId = bug.id;
    const newSummary = `${getNamePrefix(workItemId)}${defectDetails.summary}`;
    console.log(`[Info] New defect name: ${newSummary}`);
    await updateDefectSummary(defectId, constants.DefectSummaryFieldID, newSummary);

    function emitEvent(name, payload) {
        let t = triggers.find(t => t.name === name);
        return t && new Webhooks().invoke(t, payload);
    }

    function getNamePrefix(workItemId) {
        return `WI${workItemId}: `;
    }

    function getFieldById(obj, fieldId) {
        if (!obj || !obj.properties) {
            console.log(`[Warn] Obj/properties not found.`);
            return;
        }
        const prop = obj.properties.find((p) => p.field_id == fieldId);
        if (!prop) {
            console.log(`[Warn] Property with field id '${fieldId}' not found.`);
            return;
        }

        return prop;
    }

    async function getDefectDetailsByIdWithRetry(defectId) {
        let defectDetails = undefined;
        let delay = 5000;
        let attempt = 0;

        do {
            if (attempt > 0) {
                console.log(
                    `[Warn] Could not get defect details on attempt ${attempt}. Waiting ${delay} ms.`
                );
                await new Promise((r) => setTimeout(r, delay));
            }

            defectDetails = await getDefectDetailsById(defectId);

            // Return immediately once mandatory fields required by ADO are available
            if (
                defectDetails &&
                defectDetails.summary &&
                defectDetails.description &&
                defectDetails.severity
            ) {
                console.log(
                    `[Info] Successfully fetched complete defect details for '${defectId}' on attempt ${attempt + 1}.`
                );
                return defectDetails;
            }

            attempt++;
        } while (attempt < 12);

        console.log(
            `[Error] Could not get defect details. User may not have completed the initial save in qTest or the defect was abandoned.`
        );

        if (iteration < maxIterations) {
            iteration = iteration + 1;
            console.log(
                `[Info] Re-executing rule. Iteration ${iteration} of ${maxIterations}.`
            );
            event.iteration = iteration;
            emitEvent('qTestDefectSubmitted', event);
        } else {
            console.error(
                `[Error] Retry exceeded ${maxIterations} iterations. Rule execution timed out.`
            );
        }
    }

    async function getDefectDetailsById(defectId) {
        const defect = await getDefectById(defectId);

        if (!defect) return;

        const summaryField = getFieldById(defect, constants.DefectSummaryFieldID);
        const descriptionField = getFieldById(defect, constants.DefectDescriptionFieldID);
        const severityField = getFieldById(defect, constants.DefectSeverityFieldID); // Mapping severity
        const priorityField = getFieldById(defect, constants.DefectPriorityFieldID); // Mapping Priority
        const defectTypeField = getFieldById(defect, constants.DefectTypeFieldID); // Mapping Defect Type
        const statusField = getFieldById(defect, constants.DefectStatusFieldID); //Mapping Status
        const affectedReleaseField = getFieldById(defect, constants.DefectAffectedReleaseFieldID); // Mapping Affected Release/TestPhase
        const createdByField = getFieldById(defect, constants.DefectCreatedByFieldID); // Mapping Created By
        const externalReferenceField = getFieldById(defect, constants.DefectExternalReferenceFieldID); //External Reference

        if (!summaryField || !descriptionField || !severityField) {
            console.log("[Error] Fields not found, exiting.");
            return; // Prevents using undefined values
        }

        const summary = summaryField.field_value;
        console.log(`[Info] Defect summary: ${summary}`);
        const description = descriptionField.field_value;
        console.log(`[Info] Defect description: ${description}`);
        const link = defect.web_url;
        console.log(`[Info] Defect link: ${link}`);
        const severity = severityField.field_value;
        console.log(`[Info] Defect severity: ${severity}`);
        const priority = priorityField ? priorityField.field_value : null;
        console.log(`[Info] Defect priority: ${priority}`);
        const defectType = defectTypeField ? defectTypeField.field_value : null;
        console.log(`[Info] Defect type: ${defectType}`);
        const status = statusField ? statusField.field_value : null;
        console.log(`[Info] Defect status: ${status}`);
        const affectedRelease = affectedReleaseField ? affectedReleaseField.field_value : null;
        console.log(`[Info] Defect Affected Release/TestPhase: ${affectedRelease}`);
        let createdBy = createdByField ? createdByField.field_value : null;
        if (createdBy) {
            createdBy = await getQtestUserName(createdBy);  //Resolve numeric ID → readable name
        }
        console.log(`[Info] Defect Created By: ${createdBy}`);
        const externalReference = externalReferenceField ? externalReferenceField.field_value : null;
        console.log(`[Info] Defect External Reference: ${externalReference}`);

        return { summary, description, link, severity, priority, defectType, status, affectedRelease, createdBy, externalReference };
    }

    async function getDefectById(defectId) {
        const defectUrl = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/defects/${defectId}`;

        console.log(`[Info] Get defect details for '${defectId}'`);

        try {
            const response = await axios.get(defectUrl, {
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `bearer ${constants.QTEST_TOKEN}`
                }
            });
            return response.data;
        } catch (error) {
            if (error.response && error.response.status === 404) {
                console.log(`[Info] qTest returned 404 — defect '${defectId}' not found (abandoned).`);
                return null; // Important: return null, not undefined
            }
            console.log("[Error] Failed to get defect by id.", error);
            return null;
        }
    }

    function mapSeverity(qtestSeverity) {
        const severityId = parseInt(qtestSeverity);  // force numeric comparison
        switch (severityId) {
            case 10899:
                return '1 - Critical';
            case 10302:
                return '2 - High';
            case 10303:
                return '3 - Medium';
            case 10304:
                return '4 - Low';
            default:
                return '3 - Medium';
        }
    }

    // Priority Mapping (qTest → ADO)
    function mapPriority(qtestPriority) {
        const priorityId = parseInt(qtestPriority);
        switch (priorityId) {
            case 10898: return 1; // Very High
            case 10204: return 2; // High
            case 10203: return 3; // Medium
            case 10202: return 4; // Low
            default: return 3;    // Default Medium
        }
    }

    // Defect Type Mapping ((qTest → ADO))
    function mapDefectType(qtestDefectType) {
        const id = parseInt(qtestDefectType);
        switch (id) {
            case 1751: return "New_Requirement";
            case 1752: return "Code";
            case 1753: return "Data";
            case 1754: return "Environment";
            case 1792: return "Infrastructure";
            case 1755: return "User Authorization";
            case 1756: return "Configuration";
            case 1757: return "User Handling";
            case 1758: return "Translation";
            case 1759: return "Automation";
            default: return null; // null means skip this field
        }
    }

    // qTest → ADO Status Mapping
    function mapStatus(qtestStatus) {
        const statusId = parseInt(qtestStatus);
        switch (statusId) {
            case 10001: return "New";
            case 10003: return "In Analysis";
            case 11121: return "Triage";
            case 10004: return "In Resolution";
            case 10005: return "Awaiting Implementation";
            case 10006: return "Resolved";
            case 10850: return "Retest";
            case 10852: return "Reopened";
            case 10851: return "Closed";
            case 10002: return "On Hold";
            case 10853: return "Rejected";
            default: return "New"; // null means skip this field
        }
    }

    // qTest → ADO Bug Stage Mapping
    function mapAffectedRelease(qtestRelease) {
        const releaseId = parseInt(qtestRelease);
        switch (releaseId) {
            case -511: return null; // Skip "P&O Release 1"
            case 350: return "P&O_R1_SIT Dry Run";
            case 310: return "P&O_R1_SIT1";
            case 311: return "P&O_R1_SIT2";
            case 312: return "P&O_R1_DC1";
            case 347: return "P&O_R1_DC2";
            case 348: return "P&O_R1_DC3";
            case 351: return "P&O_R1_UAT";
            case 393: return "Unit Testing";
            default: return null; // Unknown or not set
        }
    }

    // Helper: Fetch qTest user display name by user ID
    async function getQtestUserName(userId) {
        if (!userId) return null;

        const userUrl = `https://${constants.ManagerURL}/api/v3/users/${userId}`;
        try {
            const response = await axios.get(userUrl, {
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `bearer ${constants.QTEST_TOKEN}`
                }
            });

            const userData = response.data;

            // Prefer formatted readable name (first_name + last_name)
            let displayName = "";
            if (userData.first_name && userData.last_name) {
                // Preserve comma pattern exactly as stored in qTest
                displayName = `${userData.first_name.trim()} ${userData.last_name.trim()}`;
            } else if (userData.last_name) {
                displayName = userData.last_name.trim();
            } else {
                displayName = userData.username || userData.email || userId.toString();
            }

            console.log(`[Info] Resolved qTest user ID '${userId}' to name '${displayName}'`);
            return displayName;

        } catch (error) {
            console.log(`[Warn] Could not resolve qTest user ID '${userId}' to name. Using ID instead.`);
            return userId.toString();
        }
    }
    async function createAzDoBug(
        defectId,
        name,
        description,
        link,
        qtestSeverity,
        qtestPriority,
        qtestDefectType,
        qtestStatus,
        qtestAffectedRelease,
        qtestCreatedBy,
        qtestExternalReference
    ) {
        console.log(`[Info] Creating bug in Azure DevOps '${defectId}'`);
        const baseUrl = encodeIfNeeded(constants.AzDoProjectURL);
        const url = `${baseUrl}/_apis/wit/workitems/$Bug?api-version=6.0`;

        const mappedStatus = mapStatus(qtestStatus);
        console.log(`[Info] Mapped qTest Status '${qtestStatus}' to ADO State '${mappedStatus}'`);

        const mappedSeverity = mapSeverity(qtestSeverity);
        console.log(`[Info] Mapped severity: ${mappedSeverity}`);

        const mappedPriority = mapPriority(qtestPriority);
        console.log(`[Info] Mapped Priority: ${mappedPriority}`);

        const mappedDefectType = mapDefectType(qtestDefectType);
        console.log(`[Info] Mapped Defect Type: ${mappedDefectType}`);

        const mappedAffectedRelease = mapAffectedRelease(qtestAffectedRelease);
        if (mappedAffectedRelease) {
            console.log(`[Info] Mapped qTest Affected Release '${qtestAffectedRelease}' to ADO Bug Stage '${mappedAffectedRelease}'`);
        } else {
            console.log(`[Info] Skipping Affected Release sync — either not set or value is 'P&O Release 1'`);
        }

        const requestBody = [
            {
                op: "add",
                path: "/fields/System.Title",
                value: name,
            },
            {
                op: "add",
                path: "/fields/Microsoft.VSTS.TCM.ReproSteps",
                value: description,
            },
            {
                op: "add",
                path: "/fields/System.Tags",
                value: "qTest",
            },
            {
                op: "add",
                path: "/fields/Microsoft.VSTS.Common.Severity",
                value: mappedSeverity,
            },
            {
                op: "add",
                path: "/fields/Microsoft.VSTS.Common.Priority",
                value: mappedPriority,
            },
            {
                op: "add",
                path: "/fields/System.AreaPath",
                value: constants.AreaPath,        //  using the AreaPath constant from qTest
            },
            {
                op: "add",
                path: "/relations/-",
                value: {
                    rel: "Hyperlink",
                    url: link,
                },
            },
        ];

        if (mappedDefectType) {
            requestBody.push({
                op: "add",
                path: "/fields/BP.ERP.DefectType",
                value: mappedDefectType,
            });
            console.log(`[Info] Added Defect Type to ADO: ${mappedDefectType}`);
        } else {
            console.log(`[Warn] Skipping Defect Type — no valid mapping found for qTest value '${qtestDefectType}'`);
        }

        if (mappedAffectedRelease) {
            requestBody.push({
                op: "add",
                path: "/fields/Custom.BugStage",
                value: mappedAffectedRelease,
            });
            console.log(`[Info] Added Bug Stage to ADO: ${mappedAffectedRelease}`);
        }

        if (qtestCreatedBy) {
            requestBody.push({
                op: "add",
                path: "/fields/Custom.bpCreatedBy",
                value: qtestCreatedBy
            });
            console.log(`[Info] Added Created By to ADO: ${qtestCreatedBy}`);
        } else {
            console.log(`[Info] Skipping Created By — no value found in qTest`);
        }

        if (qtestExternalReference) {
            requestBody.push({
                op: "add",
                path: "/fields/BP.ERP.ExternalReference",
                value: qtestExternalReference
            });
            console.log(`[Info] Added External Reference to ADO: ${qtestExternalReference}`);
        } else {
            console.log(`[Info] Skipping External Reference — no value in qTest`);
        }

        // --- Log full request before sending ---
        console.log(`[Debug] POST URL: ${url}`);
        console.log(`[Debug] Payload: ${JSON.stringify(requestBody, null, 2)}`);

        try {
            const response = await axios.post(url, requestBody, {
                headers: {
                    'Content-Type': 'application/json-patch+json',
                    'Authorization': `basic ${Buffer.from(`:${constants.AZDO_TOKEN}`).toString('base64')}`
                }
            });
            console.log(`[Info] Bug created in Azure DevOps`);
            return response.data;
        } catch (error) {
            console.log(`[Error] Failed to create bug in Azure DevOps: ${error}`);
            if (error.response) {
                console.log(`[Error] Failed to create bug in Azure DevOps. Status: ${error.response.status}`);
                console.log(`[Error] Status: ${error.response.status}`);
                console.log(`[Error] Data: ${JSON.stringify(error.response.data, null, 2)}`);

            } else if (error.request) {
                console.log(`[Error] No response received from ADO. Possible network or permission issue.`);
                console.log(`[Error] Request: ${JSON.stringify(error.request, null, 2)}`);
            } else {
                console.log(`[Error] Raw: ${JSON.stringify(error, null, 2)}`);
                console.log(`[Error] Raw: ${error.message}`);
                console.log(`[Error] Response: ${JSON.stringify(error.response.data, null, 2)}`);
                console.log(`[Debug] ADO Request Payload: ${JSON.stringify(requestBody, null, 2)}`);
            }
            console.log(`[Debug] Full error object: ${JSON.stringify(error, Object.getOwnPropertyNames(error), 2)}`);
            return null;
        }
    }

    async function updateDefectSummary(defectId, fieldId, fieldValue) {
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/defects/${defectId}`;
        const requestBody = {
            properties: [
                {
                    field_id: fieldId,
                    field_value: fieldValue,
                },
            ],
        };

        console.log(`[Info] Updating defect '${defectId}'.`);

        try {
            await axios.put(url, requestBody, {
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `bearer ${constants.QTEST_TOKEN}`
                }
            });
            console.log(`[Info] Defect '${defectId}' updated.`);
        } catch (error) {
            console.log(`[Error] Failed to update defect '${defectId}'.`, error);
        }
    }

    function encodeIfNeeded(url) {
        try {
            // Decode the URL to check if it's already encoded
            let decodedUrl = decodeURIComponent(url);
            // If decoding is successful, the URL was already encoded
            return url;
        } catch (e) {
            // If decoding fails, the URL needs to be encoded
            return encodeURIComponent(url);
        }
    }
};

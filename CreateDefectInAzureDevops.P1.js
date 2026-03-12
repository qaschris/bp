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
    //const defectPid = event.defect.pid;
    const projectId = event.defect.project_id;
    //console.log(`[Debug] PID value from qTest event: ${defectPid}`);
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
        defectDetails.externalReference,
        defectDetails.assignedToIdentity
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
        const assignedToField = constants.DefectAssignedToFieldID ? getFieldById(defect, constants.DefectAssignedToFieldID) : null; // Assigned To (user id)

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


        //return { summary: summary, description: description, link: link, severity: severity, priority: priority,};
        // Resolve qTest Assigned To -> identity (email/UPN) for ADO
        let assignedToIdentity = null;
        if (assignedToField && assignedToField.field_value) {
            assignedToIdentity = await resolveQTestUserIdToIdentity(assignedToField.field_value);
        } else {
            console.log(`[Info] Defect Assigned To is blank in qTest.`);
        }
        console.log(`[Info] Defect Assigned To Identity: ${assignedToIdentity}`);

        return { summary, description, link, severity, priority, defectType, status, affectedRelease, createdBy, externalReference, assignedToIdentity };
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

    //Severity Mapping (qTest → ADO)
    function mapSeverity(qtestSeverity) {
        const severityId = parseInt(qtestSeverity);  // force numeric comparison
        switch (severityId) {
            case 10301:
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
            case 11169: return 1; // Very High
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
            case 956: return "New_Requirement";
            case 957: return "Code";
            case 958: return "Data";
            case 959: return "Environment";
            case 960: return "Infrastructure";
            case 961: return "User Authorization";
            case 962: return "Configuration";
            case 963: return "User Handling";
            case 964: return "Translation";
            case 965: return "Automation";
            default: return null; // null means skip this field
        }
    }

    // qTest → ADO Status Mapping
    function mapStatus(qtestStatus) {
        const statusId = parseInt(qtestStatus);
        switch (statusId) {
            case 10001: return "New";
            case 10002: return "In Analysis";
            case 10004: return "In Resolution";
            case 10003: return "Awaiting Implementation";
            case 10953: return "Resolved";
            case 10880: return "Retest";
            case 10882: return "Reopened";
            case 10881: return "Closed";
            case 10883: return "On Hold";
            case 10853: return "Rejected";
            default: return "New"; // null means skip this field
        }
    }

    // qTest → ADO Bug Stage Mapping
    function mapAffectedRelease(qtestRelease) {
        const releaseId = parseInt(qtestRelease);
        switch (releaseId) {
            case -510: return null; // Skip "P&O Release 1"
            case 283: return "P&O_R1_SIT Dry Run";
            case 279: return "P&O_R1_SIT1";
            case 280: return "P&O_R1_SIT2";
            case 284: return "P&O_R1_DC1";
            case 285: return "P&O_R1_DC2";
            case 286: return "P&O_R1_DC3";
            case 287: return "P&O_R1_UAT";
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
        qtestExternalReference,
        qtestAssignedToIdentity
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
                value: "qTest-Dev",
            },
            {
                op: "add",
                path: "/fields/System.State",
                value: mappedStatus,
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

        // Assigned To (qTest -> ADO)
        const adoAssignedToRef = constants.AzDoAssignedToFieldRef || "System.AssignedTo";
        if (qtestAssignedToIdentity && String(qtestAssignedToIdentity).trim()) {
            requestBody.push({
                op: "add",
                path: `/fields/${adoAssignedToRef}`,
                value: String(qtestAssignedToIdentity).trim(),
            });
            console.log(`[Info] Added Assigned To to ADO: ${String(qtestAssignedToIdentity).trim()}`);
        } else {
            console.log(`[Info] Skipping Assigned To — qTest assignment is blank or could not be resolved to identity.`);
        }

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
                console.log(`[Error] Data: ${JSON.stringify(error.response && error.response.data ? error.response.data : null, null, 2)}`);

            } else if (error.request) {
                console.log(`[Error] No response received from ADO. Possible network or permission issue.`);
                console.log(`[Error] Request: ${JSON.stringify(error.request, null, 2)}`);
            } else {
                console.log(`[Error] Raw: ${JSON.stringify(error, null, 2)}`);
                console.log(`[Error] Raw: ${error.message}`);
                console.log(`[Error] Response: ${JSON.stringify(error.response && error.response.data ? error.response.data : null, null, 2)}`);
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

    async function resolveQTestUserIdToIdentity(userId) {
        // Returns a string we can send to ADO's System.AssignedTo.
        // BP qTest SSO: email may be blank; username typically holds the UPN/email.
        if (!userId) return null;

        // Build URL without encodeIfNeeded (it can mangle already-good URLs).
        // Prefer a single canonical base URL variable in your rule (whatever you use elsewhere).
        const base = (constants.qtesturl || constants.ManagerURL || "").trim();

        // Ensure base has scheme and no trailing slash
        const normalizedBase = base.startsWith("http")
            ? base.replace(/\/+$/, "")
            : `https://${base.replace(/\/+$/, "")}`;

        const url = `${normalizedBase}/api/v3/users/${userId}`;

        console.log(`[Debug] resolveQTestUserIdToIdentity URL = ${url}`);

        try {
            const response = await axios.get(url, {
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${constants.QTEST_TOKEN}`
                }
            });

            const data = response?.data;
            if (!data) return null;

            // Handle both shapes:
            // 1) direct user object: { id, username, email, ... }
            // 2) wrapper: { items: [ { id, username, ... } ], total, ... }
            const u = Array.isArray(data?.items) ? data.items[0] : data;
            if (!u) return null;

            // Identity preference for BP:
            // - If email is populated, it's fine to use it, but SSO often leaves it blank.
            // - username/ldap/external are usually the UPN you want for ADO.
            const identity =
                //(u.email && u.email.trim()) || //BP SSO often results in blank email, so deprioritizing email in favor of username fields
                (u.username && u.username.trim()) ||
                (u.ldap_username && u.ldap_username.trim()) ||
                (u.external_user_name && u.external_user_name.trim()) ||
                (u.external_username && u.external_username.trim()) ||
                null;

            if (identity) {
                console.log(
                    `[Info] Resolved qTest user ID '${userId}' to identity '${identity}' ` +
                    `(email may be blank for SSO users)`
                );
            } else {
                console.log(
                    `[Warn] qTest user ID '${userId}' has no usable identity fields ` +
                    `(email/username/ldap_username/external_*).`
                );
            }

            return identity;
        } catch (error) {
            // This is the key: log axios error details that explain "no response".
            const status = error?.response?.status;
            const code = error?.code;
            const msg = error?.message;

            if (status) {
                console.log(
                    `[Warn] Could not resolve qTest user ID '${userId}' to identity. ` +
                    `HTTP ${status}.`
                );
            } else {
                console.log(
                    `[Warn] Could not resolve qTest user ID '${userId}' to identity. ` +
                    `No HTTP response. code=${code || "n/a"} message=${msg || "n/a"} url=${url}`
                );
            }
            return null;
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
 

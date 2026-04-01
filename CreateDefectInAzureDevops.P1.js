const { Webhooks } = require('@qasymphony/pulse-sdk');
const axios = require('axios');

exports.handler = async function ({ event, constants, triggers }, context, callback) {
    let iteration = event.iteration != undefined ? event.iteration : 1;
    const maxIterations = 21;
    const defectId = event.defect.id;
    const projectId = event.defect.project_id;

    console.log(`[Info] Create defect event received for defect '${defectId}' in project '${projectId}'`);

    if (projectId != constants.ProjectID) {
        console.log(`[Info] Project not matching '${projectId}' != '${constants.ProjectID}', exiting.`);
        return;
    }

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

    function normalizeAreaPathLabel(value) {
        return typeof value === "string" ? value.trim() : "";
    }

    function mapAreaPathToQtestTeamValue(areaPath) {
        const label = normalizeAreaPathLabel(areaPath);
        return AREA_PATH_TO_QTEST_TEAM_VALUE[label] || DEFAULT_ASSIGNED_TO_TEAM_VALUE;
    }

    function mapQtestTeamValueToAreaPath(valueId) {
        return QTEST_TEAM_VALUE_TO_AREA_PATH[String(valueId)] || DEFAULT_AREA_PATH;
    }

    function encodeIfNeeded(url) {
        try {
            decodeURIComponent(url);
            return url;
        } catch (e) {
            return encodeURIComponent(url);
        }
    }

    function formatDateOnly(value) {
        if (!value) return value;
        const date = new Date(value);
        if (isNaN(date.getTime())) {
            console.log(`[Warn] Could not parse Target Date '${value}'. Passing original value.`);
            return value;
        }
        return date.toISOString().slice(0, 10);
    }

    async function getDefectDetailsByIdWithRetry(defectId) {
        let defectDetails = undefined;
        let delay = 5000;
        let attempt = 0;

        do {
            if (attempt > 0) {
                console.log(`[Warn] Could not get defect details on attempt ${attempt}. Waiting ${delay} ms.`);
                await new Promise((r) => setTimeout(r, delay));
            }

            defectDetails = await getDefectDetailsById(defectId);

            if (
                defectDetails &&
                defectDetails.summary &&
                defectDetails.description &&
                defectDetails.severity
            ) {
                console.log(`[Info] Successfully fetched complete defect details for '${defectId}' on attempt ${attempt + 1}.`);
                return defectDetails;
            }

            attempt++;
        } while (attempt < 12);

        console.error(`[Error] Could not get defect details after retry loop. User may not have completed the initial save in qTest or the defect was abandoned.`);

        if (iteration < maxIterations) {
            iteration = iteration + 1;
            console.log(`[Info] Re-executing rule. Iteration ${iteration} of ${maxIterations}.`);
            event.iteration = iteration;
            emitEvent('qTestDefectSubmitted', event);
            return null;
        }

        emitFriendlyFailure({
            platform: "qTest",
            objectType: "Defect",
            objectId: defectId,
            detail: `Unable to retrieve required defect details after ${maxIterations} rule iterations.`
        });

        return null;
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
                emitFriendlyFailure({
                    platform: "qTest",
                    objectType: "Defect",
                    objectId: defectId,
                    detail: "Defect was not found in qTest and may have been abandoned."
                });
                return null;
            }

            console.error("[Error] Failed to get defect by id.", error);

            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Defect",
                objectId: defectId,
                detail: "Unable to read defect details from qTest."
            });

            return null;
        }
    }

    async function getDefectDetailsById(defectId) {
        const defect = await getDefectById(defectId);
        if (!defect) return null;

        const summaryField = getFieldById(defect, constants.DefectSummaryFieldID);
        const descriptionField = getFieldById(defect, constants.DefectDescriptionFieldID);
        const severityField = getFieldById(defect, constants.DefectSeverityFieldID);
        const priorityField = getFieldById(defect, constants.DefectPriorityFieldID);
        const defectTypeField = getFieldById(defect, constants.DefectTypeFieldID);
        const statusField = getFieldById(defect, constants.DefectStatusFieldID);
        const affectedReleaseField = getFieldById(defect, constants.DefectAffectedReleaseFieldID);
        const createdByField = getFieldById(defect, constants.DefectCreatedByFieldID);
        const externalReferenceField = getFieldById(defect, constants.DefectExternalReferenceFieldID);
        const assignedToField = constants.DefectAssignedToFieldID ? getFieldById(defect, constants.DefectAssignedToFieldID) : null;
        const targetDateField = getFieldById(defect, constants.DefectTargetDateFieldID);
        const assignedToTeamField = constants.DefectAssignedToTeamFieldID ? getFieldById(defect, constants.DefectAssignedToTeamFieldID) : null;
        const targetDate = targetDateField ? targetDateField.field_value : null;

        console.log(`[Info] Defect Target Date: ${targetDate}`);

        if (!summaryField || !descriptionField || !severityField) {
            const missingFields = [];
            if (!summaryField) missingFields.push("Summary");
            if (!descriptionField) missingFields.push("Description");
            if (!severityField) missingFields.push("Severity");

            const missingFieldLabel = missingFields.join(", ");

            console.error(`[Error] Required defect fields not found: ${missingFieldLabel}.`);

            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Defect",
                objectId: defectId,
                fieldName: missingFieldLabel,
                detail: "Required field data is missing."
            });

            return null;
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
            createdBy = await getQtestUserName(createdBy);
        }
        console.log(`[Info] Defect Created By: ${createdBy}`);

        const externalReference = externalReferenceField ? externalReferenceField.field_value : null;
        console.log(`[Info] Defect External Reference: ${externalReference}`);

        let assignedToIdentity = null;
        if (assignedToField && assignedToField.field_value) {
            assignedToIdentity = await resolveQTestUserIdToIdentity(assignedToField.field_value);
        } else {
            console.log(`[Info] Defect Assigned To is blank in qTest.`);
        }
        console.log(`[Info] Defect Assigned To Identity: ${assignedToIdentity}`);

        let assignedToTeamLabel = DEFAULT_AREA_PATH;
        if (assignedToTeamField) {
            const rawTeamLabel = assignedToTeamField.field_value_name;
            const rawTeamValue = assignedToTeamField.field_value;

            assignedToTeamLabel =
                normalizeAreaPathLabel(rawTeamLabel) ||
                mapQtestTeamValueToAreaPath(rawTeamValue);

            console.log(`[Info] Defect Assigned to Team Label: ${assignedToTeamLabel}`);
        } else {
            console.log(`[Info] Defect Assigned to Team is blank in qTest. Defaulting ADO AreaPath to '${DEFAULT_AREA_PATH}'.`);
        }

        return {
            summary,
            description,
            link,
            severity,
            priority,
            defectType,
            status,
            affectedRelease,
            createdBy,
            externalReference,
            assignedToIdentity,
            assignedToTeamLabel,
            targetDate
        };
    }

    function mapSeverity(qtestSeverity) {
        const severityId = parseInt(qtestSeverity);
        switch (severityId) {
            case 10301: return '1 - Critical';
            case 10302: return '2 - High';
            case 10303: return '3 - Medium';
            case 10304: return '4 - Low';
            default: return '3 - Medium';
        }
    }

    function mapPriority(qtestPriority) {
        const priorityId = parseInt(qtestPriority);
        switch (priorityId) {
            case 11169: return 1;
            case 10204: return 2;
            case 10203: return 3;
            case 10202: return 4;
            default: return 3;
        }
    }

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
            default: return null;
        }
    }

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
            case 11376: return "Triage";
            default: return "New";
        }
    }

    function mapAffectedRelease(qtestRelease) {
        const releaseId = parseInt(qtestRelease);
        switch (releaseId) {
            case -510: return null;
            case 283: return "P&O_R1_SIT Dry Run";
            case 279: return "P&O_R1_SIT1";
            case 280: return "P&O_R1_SIT2";
            case 284: return "P&O_R1_DC1";
            case 285: return "P&O_R1_DC2";
            case 286: return "P&O_R1_DC3";
            case 287: return "P&O_R1_UAT";
            case 302: return "Unit Testing";
            default: return null;
        }
    }

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
            let displayName = "";

            if (userData.first_name && userData.last_name) {
                displayName = `${userData.first_name.trim()} ${userData.last_name.trim()}`;
            } else if (userData.last_name) {
                displayName = userData.last_name.trim();
            } else {
                displayName = userData.username || userData.email || userId.toString();
            }

            console.log(`[Info] Resolved qTest user ID '${userId}' to name '${displayName}'`);
            return displayName;
        } catch (error) {
            console.error(`[Error] Could not resolve qTest user ID '${userId}' to name. Using ID instead.`);
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
        qtestAssignedToIdentity,
        qtestAssignedToTeamLabel,
        qtestTargetDate
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
            { op: "add", path: "/fields/System.Title", value: name },
            { op: "add", path: "/fields/Microsoft.VSTS.TCM.ReproSteps", value: description },
            { op: "add", path: "/fields/System.Tags", value: "qTest-Dev" },
            { op: "add", path: "/fields/System.State", value: mappedStatus },
            { op: "add", path: "/fields/Microsoft.VSTS.Common.Severity", value: mappedSeverity },
            { op: "add", path: "/fields/Microsoft.VSTS.Common.Priority", value: mappedPriority },
            {
                op: "add",
                path: "/fields/System.AreaPath",
                value: normalizeAreaPathLabel(qtestAssignedToTeamLabel) || DEFAULT_AREA_PATH
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

        if (qtestTargetDate) {
            requestBody.push({
                op: "add",
                path: "/fields/Microsoft.VSTS.Scheduling.TargetDate",
                value: formatDateOnly(qtestTargetDate)
            });
        }

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
                console.log(`[Error] Data: ${JSON.stringify(error.response.data || null, null, 2)}`);
            } else if (error.request) {
                console.log(`[Error] No response received from ADO. Possible network or permission issue.`);
                console.log(`[Error] Request: ${JSON.stringify(error.request, null, 2)}`);
            } else {
                console.log(`[Error] Raw: ${JSON.stringify(error, null, 2)}`);
                console.log(`[Error] Message: ${error.message}`);
            }

            console.log(`[Debug] ADO Request Payload: ${JSON.stringify(requestBody, null, 2)}`);
            console.log(`[Debug] Full error object: ${JSON.stringify(error, Object.getOwnPropertyNames(error), 2)}`);

            emitFriendlyFailure({
                platform: "ADO",
                objectType: "Defect",
                objectId: defectId,
                detail: "Unable to create defect in Azure DevOps."
            });

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
            return true;
        } catch (error) {
            console.log(`[Error] Failed to update defect '${defectId}'.`, error);

            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Defect",
                objectId: defectId,
                fieldName: "Summary",
                fieldValue: fieldValue,
                detail: "Unable to update qTest defect after Azure DevOps creation."
            });

            return false;
        }
    }

    async function resolveQTestUserIdToIdentity(userId) {
        if (!userId) return null;

        const base = (constants.qtesturl || constants.ManagerURL || "").trim();
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

            const u = Array.isArray(data?.items) ? data.items[0] : data;
            if (!u) return null;

            const identity =
                (u.username && u.username.trim()) ||
                (u.ldap_username && u.ldap_username.trim()) ||
                (u.external_user_name && u.external_user_name.trim()) ||
                (u.external_username && u.external_username.trim()) ||
                null;

            if (identity) {
                console.log(`[Info] Resolved qTest user ID '${userId}' to identity '${identity}' (email may be blank for SSO users)`);
            } else {
                console.log(`[Warn] qTest user ID '${userId}' has no usable identity fields (email/username/ldap_username/external_*).`);
            }

            return identity;
        } catch (error) {
            const status = error?.response?.status;
            const code = error?.code;
            const msg = error?.message;

            if (status) {
                console.log(`[Warn] Could not resolve qTest user ID '${userId}' to identity. HTTP ${status}.`);
            } else {
                console.log(`[Warn] Could not resolve qTest user ID '${userId}' to identity. No HTTP response. code=${code || "n/a"} message=${msg || "n/a"} url=${url}`);
            }

            return null;
        }
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
        defectDetails.assignedToIdentity,
        defectDetails.assignedToTeamLabel,
        defectDetails.targetDate
    );

    if (!bug) return;

    const workItemId = bug.id;
    const newSummary = `${getNamePrefix(workItemId)}${defectDetails.summary}`;
    console.log(`[Info] New defect name: ${newSummary}`);

    const summaryUpdated = await updateDefectSummary(defectId, constants.DefectSummaryFieldID, newSummary);
    if (!summaryUpdated) return;
};
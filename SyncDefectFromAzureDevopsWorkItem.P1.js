const { Webhooks } = require('@qasymphony/pulse-sdk');
const axios = require("axios");

exports.handler = async function ({ event, constants, triggers }, context, callback) {

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

    const standardHeaders = {
        "Content-Type": "application/json",
        Authorization: `bearer ${constants.QTEST_TOKEN}`,
    };

    const eventType = {
        CREATED: "workitem.created",
        UPDATED: "workitem.updated",
        DELETED: "workitem.deleted",
    };

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

    function buildDefectDescription(eventData) {
        const fields = getFields(eventData);
        return `Link to Azure DevOps: ${eventData.resource._links.html.href}
                Repro steps: 
                ${htmlToPlainText(fields["Microsoft.VSTS.TCM.ReproSteps"])}`;
    }

    function buildDefectSummary(namePrefix, eventData) {
        const fields = getFields(eventData);
        return `${namePrefix}${fields["System.Title"]}`;
    }

    function getFields(eventData) {
        return eventData.resource.revision ? eventData.resource.revision.fields : eventData.resource.fields;
    }

    function getNamePrefix(workItemId) {
        return `WI${workItemId}: `;
    }

    function htmlToPlainText(htmlText) {
        if (!htmlText || htmlText.length === 0) return "";
        return htmlText
            .replace(/<style([\s\S]*?)<\/style>/gi, "")
            .replace(/<script([\s\S]*?)<\/script>/gi, "")
            .replace(/<\/div>/gi, "\n")
            .replace(/<\/li>/gi, "\n")
            .replace(/<li>/gi, "  *  ")
            .replace(/<\/ul>/gi, "\n")
            .replace(/<\/p>/gi, "\n")
            .replace(/<br\s*[\/]?>/gi, "\n")
            .replace(/<[^>]+>/gi, "")
            .replace(/\n\s*\n/gi, "\n");
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

    function extractUpnOrEmailFromAdoAssignedTo(raw) {
        if (!raw) return null;

        if (typeof raw === "object") {
            const candidate =
                raw.uniqueName ||
                raw.mail ||
                raw.email ||
                raw.userPrincipalName ||
                raw.displayName ||
                "";
            return candidate ? candidate.trim() : null;
        }

        if (typeof raw !== "string") return null;

        const s = raw.trim();
        const angle = s.match(/<([^>]+)>/);
        if (angle && angle[1]) return angle[1].trim();

        const email = s.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
        if (email && email[0]) return email[0].trim();

        return s || null;
    }

    async function getAdoComments(workItemId) {
        const url = `${constants.AzDoProjectURL}/_apis/wit/workitems/${workItemId}/comments?api-version=7.0-preview.3`;
        const encodedToken = Buffer.from(`:${constants.AZDO_TOKEN}`).toString("base64");

        try {
            const response = await axios.get(url, {
                headers: { Authorization: `Basic ${encodedToken}` }
            });

            return response.data?.comments || [];
        } catch (error) {
            console.error(`[Error] Failed to fetch ADO comments: ${error.message}`);
            emitFriendlyFailure({
                platform: "ADO",
                objectType: "Defect",
                objectId: workItemId,
                fieldName: "Discussion",
                detail: "Unable to retrieve comments from Azure DevOps."
            });
            return [];
        }
    }

    function formatDiscussion(comments) {
        if (!comments || comments.length === 0) return "";

        return comments.map(c => {
            const author = c.createdBy?.displayName || "Unknown";
            const date = new Date(c.createdDate).toLocaleString();
            const text = c.text || "";

            return `
                <div style="margin-bottom:10px;">
                <strong>${author}</strong>
                <span style="color:gray;">${date}</span>
                <div>${text}</div>
                </div>
                <hr/>`;
        }).join("");
    }

    async function resolveQtestUserIdByUsernameOrUpn(identity, standardHeaders) {
        if (!identity) return null;

        const url = `https://${constants.ManagerURL}/api/v3/users/search?username=${encodeURIComponent(identity)}`;

        try {
            const resp = await axios.get(url, { headers: standardHeaders });
            const data = resp && resp.data;

            const arr = Array.isArray(data)
                ? data
                : Array.isArray(data?.items)
                    ? data.items
                    : Array.isArray(data?.data)
                        ? data.data
                        : [];

            const user = arr[0];
            return user?.id ?? null;
        } catch (e) {
            console.error(
                `[Error] qTest user search failed for '${identity}'. ` +
                `Status: ${e?.response?.status ?? "n/a"}`
            );
            return null;
        }
    }

    async function getDefectByWorkItemId(workItemId) {
        const prefix = getNamePrefix(workItemId);
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/search`;
        const requestBody = {
            object_type: "defects",
            fields: ["*"],
            query: `Summary ~ '${prefix}'`,
        };

        console.log(`[Info] Get existing defect for 'WI${workItemId}'`);
        let failed = false;
        let defect = undefined;

        try {
            const response = await post(url, requestBody);

            if (!response || response.total === 0) {
                console.log("[Info] Defect not found by work item id.");
            } else if (response.total === 1) {
                defect = response.items[0];
            } else {
                failed = true;
                console.log("[Warn] Multiple Defects found by work item id.");
                emitFriendlyFailure({
                    platform: "qTest",
                    objectType: "Defect",
                    objectId: workItemId,
                    detail: "Multiple matching qTest defects were found for this Azure DevOps work item."
                });
            }
        } catch (error) {
            console.error("[Error] Failed to get defect by work item id.", error);
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Defect",
                objectId: workItemId,
                detail: "Unable to locate the matching qTest defect."
            });
            failed = true;
        }

        return { failed, defect };
    }

    async function updateDefect(
        defectToUpdate,
        summary,
        description,
        severityValue,
        priorityValue,
        rootCauseValue,
        defectTypeValue,
        statusValue,
        proposedFixValue,
        closedDateValue,
        resolvedReasonValue,
        assignedToUserId,
        targetDateValue,
        discussionValue,
        qtestAssignedToTeamValue,
        adoExternalReference,
        adoState
    ) {
        const defectId = defectToUpdate.id;
        const defectPid = defectToUpdate.pid;

        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/defects/${defectId}`;
        const requestBody = {
            properties: [
                {
                    field_id: constants.DefectSummaryFieldID,
                    field_value: summary,
                },
                {
                    field_id: constants.DefectDescriptionFieldID,
                    field_value: description,
                },
            ],
        };

        if (severityValue) {
            requestBody.properties.push({
                field_id: constants.DefectSeverityFieldID,
                field_value: parseInt(severityValue, 10),
            });
        }

        if (priorityValue) {
            requestBody.properties.push({
                field_id: constants.DefectPriorityFieldID,
                field_value: parseInt(priorityValue, 10),
            });
        }

        if (constants.DefectRootCauseFieldID) {
            requestBody.properties.push({
                field_id: constants.DefectRootCauseFieldID,
                field_value: rootCauseValue || "",
            });
        }

        if (defectTypeValue) {
            requestBody.properties.push({
                field_id: constants.DefectTypeFieldID,
                field_value: parseInt(defectTypeValue, 10),
            });
            console.log(`[Info] Added Defect Type '${defectTypeValue}' to qTest update payload.`);
        } else {
            console.log(`[Warn] No Defect Type mapping found or field is empty in ADO.`);
        }

        if (statusValue) {
            requestBody.properties.push({
                field_id: constants.DefectStatusFieldID,
                field_value: parseInt(statusValue, 10),
            });
            console.log(`[Info] Added Status '${statusValue}' to qTest update payload.`);
        } else {
            console.log(`[Warn] No Status mapping found or ADO state '${adoState}' not mapped.`);
        }

        if (constants.DefectProposedFixFieldID) {
            const formattedProposedFix = proposedFixValue ? `<p>${proposedFixValue}</p>` : "";
            requestBody.properties.push({
                field_id: constants.DefectProposedFixFieldID,
                field_value: formattedProposedFix,
            });
            console.log(`[Info] Added Proposed Fix to qTest update payload.`);
        }

        if (constants.DefectExternalReferenceFieldID) {
            requestBody.properties.push({
                field_id: constants.DefectExternalReferenceFieldID,
                field_value: adoExternalReference || "",
            });
            console.log(`[Info] Added External Reference to qTest update payload.`);
        }

        if (constants.DefectClosedDateFieldID && closedDateValue) {
            requestBody.properties.push({
                field_id: constants.DefectClosedDateFieldID,
                field_value: closedDateValue,
            });
            console.log(`[Info] Added Closed Date '${closedDateValue}' to qTest update payload.`);
        }

        if (constants.DefectResolvedReasonFieldID && resolvedReasonValue) {
            requestBody.properties.push({
                field_id: constants.DefectResolvedReasonFieldID,
                field_value: parseInt(resolvedReasonValue, 10),
            });
            console.log(`[Info] Added Resolved Reason '${resolvedReasonValue}' to qTest update payload.`);
        } else {
            console.log(`[Warn] No Resolved Reason provided or mapping not found`);
        }

        if (constants.DefectTargetDateFieldID && targetDateValue) {
            requestBody.properties.push({
                field_id: constants.DefectTargetDateFieldID,
                field_value: targetDateValue
            });
        }

        if (constants.DefectDiscussionFieldID) {
            requestBody.properties.push({
                field_id: constants.DefectDiscussionFieldID,
                field_value: discussionValue || ""
            });
            console.log(`[Info] Added Discussion to qTest update payload.`);
        }

        if (constants.DefectAssignedToFieldID) {
            requestBody.properties.push({
                field_id: constants.DefectAssignedToFieldID,
                field_value: assignedToUserId ? assignedToUserId : ""
            });

            console.log(
                assignedToUserId
                    ? `[Info] Added Assigned To userId '${assignedToUserId}' to qTest update payload.`
                    : `[Info] Clearing qTest Assigned To (Blank) in update payload.`
            );
        }

        if (constants.DefectAssignedToTeamFieldID && qtestAssignedToTeamValue) {
            requestBody.properties.push({
                field_id: constants.DefectAssignedToTeamFieldID,
                field_value: parseInt(qtestAssignedToTeamValue, 10),
            });
            console.log(
                `[Info] Added Assigned to Team '${qtestAssignedToTeamValue}' to qTest update payload.`
            );
        }

        console.log(`[Info] Updating defect '${defectId}' (${defectPid}).`);
        console.log('[Debug] Final qTest Update Payload:', JSON.stringify(requestBody, null, 2));

        try {
            await put(url, requestBody);
            console.log(`[Info] Defect '${defectId}' (${defectPid}) updated.`);
        } catch (error) {
            console.error(`[Error] Failed to update defect '${defectId}'.`, error);
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Defect",
                objectId: defectId,
                detail: "Unable to update the qTest defect from Azure DevOps."
            });
        }
    }

    function post(url, requestBody) {
        return doqTestRequest(url, "POST", requestBody);
    }

    function put(url, requestBody) {
        return doqTestRequest(url, "PUT", requestBody);
    }

    async function doqTestRequest(url, method, requestBody) {
        const opts = {
            url: url,
            json: true,
            headers: standardHeaders,
            data: method === "GET" ? undefined : requestBody,
            method: method,
        };

        try {
            const response = await axios(opts);
            return response.data;
        } catch (error) {
            const status = error?.response?.status || "Unknown";
            const message = error?.response?.data
                ? JSON.stringify(error.response.data, null, 2)
                : error.message;

            console.error(`[Error] URL: ${url}`);
            console.error(`[Error] HTTP Status: ${status}`);
            console.error(`[Error] Message: ${message}`);

            throw new Error(`qTest API ${method} ${url} failed with ${status}: ${message}`);
        }
    }

    let workItemId = undefined;
    let defectToUpdate = undefined;

    switch (event.eventType) {
        case eventType.CREATED: {
            workItemId = event.resource.workItemId;
            console.log(`[Info] Create workitem event received for 'WI${workItemId}'`);
            console.log(`[Info] New defects are not synched from Azure DevOps. The current workflow expects the defect to be created in qTest first. Exiting.`);
            return;
        }
        case eventType.UPDATED: {
            workItemId = event.resource.workItemId;
            console.log(`[Info] Update workitem event received for 'WI${workItemId}'`);

            const getDefectResult = await getDefectByWorkItemId(workItemId);
            if (getDefectResult.failed) {
                return;
            }

            if (getDefectResult.defect === undefined) {
                console.log("[Info] Corresponding defect not found. Exiting.");
                return;
            }

            defectToUpdate = getDefectResult.defect;
            break;
        }
        case eventType.DELETED: {
            workItemId = event.resource.workItemId;
            console.log(`[Info] Delete workitem event received for 'WI${workItemId}'`);
            console.log(`[Info] Defects are not deleted in qTest automatically when deleting in Azure DevOps. Exiting.`);
            return;
        }
        default:
            console.log(`[Error] Unknown workitem event type '${event.eventType}' for 'WI${workItemId}'`);
            emitFriendlyFailure({
                platform: "ADO",
                objectType: "Defect",
                objectId: workItemId || "Unknown",
                detail: `Unsupported work item event type '${event.eventType}'.`
            });
            return;
    }

    const namePrefix = getNamePrefix(workItemId);
    const defectDescription = buildDefectDescription(event);
    const defectSummary = buildDefectSummary(namePrefix, event);
    const fields = getFields(event);

    const adoAreaPath = fields[constants.AzDoAreaPathFieldRef || "System.AreaPath"] || "";
    console.log(`[Info] ADO AreaPath value: '${adoAreaPath}'`);

    const qtestAssignedToTeamValue = mapAreaPathToQtestTeamValue(adoAreaPath);
    console.log(
        `[Info] Mapped ADO AreaPath '${adoAreaPath || "(blank)"}' to qTest Assigned to Team value '${qtestAssignedToTeamValue}'`
    );

    if (!AREA_PATH_TO_QTEST_TEAM_VALUE[normalizeAreaPathLabel(adoAreaPath)]) {
        console.log(
            `[Warn] ADO AreaPath '${adoAreaPath || "(blank)"}' not found in qTest team mapping. Defaulting to '${DEFAULT_AREA_PATH}' (${DEFAULT_ASSIGNED_TO_TEAM_VALUE}).`
        );
    }

    let adoComments = await getAdoComments(workItemId);

    if (constants.SyncUserRegex) {
        const regex = new RegExp(constants.SyncUserRegex, "i");
        adoComments = adoComments.filter(c => !regex.test(c.createdBy?.displayName || ""));
    }

    const discussionHtml = formatDiscussion(adoComments);
    let qtestDiscussionValue = "";

    if (discussionHtml) {
        qtestDiscussionValue = `
                    <h3>ADO Discussion</h3>
                    ${discussionHtml}
                    `;
    }

    const adoAssignedToRaw = fields[constants.AzDoAssignedToFieldRef || "System.AssignedTo"];
    const adoAssignedToIdentity = extractUpnOrEmailFromAdoAssignedTo(adoAssignedToRaw);

    let qtestAssignedToUserId = null;
    if (adoAssignedToIdentity) {
        console.log(`[Info] Normalized ADO Assigned To to '${adoAssignedToIdentity}'`);

        qtestAssignedToUserId = await resolveQtestUserIdByUsernameOrUpn(adoAssignedToIdentity, standardHeaders);

        if (qtestAssignedToUserId) {
            console.log(`[Info] Resolved ADO Assigned To '${adoAssignedToIdentity}' -> qTest userId '${qtestAssignedToUserId}'`);
        } else {
            console.log(
                `[Warn] Could not resolve ADO Assigned To '${adoAssignedToIdentity}' in qTest. ` +
                `Will clear qTest assignment (Blank).`
            );
        }
    } else if (adoAssignedToRaw) {
        console.log(`[Warn] Could not normalize ADO Assigned To '${adoAssignedToRaw}'`);
    } else {
        console.log(`[Info] ADO Assigned To is blank/unassigned.`);
    }

    const adoSeverity = fields["Microsoft.VSTS.Common.Severity"];
    console.log(`[Info] ADO Severity value: '${adoSeverity}'`);

    const severityMap = {
        "1 - Critical": 10301,
        "2 - High": 10302,
        "3 - Medium": 10303,
        "4 - Low": 10304,
    };

    const qtestSeverityValue = severityMap[adoSeverity] || null;
    console.log(`[Info] Mapped ADO Severity '${adoSeverity}' to qTest Severity`);

    const adoPriority = fields["Microsoft.VSTS.Common.Priority"];
    console.log(`[Info] ADO Priority value: '${adoPriority}'`);

    const priorityMap = {
        1: 11169,
        2: 10204,
        3: 10203,
        4: 10202,
    };

    const qtestPriorityValue = priorityMap[adoPriority] || null;
    console.log(`[Info] Mapped ADO Priority '${adoPriority}' to qTest Priority value '${qtestPriorityValue}'`);

    const adoDefectType = fields["BP.ERP.DefectType"];
    console.log(`[Info] ADO Defect Type value: '${adoDefectType}'`);

    const defectTypeMap = {
        "New_Requirement": 956,
        "Code": 957,
        "Data": 958,
        "Environment": 959,
        "Infrastructure": 960,
        "User Authorization": 961,
        "Configuration": 962,
        "User Handling": 963,
        "Translation": 964,
        "Automation": 965,
    };

    const qtestDefectTypeValue = defectTypeMap[adoDefectType] || null;
    console.log(`[Info] Mapped ADO Defect Type '${adoDefectType}' to qTest Defect Type value '${qtestDefectTypeValue}'`);

    const adoState = fields["System.State"];
    console.log(`[Info] ADO State value: '${adoState}'`);

    const stateMap = {
        "New": 10001,
        "In Analysis": 10002,
        "Active": 10002,
        "In Resolution": 10004,
        "Awaiting Implementation": 10003,
        "Resolved": 10953,
        "Retest": 10880,
        "Reopened": 10882,
        "Closed": 10881,
        "On Hold": 10883,
        "Rejected": 10853,
        "Cancelled": 10853,
        "Triage": 11376
    };

    const qtestStatusValue = stateMap[adoState] || null;
    console.log(`[Info] Mapped ADO State '${adoState}' to qTest Status value '${qtestStatusValue}'`);
    if (!qtestStatusValue) {
        console.log(`[Warn] ADO State '${adoState}' does not match any defined qTest status.`);
    }

    const adoRootCause =
        fields["Microsoft.VSTS.CMMI.RootCause@OData.Community.Display.V1.FormattedValue"] ||
        fields["Microsoft.VSTS.CMMI.RootCause"] || "";
    console.log(`[Info] ADO Root Cause value: '${adoRootCause}'`);

    const qtestRootCauseValue = adoRootCause;

    const adoProposedFix = fields["Microsoft.VSTS.CMMI.ProposedFix"] || "";
    console.log(`[Info] ADO Proposed Fix value length: ${adoProposedFix.length}`);

    const qtestProposedFixValue = adoProposedFix;

    const adoActualCloseDate = fields["BP.ERP.ActualClose"];
    let qtestClosedDateValue = null;

    const adoTargetDate = fields["Microsoft.VSTS.Scheduling.TargetDate"];
    let qtestTargetDateValue = null;
    if (adoTargetDate) {
        qtestTargetDateValue = new Date(adoTargetDate).toISOString().replace(".000Z", "+00:00");
        console.log(`[Info] ADO Target Date: '${adoTargetDate}' => qTest Target Date: '${qtestTargetDateValue}'`);
    }

    const adoExternalReference = fields["BP.ERP.ExternalReference"] || "";
    console.log(`[Info] ADO External Reference value: '${adoExternalReference}'`);

    if (adoActualCloseDate) {
        qtestClosedDateValue = new Date(adoActualCloseDate).toISOString().replace(".000Z", "+00:00");
        console.log(`[Info] ADO Actual Close Date: '${adoActualCloseDate}' => qTest Closed Date: '${qtestClosedDateValue}'`);
    } else {
        console.log(`[Info] No Actual Close Date found in ADO.`);
    }

    const adoResolvedReason =
        fields["Microsoft.VSTS.Common.ResolvedReason"] ||
        fields["Microsoft.VSTS.Common.ResolvedReason@OData.Community.Display.V1.FormattedValue"] ||
        "";
    console.log(`[Info] ADO Resolved Reason: '${adoResolvedReason}'`);

    const resolvedReasonMap = {
        "As Designed": 1299,
        "Cannot Reproduce": 1300,
        "Copied to Backlog": 1301,
        "Deferred": 1302,
        "Duplicate": 1303,
        "Fixed": 1304,
        "Fixed and verified": 1305,
        "Obsolete": 1306,
        "Will not Fix": 1307,
    };

    const qtestResolvedReasonValue = resolvedReasonMap[adoResolvedReason] || null;
    console.log(`[Info] Mapped ADO Resolved Reason '${adoResolvedReason}' → qTest value '${qtestResolvedReasonValue}'`);

    if (defectToUpdate) {
        await updateDefect(
            defectToUpdate,
            defectSummary,
            defectDescription,
            qtestSeverityValue,
            qtestPriorityValue,
            qtestRootCauseValue,
            qtestDefectTypeValue,
            qtestStatusValue,
            qtestProposedFixValue,
            qtestClosedDateValue,
            qtestResolvedReasonValue,
            qtestAssignedToUserId,
            qtestTargetDateValue,
            qtestDiscussionValue,
            qtestAssignedToTeamValue,
            adoExternalReference,
            adoState
        );
    }
};
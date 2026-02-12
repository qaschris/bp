const axios = require("axios");

// DO NOT EDIT exported "handler" function is the entrypoint
exports.handler = async function ({ event, constants, triggers }, context, callback) {
    function buildDefectDescription(eventData) {
        const fields = getFields(eventData);
        return `Link to Azure DevOps: ${eventData.resource._links.html.href}
Type: ${fields["System.WorkItemType"]}
Area: ${fields["System.AreaPath"]}
Iteration: ${fields["System.IterationPath"]}
State: ${fields["System.State"]}
Reason: ${fields["System.Reason"]}
Repro steps: 
${htmlToPlainText(fields["Microsoft.VSTS.TCM.ReproSteps"])}
System info:
${htmlToPlainText(fields["Microsoft.VSTS.TCM.SystemInfo"])}
Acceptance criteria:
${htmlToPlainText(fields["Microsoft.VSTS.Common.AcceptanceCriteria"])}`;
    }

    function buildDefectSummary(namePrefix, eventData) {
        const fields = getFields(eventData);
        return `${namePrefix}${fields["System.Title"]}`;
    }

    function getFields(eventData) {
        // In case of update the fields can be taken from the revision, in case of create from the resource directly
        return eventData.resource.revision ? eventData.resource.revision.fields : eventData.resource.fields;
    }

    const standardHeaders = {
        "Content-Type": "application/json",
        Authorization: `bearer ${constants.QTEST_TOKEN}`,
    };
    const eventType = {
        CREATED: "workitem.created",
        UPDATED: "workitem.updated",
        DELETED: "workitem.deleted",
    };

    let workItemId = undefined;
    let defectToUpdate = undefined;
    switch (event.eventType) {
        case eventType.CREATED: {
            console.log(`[Info] Create workitem event received for 'WI${workItemId}'`);
            console.log(
                `[Info] New defects are not synched from Azure DevOps. The current workflow expects the defect to be created in qTest first. Exiting.`
            );
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
            console.log(`[Info] Delete workitem event received for 'WI${workItemId}'`);
            console.log(
                `[Info] Defects are not deleted in qTest automatically when deleting in Azure DevOps. Exiting.`
            );
            return;
        }
        default:
            console.log(`[Error] Unknown workitem event type '${event.eventType}' for 'WI${workitemId}'`);
            return;
    }

    // Prepare data to create/update requirement
    const namePrefix = getNamePrefix(workItemId);
    const defectDescription = buildDefectDescription(event);
    const defectSummary = buildDefectSummary(namePrefix, event);

    if (defectToUpdate) {
        const assignedToFieldValue = await resolveQTestAssignedToFieldValue(event, defectToUpdate);
        await updateDefect(defectToUpdate, defectSummary, defectDescription, assignedToFieldValue);
    }

    function getNamePrefix(workItemId) {
        return `WI${workItemId}: `;
    }

    // -------------------------
    // P1: Assigned To mapping (ADO -> qTest)
    // -------------------------
    let _qtestUsersCache = null;
    let _defectFieldsCache = null;
    let _assignedToDefectFieldIdCache = null;

    async function getQTestProjectUsers() {
        if (_qtestUsersCache) return _qtestUsersCache;

        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/users`;
        try {
            const response = await axios({
                url,
                method: "GET",
                headers: standardHeaders,
            });

            _qtestUsersCache = Array.isArray(response.data) ? response.data : [];
            return _qtestUsersCache;
        } catch (error) {
            console.log("[Error] Failed to get qTest project users.", error);
            _qtestUsersCache = [];
            return _qtestUsersCache;
        }
    }

    async function getDefectFields() {
        if (_defectFieldsCache) return _defectFieldsCache;

        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/settings/defects/fields`;
        try {
            const response = await axios({
                url,
                method: "GET",
                headers: standardHeaders,
            });

            _defectFieldsCache = Array.isArray(response.data) ? response.data : [];
            return _defectFieldsCache;
        } catch (error) {
            console.log("[Error] Failed to get qTest defect fields.", error);
            _defectFieldsCache = [];
            return _defectFieldsCache;
        }
    }

    async function getAssignedToDefectFieldId() {
        if (_assignedToDefectFieldIdCache) return _assignedToDefectFieldIdCache;

        if (constants.DefectAssignedToFieldID) {
            _assignedToDefectFieldIdCache = constants.DefectAssignedToFieldID;
            return _assignedToDefectFieldIdCache;
        }

        const fields = await getDefectFields();
        const assigned = fields.find(f => ((f.label || f.name || f.display_name || "") + "").toLowerCase() === "assigned to");
        _assignedToDefectFieldIdCache = assigned ? assigned.id : null;
        return _assignedToDefectFieldIdCache;
    }

    function parseEmailFromString(s) {
        if (!s) return null;
        const str = s.toString();
        const m = str.match(/<([^>]+@[^>]+)>/);
        if (m) return m[1].trim();
        const m2 = str.match(/([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})/i);
        return m2 ? m2[1].trim() : null;
    }

    function extractAdoAssignedToIdentity(raw) {
        if (!raw) return null;

        if (typeof raw === "string") {
            const email = parseEmailFromString(raw);
            return email ? { email, displayName: raw } : null;
        }

        if (typeof raw === "object") {
            // ADO identity objects commonly include uniqueName (email) and displayName
            const email = raw.uniqueName || raw.mailAddress || raw.email || raw.mail || null;
            const displayName = raw.displayName || raw.name || null;
            if (email) return { email: email.toString(), displayName: displayName ? displayName.toString() : null };

            // Sometimes AssignedTo is a string nested in an object
            const asString = raw.toString && raw.toString();
            const parsed = parseEmailFromString(asString);
            return parsed ? { email: parsed, displayName: displayName || asString } : null;
        }

        return null;
    }

    function normalize(s) {
        return (s || "").toString().trim().toLowerCase();
    }

    function getUserCandidateStrings(u) {
        const candidates = [];
        ["email", "username", "user_name", "login", "name", "displayName", "display_name"].forEach(k => {
            if (u && u[k]) candidates.push(u[k].toString());
        });

        if (u && (u.first_name || u.last_name)) {
            candidates.push(`${u.first_name || ""} ${u.last_name || ""}`.trim());
        }

        return candidates.map(normalize).filter(Boolean);
    }

    function parseQTestIdList(value) {
        if (!value) return [];
        if (Array.isArray(value)) return value.map(v => Number(v)).filter(v => !Number.isNaN(v));
        const s = value.toString().trim();
        const m = s.match(/\[([^\]]*)\]/);
        if (!m) return [];
        return m[1]
            .split(",")
            .map(x => Number((x || "").trim()))
            .filter(n => !Number.isNaN(n));
    }

    async function resolveQTestAssignedToFieldValue(eventData, defectToUpdate) {
        const fields = getFields(eventData);
        const adoAssignedRaw = fields["System.AssignedTo"];
        const identity = extractAdoAssignedToIdentity(adoAssignedRaw);
        if (!identity || !identity.email) {
            return undefined; // don't touch assigned-to if not present
        }

        const users = await getQTestProjectUsers();
        const desiredEmail = normalize(identity.email);
        const desiredName = normalize(identity.displayName);

        const match = users.find(u => {
            const cands = getUserCandidateStrings(u);
            if (cands.includes(desiredEmail)) return true;
            if (desiredName && cands.includes(desiredName)) return true;
            return false;
        });

        if (!match || !match.id) {
            console.log(`[Warn] Could not map ADO AssignedTo '${identity.email}' to a qTest user. Leaving qTest assignment unchanged.`);
            return undefined;
        }

        // If current assignment already matches, skip update
        const assignedFieldId = await getAssignedToDefectFieldId();
        if (assignedFieldId && defectToUpdate && defectToUpdate.properties) {
            const currentField = defectToUpdate.properties.find(p => Number(p.field_id) === Number(assignedFieldId));
            const currentIds = currentField ? parseQTestIdList(currentField.field_value) : [];
            if (currentIds.length && Number(currentIds[0]) === Number(match.id)) {
                return undefined;
            }
        }

        return `[${match.id}]`;
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
    async function getDefectByWorkItemId(workItemId) {
        const prefix = getNamePrefix(workItemId);
        const url = "https://" + constants.ManagerURL + "/api/v3/projects/" + constants.ProjectID + "/search";
        const requestBody = {
            object_type: "defects",
            fields: ["*"],
            query: "Summary ~ '" + prefix + "'",
        };

        console.log(`[Info] Get existing defect for 'WI${workItemId}'`);
        let failed = false;
        let defect = undefined;

        try {
            const response = await post(url, requestBody);
            console.log(response);

            if (!response || response.total === 0) {
                console.log("[Info] Defect not found by work item id.");
            } else {
                if (response.total === 1) {
                    defect = response.items[0];
                } else {
                    failed = true;
                    console.log("[Warn] Multiple Defects found by work item id.");
                }
            }
        } catch (error) {
            console.log("[Error] Failed to get defect by work item id.", error);
            failed = true;
        }

        return { failed: failed, defect: defect };
    }

    async function updateDefect(defectToUpdate, summary, description, assignedToFieldValue) {
        const defectId = defectToUpdate.id;
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/defects/${defectId}`;

        const properties = [
            {
                field_id: constants.DefectSummaryFieldID,
                field_value: summary,
            },
            {
                field_id: constants.DefectDescriptionFieldID,
                field_value: description,
            },
        ];

        // P1: Map Assigned To from ADO -> qTest when possible
        if (assignedToFieldValue !== undefined && assignedToFieldValue !== null) {
            const assignedFieldId = await getAssignedToDefectFieldId();
            if (assignedFieldId) {
                properties.push({
                    field_id: assignedFieldId,
                    field_value: assignedToFieldValue,
                });
            }
        }

        const requestBody = { properties };

        console.log(`[Info] Updating defect '${defectId}'.`);

        try {
            await put(url, requestBody);
            console.log(`[Info] Defect '${defectId}' updated.`);
        } catch (error) {
            console.log(`[Error] Failed to update defect '${defectId}'.`, error);
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
            data: requestBody,
            method: method,
        };

        try {
            const response = await axios(opts);
            return response.data;
        } catch (error) {
            console.log(`[Error] HTTP ${error.response.status}: ${error.response.data}`);
            throw error;
        }
    }
};
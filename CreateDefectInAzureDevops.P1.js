const { Webhooks } = require('@qasymphony/pulse-sdk');
const axios = require('axios');

exports.handler = async function ({ event, constants, triggers }, context, callback) {
    let iteration;
    if (event.iteration != undefined) {
        iteration = event.iteration;
    } else {
        iteration = 1;
    }
    const maxIterations = 4;
    const defectId = event.defect.id;
    const projectId = event.defect.project_id;
    console.log(`[Info] Create defect event received for defect '${defectId}' in project '${projectId}'`);

    if (projectId != constants.ProjectID) {
        console.log(`[Info] Project not matching '${projectId}' != '${constants.ProjectID}', exiting.`);
        return;
    }

    const defectDetails = await getDefectDetailsByIdWithRetry(defectId);
    if (!defectDetails) return;

    const assignedToEmail = await resolveAzDoAssignedToEmail(defectDetails.defect);

    const bug = await createAzDoBug(defectId, defectDetails.summary, defectDetails.description, defectDetails.link, assignedToEmail);

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

    // -------------------------
    // P1: Assigned To (qTest <-> ADO) mapping helpers
    // -------------------------
    let _qtestUsersCache = null; // [{id, ...}]
    let _defectFieldsCache = null; // [{id,label,...}]
    let _assignedToDefectFieldIdCache = null;

    async function getQTestProjectUsers() {
        if (_qtestUsersCache) return _qtestUsersCache;

        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/users`;
        try {
            const response = await axios.get(url, { headers: { Authorization: `bearer ${constants.QTEST_TOKEN}` } });
            _qtestUsersCache = Array.isArray(response.data) ? response.data : [];
            return _qtestUsersCache;
        } catch (error) {
            console.log("[Error] Failed to get qTest project users.", error);
            return [];
        }
    }

    async function getDefectFields() {
        if (_defectFieldsCache) return _defectFieldsCache;

        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/settings/defects/fields`;
        try {
            const response = await axios.get(url, { headers: { Authorization: `bearer ${constants.QTEST_TOKEN}` } });
            _defectFieldsCache = Array.isArray(response.data) ? response.data : [];
            return _defectFieldsCache;
        } catch (error) {
            console.log("[Error] Failed to get qTest defect fields.", error);
            return [];
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

    function parseQTestIdList(value) {
        if (!value) return [];
        if (Array.isArray(value)) return value.map(v => Number(v)).filter(v => !Number.isNaN(v));
        const s = value.toString().trim();
        // Common qTest format for user/multi-select fields: "[123,456]" or "[123]"
        const m = s.match(/\[([^\]]*)\]/);
        if (!m) return [];
        return m[1]
            .split(",")
            .map(x => Number((x || "").trim()))
            .filter(n => !Number.isNaN(n));
    }

    async function resolveAzDoAssignedToEmail(defect) {
        if (!defect) return null;

        const assignedFieldId = await getAssignedToDefectFieldId();
        if (!assignedFieldId) return null;

        const assignedField = getFieldById(defect, assignedFieldId);
        if (!assignedField || !assignedField.field_value) return null;

        const ids = parseQTestIdList(assignedField.field_value);
        if (!ids.length) return null;

        const users = await getQTestProjectUsers();
        const user = users.find(u => Number(u.id) === Number(ids[0]));
        if (!user) return null;

        // Try the most likely keys for email/username depending on qTest version
        return (user.email || user.username || user.user_name || user.login || "").toString() || null;
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
                console.log(`[Warn] Could not get defect details on attempt ${attempt}. Waiting ${delay} ms.`);
                await new Promise((r) => setTimeout(r, delay));
            }

            defectDetails = await getDefectDetailsById(defectId);

            if (defectDetails && defectDetails.summary && defectDetails.description) return defectDetails;

            attempt++;
        } while (attempt < 12);

        console.log(`[Error] Could not get defect details, user has not yet performed initial save in qTest, or defect was abandoned.`);
        if (iteration < maxIterations) {
            iteration = iteration + 1;
            console.log(`[Info] Re-executing with original parameters and iteration of ${iteration} of a maximum ${maxIterations}.`);
            event.iteration = iteration;
            emitEvent('qTestDefectSubmitted', event);
        } else {
            console.error(`[Error] Retry exceeded ${maxIterations} attempts, rule has timed out.`);
        }
    }

    async function getDefectDetailsById(defectId) {
        const defect = await getDefectById(defectId);

        if (!defect) return;

        const summaryField = getFieldById(defect, constants.DefectSummaryFieldID);
        const descriptionField = getFieldById(defect, constants.DefectDescriptionFieldID);

        if (!summaryField || !descriptionField) {
            console.log("[Error] Fields not found, exiting.");
        }

        const summary = summaryField.field_value;
        console.log(`[Info] Defect summary: ${summary}`);
        const description = descriptionField.field_value;
        console.log(`[Info] Defect description: ${description}`);
        const link = defect.web_url;
        console.log(`[Info] Defect link: ${link}`);

        return { summary: summary, description: description, link: link, defect: defect };
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
            console.log("[Error] Failed to get defect by id.", error);
        }
    }

    async function createAzDoBug(defectId, name, description, link, assignedToEmail) {
        console.log(`[Info] Creating bug in Azure DevOps '${defectId}'`);
        const baseUrl = encodeIfNeeded(constants.AzDoProjectURL);
        const url = `${baseUrl}/_apis/wit/workitems/$bug?api-version=6.0`;
        const requestBody = [
            {
                op: "add",
                path: "/fields/System.Title",
                value: name,
            },
            ...(assignedToEmail ? [{
                op: "add",
                path: "/fields/System.AssignedTo",
                value: assignedToEmail,
            }] : []),
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
                path: "/relations/-",
                value: {
                    rel: "Hyperlink",
                    url: link,
                },
            },
        ];
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

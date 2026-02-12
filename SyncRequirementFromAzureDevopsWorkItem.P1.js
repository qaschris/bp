const axios = require("axios");

exports.handler = async function ({ event, constants, triggers }, context, callback) {
    function buildRequirementDescription(eventData) {
        const fields = getFields(eventData);
        return `<a href="${eventData.resource._links.html.href}" target="_blank">Open in Azure DevOps</a><br>
<b>Type:</b> ${fields["System.WorkItemType"]}<br>
<b>Area:</b> ${fields["System.AreaPath"]}<br>
<b>Iteration:</b> ${fields["System.IterationPath"]}<br>
<b>State:</b> ${fields["System.State"]}<br>
<b>Reason:</b> ${fields["System.Reason"]}<br>
<b>Acceptance Criteria:</b> ${fields["Microsoft.VSTS.Common.AcceptanceCriteria"]}<br>
<b>Description:</b> ${fields["System.Description"] || ""}`;
    }

    function buildRequirementName(namePrefix, eventData) {
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
    let requirementToUpdate = undefined;
    switch (event.eventType) {
        case eventType.CREATED: {
            workItemId = event.resource.id;
            console.log(`[Info] Create workitem event received for 'WI${workItemId}'`);
            break;
        }
        case eventType.UPDATED: {
            workItemId = event.resource.workItemId;
            console.log(`[Info] Update workitem event received for 'WI${workItemId}'`);
            const getReqResult = await getRequirementByWorkItemId(workItemId);
            if (getReqResult.failed) {
                return;
            }
            if (getReqResult.requirement === undefined && !constants.AllowCreationOnUpdate) {
                console.log("[Info] Creation of Requirement on update event not enabled. Exiting.");
                return;
            }
            requirementToUpdate = getReqResult.requirement;
            break;
        }
        case eventType.DELETED: {
            workItemId = event.resource.id;
            console.log(`[Info] Delete workitem event received for 'WI${workItemId}'`);
            const getReqResult = await getRequirementByWorkItemId(workItemId);
            if (getReqResult.failed) {
                return;
            }
            if (getReqResult.requirement === undefined) {
                console.log(`[Info] Requirement not found to delete. Exiting.`);
                return;
            }
            // Delete requirement and finish
            await deleteRequirement(getReqResult.requirement);
            return;
        }
        default:
            console.log(`[Error] Unknown workitem event type '${event.eventType}' for 'WI${workitemId}'`);
            return;
    }

    // Prepare data to create/update requirement
    const namePrefix = getNamePrefix(workItemId);
    const requirementDescription = buildRequirementDescription(event);
    const requirementName = buildRequirementName(namePrefix, event);

    const targetParentModuleId = await resolveRequirementParentModuleId(event);

    if (requirementToUpdate) {
        await updateRequirement(requirementToUpdate, requirementName, requirementDescription, targetParentModuleId);
    } else {
        await createRequirement(requirementName, requirementDescription, targetParentModuleId);
    }

    function getNamePrefix(workItemId) {
        return `WI${workItemId}: `;
    }

    // -------------------------
    // P1: AreaPath + RICEFW module mapping helpers
    // -------------------------

    // Per requirements, ADO "Feature" work items should be stored under a root "RICEFW" module,
    // then under the ADO AreaPath module hierarchy.
    const _moduleChildrenCache = {}; // parentId -> [{id,name}]
    let _rootModulesCache = null; // [{id,name}]

    function splitAreaPath(areaPath) {
        if (!areaPath || typeof areaPath !== "string") return [];
        return areaPath
            .split("\\")
            .map(s => (s || "").trim())
            .filter(Boolean);
    }

    async function resolveRequirementParentModuleId(eventData) {
        const fields = getFields(eventData);
        const workItemType = (fields["System.WorkItemType"] || "").toString();

        // Determine base parent module
        let baseParentId = constants.RequirementParentID;

        // P1: Map ADO Feature work items into RICEFW folder
        if (workItemType.toLowerCase() === "feature") {
            const ricefwName = (constants.RicefwRootModuleName || "RICEFW").toString();
            const ricefwModule = await getOrCreateRootModuleByName(ricefwName);
            baseParentId = ricefwModule.id;
        }

        // P1: Map AreaPath into qTest modules
        const areaPath = fields["System.AreaPath"] || constants.DefaultAreaPath || constants.AreaPathDefault || "";
        const segments = splitAreaPath(areaPath);

        if (!segments.length) {
            return baseParentId;
        }

        return await ensureModulePath(baseParentId, segments);
    }

    async function getModules(parentId) {
        const baseUrl = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/modules`;
        const url = parentId ? `${baseUrl}?parentId=${parentId}` : baseUrl;

        try {
            const response = await axios({
                url,
                method: "GET",
                headers: standardHeaders,
            });

            // Response is typically an array of modules
            return Array.isArray(response.data) ? response.data : [];
        } catch (error) {
            console.log(`[Error] Failed to list modules (parentId='${parentId || ""}').`, error);
            throw error;
        }
    }

    async function createModule(name, parentId) {
        // qTest: Create a module. If parentId is provided, create it under that parent.
        // See qTest module APIs documentation.
        const baseUrl = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/modules`;
        const url = parentId ? `${baseUrl}?parentId=${parentId}` : baseUrl;

        const requestBody = {
            name: name,
            description: "",
        };

        try {
            const response = await axios({
                url,
                method: "POST",
                headers: standardHeaders,
                data: requestBody,
            });
            return response.data;
        } catch (error) {
            console.log(`[Error] Failed to create module '${name}' (parentId='${parentId || ""}').`, error);
            throw error;
        }
    }

    async function getOrCreateRootModuleByName(moduleName) {
        if (!_rootModulesCache) {
            const roots = await getModules(undefined);
            _rootModulesCache = roots.map(m => ({ id: m.id, name: m.name }));
        }

        let found = _rootModulesCache.find(m => (m.name || "").toString() === moduleName);
        if (found) {
            return found;
        }

        const created = await createModule(moduleName, undefined);
        const entry = { id: created.id, name: created.name };
        _rootModulesCache.push(entry);
        return entry;
    }

    async function listChildren(parentId) {
        if (_moduleChildrenCache[parentId]) {
            return _moduleChildrenCache[parentId];
        }

        const mods = await getModules(parentId);
        const children = mods.map(m => ({ id: m.id, name: m.name }));
        _moduleChildrenCache[parentId] = children;
        return children;
    }

    async function ensureChildModule(parentId, childName) {
        const children = await listChildren(parentId);
        let found = children.find(m => (m.name || "").toString() === childName);
        if (found) {
            return found.id;
        }

        const created = await createModule(childName, parentId);
        const entry = { id: created.id, name: created.name };
        children.push(entry);
        _moduleChildrenCache[parentId] = children;
        return entry.id;
    }

    async function ensureModulePath(baseParentId, segments) {
        let currentParentId = baseParentId;

        for (const seg of segments) {
            currentParentId = await ensureChildModule(currentParentId, seg);
        }

        return currentParentId;
    }


    async function getRequirementByWorkItemId(workItemId) {
        const prefix = getNamePrefix(workItemId);
        const url = "https://" + constants.ManagerURL + "/api/v3/projects/" + constants.ProjectID + "/search";
        const requestBody = {
            object_type: "requirements",
            fields: ["*"],
            query: "Name ~ '" + prefix + "'",
        };

        console.log(`[Info] Get existing requirement for 'WI${workItemId}'`);
        let failed = false;
        let requirement = undefined;

        try {
            const response = await post(url, requestBody);
            console.log(response);

            if (!response || response.total === 0) {
                console.log("[Info] Requirement not found by work item id.");
            } else {
                if (response.total === 1) {
                    requirement = response.items[0];
                } else {
                    failed = true;
                    console.log("[Warn] Multiple Requirements found by work item id.");
                }
            }
        } catch (error) {
            console.log("[Error] Failed to get requirement by work item id.", error);
            failed = true;
        }

        return { failed: failed, requirement: requirement };
    }

    async function updateRequirement(requirementToUpdate, name, description, parentModuleId) {
        const requirementId = requirementToUpdate.id;

        // qTest supports moving a requirement to another module via the update endpoint (see API docs).
        // For qTest 6+, we pass parentId as a request parameter; for older versions we also include parent_id in the body.
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/requirements/${requirementId}?parentId=${parentModuleId}`;
        const requestBody = {
            name: name,
            parent_id: parentModuleId,
            properties: [
                {
                    field_id: constants.RequirementDescriptionFieldID,
                    field_value: description,
                },
            ],
        };

        console.log(`[Info] Updating requirement '${requirementId}'.`);

        try {
            await put(url, requestBody);
            console.log(`[Info] Requirement '${requirementId}' updated.`);
        } catch (error) {
            console.log(`[Error] Failed to update requirement '${requirementId}'.`, error);
        }
    }

    async function createRequirement(name, description, parentModuleId) {
        // For qTest 6+, specify parentId as a request parameter; for older versions include parent_id in the body.
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/requirements?parentId=${parentModuleId}`;
        const requestBody = {
            name: name,
            parent_id: parentModuleId,
            properties: [
                {
                    field_id: constants.RequirementDescriptionFieldID,
                    field_value: description,
                },
            ],
        };

        console.log(`[Info] Creating requirement.`);

        try {
            await post(url, requestBody);
            console.log(`[Info] Requirement created.`);
        } catch (error) {
            console.log(`[Error] Failed to create requirement.`, error);
        }
    }

    async function deleteRequirement(requirementToDelete) {
        const requirementId = requirementToDelete.id;
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/requirements/${requirementId}`;

        console.log(`[Info] Deleting requirement '${requirementId}'.`);

        try {
            await doRequest(url, "DELETE", null);
            console.log(`[Info] Requirement '${requirementId}' deleted.`);
        } catch (error) {
            console.log(`[Error] Failed to delete requirement '${requirementId}'.`, error);
        }
    }

    function post(url, requestBody) {
        return doRequest(url, "POST", requestBody);
    }

    function put(url, requestBody) {
        return doRequest(url, "PUT", requestBody);
    }

    async function doRequest(url, method, requestBody) {
        const opts = {
            url: url,
            method: method,
            headers: standardHeaders,
            data: requestBody,
        };

        try {
            const response = await axios(opts);
            return response.data;
        } catch (error) {
            throw new Error(`Failed to ${method} ${url}. ${error.message}`);
        }
    }
};

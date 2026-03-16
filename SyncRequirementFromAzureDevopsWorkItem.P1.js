const axios = require("axios");

exports.handler = async function ({ event, constants }, context, callback) {
    // --- Helper functions ---
    function getFitGap(eventData) {
        // UPDATE event → delta fields
        if (eventData.eventType === "workitem.updated") {
            const delta = eventData.resource?.fields?.["BP.ERP.FitGap"];
            return delta?.newValue || null;
        }

        // CREATE event → full revision fields
        const fields = getFields(eventData);
        return fields["BP.ERP.FitGap"] || null;
    }
    function getBPEntity(eventData) {
        // UPDATE event → delta fields
        if (eventData.eventType === "workitem.updated") {
            const delta = eventData.resource?.fields?.["Custom.Entity"];
            return delta?.newValue ?? null;
        }

        // CREATE event → full revision fields
        const fields = getFields(eventData);
        return fields["Custom.Entity"] ?? null;
    }




    function getFields(eventData) {
        // In case of update the fields can be taken from the revision, in case of create from the resource directly
        return eventData.resource.revision ? eventData.resource.revision.fields : eventData.resource.fields;
    }

    function buildRequirementDescription(eventData) {
        const fields = getFields(eventData);
        return `<a href="${eventData.resource._links.html.href}" target="_blank">Open in Azure DevOps</a>

<b>Type:</b> ${fields["System.WorkItemType"]}

<b>Area:</b> ${fields["System.AreaPath"]}

<b>Iteration:</b> ${fields["System.IterationPath"]}

<b>State:</b> ${fields["System.State"]}

<b>Reason:</b> ${fields["System.Reason"]}

<b>Complexity:</b> ${fields["Custom.Complexity"] || ""}
 
<b>Acceptance Criteria:</b> ${fields["Microsoft.VSTS.Common.AcceptanceCriteria"]}

<b>Description:</b> ${fields["System.Description"] || ""}`;
    }

    function buildRequirementName(namePrefix, eventData) {
        const fields = getFields(eventData);
        return `${namePrefix}${fields["System.Title"]}`;
    }

    function getNamePrefix(workItemId) {
        return `WI${workItemId}: `;
    }

    // --- HTTP helpers ---
    const standardHeaders = {
        "Content-Type": "application/json",
        Authorization: `bearer ${constants.QTEST_TOKEN}`,
    };

    async function doRequest(url, method, requestBody) {
        const opts = {
            url,
            method,
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

    function post(url, requestBody) {
        return doRequest(url, "POST", requestBody);
    }

    function put(url, requestBody) {
        return doRequest(url, "PUT", requestBody);
    }

    const moduleChildrenCache = {};

    function normalizeAreaPathSegments(areaPath) {
        if (!areaPath) return [];
        return String(areaPath)
            .split(/[\\/]+/)
            .map(s => s.trim())
            .filter(Boolean);
    }

    async function getSubModules(parentId) {
        const cacheKey = String(parentId);
        if (moduleChildrenCache[cacheKey]) {
            return moduleChildrenCache[cacheKey];
        }

        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/modules/${parentId}?expand=descendants`;

        try {
            const response = await doRequest(url, "GET", null);
            const items = Array.isArray(response)
                ? response
                : Array.isArray(response?.items)
                    ? response.items
                    : Array.isArray(response?.data)
                        ? response.data
                        : [];

            moduleChildrenCache[cacheKey] = items;
            return items;
        } catch (error) {
            console.log(`[Error] Failed to get sub-modules for parent '${parentId}'.`, error);
            throw error;
        }
    }

    async function createModule(name, parentId) {
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/modules`;
        const requestBody = {
            name,
            parent_id: parentId,
        };

        try {
            const created = await post(url, requestBody);
            delete moduleChildrenCache[String(parentId)];
            console.log(`[Info] Created qTest module '${name}' under parent '${parentId}'.`);
            return created;
        } catch (error) {
            console.log(`[Error] Failed to create module '${name}' under parent '${parentId}'.`, error);
            throw error;
        }
    }

    async function ensureModulePath(areaPath) {
        const segments = normalizeAreaPathSegments(areaPath);
        let currentParentId = constants.RequirementParentID;

        if (!segments.length) {
            console.log(`[Info] AreaPath missing or blank. Using RequirementParentID '${currentParentId}'.`);
            return currentParentId;
        }

        console.log(`[Info] Resolving qTest module path from ADO AreaPath '${areaPath}'.`);

        for (const segment of segments) {
            const children = await getSubModules(currentParentId);
            const existing = children.find(m =>
                ((m?.name || "").trim().toLowerCase() === segment.toLowerCase())
            );

            if (existing) {
                currentParentId = existing.id;
                console.log(`[Info] Reusing qTest module '${segment}' (id: ${currentParentId}).`);
            } else {
                const created = await createModule(segment, currentParentId);
                currentParentId = created?.id;
                if (!currentParentId) {
                    throw new Error(`Module creation for '${segment}' did not return an id.`);
                }
            }
        }

        console.log(`[Info] Resolved qTest target module id '${currentParentId}' for AreaPath '${areaPath}'.`);
        return currentParentId;
    }

    // --- Requirement CRUD ---
    async function getRequirementByWorkItemId(workItemId) {
        const prefix = getNamePrefix(workItemId);
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/search`;
        const requestBody = {
            object_type: "requirements",
            fields: ["*"],
            query: `Name ~ '${prefix}'`,
        };

        let failed = false;
        let requirement = undefined;

        try {
            const response = await post(url, requestBody);
            if (!response || response.total === 0) {
                console.log("[Info] Requirement not found by work item id.");
            } else if (response.total === 1) {
                requirement = response.items[0];
            } else {
                failed = true;
                console.log("[Warn] Multiple Requirements found by work item id.");
            }
        } catch (error) {
            console.log("[Error] Failed to get requirement by work item id.", error);
            failed = true;
        }

        return { failed, requirement };
    }

    async function updateRequirement(requirementToUpdate, name, description, areaPath, complexityValue, assignedTo, iterationPath, applicationName, fitGap, bpEntity, targetModuleId) {        const requirementId = requirementToUpdate.id;
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/requirements/${requirementId}?parentId=${targetModuleId}`;        const requestBody = {
            name,
            properties: [
                { field_id: constants.RequirementDescriptionFieldID, field_value: description },
                { field_id: constants.RequirementStreamSquadFieldID, field_value: areaPath },
            ],
        };

        if (complexityValue) {
            requestBody.properties.push({
                field_id: constants.RequirementComplexityFieldID,
                field_value: complexityValue,
            });
        }

        if (assignedTo) {

            let removed_email = assignedTo.replace(/\s*<[^>]*>/g, '');


            requestBody.properties.push({
                field_id: constants.RequirementAssignedToFieldID,
                field_value: removed_email,
            });

        }
        if (iterationPath) {
            requestBody.properties.push({
                field_id: constants.RequirementIterationPathFieldID,
                field_value: iterationPath,
            });
        }
        if (applicationName) {
            requestBody.properties.push({
                field_id: constants.RequirementApplicationNameFieldID,
                field_value: applicationName,
            });
            if (fitGap !== null && fitGap !== undefined) {
                requestBody.properties.push({
                    field_id: constants.RequirementFitGapFieldID,
                    field_value: fitGap
                });
            }
            if (bpEntity !== null && bpEntity !== undefined) {
                requestBody.properties.push({
                    field_id: constants.RequirementBPEntityFieldID,
                    field_value: bpEntity
                });
            }
        }

        try {
            await put(url, requestBody);
            console.log(`[Info] Requirement '${requirementId}' updated.`);
        } catch (error) {
            console.log(`[Error] Failed to update requirement '${requirementId}'.`, error);
        }
    }

    async function createRequirement(name, description, areaPath, complexityValue, assignedTo, iterationPath, applicationName, fitGap, bpEntity, targetModuleId) {
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/requirements`;
        const requestBody = {
            name,
            parent_id: targetModuleId,
            properties: [
                { field_id: constants.RequirementDescriptionFieldID, field_value: description },
                { field_id: constants.RequirementStreamSquadFieldID, field_value: areaPath },
            ],
        };

        if (complexityValue) {
            requestBody.properties.push({
                field_id: constants.RequirementComplexityFieldID,
                field_value: complexityValue,
            });
        }

        if (assignedTo) {
            let removed_email = assignedTo.replace(/\s*<[^>]*>/g, '');
            requestBody.properties.push({
                field_id: constants.RequirementAssignedToFieldID,
                field_value: removed_email,
            });

        }
        if (iterationPath) {
            requestBody.properties.push({
                field_id: constants.RequirementIterationPathFieldID,
                field_value: iterationPath,
            });
        }
        if (applicationName) {
            requestBody.properties.push({
                field_id: constants.RequirementApplicationNameFieldID,
                field_value: applicationName,
            });
        }
        if (fitGap !== null && fitGap !== undefined) {
            requestBody.properties.push({
                field_id: constants.RequirementFitGapFieldID,
                field_value: fitGap
            });
        }
        if (bpEntity !== null && bpEntity !== undefined) {
            requestBody.properties.push({
                field_id: constants.RequirementBPEntityFieldID,
                field_value: bpEntity
            });
        }

        try {
            await post(url, requestBody);
            console.log(`[Info] Requirement created.`);
        } catch (error) {
            console.log(`[Error] Failed to create requirement`, error);
        }
    }

    async function deleteRequirement(requirementToDelete) {
        const requirementId = requirementToDelete.id;
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/requirements/${requirementId}`;
        try {
            await doRequest(url, "DELETE", null);
            console.log(`[Info] Requirement '${requirementId}' deleted.`);
        } catch (error) {
            console.log(`[Error] Failed to delete requirement '${requirementId}'.`, error);
        }
    }

    // --- Main Handler Logic ---
    const eventType = {
        CREATED: "workitem.created",
        UPDATED: "workitem.updated",
        DELETED: "workitem.deleted",
    };

    let workItemId;
    let requirementToUpdate;

    switch (event.eventType) {
        case eventType.CREATED:
            workItemId = event.resource.id;
            console.log(`[Info] Create workitem event received for 'WI${workItemId}'`);
            break;

        case eventType.UPDATED:
            workItemId = event.resource.workItemId;
            console.log(`[Info] Update workitem event received for 'WI${workItemId}'`);
            const getReqResult = await getRequirementByWorkItemId(workItemId);
            if (getReqResult.failed) return;
            if (!getReqResult.requirement && !constants.AllowCreationOnUpdate) {
                console.log("[Info] Creation of Requirement on update event not enabled. Exiting.");
                return;
            }
            requirementToUpdate = getReqResult.requirement;
            break;

        case eventType.DELETED:
            workItemId = event.resource.id;
            console.log(`[Info] Delete workitem event received for 'WI${workItemId}'`);
            const getReq = await getRequirementByWorkItemId(workItemId);
            if (getReq.failed || !getReq.requirement) return;
            await deleteRequirement(getReq.requirement);
            return;

        default:
            console.log(`[Error] Unknown workitem event type '${event.eventType}' for 'WI${workItemId}'`);
            return;
    }

    // Prepare data
    const namePrefix = getNamePrefix(workItemId);
    const requirementDescription = buildRequirementDescription(event);
    const requirementName = buildRequirementName(namePrefix, event);

    const fields = getFields(event);
    const adoAreaPath = fields["System.AreaPath"];
    const adoComplexity = fields["Custom.Complexity"];
    const qtestComplexityValue = {
        "1 - Very High": 941,
        "2 - High": 942,
        "3 - Medium": 943,
        "4 - Low": 944,
        "5 - Very Low": 945,
    }[adoComplexity] || null;

    const adoAssignedTo = fields["System.AssignedTo"];
    const iterationPath = fields["System.IterationPath"] || null;
    const applicationName = fields["Custom.ApplicationName"] || null;
    const fitGap = getFitGap(event);
    const bpEntity = getBPEntity(event);
    const targetModuleId = await ensureModulePath(adoAreaPath);


    if (requirementToUpdate) {
        await updateRequirement(requirementToUpdate, requirementName, requirementDescription, adoAreaPath, qtestComplexityValue, adoAssignedTo, iterationPath, applicationName, fitGap, bpEntity, targetModuleId);
    } else {
        await createRequirement(requirementName, requirementDescription, adoAreaPath, qtestComplexityValue, adoAssignedTo, iterationPath, applicationName, fitGap, bpEntity, targetModuleId);
    }
};
 
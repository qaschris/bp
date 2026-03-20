const axios = require("axios");

exports.handler = async function ({ event, constants }, context, callback) {
    // --- Helper functions ---
    function getFitGap(eventData) {
        if (eventData.eventType === "workitem.updated") {
            const delta = eventData.resource?.fields?.["BP.ERP.FitGap"];
            if (delta && Object.prototype.hasOwnProperty.call(delta, "newValue")) {
                return delta.newValue ?? null;
            }
        }

        const fields = getFields(eventData);
        return fields["BP.ERP.FitGap"] ?? null;
    }

    function getBPEntity(eventData) {
        if (eventData.eventType === "workitem.updated") {
            const delta = eventData.resource?.fields?.["Custom.Entity"];
            if (delta && Object.prototype.hasOwnProperty.call(delta, "newValue")) {
                return delta.newValue ?? null;
            }
        }

        const fields = getFields(eventData);
        return fields["Custom.Entity"] ?? null;
    }

    function getFields(eventData) {
        // In case of update the fields can be taken from the revision, in case of create from the resource directly
        return eventData.resource.revision ? eventData.resource.revision.fields : eventData.resource.fields;
    }

    function buildRequirementDescription(eventData) {
        const fields = getFields(eventData);

        const workItemType = fields["System.WorkItemType"] || "";
        const areaPath = fields["System.AreaPath"] || "";
        const iterationPath = fields["System.IterationPath"] || "";
        const state = fields["System.State"] || "";
        const reason = fields["System.Reason"] || "";
        const complexity = fields["Custom.Complexity"] || "";
        const acceptanceCriteria = fields["Microsoft.VSTS.Common.AcceptanceCriteria"] || "";
        const description = fields["System.Description"] || "";

        return `<a href="${eventData.resource._links.html.href}" target="_blank">Open in Azure DevOps</a><br>
    <b>Type:</b> ${workItemType}<br>
    <b>Area:</b> ${areaPath}<br>
    <b>Iteration:</b> ${iterationPath}<br>
    <b>State:</b> ${state}<br>
    <b>Reason:</b> ${reason}<br>
    <b>Complexity:</b> ${complexity}<br>
    <b>Acceptance Criteria:</b> ${acceptanceCriteria}<br>
    <b>Description:</b> ${description}`;
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

    function safeJson(value) {
        try {
            return JSON.stringify(value, null, 2);
        } catch (e) {
            return `[Unserializable: ${e.message}]`;
        }
    }

    function logDivider(title) {
        console.log(`==================== ${title} ====================`);
    }

    function sanitizeHeadersForLog(headers) {
        const clone = { ...(headers || {}) };
        if (clone.Authorization) {
            clone.Authorization = "[REDACTED]";
        }
        return clone;
    }

    async function doRequest(url, method, requestBody) {
        const opts = {
            url,
            method,
            headers: standardHeaders,
        };

        if (requestBody !== undefined && requestBody !== null && method !== "GET") {
            opts.data = requestBody;
        }

        logDivider(`HTTP ${method}`);
        console.log(`[Debug] URL: ${url}`);
        console.log(`[Debug] Headers: ${safeJson(sanitizeHeadersForLog(standardHeaders))}`);
        if (opts.data !== undefined) {
            console.log(`[Debug] Request Payload: ${safeJson(opts.data)}`);
        } else {
            console.log(`[Debug] Request Payload: <none>`);
        }

        try {
            const response = await axios(opts);

            console.log(`[Debug] HTTP Status: ${response.status}`);
            console.log(`[Debug] Response Headers: ${safeJson(response.headers)}`);
            console.log(`[Debug] Response Body: ${safeJson(response.data)}`);

            return response.data;
        } catch (error) {
            console.log(`[Error] HTTP request failed.`);
            console.log(`[Error] URL: ${url}`);
            console.log(`[Error] Method: ${method}`);
            console.log(`[Error] Message: ${error.message}`);

            if (error.response) {
                console.log(`[Error] HTTP Status: ${error.response.status}`);
                console.log(`[Error] Error Response Headers: ${safeJson(error.response.headers)}`);
                console.log(`[Error] Error Response Body: ${safeJson(error.response.data)}`);
            } else if (error.request) {
                console.log(`[Error] No HTTP response received.`);
                console.log(`[Error] Raw Request Object: ${safeJson(error.request)}`);
            } else {
                console.log(`[Error] Axios config/setup error: ${safeJson(error)}`);
            }

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

    function getReleaseFolderName(iterationPath) {
        if (!iterationPath) {
            return "TBD";
        }

        const segments = String(iterationPath)
            .split(/[\\/]+/)
            .map(s => s.trim())
            .filter(Boolean);

        // Usually the release indicator is in the second node, e.g.
        // bp_Quantum\P_O R1.0 (Castellon and Kwinana)
        // bp_Quantum\P_O R1.1 Whiting
        const candidate = segments.length > 1 ? segments[1] : segments[0];

        if (!candidate) {
            return "TBD";
        }

        const match = candidate.match(/P_O\s+R(\d+(?:\.\d+)?)/i);
        if (match && match[1]) {
            return `P&O Release ${match[1]}`;
        }

        return "TBD";
    }

    async function getSubModules(parentId) {
        const cacheKey = String(parentId);
        if (moduleChildrenCache[cacheKey]) {
            console.log(`[Debug] getSubModules cache hit for parent '${parentId}'.`);
            console.log(`[Debug] Cached children: ${moduleChildrenCache[cacheKey].map(c => `${c.name} (${c.id})`).join(", ")}`);
            return moduleChildrenCache[cacheKey];
        }

        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/modules/${parentId}?expand=descendants`;

        try {
            const response = await doRequest(url, "GET", null);

            let items = [];
            if (Array.isArray(response?.children)) {
                items = response.children;
            } else if (Array.isArray(response)) {
                items = response;
            } else if (Array.isArray(response?.items)) {
                items = response.items;
            } else if (Array.isArray(response?.data)) {
                items = response.data;
            }

            console.log(`[Debug] getSubModules parent '${parentId}' resolved ${items.length} immediate children.`);
            console.log(`[Debug] Children: ${items.map(c => `${c.name} (${c.id})`).join(", ")}`);

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

    async function ensureModulePath(areaPath, iterationPath) {
        const releaseFolderName = getReleaseFolderName(iterationPath);

        let areaSegments = normalizeAreaPathSegments(areaPath);

        // Drop the first AreaPath node if it is the project root (bp_Quantum),
        // because the release folder replaces that top-level bucket.
        if (areaSegments.length && areaSegments[0].toLowerCase() === "bp_quantum") {
            areaSegments = areaSegments.slice(1);
        }

        const segments = [releaseFolderName, ...areaSegments];
        let currentParentId = constants.RequirementParentID;

        if (!segments.length) {
            console.log(`[Info] No module segments resolved. Using RequirementParentID '${currentParentId}'.`);
            return currentParentId;
        }

        console.log(`[Info] Resolving qTest module path from IterationPath '${iterationPath}' and AreaPath '${areaPath}'.`);
        console.log(`[Info] Derived release folder: '${releaseFolderName}'.`);
        console.log(`[Info] Final module path segments: ${safeJson(segments)}`);

        for (const segment of segments) {
            const children = await getSubModules(currentParentId);
            console.log(`[Debug] Looking for segment '${segment}' under parent '${currentParentId}'.`);
            console.log(`[Debug] Children of ${currentParentId}: ${children.map(c => `${c.name} (${c.id})`).join(", ")}`);

            const existing = children.find(m =>
                ((m?.name || "").trim().toLowerCase() === segment.toLowerCase())
            );

            if (existing) {
                console.log(`[Debug] Found existing module match for '${segment}': ${safeJson(existing)}`);
                currentParentId = existing.id;
                console.log(`[Info] Reusing qTest module '${segment}' (id: ${currentParentId}).`);
            } else {
                console.log(`[Debug] No existing module found for '${segment}' under parent '${currentParentId}'. Will create.`);
                const created = await createModule(segment, currentParentId);
                currentParentId = created?.id;
                if (!currentParentId) {
                    throw new Error(`Module creation for '${segment}' did not return an id.`);
                }
            }
        }

        console.log(`[Info] Resolved qTest target module id '${currentParentId}' for release '${releaseFolderName}' and AreaPath '${areaPath}'.`);
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
            logDivider("SEARCH REQUIREMENT BY WORK ITEM ID");
            console.log(`[Debug] Search Prefix: ${prefix}`);
            console.log(`[Debug] Search Payload: ${safeJson(requestBody)}`);
            const response = await post(url, requestBody);
            console.log(`[Debug] Search Response: ${safeJson(response)}`);
            if (!response || response.total === 0) {
                console.log("[Info] Requirement not found by work item id.");
            } else if (response.total === 1) {
                console.log(`[Debug] Requirement matched for update: ${safeJson(requirement)}`);
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
        console.log(`[Debug] Raw AssignedTo before normalization: ${safeJson(assignedTo)}`);
        if (assignedTo) {
            let removed_email = assignedTo.replace(/\s*<[^>]*>/g, '');
            console.log(`[Debug] Normalized AssignedTo for qTest field: ${removed_email}`);
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
            logDivider("UPDATE REQUIREMENT");
            console.log(`[Debug] Requirement ID: ${requirementId}`);
            console.log(`[Debug] Target Module ID: ${targetModuleId}`);
            console.log(`[Debug] Update URL: ${url}`);
            console.log(`[Debug] Final Update Payload: ${safeJson(requestBody)}`);
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
        console.log(`[Debug] Raw AssignedTo before normalization: ${safeJson(assignedTo)}`);
        if (assignedTo) {
            let removed_email = assignedTo.replace(/\s*<[^>]*>/g, '');
            console.log(`[Debug] Normalized AssignedTo for qTest field: ${removed_email}`);
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
            logDivider("CREATE REQUIREMENT");
            console.log(`[Debug] Target Module ID: ${targetModuleId}`);
            console.log(`[Debug] Create URL: ${url}`);
            console.log(`[Debug] Final Create Payload: ${safeJson(requestBody)}`);
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

    logDivider("RAW INCOMING EVENT");
    console.log(safeJson(event));

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

    logDivider("BUILT REQUIREMENT CONTENT");
    console.log(`[Debug] Requirement Name: ${requirementName}`);
    console.log(`[Debug] Requirement Description: ${requirementDescription}`);

    const fields = getFields(event);

    logDivider("EXTRACTED ADO FIELDS");
    console.log(`[Debug] Event Type: ${event.eventType}`);
    console.log(`[Debug] Work Item Type: ${fields["System.WorkItemType"]}`);
    console.log(`[Debug] Title: ${fields["System.Title"]}`);
    console.log(`[Debug] AreaPath: ${fields["System.AreaPath"]}`);
    console.log(`[Debug] IterationPath: ${fields["System.IterationPath"]}`);
    console.log(`[Debug] State: ${fields["System.State"]}`);
    console.log(`[Debug] Reason: ${fields["System.Reason"]}`);
    console.log(`[Debug] Complexity: ${fields["Custom.Complexity"]}`);
    console.log(`[Debug] AssignedTo Raw: ${safeJson(fields["System.AssignedTo"])}`);
    console.log(`[Debug] ApplicationName: ${fields["Custom.ApplicationName"]}`);
    console.log(`[Debug] FitGap: ${safeJson(getFitGap(event))}`);
    console.log(`[Debug] BPEntity: ${safeJson(getBPEntity(event))}`);

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
    const targetModuleId = await ensureModulePath(adoAreaPath, iterationPath);

    if (requirementToUpdate) {
        await updateRequirement(requirementToUpdate, requirementName, requirementDescription, adoAreaPath, qtestComplexityValue, adoAssignedTo, iterationPath, applicationName, fitGap, bpEntity, targetModuleId);
    } else {
        await createRequirement(requirementName, requirementDescription, adoAreaPath, qtestComplexityValue, adoAssignedTo, iterationPath, applicationName, fitGap, bpEntity, targetModuleId);
    }
};
 
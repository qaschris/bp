const axios = require("axios");

exports.handler = async function ({ event, constants }, context, callback) {
    // --- Helper functions ---
    function getFields(eventData) {
        return eventData.resource.revision ? eventData.resource.revision.fields : eventData.resource.fields;
    }

    function getDeltaOrField(eventData, fieldRef) {
        if (eventData.eventType === "workitem.updated") {
            const delta = eventData.resource?.fields?.[fieldRef];
            if (delta && Object.prototype.hasOwnProperty.call(delta, "newValue")) {
                return delta.newValue;
            }
        }
        const fields = getFields(eventData);
        return fields[fieldRef];
    }

    function getNamePrefix(workItemId) {
        return `WI${workItemId}: `;
    }

    function normalizeAssignedTo(value) {
        if (!value) return "";
        if (typeof value === "string") {
            const match = value.match(/<([^>]+)>/);
            return match && match[1] ? match[1].trim() : value.replace(/\s*<[^>]*>/g, "").trim();
        }
        if (typeof value === "object") {
            return (
                value.uniqueName ||
                value.mail ||
                value.email ||
                value.userPrincipalName ||
                value.displayName ||
                ""
            ).toString().trim();
        }
        return "";
    }

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

    function isRicefwFeature(eventData) {
        const fields = getFields(eventData);
        const workItemType = (fields["System.WorkItemType"] || "").toString().trim();
        const title = (fields["System.Title"] || "").toString();
        return workItemType.toLowerCase() === "feature" && /\[RICEFW\]/i.test(title);
    }

    function buildRequirementName(namePrefix, eventData) {
        const fields = getFields(eventData);
        return `${namePrefix}${fields["System.Title"]}`;
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

        const featureType = fields[constants.AzDoRicefwFeatureTypeFieldRef || "Custom.FeatureType"] || "";
        const processRelease = fields[constants.AzDoRicefwProcessReleaseFieldRef || "Custom.ProcessRelease"] || "";
        const ricefwId = fields[constants.AzDoRicefwIdFieldRef || "Custom.RICEFWID"] || "";
        const ricefwConfiguration = fields[constants.AzDoRicefwConfigurationFieldRef || "Custom.RICEFWConfiguration"] || "";
        const testingStatus = fields[constants.AzDoRicefwTestingStatusFieldRef || "Custom.TestingStatus"] || "";
        const area = fields[constants.AzDoRicefwAreaFieldRef || "Custom.Area"] || "";

        return `<a href="${eventData.resource._links.html.href}" target="_blank">Open in Azure DevOps</a><br>
                <b>Type:</b> ${workItemType}<br>
                <b>Feature Type:</b> ${featureType}<br>
                <b>RICEFW ID:</b> ${ricefwId}<br>
                <b>Area:</b> ${area}<br>
                <b>Area Path:</b> ${areaPath}<br>
                <b>Iteration:</b> ${iterationPath}<br>
                <b>State:</b> ${state}<br>
                <b>Reason:</b> ${reason}<br>
                <b>Testing Status:</b> ${testingStatus}<br>
                <b>Complexity:</b> ${complexity}<br>
                <b>Process Release:</b> ${processRelease}<br>
                <b>RICEFW / Configuration:</b> ${ricefwConfiguration}<br>
                <b>Assigned To:</b> ${normalizeAssignedTo(fields["System.AssignedTo"])}<br>
                <b>Acceptance Criteria:</b> ${acceptanceCriteria}<br>
                <b>Description:</b> ${description}`;
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

    // --- qTest module helpers ---
    const moduleChildrenCache = {};

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

        const created = await post(url, requestBody);
        delete moduleChildrenCache[String(parentId)];
        console.log(`[Info] Created qTest module '${name}' under parent '${parentId}'.`);
        return created;
    }

    async function ensureModulePath(rootModuleId, areaPath, iterationPath) {
        const releaseFolderName = getReleaseFolderName(iterationPath);

        let areaSegments = normalizeAreaPathSegments(areaPath);

        if (areaSegments.length && areaSegments[0].toLowerCase() === "bp_quantum") {
            areaSegments = areaSegments.slice(1);
        }

        const segments = [releaseFolderName, ...areaSegments];
        let currentParentId = rootModuleId;

        if (!segments.length) {
            console.log(`[Info] No module segments resolved. Using root module '${currentParentId}'.`);
            return currentParentId;
        }

        console.log(`[Info] Resolving qTest module path from IterationPath '${iterationPath}' and AreaPath '${areaPath}'.`);
        console.log(`[Info] Derived release folder: '${releaseFolderName}'.`);

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

        return currentParentId;
    }

    async function ensureRicefwRootModule() {
        if (constants.RicefwRootModuleID) {
            return constants.RicefwRootModuleID;
        }

        if (!constants.RequirementParentID) {
            throw new Error("Either RicefwRootModuleID or RequirementParentID must be provided.");
        }

        const rootName = constants.RicefwRootModuleName || "RICEFW";
        const children = await getSubModules(constants.RequirementParentID);
        const existing = children.find(m => ((m?.name || "").trim().toLowerCase() === rootName.toLowerCase()));
        if (existing) {
            console.log(`[Info] Reusing RICEFW root module '${rootName}' (id: ${existing.id}).`);
            return existing.id;
        }

        const created = await createModule(rootName, constants.RequirementParentID);
        const rootId = created?.id;
        if (!rootId) {
            throw new Error(`RICEFW root module '${rootName}' did not return an id.`);
        }
        return rootId;
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

    async function updateRequirement(requirementToUpdate, name, description, fields, complexityValue, workItemTypeValue, priorityValue, typeValue, targetModuleId) {
        const requirementId = requirementToUpdate.id;
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/requirements/${requirementId}?parentId=${targetModuleId}`;
        const requestBody = {
            name,
            properties: [],
        };

        const areaPath = fields["System.AreaPath"] || "";
        const assignedTo = normalizeAssignedTo(fields["System.AssignedTo"]);
        const iterationPath = fields["System.IterationPath"] || null;
        const state = fields["System.State"] || null;
        const reason = fields["System.Reason"] || null;
        const acceptanceCriteria = fields["Microsoft.VSTS.Common.AcceptanceCriteria"] || "";
        const fullDescription = fields["System.Description"] || "";
        const applicationName = fields[constants.AzDoApplicationNameFieldRef || "Custom.ApplicationName"] || null;
        const fitGap = fields[constants.AzDoFitGapFieldRef || "BP.ERP.FitGap"] ?? null;
        const bpEntity = fields[constants.AzDoBPEntityFieldRef || "Custom.Entity"] ?? null;
        const processRelease = fields[constants.AzDoRicefwProcessReleaseFieldRef || "Custom.ProcessRelease"] || null;
        const ricefwId = fields[constants.AzDoRicefwIdFieldRef || "Custom.RICEFWID"] || null;
        const ricefwConfiguration = fields[constants.AzDoRicefwConfigurationFieldRef || "Custom.RICEFWConfiguration"] || null;
        const testingStatus = fields[constants.AzDoRicefwTestingStatusFieldRef || "Custom.TestingStatus"] || null;
        const featureType = fields[constants.AzDoRicefwFeatureTypeFieldRef || "Custom.FeatureType"] || null;
        const area = fields[constants.AzDoRicefwAreaFieldRef || "Custom.Area"] || null;

        requestBody.properties.push({ field_id: constants.RequirementDescriptionFieldID, field_value: description });
        requestBody.properties.push({ field_id: constants.RequirementStreamSquadFieldID, field_value: areaPath });

        if (constants.RequirementComplexityFieldID && complexityValue) {
            requestBody.properties.push({ field_id: constants.RequirementComplexityFieldID, field_value: complexityValue });
        }
        if (constants.RequirementWorkItemTypeFieldID && workItemTypeValue) {
            requestBody.properties.push({ field_id: constants.RequirementWorkItemTypeFieldID, field_value: workItemTypeValue });
        }
        if (constants.RequirementPriorityFieldID && priorityValue) {
            requestBody.properties.push({ field_id: constants.RequirementPriorityFieldID, field_value: priorityValue });
        }
        if (constants.RequirementTypeFieldID && typeValue) {
            requestBody.properties.push({ field_id: constants.RequirementTypeFieldID, field_value: typeValue });
        }
        if (constants.RequirementAssignedToFieldID && assignedTo) {
            requestBody.properties.push({ field_id: constants.RequirementAssignedToFieldID, field_value: assignedTo });
        }
        if (constants.RequirementIterationPathFieldID && iterationPath) {
            requestBody.properties.push({ field_id: constants.RequirementIterationPathFieldID, field_value: iterationPath });
        }
        if (constants.RequirementApplicationNameFieldID && applicationName) {
            requestBody.properties.push({ field_id: constants.RequirementApplicationNameFieldID, field_value: applicationName });
        }
        if (constants.RequirementFitGapFieldID && fitGap !== null && fitGap !== undefined) {
            requestBody.properties.push({ field_id: constants.RequirementFitGapFieldID, field_value: fitGap });
        }
        if (constants.RequirementBPEntityFieldID && bpEntity !== null && bpEntity !== undefined) {
            requestBody.properties.push({ field_id: constants.RequirementBPEntityFieldID, field_value: bpEntity });
        }

        // RICEFW-specific fields (optional constants)
        if (constants.RequirementStateFieldID && state) {
            requestBody.properties.push({ field_id: constants.RequirementStateFieldID, field_value: state });
        }
        if (constants.RequirementReasonFieldID && reason) {
            requestBody.properties.push({ field_id: constants.RequirementReasonFieldID, field_value: reason });
        }
        if (constants.RequirementAcceptanceCriteriaFieldID) {
            requestBody.properties.push({ field_id: constants.RequirementAcceptanceCriteriaFieldID, field_value: acceptanceCriteria });
        }
        if (constants.RequirementPlainDescriptionFieldID) {
            requestBody.properties.push({ field_id: constants.RequirementPlainDescriptionFieldID, field_value: fullDescription });
        }
        if (constants.RequirementProcessReleaseFieldID && processRelease) {
            requestBody.properties.push({ field_id: constants.RequirementProcessReleaseFieldID, field_value: processRelease });
        }
        if (constants.RequirementRicefwIdFieldID && ricefwId) {
            requestBody.properties.push({ field_id: constants.RequirementRicefwIdFieldID, field_value: ricefwId });
        }
        if (constants.RequirementRicefwConfigurationFieldID && ricefwConfiguration) {
            requestBody.properties.push({ field_id: constants.RequirementRicefwConfigurationFieldID, field_value: ricefwConfiguration });
        }
        if (constants.RequirementTestingStatusFieldID && testingStatus) {
            requestBody.properties.push({ field_id: constants.RequirementTestingStatusFieldID, field_value: testingStatus });
        }
        if (constants.RequirementFeatureTypeFieldID && featureType) {
            requestBody.properties.push({ field_id: constants.RequirementFeatureTypeFieldID, field_value: featureType });
        }
        if (constants.RequirementAreaFieldID && area) {
            requestBody.properties.push({ field_id: constants.RequirementAreaFieldID, field_value: area });
        }

        try {
            await put(url, requestBody);
            console.log(`[Info] RICEFW requirement '${requirementId}' updated.`);
        } catch (error) {
            console.log(`[Error] Failed to update RICEFW requirement '${requirementId}'.`, error);
        }
    }

    async function createRequirement(name, description, fields, complexityValue, workItemTypeValue, priorityValue, typeValue, targetModuleId) {
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/requirements`;
        const requestBody = {
            name,
            parent_id: targetModuleId,
            properties: [],
        };

        const areaPath = fields["System.AreaPath"] || "";
        const assignedTo = normalizeAssignedTo(fields["System.AssignedTo"]);
        const iterationPath = fields["System.IterationPath"] || null;
        const state = fields["System.State"] || null;
        const reason = fields["System.Reason"] || null;
        const acceptanceCriteria = fields["Microsoft.VSTS.Common.AcceptanceCriteria"] || "";
        const fullDescription = fields["System.Description"] || "";
        const applicationName = fields[constants.AzDoApplicationNameFieldRef || "Custom.ApplicationName"] || null;
        const fitGap = fields[constants.AzDoFitGapFieldRef || "BP.ERP.FitGap"] ?? null;
        const bpEntity = fields[constants.AzDoBPEntityFieldRef || "Custom.Entity"] ?? null;
        const processRelease = fields[constants.AzDoRicefwProcessReleaseFieldRef || "Custom.ProcessRelease"] || null;
        const ricefwId = fields[constants.AzDoRicefwIdFieldRef || "Custom.RICEFWID"] || null;
        const ricefwConfiguration = fields[constants.AzDoRicefwConfigurationFieldRef || "Custom.RICEFWConfiguration"] || null;
        const testingStatus = fields[constants.AzDoRicefwTestingStatusFieldRef || "Custom.TestingStatus"] || null;
        const featureType = fields[constants.AzDoRicefwFeatureTypeFieldRef || "Custom.FeatureType"] || null;
        const area = fields[constants.AzDoRicefwAreaFieldRef || "Custom.Area"] || null;

        requestBody.properties.push({ field_id: constants.RequirementDescriptionFieldID, field_value: description });
        requestBody.properties.push({ field_id: constants.RequirementStreamSquadFieldID, field_value: areaPath });

        if (constants.RequirementComplexityFieldID && complexityValue) {
            requestBody.properties.push({ field_id: constants.RequirementComplexityFieldID, field_value: complexityValue });
        }
        if (constants.RequirementWorkItemTypeFieldID && workItemTypeValue) {
            requestBody.properties.push({ field_id: constants.RequirementWorkItemTypeFieldID, field_value: workItemTypeValue });
        }
        if (constants.RequirementPriorityFieldID && priorityValue) {
            requestBody.properties.push({ field_id: constants.RequirementPriorityFieldID, field_value: priorityValue });
        }
        if (constants.RequirementTypeFieldID && typeValue) {
            requestBody.properties.push({ field_id: constants.RequirementTypeFieldID, field_value: typeValue });
        }
        if (constants.RequirementAssignedToFieldID && assignedTo) {
            requestBody.properties.push({ field_id: constants.RequirementAssignedToFieldID, field_value: assignedTo });
        }
        if (constants.RequirementIterationPathFieldID && iterationPath) {
            requestBody.properties.push({ field_id: constants.RequirementIterationPathFieldID, field_value: iterationPath });
        }
        if (constants.RequirementApplicationNameFieldID && applicationName) {
            requestBody.properties.push({ field_id: constants.RequirementApplicationNameFieldID, field_value: applicationName });
        }
        if (constants.RequirementFitGapFieldID && fitGap !== null && fitGap !== undefined) {
            requestBody.properties.push({ field_id: constants.RequirementFitGapFieldID, field_value: fitGap });
        }
        if (constants.RequirementBPEntityFieldID && bpEntity !== null && bpEntity !== undefined) {
            requestBody.properties.push({ field_id: constants.RequirementBPEntityFieldID, field_value: bpEntity });
        }

        // RICEFW-specific fields (optional constants)
        if (constants.RequirementStateFieldID && state) {
            requestBody.properties.push({ field_id: constants.RequirementStateFieldID, field_value: state });
        }
        if (constants.RequirementReasonFieldID && reason) {
            requestBody.properties.push({ field_id: constants.RequirementReasonFieldID, field_value: reason });
        }
        if (constants.RequirementAcceptanceCriteriaFieldID) {
            requestBody.properties.push({ field_id: constants.RequirementAcceptanceCriteriaFieldID, field_value: acceptanceCriteria });
        }
        if (constants.RequirementPlainDescriptionFieldID) {
            requestBody.properties.push({ field_id: constants.RequirementPlainDescriptionFieldID, field_value: fullDescription });
        }
        if (constants.RequirementProcessReleaseFieldID && processRelease) {
            requestBody.properties.push({ field_id: constants.RequirementProcessReleaseFieldID, field_value: processRelease });
        }
        if (constants.RequirementRicefwIdFieldID && ricefwId) {
            requestBody.properties.push({ field_id: constants.RequirementRicefwIdFieldID, field_value: ricefwId });
        }
        if (constants.RequirementRicefwConfigurationFieldID && ricefwConfiguration) {
            requestBody.properties.push({ field_id: constants.RequirementRicefwConfigurationFieldID, field_value: ricefwConfiguration });
        }
        if (constants.RequirementTestingStatusFieldID && testingStatus) {
            requestBody.properties.push({ field_id: constants.RequirementTestingStatusFieldID, field_value: testingStatus });
        }
        if (constants.RequirementFeatureTypeFieldID && featureType) {
            requestBody.properties.push({ field_id: constants.RequirementFeatureTypeFieldID, field_value: featureType });
        }
        if (constants.RequirementAreaFieldID && area) {
            requestBody.properties.push({ field_id: constants.RequirementAreaFieldID, field_value: area });
        }

        try {
            await post(url, requestBody);
            console.log(`[Info] RICEFW requirement created.`);
        } catch (error) {
            console.log(`[Error] Failed to create RICEFW requirement.`, error);
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

    if (!isRicefwFeature(event)) {
        console.log("[Info] Work item is not a RICEFW Feature. Exiting.");
        return;
    }

    let workItemId;
    let requirementToUpdate;

    switch (event.eventType) {
        case eventType.CREATED:
            workItemId = event.resource.id;
            console.log(`[Info] Create RICEFW Feature event received for 'WI${workItemId}'`);
            break;

        case eventType.UPDATED:
            workItemId = event.resource.workItemId;
            console.log(`[Info] Update RICEFW Feature event received for 'WI${workItemId}'`);
            const getReqResult = await getRequirementByWorkItemId(workItemId);
            if (getReqResult.failed) return;
            if (!getReqResult.requirement && !constants.AllowCreationOnUpdate) {
                console.log("[Info] Creation of RICEFW Requirement on update event not enabled. Exiting.");
                return;
            }
            requirementToUpdate = getReqResult.requirement;
            break;

        case eventType.DELETED:
            workItemId = event.resource.id;
            console.log(`[Info] Delete RICEFW Feature event received for 'WI${workItemId}'`);
            const getReq = await getRequirementByWorkItemId(workItemId);
            if (getReq.failed || !getReq.requirement) return;
            await deleteRequirement(getReq.requirement);
            return;

        default:
            console.log(`[Error] Unknown workitem event type '${event.eventType}' for 'WI${workItemId}'`);
            return;
    }

    const fields = getFields(event);
    const adoAreaPath = fields["System.AreaPath"] || "";
    const adoComplexity = fields["Custom.Complexity"];
    const qtestComplexityValue = {
        "1 - Very High": 941,
        "2 - High": 942,
        "3 - Medium": 943,
        "4 - Low": 944,
        "5 - Very Low": 945,
    }[adoComplexity] || null;

    const qtestWorkItemTypeValue = 1359;

    const adoPriority = fields["Microsoft.VSTS.Common.Priority"];
    const qtestPriorityValue = {
        1: 11355,
        2: 11356,
        3: 11357,
        4: 11358,
    }[adoPriority] || null;

    const adoRequirementCategory = fields["Custom.RequirementCategory"] || null;
    const qtestTypeValue = {
        "Change Request": 11369,
        "RICEFW": 11370,
    }[adoRequirementCategory] || null;
    const rootModuleId = await ensureRicefwRootModule();
    const iterationPath = fields["System.IterationPath"] || null;
    const targetModuleId = await ensureModulePath(rootModuleId, adoAreaPath, iterationPath);
    const namePrefix = getNamePrefix(workItemId);
    const requirementDescription = buildRequirementDescription(event);
    const requirementName = buildRequirementName(namePrefix, event);

    if (requirementToUpdate) {
        await updateRequirement(
            requirementToUpdate,
            requirementName,
            requirementDescription,
            fields,
            qtestComplexityValue,
            qtestWorkItemTypeValue,
            qtestPriorityValue,
            qtestTypeValue,
            targetModuleId
        );
    } else {
        await createRequirement(
            requirementName,
            requirementDescription,
            fields,
            qtestComplexityValue,
            qtestWorkItemTypeValue,
            qtestPriorityValue,
            qtestTypeValue,
            targetModuleId
        );
    }
};

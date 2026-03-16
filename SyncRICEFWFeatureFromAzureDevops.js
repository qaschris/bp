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
        const acceptanceCriteria = fields["Microsoft.VSTS.Common.AcceptanceCriteria"] || "";
        const description = fields["System.Description"] || "";
        const featureType = fields[constants.AzDoRicefwFeatureTypeFieldRef || "Custom.FeatureType"] || "";
        const processRelease = fields[constants.AzDoRicefwProcessReleaseFieldRef || "Custom.ProcessRelease"] || "";
        const ricefwId = fields[constants.AzDoRicefwIdFieldRef || "Custom.RICEFWID"] || "";
        const ricefwConfiguration = fields[constants.AzDoRicefwConfigurationFieldRef || "Custom.RICEFWConfiguration"] || "";
        const testingStatus = fields[constants.AzDoRicefwTestingStatusFieldRef || "Custom.TestingStatus"] || "";
        const area = fields[constants.AzDoRicefwAreaFieldRef || "Custom.Area"] || "";

        return `<a href="${eventData.resource._links.html.href}" target="_blank">Open in Azure DevOps</a>

<b>Type:</b> ${fields["System.WorkItemType"] || ""}

<b>Feature Type:</b> ${featureType}

<b>RICEFW ID:</b> ${ricefwId}

<b>Area:</b> ${area}

<b>Area Path:</b> ${fields["System.AreaPath"] || ""}

<b>Iteration:</b> ${fields["System.IterationPath"] || ""}

<b>State:</b> ${fields["System.State"] || ""}

<b>Reason:</b> ${fields["System.Reason"] || ""}

<b>Testing Status:</b> ${testingStatus}

<b>Complexity:</b> ${fields["Custom.Complexity"] || ""}

<b>Process Release:</b> ${processRelease}

<b>RICEFW / Configuration:</b> ${ricefwConfiguration}

<b>Assigned To:</b> ${normalizeAssignedTo(fields["System.AssignedTo"])}

<b>Acceptance Criteria:</b> ${acceptanceCriteria}

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
            return moduleChildrenCache[cacheKey];
        }

        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/modules/${parentId}?expand=descendants`;
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

    async function ensureModulePath(rootModuleId, areaPath) {
        const segments = normalizeAreaPathSegments(areaPath);
        let currentParentId = rootModuleId;

        if (!segments.length) {
            console.log(`[Info] AreaPath missing or blank. Using root module '${currentParentId}'.`);
            return currentParentId;
        }

        for (const segment of segments) {
            const children = await getSubModules(currentParentId);
            const existing = children.find(m => ((m?.name || "").trim().toLowerCase() === segment.toLowerCase()));

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

    async function updateRequirement(requirementToUpdate, name, description, fields, targetModuleId) {
        const requirementId = requirementToUpdate.id;
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/requirements/${requirementId}?parentId=${targetModuleId}`;
        const requestBody = {
            name,
            properties: [],
        };

        const areaPath = fields["System.AreaPath"] || "";
        const complexityValue = fields["Custom.Complexity"] || null;
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

    async function createRequirement(name, description, fields, targetModuleId) {
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/requirements`;
        const requestBody = {
            name,
            parent_id: targetModuleId,
            properties: [],
        };

        const areaPath = fields["System.AreaPath"] || "";
        const complexityValue = fields["Custom.Complexity"] || null;
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
    const rootModuleId = await ensureRicefwRootModule();
    const targetModuleId = await ensureModulePath(rootModuleId, adoAreaPath);
    const namePrefix = getNamePrefix(workItemId);
    const requirementDescription = buildRequirementDescription(event);
    const requirementName = buildRequirementName(namePrefix, event);

    if (requirementToUpdate) {
        await updateRequirement(requirementToUpdate, requirementName, requirementDescription, fields, targetModuleId);
    } else {
        await createRequirement(requirementName, requirementDescription, fields, targetModuleId);
    }
};

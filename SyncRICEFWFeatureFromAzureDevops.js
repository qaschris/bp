const axios = require("axios");
const { Webhooks } = require("@qasymphony/pulse-sdk");

exports.handler = async function ({ event, constants, triggers }, context, callback) {
    const qtestMetadataCache = {};
    const emittedMessageKeys = new Set();
    let adoFieldRefs = null;
    let relevantUpdatedFieldRefs = new Set();

    function emitEvent(name, payload) {
        const trigger = triggers.find(item => item.name === name);
        return trigger
            ? new Webhooks().invoke(trigger, payload)
            : console.error(`[ERROR]: (emitEvent) Webhook named '${name}' not found.`);
    }

    function emitFriendlyFailure(details = {}) {
        const platform = details.platform || "Unknown";
        const objectType = details.objectType || "Object";
        const objectId = details.objectId != null ? details.objectId : "Unknown";
        const objectPid = details.objectPid ? ` Object PID: ${details.objectPid}.` : "";
        const fieldName = details.fieldName ? ` Field: ${details.fieldName}.` : "";
        const fieldValue = details.fieldValue != null && details.fieldValue !== ""
            ? ` Value: ${details.fieldValue}.`
            : "";
        const detail = details.detail || "Sync failed.";

        const message =
            `Sync failed. Platform: ${platform}. Object Type: ${objectType}. Object ID: ${objectId}.${objectPid}${fieldName}${fieldValue} Detail: ${detail}`;
        const dedupKey = details.dedupKey || `failure|${message}`;

        if (emittedMessageKeys.has(dedupKey)) return false;
        emittedMessageKeys.add(dedupKey);
        console.error(`[Error] ${message}`);
        emitEvent("ChatOpsEvent", { message });
        return true;
    }

    function emitFriendlyWarning(details = {}) {
        const platform = details.platform || "Unknown";
        const objectType = details.objectType || "Object";
        const objectId = details.objectId != null ? details.objectId : "Unknown";
        const objectPid = details.objectPid ? ` Object PID: ${details.objectPid}.` : "";
        const fieldName = details.fieldName ? ` Field: ${details.fieldName}.` : "";
        const fieldValue = details.fieldValue != null && details.fieldValue !== ""
            ? ` Value: ${details.fieldValue}.`
            : "";
        const detail = details.detail || "Sync warning.";

        const message =
            `Sync warning. Platform: ${platform}. Object Type: ${objectType}. Object ID: ${objectId}.${objectPid}${fieldName}${fieldValue} Detail: ${detail}`;
        const dedupKey = details.dedupKey || `warning|${message}`;

        if (emittedMessageKeys.has(dedupKey)) return false;
        emittedMessageKeys.add(dedupKey);
        console.log(`[Warn] ${message}`);
        emitEvent("ChatOpsEvent", { message });
        return true;
    }

    function normalizeText(value) {
        return value == null
            ? ""
            : String(value)
                .normalize("NFKC")
                .replace(/[\u200B-\u200D\uFEFF]/g, "")
                .replace(/<0x(?:200b|200c|200d|feff)>/gi, "")
                .trim();
    }

    function firstNonEmpty(...values) {
        for (const value of values) {
            if (value !== undefined && value !== null && value !== "") return value;
        }
        return "";
    }

    function safeJson(value) {
        try {
            return JSON.stringify(value, null, 2);
        } catch (error) {
            return `[Unserializable: ${error.message}]`;
        }
    }

    function markFriendlyFailure(error) {
        if (error && typeof error === "object") {
            error.__friendlyFailureEmitted = true;
        }
        return error;
    }

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

        if (typeof value === "object") {
            value =
                value.displayName ||
                value.name ||
                value.uniqueName ||
                value.mail ||
                value.email ||
                "";
        }

        if (typeof value !== "string") {
            return "";
        }

        const cleaned = value
            .replace(/\s*<[^>]*>/g, "")
            .trim();

        return cleaned.replace(/_/g, " ").trim();
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
        const workItemType = normalizeText(getAdoFieldValue(fields, adoFieldRefs.workItemType));
        const featureType = normalizeText(getAdoFieldValue(fields, adoFieldRefs.featureType));
        const ricefwConfiguration = normalizeText(getAdoFieldValue(fields, adoFieldRefs.ricefwConfiguration));
        const featureState = normalizeText(getAdoFieldValue(fields, adoFieldRefs.state));

        console.log(
            `[Debug] Evaluating if work item is a RICEFW Feature: ` +
            `WorkItemType='${workItemType}', FeatureType='${featureType}', ` +
            `RICEFWConfiguration='${ricefwConfiguration}', State='${featureState}'`
        );

        const isRicefwFeature =
            workItemType.toLowerCase() === "feature" &&
            (featureType.toLowerCase() === "ricefw" || featureType.toLowerCase() === "change request") &&
            (
                ricefwConfiguration.toLowerCase() === "enhancement" ||
                ricefwConfiguration.toLowerCase() === "form" ||
                ricefwConfiguration.toLowerCase() === "interface" ||
                ricefwConfiguration.toLowerCase() === "report" ||
                ricefwConfiguration.toLowerCase() === "workflow"
            ) &&
            (featureState.toLowerCase() !== "rejected" && featureState.toLowerCase() !== "cancelled");

        console.log(`[Debug] '${isRicefwFeature}' - Work item ${isRicefwFeature ? "meets" : "does not meet"} criteria for RICEFW Feature.`);
        return isRicefwFeature;
    }

    function buildRequirementName(namePrefix, eventData) {
        return `${namePrefix}${getAdoFieldValue(getFields(eventData), adoFieldRefs.title)}`;
    }

    function buildRequirementDescription(eventData) {
        const fields = getFields(eventData);

        const workItemType = getAdoFieldValue(fields, adoFieldRefs.workItemType) || "";
        const areaPath = getAdoFieldValue(fields, adoFieldRefs.areaPath) || "";
        const iterationPath = getAdoFieldValue(fields, adoFieldRefs.iterationPath) || "";
        const state = getAdoFieldValue(fields, adoFieldRefs.state) || "";
        const reason = getAdoFieldValue(fields, adoFieldRefs.reason) || "";
        const acceptanceCriteria = getAdoFieldValue(fields, adoFieldRefs.acceptanceCriteria) || "";
        const description = getAdoFieldValue(fields, adoFieldRefs.description) || "";
        const ricefwId = getAdoFieldValue(fields, adoFieldRefs.ricefwId) || "";
        const htmlHref = firstNonEmpty(eventData?.resource?._links?.html?.href, eventData?.resource?.revision?._links?.html?.href);

        return `<a href="${htmlHref}" target="_blank">Open in Azure DevOps</a><br>
                <b>Type:</b> ${workItemType}<br>
                <b>Area Path:</b> ${areaPath}<br>
                <b>Iteration:</b> ${iterationPath}<br>
                <b>State:</b> ${state}<br>
                <b>Reason:</b> ${reason}<br>
                <b>RICEFW ID:</b> ${ricefwId}<br>
                <b>Acceptance Criteria:</b> ${acceptanceCriteria}<br>
                <b>Description:</b> ${description}`;
    }

    const standardHeaders = {
        "Content-Type": "application/json",
        Authorization: `bearer ${constants.QTEST_TOKEN}`,
    };

    function normalizeBaseUrl(value) {
        const raw = normalizeText(value).replace(/\/+$/, "");
        if (!raw) throw new Error("A qTest base URL is required.");
        return raw.startsWith("http://") || raw.startsWith("https://") ? raw : `https://${raw}`;
    }

    function normalizeLabel(value) {
        return normalizeText(value).replace(/\s+/g, " ").toLowerCase();
    }

    function normalizeFieldResponse(data) {
        if (Array.isArray(data)) return data;
        if (Array.isArray(data?.items)) return data.items;
        if (Array.isArray(data?.data)) return data.data;
        if (Array.isArray(data?.users)) return data.users;
        return [];
    }

    function getAllowedValues(fieldDefinition, options = {}) {
        const includeInactive = options.includeInactive === true;
        return Array.isArray(fieldDefinition?.allowed_values)
            ? fieldDefinition.allowed_values.filter(v => includeInactive || v?.is_active !== false)
            : [];
    }

    function getAdoFieldValue(fields, fieldRef, options = {}) {
        if (!fields || !fieldRef) return "";
        if (options.preferFormatted) {
            const formatted = fields[`${fieldRef}@OData.Community.Display.V1.FormattedValue`];
            if (formatted !== undefined && formatted !== null && formatted !== "") return formatted;
        }
        return firstNonEmpty(fields[fieldRef]);
    }

    function buildAdoFieldRefs() {
        return {
            title: normalizeText(constants.AzDoTitleFieldRef),
            workItemType: normalizeText(constants.AzDoWorkItemTypeFieldRef),
            areaPath: normalizeText(constants.AzDoAreaPathFieldRef),
            iterationPath: normalizeText(constants.AzDoIterationPathFieldRef),
            state: normalizeText(constants.AzDoStateFieldRef),
            reason: normalizeText(constants.AzDoReasonFieldRef),
            assignedTo: normalizeText(constants.AzDoAssignedToFieldRef),
            description: normalizeText(constants.AzDoDescriptionFieldRef),
            acceptanceCriteria: normalizeText(constants.AzDoAcceptanceCriteriaFieldRef),
            priority: normalizeText(constants.AzDoPriorityFieldRef),
            complexity: normalizeText(constants.AzDoComplexityFieldRef),
            ricefwId: normalizeText(constants.AzDoRICEFWIdFieldRef),
            testingStatus: normalizeText(constants.AzDoTestingStatusFieldRef),
            featureType: normalizeText(constants.AzDoBPFeatureTypeFieldRef),
            ricefwConfiguration: normalizeText(constants.AzDoRICEFWConfigurationFieldRef),
        };
    }

    function buildRelevantUpdatedFieldRefs() {
        return new Set([
            adoFieldRefs.title,
            adoFieldRefs.workItemType,
            adoFieldRefs.areaPath,
            adoFieldRefs.iterationPath,
            adoFieldRefs.state,
            adoFieldRefs.reason,
            adoFieldRefs.assignedTo,
            adoFieldRefs.description,
            adoFieldRefs.acceptanceCriteria,
            adoFieldRefs.priority,
            adoFieldRefs.complexity,
            adoFieldRefs.ricefwId,
            adoFieldRefs.testingStatus,
            adoFieldRefs.featureType,
            adoFieldRefs.ricefwConfiguration,
        ].filter(Boolean));
    }

    function validateRequiredConfiguration() {
        const missingQtestConstants = [
            "RequirementDescriptionFieldID",
            "RequirementStreamSquadFieldID",
            "RequirementWorkItemTypeFieldID",
            "RequirementAssignedToFieldID",
            "RequirementIterationPathFieldID",
            "RequirementStateFieldID",
            "RequirementReasonFieldID",
            "RequirementAcceptanceCriteriaFieldID",
            "RequirementPlainDescriptionFieldID",
            "RequirementRICEFWConfigurationFieldID",
            "RequirementStatusFieldID",
            "FeatureParentID",
        ].filter(name => !normalizeText(constants[name]));

        if (missingQtestConstants.length) {
            emitFriendlyFailure({
                platform: "Pulse",
                objectType: "Configuration",
                objectId: "Unknown",
                fieldName: missingQtestConstants.join(", "),
                detail: "Required qTest RICEFW field constants are missing in Pulse.",
                dedupKey: `ricefw-config-qtest:${missingQtestConstants.join("|")}`,
            });
            return false;
        }

        adoFieldRefs = buildAdoFieldRefs();
        const requiredAdoRefKeys = [
            "title", "workItemType", "areaPath", "iterationPath", "state", "reason",
            "assignedTo", "description", "acceptanceCriteria", "priority", "complexity",
            "ricefwId", "testingStatus", "featureType",
            "ricefwConfiguration",
        ];
        const missingAdoRefs = requiredAdoRefKeys.filter(key => !adoFieldRefs[key]);
        if (missingAdoRefs.length) {
            emitFriendlyFailure({
                platform: "Pulse",
                objectType: "Configuration",
                objectId: "Unknown",
                fieldName: missingAdoRefs.join(", "),
                detail: "Required Azure DevOps RICEFW field reference constants are missing in Pulse.",
                dedupKey: `ricefw-config-ado:${missingAdoRefs.join("|")}`,
            });
            return false;
        }

        relevantUpdatedFieldRefs = buildRelevantUpdatedFieldRefs();
        return true;
    }

    async function getFieldDefinitions(objectType) {
        const cacheKey = `${constants.ProjectID}:${objectType}`;
        if (qtestMetadataCache[cacheKey]) {
            return qtestMetadataCache[cacheKey];
        }

        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/settings/${objectType}/fields`;
        console.log(`[Debug] Fetching qTest field definitions for '${objectType}' from '${url}'.`);

        const response = await axios.get(url, { headers: standardHeaders });
        const fields = normalizeFieldResponse(response.data);
        qtestMetadataCache[cacheKey] = fields;
        return fields;
    }

    async function resolveFieldValue(fieldId, rawValue, objectType, options = {}) {
        if (rawValue === undefined || rawValue === null || rawValue === "") {
            return null;
        }

        const fields = await getFieldDefinitions(objectType);
        const fieldDefinition = fields.find(field => String(field?.id) === String(fieldId));

        if (!fieldDefinition) {
            throw new Error(`Field definition '${fieldId}' was not found for '${objectType}'.`);
        }

        if (!fieldDefinition.constrained) {
            return rawValue;
        }

        const allowedValues = getAllowedValues(fieldDefinition, options);
        const normalizedRawValue = normalizeLabel(rawValue);

        const exactValueMatch = allowedValues.find(option => String(option?.value) === String(rawValue));
        if (exactValueMatch) {
            return exactValueMatch.value;
        }

        const exactLabelMatch = allowedValues.find(option => normalizeLabel(option?.label) === normalizedRawValue);
        if (exactLabelMatch) {
            return exactLabelMatch.value;
        }

        throw new Error(
            `Unable to resolve qTest option for field '${fieldDefinition.label}' (${fieldId}) from value '${rawValue}'.`
        );
    }

    function getChangedFieldRefs(eventData) {
        if (eventData?.eventType !== "workitem.updated") {
            return [];
        }

        return Object.keys(eventData?.resource?.fields || {});
    }

    function shouldProcessRicefwUpdate(eventData) {
        const changedFieldRefs = getChangedFieldRefs(eventData);

        if (!changedFieldRefs.length) {
            console.log("[Info] Updated event did not include field deltas. Continuing with sync.");
            return true;
        }

        const hasRelevantChange = changedFieldRefs.some(fieldRef => relevantUpdatedFieldRefs.has(fieldRef));
        if (!hasRelevantChange) {
            console.log("[Info] Updated event does not include any qTest-synced RICEFW fields. Skipping to prevent loop.");
            return false;
        }

        return true;
    }

    async function resolveRequirementFieldValue(fieldId, rawValue, fieldLabel, requirementContext = null, options = {}) {
        if (!fieldId || rawValue === undefined || rawValue === null || rawValue === "") {
            return null;
        }

        try {
            return await resolveFieldValue(fieldId, rawValue, "requirements", options);
        } catch (error) {
            if (options.emitFailure === false) throw error;
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "RICEFW/Feature",
                objectId: requirementContext?.id || event?.resource?.workItemId || event?.resource?.id || "Unknown",
                objectPid: requirementContext?.pid,
                fieldName: fieldLabel,
                fieldValue: rawValue,
                detail: error.message,
                dedupKey: `ricefw-field-failure:${fieldId}:${normalizeLabel(rawValue)}`,
            });
            throw markFriendlyFailure(error);
        }
    }

    async function resolveOptionalRequirementFieldValue(fieldId, rawValue, fieldLabel, requirementContext = null) {
        if (!fieldId || rawValue === undefined || rawValue === null || rawValue === "") {
            return { value: null, warningDetails: null };
        }

        try {
            const value = await resolveRequirementFieldValue(
                fieldId,
                rawValue,
                fieldLabel,
                requirementContext,
                { includeInactive: true, emitFailure: false }
            );
            return { value, warningDetails: null };
        } catch (error) {
            return {
                value: null,
                warningDetails: {
                    platform: "qTest",
                    objectType: "RICEFW/Feature",
                    objectId: requirementContext?.id || event?.resource?.workItemId || event?.resource?.id || "Unknown",
                    objectPid: requirementContext?.pid,
                    fieldName: fieldLabel,
                    fieldValue: rawValue,
                    detail: `${error.message} The field was left unchanged.`,
                    dedupKey: `ricefw-field-warning:${fieldId}:${normalizeLabel(rawValue)}`,
                },
            };
        }
    }

    function resolveRequirementAssignedToText(adoAssignedTo) {
        return normalizeAssignedTo(adoAssignedTo);
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

        try {
            const response = await axios(opts);
            return response.data;
        } catch (error) {
            console.error(`[Error] HTTP request failed for ${method} ${url}:`, error);
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
            console.error(`[Error] Failed to get sub-modules for parent '${parentId}'.`, error);
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "RICEFW/Feature",
                objectId: "ModulePath",
                detail: `Unable to retrieve qTest module children for parent '${parentId}'.`,
                dedupKey: `ricefw-modules-get:${parentId}`,
            });
            throw markFriendlyFailure(error);
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
            console.error(`[Error] Failed to create module '${name}' under parent '${parentId}'.`, error);
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "RICEFW/Feature",
                objectId: "ModulePath",
                fieldName: "Module",
                fieldValue: name,
                detail: `Unable to create qTest module under parent '${parentId}'.`,
                dedupKey: `ricefw-module-create:${parentId}:${normalizeLabel(name)}`,
            });
            throw markFriendlyFailure(error);
        }
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
                    console.error(`[Error] Module creation for '${segment}' did not return an id.`);
                    emitFriendlyFailure({
                        platform: "qTest",
                        objectType: "RICEFW/Feature",
                        objectId: "ModulePath",
                        fieldName: "Module",
                        fieldValue: segment,
                        detail: "Module creation did not return an id.",
                        dedupKey: `ricefw-module-missing-id:${normalizeLabel(segment)}`,
                    });
                    throw markFriendlyFailure(new Error(`Module creation for '${segment}' did not return an id.`));
                }
            }
        }

        return currentParentId;
    }

    async function getRequirementByWorkItemId(workItemId) {
        const prefix = getNamePrefix(workItemId);
        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/search`;
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
                emitFriendlyFailure({
                    platform: "qTest",
                    objectType: "RICEFW/Feature",
                    objectId: workItemId,
                    detail: "Multiple matching qTest requirements were found for this Azure DevOps feature."
                });
            }
        } catch (error) {
            console.error("[Error] Failed to get requirement by work item id.", error);
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "RICEFW/Feature",
                objectId: workItemId,
                detail: "Unable to locate the matching qTest requirement."
            });
            failed = true;
        }

        return { failed, requirement };
    }

    async function getRequirementDetails(requirementId) {
        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/requirements/${requirementId}`;
        try {
            return await doRequest(url, "GET", null);
        } catch (error) {
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "RICEFW/Feature",
                objectId: requirementId,
                detail: "Unable to retrieve the current qTest requirement details.",
                dedupKey: `ricefw-details:${requirementId}`,
            });
            throw markFriendlyFailure(error);
        }
    }

    function getPropertyById(requirement, fieldId) {
        const properties = Array.isArray(requirement?.properties) ? requirement.properties : [];
        return properties.find(property => String(property?.field_id) === String(fieldId)) || null;
    }

    function getRequirementParentId(requirement) {
        return firstNonEmpty(
            requirement?.parent_id,
            requirement?.parentId,
            requirement?.parent?.id,
            requirement?.module_id,
            requirement?.moduleId
        );
    }

    function valuesEqual(left, right) {
        const normalizeValue = value => {
            if (value === undefined || value === null || value === "") return "";
            return normalizeText(value).replace(/\s+/g, " ");
        };
        return normalizeValue(left) === normalizeValue(right);
    }

    function buildRequirementProperties(desiredState) {
        const properties = [
            { field_id: constants.RequirementDescriptionFieldID, field_value: desiredState.description },
            { field_id: constants.RequirementStreamSquadFieldID, field_value: desiredState.areaPath },
            { field_id: constants.RequirementWorkItemTypeFieldID, field_value: desiredState.workItemTypeValue },
            { field_id: constants.RequirementStateFieldID, field_value: desiredState.stateValue || "" },
            { field_id: constants.RequirementReasonFieldID, field_value: desiredState.reasonValue || "" },
            { field_id: constants.RequirementAcceptanceCriteriaFieldID, field_value: desiredState.acceptanceCriteriaValue || "" },
            { field_id: constants.RequirementPlainDescriptionFieldID, field_value: desiredState.plainDescriptionValue || "" },
            { field_id: constants.RequirementAssignedToFieldID, field_value: desiredState.assignedToText || "" },
        ];

        if (normalizeText(constants.RequirementComplexityFieldID) && desiredState.complexityValue) {
            properties.push({ field_id: constants.RequirementComplexityFieldID, field_value: desiredState.complexityValue });
        }
        if (normalizeText(constants.RequirementPriorityFieldID) && desiredState.priorityValue) {
            properties.push({ field_id: constants.RequirementPriorityFieldID, field_value: desiredState.priorityValue });
        }
        if (normalizeText(constants.RequirementTypeFieldID) && desiredState.typeValue) {
            properties.push({ field_id: constants.RequirementTypeFieldID, field_value: desiredState.typeValue });
        }
        if (desiredState.iterationPathValue) {
            properties.push({ field_id: constants.RequirementIterationPathFieldID, field_value: desiredState.iterationPathValue });
        }
        if (desiredState.ricefwConfigurationValue) {
            properties.push({ field_id: constants.RequirementRICEFWConfigurationFieldID, field_value: desiredState.ricefwConfigurationValue });
        }
        if (desiredState.testingStatusValue) {
            properties.push({ field_id: constants.RequirementStatusFieldID, field_value: desiredState.testingStatusValue });
        }

        return properties;
    }

    function evaluateRequirementUpdate(requirementDetails, desiredState) {
        const desiredProperties = buildRequirementProperties(desiredState);
        const requestBody = { name: desiredState.name, properties: desiredProperties };
        const changedFields = [];

        if (!valuesEqual(requirementDetails?.name, desiredState.name)) {
            changedFields.push("name");
        }

        for (const property of desiredProperties) {
            const currentProperty = getPropertyById(requirementDetails, property.field_id);
            const currentValue = firstNonEmpty(currentProperty?.field_value, currentProperty?.field_value_name);
            if (!valuesEqual(currentValue, property.field_value)) {
                changedFields.push(`field:${property.field_id}`);
            }
        }

        const currentParentId = getRequirementParentId(requirementDetails);
        const parentChanged = desiredState.targetModuleId &&
            String(currentParentId || "") !== String(desiredState.targetModuleId || "");
        if (parentChanged) {
            changedFields.push("parentId");
        }

        return { needsUpdate: changedFields.length > 0, changedFields, parentChanged, requestBody };
    }

    function emitWarnings(warnings, requirementContext = null, fallbackObjectId = null) {
        for (const warning of warnings || []) {
            if (!warning) continue;
            emitFriendlyWarning({
                ...warning,
                objectId: requirementContext?.id || warning.objectId || fallbackObjectId || "Unknown",
                objectPid: requirementContext?.pid || warning.objectPid,
            });
        }
    }

    async function buildDesiredRequirementState(eventData, requirementContext = null) {
        const fields = getFields(eventData);
        const workItemId = eventData?.resource?.workItemId || eventData?.resource?.id;
        const warnings = [];
        const adoAreaPath = getAdoFieldValue(fields, adoFieldRefs.areaPath);
        const adoIterationPath = getAdoFieldValue(fields, adoFieldRefs.iterationPath);
        const adoFeatureType = getAdoFieldValue(fields, adoFieldRefs.featureType);

        const complexityValue = normalizeText(constants.RequirementComplexityFieldID)
            ? await resolveOptionalRequirementFieldValue(constants.RequirementComplexityFieldID, getAdoFieldValue(fields, adoFieldRefs.complexity), "Complexity", requirementContext).then(result => { if (result.warningDetails) warnings.push(result.warningDetails); return result.value; })
            : null;
        const workItemTypeValue = await resolveRequirementFieldValue(constants.RequirementWorkItemTypeFieldID, "Feature", "Work Item Type", requirementContext);
        const priorityValue = normalizeText(constants.RequirementPriorityFieldID)
            ? await resolveOptionalRequirementFieldValue(constants.RequirementPriorityFieldID, getAdoFieldValue(fields, adoFieldRefs.priority), "Priority", requirementContext).then(result => { if (result.warningDetails) warnings.push(result.warningDetails); return result.value; })
            : null;
        const typeValue = normalizeText(constants.RequirementTypeFieldID)
            ? await resolveOptionalRequirementFieldValue(constants.RequirementTypeFieldID, adoFeatureType, "Requirement Category", requirementContext).then(result => { if (result.warningDetails) warnings.push(result.warningDetails); return result.value; })
            : null;
        const assignedToText = resolveRequirementAssignedToText(fields[adoFieldRefs.assignedTo]);

        const iterationResolution = await resolveOptionalRequirementFieldValue(constants.RequirementIterationPathFieldID, adoIterationPath, "Iteration Path", requirementContext);
        if (iterationResolution.warningDetails) warnings.push(iterationResolution.warningDetails);

        const ricefwConfigurationResolution = await resolveOptionalRequirementFieldValue(constants.RequirementRICEFWConfigurationFieldID, getAdoFieldValue(fields, adoFieldRefs.ricefwConfiguration), "RICEFW Configuration", requirementContext);
        if (ricefwConfigurationResolution.warningDetails) warnings.push(ricefwConfigurationResolution.warningDetails);
        const testingStatusResolution = await resolveOptionalRequirementFieldValue(constants.RequirementStatusFieldID, getAdoFieldValue(fields, adoFieldRefs.testingStatus), "Testing Status", requirementContext);
        if (testingStatusResolution.warningDetails) warnings.push(testingStatusResolution.warningDetails);
        const stateResolution = await resolveOptionalRequirementFieldValue(constants.RequirementStateFieldID, getAdoFieldValue(fields, adoFieldRefs.state), "State", requirementContext);
        if (stateResolution.warningDetails) warnings.push(stateResolution.warningDetails);
        const reasonResolution = await resolveOptionalRequirementFieldValue(constants.RequirementReasonFieldID, getAdoFieldValue(fields, adoFieldRefs.reason), "Reason", requirementContext);
        if (reasonResolution.warningDetails) warnings.push(reasonResolution.warningDetails);

        return {
            workItemId,
            name: buildRequirementName(getNamePrefix(workItemId), eventData),
            description: buildRequirementDescription(eventData),
            areaPath: adoAreaPath,
            complexityValue,
            workItemTypeValue,
            priorityValue,
            typeValue,
            assignedToText,
            iterationPathValue: iterationResolution.value,
            stateValue: stateResolution.value,
            reasonValue: reasonResolution.value,
            acceptanceCriteriaValue: getAdoFieldValue(fields, adoFieldRefs.acceptanceCriteria) || "",
            plainDescriptionValue: getAdoFieldValue(fields, adoFieldRefs.description) || "",
            ricefwConfigurationValue: ricefwConfigurationResolution.value,
            testingStatusValue: testingStatusResolution.value,
            targetModuleId: await ensureModulePath(constants.FeatureParentID, adoAreaPath, adoIterationPath),
            warnings,
        };
    }

    async function updateRequirement(requirementToUpdate, desiredState) {
        const requirementDetails = await getRequirementDetails(requirementToUpdate.id);
        const evaluation = evaluateRequirementUpdate(requirementDetails, desiredState);
        if (!evaluation.needsUpdate) {
            console.log(`[Info] RICEFW requirement '${requirementToUpdate.id}' is already in sync. Skipping update.`);
            emitWarnings(desiredState.warnings, requirementDetails, desiredState.workItemId);
            return requirementDetails;
        }

        const query = evaluation.parentChanged ? `?parentId=${desiredState.targetModuleId}` : "";
        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/requirements/${requirementToUpdate.id}${query}`;

        try {
            await put(url, evaluation.requestBody);
            console.log(`[Info] RICEFW requirement '${requirementToUpdate.id}' updated.`);
            emitWarnings(desiredState.warnings, requirementDetails, desiredState.workItemId);
        } catch (error) {
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "RICEFW/Feature",
                objectId: requirementToUpdate.id,
                objectPid: requirementToUpdate?.pid || requirementDetails?.pid,
                detail: "Unable to update the qTest requirement from Azure DevOps.",
                dedupKey: `ricefw-update:${requirementToUpdate.id}`,
            });
            throw markFriendlyFailure(error);
        }
    }

    async function createRequirement(desiredState) {
        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/requirements`;
        const requestBody = {
            name: desiredState.name,
            parent_id: desiredState.targetModuleId,
            properties: buildRequirementProperties(desiredState),
        };

        try {
            await post(url, requestBody);
            console.log("[Info] RICEFW requirement created.");
            emitWarnings(desiredState.warnings, null, desiredState.workItemId);
        } catch (error) {
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "RICEFW/Feature",
                objectId: desiredState.workItemId || "New",
                detail: "Unable to create the qTest requirement from Azure DevOps.",
                dedupKey: `ricefw-create:${desiredState.workItemId || "unknown"}`,
            });
            throw markFriendlyFailure(error);
        }
    }

    async function deleteRequirement(requirementToDelete) {
        const requirementId = requirementToDelete.id;
        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/requirements/${requirementId}`;
        try {
            await doRequest(url, "DELETE", null);
            console.log(`[Info] Requirement '${requirementId}' deleted.`);
        } catch (error) {
            console.error(`[Error] Failed to delete requirement '${requirementId}'.`, error);
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "RICEFW/Feature",
                objectId: requirementId,
                detail: "Unable to delete the qTest requirement."
            });
        }
    }

    const eventType = {
        CREATED: "workitem.created",
        UPDATED: "workitem.updated",
        DELETED: "workitem.deleted",
    };

    let workItemId = null;
    let requirementToUpdate = null;

    try {
        if (!validateRequiredConfiguration()) return;

        switch (event.eventType) {
            case eventType.CREATED:
                workItemId = event?.resource?.id;
                console.log(`[Info] Create RICEFW Feature event received for 'WI${workItemId}'`);
                break;

            case eventType.UPDATED:
                workItemId = event?.resource?.workItemId;
                console.log(`[Info] Update RICEFW Feature event received for 'WI${workItemId}'`);
                if (!shouldProcessRicefwUpdate(event)) return;
                {
                    const getReqResult = await getRequirementByWorkItemId(workItemId);
                    if (getReqResult.failed) return;
                    const allowCreationOnUpdate = String(constants.AllowCreationOnUpdate).toLowerCase() === "true";
                    if (!getReqResult.requirement && !allowCreationOnUpdate) {
                        console.log("[Info] Creation of RICEFW Requirement on update event not enabled. Exiting.");
                        return;
                    }
                    requirementToUpdate = getReqResult.requirement;
                }
                break;

            case eventType.DELETED:
                workItemId = event?.resource?.id;
                console.log(`[Info] Delete RICEFW Feature event received for 'WI${workItemId}'`);
                {
                    const getReq = await getRequirementByWorkItemId(workItemId);
                    if (getReq.failed || !getReq.requirement) return;
                    await deleteRequirement(getReq.requirement);
                    return;
                }

            default:
                emitFriendlyFailure({
                    platform: "ADO",
                    objectType: "RICEFW/Feature",
                    objectId: workItemId || "Unknown",
                    detail: `Unsupported work item event type '${event.eventType}'.`,
                    dedupKey: `ricefw-eventtype:${event.eventType}`,
                });
                return;
        }

        if (!isRicefwFeature(event)) {
            if (workItemId && event.eventType === eventType.UPDATED) {
                const requirementSearch = requirementToUpdate ? { failed: false, requirement: requirementToUpdate } : await getRequirementByWorkItemId(workItemId);
                if (!requirementSearch.failed && requirementSearch.requirement) {
                    emitFriendlyWarning({
                        platform: "ADO",
                        objectType: "RICEFW/Feature",
                        objectId: workItemId,
                        objectPid: requirementSearch.requirement.pid,
                        detail: "The work item no longer meets the configured RICEFW/Feature criteria. The existing qTest item was left unchanged pending business-rule confirmation.",
                        dedupKey: `ricefw-out-of-scope:${workItemId}`,
                    });
                } else {
                    console.log("[Info] Work item is not a RICEFW Feature. Exiting.");
                }
            } else {
                console.log("[Info] Work item is not a RICEFW Feature. Exiting.");
            }
            return;
        }

        const desiredState = await buildDesiredRequirementState(event, requirementToUpdate);
        console.log(`[Debug] RICEFW Requirement Name: ${desiredState.name}`);
        console.log(`[Debug] RICEFW Requirement Description: ${desiredState.description}`);
        console.log(`[Debug] RICEFW Target Module ID: ${desiredState.targetModuleId}`);
        console.log(`[Debug] RICEFW Desired Properties: ${safeJson(buildRequirementProperties(desiredState))}`);

        if (requirementToUpdate) {
            await updateRequirement(requirementToUpdate, desiredState);
        } else {
            await createRequirement(desiredState);
        }
    } catch (error) {
        if (!error?.__friendlyFailureEmitted) {
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "RICEFW/Feature",
                objectId: workItemId || "Unknown",
                objectPid: requirementToUpdate?.pid,
                detail: "Unexpected error occurred during RICEFW requirement sync.",
                dedupKey: `ricefw-fatal:${workItemId || "unknown"}`,
            });
        }
        console.error(error);
    }
};

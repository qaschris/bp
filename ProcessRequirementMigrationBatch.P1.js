const axios = require("axios");
const { Webhooks } = require("@qasymphony/pulse-sdk");

exports.handler = async function ({ event, constants, triggers }, context, callback) {
    const qtestMetadataCache = {};
    const moduleChildrenCache = {};
    const emittedMessageKeys = new Set();
    let adoFieldRefs = null;

    const runId = normalizeText(event?.runId) || `requirement-migration-${Date.now()}`;
    const migrationTargetParentId = normalizeText(firstNonEmpty(event?.targetParentId, constants.RequirementParentID));
    const batchTriggerName = "RequirementMigrationBatchEvent";
    const startTimeMs = Date.now();
    const maxRunMs = parsePositiveInt(event?.maxRunMs, 240000);

    function emitEvent(name, payload) {
        const trigger = triggers.find(item => item.name === name);
        return trigger
            ? new Webhooks().invoke(trigger, payload)
            : console.error(`[ERROR]: (emitEvent) Webhook named '${name}' not found.`);
    }

    function markFriendlyFailure(error) {
        if (error && typeof error === "object") {
            error.__friendlyFailureEmitted = true;
        }
        return error;
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

    function normalizeLabel(value) {
        return normalizeText(value).replace(/\s+/g, " ").toLowerCase();
    }

    function normalizeAssignedToText(value) {
        if (!value) return "";

        if (typeof value === "object") {
            value =
                value.displayName ||
                value.name ||
                value.uniqueName ||
                value.mail ||
                value.email ||
                value.userPrincipalName ||
                "";
        }

        if (typeof value !== "string") {
            return "";
        }

        return value
            .replace(/\s*<[^>]*>/g, "")
            .replace(/_/g, " ")
            .trim();
    }

    function firstNonEmpty(...values) {
        for (const value of values) {
            if (value !== undefined && value !== null && value !== "") return value;
        }
        return "";
    }

    function safeJson(value) {
        try { return JSON.stringify(value, null, 2); } catch (error) { return `[Unserializable: ${error.message}]`; }
    }

    function parsePositiveInt(value, fallback = 0) {
        const parsed = Number.parseInt(normalizeText(value), 10);
        return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
    }

    function getPayloadArray(value) {
        if (Array.isArray(value)) return value;
        if (typeof value === "string") {
            const trimmed = value.trim();
            if (!trimmed) return [];

            if (trimmed.startsWith("[")) {
                try {
                    const parsed = JSON.parse(trimmed);
                    return Array.isArray(parsed) ? parsed : [];
                } catch (error) {
                    return trimmed.split(",").map(item => item.trim()).filter(Boolean);
                }
            }

            return trimmed.split(",").map(item => item.trim()).filter(Boolean);
        }

        return [];
    }

    function normalizeBaseUrl(value) {
        const raw = (value || "").toString().trim().replace(/\/+$/, "");
        if (!raw) throw new Error("A qTest base URL is required.");
        return raw.startsWith("http://") || raw.startsWith("https://") ? raw : `https://${raw}`;
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
            requirementCategory: normalizeText(constants.AzDoRequirementCategoryFieldRef),
            applicationName: normalizeText(constants.AzDoApplicationNameFieldRef),
            fitGap: normalizeText(constants.AzDoFitGapFieldRef),
            entity: normalizeText(constants.AzDoEntityFieldRef),
        };
    }

    function validateRequiredConfiguration() {
        const missingQtestConstants = [
            "QTEST_TOKEN",
            "ManagerURL",
            "ProjectID",
            "RequirementDescriptionFieldID",
            "RequirementStreamSquadFieldID",
            "RequirementComplexityFieldID",
            "RequirementWorkItemTypeFieldID",
            "RequirementPriorityFieldID",
            "RequirementTypeFieldID",
            "RequirementAssignedToFieldID",
            "RequirementIterationPathFieldID",
        ].filter(name => !normalizeText(constants[name]));

        if (!migrationTargetParentId) {
            missingQtestConstants.push("RequirementParentID");
        }

        if (missingQtestConstants.length) {
            emitFriendlyFailure({
                platform: "Pulse",
                objectType: "Configuration",
                objectId: "RequirementMigrationBatch",
                fieldName: missingQtestConstants.join(", "),
                detail: "Required qTest migration constants are missing in Pulse.",
                dedupKey: `requirement-migration-config:qtest:${missingQtestConstants.join("|")}`,
            });
            return false;
        }

        const missingAdoConstants = [
            "AZDO_TOKEN",
            "AzDoProjectURL",
        ].filter(name => !normalizeText(constants[name]));

        if (missingAdoConstants.length) {
            emitFriendlyFailure({
                platform: "Pulse",
                objectType: "Configuration",
                objectId: "RequirementMigrationBatch",
                fieldName: missingAdoConstants.join(", "),
                detail: "Required Azure DevOps migration constants are missing in Pulse.",
                dedupKey: `requirement-migration-config:ado:${missingAdoConstants.join("|")}`,
            });
            return false;
        }

        adoFieldRefs = buildAdoFieldRefs();
        const missingAdoRefs = [
            "title", "workItemType", "areaPath", "iterationPath", "state", "reason",
            "assignedTo", "description", "acceptanceCriteria", "priority", "complexity",
            "requirementCategory",
        ].filter(key => !adoFieldRefs[key]);

        if (missingAdoRefs.length) {
            emitFriendlyFailure({
                platform: "Pulse",
                objectType: "Configuration",
                objectId: "RequirementMigrationBatch",
                fieldName: missingAdoRefs.join(", "),
                detail: "Required Azure DevOps requirement field reference constants are missing in Pulse.",
                dedupKey: `requirement-migration-config:adoRefs:${missingAdoRefs.join("|")}`,
            });
            return false;
        }

        if (!triggers.find(item => item.name === batchTriggerName)) {
            emitFriendlyFailure({
                platform: "Pulse",
                objectType: "Configuration",
                objectId: "RequirementMigrationBatch",
                fieldName: `Trigger:${batchTriggerName}`,
                detail: "The migration batch trigger could not be found in the Pulse trigger list.",
                dedupKey: `requirement-migration-config:missingTrigger:${batchTriggerName}`,
            });
            return false;
        }

        return true;
    }

    function logDivider(title) {
        console.log(`==================== ${title} ====================`);
    }

    function sanitizeHeadersForLog(headers) {
        const clone = { ...(headers || {}) };
        if (clone.Authorization) clone.Authorization = "[REDACTED]";
        return clone;
    }

    async function doRequest({ url, method, headers, requestBody, params }) {
        const opts = { url, method, headers, params };
        if (requestBody !== undefined && requestBody !== null && method !== "GET") {
            opts.data = requestBody;
        }

        logDivider(`HTTP ${method}`);
        console.log(`[Debug] URL: ${url}`);
        console.log(`[Debug] Headers: ${safeJson(sanitizeHeadersForLog(headers))}`);
        console.log(`[Debug] Params: ${safeJson(params || {})}`);
        console.log(`[Debug] Request Payload: ${opts.data !== undefined ? safeJson(opts.data) : "<none>"}`);

        try {
            const response = await axios(opts);
            console.log(`[Debug] HTTP Status: ${response.status}`);
            console.log(`[Debug] Response Body: ${safeJson(response.data)}`);
            return response.data;
        } catch (error) {
            console.error("[Error] HTTP request failed.");
            console.error(`[Error] URL: ${url}`);
            console.error(`[Error] Method: ${method}`);
            console.error(`[Error] Message: ${error.message}`);
            if (error.response) {
                console.error(`[Error] HTTP Status: ${error.response.status}`);
                console.error(`[Error] Error Response Body: ${safeJson(error.response.data)}`);
            }
            throw error;
        }
    }

    function qtestHeaders() {
        return {
            "Content-Type": "application/json",
            Authorization: `bearer ${constants.QTEST_TOKEN}`,
        };
    }

    function adoHeaders() {
        return {
            "Content-Type": "application/json",
            Authorization: `Basic ${Buffer.from(`:${constants.AZDO_TOKEN}`).toString("base64")}`,
        };
    }

    async function qtestGet(url, params) {
        return doRequest({ url, method: "GET", headers: qtestHeaders(), params });
    }

    async function qtestPost(url, requestBody, params) {
        return doRequest({ url, method: "POST", headers: qtestHeaders(), requestBody, params });
    }

    async function qtestPut(url, requestBody, params) {
        return doRequest({ url, method: "PUT", headers: qtestHeaders(), requestBody, params });
    }

    async function adoGet(url, params) {
        return doRequest({ url, method: "GET", headers: adoHeaders(), params });
    }

    async function getFieldDefinitions(objectType) {
        const cacheKey = `${constants.ProjectID}:${objectType}`;
        if (qtestMetadataCache[cacheKey]) return qtestMetadataCache[cacheKey];

        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/settings/${objectType}/fields`;
        const response = await qtestGet(url);
        const fields = normalizeFieldResponse(response);
        qtestMetadataCache[cacheKey] = fields;
        return fields;
    }

    async function resolveFieldValue(fieldId, rawValue, objectType, options = {}) {
        if (rawValue === undefined || rawValue === null || rawValue === "") return null;

        const fields = await getFieldDefinitions(objectType);
        const fieldDefinition = fields.find(field => String(field?.id) === String(fieldId));
        if (!fieldDefinition) throw new Error(`Field definition '${fieldId}' was not found for '${objectType}'.`);
        if (!fieldDefinition.constrained) return rawValue;

        const allowedValues = getAllowedValues(fieldDefinition, options);
        const normalizedRawValue = normalizeLabel(rawValue);
        const exactValueMatch = allowedValues.find(option => String(option?.value) === String(rawValue));
        if (exactValueMatch) return exactValueMatch.value;
        const exactLabelMatch = allowedValues.find(option => normalizeLabel(option?.label) === normalizedRawValue);
        if (exactLabelMatch) return exactLabelMatch.value;

        throw new Error(`Unable to resolve qTest option for field '${fieldDefinition.label}' (${fieldId}) from value '${rawValue}'.`);
    }

    async function resolveRequirementFieldValue(fieldId, rawValue, fieldLabel, requirementContext = null, options = {}) {
        if (!fieldId || rawValue === undefined || rawValue === null || rawValue === "") return null;

        try {
            return await resolveFieldValue(fieldId, rawValue, "requirements", options);
        } catch (error) {
            if (options.emitFailure === false) throw error;
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Requirement",
                objectId: requirementContext?.id || "Unknown",
                objectPid: requirementContext?.pid,
                fieldName: fieldLabel,
                fieldValue: rawValue,
                detail: error.message,
                dedupKey: `requirement-migration-field-failure:${fieldId}:${normalizeLabel(rawValue)}`,
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
                    objectType: "Requirement",
                    objectId: requirementContext?.id || "Unknown",
                    objectPid: requirementContext?.pid,
                    fieldName: fieldLabel,
                    fieldValue: rawValue,
                    detail: `${error.message} The field was left unchanged.`,
                    dedupKey: `requirement-migration-field-warning:${fieldId}:${normalizeLabel(rawValue)}`,
                },
            };
        }
    }

    function normalizeAreaPathSegments(areaPath) {
        if (!areaPath) return [];
        return String(areaPath)
            .split(/[\\/]+/)
            .map(segment => segment.trim())
            .filter(Boolean);
    }

    function getReleaseFolderName(iterationPath) {
        if (!iterationPath) return "TBD";

        const segments = String(iterationPath)
            .split(/[\\/]+/)
            .map(segment => segment.trim())
            .filter(Boolean);
        const candidate = segments.length > 1 ? segments[1] : segments[0];
        if (!candidate) return "TBD";

        const match = candidate.match(/P_O\s+R(\d+(?:\.\d+)?)/i);
        if (match && match[1]) return `P&O Release ${match[1]}`;
        return "TBD";
    }

    async function getSubModules(parentId) {
        const cacheKey = String(parentId);
        if (moduleChildrenCache[cacheKey]) return moduleChildrenCache[cacheKey];

        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/modules/${parentId}?expand=descendants`;
        try {
            const response = await qtestGet(url);

            let items = [];
            if (Array.isArray(response?.children)) items = response.children;
            else if (Array.isArray(response)) items = response;
            else if (Array.isArray(response?.items)) items = response.items;
            else if (Array.isArray(response?.data)) items = response.data;

            moduleChildrenCache[cacheKey] = items;
            return items;
        } catch (error) {
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Requirement",
                objectId: "ModulePath",
                detail: `Unable to retrieve qTest module children for parent '${parentId}'.`,
                dedupKey: `requirement-migration-modules-get:${parentId}`,
            });
            throw markFriendlyFailure(error);
        }
    }

    async function createModule(name, parentId) {
        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/modules`;
        try {
            const created = await qtestPost(url, { name, parent_id: parentId });
            delete moduleChildrenCache[String(parentId)];
            console.log(`[Info] Created qTest module '${name}' under parent '${parentId}'.`);
            return created;
        } catch (error) {
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Requirement",
                objectId: "ModulePath",
                fieldName: "Module",
                fieldValue: name,
                detail: `Unable to create qTest module under parent '${parentId}'.`,
                dedupKey: `requirement-migration-module-create:${parentId}:${normalizeLabel(name)}`,
            });
            throw markFriendlyFailure(error);
        }
    }

    async function ensureModulePath(areaPath, iterationPath) {
        const releaseFolderName = getReleaseFolderName(iterationPath);
        let areaSegments = normalizeAreaPathSegments(areaPath);
        if (areaSegments.length && areaSegments[0].toLowerCase() === "bp_quantum") {
            areaSegments = areaSegments.slice(1);
        }

        const segments = [releaseFolderName, ...areaSegments];
        let currentParentId = migrationTargetParentId;

        for (const segment of segments) {
            const children = await getSubModules(currentParentId);
            const existing = children.find(module => normalizeLabel(module?.name) === normalizeLabel(segment));
            if (existing) {
                currentParentId = existing.id;
                continue;
            }

            const created = await createModule(segment, currentParentId);
            currentParentId = created?.id;
            if (!currentParentId) throw new Error(`Module creation for '${segment}' did not return an id.`);
        }

        return currentParentId;
    }

    async function getRequirementDetails(requirementId) {
        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/requirements/${requirementId}`;
        try {
            return await qtestGet(url);
        } catch (error) {
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Requirement",
                objectId: requirementId,
                detail: "Unable to retrieve the current qTest requirement details.",
                dedupKey: `requirement-migration-details:${requirementId}`,
            });
            throw markFriendlyFailure(error);
        }
    }

    async function fetchAdoWorkItem(workItemId, requirementContext = null) {
        const url = `${constants.AzDoProjectURL}/_apis/wit/workitems/${workItemId}`;
        try {
            return await adoGet(url, { "api-version": "7.1-preview.3" });
        } catch (error) {
            emitFriendlyFailure({
                platform: "Azure DevOps",
                objectType: "Work Item",
                objectId: workItemId,
                objectPid: requirementContext?.pid,
                detail: "Unable to retrieve the Azure DevOps work item for requirement migration.",
                dedupKey: `requirement-migration-ado-read:${workItemId}`,
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

    function extractFieldsFromAdoWorkItem(workItem) {
        return workItem?.fields || {};
    }

    function escapeHtml(value) {
        return String(value ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#39;");
    }

    function getAdoHtmlUrl(workItem) {
        return firstNonEmpty(
            workItem?._links?.html?.href,
            `${constants.AzDoProjectURL}/_workitems/edit/${workItem?.id}`
        );
    }

    function buildRequirementDescription(workItem) {
        const fields = extractFieldsFromAdoWorkItem(workItem);
        const sections = [];

        sections.push(`<a href="${escapeHtml(getAdoHtmlUrl(workItem))}" target="_blank">Open in Azure DevOps</a>`);
        sections.push(`<b>Type:</b> ${escapeHtml(getAdoFieldValue(fields, adoFieldRefs.workItemType))}`);
        sections.push(`<b>Area:</b> ${escapeHtml(getAdoFieldValue(fields, adoFieldRefs.areaPath))}`);
        sections.push(`<b>Iteration:</b> ${escapeHtml(getAdoFieldValue(fields, adoFieldRefs.iterationPath))}`);
        sections.push(`<b>State:</b> ${escapeHtml(getAdoFieldValue(fields, adoFieldRefs.state))}`);
        sections.push(`<b>Reason:</b> ${escapeHtml(getAdoFieldValue(fields, adoFieldRefs.reason))}`);
        sections.push(`<b>Complexity:</b> ${escapeHtml(getAdoFieldValue(fields, adoFieldRefs.complexity))}`);
        sections.push(`<b>Acceptance Criteria:</b> ${escapeHtml(getAdoFieldValue(fields, adoFieldRefs.acceptanceCriteria))}`);
        sections.push(`<b>Description:</b> ${escapeHtml(getAdoFieldValue(fields, adoFieldRefs.description))}`);

        return sections.join("<br>");
    }

    function buildRequirementName(workItemId, workItem) {
        return `WI${workItemId}: ${getAdoFieldValue(extractFieldsFromAdoWorkItem(workItem), adoFieldRefs.title)}`;
    }

    function extractWorkItemIdFromRequirement(requirement) {
        const name = requirement?.name || "";
        const match = name.match(/^WI(\d+):/i);
        return match ? Number(match[1]) : null;
    }

    function buildRequirementProperties(desiredState) {
        const properties = [
            { field_id: constants.RequirementDescriptionFieldID, field_value: desiredState.description },
            { field_id: constants.RequirementStreamSquadFieldID, field_value: desiredState.areaPath },
        ];

        if (desiredState.complexityValue) {
            properties.push({ field_id: constants.RequirementComplexityFieldID, field_value: desiredState.complexityValue });
        }

        if (desiredState.workItemTypeValue) {
            properties.push({ field_id: constants.RequirementWorkItemTypeFieldID, field_value: desiredState.workItemTypeValue });
        }

        if (desiredState.priorityValue) {
            properties.push({ field_id: constants.RequirementPriorityFieldID, field_value: desiredState.priorityValue });
        }

        if (desiredState.typeValue) {
            properties.push({ field_id: constants.RequirementTypeFieldID, field_value: desiredState.typeValue });
        }

        properties.push({ field_id: constants.RequirementAssignedToFieldID, field_value: desiredState.assignedToText || "" });

        if (desiredState.iterationPathValue && normalizeText(constants.RequirementIterationPathFieldID)) {
            properties.push({ field_id: constants.RequirementIterationPathFieldID, field_value: desiredState.iterationPathValue });
        }

        if (desiredState.applicationNameValue && normalizeText(constants.RequirementApplicationNameFieldID)) {
            properties.push({ field_id: constants.RequirementApplicationNameFieldID, field_value: desiredState.applicationNameValue });
        }

        if (desiredState.fitGapValue && normalizeText(constants.RequirementFitGapFieldID)) {
            properties.push({ field_id: constants.RequirementFitGapFieldID, field_value: desiredState.fitGapValue });
        }

        if (desiredState.bpEntityValue && normalizeText(constants.RequirementBPEntityFieldID)) {
            properties.push({ field_id: constants.RequirementBPEntityFieldID, field_value: desiredState.bpEntityValue });
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

        return {
            needsUpdate: changedFields.length > 0,
            changedFields,
            parentChanged,
            requestBody,
        };
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

    async function buildDesiredRequirementState(workItem, requirementContext = null) {
        const fields = extractFieldsFromAdoWorkItem(workItem);
        const workItemId = workItem?.id;
        const warnings = [];

        const adoAreaPath = getAdoFieldValue(fields, adoFieldRefs.areaPath);
        const adoIterationPath = getAdoFieldValue(fields, adoFieldRefs.iterationPath);
        const adoComplexity = getAdoFieldValue(fields, adoFieldRefs.complexity);
        const adoWorkItemType = getAdoFieldValue(fields, adoFieldRefs.workItemType);
        const adoPriority = getAdoFieldValue(fields, adoFieldRefs.priority);
        const adoRequirementCategory = getAdoFieldValue(fields, adoFieldRefs.requirementCategory);
        const adoApplicationName = getAdoFieldValue(fields, adoFieldRefs.applicationName);
        const adoFitGap = getAdoFieldValue(fields, adoFieldRefs.fitGap);
        const adoEntity = getAdoFieldValue(fields, adoFieldRefs.entity);
        const adoAssignedTo = firstNonEmpty(fields[adoFieldRefs.assignedTo]);

        logDivider("EXTRACTED ADO FIELDS");
        console.log(`[Debug] Work Item ID: ${workItemId}`);
        console.log(`[Debug] Work Item Type: ${adoWorkItemType}`);
        console.log(`[Debug] Title: ${getAdoFieldValue(fields, adoFieldRefs.title)}`);
        console.log(`[Debug] AreaPath: ${adoAreaPath}`);
        console.log(`[Debug] IterationPath: ${adoIterationPath}`);
        console.log(`[Debug] State: ${getAdoFieldValue(fields, adoFieldRefs.state)}`);
        console.log(`[Debug] Reason: ${getAdoFieldValue(fields, adoFieldRefs.reason)}`);
        console.log(`[Debug] Complexity: ${adoComplexity}`);
        console.log(`[Debug] AssignedTo Raw: ${safeJson(adoAssignedTo)}`);

        const complexityValue = await resolveRequirementFieldValue(
            constants.RequirementComplexityFieldID,
            adoComplexity,
            "Complexity",
            requirementContext
        );
        const workItemTypeValue = await resolveRequirementFieldValue(
            constants.RequirementWorkItemTypeFieldID,
            adoWorkItemType,
            "Work Item Type",
            requirementContext
        );
        const priorityValue = await resolveRequirementFieldValue(
            constants.RequirementPriorityFieldID,
            adoPriority,
            "Priority",
            requirementContext
        );
        const typeValue = await resolveRequirementFieldValue(
            constants.RequirementTypeFieldID,
            adoRequirementCategory,
            "Requirement Category",
            requirementContext
        );

        const iterationResolution = await resolveOptionalRequirementFieldValue(
            constants.RequirementIterationPathFieldID,
            adoIterationPath,
            "Iteration Path",
            requirementContext
        );
        if (iterationResolution.warningDetails) warnings.push(iterationResolution.warningDetails);

        let applicationNameValue = null;
        if (adoApplicationName) {
            if (!normalizeText(constants.RequirementApplicationNameFieldID)) {
                warnings.push({
                    platform: "Pulse",
                    objectType: "Configuration",
                    objectId: requirementContext?.id || workItemId || "Unknown",
                    objectPid: requirementContext?.pid,
                    fieldName: "RequirementApplicationNameFieldID",
                    detail: "Application Name has a source value but the qTest field id constant is not configured. The field was left unchanged.",
                    dedupKey: `requirement-migration-config-warning:application:${workItemId}`,
                });
            } else {
                const resolution = await resolveOptionalRequirementFieldValue(
                    constants.RequirementApplicationNameFieldID,
                    adoApplicationName,
                    "Application Name",
                    requirementContext
                );
                applicationNameValue = resolution.value;
                if (resolution.warningDetails) warnings.push(resolution.warningDetails);
            }
        }

        let fitGapValue = null;
        if (adoFitGap !== undefined && adoFitGap !== null && adoFitGap !== "") {
            if (!normalizeText(constants.RequirementFitGapFieldID)) {
                warnings.push({
                    platform: "Pulse",
                    objectType: "Configuration",
                    objectId: requirementContext?.id || workItemId || "Unknown",
                    objectPid: requirementContext?.pid,
                    fieldName: "RequirementFitGapFieldID",
                    detail: "Fit Gap has a source value but the qTest field id constant is not configured. The field was left unchanged.",
                    dedupKey: `requirement-migration-config-warning:fitgap:${workItemId}`,
                });
            } else {
                const resolution = await resolveOptionalRequirementFieldValue(
                    constants.RequirementFitGapFieldID,
                    adoFitGap,
                    "Fit Gap",
                    requirementContext
                );
                fitGapValue = resolution.value;
                if (resolution.warningDetails) warnings.push(resolution.warningDetails);
            }
        }

        let bpEntityValue = null;
        if (adoEntity !== undefined && adoEntity !== null && adoEntity !== "") {
            if (!normalizeText(constants.RequirementBPEntityFieldID)) {
                warnings.push({
                    platform: "Pulse",
                    objectType: "Configuration",
                    objectId: requirementContext?.id || workItemId || "Unknown",
                    objectPid: requirementContext?.pid,
                    fieldName: "RequirementBPEntityFieldID",
                    detail: "BP Entity has a source value but the qTest field id constant is not configured. The field was left unchanged.",
                    dedupKey: `requirement-migration-config-warning:entity:${workItemId}`,
                });
            } else {
                const resolution = await resolveOptionalRequirementFieldValue(
                    constants.RequirementBPEntityFieldID,
                    adoEntity,
                    "BP Entity",
                    requirementContext
                );
                bpEntityValue = resolution.value;
                if (resolution.warningDetails) warnings.push(resolution.warningDetails);
            }
        }

        return {
            workItemId,
            name: buildRequirementName(workItemId, workItem),
            description: buildRequirementDescription(workItem),
            areaPath: adoAreaPath,
            complexityValue,
            workItemTypeValue,
            priorityValue,
            typeValue,
            assignedToText: normalizeAssignedToText(adoAssignedTo),
            iterationPathValue: iterationResolution.value,
            applicationNameValue,
            fitGapValue,
            bpEntityValue,
            targetModuleId: await ensureModulePath(adoAreaPath, adoIterationPath),
            warnings,
        };
    }

    async function updateRequirement(requirementDetails, desiredState, evaluation) {
        if (!evaluation.needsUpdate) {
            console.log(`[Info] Requirement '${requirementDetails.id}' is already in sync. Skipping update.`);
            emitWarnings(desiredState.warnings, requirementDetails, desiredState.workItemId);
            return { updated: false, requirement: requirementDetails };
        }

        const params = evaluation.parentChanged ? { parentId: desiredState.targetModuleId } : undefined;
        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/requirements/${requirementDetails.id}`;

        try {
            logDivider("UPDATE REQUIREMENT");
            console.log(`[Debug] Requirement ID: ${requirementDetails.id}`);
            console.log(`[Debug] Changed Fields: ${safeJson(evaluation.changedFields)}`);
            console.log(`[Debug] Final Update Payload: ${safeJson(evaluation.requestBody)}`);
            console.log(`[Debug] Final Update Params: ${safeJson(params || {})}`);
            const updated = await qtestPut(url, evaluation.requestBody, params);
            emitWarnings(desiredState.warnings, requirementDetails, desiredState.workItemId);
            return { updated: true, requirement: updated || requirementDetails };
        } catch (error) {
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Requirement",
                objectId: requirementDetails.id,
                objectPid: requirementDetails?.pid,
                detail: "Unable to update the qTest requirement from Azure DevOps during migration.",
                dedupKey: `requirement-migration-update:${requirementDetails.id}`,
            });
            throw markFriendlyFailure(error);
        }
    }

    function shouldYield() {
        return Date.now() - startTimeMs >= maxRunMs;
    }

    function getRequestedRequirementIds() {
        const payloadRequirementIds = getPayloadArray(event?.requirementIds)
            .map(item => parsePositiveInt(item, 0))
            .filter(Boolean);
        if (payloadRequirementIds.length) return payloadRequirementIds;

        const singleRequirementId = parsePositiveInt(firstNonEmpty(event?.singleRequirementId, event?.requirementId), 0);
        return singleRequirementId ? [singleRequirementId] : [];
    }

    async function processRequirement(requirementId) {
        const requirementDetails = await getRequirementDetails(requirementId);
        const workItemId = extractWorkItemIdFromRequirement(requirementDetails);

        if (!workItemId) {
            emitFriendlyWarning({
                platform: "qTest",
                objectType: "Requirement",
                objectId: requirementDetails?.id || requirementId,
                objectPid: requirementDetails?.pid,
                detail: "Requirement name does not include a 'WI<id>:' prefix. The record was skipped.",
                dedupKey: `requirement-migration-missing-workitem:${requirementDetails?.id || requirementId}`,
            });
            return { status: "skipped", changedFields: [] };
        }

        const workItem = await fetchAdoWorkItem(workItemId, requirementDetails);
        const desiredState = await buildDesiredRequirementState(workItem, requirementDetails);
        const evaluation = evaluateRequirementUpdate(requirementDetails, desiredState);
        const updateResult = await updateRequirement(requirementDetails, desiredState, evaluation);

        console.log(`[Info] Requirement '${requirementId}' processed with work item '${workItemId}'. Updated='${updateResult.updated}'.`);
        return {
            status: updateResult.updated ? "updated" : "skipped",
            changedFields: evaluation.changedFields,
        };
    }

    const counters = {
        processed: 0,
        updated: 0,
        skipped: 0,
        failed: 0,
        deferred: 0,
    };

    try {
        if (!validateRequiredConfiguration()) return;

        const requirementIds = getRequestedRequirementIds();
        const continuationCount = parsePositiveInt(event?.continuationCount, 0);
        const batchNumber = parsePositiveInt(event?.batchNumber, 1);

        console.log(`[Info] Requirement migration batch invoked. RunId='${runId}', Batch='${batchNumber}', Continuation='${continuationCount}', Count='${requirementIds.length}'.`);
        console.log(`[Debug] Incoming Event: ${safeJson(event)}`);

        if (!requirementIds.length) {
            console.log("[Info] No requirement ids were provided to the migration batch worker. Exiting.");
            return;
        }

        const remainingIds = [];

        for (let index = 0; index < requirementIds.length; index += 1) {
            if (index > 0 && shouldYield()) {
                remainingIds.push(...requirementIds.slice(index));
                counters.deferred = remainingIds.length;
                console.log(`[Info] Yielding migration batch with '${remainingIds.length}' requirement ids remaining.`);
                break;
            }

            const requirementId = requirementIds[index];
            counters.processed += 1;

            try {
                const result = await processRequirement(requirementId);
                if (result.status === "updated") counters.updated += 1;
                else counters.skipped += 1;
            } catch (error) {
                counters.failed += 1;
                if (!error?.__friendlyFailureEmitted) {
                    emitFriendlyFailure({
                        platform: "Pulse",
                        objectType: "Requirement",
                        objectId: requirementId,
                        detail: error.response?.data ? safeJson(error.response.data) : error.message,
                        dedupKey: `requirement-migration-item-fatal:${requirementId}`,
                    });
                }
            }
        }

        if (remainingIds.length) {
            await emitEvent(batchTriggerName, {
                ...event,
                runId,
                targetParentId: migrationTargetParentId,
                requirementIds: remainingIds,
                continuationCount: continuationCount + 1,
                maxRunMs,
            });
        }

        console.log(`[Info] Requirement migration batch summary: ${safeJson(counters)}`);
    } catch (error) {
        emitFriendlyFailure({
            platform: "Pulse",
            objectType: "RequirementMigrationBatch",
            objectId: runId || "Unknown",
            detail: error.response?.data ? safeJson(error.response.data) : error.message,
            dedupKey: `requirement-migration-batch-fatal:${runId || "unknown"}`,
        });
        throw error;
    }
};

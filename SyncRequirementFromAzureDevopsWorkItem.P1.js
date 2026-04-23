const axios = require("axios");
const { Webhooks } = require("@qasymphony/pulse-sdk");

exports.handler = async function ({ event, constants, triggers }, context, callback) {
    const qtestMetadataCache = {};
    const moduleChildrenCache = {};
    const emittedMessageKeys = new Set();
    let adoFieldRefs = null;
    let relevantUpdatedFieldRefs = new Set();

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

    function getFields(eventData) {
        return eventData.resource.revision ? eventData.resource.revision.fields : eventData.resource.fields;
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

    function buildRelevantUpdatedFieldRefs() {
        return new Set([
            adoFieldRefs.title,
            adoFieldRefs.workItemType,
            adoFieldRefs.areaPath,
            adoFieldRefs.iterationPath,
            adoFieldRefs.assignedTo,
            adoFieldRefs.description,
            adoFieldRefs.acceptanceCriteria,
            adoFieldRefs.priority,
            adoFieldRefs.complexity,
            adoFieldRefs.requirementCategory,
            adoFieldRefs.applicationName,
            adoFieldRefs.fitGap,
            adoFieldRefs.entity,
        ].filter(Boolean));
    }

    function validateRequiredConfiguration() {
        const missingQtestConstants = [
            "RequirementDescriptionFieldID",
            "RequirementStreamSquadFieldID",
            "RequirementComplexityFieldID",
            "RequirementWorkItemTypeFieldID",
            "RequirementPriorityFieldID",
            "RequirementTypeFieldID",
            "RequirementAssignedToFieldID",
            "RequirementIterationPathFieldID",
        ].filter(name => !normalizeText(constants[name]));

        if (missingQtestConstants.length) {
            emitFriendlyFailure({
                platform: "Pulse",
                objectType: "Configuration",
                objectId: "Unknown",
                fieldName: missingQtestConstants.join(", "),
                detail: "Required qTest requirement field constants are missing in Pulse.",
                dedupKey: `config:qtest:${missingQtestConstants.join("|")}`,
            });
            return false;
        }

        adoFieldRefs = buildAdoFieldRefs();
        const requiredAdoRefKeys = [
            "title", "workItemType", "areaPath", "iterationPath",
            "assignedTo", "description", "acceptanceCriteria", "priority", "complexity",
            "requirementCategory",
        ];
        const missingAdoRefs = requiredAdoRefKeys.filter(key => !adoFieldRefs[key]);

        if (missingAdoRefs.length) {
            emitFriendlyFailure({
                platform: "Pulse",
                objectType: "Configuration",
                objectId: "Unknown",
                fieldName: missingAdoRefs.join(", "),
                detail: "Required Azure DevOps requirement field reference constants are missing in Pulse.",
                dedupKey: `config:ado:${missingAdoRefs.join("|")}`,
            });
            return false;
        }

        relevantUpdatedFieldRefs = buildRelevantUpdatedFieldRefs();
        return true;
    }

    function getChangedFieldRefs(eventData) {
        if (eventData?.eventType !== "workitem.updated") return [];
        return Object.keys(eventData?.resource?.fields || {});
    }

    function shouldProcessRequirementUpdate(eventData) {
        const changedFieldRefs = getChangedFieldRefs(eventData);
        if (!changedFieldRefs.length) {
            console.log("[Info] Updated event did not include field deltas. Continuing with sync.");
            return true;
        }

        if (!changedFieldRefs.some(fieldRef => relevantUpdatedFieldRefs.has(fieldRef))) {
            console.log("[Info] Updated event does not include any qTest-synced requirement fields. Skipping to prevent loop.");
            return false;
        }

        return true;
    }

    const standardHeaders = {
        "Content-Type": "application/json",
        Authorization: `bearer ${constants.QTEST_TOKEN}`,
    };

    function getADOHeaders() {
        const token = Buffer.from(`:${constants.AZDO_TOKEN}`).toString("base64");
        return {
            "Content-Type": "application/json",
            Authorization: `Basic ${token}`,
        };
    }

    function logDivider(title) {
        console.log(`==================== ${title} ====================`);
    }

    async function doRequest(url, method, requestBody, headers = standardHeaders) {
        const opts = { url, method, headers };
        if (requestBody !== undefined && requestBody !== null && method !== "GET") opts.data = requestBody;

        try {
            const response = await axios(opts);
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
            throw new Error(`Failed to ${method} ${url}. ${error.message}`);
        }
    }

    function post(url, requestBody) { return doRequest(url, "POST", requestBody); }
    function put(url, requestBody) { return doRequest(url, "PUT", requestBody); }

    async function getQtestComments(requirementId) {
        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/requirements/${requirementId}/comments?size=1000&page=1`;
        const response = await doRequest(url, "GET", null, standardHeaders);
        return response?.items || [];
    }

    async function createQtestComment(requirementId, content) {
        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/requirements/${requirementId}/comments`;
        return doRequest(url, "POST", { content }, standardHeaders);
    }

    async function updateQtestComment(requirementId, commentId, content) {
        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/requirements/${requirementId}/comments/${commentId}`;
        return doRequest(url, "PUT", { content }, standardHeaders);
    }

    function normalizeCommentText(value) {
        if (!value) return "";

        return String(value)
            .replace(/<style[\s\S]*?<\/style>/gi, "")
            .replace(/<script[\s\S]*?<\/script>/gi, "")
            .replace(/<\/div>|<\/li>|<\/p>/gi, "\r\n")
            .replace(/<li>/gi, "- ")
            .replace(/<br\s*\/?>/gi, "\r\n")
            .replace(/<[^>]+>/gi, "")
            .replace(/&nbsp;/gi, " ")
            .replace(/\r?\n[ \t]*\r?\n[ \t]*/g, "\r\n")
            .trim();
    }

    function formatCommentDate(value) {
        if (!value) return "Unknown date";

        const parsed = new Date(value);
        return Number.isNaN(parsed.getTime()) ? String(value) : parsed.toLocaleString();
    }

    function normalizeStoredCommentContent(value) {
        return String(value || "")
            .replace(/\r\n/g, "\n")
            .trim();
    }

    function buildAdoCommentContent(comment) {
        const marker = `[CID:${comment.commentId}]`;
        const author =
            normalizeText(comment.author) ||
            "Unknown";
        const date = formatCommentDate(comment.createdDate);

        return `[From ADO]
${author} - ${date}
${comment.commentText}

${marker}`;
    }

    async function extractAllComments(workItemId) {
        const url = `${constants.AzDoProjectURL}/_apis/wit/workItems/${workItemId}/comments?api-version=7.0-preview.3`;
        const maxAttempts = 3;

        for (let attempt = 0; attempt < maxAttempts; attempt++) {
            try {
                const response = await doRequest(url, "GET", null, getADOHeaders());
                const comments = response?.comments || [];

                if (comments.length > 0) {
                    console.log(`[Info] Fetched ${comments.length} comments`);
                    return comments
                        .map(comment => ({
                            commentText: normalizeCommentText(comment?.text),
                            commentId: comment.id,
                            author:
                                comment?.createdBy?.displayName ||
                                comment?.createdOnBehalfOf?.displayName ||
                                comment?.revisedBy?.displayName ||
                                comment?.modifiedBy?.displayName ||
                                "",
                            createdDate:
                                comment?.createdOnBehalfDate ||
                                comment?.createdDate ||
                                comment?.revisedDate ||
                                comment?.modifiedDate ||
                                null,
                        }))
                        .filter(comment => comment.commentText)
                        ;
                }

                console.log("[Info] No comments yet, retrying...");
            } catch (error) {
                console.warn("[Warn] Comment fetch failed, retrying...");
            }

            await new Promise(resolve => setTimeout(resolve, 1500));
        }

        console.log("[Info] No comments found after retries");
        return [];
    }

    function isCommentAdded(eventData) {
        const newCount = eventData.resource?.revision?.fields?.["System.CommentCount"] || 0;
        const oldCount = eventData.resource?.fields?.["System.CommentCount"]?.oldValue || 0;
        return newCount > oldCount;
    }

    function hasCommentActivity(eventData) {
        return isCommentAdded(eventData) || Boolean(eventData.resource?.revision?.commentVersionRef?.commentId);
    }

    async function syncCommentsIfNeeded(eventData, workItemId, qtestRequirementId) {
        if (!qtestRequirementId) return;

        if (!hasCommentActivity(eventData)) {
            console.log("[Info] No comment change detected");
            return;
        }

        console.log("[Info] Comment detected. Fetching all from API...");
        const comments = await extractAllComments(workItemId);

        if (comments.length === 0) {
            console.log("[Info] No comment text found");
            return;
        }

        let existing = [];
        try {
            existing = await getQtestComments(qtestRequirementId);
        } catch (error) {
            console.error("[Error] Failed to fetch qTest comments");
            return;
        }

        for (const comment of comments) {
            const marker = `[CID:${comment.commentId}]`;
            const content = buildAdoCommentContent(comment);
            const matchingComment = existing.find(existingComment => (existingComment.content || "").includes(marker));

            try {
                if (!matchingComment) {
                    await createQtestComment(qtestRequirementId, content);
                    existing.unshift({ content });
                    console.log(`[Info] Comment ${comment.commentId} synced successfully`);
                    continue;
                }

                const currentContent = normalizeStoredCommentContent(matchingComment.content);
                const desiredContent = normalizeStoredCommentContent(content);

                if (currentContent === desiredContent) {
                    console.log(`[Info] Comment ${comment.commentId} already synced. Skipping...`);
                    continue;
                }

                await updateQtestComment(qtestRequirementId, matchingComment.id, content);
                matchingComment.content = content;
                console.log(`[Info] Comment ${comment.commentId} updated successfully`);
            } catch (error) {
                console.error(`[Error] Failed to sync qTest comment for ${comment.commentId}`, error);
            }
        }
    }

    async function getFieldDefinitions(objectType) {
        const cacheKey = `${constants.ProjectID}:${objectType}`;
        if (qtestMetadataCache[cacheKey]) return qtestMetadataCache[cacheKey];

        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/settings/${objectType}/fields`;
        const response = await axios.get(url, { headers: standardHeaders });
        const fields = normalizeFieldResponse(response.data);
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
                objectId: requirementContext?.id || event?.resource?.workItemId || event?.resource?.id || "Unknown",
                objectPid: requirementContext?.pid,
                fieldName: fieldLabel,
                fieldValue: rawValue,
                detail: error.message,
                dedupKey: `requirement-field-failure:${fieldId}:${normalizeLabel(rawValue)}`,
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
                    objectId: requirementContext?.id || event?.resource?.workItemId || event?.resource?.id || "Unknown",
                    objectPid: requirementContext?.pid,
                    fieldName: fieldLabel,
                    fieldValue: rawValue,
                    detail: `${error.message} The field was left unchanged.`,
                    dedupKey: `requirement-field-warning:${fieldId}:${normalizeLabel(rawValue)}`,
                },
            };
        }
    }

    function buildRequirementDescription(eventData, options = {}) {
        const fields = getFields(eventData);
        const workItemType = getAdoFieldValue(fields, adoFieldRefs.workItemType);
        const areaPath = getAdoFieldValue(fields, adoFieldRefs.areaPath);
        const iterationPath = getAdoFieldValue(fields, adoFieldRefs.iterationPath);
        const complexity = getAdoFieldValue(fields, adoFieldRefs.complexity);
        const acceptanceCriteria = getAdoFieldValue(fields, adoFieldRefs.acceptanceCriteria);
        const description = getAdoFieldValue(fields, adoFieldRefs.description);
        const htmlHref = firstNonEmpty(eventData?.resource?._links?.html?.href, eventData?.resource?.revision?._links?.html?.href);
        const sections = [];

        if (htmlHref) sections.push(`<a href="${htmlHref}" target="_blank">Open in Azure DevOps</a>`);
        sections.push(`<b>Type:</b> ${workItemType}`);
        sections.push(`<b>Area:</b> ${areaPath}`);
        sections.push(`<b>Iteration:</b> ${iterationPath}`);
        if (options.includeStatusDetails === true) {
            sections.push(`<b>State:</b> ${getAdoFieldValue(fields, adoFieldRefs.state)}`);
            sections.push(`<b>Reason:</b> ${getAdoFieldValue(fields, adoFieldRefs.reason)}`);
        }
        sections.push(`<b>Complexity:</b> ${complexity}`);
        sections.push(`<b>Acceptance Criteria:</b> ${acceptanceCriteria}`);
        sections.push(`<b>Description:</b> ${description}`);

        return sections.join("<br>");
    }

    function getNamePrefix(workItemId) {
        return `WI${workItemId}: `;
    }

    function buildRequirementName(namePrefix, eventData) {
        return `${namePrefix}${getAdoFieldValue(getFields(eventData), adoFieldRefs.title)}`;
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
            const response = await doRequest(url, "GET", null);

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
                dedupKey: `requirement-modules-get:${parentId}`,
            });
            throw markFriendlyFailure(error);
        }
    }

    async function createModule(name, parentId) {
        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/modules`;
        try {
            const created = await post(url, { name, parent_id: parentId });
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
                dedupKey: `requirement-module-create:${parentId}:${normalizeLabel(name)}`,
            });
            throw markFriendlyFailure(error);
        }
    }

    async function ensureModulePath(areaPath, iterationPath) {
        const releaseFolderName = getReleaseFolderName(iterationPath);
        let areaSegments = normalizeAreaPathSegments(areaPath);
        if (areaSegments.length && areaSegments[0].toLowerCase() === "bp_quantum") areaSegments = areaSegments.slice(1);

        const segments = [releaseFolderName, ...areaSegments];
        let currentParentId = constants.RequirementParentID;
        if (!segments.length) return currentParentId;

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

    async function getRequirementByWorkItemId(workItemId) {
        const prefix = getNamePrefix(workItemId);
        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/search`;
        const requestBody = {
            object_type: "requirements",
            fields: ["*"],
            query: `Name ~ '${prefix}'`,
        };

        try {
            logDivider("SEARCH REQUIREMENT BY WORK ITEM ID");
            const response = await post(url, requestBody);

            if (!response || response.total === 0) {
                console.log(`[Info] Requirement not found for Azure DevOps work item '${workItemId}'.`);
                return { failed: false, requirement: null };
            }

            if (response.total > 1) {
                emitFriendlyFailure({
                    platform: "qTest",
                    objectType: "Requirement",
                    objectId: workItemId,
                    detail: "Multiple matching qTest requirements were found for this Azure DevOps work item.",
                    dedupKey: `requirement-search-multiple:${workItemId}`,
                });
                return { failed: true, requirement: null };
            }

            return { failed: false, requirement: response.items[0] };
        } catch (error) {
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Requirement",
                objectId: workItemId,
                detail: "Unable to locate the matching qTest requirement.",
                dedupKey: `requirement-search-failure:${workItemId}`,
            });
            return { failed: true, requirement: null };
        }
    }

    async function getRequirementDetails(requirementId) {
        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/requirements/${requirementId}`;
        try {
            return await doRequest(url, "GET", null);
        } catch (error) {
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Requirement",
                objectId: requirementId,
                detail: "Unable to retrieve the current qTest requirement details.",
                dedupKey: `requirement-details:${requirementId}`,
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
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Requirement",
                objectId: requirementId,
                objectPid: requirementToDelete?.pid,
                detail: "Unable to delete the qTest requirement.",
                dedupKey: `requirement-delete:${requirementId}`,
            });
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

    async function buildDesiredRequirementState(eventData, requirementContext = null, options = {}) {
        const fields = getFields(eventData);
        const workItemId = eventData?.resource?.workItemId || eventData?.resource?.id;
        const warnings = [];
        const namePrefix = getNamePrefix(workItemId);

        const adoTitle = getAdoFieldValue(fields, adoFieldRefs.title);
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
                    dedupKey: `requirement-config-warning:application:${workItemId}`,
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
                    dedupKey: `requirement-config-warning:fitgap:${workItemId}`,
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
                    dedupKey: `requirement-config-warning:entity:${workItemId}`,
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

        const assignedToText = normalizeAssignedToText(adoAssignedTo);

        return {
            workItemId,
            namePrefix,
            name: buildRequirementName(namePrefix, eventData),
            description: buildRequirementDescription(eventData, {
                includeStatusDetails: options.includeStatusDetails === true,
            }),
            areaPath: adoAreaPath,
            complexityValue,
            workItemTypeValue,
            priorityValue,
            typeValue,
            assignedToText,
            iterationPathValue: iterationResolution.value,
            applicationNameValue,
            fitGapValue,
            bpEntityValue,
            targetModuleId: await ensureModulePath(adoAreaPath, adoIterationPath),
            warnings,
        };
    }

    async function updateRequirement(requirementToUpdate, desiredState) {
        const requirementDetails = await getRequirementDetails(requirementToUpdate.id);
        const evaluation = evaluateRequirementUpdate(requirementDetails, desiredState);

        if (!evaluation.needsUpdate) {
            console.log(`[Info] Requirement '${requirementToUpdate.id}' is already in sync. Skipping update.`);
            emitWarnings(desiredState.warnings, requirementDetails, desiredState.workItemId);
            return requirementDetails;
        }

        const query = evaluation.parentChanged ? `?parentId=${desiredState.targetModuleId}` : "";
        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/requirements/${requirementToUpdate.id}${query}`;

        try {
            const updated = await put(url, evaluation.requestBody);
            emitWarnings(desiredState.warnings, requirementDetails, desiredState.workItemId);
            return updated || requirementDetails;
        } catch (error) {
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Requirement",
                objectId: requirementToUpdate.id,
                objectPid: requirementToUpdate?.pid || requirementDetails?.pid,
                detail: "Unable to update the qTest requirement from Azure DevOps.",
                dedupKey: `requirement-update:${requirementToUpdate.id}`,
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
            const created = await post(url, requestBody);
            emitWarnings(desiredState.warnings, created, desiredState.workItemId);
            return created;
        } catch (error) {
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Requirement",
                objectId: desiredState.workItemId || "New",
                detail: "Unable to create the qTest requirement from Azure DevOps.",
                dedupKey: `requirement-create:${desiredState.workItemId || "unknown"}`,
            });
            throw markFriendlyFailure(error);
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
                console.log(`[Info] Create workitem event received for 'WI${workItemId}'.`);
                break;

            case eventType.UPDATED: {
                workItemId = event?.resource?.workItemId;
                console.log(`[Info] Update workitem event received for 'WI${workItemId}'.`);
                const commentChanged = hasCommentActivity(event);
                const shouldProcessUpdate = shouldProcessRequirementUpdate(event);

                const getRequirementResult = await getRequirementByWorkItemId(workItemId);
                if (getRequirementResult.failed) return;

                const allowCreationOnUpdate = String(constants.AllowCreationOnUpdate).toLowerCase() === "true";
                if (!getRequirementResult.requirement && !allowCreationOnUpdate) {
                    console.log("[Info] Creation of Requirement on update event not enabled. Exiting.");
                    return;
                }

                requirementToUpdate = getRequirementResult.requirement;
                if (requirementToUpdate && !shouldProcessUpdate && !commentChanged) return;
                break;
            }

            case eventType.DELETED: {
                workItemId = event?.resource?.id;
                console.log(`[Info] Delete workitem event received for 'WI${workItemId}'.`);
                const getRequirementResult = await getRequirementByWorkItemId(workItemId);
                if (getRequirementResult.failed || !getRequirementResult.requirement) return;
                await deleteRequirement(getRequirementResult.requirement);
                return;
            }

            default:
                emitFriendlyFailure({
                    platform: "ADO",
                    objectType: "Requirement",
                    objectId: workItemId || "Unknown",
                    detail: `Unsupported work item event type '${event.eventType}'.`,
                    dedupKey: `requirement-eventtype:${event.eventType}`,
                });
                return;
        }

        const isCreatingRequirement = !requirementToUpdate;
        const desiredState = await buildDesiredRequirementState(event, requirementToUpdate, {
            includeStatusDetails: isCreatingRequirement,
        });

        let syncedRequirement;
        if (!isCreatingRequirement) {
            syncedRequirement = await updateRequirement(requirementToUpdate, desiredState);
        } else {
            syncedRequirement = await createRequirement(desiredState);
        }

        let qtestRequirementId = syncedRequirement?.id || requirementToUpdate?.id;
        if (!qtestRequirementId) {
            const getRequirementResult = await getRequirementByWorkItemId(workItemId);
            if (!getRequirementResult.failed) {
                qtestRequirementId = getRequirementResult.requirement?.id;
            }
        }

        await syncCommentsIfNeeded(event, workItemId, qtestRequirementId);
    } catch (error) {
        if (!error?.__friendlyFailureEmitted) {
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Requirement",
                objectId: workItemId || "Unknown",
                objectPid: requirementToUpdate?.pid,
                detail: "Unexpected error occurred during requirement sync.",
                dedupKey: `fatal:${workItemId || "unknown"}`,
            });
        }
        console.error(error);
    }
};

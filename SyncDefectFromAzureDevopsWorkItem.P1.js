const { Webhooks } = require('@qasymphony/pulse-sdk');
const axios = require("axios");

exports.handler = async function ({ event, constants, triggers }, context, callback) {
    const qtestMetadataCache = {};
    let failureReported = false;
    const emittedMessageKeys = new Set();
    let adoFieldRefs = null;
    const DEFECT_APPLICATION_FIELD_ID = normalizeText(constants.DefectApplicationFieldID) || "1566";
    const DEFECT_SITE_NAME_FIELD_ID = normalizeText(constants.DefectSiteNameFieldID) || "1569";
    const DEFECT_LINK_TO_AZURE_DEVOPS_LABEL = "Link to Azure DevOps";

    const DEFAULT_AREA_PATH = constants.AreaPath;
    const DEFAULT_QTEST_ASSIGNED_TO_IDENTITY = "ado-qtest-svc@bp.com";

    const standardHeaders = {
        "Content-Type": "application/json",
        Authorization: `bearer ${constants.QTEST_TOKEN}`,
    };

    function normalizeBaseUrl(value) {
        const raw = (value || "").toString().trim().replace(/\/+$/, "");
        if (!raw) {
            throw new Error("A qTest base URL is required.");
        }

        return raw.startsWith("http://") || raw.startsWith("https://")
            ? raw
            : `https://${raw}`;
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

    function normalizeFieldResponse(data) {
        if (Array.isArray(data)) return data;
        if (Array.isArray(data?.items)) return data.items;
        if (Array.isArray(data?.data)) return data.data;
        return [];
    }

    function safeJson(value) {
        try {
            return JSON.stringify(value);
        } catch (error) {
            return `[Unserializable: ${error.message}]`;
        }
    }

    function getAllowedValues(fieldDefinition, options = {}) {
        const includeInactive = options.includeInactive === true;
        return Array.isArray(fieldDefinition?.allowed_values)
            ? fieldDefinition.allowed_values.filter(v => includeInactive || v?.is_active !== false)
            : [];
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

    async function resolveProjectUserId(identity) {
        if (!identity) {
            return null;
        }

        const normalizedIdentity = normalizeLabel(identity);
        if (!normalizedIdentity) {
            return null;
        }

        const cacheKey = `projectUsers:${constants.ProjectID}`;
        let users = qtestMetadataCache[cacheKey];

        if (!users) {
            const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/users`;
            console.log(`[Debug] Fetching active project users from '${url}?inactive=false'.`);
            const response = await axios.get(url, {
                headers: standardHeaders,
                params: { inactive: false },
            });
            users = normalizeFieldResponse(response.data);
            qtestMetadataCache[cacheKey] = users;
        }

        const user = users.find(candidate => {
            const keys = [
                candidate?.username,
                candidate?.ldap_username,
                candidate?.external_user_name,
            ];

            return keys.some(value => normalizeLabel(value) === normalizedIdentity);
        });

        return user?.id ?? null;
    }

    async function resolveDefectFieldValue(fieldId, rawValue, fieldLabel, defectContext = null, options = {}) {
        if (!fieldId || rawValue === undefined || rawValue === null || rawValue === "") {
            return null;
        }

        try {
            return await resolveFieldValue(fieldId, rawValue, "defects", options);
        } catch (error) {
            if (options.emitFailure !== false) {
                emitFriendlyFailure({
                    platform: "qTest",
                    objectType: "Defect",
                    objectId: defectContext?.id ?? event?.resource?.workItemId ?? "Unknown",
                    objectPid: defectContext?.pid,
                    fieldName: fieldLabel,
                    fieldValue: rawValue,
                    detail: error.message,
                });
            }
            throw error;
        }
    }

    async function resolveOptionalDefectFieldValue(fieldId, rawValue, fieldLabel, defectContext = null) {
        if (!fieldId || rawValue === undefined || rawValue === null || rawValue === "") {
            return { value: null, warningDetails: null };
        }

        try {
            const value = await resolveDefectFieldValue(
                fieldId,
                rawValue,
                fieldLabel,
                defectContext,
                { emitFailure: false, includeInactive: true }
            );

            return { value, warningDetails: null };
        } catch (error) {
            return {
                value: null,
                warningDetails: {
                    platform: "qTest",
                    objectType: "Defect",
                    objectId: defectContext?.id ?? event?.resource?.workItemId ?? "Unknown",
                    objectPid: defectContext?.pid,
                    fieldName: fieldLabel,
                    fieldValue: rawValue,
                    detail: error.message,
                    dedupKey: `warning:${fieldLabel}:${defectContext?.id ?? event?.resource?.workItemId ?? "unknown"}:${normalizeLabel(rawValue)}`,
                },
            };
        }
    }

    async function getDefectFieldIdByLabel(fieldLabel) {
        const normalizedFieldLabel = normalizeLabel(fieldLabel);
        if (!normalizedFieldLabel) {
            return null;
        }

        const fields = await getFieldDefinitions("defects");
        const fieldDefinition = fields.find(field => normalizeLabel(field?.label) === normalizedFieldLabel);
        return fieldDefinition?.id ?? null;
    }

    function buildAdoFieldRefs() {
        return {
            title: normalizeText(constants.AzDoTitleFieldRef),
            reproSteps: normalizeText(constants.AzDoReproStepsFieldRef),
            state: normalizeText(constants.AzDoStateFieldRef),
            severity: normalizeText(constants.AzDoSeverityFieldRef),
            priority: normalizeText(constants.AzDoPriorityFieldRef),
            defectType: normalizeText(constants.AzDoDefectTypeFieldRef),
            rootCause: normalizeText(constants.AzDoRootCauseFieldRef),
            proposedFix: normalizeText(constants.AzDoProposedFixFieldRef),
            closedDate: normalizeText(constants.AzDoClosedDateFieldRef),
            targetDate: normalizeText(constants.AzDoTargetDateFieldRef),
            externalReference: normalizeText(constants.AzDoExternalReferenceFieldRef),
            resolvedReason: normalizeText(constants.AzDoResolvedReasonFieldRef),
            areaPath: normalizeText(constants.AzDoAreaPathFieldRef),
            assignedTo: normalizeText(constants.AzDoAssignedToFieldRef),
            application: normalizeText(constants.AzDoApplicationFieldRef),
            siteName: normalizeText(constants.AzDoSiteNameFieldRef),
        };
    }

    function validateRequiredConfiguration() {
        const missingQtestConstants = [
            "DefectSummaryFieldID",
            "DefectDescriptionFieldID",
            "DefectSeverityFieldID",
            "DefectPriorityFieldID",
            "DefectTypeFieldID",
            "DefectStatusFieldID",
            "DefectRootCauseFieldID",
            "DefectExternalReferenceFieldID",
            "DefectResolvedReasonFieldID",
            "DefectAssignedToFieldID",
            "DefectAssignedToTeamFieldID",
            "DefectTargetDateFieldID",
        ].filter(name => !normalizeText(constants[name]));

        if (missingQtestConstants.length) {
            emitFriendlyFailure({
                platform: "Pulse",
                objectType: "Configuration",
                objectId: "Unknown",
                fieldName: missingQtestConstants.join(", "),
                detail: "Required qTest defect field constants are missing in Pulse.",
                dedupKey: `config:qtest:${missingQtestConstants.join("|")}`,
            });
            return false;
        }

        adoFieldRefs = buildAdoFieldRefs();
        const requiredAdoRefKeys = [
            "title",
            "reproSteps",
            "state",
            "severity",
            "priority",
            "defectType",
            "rootCause",
            "proposedFix",
            "closedDate",
            "targetDate",
            "externalReference",
            "resolvedReason",
            "areaPath",
            "assignedTo",
        ];
        const missingAdoRefs = requiredAdoRefKeys
            .filter(key => !adoFieldRefs[key]);

        if (missingAdoRefs.length) {
            emitFriendlyFailure({
                platform: "Pulse",
                objectType: "Configuration",
                objectId: "Unknown",
                fieldName: missingAdoRefs.join(", "),
                detail: "Required Azure DevOps field reference constants are missing in Pulse.",
                dedupKey: `config:ado:${missingAdoRefs.join("|")}`,
            });
            return false;
        }

        return true;
    }

    function getAdoFieldValue(fields, fieldRef, options = {}) {
        if (!fieldRef) {
            return "";
        }

        const formattedKey = `${fieldRef}@OData.Community.Display.V1.FormattedValue`;
        const value = options.preferFormatted
            ? fields?.[formattedKey] ?? fields?.[fieldRef]
            : fields?.[fieldRef] ?? fields?.[formattedKey];

        return value == null ? "" : value;
    }

    const eventType = {
        CREATED: "workitem.created",
        UPDATED: "workitem.updated",
        DELETED: "workitem.deleted",
    };

    function emitEvent(name, payload) {
        return (t = triggers.find(t => t.name === name))
            ? new Webhooks().invoke(t, payload)
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
        if (emittedMessageKeys.has(dedupKey)) {
            return false;
        }

        emittedMessageKeys.add(dedupKey);
        failureReported = true;
        console.error(`[Error] ${message}`);
        emitEvent('ChatOpsEvent', { message });
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
        const detail = details.detail || "Warning.";

        const message =
            `Sync warning. Platform: ${platform}. Object Type: ${objectType}. Object ID: ${objectId}.${objectPid}${fieldName}${fieldValue} Detail: ${detail}`;

        const dedupKey = details.dedupKey || `warning|${message}`;
        if (emittedMessageKeys.has(dedupKey)) {
            return false;
        }

        emittedMessageKeys.add(dedupKey);
        console.log(`[Warn] ${message}`);
        emitEvent('ChatOpsEvent', { message });
        return true;
    }

    function stripEmbeddedAdoLinkText(value) {
        if (!value) {
            return "";
        }

        return String(value)
            .replace(/(?:Link to Azure DevOps:\s*https?:\/\/\S+\s*Repro steps:\s*)+/gi, "")
            .replace(/(?:Link to Azure DevOps:\s*https?:\/\/\S+\s*)+/gi, "")
            .replace(/^(?:Repro steps:\s*)+/i, "")
            .trim();
    }

    function buildDefectDescription(eventData) {
        const fields = getFields(eventData);
        const reproSteps = stripEmbeddedAdoLinkText(
            htmlToPlainText(getAdoFieldValue(fields, adoFieldRefs.reproSteps))
        );
        return reproSteps;
    }

    function buildDefectSummary(namePrefix, eventData) {
        const fields = getFields(eventData);
        return `${namePrefix}${getAdoFieldValue(fields, adoFieldRefs.title)}`;
    }

    function getFields(eventData) {
        return eventData.resource.revision ? eventData.resource.revision.fields : eventData.resource.fields;
    }

    function getNamePrefix(workItemId) {
        return `WI${workItemId}: `;
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

    function normalizeAreaPathLabel(value) {
        return normalizeText(value);
    }

    function mapAdoStatusToQtestLabel(status) {
        const normalizedStatus = normalizeText(status);
        if (!normalizedStatus) {
            return "";
        }

        if (["active", "cancelled"].includes(normalizeLabel(normalizedStatus))) {
            return null;
        }

        return normalizedStatus;
    }

    function mapAdoSeverityToQtestValue(adoSeverity) {
        const severityMap = {
            "1 - Critical": 10301,
            "2 - High": 10302,
            "3 - Medium": 10303,
            "4 - Low": 10304,
        };

        return severityMap[adoSeverity] || null;
    }

    function mapAdoPriorityToQtestValue(adoPriority) {
        const priorityMap = {
            1: 11169,
            2: 10204,
            3: 10203,
            4: 10202,
        };

        return priorityMap[adoPriority] || null;
    }

    async function mapAreaPathToQtestTeamValue(areaPath, defectContext = null) {
        const label = normalizeAreaPathLabel(areaPath);
        const fallbackAreaPath = normalizeAreaPathLabel(DEFAULT_AREA_PATH);
        const candidates = [];

        if (label) {
            candidates.push({ label, usedDefault: false });
        }

        if (fallbackAreaPath && normalizeLabel(fallbackAreaPath) !== normalizeLabel(label)) {
            candidates.push({ label: fallbackAreaPath, usedDefault: true });
        }

        for (const candidate of candidates) {
            try {
                const resolvedValue = await resolveDefectFieldValue(
                    constants.DefectAssignedToTeamFieldID,
                    candidate.label,
                    "Assigned to Team",
                    defectContext,
                    { emitFailure: false }
                );

                return {
                    value: resolvedValue,
                    usedDefault: candidate.usedDefault,
                    warningValue: label || "(blank)",
                    warningDetail: candidate.usedDefault
                        ? `ADO AreaPath '${label || "(blank)"}' could not be resolved to qTest Assigned to Team. Defaulted to area path '${candidate.label}'.`
                        : "",
                };
            } catch (error) {
                console.log(
                    `[Warn] Area Path '${candidate.label}' could not be resolved in qTest Assigned to Team.`
                );
            }
        }

        const warningValue = label || "(blank)";
        if (fallbackAreaPath) {
            return {
                value: null,
                usedDefault: true,
                warningValue,
                warningDetail:
                    `ADO AreaPath '${warningValue}' could not be resolved to qTest Assigned to Team, and the configured default area path '${fallbackAreaPath}' could not be resolved either. Assigned to Team was left unchanged.`,
            };
        }

        return {
            value: null,
            usedDefault: true,
            warningValue,
            warningDetail:
                `ADO AreaPath '${warningValue}' could not be resolved to qTest Assigned to Team, and no default area path is configured. Assigned to Team was left unchanged.`,
        };
    }

    function extractUpnOrEmailFromAdoAssignedTo(raw) {
        if (!raw) return null;

        if (typeof raw === "object") {
            const candidate =
                raw.uniqueName ||
                raw.mail ||
                raw.email ||
                raw.userPrincipalName ||
                raw.displayName ||
                "";
            return candidate ? candidate.trim() : null;
        }

        if (typeof raw !== "string") return null;

        const s = raw.trim();
        const angle = s.match(/<([^>]+)>/);
        if (angle && angle[1]) return angle[1].trim();

        const email = s.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
        if (email && email[0]) return email[0].trim();

        return s || null;
    }

    async function getAdoComments(workItemId, defectContext = null) {
        const url = `${constants.AzDoProjectURL}/_apis/wit/workitems/${workItemId}/comments?api-version=7.0-preview.3`;
        const encodedToken = Buffer.from(`:${constants.AZDO_TOKEN}`).toString("base64");

        try {
            const response = await axios.get(url, {
                headers: { Authorization: `Basic ${encodedToken}` }
            });

            return response.data?.comments || [];
        } catch (error) {
            console.error(`[Error] Failed to fetch ADO comments: ${error.message}`);
            emitFriendlyWarning({
                platform: "ADO",
                objectType: "Defect",
                objectId: workItemId,
                objectPid: defectContext?.pid,
                fieldName: "Discussion",
                detail: "Unable to retrieve comments from Azure DevOps. Discussion sync was skipped.",
                dedupKey: `warning:comments:${workItemId}`,
            });
            return [];
        }
    }

    function formatDiscussion(comments) {
        if (!comments || comments.length === 0) return "";

        return comments.map(c => {
            const author = c.createdBy?.displayName || "Unknown";
            const date = new Date(c.createdDate).toLocaleString();
            const text = c.text || "";

            return `
                <div style="margin-bottom:10px;">
                <strong>${author}</strong>
                <span style="color:gray;">${date}</span>
                <div>${text}</div>
                </div>
                <hr/>`;
        }).join("");
    }

    async function resolveQtestUserIdByUsernameOrUpn(identity, standardHeaders) {
        if (!identity) return null;

        try {
            return await resolveProjectUserId(identity);
        } catch (e) {
            console.error(
                `[Error] qTest project user lookup failed for '${identity}'. ` +
                `Status: ${e?.response?.status ?? "n/a"}`
            );
            return null;
        }
    }

    async function getDefectByWorkItemId(workItemId) {
        const prefix = getNamePrefix(workItemId);
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/search`;
        const requestBody = {
            object_type: "defects",
            fields: ["*"],
            query: `Summary ~ '${prefix}'`,
        };

        console.log(`[Info] Get existing defect for 'WI${workItemId}'`);
        let failed = false;
        let defect = undefined;

        try {
            const response = await post(url, requestBody);

            if (!response || response.total === 0) {
                console.log("[Info] Defect not found by work item id.");
            } else if (response.total === 1) {
                defect = response.items[0];
            } else {
                failed = true;
                console.log("[Warn] Multiple Defects found by work item id.");
                emitFriendlyFailure({
                    platform: "qTest",
                    objectType: "Defect",
                    objectId: workItemId,
                    detail: "Multiple matching qTest defects were found for this Azure DevOps work item."
                });
            }
        } catch (error) {
            console.error("[Error] Failed to get defect by work item id.", error);
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Defect",
                objectId: workItemId,
                detail: "Unable to locate the matching qTest defect."
            });
            failed = true;
        }

        return { failed, defect };
    }

    async function updateDefect(
        defectToUpdate,
        summary,
        description,
        linkFieldId,
        linkValue,
        applicationValue,
        siteNameValue,
        severityValue,
        priorityValue,
        rootCauseValue,
        defectTypeValue,
        statusValue,
        proposedFixValue,
        closedDateValue,
        resolvedReasonValue,
        assignedToUserId,
        targetDateValue,
        discussionValue,
        qtestAssignedToTeamValue,
        adoExternalReference,
        adoState
    ) {
        const defectId = defectToUpdate.id;
        const defectPid = defectToUpdate.pid;

        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/defects/${defectId}`;
        const requestBody = {
            properties: [
                {
                    field_id: constants.DefectSummaryFieldID,
                    field_value: summary,
                },
                {
                    field_id: constants.DefectDescriptionFieldID,
                    field_value: description,
                },
            ],
        };

        if (linkFieldId && linkValue) {
            const parsedLinkFieldId = parseInt(linkFieldId, 10);
            requestBody.properties.push({
                field_id: Number.isNaN(parsedLinkFieldId) ? linkFieldId : parsedLinkFieldId,
                field_value: linkValue,
            });
        }

        if (applicationValue !== null && applicationValue !== undefined && applicationValue !== "") {
            const parsedApplicationValue = parseInt(applicationValue, 10);
            const parsedApplicationFieldId = parseInt(DEFECT_APPLICATION_FIELD_ID, 10);
            requestBody.properties.push({
                field_id: Number.isNaN(parsedApplicationFieldId) ? DEFECT_APPLICATION_FIELD_ID : parsedApplicationFieldId,
                field_value: Number.isNaN(parsedApplicationValue) ? applicationValue : parsedApplicationValue,
            });
        }

        if (siteNameValue !== null && siteNameValue !== undefined && siteNameValue !== "") {
            const parsedSiteNameValue = parseInt(siteNameValue, 10);
            const parsedSiteNameFieldId = parseInt(DEFECT_SITE_NAME_FIELD_ID, 10);
            requestBody.properties.push({
                field_id: Number.isNaN(parsedSiteNameFieldId) ? DEFECT_SITE_NAME_FIELD_ID : parsedSiteNameFieldId,
                field_value: Number.isNaN(parsedSiteNameValue) ? siteNameValue : parsedSiteNameValue,
            });
        }

        if (severityValue) {
            requestBody.properties.push({
                field_id: constants.DefectSeverityFieldID,
                field_value: parseInt(severityValue, 10),
            });
        }

        if (priorityValue) {
            requestBody.properties.push({
                field_id: constants.DefectPriorityFieldID,
                field_value: parseInt(priorityValue, 10),
            });
        }

        if (constants.DefectRootCauseFieldID && rootCauseValue !== null && rootCauseValue !== undefined && rootCauseValue !== "") {
            const parsedRootCauseValue = parseInt(rootCauseValue, 10);
            requestBody.properties.push({
                field_id: constants.DefectRootCauseFieldID,
                field_value: Number.isNaN(parsedRootCauseValue) ? rootCauseValue : parsedRootCauseValue,
            });
            console.log(`[Info] Added Root Cause '${rootCauseValue}' to qTest update payload.`);
        }

        if (defectTypeValue) {
            requestBody.properties.push({
                field_id: constants.DefectTypeFieldID,
                field_value: parseInt(defectTypeValue, 10),
            });
            console.log(`[Info] Added Defect Type '${defectTypeValue}' to qTest update payload.`);
        } else {
            console.log(`[Warn] No Defect Type mapping found or field is empty in ADO.`);
        }

        if (statusValue) {
            requestBody.properties.push({
                field_id: constants.DefectStatusFieldID,
                field_value: parseInt(statusValue, 10),
            });
            console.log(`[Info] Added Status '${statusValue}' to qTest update payload.`);
        } else if (["active", "cancelled"].includes(normalizeLabel(adoState))) {
            console.log(`[Info] Intentionally skipped qTest Status update for ADO State '${adoState}' pending business confirmation.`);
        } else {
            console.log(`[Warn] No Status mapping found or ADO state '${adoState}' not mapped.`);
        }

        if (constants.DefectProposedFixFieldID) {
            const formattedProposedFix = proposedFixValue ? `<p>${proposedFixValue}</p>` : "";
            requestBody.properties.push({
                field_id: constants.DefectProposedFixFieldID,
                field_value: formattedProposedFix,
            });
            console.log(`[Info] Added Proposed Fix to qTest update payload.`);
        }

        if (constants.DefectExternalReferenceFieldID) {
            requestBody.properties.push({
                field_id: constants.DefectExternalReferenceFieldID,
                field_value: adoExternalReference || "",
            });
            console.log(`[Info] Added External Reference to qTest update payload.`);
        }

        if (constants.DefectClosedDateFieldID && closedDateValue) {
            requestBody.properties.push({
                field_id: constants.DefectClosedDateFieldID,
                field_value: closedDateValue,
            });
            console.log(`[Info] Added Closed Date '${closedDateValue}' to qTest update payload.`);
        }

        if (constants.DefectResolvedReasonFieldID && resolvedReasonValue) {
            const parsedResolvedReasonValue = parseInt(resolvedReasonValue, 10);
            requestBody.properties.push({
                field_id: constants.DefectResolvedReasonFieldID,
                field_value: Number.isNaN(parsedResolvedReasonValue)
                    ? resolvedReasonValue
                    : parsedResolvedReasonValue,
            });
            console.log(`[Info] Added Resolved Reason '${resolvedReasonValue}' to qTest update payload.`);
        } else {
            console.log(`[Warn] No Resolved Reason provided or mapping not found`);
        }

        if (constants.DefectTargetDateFieldID && targetDateValue) {
            requestBody.properties.push({
                field_id: constants.DefectTargetDateFieldID,
                field_value: targetDateValue
            });
        }

        if (constants.DefectDiscussionFieldID) {
            requestBody.properties.push({
                field_id: constants.DefectDiscussionFieldID,
                field_value: discussionValue || ""
            });
            console.log(`[Info] Added Discussion to qTest update payload.`);
        }

        if (constants.DefectAssignedToFieldID) {
            requestBody.properties.push({
                field_id: constants.DefectAssignedToFieldID,
                field_value: assignedToUserId ? assignedToUserId : ""
            });

            console.log(
                assignedToUserId
                    ? `[Info] Added Assigned To userId '${assignedToUserId}' to qTest update payload.`
                    : `[Info] Clearing qTest Assigned To (Blank) in update payload.`
            );
        }

        if (constants.DefectAssignedToTeamFieldID && qtestAssignedToTeamValue) {
            requestBody.properties.push({
                field_id: constants.DefectAssignedToTeamFieldID,
                field_value: parseInt(qtestAssignedToTeamValue, 10),
            });
            console.log(
                `[Info] Added Assigned to Team '${qtestAssignedToTeamValue}' to qTest update payload.`
            );
        }

        console.log(`[Info] Updating defect '${defectId}' (${defectPid}).`);
        console.log('[Debug] Final qTest Update Payload:', JSON.stringify(requestBody, null, 2));

        try {
            await put(url, requestBody);
            console.log(`[Info] Defect '${defectId}' (${defectPid}) updated.`);
            return true;
        } catch (error) {
            console.error(`[Error] Failed to update defect '${defectId}'.`, error);
            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Defect",
                objectId: defectId,
                objectPid: defectPid,
                detail: "Unable to update the qTest defect from Azure DevOps."
            });
            return false;
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
            data: method === "GET" ? undefined : requestBody,
            method: method,
        };

        try {
            const response = await axios(opts);
            return response.data;
        } catch (error) {
            const status = error?.response?.status || "Unknown";
            const message = error?.response?.data
                ? JSON.stringify(error.response.data, null, 2)
                : error.message;

            console.error(`[Error] URL: ${url}`);
            console.error(`[Error] HTTP Status: ${status}`);
            console.error(`[Error] Message: ${message}`);

            throw new Error(`qTest API ${method} ${url} failed with ${status}: ${message}`);
        }
    }

    let workItemId = undefined;
    let defectToUpdate = undefined;

    switch (event.eventType) {
        case eventType.CREATED: {
            workItemId = event.resource.workItemId;
            console.log(`[Info] Create workitem event received for 'WI${workItemId}'`);
            console.log(`[Info] New defects are not synched from Azure DevOps. The current workflow expects the defect to be created in qTest first. Exiting.`);
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
            workItemId = event.resource.workItemId;
            console.log(`[Info] Delete workitem event received for 'WI${workItemId}'`);
            console.log(`[Info] Defects are not deleted in qTest automatically when deleting in Azure DevOps. Exiting.`);
            return;
        }
        default:
            console.log(`[Error] Unknown workitem event type '${event.eventType}' for 'WI${workItemId}'`);
            emitFriendlyFailure({
                platform: "ADO",
                objectType: "Defect",
                objectId: workItemId || "Unknown",
                detail: `Unsupported work item event type '${event.eventType}'.`
            });
            return;
    }

    if (!validateRequiredConfiguration()) {
        return;
    }

    const namePrefix = getNamePrefix(workItemId);
    const defectDescription = buildDefectDescription(event);
    const defectSummary = buildDefectSummary(namePrefix, event);
    const fields = getFields(event);
    const defectContext = defectToUpdate
        ? { id: defectToUpdate.id, pid: defectToUpdate.pid }
        : null;
    const adoLinkValue = normalizeText(event?.resource?._links?.html?.href);
    const qtestLinkFieldId = normalizeText(constants.DefectLinkToAzureDevOpsFieldID)
        || String(await getDefectFieldIdByLabel(DEFECT_LINK_TO_AZURE_DEVOPS_LABEL) || "");
    let linkFieldWarningDetails = null;

    if (adoLinkValue && !qtestLinkFieldId) {
        linkFieldWarningDetails = {
            platform: "qTest",
            objectType: "Defect",
            objectId: defectContext?.id ?? workItemId,
            objectPid: defectContext?.pid,
            fieldName: DEFECT_LINK_TO_AZURE_DEVOPS_LABEL,
            fieldValue: adoLinkValue,
            detail: `qTest field '${DEFECT_LINK_TO_AZURE_DEVOPS_LABEL}' was not found. Azure DevOps link sync was skipped.`,
            dedupKey: `warning:link-field:${defectContext?.id ?? workItemId}`,
        };
    }

    const adoAreaPath = getAdoFieldValue(fields, adoFieldRefs.areaPath);
    console.log(`[Info] ADO AreaPath value: '${adoAreaPath}'`);

    const qtestAssignedToTeamResult = await mapAreaPathToQtestTeamValue(adoAreaPath, defectContext);
    const qtestAssignedToTeamValue = qtestAssignedToTeamResult.value;
    let assignedToTeamWarningDetails = null;
    console.log(
        `[Info] Mapped ADO AreaPath '${adoAreaPath || "(blank)"}' to qTest Assigned to Team value '${qtestAssignedToTeamValue || "(unchanged)"}'`
    );

    if (qtestAssignedToTeamResult.warningDetail) {
        assignedToTeamWarningDetails = {
            platform: "qTest",
            objectType: "Defect",
            objectId: defectContext?.id ?? workItemId,
            objectPid: defectContext?.pid,
            fieldName: "Assigned to Team",
            fieldValue: qtestAssignedToTeamResult.warningValue,
            detail: qtestAssignedToTeamResult.warningDetail,
            dedupKey: `warning:assigned-to-team:${defectContext?.id ?? workItemId}:${normalizeLabel(qtestAssignedToTeamResult.warningValue)}`,
        };
    }

    let adoComments = await getAdoComments(workItemId, defectContext);

    if (constants.SyncUserRegex) {
        const regex = new RegExp(constants.SyncUserRegex, "i");
        adoComments = adoComments.filter(c => !regex.test(c.createdBy?.displayName || ""));
    }

    const discussionHtml = formatDiscussion(adoComments);
    let qtestDiscussionValue = "";

    if (discussionHtml) {
        qtestDiscussionValue = `
                    <h3>ADO Discussion</h3>
                    ${discussionHtml}
                    `;
    }

    const adoAssignedToRaw = getAdoFieldValue(fields, adoFieldRefs.assignedTo);
    const adoAssignedToIdentity = extractUpnOrEmailFromAdoAssignedTo(adoAssignedToRaw);
    let assignedToWarningDetails = null;

    let qtestAssignedToUserId = null;
    if (adoAssignedToIdentity) {
        console.log(`[Info] Normalized ADO Assigned To to '${adoAssignedToIdentity}'`);

        qtestAssignedToUserId = await resolveQtestUserIdByUsernameOrUpn(adoAssignedToIdentity, standardHeaders);

        if (qtestAssignedToUserId) {
            console.log(`[Info] Resolved ADO Assigned To '${adoAssignedToIdentity}' -> qTest userId '${qtestAssignedToUserId}'`);
        } else {
            console.log(
                `[Warn] Could not resolve ADO Assigned To '${adoAssignedToIdentity}' in qTest. ` +
                `Attempting fallback to '${DEFAULT_QTEST_ASSIGNED_TO_IDENTITY}'.`
            );
        }
    } else if (adoAssignedToRaw) {
        console.log(`[Warn] Could not normalize ADO Assigned To '${adoAssignedToRaw}'`);
    } else {
        console.log(`[Info] ADO Assigned To is blank/unassigned.`);
    }

    if (!qtestAssignedToUserId && adoAssignedToRaw) {
        qtestAssignedToUserId = await resolveQtestUserIdByUsernameOrUpn(DEFAULT_QTEST_ASSIGNED_TO_IDENTITY, standardHeaders);

        if (qtestAssignedToUserId) {
            console.log(
                `[Info] Falling back qTest Assigned To to '${DEFAULT_QTEST_ASSIGNED_TO_IDENTITY}' ` +
                `with userId '${qtestAssignedToUserId}'.`
            );
            assignedToWarningDetails = {
                platform: "qTest",
                objectType: "Defect",
                objectId: defectContext?.id ?? workItemId,
                objectPid: defectContext?.pid,
                fieldName: "Assigned To",
                fieldValue: adoAssignedToIdentity || normalizeText(adoAssignedToRaw) || "(blank)",
                detail:
                    `ADO Assigned To '${adoAssignedToIdentity || normalizeText(adoAssignedToRaw) || "(blank)"}' could not be resolved in qTest. ` +
                    `Defaulted qTest Assigned To to '${DEFAULT_QTEST_ASSIGNED_TO_IDENTITY}'.`,
                dedupKey: `warning:assigned-to:${defectContext?.id ?? workItemId}:${normalizeLabel(adoAssignedToIdentity || adoAssignedToRaw || "(blank)")}`,
            };
        } else {
            console.log(
                `[Warn] Fallback qTest Assigned To identity '${DEFAULT_QTEST_ASSIGNED_TO_IDENTITY}' ` +
                `was not found in the qTest project. Assignment will be cleared.`
            );
            assignedToWarningDetails = {
                platform: "qTest",
                objectType: "Defect",
                objectId: defectContext?.id ?? workItemId,
                objectPid: defectContext?.pid,
                fieldName: "Assigned To",
                fieldValue: adoAssignedToIdentity || normalizeText(adoAssignedToRaw) || "(blank)",
                detail:
                    `ADO Assigned To '${adoAssignedToIdentity || normalizeText(adoAssignedToRaw) || "(blank)"}' could not be resolved in qTest, ` +
                    `and fallback user '${DEFAULT_QTEST_ASSIGNED_TO_IDENTITY}' was not found in the qTest project. Assigned To was cleared.`,
                dedupKey: `warning:assigned-to-clear:${defectContext?.id ?? workItemId}:${normalizeLabel(adoAssignedToIdentity || adoAssignedToRaw || "(blank)")}`,
            };
        }
    }

    const adoSeverity = getAdoFieldValue(fields, adoFieldRefs.severity);
    console.log(`[Info] ADO Severity value: '${adoSeverity}'`);

    const qtestSeverityValue = mapAdoSeverityToQtestValue(adoSeverity);
    console.log(`[Info] Mapped ADO Severity '${adoSeverity}' to qTest Severity value '${qtestSeverityValue}'`);

    const adoPriority = getAdoFieldValue(fields, adoFieldRefs.priority);
    console.log(`[Info] ADO Priority value: '${adoPriority}'`);

    const qtestPriorityValue = mapAdoPriorityToQtestValue(adoPriority);
    console.log(`[Info] Mapped ADO Priority '${adoPriority}' to qTest Priority value '${qtestPriorityValue}'`);

    const adoDefectType = getAdoFieldValue(fields, adoFieldRefs.defectType);
    console.log(`[Info] ADO Defect Type value: '${adoDefectType}'`);

    const qtestDefectTypeValue = await resolveDefectFieldValue(
        constants.DefectTypeFieldID,
        adoDefectType,
        "Defect Type",
        defectContext
    );
    console.log(`[Info] Mapped ADO Defect Type '${adoDefectType}' to qTest Defect Type value '${qtestDefectTypeValue}'`);

    const adoState = getAdoFieldValue(fields, adoFieldRefs.state);
    console.log(`[Info] ADO State value: '${adoState}'`);
    const qtestStatusLabel = mapAdoStatusToQtestLabel(adoState);
    let qtestStatusValue = null;
    if (qtestStatusLabel === null) {
        console.log(`[Info] Skipping qTest status update for ADO State '${adoState}' pending business confirmation.`);
    } else {
        qtestStatusValue = await resolveDefectFieldValue(
            constants.DefectStatusFieldID,
            qtestStatusLabel,
            "Status",
            defectContext
        );
        console.log(`[Info] Mapped ADO State '${adoState}' to qTest Status label '${qtestStatusLabel}' and qTest Status value '${qtestStatusValue}'`);
        if (!qtestStatusValue) {
            console.log(`[Warn] ADO State '${adoState}' does not match any defined qTest status.`);
        }
    }

    const adoRootCauseRaw = adoFieldRefs.rootCause ? fields?.[adoFieldRefs.rootCause] : "";
    const adoRootCauseFormatted = adoFieldRefs.rootCause
        ? fields?.[`${adoFieldRefs.rootCause}@OData.Community.Display.V1.FormattedValue`]
        : "";
    const adoRootCause = getAdoFieldValue(fields, adoFieldRefs.rootCause, { preferFormatted: true });
    console.log(`[Debug] ADO Root Cause diagnostics: ${safeJson({
        fieldRef: adoFieldRefs.rootCause,
        raw: adoRootCauseRaw,
        formatted: adoRootCauseFormatted,
        selected: adoRootCause,
    })}`);

    const qtestRootCauseResult = await resolveOptionalDefectFieldValue(
        constants.DefectRootCauseFieldID,
        adoRootCause,
        "Root Cause",
        defectContext
    );
    const qtestRootCauseValue = qtestRootCauseResult.value;
    if (qtestRootCauseValue) {
        console.log(`[Info] Mapped ADO Root Cause '${adoRootCause}' to qTest Root Cause value '${qtestRootCauseValue}'`);
    } else if (adoRootCause) {
        console.log(`[Warn] ADO Root Cause '${adoRootCause}' could not be mapped to qTest. Root Cause update will be skipped.`);
    }

    const adoProposedFix = getAdoFieldValue(fields, adoFieldRefs.proposedFix);
    console.log(`[Info] ADO Proposed Fix value length: ${adoProposedFix.length}`);

    const qtestProposedFixValue = adoProposedFix;

    const adoActualCloseDate = getAdoFieldValue(fields, adoFieldRefs.closedDate);
    let qtestClosedDateValue = null;

    const adoTargetDate = getAdoFieldValue(fields, adoFieldRefs.targetDate);
    let qtestTargetDateValue = null;
    if (adoTargetDate) {
        qtestTargetDateValue = new Date(adoTargetDate).toISOString().replace(".000Z", "+00:00");
        console.log(`[Info] ADO Target Date: '${adoTargetDate}' => qTest Target Date: '${qtestTargetDateValue}'`);
    }

    const adoExternalReference = getAdoFieldValue(fields, adoFieldRefs.externalReference);
    console.log(`[Info] ADO External Reference value: '${adoExternalReference}'`);

    if (adoActualCloseDate) {
        qtestClosedDateValue = new Date(adoActualCloseDate).toISOString().replace(".000Z", "+00:00");
        console.log(`[Info] ADO Actual Close Date: '${adoActualCloseDate}' => qTest Closed Date: '${qtestClosedDateValue}'`);
    } else {
        console.log(`[Info] No Actual Close Date found in ADO.`);
    }

    const adoResolvedReason = getAdoFieldValue(fields, adoFieldRefs.resolvedReason, { preferFormatted: true });
    console.log(`[Info] ADO Resolved Reason: '${adoResolvedReason}'`);

    const qtestResolvedReasonValue = await resolveDefectFieldValue(
        constants.DefectResolvedReasonFieldID,
        adoResolvedReason,
        "Resolved Reason",
        defectContext
    );
    console.log(`[Info] Mapped ADO Resolved Reason '${adoResolvedReason}' → qTest value '${qtestResolvedReasonValue}'`);

    const adoApplication = adoFieldRefs.application
        ? getAdoFieldValue(fields, adoFieldRefs.application, { preferFormatted: true })
        : "";
    console.log(`[Info] ADO Application value: '${adoApplication}'`);

    const qtestApplicationResult = await resolveOptionalDefectFieldValue(
        DEFECT_APPLICATION_FIELD_ID,
        adoApplication,
        "Application",
        defectContext
    );
    const qtestApplicationValue = qtestApplicationResult.value;
    if (qtestApplicationValue) {
        console.log(`[Info] Mapped ADO Application '${adoApplication}' to qTest Application value '${qtestApplicationValue}'`);
    } else if (adoApplication) {
        console.log(`[Warn] ADO Application '${adoApplication}' could not be mapped to qTest. Application update will be skipped.`);
    }

    const adoSiteName = adoFieldRefs.siteName
        ? getAdoFieldValue(fields, adoFieldRefs.siteName, { preferFormatted: true })
        : "";
    console.log(`[Info] ADO Site Name value: '${adoSiteName}'`);

    const qtestSiteNameResult = await resolveOptionalDefectFieldValue(
        DEFECT_SITE_NAME_FIELD_ID,
        adoSiteName,
        "Site Name",
        defectContext
    );
    const qtestSiteNameValue = qtestSiteNameResult.value;
    if (qtestSiteNameValue) {
        console.log(`[Info] Mapped ADO Site Name '${adoSiteName}' to qTest Site Name value '${qtestSiteNameValue}'`);
    } else if (adoSiteName) {
        console.log(`[Warn] ADO Site Name '${adoSiteName}' could not be mapped to qTest. Site Name update will be skipped.`);
    }

    if (defectToUpdate) {
        const updateSucceeded = await updateDefect(
            defectToUpdate,
            defectSummary,
            defectDescription,
            qtestLinkFieldId,
            adoLinkValue,
            qtestApplicationValue,
            qtestSiteNameValue,
            qtestSeverityValue,
            qtestPriorityValue,
            qtestRootCauseValue,
            qtestDefectTypeValue,
            qtestStatusValue,
            qtestProposedFixValue,
            qtestClosedDateValue,
            qtestResolvedReasonValue,
            qtestAssignedToUserId,
            qtestTargetDateValue,
            qtestDiscussionValue,
            qtestAssignedToTeamValue,
            adoExternalReference,
            adoState
        );

        if (updateSucceeded) {
            if (linkFieldWarningDetails) {
                emitFriendlyWarning(linkFieldWarningDetails);
            }

            if (assignedToWarningDetails) {
                emitFriendlyWarning(assignedToWarningDetails);
            }

            if (assignedToTeamWarningDetails) {
                emitFriendlyWarning(assignedToTeamWarningDetails);
            }

            if (qtestRootCauseResult.warningDetails) {
                emitFriendlyWarning(qtestRootCauseResult.warningDetails);
            }

            if (qtestApplicationResult.warningDetails) {
                emitFriendlyWarning(qtestApplicationResult.warningDetails);
            }

            if (qtestSiteNameResult.warningDetails) {
                emitFriendlyWarning(qtestSiteNameResult.warningDetails);
            }
        }
    }
};

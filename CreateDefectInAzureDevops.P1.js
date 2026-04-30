const { Webhooks } = require('@qasymphony/pulse-sdk');
const axios = require('axios');

exports.handler = async function ({ event, constants, triggers }, context, callback) {
    let iteration = event.iteration != undefined ? event.iteration : 1;
    const maxIterations = 21;
    const retryAttemptsPerIteration = 12;
    const retryDelayMs = 5000;
    const defectId = event.defect.id;
    const projectId = event.defect.project_id;
    const qtestMetadataCache = {};
    const DEFAULT_ADO_ASSIGNED_TO = "ado-qtest-svc@bp.com";
    const approximateTimeoutMinutes = Math.ceil(
        (maxIterations * Math.max(retryAttemptsPerIteration - 1, 0) * retryDelayMs) / 60000
    );
    let failureReported = false;
    const emittedMessageKeys = new Set();
    let adoFieldRefs = null;
    const DEFECT_APPLICATION_FIELD_ID = normalizeText(constants.DefectApplicationFieldID) || "1566";
    const DEFECT_SITE_NAME_FIELD_ID = normalizeText(constants.DefectSiteNameFieldID) || "1569";
    const DEFECT_ITERATION_PATH_FIELD_ID = normalizeText(constants.DefectIterationPathFieldID) || "1603";
    const DEFECT_LINK_TO_AZURE_DEVOPS_LABEL = "Link to Azure DevOps";
    const OPTIONAL_NONE_LABEL = "None";

    console.log(`[Info] Create defect event received for defect '${defectId}' in project '${projectId}'`);

    if (projectId != constants.ProjectID) {
        console.log(`[Info] Project not matching '${projectId}' != '${constants.ProjectID}', exiting.`);
        return;
    }

    if (!validateRequiredConfiguration()) {
        return;
    }

    const DEFAULT_AREA_PATH = constants.AreaPath;

    function normalizeText(value) {
        return value == null
            ? ""
            : String(value)
                .normalize("NFKC")
                .replace(/[\u200B-\u200D\uFEFF]/g, "")
                .replace(/<0x(?:200b|200c|200d|feff)>/gi, "")
                .trim();
    }

    function normalizeLookupLabel(value) {
        return normalizeText(value).replace(/\s+/g, " ").toLowerCase();
    }

    function normalizeAdoPicklistValue(value) {
        return normalizeText(value).replace(/\s+/g, " ").trim();
    }

    function normalizeOptionalNonePicklistToBlank(value) {
        const normalizedValue = normalizeAdoPicklistValue(value);
        return normalizeLookupLabel(normalizedValue) === normalizeLookupLabel(OPTIONAL_NONE_LABEL)
            ? ""
            : normalizedValue;
    }

    function describeCodePoints(value) {
        return Array.from(String(value || ""))
            .map(ch => `U+${ch.codePointAt(0).toString(16).toUpperCase().padStart(4, "0")}`)
            .join(" ");
    }

    function buildAdoFieldRefs() {
        return {
            title: normalizeText(constants.AzDoTitleFieldRef),
            reproSteps: normalizeText(constants.AzDoReproStepsFieldRef),
            tags: normalizeText(constants.AzDoTagsFieldRef),
            state: normalizeText(constants.AzDoStateFieldRef),
            severity: normalizeText(constants.AzDoSeverityFieldRef),
            priority: normalizeText(constants.AzDoPriorityFieldRef),
            areaPath: normalizeText(constants.AzDoAreaPathFieldRef),
            assignedTo: normalizeText(constants.AzDoAssignedToFieldRef),
            defectType: normalizeText(constants.AzDoDefectTypeFieldRef),
            bugStage: normalizeText(constants.AzDoBugStageFieldRef),
            createdBy: normalizeText(constants.AzDoCreatedByFieldRef),
            externalReference: normalizeText(constants.AzDoExternalReferenceFieldRef),
            rootCause: normalizeText(constants.AzDoRootCauseFieldRef),
            proposedFix: normalizeText(constants.AzDoProposedFixFieldRef),
            application: normalizeText(constants.AzDoApplicationFieldRef),
            siteName: normalizeText(constants.AzDoSiteNameFieldRef),
            iterationPath: normalizeText(constants.AzDoIterationPathFieldRef),
            resolvedReason: normalizeText(constants.AzDoResolvedReasonFieldRef),
            targetDate: normalizeText(constants.AzDoTargetDateFieldRef),
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
            "DefectAffectedReleaseFieldID",
            "DefectCreatedByFieldID",
            "DefectExternalReferenceFieldID",
            "DefectRootCauseFieldID",
            "DefectAssignedToFieldID",
            "DefectAssignedToTeamFieldID",
            "DefectTargetDateFieldID",
            "DefectWorkItemTagFieldID",
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
            "tags",
            "state",
            "severity",
            "priority",
            "areaPath",
            "assignedTo",
            "defectType",
            "bugStage",
            "createdBy",
            "externalReference",
            "rootCause",
            "proposedFix",
            "targetDate",
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

    function emitEvent(name, payload) {
        return (t = triggers.find(t => t.name === name))
            ? new Webhooks().invoke(t, payload)
            : console.error(`[ERROR]: (emitEvent) Webhook named '${name}' not found.`);
    }

    function emitFriendlyFailure(details = {}) {
        const platform = details.platform || "Unknown";
        const objectType = details.objectType || "Object";
        const objectId = details.objectId != null ? details.objectId : "Unknown";
        const objectPid = details.objectPid != null && details.objectPid !== ""
            ? ` Object PID: ${details.objectPid}.`
            : "";
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
        const objectPid = details.objectPid != null && details.objectPid !== ""
            ? ` Object PID: ${details.objectPid}.`
            : "";
        const fieldName = details.fieldName ? ` Field: ${details.fieldName}.` : "";
        const fieldValue = details.fieldValue != null && details.fieldValue !== ""
            ? ` Value: ${details.fieldValue}.`
            : "";
        const detail = details.detail || "Sync warning.";

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

    function getWorkItemTag(workItemId) {
        return `WI${workItemId}`;
    }

    function getFieldById(obj, fieldId) {
        if (!obj || !obj.properties) {
            console.log(`[Warn] Obj/properties not found.`);
            return;
        }

        const prop = obj.properties.find((p) => p.field_id == fieldId);
        if (!prop) {
            console.log(`[Warn] Property with field id '${fieldId}' not found.`);
            return;
        }

        return prop;
    }

    function normalizeAreaPathLabel(value) {
        return normalizeText(value);
    }

    function normalizeBaseUrl(value) {
        const raw = (value || "").toString().trim().replace(/\/+$/, "");
        if (!raw) {
            throw new Error("A qTest base URL is required.");
        }

        return raw.startsWith("http://") || raw.startsWith("https://")
            ? raw
            : `https://${raw}`;
    }

    function normalizeFieldResponse(data) {
        if (Array.isArray(data)) return data;
        if (Array.isArray(data?.items)) return data.items;
        if (Array.isArray(data?.data)) return data.data;
        return [];
    }

    function getAllowedValues(fieldDefinition, options = {}) {
        const values = Array.isArray(fieldDefinition?.allowed_values)
            ? fieldDefinition.allowed_values
            : [];

        return options.includeInactive
            ? values
            : values.filter(v => v?.is_active !== false);
    }

    async function getDefectFieldDefinitions() {
        const cacheKey = `${constants.ProjectID}:defects`;
        if (qtestMetadataCache[cacheKey]) {
            return qtestMetadataCache[cacheKey];
        }

        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/settings/defects/fields`;
        const response = await axios.get(url, {
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `bearer ${constants.QTEST_TOKEN}`
            }
        });

        const fields = normalizeFieldResponse(response.data);
        qtestMetadataCache[cacheKey] = fields;
        return fields;
    }

    async function getDefectFieldIdByLabel(fieldLabel) {
        const normalizedFieldLabel = normalizeLookupLabel(fieldLabel);
        if (!normalizedFieldLabel) {
            return null;
        }

        const fields = await getDefectFieldDefinitions();
        const fieldDefinition = fields.find(field => normalizeLookupLabel(field?.label) === normalizedFieldLabel);
        return fieldDefinition?.id ?? null;
    }

    async function getDefectFieldOptionLabelByValue(fieldId, rawValue) {
        if (!fieldId || rawValue === undefined || rawValue === null || rawValue === "") {
            return "";
        }

        const fields = await getDefectFieldDefinitions();
        const fieldDefinition = fields.find(field => String(field?.id) === String(fieldId));
        if (!fieldDefinition) {
            return "";
        }

        const option = getAllowedValues(fieldDefinition, { includeInactive: true })
            .find(allowedValue => String(allowedValue?.value) === String(rawValue));

        return normalizeAreaPathLabel(option?.label);
    }

    async function getDefectFieldLabel(fieldId, field) {
        if (!field) {
            return "";
        }

        const directLabel = normalizeAreaPathLabel(field.field_value_name);
        if (directLabel) {
            return directLabel;
        }

        const resolvedLabel = await getDefectFieldOptionLabelByValue(fieldId, field.field_value);
        if (resolvedLabel) {
            return resolvedLabel;
        }

        return field.field_value != null ? String(field.field_value).trim() : "";
    }

    function encodeIfNeeded(url) {
        try {
            decodeURIComponent(url);
            return url;
        } catch (e) {
            return encodeURIComponent(url);
        }
    }

    function normalizeAdoClassificationPath(value) {
        return normalizeText(value)
            .replace(/[\\/]+/g, "\\")
            .replace(/^\\+/, "")
            .replace(/\s+/g, " ")
            .trim();
    }

    function getClassificationStructuralSegmentNames(classificationType) {
        return classificationType === "areas"
            ? new Set(["area", "areas"])
            : new Set(["iteration", "iterations"]);
    }

    function buildAdoClassificationPathAliases(value, classificationType) {
        const normalizedPath = normalizeAdoClassificationPath(value);
        if (!normalizedPath) {
            return [];
        }

        const segments = normalizedPath.split("\\").filter(Boolean);
        const aliases = new Set();
        const structuralNames = getClassificationStructuralSegmentNames(classificationType);

        aliases.add(segments.join("\\"));

        if (segments.length > 1 && structuralNames.has(segments[1].toLowerCase())) {
            aliases.add([segments[0], ...segments.slice(2)].join("\\"));
        }

        if (segments.length > 1) {
            aliases.add(segments.slice(1).join("\\"));
        }

        return [...aliases].filter(Boolean);
    }

    function selectPreferredAdoFieldPath(aliases, classificationType) {
        if (!Array.isArray(aliases) || !aliases.length) {
            return "";
        }

        const structuralNames = getClassificationStructuralSegmentNames(classificationType);
        const rankedAliases = aliases
            .map(alias => {
                const normalizedAlias = normalizeAdoClassificationPath(alias);
                const segments = normalizedAlias.split("\\").filter(Boolean);
                const secondSegment = segments[1]?.toLowerCase() || "";
                const hasStructuralSecondSegment = structuralNames.has(secondSegment);

                return {
                    alias: normalizedAlias,
                    hasStructuralSecondSegment,
                    segmentCount: segments.length,
                };
            })
            .filter(item => item.alias);

        rankedAliases.sort((left, right) => {
            if (left.hasStructuralSecondSegment !== right.hasStructuralSecondSegment) {
                return left.hasStructuralSecondSegment ? 1 : -1;
            }

            return right.segmentCount - left.segmentCount;
        });

        return rankedAliases[0]?.alias || "";
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

    async function getAdoClassificationPathLookup(classificationType) {
        const cacheKey = `adoClassificationPaths:${classificationType}:${constants.AzDoProjectURL}`;
        if (qtestMetadataCache[cacheKey]) {
            return qtestMetadataCache[cacheKey];
        }

        const baseUrl = encodeIfNeeded(constants.AzDoProjectURL);
        const url = `${baseUrl}/_apis/wit/classificationnodes/${classificationType}?$depth=10&api-version=6.0`;
        const encodedToken = Buffer.from(`:${constants.AZDO_TOKEN}`).toString("base64");
        console.log(`[Debug] Fetching ADO ${classificationType} paths from '${url}'.`);

        const response = await axios.get(url, {
            headers: {
                Authorization: `Basic ${encodedToken}`,
            },
        });

        const pathLookup = new Map();
        const visitNode = (node, parentPath = "") => {
            if (!node) {
                return;
            }

            const fallbackPath = [parentPath, normalizeText(node.name)]
                .filter(Boolean)
                .join("\\");
            const candidatePaths = [node.path, fallbackPath];
            const allAliases = candidatePaths.flatMap(candidatePath =>
                buildAdoClassificationPathAliases(candidatePath, classificationType)
            );
            const preferredFieldPath = selectPreferredAdoFieldPath(allAliases, classificationType);

            for (const candidatePath of candidatePaths) {
                const aliases = buildAdoClassificationPathAliases(candidatePath, classificationType);
                for (const alias of aliases) {
                    const normalizedAlias = normalizeLookupLabel(alias);
                    if (!pathLookup.has(normalizedAlias)) {
                        pathLookup.set(normalizedAlias, preferredFieldPath || normalizeAdoClassificationPath(alias));
                    }
                }
            }

            if (Array.isArray(node.children)) {
                node.children.forEach(child => visitNode(child, preferredFieldPath || fallbackPath));
            }
        };

        visitNode(response.data);
        qtestMetadataCache[cacheKey] = pathLookup;
        return pathLookup;
    }

    async function resolveAdoClassificationPathForDefect({
        rawPath,
        classificationType,
        fallbackPath,
        fieldLabel,
        rawFieldLabel = fieldLabel,
    }) {
        const requestedPath = normalizeAdoClassificationPath(rawPath);
        const configuredDefault = normalizeAdoClassificationPath(fallbackPath);

        const targetFieldRef = classificationType === "iterations"
            ? adoFieldRefs.iterationPath
            : adoFieldRefs.areaPath;

        if (!targetFieldRef) {
            return {
                value: requestedPath || configuredDefault,
                warningDetail: null,
                warningValue: null,
            };
        }

        try {
            const pathLookup = await getAdoClassificationPathLookup(classificationType);
            const requestedAliases = buildAdoClassificationPathAliases(requestedPath, classificationType);
            const defaultAliases = buildAdoClassificationPathAliases(configuredDefault, classificationType);

            for (const requestedAlias of requestedAliases) {
                const normalizedRequested = normalizeLookupLabel(requestedAlias);
                if (pathLookup.has(normalizedRequested)) {
                    return {
                        value: pathLookup.get(normalizedRequested),
                        warningDetail: null,
                        warningValue: null,
                    };
                }
            }

            if (configuredDefault) {
                let resolvedDefault = configuredDefault;
                for (const defaultAlias of defaultAliases) {
                    const normalizedDefault = normalizeLookupLabel(defaultAlias);
                    if (pathLookup.has(normalizedDefault)) {
                        resolvedDefault = pathLookup.get(normalizedDefault);
                        break;
                    }
                }
                const warningDetail = requestedPath
                    ? `${rawFieldLabel} '${requestedPath}' was not found in Azure DevOps. Defaulted ADO ${fieldLabel} to '${resolvedDefault}'.`
                    : `${fieldLabel} was blank in qTest. Defaulted ADO ${fieldLabel} to '${resolvedDefault}'.`;

                if (requestedPath) {
                    console.log(
                        `[Debug] Requested ${classificationType} aliases: ${requestedAliases.join(" | ") || "(none)"}.`
                    );
                }
                console.log(`[Warn] ${warningDetail}`);
                return {
                    value: resolvedDefault,
                    warningDetail,
                    warningValue: requestedPath || "(blank)",
                };
            }

            if (requestedPath) {
                const warningDetail =
                    `${rawFieldLabel} '${requestedPath}' was not found in Azure DevOps, ` +
                    `and no default ${fieldLabel} constant is configured. ${fieldLabel} sync was skipped.`;
                console.log(`[Warn] ${warningDetail}`);
                return {
                    value: "",
                    warningDetail,
                    warningValue: requestedPath,
                };
            }
        } catch (error) {
            console.log(
                `[Warn] Could not validate qTest ${fieldLabel} '${requestedPath || "(blank)"}' ` +
                `against Azure DevOps ${classificationType}. ${error.message}`
            );
        }

        return {
            value: requestedPath || configuredDefault,
            warningDetail: null,
            warningValue: null,
        };
    }

    async function resolveAdoIterationPathForDefect(rawIterationPath) {
        return resolveAdoClassificationPathForDefect({
            rawPath: rawIterationPath,
            classificationType: "iterations",
            fallbackPath: constants.IterationPath || constants.AzDoDefaultIterationPath,
            fieldLabel: "Iteration Path",
        });
    }

    async function resolveAdoAreaPathForDefect(rawAreaPath) {
        return resolveAdoClassificationPathForDefect({
            rawPath: rawAreaPath,
            classificationType: "areas",
            fallbackPath: DEFAULT_AREA_PATH,
            fieldLabel: "AreaPath",
            rawFieldLabel: "ADO AreaPath",
        });
    }

    function formatDateOnly(value) {
        if (!value) return value;
        const date = new Date(value);
        if (isNaN(date.getTime())) {
            console.log(`[Warn] Could not parse Target Date '${value}'. Passing original value.`);
            return value;
        }
        return date.toISOString().slice(0, 10);
    }

    async function getDefectDetailsByIdWithRetry(defectId) {
        let defectDetails = undefined;
        let attempt = 0;

        do {
            if (attempt > 0) {
                console.log(`[Warn] Could not get defect details on attempt ${attempt}. Waiting ${retryDelayMs} ms.`);
                await new Promise((r) => setTimeout(r, retryDelayMs));
            }

            defectDetails = await getDefectDetailsById(defectId);

            if (
                defectDetails &&
                defectDetails.summary &&
                defectDetails.description &&
                defectDetails.severity
            ) {
                console.log(`[Info] Successfully fetched complete defect details for '${defectId}' on attempt ${attempt + 1}.`);
                return defectDetails;
            }

            attempt++;
        } while (attempt < retryAttemptsPerIteration);

        console.error(`[Error] Could not get defect details after retry loop. User may not have completed the initial save in qTest or the defect was abandoned.`);

        if (iteration < maxIterations) {
            iteration = iteration + 1;
            console.log(`[Info] Re-executing rule. Iteration ${iteration} of ${maxIterations}.`);
            event.iteration = iteration;
            emitEvent('qTestDefectSubmitted', event);
            return null;
        }

        emitFriendlyFailure({
            platform: "qTest",
            objectType: "Defect",
            objectId: defectId,
            detail:
                `Unable to retrieve the required defect details from qTest after approximately ` +
                `${approximateTimeoutMinutes} minutes. This usually happens when the defect form ` +
                `was left open too long before being saved or abandoned.`,
            dedupKey: `timeout:${defectId}`
        });

        return null;
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
            if (error.response && error.response.status === 404) {
                emitFriendlyFailure({
                    platform: "qTest",
                    objectType: "Defect",
                    objectId: defectId,
                    detail: "Defect was not found in qTest and may have been abandoned."
                });
                return null;
            }

            console.error("[Error] Failed to get defect by id.", error);

            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Defect",
                objectId: defectId,
                detail: "Unable to read defect details from qTest."
            });

            return null;
        }
    }

    async function getDefectDetailsById(defectId) {
        const defect = await getDefectById(defectId);
        if (!defect) return null;

        const defectPid = defect.pid || null;
        console.log(`[Info] Defect PID: ${defectPid}`);
        const summaryField = getFieldById(defect, constants.DefectSummaryFieldID);
        const descriptionField = getFieldById(defect, constants.DefectDescriptionFieldID);
        const severityField = getFieldById(defect, constants.DefectSeverityFieldID);
        const priorityField = getFieldById(defect, constants.DefectPriorityFieldID);
        const defectTypeField = getFieldById(defect, constants.DefectTypeFieldID);
        const statusField = getFieldById(defect, constants.DefectStatusFieldID);
        const affectedReleaseField = getFieldById(defect, constants.DefectAffectedReleaseFieldID);
        const createdByField = getFieldById(defect, constants.DefectCreatedByFieldID);
        const externalReferenceField = getFieldById(defect, constants.DefectExternalReferenceFieldID);
        const applicationField = getFieldById(defect, DEFECT_APPLICATION_FIELD_ID);
        const siteNameField = getFieldById(defect, DEFECT_SITE_NAME_FIELD_ID);
        const rootCauseField = constants.DefectRootCauseFieldID ? getFieldById(defect, constants.DefectRootCauseFieldID) : null;
        const proposedFixField = constants.DefectProposedFixFieldID ? getFieldById(defect, constants.DefectProposedFixFieldID) : null;
        const resolvedReasonField = constants.DefectResolvedReasonFieldID ? getFieldById(defect, constants.DefectResolvedReasonFieldID) : null;
        const assignedToField = constants.DefectAssignedToFieldID ? getFieldById(defect, constants.DefectAssignedToFieldID) : null;
        const targetDateField = getFieldById(defect, constants.DefectTargetDateFieldID);
        const assignedToTeamField = constants.DefectAssignedToTeamFieldID ? getFieldById(defect, constants.DefectAssignedToTeamFieldID) : null;
        const iterationPathField = getFieldById(defect, DEFECT_ITERATION_PATH_FIELD_ID);
        const targetDate = targetDateField ? targetDateField.field_value : null;

        console.log(`[Info] Defect Target Date: ${targetDate}`);

        if (!summaryField || !descriptionField || !severityField) {
            const missingFields = [];
            if (!summaryField) missingFields.push("Summary");
            if (!descriptionField) missingFields.push("Description");
            if (!severityField) missingFields.push("Severity");

            const missingFieldLabel = missingFields.join(", ");

            console.error(`[Error] Required defect fields not found: ${missingFieldLabel}.`);

            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Defect",
                objectId: defectId,
                objectPid: defectPid,
                fieldName: missingFieldLabel,
                detail: "Required field data is missing."
            });

            return null;
        }

        const summary = summaryField.field_value;
        console.log(`[Info] Defect summary: ${summary}`);

        const description = descriptionField.field_value;
        console.log(`[Info] Defect description: ${description}`);

        const link = defect.web_url;
        console.log(`[Info] Defect link: ${link}`);

        const severity = severityField.field_value;
        console.log(`[Info] Defect severity: ${severity}`);

        const priority = priorityField ? priorityField.field_value : null;
        console.log(`[Info] Defect priority: ${priority}`);

        const defectType = defectTypeField ? defectTypeField.field_value : null;
        console.log(`[Info] Defect type: ${defectType}`);

        const status = statusField ? statusField.field_value : null;
        const statusLabel = await getDefectFieldLabel(constants.DefectStatusFieldID, statusField);
        console.log(`[Info] Defect status: ${status} (${statusLabel || "no label"})`);

        const affectedRelease = affectedReleaseField ? affectedReleaseField.field_value : null;
        console.log(`[Info] Defect Affected Release/TestPhase: ${affectedRelease}`);

        let createdBy = createdByField ? createdByField.field_value : null;
        if (createdBy) {
            createdBy = await getQtestUserName(createdBy);
        }
        console.log(`[Info] Defect Created By: ${createdBy}`);

        const externalReference = externalReferenceField ? externalReferenceField.field_value : null;
        console.log(`[Info] Defect External Reference: ${externalReference}`);

        const application = await getDefectFieldLabel(DEFECT_APPLICATION_FIELD_ID, applicationField);
        console.log(`[Info] Defect Application: ${application}`);

        const siteName = await getDefectFieldLabel(DEFECT_SITE_NAME_FIELD_ID, siteNameField);
        console.log(`[Info] Defect Site Name: ${siteName}`);

        const iterationPath = normalizeText(await getDefectFieldLabel(DEFECT_ITERATION_PATH_FIELD_ID, iterationPathField));
        console.log(`[Info] Defect Iteration Path: ${iterationPath}`);
        if (!iterationPath) {
            console.log(`[Info] Defect Iteration Path is blank in qTest.`);
        }

        console.log(`[Debug] Root Cause target ADO field ref: '${adoFieldRefs.rootCause}'`);
        console.log(
            `[Debug] qTest Root Cause source: ${JSON.stringify({
                fieldId: constants.DefectRootCauseFieldID,
                rawValue: rootCauseField?.field_value ?? null,
                rawLabel: rootCauseField?.field_value_name ?? null,
            })}`
        );
        const rootCause = await getDefectFieldLabel(constants.DefectRootCauseFieldID, rootCauseField);
        console.log(`[Info] Defect Root Cause: ${rootCause}`);

        const proposedFix = proposedFixField ? proposedFixField.field_value : null;
        console.log(`[Info] Defect Proposed Fix length: ${proposedFix ? proposedFix.length : 0}`);

        const resolvedReason = await getDefectFieldLabel(constants.DefectResolvedReasonFieldID, resolvedReasonField);
        console.log(`[Info] Defect Resolved Reason: ${resolvedReason}`);

        let assignedToIdentity = null;
        let assignedToWarning = null;
        let assignedToWarningValue = "(blank)";
        if (assignedToField && assignedToField.field_value) {
            assignedToIdentity = await resolveQTestUserIdToIdentity(assignedToField.field_value);
            if (!assignedToIdentity) {
                assignedToIdentity = DEFAULT_ADO_ASSIGNED_TO;
                assignedToWarningValue = normalizeText(assignedToField.field_value_name || assignedToField.field_value);
                assignedToWarning =
                    `Assigned To in qTest could not be resolved to an Azure DevOps identity. ` +
                    `Defaulted ADO Assigned To to '${DEFAULT_ADO_ASSIGNED_TO}'.`;
                console.log(
                    `[Warn] Defect Assigned To user '${assignedToField.field_value}' could not be resolved. ` +
                    `Defaulting ADO Assigned To to '${DEFAULT_ADO_ASSIGNED_TO}'.`
                );
            }
        } else {
            console.log(`[Info] Defect Assigned To is blank in qTest.`);
        }
        console.log(`[Info] Defect Assigned To Identity: ${assignedToIdentity}`);

        let assignedToTeamLabel = DEFAULT_AREA_PATH;
        let assignedToTeamWarning = null;
        let assignedToTeamWarningValue = "(blank)";

        if (assignedToTeamField) {
            const rawTeamLabel = normalizeAreaPathLabel(assignedToTeamField.field_value_name);
            const rawTeamValue = assignedToTeamField.field_value;
            let resolvedAreaPath = "";

            try {
                resolvedAreaPath = await getDefectFieldOptionLabelByValue(
                    constants.DefectAssignedToTeamFieldID,
                    rawTeamValue
                );
            } catch (error) {
                console.log(
                    `[Warn] Could not resolve qTest Assigned to Team value '${rawTeamValue}' via Fields API. ` +
                    `Error: ${error.message}`
                );
            }

            assignedToTeamLabel = normalizeAreaPathLabel(resolvedAreaPath || rawTeamLabel);

            if (assignedToTeamLabel) {
                console.log(`[Info] Defect Assigned to Team resolved from qTest value '${rawTeamValue}' to ADO AreaPath '${assignedToTeamLabel}'`);
            } else {
                assignedToTeamLabel = DEFAULT_AREA_PATH;
                assignedToTeamWarningValue = `${rawTeamLabel || rawTeamValue || "(blank)"}`;
                assignedToTeamWarning =
                    `Assigned to Team value in qTest could not be resolved. Raw label='${rawTeamLabel || ""}', raw value='${rawTeamValue || ""}'. Defaulted ADO AreaPath to '${DEFAULT_AREA_PATH}'.`;

                console.log(`[Warn] ${assignedToTeamWarning}`);
            }
        } else {
            assignedToTeamLabel = DEFAULT_AREA_PATH;
            assignedToTeamWarningValue = "(blank)";
            assignedToTeamWarning =
                `Assigned to Team was blank in qTest. Defaulted ADO AreaPath to '${DEFAULT_AREA_PATH}'.`;

            console.log(`[Warn] ${assignedToTeamWarning}`);
        }

        return {
            pid: defectPid,
            summary,
            description,
            link,
            severity,
            priority,
            defectType,
            status,
            statusLabel,
            affectedRelease,
            createdBy,
            externalReference,
            application,
            siteName,
            rootCause,
            proposedFix,
            resolvedReason,
            assignedToIdentity,
            assignedToWarning,
            assignedToWarningValue,
            assignedToTeamLabel,
            assignedToTeamWarning,
            assignedToTeamWarningValue,
            targetDate,
            iterationPath
        };
    }

    function mapSeverity(qtestSeverity) {
        const severityId = parseInt(qtestSeverity);
        switch (severityId) {
            case 10899: return '1 - Critical';
            case 10302: return '2 - High';
            case 10303: return '3 - Medium';
            case 10304: return '4 - Low';
            default: return '3 - Medium';
        }
    }

    function mapPriority(qtestPriority) {
        const priorityId = parseInt(qtestPriority);
        switch (priorityId) {
            case 10898: return '4 - Critical';
            case 10204: return '3 - High';
            case 10203: return '2 - Medium';
            case 10202: return '1 - Low';
            default: return 3;
        }
    }

    function mapDefectType(qtestDefectType) {
        const id = parseInt(qtestDefectType);
        switch (id) {
            case 1751: return "New_Requirement";
            case 1752: return "Code";
            case 1753: return "Data";
            case 1754: return "Environment";
            case 1792: return "Infrastructure";
            case 1755: return "User Authorization";
            case 1756: return "Configuration";
            case 1757: return "User Handling";
            case 1758: return "Translation";
            case 1759: return "Automation";
            default: return null;
        }
    }

    function mapStatus(qtestStatus) {
        const statusId = parseInt(qtestStatus);
        switch (statusId) {
            case 10001: return "New";
            case 10003: return "In Analysis";
            case 10004: return "In Resolution";
            case 10005: return "Awaiting Implementation";
            case 10006: return "Resolved";
            case 10850: return "Retest";
            case 10852: return "Reopened";
            case 10851: return "Closed";
            case 10002: return "On Hold";
            case 10853: return "Rejected";
            case 11121: return "Triage";
            default: return "New";
        }
    }

    function mapAffectedRelease(qtestRelease) {
        const releaseId = parseInt(qtestRelease);
        switch (releaseId) {
            case -511: return null;
            case -542: return null;
            case 350: return "P&O_R1_SIT Dry Run";
            case 310: return "P&O_R1_SIT1";
            case 311: return "P&O_R1_SIT2";
            case 312: return "P&O_R1_DC1";
            case 347: return "P&O_R1_DC2";
            case 348: return "P&O_R1_DC3";
            case 351: return "P&O_R1_UAT";
            case 393: return "Unit Testing";
            case 387: return "P&O_R1.1_SIT";
            case 388: return "P&O_R1.1_UAT";
            case 428: return "P&O_R1.1_DC1";
            case 429: return "P&O_R1.1_DC2";
            case 430: return "P&O_R1.1_DC3";
            default: return null;
        }
    }

    async function getQtestUserName(userId) {
        if (!userId) return null;

        const userUrl = `https://${constants.ManagerURL}/api/v3/users/${userId}`;
        try {
            const response = await axios.get(userUrl, {
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `bearer ${constants.QTEST_TOKEN}`
                }
            });

            const userData = response.data;
            let displayName = "";

            if (userData.first_name && userData.last_name) {
                displayName = `${userData.first_name.trim()} ${userData.last_name.trim()}`;
            } else if (userData.last_name) {
                displayName = userData.last_name.trim();
            } else {
                displayName = userData.username || userData.email || userId.toString();
            }

            console.log(`[Info] Resolved qTest user ID '${userId}' to name '${displayName}'`);
            return displayName;
        } catch (error) {
            console.error(`[Error] Could not resolve qTest user ID '${userId}' to name. Using ID instead.`);
            return userId.toString();
        }
    }

    // Legacy helper kept for reference only. The active create path is
    // createAzDoBugWithFallback(), which intentionally does not push
    // Closed Date or Resolved Reason during initial defect creation.
    async function createAzDoBug(
        defectId,
        defectPid,
        name,
        description,
        link,
        qtestSeverity,
        qtestPriority,
        qtestDefectType,
        qtestStatus,
        qtestStatusLabel,
        qtestAffectedRelease,
        qtestCreatedBy,
        qtestExternalReference,
        qtestRootCause,
        qtestProposedFix,
        qtestResolvedReason,
        qtestAssignedToIdentity,
        qtestAssignedToTeamLabel,
        qtestTargetDate,
        qtestIterationPath
    ) {
        console.log(`[Info] Creating bug in Azure DevOps '${defectId}'`);

        const baseUrl = encodeIfNeeded(constants.AzDoProjectURL);
        const url = `${baseUrl}/_apis/wit/workitems/$Bug?api-version=6.0`;

        const mappedStatus = mapStatus(qtestStatus);
        console.log(`[Info] Mapped qTest Status '${qtestStatus}' (${qtestStatusLabel || "no label"}) to ADO State '${mappedStatus}'`);

        const mappedSeverity = mapSeverity(qtestSeverity);
        console.log(`[Info] Mapped severity: ${mappedSeverity}`);

        const mappedPriority = mapPriority(qtestPriority);
        console.log(`[Info] Mapped Priority: ${mappedPriority}`);

        const mappedDefectType = mapDefectType(qtestDefectType);
        console.log(`[Info] Mapped Defect Type: ${mappedDefectType}`);

        const mappedAffectedRelease = mapAffectedRelease(qtestAffectedRelease);
        if (mappedAffectedRelease) {
            console.log(`[Info] Mapped qTest Affected Release '${qtestAffectedRelease}' to ADO Bug Stage '${mappedAffectedRelease}'`);
        } else {
            console.log(`[Info] Skipping Affected Release sync — either not set or value is 'P&O Release 1'`);
        }

        const finalAreaPath = normalizeAreaPathLabel(qtestAssignedToTeamLabel) || DEFAULT_AREA_PATH;
        console.log(`[Info] Final ADO AreaPath to send: ${finalAreaPath}`);

        const requestBody = [
            buildFieldPatchOperation(adoFieldRefs.title, name),
            buildFieldPatchOperation(adoFieldRefs.reproSteps, description),
            buildFieldPatchOperation(adoFieldRefs.tags, "qTest"),
            buildFieldPatchOperation(adoFieldRefs.state, mappedStatus),
            buildFieldPatchOperation(adoFieldRefs.severity, mappedSeverity),
            buildFieldPatchOperation(adoFieldRefs.priority, mappedPriority),
            {
                op: "add",
                path: `/fields/${adoFieldRefs.areaPath}`,
                value: finalAreaPath
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

        if (qtestAssignedToIdentity && String(qtestAssignedToIdentity).trim()) {
            requestBody.push({
                op: "add",
                path: `/fields/${adoFieldRefs.assignedTo}`,
                value: String(qtestAssignedToIdentity).trim(),
            });
            console.log(`[Info] Added Assigned To to ADO: ${String(qtestAssignedToIdentity).trim()}`);
        } else {
            console.log(`[Info] Skipping Assigned To — qTest assignment is blank or could not be resolved to identity.`);
        }

        if (mappedDefectType) {
            requestBody.push({
                op: "add",
                path: `/fields/${adoFieldRefs.defectType}`,
                value: mappedDefectType,
            });
            console.log(`[Info] Added Defect Type to ADO: ${mappedDefectType}`);
        } else {
            console.log(`[Warn] Skipping Defect Type — no valid mapping found for qTest value '${qtestDefectType}'`);
        }

        if (mappedAffectedRelease) {
            requestBody.push({
                op: "add",
                path: `/fields/${adoFieldRefs.bugStage}`,
                value: mappedAffectedRelease,
            });
            console.log(`[Info] Added Bug Stage to ADO: ${mappedAffectedRelease}`);
        }

        if (qtestCreatedBy) {
            requestBody.push({
                op: "add",
                path: `/fields/${adoFieldRefs.createdBy}`,
                value: qtestCreatedBy
            });
            console.log(`[Info] Added Created By to ADO: ${qtestCreatedBy}`);
        } else {
            console.log(`[Info] Skipping Created By — no value found in qTest`);
        }

        if (qtestExternalReference) {
            requestBody.push({
                op: "add",
                path: `/fields/${adoFieldRefs.externalReference}`,
                value: qtestExternalReference
            });
            console.log(`[Info] Added External Reference to ADO: ${qtestExternalReference}`);
        } else {
            console.log(`[Info] Skipping External Reference — no value in qTest`);
        }

        const normalizedRootCause = normalizeOptionalNonePicklistToBlank(qtestRootCause);
        if (normalizedRootCause) {
            requestBody.push({
                op: "add",
                path: `/fields/${adoFieldRefs.rootCause}`,
                value: normalizedRootCause
            });
            console.log(`[Info] Added Root Cause to ADO: ${normalizedRootCause}`);
            console.log(`[Debug] Root Cause code points: ${describeCodePoints(normalizedRootCause)}`);
        } else {
            console.log(`[Info] Skipping Root Cause — no value in qTest`);
        }

        if (qtestProposedFix) {
            requestBody.push({
                op: "add",
                path: `/fields/${adoFieldRefs.proposedFix}`,
                value: qtestProposedFix
            });
            console.log(`[Info] Added Proposed Fix to ADO.`);
        } else {
            console.log(`[Info] Skipping Proposed Fix — no value in qTest`);
        }

        if (qtestClosedDate) {
            const formattedClosedDate = formatDateOnly(qtestClosedDate);
            requestBody.push({
                op: "add",
                path: `/fields/${adoFieldRefs.closedDate}`,
                value: formattedClosedDate
            });
            console.log(`[Info] Added Closed Date to ADO: ${formattedClosedDate}`);
        } else {
            console.log(`[Info] Skipping Closed Date — no value in qTest`);
        }

        if (qtestTargetDate) {
            requestBody.push({
                op: "add",
                path: `/fields/${adoFieldRefs.targetDate}`,
                value: formatDateOnly(qtestTargetDate)
            });
        }

        console.log(`[Debug] POST URL: ${url}`);
        console.log(`[Debug] Payload: ${JSON.stringify(requestBody, null, 2)}`);

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
            if (error.response) {
                console.log(`[Error] Failed to create bug in Azure DevOps. Status: ${error.response.status}`);
                console.log(`[Error] Data: ${JSON.stringify(error.response.data || null, null, 2)}`);
            } else if (error.request) {
                console.log(`[Error] No response received from ADO. Possible network or permission issue.`);
                console.log(`[Error] Request: ${JSON.stringify(error.request, null, 2)}`);
            } else {
                console.log(`[Error] Raw: ${JSON.stringify(error, null, 2)}`);
                console.log(`[Error] Message: ${error.message}`);
            }

            console.log(`[Debug] ADO Request Payload: ${JSON.stringify(requestBody, null, 2)}`);
            console.log(`[Debug] Full error object: ${JSON.stringify(error, Object.getOwnPropertyNames(error), 2)}`);

            emitFriendlyFailure({
                platform: "ADO",
                objectType: "Defect",
                objectId: defectId,
                objectPid: defectPid,
                detail: "Unable to create defect in Azure DevOps."
            });

            return null;
        }
    }

    function buildFieldPatchOperation(fieldRef, value) {
        return {
            op: "add",
            path: `/fields/${fieldRef}`,
            value,
        };
    }

    function cloneJson(value) {
        return JSON.parse(JSON.stringify(value));
    }

    function upsertAssignedToOperation(requestBody, assignedToValue) {
        const nextBody = cloneJson(requestBody);
        const assignedToPath = `/fields/${adoFieldRefs.assignedTo}`;
        const existingOp = nextBody.find(operation => operation.path === assignedToPath);

        if (assignedToValue) {
            if (existingOp) {
                existingOp.value = assignedToValue;
            } else {
                nextBody.push(buildFieldPatchOperation(adoFieldRefs.assignedTo, assignedToValue));
            }
        } else if (existingOp) {
            nextBody.splice(nextBody.indexOf(existingOp), 1);
        }

        return nextBody;
    }

    function shouldRetryAssignedToWithFallback(error, assignedToValue) {
        const status = error?.response?.status ?? error?.status;
        return Boolean(
            status === 400 &&
            normalizeText(assignedToValue) &&
            normalizeLookupLabel(assignedToValue) !== normalizeLookupLabel(DEFAULT_ADO_ASSIGNED_TO)
        );
    }

    function buildAdoTags(defectPid) {
        const tags = ["qTest"];
        const normalizedPid = normalizeText(defectPid);
        if (normalizedPid) {
            tags.push(normalizedPid);
        }
        return tags.join("; ");
    }

    async function postAdoBug(url, requestBody) {
        const response = await axios.post(url, requestBody, {
            headers: {
                'Content-Type': 'application/json-patch+json',
                'Authorization': `basic ${Buffer.from(`:${constants.AZDO_TOKEN}`).toString('base64')}`
            },
            validateStatus: () => true,
        });

        if (response.status >= 200 && response.status < 300) {
            return response;
        }

        const error = new Error(`Azure DevOps create failed with status ${response.status}`);
        error.status = response.status;
        error.response = response;
        throw error;
    }

    async function createAzDoBugWithFallback(
        defectId,
        defectPid,
        name,
        description,
        link,
        qtestSeverity,
        qtestPriority,
        qtestDefectType,
        qtestStatus,
        qtestStatusLabel,
        qtestAffectedRelease,
        qtestCreatedBy,
        qtestExternalReference,
        qtestApplication,
        qtestSiteName,
        qtestRootCause,
        qtestProposedFix,
        qtestResolvedReason,
        qtestAssignedToIdentity,
        qtestAssignedToTeamLabel,
        qtestTargetDate,
        qtestIterationPath
    ) {
        console.log(`[Info] Creating bug in Azure DevOps '${defectId}'`);

        const baseUrl = encodeIfNeeded(constants.AzDoProjectURL);
        const url = `${baseUrl}/_apis/wit/workitems/$Bug?api-version=6.0`;

        const mappedStatus = mapStatus(qtestStatus);
        console.log(`[Info] Mapped qTest Status '${qtestStatus}' (${qtestStatusLabel || "no label"}) to ADO State '${mappedStatus}'`);

        const mappedSeverity = mapSeverity(qtestSeverity);
        const mappedPriority = mapPriority(qtestPriority);
        const mappedDefectType = mapDefectType(qtestDefectType);
        const mappedAffectedRelease = mapAffectedRelease(qtestAffectedRelease);
        const sanitizedDescription = stripEmbeddedAdoLinkText(description);
        const areaPathResolution = await resolveAdoAreaPathForDefect(qtestAssignedToTeamLabel);
        const finalAreaPath = areaPathResolution.value || DEFAULT_AREA_PATH;
        const iterationPathResolution = await resolveAdoIterationPathForDefect(qtestIterationPath);
        const finalIterationPath = iterationPathResolution.value;

        console.log(`[Info] Mapped severity: ${mappedSeverity}`);
        console.log(`[Info] Mapped Priority: ${mappedPriority}`);
        console.log(`[Info] Mapped Defect Type: ${mappedDefectType}`);
        if (mappedAffectedRelease) {
            console.log(`[Info] Mapped qTest Affected Release '${qtestAffectedRelease}' to ADO Bug Stage '${mappedAffectedRelease}'`);
        } else {
            console.log(`[Info] Skipping Affected Release sync - either not set or value is 'P&O Release 1'`);
        }
        console.log(`[Info] Final ADO AreaPath to send: ${finalAreaPath}`);
        if (finalIterationPath) {
            console.log(`[Info] Final ADO Iteration Path to send: ${finalIterationPath}`);
        } else {
            console.log("[Info] Skipping Iteration Path sync - no qTest value and no default configured.");
        }
        console.log(`[Debug] Root Cause target ADO field ref: '${adoFieldRefs.rootCause}'`);

        const requestBody = [
            buildFieldPatchOperation(adoFieldRefs.title, name),
            buildFieldPatchOperation(adoFieldRefs.reproSteps, sanitizedDescription),
            buildFieldPatchOperation(adoFieldRefs.tags, buildAdoTags(defectPid)),
            buildFieldPatchOperation(adoFieldRefs.state, mappedStatus),
            buildFieldPatchOperation(adoFieldRefs.severity, mappedSeverity),
            buildFieldPatchOperation(adoFieldRefs.priority, mappedPriority),
            buildFieldPatchOperation(adoFieldRefs.areaPath, finalAreaPath),
            {
                op: "add",
                path: "/relations/-",
                value: {
                    rel: "Hyperlink",
                    url: link,
                },
            },
        ];

        const desiredAssignedTo = normalizeText(qtestAssignedToIdentity);
        if (desiredAssignedTo) {
            requestBody.push(buildFieldPatchOperation(adoFieldRefs.assignedTo, desiredAssignedTo));
            console.log(`[Info] Added Assigned To to ADO: ${desiredAssignedTo}`);
        } else {
            console.log(`[Info] Skipping Assigned To - qTest assignment is blank or could not be resolved to identity.`);
        }

        if (mappedDefectType) {
            requestBody.push(buildFieldPatchOperation(adoFieldRefs.defectType, mappedDefectType));
        }

        if (mappedAffectedRelease) {
            requestBody.push(buildFieldPatchOperation(adoFieldRefs.bugStage, mappedAffectedRelease));
        }

        if (qtestCreatedBy) {
            requestBody.push(buildFieldPatchOperation(adoFieldRefs.createdBy, qtestCreatedBy));
        }

        if (qtestExternalReference) {
            requestBody.push(buildFieldPatchOperation(adoFieldRefs.externalReference, qtestExternalReference));
        }

        if (adoFieldRefs.application && qtestApplication) {
            requestBody.push(buildFieldPatchOperation(adoFieldRefs.application, qtestApplication));
        }

        if (adoFieldRefs.siteName && qtestSiteName) {
            requestBody.push(buildFieldPatchOperation(adoFieldRefs.siteName, qtestSiteName));
        }

        if (adoFieldRefs.iterationPath && finalIterationPath) {
            requestBody.push(buildFieldPatchOperation(adoFieldRefs.iterationPath, finalIterationPath));
        }

        const normalizedRootCause = normalizeOptionalNonePicklistToBlank(qtestRootCause);
        if (normalizedRootCause) {
            requestBody.push(buildFieldPatchOperation(adoFieldRefs.rootCause, normalizedRootCause));
            console.log(`[Info] Added Root Cause to ADO: ${normalizedRootCause}`);
            console.log(`[Debug] Root Cause code points: ${describeCodePoints(normalizedRootCause)}`);
        } else {
            console.log("[Info] Skipping Root Cause - no value in qTest.");
        }

        if (qtestProposedFix) {
            requestBody.push(buildFieldPatchOperation(adoFieldRefs.proposedFix, qtestProposedFix));
        }

        if (qtestTargetDate) {
            requestBody.push(buildFieldPatchOperation(adoFieldRefs.targetDate, formatDateOnly(qtestTargetDate)));
        }

        console.log(`[Debug] POST URL: ${url}`);
        console.log(`[Debug] Payload: ${JSON.stringify(requestBody, null, 2)}`);

        try {
            let response;

            try {
                response = await postAdoBug(url, requestBody);
            } catch (error) {
                if (!shouldRetryAssignedToWithFallback(error, desiredAssignedTo)) {
                    throw error;
                }

                console.log(
                    `[Warn] Initial Azure DevOps create failed while assigning '${desiredAssignedTo}'. ` +
                    `Retrying with fallback '${DEFAULT_ADO_ASSIGNED_TO}'.`
                );

                const retryBody = upsertAssignedToOperation(requestBody, DEFAULT_ADO_ASSIGNED_TO);
                response = await postAdoBug(url, retryBody);

                return {
                    bug: response.data,
                    assignedToFallbackWarning: {
                        platform: "ADO",
                        objectType: "Defect",
                        objectId: response.data?.id ?? "Unknown",
                        objectPid: defectPid,
                        fieldName: adoFieldRefs.assignedTo,
                        fieldValue: desiredAssignedTo,
                        detail:
                            `Azure DevOps rejected the original Assigned To value. ` +
                            `Defaulted Assigned To to '${DEFAULT_ADO_ASSIGNED_TO}'.`,
                        dedupKey: `create:fallback-assigned-to:${response.data?.id ?? defectId}`,
                    },
                    iterationPathWarning: iterationPathResolution.warningDetail
                        ? {
                            platform: "ADO",
                            objectType: "Defect",
                            objectId: response.data?.id ?? "Unknown",
                            objectPid: defectPid,
                            fieldName: adoFieldRefs.iterationPath,
                            fieldValue: iterationPathResolution.warningValue,
                            detail: iterationPathResolution.warningDetail,
                            dedupKey: `create:iteration-path-warning:${response.data?.id ?? defectId}`,
                        }
                        : null,
                    areaPathWarning: areaPathResolution.warningDetail
                        ? {
                            platform: "ADO",
                            objectType: "Defect",
                            objectId: response.data?.id ?? "Unknown",
                            objectPid: defectPid,
                            fieldName: adoFieldRefs.areaPath,
                            fieldValue: areaPathResolution.warningValue,
                            detail: areaPathResolution.warningDetail,
                            dedupKey: `create:area-path-validation-warning:${response.data?.id ?? defectId}`,
                        }
                        : null,
                };
            }

            console.log(`[Info] Bug created in Azure DevOps`);
            return {
                bug: response.data,
                assignedToFallbackWarning: null,
                iterationPathWarning: iterationPathResolution.warningDetail
                    ? {
                        platform: "ADO",
                        objectType: "Defect",
                        objectId: response.data?.id ?? "Unknown",
                        objectPid: defectPid,
                        fieldName: adoFieldRefs.iterationPath,
                        fieldValue: iterationPathResolution.warningValue,
                        detail: iterationPathResolution.warningDetail,
                        dedupKey: `create:iteration-path-warning:${response.data?.id ?? defectId}`,
                    }
                    : null,
                areaPathWarning: areaPathResolution.warningDetail
                    ? {
                        platform: "ADO",
                        objectType: "Defect",
                        objectId: response.data?.id ?? "Unknown",
                        objectPid: defectPid,
                        fieldName: adoFieldRefs.areaPath,
                        fieldValue: areaPathResolution.warningValue,
                        detail: areaPathResolution.warningDetail,
                        dedupKey: `create:area-path-validation-warning:${response.data?.id ?? defectId}`,
                    }
                    : null,
            };
        } catch (error) {
            console.log(`[Error] Failed to create bug in Azure DevOps: ${error}`);
            if (error.response) {
                console.log(`[Error] Failed to create bug in Azure DevOps. Status: ${error.response.status}`);
                console.log(`[Error] Data: ${JSON.stringify(error.response.data || null, null, 2)}`);
            } else if (error.request) {
                console.log(`[Error] No response received from ADO. Possible network or permission issue.`);
                console.log(`[Error] Request: ${JSON.stringify(error.request, null, 2)}`);
            } else {
                console.log(`[Error] Raw: ${JSON.stringify(error, null, 2)}`);
                console.log(`[Error] Message: ${error.message}`);
            }

            console.log(`[Debug] ADO Request Payload: ${JSON.stringify(requestBody, null, 2)}`);
            console.log(`[Debug] Full error object: ${JSON.stringify(error, Object.getOwnPropertyNames(error), 2)}`);

            emitFriendlyFailure({
                platform: "ADO",
                objectType: "Defect",
                objectId: "Unknown",
                objectPid: defectPid,
                detail:
                    `Unable to create defect in Azure DevOps from qTest defect '${defectId}'. ` +
                    `${error?.response?.status ? `ADO returned HTTP ${error.response.status}.` : ""}`.trim(),
                dedupKey: `ado-create:${defectId}:${defectPid || "nopid"}`,
            });

            return null;
        }
    }

    async function updateDefectFields(defectId, properties, failureDetails = {}) {
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/defects/${defectId}`;
        const requestBody = {
            properties,
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
            return true;
        } catch (error) {
            console.log(`[Error] Failed to update defect '${defectId}'.`, error);

            emitFriendlyFailure({
                platform: "qTest",
                objectType: "Defect",
                objectId: defectId,
                fieldName: failureDetails.fieldName,
                fieldValue: failureDetails.fieldValue,
                detail: failureDetails.detail || "Unable to update qTest defect after Azure DevOps creation."
            });

            return false;
        }
    }

    async function resolveQTestUserIdToIdentity(userId) {
        if (!userId) return null;

        const base = (constants.qtesturl || constants.ManagerURL || "").trim();
        const normalizedBase = base.startsWith("http")
            ? base.replace(/\/+$/, "")
            : `https://${base.replace(/\/+$/, "")}`;

        const url = `${normalizedBase}/api/v3/users/${userId}`;

        console.log(`[Debug] resolveQTestUserIdToIdentity URL = ${url}`);

        try {
            const response = await axios.get(url, {
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${constants.QTEST_TOKEN}`
                }
            });

            const data = response?.data;
            if (!data) return null;

            const u = Array.isArray(data?.items) ? data.items[0] : data;
            if (!u) return null;

            const identity =
                (u.username && u.username.trim()) ||
                (u.ldap_username && u.ldap_username.trim()) ||
                (u.external_user_name && u.external_user_name.trim()) ||
                (u.external_username && u.external_username.trim()) ||
                null;

            if (identity) {
                console.log(`[Info] Resolved qTest user ID '${userId}' to identity '${identity}' (email may be blank for SSO users)`);
            } else {
                console.log(`[Warn] qTest user ID '${userId}' has no usable identity fields (email/username/ldap_username/external_*).`);
            }

            return identity;
        } catch (error) {
            const status = error?.response?.status;
            const code = error?.code;
            const msg = error?.message;

            if (status) {
                console.log(`[Warn] Could not resolve qTest user ID '${userId}' to identity. HTTP ${status}.`);
            } else {
                console.log(`[Warn] Could not resolve qTest user ID '${userId}' to identity. No HTTP response. code=${code || "n/a"} message=${msg || "n/a"} url=${url}`);
            }

            return null;
        }
    }

    const defectDetails = await getDefectDetailsByIdWithRetry(defectId);
    if (!defectDetails) return;

    const createResult = await createAzDoBugWithFallback(
        defectId,
        defectDetails.pid,
        defectDetails.summary,
        defectDetails.description,
        defectDetails.link,
        defectDetails.severity,
        defectDetails.priority,
        defectDetails.defectType,
        defectDetails.status,
        defectDetails.statusLabel,
        defectDetails.affectedRelease,
        defectDetails.createdBy,
        defectDetails.externalReference,
        defectDetails.application,
        defectDetails.siteName,
        defectDetails.rootCause,
        defectDetails.proposedFix,
        defectDetails.resolvedReason,
        defectDetails.assignedToIdentity,
        defectDetails.assignedToTeamLabel,
        defectDetails.targetDate,
        defectDetails.iterationPath
    );

    if (!createResult || !createResult.bug) return;

    const workItemId = createResult.bug.id;
    const workItemTag = getWorkItemTag(workItemId);
    const parsedWorkItemTagFieldId = parseInt(constants.DefectWorkItemTagFieldID, 10);
    const adoLinkValue = normalizeText(
        createResult?.bug?._links?.html?.href ||
        createResult?.bug?.url
    );
    const qtestLinkFieldId = normalizeText(constants.DefectLinkToAzureDevOpsFieldID)
        || String(await getDefectFieldIdByLabel(DEFECT_LINK_TO_AZURE_DEVOPS_LABEL) || "");
    let linkFieldWarningDetails = null;
    if (adoLinkValue && !qtestLinkFieldId) {
        linkFieldWarningDetails = {
            platform: "qTest",
            objectType: "Defect",
            objectId: defectId,
            objectPid: defectDetails.pid,
            fieldName: DEFECT_LINK_TO_AZURE_DEVOPS_LABEL,
            fieldValue: adoLinkValue,
            detail: `qTest field '${DEFECT_LINK_TO_AZURE_DEVOPS_LABEL}' was not found. Azure DevOps link sync was skipped.`,
        };
    }
    console.log(`[Info] Linking defect '${defectId}' with work item tag '${workItemTag}'.`);

    const defectUpdateProperties = [
        {
            field_id: Number.isNaN(parsedWorkItemTagFieldId)
                ? constants.DefectWorkItemTagFieldID
                : parsedWorkItemTagFieldId,
            field_value: workItemTag,
        },
    ];

    if (qtestLinkFieldId && adoLinkValue) {
        const parsedLinkFieldId = parseInt(qtestLinkFieldId, 10);
        defectUpdateProperties.push({
            field_id: Number.isNaN(parsedLinkFieldId) ? qtestLinkFieldId : parsedLinkFieldId,
            field_value: adoLinkValue,
        });
        console.log(`[Info] Linking defect '${defectId}' with Azure DevOps URL '${adoLinkValue}'.`);
    }

    const tagUpdated = await updateDefectFields(
        defectId,
        defectUpdateProperties,
        {
            fieldName: "Work Item Tag",
            fieldValue: workItemTag,
            detail: "Unable to update the qTest defect link metadata after Azure DevOps creation.",
        }
    );
    if (!tagUpdated) return;

    if (linkFieldWarningDetails) {
        emitFriendlyWarning(linkFieldWarningDetails);
    }

    if (defectDetails.assignedToWarning) {
        emitFriendlyWarning({
            platform: "ADO",
            objectType: "Defect",
            objectId: workItemId,
            objectPid: defectDetails.pid,
            fieldName: adoFieldRefs.assignedTo,
            fieldValue: defectDetails.assignedToWarningValue,
            detail: defectDetails.assignedToWarning,
            dedupKey: `create:assigned-to-warning:${workItemId}`,
        });
    }

    if (createResult.assignedToFallbackWarning) {
        emitFriendlyWarning(createResult.assignedToFallbackWarning);
    }

    if (createResult.areaPathWarning) {
        emitFriendlyWarning(createResult.areaPathWarning);
    }

    if (defectDetails.assignedToTeamWarning) {
        emitFriendlyWarning({
            platform: "ADO",
            objectType: "Defect",
            objectId: workItemId,
            objectPid: defectDetails.pid,
            fieldName: adoFieldRefs.areaPath,
            fieldValue: defectDetails.assignedToTeamWarningValue,
            detail: defectDetails.assignedToTeamWarning,
            dedupKey: `create:area-path-warning:${workItemId}`,
        });
    }

    if (createResult.iterationPathWarning) {
        emitFriendlyWarning(createResult.iterationPathWarning);
    }
};

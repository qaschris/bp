const { Webhooks } = require('@qasymphony/pulse-sdk');
const axios = require('axios');

exports.handler = async function ({ event, constants, triggers }, context, callback) {
  const qtestMetadataCache = {};
  let defectPid = null;
  let failureReported = false;
  const emittedMessageKeys = new Set();
  let adoFieldRefs = null;
  const DEFECT_APPLICATION_FIELD_ID = normalizeText(constants.DefectApplicationFieldID) || "1566";
  const DEFECT_SOURCE_TEAM_FIELD_ID = normalizeText(constants.DefectSourceTeamFieldID);
  const DEFECT_SITE_NAME_FIELD_ID = normalizeText(constants.DefectSiteNameFieldID) || "1569";
  const DEFECT_ITERATION_PATH_FIELD_ID = normalizeText(constants.DefectIterationPathFieldID) || "1603";

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

  function describeCodePoints(value) {
    return Array.from(String(value || ""))
      .map(ch => `U+${ch.codePointAt(0).toString(16).toUpperCase().padStart(4, "0")}`)
      .join(" ");
  }

  // Legacy helper retained only for compatibility. The active Source Team path
  // now sends and compares the raw sanitized value directly, without any dash
  // replacement logic.
  function normalizeAdoDashVariants(value) {
    return normalizeText(value)
      .replace(/â€“/gi, "–")
      .replace(/â€”/gi, "—")
      .replace(/\s+[‐‑‒–—−-]\s+/g, " – ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function normalizeComparableConstrainedLabel(value) {
    return normalizeAdoDashVariants(value).toLowerCase();
  }

  function normalizeAreaPathLabel(value) {
    return normalizeText(value);
  }

  function normalizeAdoIterationPath(value) {
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
    const normalizedPath = normalizeAdoIterationPath(value);
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
        const normalizedAlias = normalizeAdoIterationPath(alias);
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

  function extractAdoIdentityEmail(v) {
    if (!v) return "";
    if (typeof v === "string") return v.trim();

    if (typeof v === "object") {
      return normalizeText(v.uniqueName || v.userPrincipalName || v.mail || v.email || v.displayName || "");
    }
    return "";
  }

  async function getAdoWorkItem(workItemId, token, baseUrl) {
    const url = `${baseUrl}/_apis/wit/workitems/${workItemId}?api-version=6.0&$expand=Relations`;
    const encodedToken = Buffer.from(`:${token}`).toString("base64");
    const resp = await axios.get(url, {
      headers: { Authorization: `Basic ${encodedToken}` }
    });
    return resp.data;
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

    const url = `${constants.AzDoProjectURL}/_apis/wit/classificationnodes/${classificationType}?$depth=10&api-version=6.0`;
    const encodedToken = Buffer.from(`:${constants.AZDO_TOKEN}`).toString("base64");
    console.log(`[Debug] Fetching ADO ${classificationType} paths:`, url);

    const resp = await axios.get(url, {
      headers: { Authorization: `Basic ${encodedToken}` }
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
            pathLookup.set(normalizedAlias, preferredFieldPath || normalizeAdoIterationPath(alias));
          }
        }
      }

      if (Array.isArray(node.children)) {
        node.children.forEach(child => visitNode(child, preferredFieldPath || fallbackPath));
      }
    };

    visitNode(resp.data);
    qtestMetadataCache[cacheKey] = pathLookup;
    return pathLookup;
  }

  async function resolveOutboundAdoClassificationPath({
    rawPath,
    currentAdoPath,
    classificationType,
    fallbackPath,
    fieldLabel,
    rawFieldLabel = fieldLabel,
    workItemId,
    currentDefectPid
  }) {
    const requestedPath = normalizeAdoIterationPath(rawPath);
    const currentPath = normalizeAdoIterationPath(currentAdoPath);
    const configuredDefault = normalizeAdoIterationPath(fallbackPath);
    const targetFieldRef = classificationType === "iterations"
      ? adoFieldRefs.iterationPath
      : adoFieldRefs.areaPath;

    if (!targetFieldRef) {
      return {
        value: requestedPath || configuredDefault,
        warningDetails: null,
      };
    }

    try {
      const pathLookup = await getAdoClassificationPathLookup(classificationType);
      const requestedAliases = buildAdoClassificationPathAliases(requestedPath, classificationType);
      const currentAliases = buildAdoClassificationPathAliases(currentPath, classificationType);
      const defaultAliases = buildAdoClassificationPathAliases(configuredDefault, classificationType);

      for (const requestedAlias of requestedAliases) {
        const normalizedRequested = normalizeLookupLabel(requestedAlias);
        if (pathLookup.has(normalizedRequested)) {
          return {
            value: pathLookup.get(normalizedRequested),
            warningDetails: null,
          };
        }
      }

      if (requestedAliases.length && currentAliases.length) {
        const currentAliasSet = new Set(currentAliases.map(alias => normalizeLookupLabel(alias)));
        const matchesCurrentValue = requestedAliases.some(alias => currentAliasSet.has(normalizeLookupLabel(alias)));
        if (matchesCurrentValue) {
          console.log(
            `[Info] qTest ${fieldLabel} '${requestedPath}' was not recognized in the Azure DevOps ${classificationType} lookup, ` +
            `but it already matches the current ADO value '${currentPath}'. Leaving it unchanged.`
          );
          return {
            value: currentPath,
            warningDetails: null,
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
        const detail = requestedPath
          ? `${rawFieldLabel} '${requestedPath}' was not found in Azure DevOps. Defaulted ADO ${fieldLabel} to '${resolvedDefault}'.`
          : `${fieldLabel} was blank in qTest. Defaulted ADO ${fieldLabel} to '${resolvedDefault}'.`;

        if (requestedPath) {
          console.log(`[Debug] Requested ${classificationType} aliases: ${requestedAliases.join(" | ") || "(none)"}`);
          if (currentPath) {
            console.log(`[Debug] Current ADO ${fieldLabel} aliases: ${currentAliases.join(" | ") || "(none)"}`);
          }
        }
        console.log(`[Warn] ${detail}`);
        return {
          value: resolvedDefault,
          warningDetails: {
            platform: "ADO",
            objectType: "Defect",
            objectId: workItemId,
            objectPid: currentDefectPid,
            fieldName: targetFieldRef,
            fieldValue: requestedPath || "(blank)",
            detail,
            dedupKey: `update:${classificationType}-warning:${workItemId}:${normalizeLookupLabel(requestedPath || resolvedDefault)}`,
          },
        };
      }

      if (requestedPath) {
        const detail =
          `${rawFieldLabel} '${requestedPath}' was not found in Azure DevOps, ` +
          `and no default ${fieldLabel} constant is configured. ${fieldLabel} sync was skipped.`;
        console.log(`[Warn] ${detail}`);
        return {
          value: "",
          warningDetails: {
            platform: "ADO",
            objectType: "Defect",
            objectId: workItemId,
            objectPid: currentDefectPid,
            fieldName: targetFieldRef,
            fieldValue: requestedPath,
            detail,
            dedupKey: `update:${classificationType}-skip:${workItemId}:${normalizeLookupLabel(requestedPath)}`,
          },
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
      warningDetails: null,
    };
  }

  async function resolveOutboundIterationPath(rawIterationPathLabel, workItemId, currentDefectPid, currentAdoPath = "") {
    return resolveOutboundAdoClassificationPath({
      rawPath: rawIterationPathLabel,
      currentAdoPath,
      classificationType: "iterations",
      fallbackPath: constants.IterationPath || constants.AzDoDefaultIterationPath,
      fieldLabel: "Iteration Path",
      workItemId,
      currentDefectPid
    });
  }

  async function resolveOutboundAreaPath(rawAreaPath, workItemId, currentDefectPid, currentAdoPath = "") {
    return resolveOutboundAdoClassificationPath({
      rawPath: rawAreaPath,
      currentAdoPath,
      classificationType: "areas",
      fallbackPath: constants.AreaPath,
      fieldLabel: "AreaPath",
      rawFieldLabel: "ADO AreaPath",
      workItemId,
      currentDefectPid
    });
  }

  function buildAdoFieldRefs() {
    return {
      title: normalizeText(constants.AzDoTitleFieldRef),
      reproSteps: normalizeText(constants.AzDoReproStepsFieldRef),
      state: normalizeText(constants.AzDoStateFieldRef),
      severity: normalizeText(constants.AzDoSeverityFieldRef),
      priority: normalizeText(constants.AzDoPriorityFieldRef),
      defectType: normalizeText(constants.AzDoDefectTypeFieldRef),
      externalReference: normalizeText(constants.AzDoExternalReferenceFieldRef),
      bugStage: normalizeText(constants.AzDoBugStageFieldRef),
      rootCause: normalizeText(constants.AzDoRootCauseFieldRef),
      proposedFix: normalizeText(constants.AzDoProposedFixFieldRef),
      closedDate: normalizeText(constants.AzDoClosedDateFieldRef),
      resolvedReason: normalizeText(constants.AzDoResolvedReasonFieldRef),
      application: normalizeText(constants.AzDoApplicationFieldRef),
      sourceTeam: normalizeText(constants.AzDoSourceTeamFieldRef),
      siteName: normalizeText(constants.AzDoSiteNameFieldRef),
      areaPath: normalizeText(constants.AzDoAreaPathFieldRef),
      assignedTo: normalizeText(constants.AzDoAssignedToFieldRef),
      targetDate: normalizeText(constants.AzDoTargetDateFieldRef),
      iterationPath: normalizeText(constants.AzDoIterationPathFieldRef),
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
      "DefectExternalReferenceFieldID",
      "DefectRootCauseFieldID",
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
      "externalReference",
      "bugStage",
      "rootCause",
      "proposedFix",
      "closedDate",
      "resolvedReason",
      "areaPath",
      "assignedTo",
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

  function buildFieldPatchOperation(fieldRef, value) {
    return { op: "add", path: `/fields/${fieldRef}`, value };
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

  try {
    console.log("[Info] Defect update event received.");

    const iteration = event.iteration !== undefined ? event.iteration : 1;
    console.log("[Info] Iteration:", iteration);

    const DEFAULT_AREA_PATH = constants.AreaPath;
    const DEFAULT_ADO_ASSIGNED_TO = "ado-qtest-svc@bp.com";

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
        headers: { Authorization: `Bearer ${constants.QTEST_TOKEN}` }
      });

      const fields = normalizeFieldResponse(response.data);
      qtestMetadataCache[cacheKey] = fields;
      return fields;
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

      return norm(option?.label);
    }

    function stripWorkItemPrefix(value) {
      return typeof value === "string"
        ? value.replace(/^WI[-\s:]?\d+\s*:\s*/i, "").trim()
        : "";
    }

    function htmlToPlainText(htmlText) {
      if (!htmlText) return "";
      return String(htmlText)
        .replace(/<style([\s\S]*?)<\/style>/gi, "")
        .replace(/<script([\s\S]*?)<\/script>/gi, "")
        .replace(/<\/div>/gi, "\n")
        .replace(/<\/li>/gi, "\n")
        .replace(/<li>/gi, "  *  ")
        .replace(/<\/ul>/gi, "\n")
        .replace(/<\/p>/gi, "\n")
        .replace(/<br\s*[\/]?>/gi, "\n")
        .replace(/<[^>]+>/gi, "")
        .replace(/\n\s*\n/gi, "\n")
        .trim();
    }

    function formatDateOnly(value) {
      if (!value) return "";
      const date = new Date(value);
      return isNaN(date.getTime()) ? String(value).trim() : date.toISOString().slice(0, 10);
    }

    function normalizeDateOnly(value) {
      if (!value) return "";
      const date = new Date(value);
      return isNaN(date.getTime()) ? "" : date.toISOString().slice(0, 10);
    }

    function formatUtcDateTime(value) {
      if (!value) return "";
      const date = new Date(value);
      if (isNaN(date.getTime())) {
        return normalizeText(value);
      }

      const year = date.getUTCFullYear();
      const month = String(date.getUTCMonth() + 1).padStart(2, "0");
      const day = String(date.getUTCDate()).padStart(2, "0");
      const hours = String(date.getUTCHours()).padStart(2, "0");
      const minutes = String(date.getUTCMinutes()).padStart(2, "0");
      const seconds = String(date.getUTCSeconds()).padStart(2, "0");
      return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}Z`;
    }

    function normalizeUtcDateTime(value) {
      if (!value) return "";
      const date = new Date(value);
      return isNaN(date.getTime()) ? "" : formatUtcDateTime(date);
    }

    async function getDefectFieldLabel(fieldId, prop) {
      if (!prop) {
        return "";
      }

      const directLabel = normalizeText(prop.field_value_name);
      if (directLabel) {
        return directLabel;
      }

      const resolvedLabel = await getDefectFieldOptionLabelByValue(fieldId, prop.field_value);
      if (resolvedLabel) {
        return resolvedLabel;
      }

      return normalizeText(prop.field_value);
    }

    function mapSeverity(qtestSeverity) {
      const severityId = parseInt(qtestSeverity, 10);
      switch (severityId) {
        case 10301: return "1 - Critical";
        case 10302: return "2 - High";
        case 10303: return "3 - Medium";
        case 10304: return "4 - Low";
        default: return "";
      }
    }

    function mapPriority(qtestPriority) {
      const priorityId = parseInt(qtestPriority, 10);
      switch (priorityId) {
        case 11169: return 1;
        case 10204: return 2;
        case 10203: return 3;
        case 10202: return 4;
        default: return null;
      }
    }

    function mapDefectType(qtestDefectType) {
      const id = parseInt(qtestDefectType, 10);
      switch (id) {
        case 956: return "New_Requirement";
        case 957: return "Code";
        case 958: return "Data";
        case 959: return "Environment";
        case 960: return "Infrastructure";
        case 961: return "User Authorization";
        case 962: return "Configuration";
        case 963: return "User Handling";
        case 964: return "Translation";
        case 965: return "Automation";
        default: return "";
      }
    }

    function mapStatus(qtestStatus, qtestStatusLabel) {
      const normalizedStatusLabel = normalizeLookupLabel(qtestStatusLabel);
      switch (normalizedStatusLabel) {
        case "new": return "New";
        case "active": return "Active";
        case "in analysis": return "In Analysis";
        case "in resolution": return "In Resolution";
        case "awaiting implementation": return "Awaiting Implementation";
        case "resolved": return "Resolved";
        case "retest": return "Retest";
        case "reopened": return "Reopened";
        case "closed": return "Closed";
        case "on hold": return "On Hold";
        case "rejected": return "Rejected";
        case "triage": return "Triage";
        default: break;
      }

      const statusId = parseInt(qtestStatus, 10);
      switch (statusId) {
        case 10001: return "New";
        case 10002: return "In Analysis";
        case 10004: return "In Resolution";
        case 10003: return "Awaiting Implementation";
        case 10953: return "Resolved";
        case 10880: return "Retest";
        case 10882: return "Reopened";
        case 10881: return "Closed";
        case 10883: return "On Hold";
        case 10853: return "Rejected";
        case 11376: return "Triage";
        default: return "";
      }
    }

    function mapAffectedRelease(qtestRelease) {
      const releaseId = parseInt(qtestRelease, 10);
      switch (releaseId) {
        case -510: return "";
        case 283: return "P&O_R1_SIT Dry Run";
        case 279: return "P&O_R1_SIT1";
        case 280: return "P&O_R1_SIT2";
        case 284: return "P&O_R1_DC1";
        case 285: return "P&O_R1_DC2";
        case 286: return "P&O_R1_DC3";
        case 287: return "P&O_R1_UAT";
        case 302: return "Unit Testing";
        default: return "";
      }
    }

    const defectId = event.defect?.id || event.entityId;
    const projectId = event.defect?.project_id || event.projectId;

    if (!defectId || !projectId) {
      console.error("[Error] Missing defect or project ID in event.");
      emitFriendlyFailure({
        platform: "qTest",
        objectType: "Defect",
        objectId: defectId || "Unknown",
        detail: "Event did not include the required defect or project identifier."
      });
      return;
    }

    if (projectId != constants.ProjectID) {
      console.log(`[Info] Project not matching '${projectId}' != '${constants.ProjectID}', exiting.`);
      return;
    }

    if (!validateRequiredConfiguration()) {
      return;
    }

    const qTestDefectUrl = `https://${constants.ManagerURL}/api/v3/projects/${projectId}/defects/${defectId}`;
    console.log("[Info] Fetching defect details:", qTestDefectUrl);

    let defect;
    try {
      const qTestResponse = await axios.get(qTestDefectUrl, {
        headers: { Authorization: `Bearer ${constants.QTEST_TOKEN}` }
      });
      defect = qTestResponse.data;
    } catch (err) {
      console.error("[Error] Failed to fetch defect:", err.message);
      emitFriendlyFailure({
        platform: "qTest",
        objectType: "Defect",
        objectId: defectId,
        detail: "Unable to read defect details from qTest."
      });
      return;
    }

    if (constants.SyncUserRegex) {
      const updaterName = defect?.updated_by?.name || defect?.last_modified_by?.name || "";
      if (new RegExp(constants.SyncUserRegex, "i").test(updaterName)) {
        console.log("[Info] Update appears from ADO sync user; skipping to avoid loop.");
        return;
      }
    }

    const props = Array.isArray(defect.properties) ? defect.properties : [];
    const getPropById = (fid) => props.find(p => p.field_id == fid);
    const firstNonEmpty = (...vals) => vals.find(v => v && String(v).trim().length) || "";
    const norm = (s) => (typeof s === "string" ? s.trim() : s);

    const summary = firstNonEmpty(getPropById(constants.DefectSummaryFieldID)?.field_value, defect.name);
    const description = firstNonEmpty(getPropById(constants.DefectDescriptionFieldID)?.field_value, defect.description);
    defectPid = defect?.pid || null;
    console.log("[Info] qTest Defect PID:", defectPid);

    const wiRegex = /WI[-\s:]?(\d+)/i;
    let wiMatch = wiRegex.exec(summary) || wiRegex.exec(description || "");

    if (!wiMatch && props.length) {
      for (const p of props) {
        const v = firstNonEmpty(p.field_value, p.field_value_name);
        const m = v ? wiRegex.exec(v) : null;
        if (m) {
          wiMatch = m;
          break;
        }
      }
    }

    if (!wiMatch) {
      console.log("[Info] qTest defect is not yet linked to an Azure DevOps work item. Skipping update.");
      return;
    }

    const workItemId = wiMatch[1];
    console.log("[Info] Found Azure Work Item ID:", workItemId);

    let adoCurrent;
    try {
      adoCurrent = await getAdoWorkItem(workItemId, constants.AZDO_TOKEN, constants.AzDoProjectURL);
    } catch (e) {
      console.error("[Error] Failed to read ADO work item:", e.response?.data || e.message);
      emitFriendlyFailure({
        platform: "ADO",
        objectType: "Defect",
        objectId: workItemId,
        detail: "Unable to read the Azure DevOps work item."
      });
      return;
    }

    const cur = adoCurrent?.fields || {};
    const curAreaPath = norm(getAdoFieldValue(cur, adoFieldRefs.areaPath));
    const curIterationPath = norm(getAdoFieldValue(cur, adoFieldRefs.iterationPath));

    const applicationProp = getPropById(DEFECT_APPLICATION_FIELD_ID);
    const sourceTeamProp = DEFECT_SOURCE_TEAM_FIELD_ID ? getPropById(DEFECT_SOURCE_TEAM_FIELD_ID) : null;
    const siteNameProp = getPropById(DEFECT_SITE_NAME_FIELD_ID);
    const severityProp = getPropById(constants.DefectSeverityFieldID);
    const priorityProp = getPropById(constants.DefectPriorityFieldID);
    const defectTypeProp = getPropById(constants.DefectTypeFieldID);
    const statusProp = getPropById(constants.DefectStatusFieldID);
    const affectedReleaseProp = getPropById(constants.DefectAffectedReleaseFieldID);
    const externalReferenceProp = getPropById(constants.DefectExternalReferenceFieldID);
    const rootCauseProp = constants.DefectRootCauseFieldID ? getPropById(constants.DefectRootCauseFieldID) : null;
    const proposedFixProp = constants.DefectProposedFixFieldID ? getPropById(constants.DefectProposedFixFieldID) : null;
    const closedDateProp = constants.DefectClosedDateFieldID ? getPropById(constants.DefectClosedDateFieldID) : null;
    const resolvedReasonProp = constants.DefectResolvedReasonFieldID ? getPropById(constants.DefectResolvedReasonFieldID) : null;
    const assignedToProp = getPropById(constants.DefectAssignedToFieldID);
    const assignedToTeamProp = getPropById(constants.DefectAssignedToTeamFieldID);
    const targetDateProp = getPropById(constants.DefectTargetDateFieldID);
    const iterationPathProp = getPropById(DEFECT_ITERATION_PATH_FIELD_ID);
    const targetDate = norm(firstNonEmpty(targetDateProp?.field_value));

    const adoTitle = stripWorkItemPrefix(summary);
    const appLabel = normalizeText(await getDefectFieldLabel(DEFECT_APPLICATION_FIELD_ID, applicationProp));
    const srcLabel = normalizeText(await getDefectFieldLabel(DEFECT_SOURCE_TEAM_FIELD_ID, sourceTeamProp));
    const siteLabel = normalizeText(await getDefectFieldLabel(DEFECT_SITE_NAME_FIELD_ID, siteNameProp));
    const rawIterationPathLabel = normalizeText(await getDefectFieldLabel(DEFECT_ITERATION_PATH_FIELD_ID, iterationPathProp));
    const iterationPathResolution = await resolveOutboundIterationPath(rawIterationPathLabel, workItemId, defectPid, curIterationPath);
    const desiredIterationPath = iterationPathResolution.value;
    const iterationPathWarningDetails = iterationPathResolution.warningDetails;
    const assignedLabel = norm(firstNonEmpty(assignedToProp?.field_value));
    const assignedToLabel = norm(firstNonEmpty(assignedToProp?.field_value_name));
    const externalReference = norm(firstNonEmpty(externalReferenceProp?.field_value));
    const proposedFix = firstNonEmpty(proposedFixProp?.field_value);
    const closedDate = norm(firstNonEmpty(closedDateProp?.field_value));
    const mappedSeverity = mapSeverity(severityProp?.field_value);
    const mappedPriority = mapPriority(priorityProp?.field_value);
    const mappedDefectType = mapDefectType(defectTypeProp?.field_value);
    const statusLabel = await getDefectFieldLabel(constants.DefectStatusFieldID, statusProp);
    const mappedStatus = mapStatus(statusProp?.field_value, statusLabel);
    const isResolvedReasonLockedState = ["new", "active"].includes(
      normalizeLookupLabel(mappedStatus || statusLabel || "")
    );
    const mappedBugStage = mapAffectedRelease(affectedReleaseProp?.field_value);
    console.log("[Debug] Root Cause target ADO field ref:", adoFieldRefs.rootCause);
    console.log("[Debug] qTest Root Cause source:", {
      fieldId: constants.DefectRootCauseFieldID,
      rawValue: rootCauseProp?.field_value ?? null,
      rawLabel: rootCauseProp?.field_value_name ?? null,
    });
    const rootCauseLabel = normalizeAdoPicklistValue(
      await getDefectFieldLabel(constants.DefectRootCauseFieldID, rootCauseProp)
    );
    const resolvedReasonLabel = await getDefectFieldLabel(constants.DefectResolvedReasonFieldID, resolvedReasonProp);

    let userName = "";
    let assignedToWarningDetails = null;

    if (assignedLabel) {
      const userApiUrl = `https://${constants.ManagerURL}/api/v3/users/${assignedLabel}`;
      console.log("[Info] Fetching qTest user details:", userApiUrl);

      try {
        const userResp = await axios.get(userApiUrl, {
          headers: { Authorization: `Bearer ${constants.QTEST_TOKEN}` }
        });

        const u = userResp?.data || {};
        const identity =
          (u.username || "").trim() ||
          (u.ldap_username || "").trim() ||
          (u.external_user_name || "").trim();
        userName = identity;
      } catch (e) {
        console.error("[Error] Failed to fetch qTest user details:", e.response?.data || e.message);
        // Log-only: do not emit ChatOps for non-fatal assigned user lookup issues
      }

      if (!userName) {
        userName = DEFAULT_ADO_ASSIGNED_TO;
        assignedToWarningDetails = {
          platform: "ADO",
          objectType: "Defect",
          objectId: workItemId,
          objectPid: defectPid,
          fieldName: adoFieldRefs.assignedTo,
          fieldValue: assignedToLabel || assignedLabel || "(blank)",
          detail:
            `Assigned To in qTest could not be resolved to an Azure DevOps identity. ` +
            `Defaulted Assigned To to '${DEFAULT_ADO_ASSIGNED_TO}'.`,
          dedupKey: `update:assigned-to-warning:${workItemId}`,
        };
        console.log(
          `[Warn] qTest Assigned To '${assignedLabel}' could not be resolved to an ADO identity. ` +
          `Defaulting AssignedTo to '${DEFAULT_ADO_ASSIGNED_TO}'.`
        );
      }
    }

    let assignedToTeamLabel = DEFAULT_AREA_PATH;
    let assignedToTeamWarningDetails = null;

    if (assignedToTeamProp) {
      const rawTeamLabel = norm(assignedToTeamProp.field_value_name);
      const rawTeamValue = assignedToTeamProp.field_value;
      let resolvedAreaPath = "";

      try {
        resolvedAreaPath = await getDefectFieldOptionLabelByValue(
          constants.DefectAssignedToTeamFieldID,
          rawTeamValue
        );
      } catch (e) {
        console.log(
          `[Warn] Could not resolve qTest Assigned to Team value '${rawTeamValue}' via Fields API. ` +
          `Error: ${e.message}`
        );
      }

      assignedToTeamLabel = norm(resolvedAreaPath || rawTeamLabel);

      if (assignedToTeamLabel) {
        console.log(`[Info] Resolved qTest Assigned to Team value '${rawTeamValue}' to ADO AreaPath '${assignedToTeamLabel}'`);
      } else {
        assignedToTeamLabel = DEFAULT_AREA_PATH;

        const warningMessage =
          `Assigned to Team value in qTest could not be resolved. Raw label='${rawTeamLabel || ""}', raw value='${rawTeamValue || ""}'. Defaulted ADO AreaPath to '${DEFAULT_AREA_PATH}'.`;

        console.log(`[Warn] ${warningMessage}`);
        assignedToTeamWarningDetails = {
          platform: "ADO",
          objectType: "Defect",
          objectId: workItemId,
          objectPid: defectPid,
          fieldName: adoFieldRefs.areaPath,
          fieldValue: rawTeamLabel || rawTeamValue || "(blank)",
          detail: warningMessage,
          dedupKey: `update:area-path-warning:${workItemId}`,
        };
      }
    } else {
      assignedToTeamLabel = DEFAULT_AREA_PATH;

      const warningMessage =
        `Assigned to Team was blank in qTest. Defaulted ADO AreaPath to '${DEFAULT_AREA_PATH}'.`;

      console.log(`[Warn] ${warningMessage}`);
      assignedToTeamWarningDetails = {
        platform: "ADO",
        objectType: "Defect",
        objectId: workItemId,
        objectPid: defectPid,
        fieldName: adoFieldRefs.areaPath,
        fieldValue: "(blank)",
        detail: warningMessage,
        dedupKey: `update:area-path-warning:${workItemId}`,
      };
    }

    const areaPathResolution = await resolveOutboundAreaPath(assignedToTeamLabel, workItemId, defectPid, curAreaPath);
    assignedToTeamLabel = areaPathResolution.value || DEFAULT_AREA_PATH;
    if (!assignedToTeamWarningDetails && areaPathResolution.warningDetails) {
      assignedToTeamWarningDetails = areaPathResolution.warningDetails;
    }

    let isoDate;
    if (targetDate) {
      const d = new Date(targetDate);
      if (!isNaN(d.getTime())) isoDate = d.toISOString().replace(".000Z", "+00:00");
    }

    console.log("[Info] ADO Title:", adoTitle);
    console.log("[Info] Mapped Severity:", mappedSeverity);
    console.log("[Info] Mapped Priority:", mappedPriority);
    console.log("[Info] Mapped State:", mappedStatus);
    console.log("[Info] Mapped Defect Type:", mappedDefectType);
    console.log("[Info] Mapped Bug Stage:", mappedBugStage);
    console.log("[Info] Root Cause:", rootCauseLabel);
    console.log("[Info] Resolved Reason:", resolvedReasonLabel);
    console.log("[Info] Source Team:", srcLabel);
    console.log("[Info] Assigned To Username:", userName);
    console.log("[Info] Assigned To Email:", userName);
    console.log("[Info] Assigned To Label:", assignedToLabel);

    const curTitle = norm(getAdoFieldValue(cur, adoFieldRefs.title));
    const curDescription = htmlToPlainText(getAdoFieldValue(cur, adoFieldRefs.reproSteps));
    const curSeverity = norm(getAdoFieldValue(cur, adoFieldRefs.severity));
    const curPriorityRaw = getAdoFieldValue(cur, adoFieldRefs.priority);
    const curPriority = curPriorityRaw != null ? normalizeText(curPriorityRaw) : "";
    const curState = norm(getAdoFieldValue(cur, adoFieldRefs.state));
    const curDefectType = norm(getAdoFieldValue(cur, adoFieldRefs.defectType));
    const curExternalReference = norm(getAdoFieldValue(cur, adoFieldRefs.externalReference));
    const curBugStage = norm(getAdoFieldValue(cur, adoFieldRefs.bugStage));
    const curRootCause = normalizeAdoPicklistValue(getAdoFieldValue(cur, adoFieldRefs.rootCause, { preferFormatted: true }));
    const curProposedFix = htmlToPlainText(getAdoFieldValue(cur, adoFieldRefs.proposedFix));
    const curClosedDate = normalizeUtcDateTime(getAdoFieldValue(cur, adoFieldRefs.closedDate));
    const curResolvedReason = norm(getAdoFieldValue(cur, adoFieldRefs.resolvedReason, { preferFormatted: true }));
    const curApp = norm(getAdoFieldValue(cur, adoFieldRefs.application));
    const curSrc = norm(getAdoFieldValue(cur, adoFieldRefs.sourceTeam));
    const curSite = norm(getAdoFieldValue(cur, adoFieldRefs.siteName));
    const curAssignedToRaw = getAdoFieldValue(cur, adoFieldRefs.assignedTo);
    const curAssignedTo = norm(extractAdoIdentityEmail(curAssignedToRaw));
    const curTargetDate = normalizeDateOnly(getAdoFieldValue(cur, adoFieldRefs.targetDate));

    const desiredDescription = stripEmbeddedAdoLinkText(description || "");
    const desiredDescriptionPlain = htmlToPlainText(desiredDescription);
    const desiredProposedFix = proposedFix || "";
    const desiredProposedFixPlain = htmlToPlainText(desiredProposedFix);
    const desiredTargetDate = formatDateOnly(targetDate);
    const desiredClosedDate = formatUtcDateTime(closedDate);
    const shouldSyncClosedDate = ["Closed", "Rejected", "Resolved"].includes(mappedStatus);

    console.log("curTargetDate", curTargetDate);
    console.log("newTargetDate", desiredTargetDate, isoDate);

    const patchData = [];

    if (adoTitle && curTitle !== adoTitle) {
      console.log("[Info] Updating Title:", { from: curTitle || "(empty)", to: adoTitle });
      patchData.push(buildFieldPatchOperation(adoFieldRefs.title, adoTitle));
    }

    if (curDescription !== desiredDescriptionPlain) {
      console.log("[Info] Updating Repro Steps.");
      patchData.push(buildFieldPatchOperation(adoFieldRefs.reproSteps, desiredDescription));
    }

    if (mappedSeverity && curSeverity !== mappedSeverity) {
      console.log("[Info] Updating Severity:", { from: curSeverity || "(empty)", to: mappedSeverity });
      patchData.push(buildFieldPatchOperation(adoFieldRefs.severity, mappedSeverity));
    }

    if (mappedPriority != null && curPriority !== String(mappedPriority)) {
      console.log("[Info] Updating Priority:", { from: curPriority || "(empty)", to: mappedPriority });
      patchData.push(buildFieldPatchOperation(adoFieldRefs.priority, mappedPriority));
    }

    if (mappedStatus && curState !== mappedStatus) {
      console.log("[Info] Updating State:", { from: curState || "(empty)", to: mappedStatus });
      patchData.push(buildFieldPatchOperation(adoFieldRefs.state, mappedStatus));
    }

    if (mappedDefectType && curDefectType !== mappedDefectType) {
      console.log("[Info] Updating Defect Type:", { from: curDefectType || "(empty)", to: mappedDefectType });
      patchData.push(buildFieldPatchOperation(adoFieldRefs.defectType, mappedDefectType));
    }

    if (curExternalReference !== externalReference) {
      console.log("[Info] Updating External Reference:", { from: curExternalReference || "(empty)", to: externalReference || "(empty)" });
      patchData.push(buildFieldPatchOperation(adoFieldRefs.externalReference, externalReference || ""));
    }

    if (mappedBugStage && curBugStage !== mappedBugStage) {
      console.log("[Info] Updating Bug Stage:", { from: curBugStage || "(empty)", to: mappedBugStage });
      patchData.push(buildFieldPatchOperation(adoFieldRefs.bugStage, mappedBugStage));
    }

    if (rootCauseLabel && curRootCause !== rootCauseLabel) {
      console.log("[Info] Updating Root Cause:", { from: curRootCause || "(empty)", to: rootCauseLabel });
      console.log(`[Debug] Root Cause code points: ${describeCodePoints(rootCauseLabel)}`);
      patchData.push(buildFieldPatchOperation(adoFieldRefs.rootCause, rootCauseLabel));
    }

    if (curProposedFix !== desiredProposedFixPlain) {
      console.log("[Info] Updating Proposed Fix.");
      patchData.push(buildFieldPatchOperation(adoFieldRefs.proposedFix, desiredProposedFix));
    }

    if (shouldSyncClosedDate && desiredClosedDate && curClosedDate !== normalizeUtcDateTime(desiredClosedDate)) {
      console.log("[Info] Updating Closed Date:", { from: curClosedDate || "(empty)", to: desiredClosedDate });
      patchData.push(buildFieldPatchOperation(adoFieldRefs.closedDate, desiredClosedDate));
    } else if (desiredClosedDate && !shouldSyncClosedDate) {
      console.log(`[Info] Skipping Closed Date sync because outbound ADO State is '${mappedStatus || "(blank)"}'.`);
    }

    if (resolvedReasonLabel && !isResolvedReasonLockedState && curResolvedReason !== resolvedReasonLabel) {
      console.log("[Info] Updating Resolved Reason:", { from: curResolvedReason || "(empty)", to: resolvedReasonLabel });
      patchData.push(buildFieldPatchOperation(adoFieldRefs.resolvedReason, resolvedReasonLabel));
    } else if (resolvedReasonLabel && isResolvedReasonLockedState) {
      console.log(`[Info] Skipping Resolved Reason sync because outbound ADO State is '${mappedStatus || statusLabel || "(blank)"}'.`);
    }

    if (adoFieldRefs.application && appLabel && curApp !== appLabel) {
      console.log("[Info] Updating Application:", { from: curApp || "(empty)", to: appLabel });
      patchData.push(buildFieldPatchOperation(adoFieldRefs.application, appLabel));
    }

    if (
      adoFieldRefs.sourceTeam &&
      srcLabel &&
      curSrc !== srcLabel
    ) {
      console.log("[Info] Updating SourceTeam:", { from: curSrc || "(empty)", to: srcLabel });
      patchData.push(buildFieldPatchOperation(adoFieldRefs.sourceTeam, srcLabel));
    }

    if (adoFieldRefs.siteName && siteLabel && curSite !== siteLabel) {
      console.log("[Info] Updating SubEntity:", { from: curSite || "(empty)", to: siteLabel });
      patchData.push(buildFieldPatchOperation(adoFieldRefs.siteName, siteLabel));
    }

    if (assignedToTeamLabel && curAreaPath !== assignedToTeamLabel) {
      console.log("[Info] Updating AreaPath from qTest Assigned to Team:", {
        from: curAreaPath || "(empty)",
        to: assignedToTeamLabel
      });
      patchData.push(buildFieldPatchOperation(adoFieldRefs.areaPath, assignedToTeamLabel));
    }

    const desiredAssignedTo = norm(userName);
    {
      if (desiredAssignedTo) {
        if (curAssignedTo !== desiredAssignedTo) {
          console.log("[Info] Updating AssignedTo:", { from: curAssignedTo || "(unassigned)", to: desiredAssignedTo });
          patchData.push(buildFieldPatchOperation(adoFieldRefs.assignedTo, desiredAssignedTo));
        }
      } else if (curAssignedTo) {
        console.log("[Info] Clearing AssignedTo (ADO Unassigned):", { from: curAssignedTo });
        patchData.push({ op: "remove", path: `/fields/${adoFieldRefs.assignedTo}` });
      }
    }

    if (desiredTargetDate && curTargetDate !== normalizeDateOnly(desiredTargetDate)) {
      console.log("[Info] Updating TargetDate:", { from: curTargetDate || "(empty)", to: desiredTargetDate });
      patchData.push(buildFieldPatchOperation(adoFieldRefs.targetDate, desiredTargetDate));
    }

    if (adoFieldRefs.iterationPath && desiredIterationPath && curIterationPath !== desiredIterationPath) {
      console.log("[Info] Updating Iteration Path:", { from: curIterationPath || "(empty)", to: desiredIterationPath });
      patchData.push(buildFieldPatchOperation(adoFieldRefs.iterationPath, desiredIterationPath));
    }

    const backlink = defect.web_url || qTestDefectUrl;
    const hasLink = (adoCurrent?.relations || []).some(
      r => r.rel === "Hyperlink" && (r.url || "").toLowerCase() === (backlink || "").toLowerCase()
    );
    if (backlink && !hasLink) {
      patchData.push({ op: "add", path: "/relations/-", value: { rel: "Hyperlink", url: backlink } });
    }

    if (patchData.length === 0) {
      if (assignedToWarningDetails) {
        emitFriendlyWarning(assignedToWarningDetails);
      }
      if (assignedToTeamWarningDetails) {
        emitFriendlyWarning(assignedToTeamWarningDetails);
      }
      if (iterationPathWarningDetails) {
        emitFriendlyWarning(iterationPathWarningDetails);
      }
      console.log("[Info] No ADO changes detected; skipping patch (prevents loops).");
      return;
    }

    const adoPatchUrl = `${constants.AzDoProjectURL}/_apis/wit/workitems/${workItemId}?api-version=6.0`;
    const encodedToken = Buffer.from(`:${constants.AZDO_TOKEN}`).toString("base64");

    console.log("[Info] Sending update to ADO:", JSON.stringify(patchData, null, 2));

    try {
      try {
        await axios.patch(adoPatchUrl, patchData, {
          headers: {
            Authorization: `Basic ${encodedToken}`,
            "Content-Type": "application/json-patch+json"
          }
        });
      } catch (err) {
        const assignedToPath = `/fields/${adoFieldRefs.assignedTo}`;
        const canRetryAssignedTo =
          err?.response?.status === 400 &&
          normalizeText(desiredAssignedTo) &&
          normalizeLookupLabel(desiredAssignedTo) !== normalizeLookupLabel(DEFAULT_ADO_ASSIGNED_TO) &&
          patchData.some(operation => operation.path === assignedToPath);

        if (!canRetryAssignedTo) {
          throw err;
        }

        console.log(
          `[Warn] Azure DevOps update failed while assigning '${desiredAssignedTo}'. ` +
          `Retrying with fallback '${DEFAULT_ADO_ASSIGNED_TO}'.`
        );

        const retryPatchData = patchData.map(operation => {
          if (operation.path !== assignedToPath) {
            return operation;
          }

          return buildFieldPatchOperation(adoFieldRefs.assignedTo, DEFAULT_ADO_ASSIGNED_TO);
        });

        await axios.patch(adoPatchUrl, retryPatchData, {
          headers: {
            Authorization: `Basic ${encodedToken}`,
            "Content-Type": "application/json-patch+json"
          }
        });

        assignedToWarningDetails = {
          platform: "ADO",
          objectType: "Defect",
          objectId: workItemId,
          objectPid: defectPid,
          fieldName: adoFieldRefs.assignedTo,
          fieldValue: desiredAssignedTo,
          detail:
            `Azure DevOps rejected the original Assigned To value. ` +
            `Defaulted Assigned To to '${DEFAULT_ADO_ASSIGNED_TO}'.`,
          dedupKey: `update:assigned-to-warning:${workItemId}`,
        };
      }

      console.log("[Info] Successfully updated Azure DevOps work item.");
      if (assignedToWarningDetails) {
        emitFriendlyWarning(assignedToWarningDetails);
      }
      if (assignedToTeamWarningDetails) {
        emitFriendlyWarning(assignedToTeamWarningDetails);
      }
      if (iterationPathWarningDetails) {
        emitFriendlyWarning(iterationPathWarningDetails);
      }
    } catch (err) {
      console.error("[Error] Azure update failed:", err.response?.data || err.message);
      if (err?.response?.status) {
        console.error("[Error] Azure update status:", err.response.status);
      }
      if (err?.response?.data) {
        try {
          console.error("[Error] Azure update response body:", JSON.stringify(err.response.data, null, 2));
        } catch (stringifyError) {
          console.error("[Error] Azure update response body could not be stringified:", stringifyError.message);
        }
      }
      emitFriendlyFailure({
        platform: "ADO",
        objectType: "Defect",
        objectId: workItemId,
        objectPid: defectPid,
        detail: "Unable to update the Azure DevOps work item from qTest.",
        dedupKey: `ado-update:${workItemId}:${defectPid || "nopid"}`,
      });
    }

  } catch (fatal) {
    console.error("[Fatal] Unexpected error:", fatal.message);
    if (!failureReported) {
      emitFriendlyFailure({
        platform: "ADO",
        objectType: "Defect",
        objectId: event?.defect?.id || event?.entityId || "Unknown",
        objectPid: defectPid,
        detail: "Unexpected error occurred during defect sync.",
        dedupKey: `fatal:${event?.defect?.id || event?.entityId || "unknown"}`,
      });
    }
  }
};

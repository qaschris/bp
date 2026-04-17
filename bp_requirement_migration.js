const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { resolveFieldValue } = require('./qtestApiUtils');

const CONFIG = {
    QTEST_TOKEN: 'REPLACE_ME',
    ADO_TOKEN: 'REPLACE_ME',
    QTEST_BASE_URL: 'https://base.qtestnet.com',
    ADO_BASE_URL: 'https://dev.azure.com/organization/project',
    QTEST_PROJECT_ID: 123456,
    PARENT_MODULE_ID: 12345678,
    TEST_MODE: true,  // Set to true to process only a single requirement defined by TEST_REQUIREMENT_ID; false to process all requirements under PARENT_MODULE_ID
    TEST_REQUIREMENT_ID: 123456789, // Only used if TEST_MODE is true; the ID of a single requirement to process for testing purposes
    DRY_RUN: true, // Set to true to skip all write operations to qTest, allowing you to validate the script behavior without making any changes; false to allow updates to be made to qTest
    PAGE_SIZE: 100,
    LOG_DIRECTORY: './logs',
    REQUEST_TIMEOUT_MS: 60000,

    FIELD_IDS: {
        REQUIREMENT_DESCRIPTION: 0,
        REQUIREMENT_STREAM_SQUAD: 0,
        REQUIREMENT_COMPLEXITY: 0,
        REQUIREMENT_WORK_ITEM_TYPE: 0,
        REQUIREMENT_PRIORITY: 0,
        REQUIREMENT_TYPE: 0,
        REQUIREMENT_ASSIGNED_TO: 0,
        REQUIREMENT_ITERATION_PATH: 0,
        REQUIREMENT_APPLICATION_NAME: 0,
        REQUIREMENT_FIT_GAP: 0,
        REQUIREMENT_BP_ENTITY: 0,
        REQUIREMENT_ACCEPTANCE_CRITERIA: 0, // optional; leave 0 to keep acceptance criteria embedded in description only
    },

};

const STATE = {
    logFile: null,
    errorFile: null,
    moduleChildrenCache: {},
    qtestMetadataCache: {},
    counters: {
        discovered: 0,
        processed: 0,
        updated: 0,
        skipped: 0,
        failed: 0,
    },
};

function ensureLogDirectory() {
    fs.mkdirSync(CONFIG.LOG_DIRECTORY, { recursive: true });
    const stamp = new Date().toISOString().replace(/[.:]/g, '-');
    STATE.logFile = path.join(CONFIG.LOG_DIRECTORY, `bp-requirement-migration-${stamp}.log`);
    STATE.errorFile = path.join(CONFIG.LOG_DIRECTORY, `bp-requirement-migration-errors-${stamp}.log`);
}

function appendFile(filePath, message) {
    fs.appendFileSync(filePath, `${message}\n`, 'utf8');
}

function safeJson(value) {
    try {
        return JSON.stringify(value, null, 2);
    } catch (error) {
        return `[Unserializable: ${error.message}]`;
    }
}

function redactSecrets(text) {
    if (!text) return text;

    let output = String(text);
    const secrets = [CONFIG.QTEST_TOKEN, CONFIG.ADO_TOKEN].filter(Boolean);

    for (const secret of secrets) {
        if (secret && secret !== 'REPLACE_ME') {
            output = output.split(secret).join('[REDACTED]');
        }
    }

    return output;
}

function log(level, message, payload) {
    const timestamp = new Date().toISOString();
    const line = `[${timestamp}] [${level}] ${message}`;
    console.log(line);
    appendFile(STATE.logFile, line);

    if (payload !== undefined) {
        const rendered = redactSecrets(typeof payload === 'string' ? payload : safeJson(payload));
        console.log(rendered);
        appendFile(STATE.logFile, rendered);
    }
}

function logError(message, payload) {
    const timestamp = new Date().toISOString();
    const line = `[${timestamp}] [ERROR] ${message}`;
    console.error(line);
    appendFile(STATE.logFile, line);
    appendFile(STATE.errorFile, line);

    if (payload !== undefined) {
        const rendered = redactSecrets(typeof payload === 'string' ? payload : safeJson(payload));
        console.error(rendered);
        appendFile(STATE.logFile, rendered);
        appendFile(STATE.errorFile, rendered);
    }
}

function logWarning(message, payload) {
    const timestamp = new Date().toISOString();
    const line = `[${timestamp}] [WARN] ${message}`;
    console.warn(line);
    appendFile(STATE.logFile, line);

    if (payload !== undefined) {
        const rendered = redactSecrets(typeof payload === 'string' ? payload : safeJson(payload));
        console.warn(rendered);
        appendFile(STATE.logFile, rendered);
    }
}

function normalizeText(value) {
    return value == null
        ? ''
        : String(value)
            .normalize('NFKC')
            .replace(/[\u200B-\u200D\uFEFF]/g, '')
            .replace(/<0x(?:200b|200c|200d|feff)>/gi, '')
            .trim();
}

function firstNonEmpty(...values) {
    for (const value of values) {
        if (value !== undefined && value !== null && value !== '') {
            return value;
        }
    }

    return '';
}

function validateConfig() {
    const required = [
        'QTEST_TOKEN',
        'ADO_TOKEN',
        'QTEST_BASE_URL',
        'ADO_BASE_URL',
        'QTEST_PROJECT_ID',
        'PARENT_MODULE_ID',
    ];

    for (const key of required) {
        const value = CONFIG[key];
        if (value === undefined || value === null || value === '' || value === 'REPLACE_ME') {
            throw new Error(`CONFIG.${key} must be populated before running the script.`);
        }
    }

    if (CONFIG.TEST_MODE && !CONFIG.TEST_REQUIREMENT_ID) {
        throw new Error('CONFIG.TEST_REQUIREMENT_ID must be populated when CONFIG.TEST_MODE is true.');
    }
}

function sanitizeHeadersForLog(headers) {
    const clone = { ...(headers || {}) };
    if (clone.Authorization) clone.Authorization = '[REDACTED]';
    return clone;
}

async function doRequest({ method, url, headers, data, params, timeout }) {
    const requestConfig = {
        method,
        url,
        headers,
        data,
        params,
        timeout: timeout || CONFIG.REQUEST_TIMEOUT_MS,
    };

    log('INFO', `HTTP ${method} ${url}`, {
        headers: sanitizeHeadersForLog(headers),
        params,
        data,
    });

    try {
        const response = await axios(requestConfig);
        log('INFO', `HTTP ${method} succeeded with status ${response.status}`, response.data);
        return response.data;
    } catch (error) {
        const errorPayload = {
            message: error.message,
            status: error.response?.status,
            response: error.response?.data,
        };
        logError(`HTTP ${method} failed for ${url}`, errorPayload);
        throw error;
    }
}

function qTestHeaders() {
    return {
        'Content-Type': 'application/json',
        Authorization: `bearer ${CONFIG.QTEST_TOKEN}`,
    };
}

function adoHeaders() {
    return {
        'Content-Type': 'application/json',
        Authorization: `Basic ${Buffer.from(`:${CONFIG.ADO_TOKEN}`).toString('base64')}`,
    };
}

async function resolveRequirementFieldValue(fieldId, rawValue, fieldLabel, options = {}) {
    if (!fieldId || rawValue === undefined || rawValue === null || rawValue === '') {
        return null;
    }

    try {
        return await resolveFieldValue({
            baseUrl: CONFIG.QTEST_BASE_URL,
            projectId: CONFIG.QTEST_PROJECT_ID,
            objectType: 'requirements',
            fieldId,
            rawValue,
            includeInactive: options.includeInactive === true,
            headers: qTestHeaders(),
            cache: STATE.qtestMetadataCache,
            logger: message => log('INFO', message),
        });
    } catch (error) {
        if (options.emitFailure !== false) {
            logError(`Unable to resolve qTest option value for '${fieldLabel}'.`, {
                fieldId,
                fieldLabel,
                rawValue,
                error: error.message,
            });
        }
        throw error;
    }
}

async function resolveOptionalRequirementFieldValue(fieldId, rawValue, fieldLabel) {
    if (!fieldId || rawValue === undefined || rawValue === null || rawValue === '') {
        return { value: null, warning: null };
    }

    try {
        const value = await resolveRequirementFieldValue(fieldId, rawValue, fieldLabel, {
            includeInactive: true,
            emitFailure: false,
        });
        return { value, warning: null };
    } catch (error) {
        return {
            value: null,
            warning: {
                fieldName: fieldLabel,
                fieldValue: rawValue,
                detail: `${error.message} The field will be left unchanged.`,
            },
        };
    }
}

async function qtestGet(url, params) {
    return doRequest({ method: 'GET', url, headers: qTestHeaders(), params });
}

async function qtestPost(url, data, params) {
    return doRequest({ method: 'POST', url, headers: qTestHeaders(), data, params });
}

async function qtestPut(url, data, params) {
    return doRequest({ method: 'PUT', url, headers: qTestHeaders(), data, params });
}

async function adoGet(url, params) {
    return doRequest({ method: 'GET', url, headers: adoHeaders(), params });
}

function normalizeAssignedTo(value) {
    if (!value) return '';

    let resolved = value;
    if (typeof resolved === 'object') {
        resolved =
            resolved.displayName ||
            resolved.name ||
            resolved.uniqueName ||
            resolved.mail ||
            resolved.email ||
            '';
    }

    if (typeof resolved !== 'string') return '';

    return resolved
        .replace(/\s*<[^>]*>/g, '')
        .replace(/_/g, ' ')
        .trim();
}

function normalizeAreaPathSegments(areaPath) {
    if (!areaPath) return [];

    return String(areaPath)
        .split(/[\\/]+/)
        .map(segment => segment.trim())
        .filter(Boolean);
}

function getReleaseFolderName(iterationPath) {
    if (!iterationPath) return 'TBD';

    const segments = String(iterationPath)
        .split(/[\\/]+/)
        .map(segment => segment.trim())
        .filter(Boolean);

    const candidate = segments.length > 1 ? segments[1] : segments[0];
    if (!candidate) return 'TBD';

    const match = candidate.match(/P_O\s+R(\d+(?:\.\d+)?)/i);
    if (match && match[1]) {
        return `P&O Release ${match[1]}`;
    }

    return 'TBD';
}

function escapeHtml(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function extractFieldsFromAdoWorkItem(workItem) {
    return workItem?.fields || {};
}

function getAdoHtmlUrl(workItem) {
    return workItem?._links?.html?.href || `${CONFIG.ADO_BASE_URL}/_workitems/edit/${workItem?.id}`;
}

function buildRequirementDescription(workItem) {
    const fields = extractFieldsFromAdoWorkItem(workItem);

    const workItemType = fields['System.WorkItemType'] || '';
    const areaPath = fields['System.AreaPath'] || '';
    const iterationPath = fields['System.IterationPath'] || '';
    const state = fields['System.State'] || '';
    const reason = fields['System.Reason'] || '';
    const complexity = fields['Custom.Complexity'] || '';
    const acceptanceCriteria = fields['Microsoft.VSTS.Common.AcceptanceCriteria'] || '';
    const description = fields['System.Description'] || '';
    const htmlUrl = getAdoHtmlUrl(workItem);

    return `<a href="${escapeHtml(htmlUrl)}" target="_blank">Open in Azure DevOps</a><br>
<b>Type:</b> ${workItemType}<br>
<b>Area:</b> ${areaPath}<br>
<b>Iteration:</b> ${iterationPath}<br>
<b>State:</b> ${state}<br>
<b>Reason:</b> ${reason}<br>
<b>Complexity:</b> ${complexity}<br>
<b>Acceptance Criteria:</b> ${acceptanceCriteria}<br>
<b>Description:</b> ${description}`;
}

function buildRequirementName(workItemId, workItem) {
    const title = extractFieldsFromAdoWorkItem(workItem)['System.Title'] || '';
    return `WI${workItemId}: ${title}`;
}

function extractWorkItemIdFromRequirement(requirement) {
    const name = requirement?.name || '';
    const match = name.match(/^WI(\d+):/i);
    return match ? Number(match[1]) : null;
}

function buildOptionalFieldConfigWarning(fieldKey, sourceLabel, rawValue) {
    return {
        fieldName: `CONFIG.FIELD_IDS.${fieldKey}`,
        fieldValue: rawValue,
        detail: `${sourceLabel} has a source value but the qTest field id is not configured. The field will be left unchanged.`,
    };
}

function buildRequirementProperties(desiredState) {
    const properties = [];

    if (CONFIG.FIELD_IDS.REQUIREMENT_DESCRIPTION) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_DESCRIPTION,
            field_value: desiredState.description,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_STREAM_SQUAD) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_STREAM_SQUAD,
            field_value: desiredState.areaPath,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_COMPLEXITY && desiredState.complexityValue) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_COMPLEXITY,
            field_value: desiredState.complexityValue,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_WORK_ITEM_TYPE && desiredState.workItemTypeValue) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_WORK_ITEM_TYPE,
            field_value: desiredState.workItemTypeValue,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_PRIORITY && desiredState.priorityValue) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_PRIORITY,
            field_value: desiredState.priorityValue,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_TYPE && desiredState.typeValue) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_TYPE,
            field_value: desiredState.typeValue,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_ASSIGNED_TO) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_ASSIGNED_TO,
            field_value: desiredState.assignedToText || '',
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_ITERATION_PATH && desiredState.iterationPathValue) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_ITERATION_PATH,
            field_value: desiredState.iterationPathValue,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_APPLICATION_NAME && desiredState.applicationNameValue) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_APPLICATION_NAME,
            field_value: desiredState.applicationNameValue,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_FIT_GAP && desiredState.fitGapValue !== null && desiredState.fitGapValue !== undefined) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_FIT_GAP,
            field_value: desiredState.fitGapValue,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_BP_ENTITY && desiredState.bpEntityValue !== null && desiredState.bpEntityValue !== undefined) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_BP_ENTITY,
            field_value: desiredState.bpEntityValue,
        });
    }

    return properties;
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
        if (value === undefined || value === null || value === '') return '';
        return normalizeText(value).replace(/\s+/g, ' ');
    };

    return normalizeValue(left) === normalizeValue(right);
}

function evaluateRequirementUpdate(requirementDetails, desiredState) {
    const desiredProperties = buildRequirementProperties(desiredState);
    const requestBody = { name: desiredState.name, properties: desiredProperties };
    const changedFields = [];

    if (!valuesEqual(requirementDetails?.name, desiredState.name)) {
        changedFields.push('name');
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
        String(currentParentId || '') !== String(desiredState.targetModuleId || '');
    if (parentChanged) {
        changedFields.push('parentId');
    }

    return {
        needsUpdate: changedFields.length > 0,
        changedFields,
        parentChanged,
        requestBody,
    };
}

function emitWarnings(warnings, requirementId, workItemId) {
    for (const warning of warnings || []) {
        if (!warning) continue;
        logWarning(
            `Optional requirement field '${warning.fieldName}' was left unchanged for qTest requirement '${requirementId}' from ADO work item '${workItemId}'.`,
            warning
        );
    }
}

async function buildDesiredRequirementState(workItemId, workItem) {
    const fields = extractFieldsFromAdoWorkItem(workItem);
    const warnings = [];
    const adoAreaPath = fields['System.AreaPath'] || '';
    const adoIterationPath = fields['System.IterationPath'] || '';
    const adoComplexity = fields['Custom.Complexity'] || '';
    const adoWorkItemType = fields['System.WorkItemType'] || '';
    const adoPriority = fields['Microsoft.VSTS.Common.Priority'];
    const adoRequirementCategory = fields['Custom.RequirementCategory'] || '';
    const adoApplicationName = fields['Custom.ApplicationName'] || '';
    const adoFitGap = fields['BP.ERP.FitGap'];
    const adoEntity = fields['Custom.Entity'];
    const adoAssignedTo = firstNonEmpty(fields['System.AssignedTo']);

    log('INFO', `Building desired state for work item '${workItemId}'.`, {
        title: fields['System.Title'] || '',
        areaPath: adoAreaPath,
        iterationPath: adoIterationPath,
        complexity: adoComplexity,
        workItemType: adoWorkItemType,
        priority: adoPriority,
        requirementCategory: adoRequirementCategory,
        applicationName: adoApplicationName,
        fitGap: adoFitGap,
        bpEntity: adoEntity,
        assignedTo: adoAssignedTo,
    });

    const complexityValue = await resolveRequirementFieldValue(
        CONFIG.FIELD_IDS.REQUIREMENT_COMPLEXITY,
        adoComplexity,
        'Complexity'
    );
    const workItemTypeValue = await resolveRequirementFieldValue(
        CONFIG.FIELD_IDS.REQUIREMENT_WORK_ITEM_TYPE,
        adoWorkItemType,
        'Work Item Type'
    );
    const priorityValue = await resolveRequirementFieldValue(
        CONFIG.FIELD_IDS.REQUIREMENT_PRIORITY,
        adoPriority,
        'Priority'
    );
    const typeValue = await resolveRequirementFieldValue(
        CONFIG.FIELD_IDS.REQUIREMENT_TYPE,
        adoRequirementCategory,
        'Requirement Category'
    );

    const iterationResolution = await resolveOptionalRequirementFieldValue(
        CONFIG.FIELD_IDS.REQUIREMENT_ITERATION_PATH,
        adoIterationPath,
        'Iteration Path'
    );
    if (iterationResolution.warning) {
        warnings.push(iterationResolution.warning);
    }

    let applicationNameValue = null;
    if (adoApplicationName) {
        if (!CONFIG.FIELD_IDS.REQUIREMENT_APPLICATION_NAME) {
            warnings.push(buildOptionalFieldConfigWarning(
                'REQUIREMENT_APPLICATION_NAME',
                'Application Name',
                adoApplicationName
            ));
        } else {
            const resolution = await resolveOptionalRequirementFieldValue(
                CONFIG.FIELD_IDS.REQUIREMENT_APPLICATION_NAME,
                adoApplicationName,
                'Application Name'
            );
            applicationNameValue = resolution.value;
            if (resolution.warning) {
                warnings.push(resolution.warning);
            }
        }
    }

    let fitGapValue = null;
    if (adoFitGap !== undefined && adoFitGap !== null && adoFitGap !== '') {
        if (!CONFIG.FIELD_IDS.REQUIREMENT_FIT_GAP) {
            warnings.push(buildOptionalFieldConfigWarning(
                'REQUIREMENT_FIT_GAP',
                'Fit Gap',
                adoFitGap
            ));
        } else {
            const resolution = await resolveOptionalRequirementFieldValue(
                CONFIG.FIELD_IDS.REQUIREMENT_FIT_GAP,
                adoFitGap,
                'Fit Gap'
            );
            fitGapValue = resolution.value;
            if (resolution.warning) {
                warnings.push(resolution.warning);
            }
        }
    }

    let bpEntityValue = null;
    if (adoEntity !== undefined && adoEntity !== null && adoEntity !== '') {
        if (!CONFIG.FIELD_IDS.REQUIREMENT_BP_ENTITY) {
            warnings.push(buildOptionalFieldConfigWarning(
                'REQUIREMENT_BP_ENTITY',
                'BP Entity',
                adoEntity
            ));
        } else {
            const resolution = await resolveOptionalRequirementFieldValue(
                CONFIG.FIELD_IDS.REQUIREMENT_BP_ENTITY,
                adoEntity,
                'BP Entity'
            );
            bpEntityValue = resolution.value;
            if (resolution.warning) {
                warnings.push(resolution.warning);
            }
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
        assignedToText: normalizeAssignedTo(adoAssignedTo),
        iterationPathValue: iterationResolution.value,
        applicationNameValue,
        fitGapValue,
        bpEntityValue,
        targetModuleId: await ensureModulePath(adoAreaPath, adoIterationPath),
        warnings,
    };
}

async function getSubModules(parentId) {
    const cacheKey = String(parentId);
    if (STATE.moduleChildrenCache[cacheKey]) {
        return STATE.moduleChildrenCache[cacheKey];
    }

    const url = `${CONFIG.QTEST_BASE_URL}/api/v3/projects/${CONFIG.QTEST_PROJECT_ID}/modules/${parentId}`;
    const response = await qtestGet(url, { expand: 'descendants' });

    let children = [];
    if (Array.isArray(response?.children)) {
        children = response.children;
    } else if (Array.isArray(response?.items)) {
        children = response.items;
    } else if (Array.isArray(response)) {
        children = response;
    }

    STATE.moduleChildrenCache[cacheKey] = children;
    return children;
}

async function createModule(name, parentId) {
    const url = `${CONFIG.QTEST_BASE_URL}/api/v3/projects/${CONFIG.QTEST_PROJECT_ID}/modules`;
    const payload = {
        name,
        parent_id: parentId,
    };

    if (CONFIG.DRY_RUN) {
        log('INFO', `[DRY RUN] Would create module '${name}' under parent '${parentId}'.`, payload);
        return { id: `DRYRUN-${parentId}-${name}`, name };
    }

    const created = await qtestPost(url, payload);
    delete STATE.moduleChildrenCache[String(parentId)];
    return created;
}

async function ensureModulePath(areaPath, iterationPath) {
    const releaseFolderName = getReleaseFolderName(iterationPath);
    let areaSegments = normalizeAreaPathSegments(areaPath);

    if (areaSegments.length && areaSegments[0].toLowerCase() === 'bp_quantum') {
        areaSegments = areaSegments.slice(1);
    }

    const segments = [releaseFolderName, ...areaSegments];
    let currentParentId = CONFIG.PARENT_MODULE_ID;

    for (const segment of segments) {
        const children = await getSubModules(currentParentId);
        const existing = children.find(child => (child?.name || '').trim().toLowerCase() === segment.toLowerCase());

        if (existing) {
            currentParentId = existing.id;
            continue;
        }

        const created = await createModule(segment, currentParentId);
        currentParentId = created.id;
    }

    return currentParentId;
}

async function fetchRequirementById(requirementId) {
    const url = `${CONFIG.QTEST_BASE_URL}/api/v3/projects/${CONFIG.QTEST_PROJECT_ID}/requirements/${requirementId}`;
    return qtestGet(url);
}

function normalizeRequirementPage(response) {
    if (Array.isArray(response)) return response;
    if (Array.isArray(response?.items)) return response.items;
    if (Array.isArray(response?.data)) return response.data;
    return [];
}

function getTotalFromPage(response, itemsLength) {
    return Number(response?.total ?? response?.count ?? itemsLength ?? 0);
}

async function fetchRequirementsUnderParent(parentId) {
    const allItems = [];
    let page = 1;

    while (true) {
        const url = `${CONFIG.QTEST_BASE_URL}/api/v3/projects/${CONFIG.QTEST_PROJECT_ID}/requirements`;
        const response = await qtestGet(url, {
            parentId,
            page,
            size: CONFIG.PAGE_SIZE,
        });

        const items = normalizeRequirementPage(response);
        const total = getTotalFromPage(response, items.length);

        if (!items.length) {
            break;
        }

        allItems.push(...items);
        log('INFO', `Fetched page ${page} of qTest requirements. Running total: ${allItems.length}/${total || '?'}.`);

        if (items.length < CONFIG.PAGE_SIZE || (total && allItems.length >= total)) {
            break;
        }

        page += 1;
    }

    return allItems;
}

async function fetchAdoWorkItem(workItemId) {
    const url = `${CONFIG.ADO_BASE_URL}/_apis/wit/workitems/${workItemId}`;
    return adoGet(url, { 'api-version': '7.1-preview.3' });
}

async function updateRequirement(requirementDetails, desiredState, evaluation) {
    if (!evaluation.needsUpdate) {
        log('INFO', `Requirement '${requirementDetails.id}' is already in sync. Skipping update.`, {
            changedFields: evaluation.changedFields,
        });
        emitWarnings(desiredState.warnings, requirementDetails.id, desiredState.workItemId);
        return requirementDetails;
    }

    const url = `${CONFIG.QTEST_BASE_URL}/api/v3/projects/${CONFIG.QTEST_PROJECT_ID}/requirements/${requirementDetails.id}`;
    const params = evaluation.parentChanged ? { parentId: desiredState.targetModuleId } : undefined;

    if (CONFIG.DRY_RUN) {
        log('INFO', `[DRY RUN] Would update requirement '${requirementDetails.id}'.`, {
            changedFields: evaluation.changedFields,
            parentChanged: evaluation.parentChanged,
            params,
            payload: evaluation.requestBody,
        });
        emitWarnings(desiredState.warnings, requirementDetails.id, desiredState.workItemId);
        return requirementDetails;
    }

    const updated = await qtestPut(url, evaluation.requestBody, params);
    emitWarnings(desiredState.warnings, requirementDetails.id, desiredState.workItemId);
    return updated || requirementDetails;
}

async function processRequirement(requirement) {
    const requirementId = requirement?.id;
    const requirementName = requirement?.name;

    STATE.counters.processed += 1;

    log('INFO', `Processing qTest requirement '${requirementId}' - '${requirementName}'.`);

    const workItemId = extractWorkItemIdFromRequirement(requirement);
    if (!workItemId) {
        STATE.counters.skipped += 1;
        log('INFO', `Skipping requirement '${requirementId}' because no ADO work item id could be parsed from the name.`);
        return;
    }

    try {
        const workItem = await fetchAdoWorkItem(workItemId);
        const requirementDetails = await fetchRequirementById(requirementId);
        const desiredState = await buildDesiredRequirementState(workItemId, workItem);
        const evaluation = evaluateRequirementUpdate(requirementDetails, desiredState);

        await updateRequirement(requirementDetails, desiredState, evaluation);

        if (evaluation.needsUpdate) {
            STATE.counters.updated += 1;
        } else {
            STATE.counters.skipped += 1;
        }

        log('INFO', `Completed requirement '${requirementId}' using ADO work item '${workItemId}'.`, {
            changedFields: evaluation.changedFields,
            updated: evaluation.needsUpdate,
            parentChanged: evaluation.parentChanged,
        });
    } catch (error) {
        STATE.counters.failed += 1;
        logError(`Failed processing requirement '${requirementId}'.`, {
            requirementId,
            requirementName,
            workItemId,
            error: error.message,
            status: error.response?.status,
            response: error.response?.data,
        });
    }
}

async function run() {
    ensureLogDirectory();
    validateConfig();

    log('INFO', 'Starting BP requirement migration script.', {
        qtestBaseUrl: CONFIG.QTEST_BASE_URL,
        adoBaseUrl: CONFIG.ADO_BASE_URL,
        qtestProjectId: CONFIG.QTEST_PROJECT_ID,
        parentModuleId: CONFIG.PARENT_MODULE_ID,
        testMode: CONFIG.TEST_MODE,
        testRequirementId: CONFIG.TEST_REQUIREMENT_ID,
        dryRun: CONFIG.DRY_RUN,
        pageSize: CONFIG.PAGE_SIZE,
    });

    let requirements = [];

    if (CONFIG.TEST_MODE) {
        const singleRequirement = await fetchRequirementById(CONFIG.TEST_REQUIREMENT_ID);
        requirements = [singleRequirement];
        log('INFO', `Loaded single requirement '${CONFIG.TEST_REQUIREMENT_ID}' for test mode.`);
    } else {
        requirements = await fetchRequirementsUnderParent(CONFIG.PARENT_MODULE_ID);
        log('INFO', `Loaded ${requirements.length} qTest requirements under parent '${CONFIG.PARENT_MODULE_ID}'.`);
    }

    STATE.counters.discovered = requirements.length;

    for (const requirement of requirements) {
        await processRequirement(requirement);
    }

    log('INFO', 'BP requirement migration script complete.', STATE.counters);
}

run().catch(error => {
    try {
        if (!STATE.logFile || !STATE.errorFile) {
            ensureLogDirectory();
        }
        logError('Fatal script failure.', {
            message: error.message,
            stack: error.stack,
            response: error.response?.data,
        });
    } catch (loggingError) {
        console.error('Unable to write fatal error log.', loggingError);
    }

    process.exitCode = 1;
});

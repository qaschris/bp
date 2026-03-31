const fs = require('fs');
const path = require('path');
const axios = require('axios');

const CONFIG = {
    QTEST_TOKEN: 'REPLACE_ME',
    ADO_TOKEN: 'REPLACE_ME',
    QTEST_BASE_URL: 'https://base.qtestnet.com',
    ADO_BASE_URL: 'https://dev.azure.com/organization/project',
    QTEST_PROJECT_ID: 123456,
    PARENT_MODULE_ID: 12345678,
    TEST_MODE: true,
    TEST_REQUIREMENT_ID: 123456789,
    DRY_RUN: true,
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

    VALUE_MAPPINGS: {
        COMPLEXITY: {
            '1 - Very High': 941,
            '2 - High': 942,
            '3 - Medium': 943,
            '4 - Low': 944,
            '5 - Very Low': 945,
        },
        WORK_ITEM_TYPE: {
            Requirement: 1358,
        },
        PRIORITY: {
            1: 11355,
            2: 11356,
            3: 11357,
            4: 11358,
        },
        REQUIREMENT_CATEGORY: {
            'AI Deliverable': 11333,
            'Business Risk': 11334,
            'Business Value': 11335,
            Control: 11336,
            Demo: 11337,
            Legal: 11338,
            'Regulatory or Business Critical': 11339,
            Security: 11340,
            Statutory: 11341,
            Tax: 11342,
        },
    },
};

const STATE = {
    logFile: null,
    errorFile: null,
    moduleChildrenCache: {},
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

function getMappedValues(workItem) {
    const fields = extractFieldsFromAdoWorkItem(workItem);
    const complexity = fields['Custom.Complexity'] || null;
    const workItemType = fields['System.WorkItemType'] || null;
    const priority = fields['Microsoft.VSTS.Common.Priority'];
    const category = fields['Custom.RequirementCategory'] || null;

    return {
        qtestComplexityValue: CONFIG.VALUE_MAPPINGS.COMPLEXITY[complexity] || null,
        qtestWorkItemTypeValue: CONFIG.VALUE_MAPPINGS.WORK_ITEM_TYPE[workItemType] || null,
        qtestPriorityValue: CONFIG.VALUE_MAPPINGS.PRIORITY[priority] || null,
        qtestTypeValue: CONFIG.VALUE_MAPPINGS.REQUIREMENT_CATEGORY[category] || null,
    };
}

function buildRequirementProperties({
    description,
    acceptanceCriteria,
    areaPath,
    complexityValue,
    workItemTypeValue,
    priorityValue,
    typeValue,
    assignedTo,
    iterationPath,
    applicationName,
    fitGap,
    bpEntity,
}) {
    const properties = [];

    if (CONFIG.FIELD_IDS.REQUIREMENT_DESCRIPTION) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_DESCRIPTION,
            field_value: description,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_STREAM_SQUAD && areaPath) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_STREAM_SQUAD,
            field_value: areaPath,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_COMPLEXITY && complexityValue) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_COMPLEXITY,
            field_value: complexityValue,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_WORK_ITEM_TYPE && workItemTypeValue) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_WORK_ITEM_TYPE,
            field_value: workItemTypeValue,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_PRIORITY && priorityValue) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_PRIORITY,
            field_value: priorityValue,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_TYPE && typeValue) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_TYPE,
            field_value: typeValue,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_ASSIGNED_TO && assignedTo) {
        const normalizedAssignedTo = normalizeAssignedTo(assignedTo);
        if (normalizedAssignedTo) {
            properties.push({
                field_id: CONFIG.FIELD_IDS.REQUIREMENT_ASSIGNED_TO,
                field_value: normalizedAssignedTo,
            });
        }
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_ITERATION_PATH && iterationPath) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_ITERATION_PATH,
            field_value: iterationPath,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_APPLICATION_NAME && applicationName) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_APPLICATION_NAME,
            field_value: applicationName,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_FIT_GAP && fitGap !== null && fitGap !== undefined) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_FIT_GAP,
            field_value: fitGap,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_BP_ENTITY && bpEntity !== null && bpEntity !== undefined) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_BP_ENTITY,
            field_value: bpEntity,
        });
    }

    if (CONFIG.FIELD_IDS.REQUIREMENT_ACCEPTANCE_CRITERIA && acceptanceCriteria) {
        properties.push({
            field_id: CONFIG.FIELD_IDS.REQUIREMENT_ACCEPTANCE_CRITERIA,
            field_value: acceptanceCriteria,
        });
    }

    return properties;
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

async function updateRequirement(requirement, payload, targetModuleId) {
    const url = `${CONFIG.QTEST_BASE_URL}/api/v3/projects/${CONFIG.QTEST_PROJECT_ID}/requirements/${requirement.id}`;
    const params = { parentId: targetModuleId };

    if (CONFIG.DRY_RUN) {
        log('INFO', `[DRY RUN] Would update requirement '${requirement.id}' in target module '${targetModuleId}'.`, payload);
        return;
    }

    await qtestPut(url, payload, params);
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
        const fields = extractFieldsFromAdoWorkItem(workItem);
        const description = buildRequirementDescription(workItem);
        const name = buildRequirementName(workItemId, workItem);
        const mappedValues = getMappedValues(workItem);
        const areaPath = fields['System.AreaPath'] || '';
        const iterationPath = fields['System.IterationPath'] || null;
        const applicationName = fields['Custom.ApplicationName'] || null;
        const fitGap = fields['BP.ERP.FitGap'] ?? null;
        const bpEntity = fields['Custom.Entity'] ?? null;
        const acceptanceCriteria = fields['Microsoft.VSTS.Common.AcceptanceCriteria'] || '';
        const targetModuleId = await ensureModulePath(areaPath, iterationPath);

        const payload = {
            name,
            properties: buildRequirementProperties({
                description,
                acceptanceCriteria,
                areaPath,
                complexityValue: mappedValues.qtestComplexityValue,
                workItemTypeValue: mappedValues.qtestWorkItemTypeValue,
                priorityValue: mappedValues.qtestPriorityValue,
                typeValue: mappedValues.qtestTypeValue,
                assignedTo: fields['System.AssignedTo'],
                iterationPath,
                applicationName,
                fitGap,
                bpEntity,
            }),
        };

        await updateRequirement(requirement, payload, targetModuleId);
        STATE.counters.updated += 1;
        log('INFO', `Completed requirement '${requirementId}' using ADO work item '${workItemId}'.`);
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
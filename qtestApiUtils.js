const axios = require("axios");

function normalizeBaseUrl(value) {
    const raw = (value || "").toString().trim().replace(/\/+$/, "");
    if (!raw) {
        throw new Error("A qTest base URL is required.");
    }

    return raw.startsWith("http://") || raw.startsWith("https://")
        ? raw
        : `https://${raw}`;
}

function normalizeLabel(value) {
    return value == null ? "" : String(value).trim().toLowerCase();
}

function normalizeFieldResponse(data) {
    if (Array.isArray(data)) return data;
    if (Array.isArray(data?.items)) return data.items;
    if (Array.isArray(data?.data)) return data.data;
    return [];
}

function getAllowedValues(fieldDefinition, options = {}) {
    const includeInactive = options.includeInactive === true;
    return Array.isArray(fieldDefinition?.allowed_values)
        ? fieldDefinition.allowed_values.filter(v => includeInactive || v?.is_active !== false)
        : [];
}

async function getFieldDefinitions({
    baseUrl,
    projectId,
    objectType,
    headers,
    cache,
    logger,
}) {
    const cacheKey = `${projectId}:${objectType}`;
    if (cache[cacheKey]) {
        return cache[cacheKey];
    }

    const url = `${normalizeBaseUrl(baseUrl)}/api/v3/projects/${projectId}/settings/${objectType}/fields`;
    logger?.(`[Debug] Fetching qTest field definitions for '${objectType}' from '${url}'.`);

    const response = await axios.get(url, { headers });
    const fields = normalizeFieldResponse(response.data);
    cache[cacheKey] = fields;
    return fields;
}

async function getFieldDefinitionById({
    baseUrl,
    projectId,
    objectType,
    fieldId,
    headers,
    cache,
    logger,
}) {
    const fields = await getFieldDefinitions({
        baseUrl,
        projectId,
        objectType,
        headers,
        cache,
        logger,
    });

    return fields.find(field => String(field?.id) === String(fieldId)) || null;
}

async function resolveFieldValue({
    baseUrl,
    projectId,
    objectType,
    fieldId,
    rawValue,
    includeInactive,
    headers,
    cache,
    logger,
}) {
    if (rawValue === undefined || rawValue === null || rawValue === "") {
        return null;
    }

    const fieldDefinition = await getFieldDefinitionById({
        baseUrl,
        projectId,
        objectType,
        fieldId,
        headers,
        cache,
        logger,
    });

    if (!fieldDefinition) {
        throw new Error(`Field definition '${fieldId}' was not found for '${objectType}'.`);
    }

    if (!fieldDefinition.constrained) {
        return rawValue;
    }

    const allowedValues = getAllowedValues(fieldDefinition, { includeInactive });
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

async function getProjectUsers({
    baseUrl,
    projectId,
    headers,
    cache,
    logger,
}) {
    const cacheKey = `projectUsers:${projectId}`;
    if (cache[cacheKey]) {
        return cache[cacheKey];
    }

    const url = `${normalizeBaseUrl(baseUrl)}/api/v3/projects/${projectId}/users`;
    logger?.(`[Debug] Fetching active project users from '${url}?inactive=false'.`);

    const response = await axios.get(url, {
        headers,
        params: { inactive: false },
    });

    const users = normalizeFieldResponse(response.data);
    cache[cacheKey] = users;
    return users;
}

async function resolveProjectUserId({
    baseUrl,
    projectId,
    headers,
    identity,
    cache,
    logger,
}) {
    if (!identity) {
        return null;
    }

    const normalizedIdentity = normalizeLabel(identity);
    if (!normalizedIdentity) {
        return null;
    }

    const users = await getProjectUsers({
        baseUrl,
        projectId,
        headers,
        cache,
        logger,
    });

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

module.exports = {
    getFieldDefinitionById,
    getFieldDefinitions,
    getProjectUsers,
    normalizeBaseUrl,
    normalizeLabel,
    resolveFieldValue,
    resolveProjectUserId,
};

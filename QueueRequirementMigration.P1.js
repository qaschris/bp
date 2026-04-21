const axios = require("axios");
const { Webhooks } = require("@qasymphony/pulse-sdk");

exports.handler = async function ({ event, constants, triggers }, context, callback) {
    const emittedMessageKeys = new Set();
    const QUEUE_TRIGGER_NAME = "QueueRequirementMigration.P1";
    const BATCH_TRIGGER_NAME = "ProcessRequirementMigrationBatch.P1";

    function emitEvent(name, payload) {
        const trigger = triggers.find(item => item.name === name);
        return trigger
            ? new Webhooks().invoke(trigger, payload)
            : console.error(`[ERROR]: (emitEvent) Webhook named '${name}' not found.`);
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
        try { return JSON.stringify(value, null, 2); } catch (error) { return `[Unserializable: ${error.message}]`; }
    }

    function normalizeBaseUrl(value) {
        const raw = (value || "").toString().trim().replace(/\/+$/, "");
        if (!raw) throw new Error("A qTest base URL is required.");
        return raw.startsWith("http://") || raw.startsWith("https://") ? raw : `https://${raw}`;
    }

    function parsePositiveInt(value, fallback = 0) {
        const parsed = Number.parseInt(normalizeText(value), 10);
        return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
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

    function chunk(items, size) {
        const output = [];
        for (let index = 0; index < items.length; index += size) {
            output.push(items.slice(index, index + size));
        }
        return output;
    }

    function emitFriendlyFailure(details = {}) {
        const platform = details.platform || "Unknown";
        const objectType = details.objectType || "Object";
        const objectId = details.objectId != null ? details.objectId : "Unknown";
        const fieldName = details.fieldName ? ` Field: ${details.fieldName}.` : "";
        const fieldValue = details.fieldValue != null && details.fieldValue !== ""
            ? ` Value: ${details.fieldValue}.`
            : "";
        const detail = details.detail || "Sync failed.";
        const message =
            `Sync failed. Platform: ${platform}. Object Type: ${objectType}. Object ID: ${objectId}.${fieldName}${fieldValue} Detail: ${detail}`;
        const dedupKey = details.dedupKey || `failure|${message}`;

        if (emittedMessageKeys.has(dedupKey)) return false;
        emittedMessageKeys.add(dedupKey);
        console.error(`[Error] ${message}`);
        emitEvent("ChatOpsEvent", { message });
        return true;
    }

    function emitFriendlyInfo(details = {}) {
        const detail = details.detail || "Migration info.";
        const message = `Requirement migration. ${detail}`;
        const dedupKey = details.dedupKey || `info|${message}`;

        if (emittedMessageKeys.has(dedupKey)) return false;
        emittedMessageKeys.add(dedupKey);
        console.log(`[Info] ${message}`);
        emitEvent("ChatOpsEvent", { message });
        return true;
    }

    function validateRequiredConfiguration({ singleRequirementId, sourceParentId, targetParentId }) {
        const missing = [
            "QTEST_TOKEN",
            "ManagerURL",
            "ProjectID",
        ].filter(name => !normalizeText(constants[name]));

        if (!singleRequirementId && !normalizeText(sourceParentId)) {
            missing.push("event.sourceParentId");
        }

        if (!normalizeText(targetParentId)) {
            missing.push("RequirementParentID");
        }

        if (normalizeText(QUEUE_TRIGGER_NAME) && !triggers.find(item => item.name === QUEUE_TRIGGER_NAME)) {
            missing.push(`Trigger:${QUEUE_TRIGGER_NAME}`);
        }

        if (normalizeText(BATCH_TRIGGER_NAME) && !triggers.find(item => item.name === BATCH_TRIGGER_NAME)) {
            missing.push(`Trigger:${BATCH_TRIGGER_NAME}`);
        }

        if (missing.length) {
            emitFriendlyFailure({
                platform: "Pulse",
                objectType: "Configuration",
                objectId: "RequirementMigrationQueue",
                fieldName: missing.join(", "),
                detail: "Required migration queue configuration is missing in Pulse.",
                dedupKey: `requirement-migration-queue-config:${missing.join("|")}`,
            });
            return false;
        }

        return true;
    }

    async function qtestGet(url, params) {
        const headers = {
            "Content-Type": "application/json",
            Authorization: `bearer ${constants.QTEST_TOKEN}`,
        };

        console.log(`[Debug] GET ${url}`);
        console.log(`[Debug] Params: ${safeJson(params || {})}`);

        const response = await axios.get(url, { headers, params });
        console.log(`[Debug] HTTP Status: ${response.status}`);
        return response.data;
    }

    async function fetchRequirementsPage(parentId, page, size) {
        const url = `${normalizeBaseUrl(constants.ManagerURL)}/api/v3/projects/${constants.ProjectID}/requirements`;
        return qtestGet(url, {
            parentId,
            page,
            size,
        });
    }

    const singleRequirementId = parsePositiveInt(firstNonEmpty(event?.singleRequirementId, event?.requirementId), 0);
    const sourceParentId = firstNonEmpty(event?.sourceParentId);
    const targetParentId = firstNonEmpty(
        event?.targetParentId,
        constants.RequirementParentID
    );
    const page = parsePositiveInt(event?.page, 1);
    const pageSize = parsePositiveInt(event?.pageSize, 100);
    const batchSize = parsePositiveInt(event?.batchSize, 20);
    const nextBatchNumber = parsePositiveInt(event?.nextBatchNumber, 1);
    const runId = normalizeText(event?.runId) || `requirement-migration-${Date.now()}`;

    try {
        if (!validateRequiredConfiguration({ singleRequirementId, sourceParentId, targetParentId })) {
            return;
        }

        console.log(`[Info] Requirement migration queue invoked. RunId='${runId}', Page='${page}', SingleRequirementId='${singleRequirementId || ""}'.`);

        if (singleRequirementId) {
            await emitEvent(BATCH_TRIGGER_NAME, {
                runId,
                sourceParentId,
                targetParentId,
                requirementIds: [singleRequirementId],
                batchNumber: nextBatchNumber,
                batchSize,
                page: 1,
                mode: "single",
            });

            emitFriendlyInfo({
                detail: `Queued single requirement '${singleRequirementId}' for migration under run '${runId}'.`,
                dedupKey: `requirement-migration-single:${runId}:${singleRequirementId}`,
            });
            return;
        }

        const response = await fetchRequirementsPage(sourceParentId, page, pageSize);
        const requirements = normalizeRequirementPage(response);
        const total = getTotalFromPage(response, requirements.length);
        const requirementIds = requirements
            .map(item => item?.id)
            .filter(id => id !== undefined && id !== null && id !== "");
        const batches = chunk(requirementIds, batchSize);

        console.log(`[Info] Loaded ${requirementIds.length} requirement ids from page '${page}'. Total='${total || "Unknown"}'.`);

        for (let index = 0; index < batches.length; index += 1) {
            const batchNumber = nextBatchNumber + index;
            const ids = batches[index];

            await emitEvent(BATCH_TRIGGER_NAME, {
                runId,
                sourceParentId,
                targetParentId,
                requirementIds: ids,
                batchNumber,
                batchSize,
                page,
                mode: "batch",
            });
        }

        if (!requirements.length) {
            emitFriendlyInfo({
                detail: `Requirement migration queue completed for run '${runId}'. No additional requirements were found on page '${page}'.`,
                dedupKey: `requirement-migration-queue-empty:${runId}:${page}`,
            });
            return;
        }

        const loadedCount = (page - 1) * pageSize + requirements.length;
        const hasMore = requirements.length === pageSize && (!total || loadedCount < total);

        if (hasMore) {
            await emitEvent(QUEUE_TRIGGER_NAME, {
                runId,
                sourceParentId,
                targetParentId,
                page: page + 1,
                pageSize,
                batchSize,
                nextBatchNumber: nextBatchNumber + batches.length,
                mode: "queue",
            });

            console.log(`[Info] Queued next requirement migration page '${page + 1}' for run '${runId}'.`);
            return;
        }

        emitFriendlyInfo({
            detail: `Requirement migration queue finished enqueuing run '${runId}'. Queued ${loadedCount} requirement records.`,
            dedupKey: `requirement-migration-queue-finished:${runId}`,
        });
    } catch (error) {
        emitFriendlyFailure({
            platform: "Pulse",
            objectType: "RequirementMigrationQueue",
            objectId: runId || "Unknown",
            detail: error.response?.data ? safeJson(error.response.data) : error.message,
            dedupKey: `requirement-migration-queue-fatal:${runId || "unknown"}`,
        });
        throw error;
    }
};

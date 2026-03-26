const { Webhooks } = require('@qasymphony/pulse-sdk');
const axios = require('axios');

exports.handler = async function ({ event, constants, triggers }, context, callback) {
  try {
    function emitEvent(name, payload) {
        let t = triggers.find(t => t.name === name);
        return t && new Webhooks().invoke(t, payload);
    }

    console.log("[Info] Defect update event received.");

    const iteration = event.iteration !== undefined ? event.iteration : 1;
    console.log("[Info] Iteration:", iteration);

    const defectId = event.defect ?.id || event.entityId;
    const projectId = event.defect ?.project_id || event.projectId;

    if (!defectId || !projectId) {
      console.error("[Error] Missing defect or project ID in event.");
      emitEvent('ChatOpsEvent', { message: '[Error] Missing defect or project ID in event.' });
      return;
    }
    if (projectId != constants.ProjectID) {
      console.log(`[Info] Project not matching '${projectId}' != '${constants.ProjectID}', exiting.`);
      return;
    }

    // ---- Fetch qTest defect ----
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
      emitEvent('ChatOpsEvent', { message: '[Error] Failed to fetch defect.' });
      return;
    }

    // ---- Echo guard ----
    if (constants.SyncUserRegex) {
      const updaterName = defect ?.updated_by ?.name || defect ?.last_modified_by ?.name || "";
      if (new RegExp(constants.SyncUserRegex, "i").test(updaterName)) {
        console.log("[Info] Update appears from ADO sync user; skipping to avoid loop.");
        return;
      }
    }

    // ---- Extract qTest props ----
    const props = Array.isArray(defect.properties) ? defect.properties : [];
    const getPropById = (fid) => props.find(p => p.field_id == fid);
    const firstNonEmpty = (...vals) => vals.find(v => v && String(v).trim().length) || "";
    const norm = (s) => (typeof s === "string" ? s.trim() : s);

    const summary = firstNonEmpty(getPropById(constants.DefectSummaryFieldID) ?.field_value, defect.name);
    const description = firstNonEmpty(getPropById(constants.DefectDescriptionFieldID) ?.field_value, defect.description);

    // ---- Extract ADO Work Item ID ----
    const wiRegex = /WI[-\s:]?(\d+)/i;
    let wiMatch = wiRegex.exec(summary) || wiRegex.exec(description || "");
    if (!wiMatch && props.length) {
      for (const p of props) {
        const v = firstNonEmpty(p.field_value, p.field_value_name);
        const m = v ? wiRegex.exec(v) : null;
        if (m) { wiMatch = m; break; }
      }
    }
    if (!wiMatch) {
      console.error("[Error] Could not extract Azure Work Item ID.");
      emitEvent('ChatOpsEvent', { message: '[Error] Could not extract Azure Work Item ID.' });
      return;
    }
    const workItemId = wiMatch[1];
    console.log("[Info] Found Azure Work Item ID:", workItemId);

    // ---- qTest → ADO fields ----
    const applicationProp = getPropById(constants.DefectApplicationFieldID);
    const sourceTeamProp = getPropById(constants.DefectSourceTeamFieldID);
    const siteNameProp = getPropById(constants.DefectSiteNameFieldID);
    const assignedToProp = getPropById(constants.DefectAssignedToFieldID);

    const appLabel = norm(firstNonEmpty(applicationProp ?.field_value_name));
    const srcLabel = norm(firstNonEmpty(sourceTeamProp ?.field_value_name));
    const siteLabel = norm(firstNonEmpty(siteNameProp ?.field_value_name));
    const assignedLabel = norm(firstNonEmpty(assignedToProp ?.field_value));
    const assignedToLabel = norm(firstNonEmpty(assignedToProp ?.field_value_name));

    let userEmail = "";
    let userName = "";

    if (assignedLabel) {
      const userApiUrl = `https://${constants.ManagerURL}/api/v3/users/${assignedLabel}`;
      console.log("[Info] Fetching qTest user details:", userApiUrl);

      try {
        const userResp = await axios.get(userApiUrl, {
          headers: { Authorization: `Bearer ${constants.QTEST_TOKEN}` }
        });

                const u = userResp?.data || {};
        const identity = (u.username || "").trim() || (u.ldap_username || "").trim() || (u.external_user_name || "").trim();
        userEmail = (u.email || "").trim();
        userName = identity;
      } catch (e) {
        console.error("[Error] Failed to fetch qTest user details:", e.response?.data || e.message);
        emitEvent('ChatOpsEvent', { message: '[Error] Failed to fetch qTest user details.' });
      }
    }

    console.log("[Info] Assigned To Username:", userName);
    console.log("[Info] Assigned To Email:", userEmail);

    // ---- Read ADO Work Item ----
    let adoCurrent;
    try {
      adoCurrent = await getAdoWorkItem(workItemId, constants.AZDO_TOKEN, constants.AzDoProjectURL);
    } catch (e) {
      console.error("[Error] Failed to read ADO work item:", e.response ?.data || e.message);
      emitEvent('ChatOpsEvent', { message: '[Error] Failed to read ADO work item.' });
      return;
    }
    const cur = adoCurrent ?.fields || {};

    const curApp = norm(constants.AzDoApplicationFieldRef ? cur[constants.AzDoApplicationFieldRef] : "");
    const curSrc = norm(constants.AzDoSourceTeamFieldRef ? cur[constants.AzDoSourceTeamFieldRef] : "");
    const curSite = norm(constants.AzDoSiteNameFieldRef ? cur[constants.AzDoSiteNameFieldRef] : "");
    const curAssignedToRaw = cur[constants.AzDoAssignedToFieldRef];
    const curAssignedTo = norm(extractAdoIdentityEmail(curAssignedToRaw));

    // ---- Build patch ----
    const patchData = [];

    if (constants.AzDoApplicationFieldRef && appLabel && curApp !== appLabel) {
      console.log("[Info] Updating Application:", { from: curApp || "(empty)", to: appLabel });
      patchData.push({ op: "add", path: `/fields/${constants.AzDoApplicationFieldRef}`, value: appLabel });
    }

    if (constants.AzDoSourceTeamFieldRef && srcLabel && curSrc !== srcLabel) {
      console.log("[Info] Updating SourceTeam:", { from: curSrc || "(empty)", to: srcLabel });
      patchData.push({ op: "add", path: `/fields/${constants.AzDoSourceTeamFieldRef}`, value: srcLabel });
    }

    if (constants.AzDoSiteNameFieldRef && siteLabel && curSite !== siteLabel) {
      console.log("[Info] Updating SubEntity:", { from: curSite || "(empty)", to: siteLabel });
      patchData.push({ op: "add", path: `/fields/${constants.AzDoSiteNameFieldRef}`, value: siteLabel });
    }
    // ---- Assigned To (qTest -> ADO) ----
    // Requirement: match users by email; if missing/unresolved, ADO should be Unassigned.
    if (constants.AzDoAssignedToFieldRef) {
      const desiredAssignedTo = norm(userEmail || userName);

      if (desiredAssignedTo) {
        if (curAssignedTo !== desiredAssignedTo) {
          console.log("[Info] Updating AssignedTo:", { from: curAssignedTo || "(unassigned)", to: desiredAssignedTo });
          patchData.push({ op: "add", path: `/fields/${constants.AzDoAssignedToFieldRef}`, value: desiredAssignedTo });
        }
      } else if (curAssignedTo) {
        console.log("[Info] Clearing AssignedTo (ADO Unassigned):", { from: curAssignedTo });
        patchData.push({ op: "remove", path: `/fields/${constants.AzDoAssignedToFieldRef}` });
      }
    }

    const backlink = defect.web_url || qTestDefectUrl;
    const hasLink = (adoCurrent ?.relations || []).some(
      r => r.rel === "Hyperlink" && (r.url || "").toLowerCase() === (backlink || "").toLowerCase()
    );
    if (backlink && !hasLink) {
      patchData.push({ op: "add", path: "/relations/-", value: { rel: "Hyperlink", url: backlink } });
    }

    if (patchData.length === 0) {
      console.log("[Info] No ADO changes detected; skipping patch (prevents loops).");
      return;
    }

    // ---- PATCH ADO ----
    const adoPatchUrl = `${constants.AzDoProjectURL}/_apis/wit/workitems/${workItemId}?api-version=6.0`;
    const encodedToken = Buffer.from(`:${constants.AZDO_TOKEN}`).toString("base64");

    console.log("[Info] Sending update to ADO:", JSON.stringify(patchData, null, 2));

    try {
      await axios.patch(adoPatchUrl, patchData, {
        headers: {
          Authorization: `Basic ${encodedToken}`,
          "Content-Type": "application/json-patch+json"
        }
      });
      console.log("[Info] Successfully updated Azure DevOps work item.");
    } catch (err) {
      console.error("[Error] Azure update failed:", err.response ?.data || err.message);
      emitEvent('ChatOpsEvent', { message: '[Error] Azure update failed.' });
    }

  } catch (fatal) {
    console.error("[Fatal] Unexpected error:", fatal.message);
    emitEvent('ChatOpsEvent', { message: '[Fatal] Unexpected error occurred.' });
  }

  
  function extractAdoIdentityEmail(v) {
    // ADO may return identity objects for System.AssignedTo; normalize to an email/UPN string.
    if (!v) return "";
    if (typeof v === "string") return v.trim();

    if (typeof v === "object") {
      // Common identity shapes: { uniqueName }, { mail }, { email }, etc.
      return (v.uniqueName || v.mail || v.email || v.displayName || "").toString().trim();
    }
    return "";
  }

// ---- helpers ----
  async function getAdoWorkItem(workItemId, token, baseUrl) {
    const url = `${baseUrl}/_apis/wit/workitems/${workItemId}?api-version=6.0&$expand=Relations`;
    const encodedToken = Buffer.from(`:${token}`).toString("base64");
    const resp = await axios.get(url, { headers: { Authorization: `Basic ${encodedToken}` } });
    return resp.data;
  }
};

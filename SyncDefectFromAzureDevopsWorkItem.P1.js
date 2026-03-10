    const axios = require("axios"); 

    // DO NOT EDIT exported "handler" function is the entrypoint
    exports.handler = async function ({ event, constants, triggers }, context, callback) {
        function buildDefectDescription(eventData) {
            const fields = getFields(eventData);
            return `Link to Azure DevOps: ${eventData.resource._links.html.href}
    Repro steps: 
    ${htmlToPlainText(fields["Microsoft.VSTS.TCM.ReproSteps"])}`;
        }
    //     return `Link to Azure DevOps: ${eventData.resource._links.html.href}
    // Type: ${fields["System.WorkItemType"]}
    // Area: ${fields["System.AreaPath"]}
    // Iteration: ${fields["System.IterationPath"]}
    // State: ${fields["System.State"]}
    // Reason: ${fields["System.Reason"]}
    // Severity: ${fields["Microsoft.VSTS.Common.Severity"] || ""}
    // Priority: ${fields["Microsoft.VSTS.Common.Priority"] || ""}
    // Root Cause: ${fields["Microsoft.VSTS.CMMI.RootCause"] || ""}
    // Repro steps: 
    // ${htmlToPlainText(fields["Microsoft.VSTS.TCM.ReproSteps"])}
    // System info:
    // ${htmlToPlainText(fields["Microsoft.VSTS.TCM.SystemInfo"])}
    // Acceptance criteria:
    // ${htmlToPlainText(fields["Microsoft.VSTS.Common.AcceptanceCriteria"])}`;
    //     }

        function buildDefectSummary(namePrefix, eventData) {
            const fields = getFields(eventData);
            return `${namePrefix}${fields["System.Title"]}`;
        }

        function getFields(eventData) {
            // In case of update the fields can be taken from the revision, in case of create from the resource directly
            return eventData.resource.revision ? eventData.resource.revision.fields : eventData.resource.fields;
        }

        function extractUpnOrEmailFromAdoAssignedTo(raw) {
            // Handles ADO identity object OR string formats.
            // Returns a normalized email/UPN string, or null if not resolvable.

            if (!raw) return null;

            // If ADO sends an identity object, prefer uniqueName (usually UPN/email).
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

            // Common ADO display format: "Last, First (ORG) <user@domain>"
            const angle = s.match(/<([^>]+)>/);
            if (angle && angle[1]) return angle[1].trim();

            // If no angle brackets, try to find an email anywhere in the string.
            const email = s.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
            if (email && email[0]) return email[0].trim();

            // Fallback: sometimes it might already just be a UPN-ish string
            return s || null;
        }

        async function resolveQtestUserIdByUsernameOrUpn(identity, standardHeaders) {
            if (!identity) return null;

            const url = `https://${constants.ManagerURL}/api/v3/users/search?username=${encodeURIComponent(identity)}`;

            try {
                const resp = await axios.get(url, { headers: standardHeaders });
                const data = resp && resp.data;

                // API shape can vary: sometimes array, sometimes {items:[...]} or {data:[...]}
                const arr = Array.isArray(data)
                    ? data
                    : Array.isArray(data?.items)
                        ? data.items
                        : Array.isArray(data?.data)
                            ? data.data
                            : [];

                const user = arr[0];
                return user?.id ?? null;
            } catch (e) {
                console.log(
                    `[Warn] qTest user search failed for '${identity}'. ` +
                    `Status: ${e?.response?.status ?? "n/a"}`
                );
                return null;
            }
        }

        const standardHeaders = {
            "Content-Type": "application/json",
            Authorization: `bearer ${constants.QTEST_TOKEN}`,
        };
        const eventType = {
            CREATED: "workitem.created",
            UPDATED: "workitem.updated",
            DELETED: "workitem.deleted",
        };

        let workItemId = undefined;
        let defectToUpdate = undefined;
        switch (event.eventType) {
            case eventType.CREATED: {
                console.log(`[Info] Create workitem event received for 'WI${workItemId}'`);
                console.log(
                    `[Info] New defects are not synched from Azure DevOps. The current workflow expects the defect to be created in qTest first. Exiting.`
                );
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
                console.log(`[Info] Delete workitem event received for 'WI${workItemId}'`);
                console.log(
                    `[Info] Defects are not deleted in qTest automatically when deleting in Azure DevOps. Exiting.`
                );
                return;
            }
            default:
                console.log(`[Error] Unknown workitem event type '${event.eventType}' for 'WI${workitemId}'`);
                return;
        }

        // Prepare data to create/update defect
        const namePrefix = getNamePrefix(workItemId);
        const defectDescription = buildDefectDescription(event);
        const defectSummary = buildDefectSummary(namePrefix, event);
        const fields = getFields(event);

        // Assigned To: ADO -> qTest (P1). Match users by email/UPN; BP SSO commonly uses uniqueName.
        const adoAssignedToRaw = fields[constants.AzDoAssignedToFieldRef || "System.AssignedTo"];
        const adoAssignedToIdentity = extractUpnOrEmailFromAdoAssignedTo(adoAssignedToRaw);

        if (adoAssignedToIdentity) {
            console.log(`[Info] Normalized ADO Assigned To to '${adoAssignedToIdentity}'`);
        } else if (adoAssignedToRaw) {
            console.log(`[Warn] Could not normalize ADO Assigned To '${adoAssignedToRaw}'`);
        } else {
            console.log(`[Info] ADO Assigned To is blank/unassigned.`);
        }

        // Get Severity from ADO and map to qTest allowed values
        const adoSeverity = fields["Microsoft.VSTS.Common.Severity"];
        console.log(`[Info] ADO Severity value: '${adoSeverity}'`);

        const severityMap = {
            "1 - Critical": 10301,
            "2 - High": 10302,
            "3 - Medium": 10303,
            "4 - Low": 10304,
        };

        const qtestSeverityValue = severityMap[adoSeverity] || null;
        console.log(`[Info] Mapped ADO Severity '${adoSeverity}' to qTest Severity`);

        // Get Priority from ADO and map to qTest allowed values
        const adoPriority = fields["Microsoft.VSTS.Common.Priority"];
        console.log(`[Info] ADO Priority value: '${adoPriority}'`);

        const priorityMap = {
            1: 11169, // Very High
            2: 10204, // High
            3: 10203, // Medium
            4: 10202, // Low
        };

        const qtestPriorityValue = priorityMap[adoPriority] || null;
        console.log(`[Info] Mapped ADO Priority '${adoPriority}' to qTest Priority value '${qtestPriorityValue}'`);

        // Get Defect Type from ADO and map to qTest allowed values
        const adoDefectType = fields["BP.ERP.DefectType"];
        console.log(`[Info] ADO Defect Type value: '${adoDefectType}'`);

        const defectTypeMap = {
            "New_Requirement": 956,
            "Code": 957,
            "Data": 958,
            "Environment": 959,
            "Infrastructure": 960,
            "User Authorization": 961,
            "Configuration": 962,
            "User Handling": 963,
            "Translation": 964,
            "Automation": 965,
        };

        const qtestDefectTypeValue = defectTypeMap[adoDefectType] || null;
        console.log(`[Info] Mapped ADO Defect Type '${adoDefectType}' to qTest Defect Type value '${qtestDefectTypeValue}'`);

        // Get state from ADO and map to qTest allowed values
        const adoState = fields["System.State"];
        console.log(`[Info] ADO State value: '${adoState}'`);

        const stateMap = {
        "New": 10001,
        "In Analysis": 10002,
        "Active": 10002,
        "Triage": 10002,               
        "In Resolution": 10004,
        "Awaiting Implementation": 10003,               
        "Resolved": 10953,                  
        "Retest": 10880,
        "Reopened": 10882,
        "Closed": 10881,
        "On Hold": 10883,
        "Rejected": 10853,
        "Cancelled": 10853
        };

        const qtestStatusValue = stateMap[adoState] || null;
        console.log(`[Info] Mapped ADO State '${adoState}' to qTest Status value '${qtestStatusValue}'`);
        if (!qtestStatusValue) {
        console.log(`[Warn] ADO State '${adoState}' does not match any defined qTest status.`);
        }

        //Get Root Cause from ADO (dropdown) to qTest (text)
        const adoRootCause = 
            fields["Microsoft.VSTS.CMMI.RootCause@OData.Community.Display.V1.FormattedValue"] ||
            fields["Microsoft.VSTS.CMMI.RootCause"] || "";
        console.log(`[Info] ADO Root Cause value: '${adoRootCause}'`);

        // Directly sync as plain text (can also clear qTest if ADO field is empty)
        const qtestRootCauseValue = adoRootCause;

        // Get Proposed Fix from ADO (HTML/text) and sync to qTest (LongText)
        const adoProposedFix =
            fields["Microsoft.VSTS.CMMI.ProposedFix"] || "";
        console.log(`[Info] ADO Proposed Fix value length: ${adoProposedFix.length}`);

        // Directly send HTML/text as-is to qTest Proposed Fix field
        const qtestProposedFixValue = adoProposedFix;

        // Get Actual Close Date from ADO and sync to qTest Closed Date
        const adoActualCloseDate = fields["BP.ERP.ActualClose"];
        let qtestClosedDateValue = null;

        // Get External Reference from ADO and sync to qTest
        const adoExternalReference = fields["BP.ERP.ExternalReference"] || "";
        console.log(`[Info] ADO External Reference value: '${adoExternalReference}'`);

        if (adoActualCloseDate) {
            const formattedDate = new Date(adoActualCloseDate);
            const month = String(formattedDate.getUTCMonth() + 1).padStart(2, '0');
            const day = String(formattedDate.getUTCDate()).padStart(2, '0');
            const year = formattedDate.getUTCFullYear();
            qtestClosedDateValue = new Date(adoActualCloseDate).toISOString().replace(".000Z", "+00:00");
            console.log(`[Info] ADO Actual Close Date: '${adoActualCloseDate}' => qTest Closed Date: '${qtestClosedDateValue}'`);
        } else {
            console.log(`[Info] No Actual Close Date found in ADO.`);
        }
        // Get Resolved Reason from ADO and map to qTest allowed values
        const adoResolvedReason =
            fields["Microsoft.VSTS.Common.ResolvedReason"] ||
            fields["Microsoft.VSTS.Common.ResolvedReason@OData.Community.Display.V1.FormattedValue"] ||
            "";
        console.log(`[Info] ADO Resolved Reason: '${adoResolvedReason}'`);

        const resolvedReasonMap = {
            "As Designed": 1299,
            "Cannot Reproduce": 1300,
            "Copied to Backlog": 1301,
            "Deferred": 1302,
            "Duplicate": 1303,
            "Fixed": 1304,
            "Fixed and verified": 1305,
            "Obsolete": 1306,
            "Will not Fix": 1307,
        };

        const qtestResolvedReasonValue = resolvedReasonMap[adoResolvedReason] || null;
        console.log(
            `[Info] Mapped ADO Resolved Reason '${adoResolvedReason}' → qTest value '${qtestResolvedReasonValue}'`
        );

        if (defectToUpdate) {
            await updateDefect(
                defectToUpdate,
                defectSummary,
                defectDescription,
                qtestSeverityValue,
                qtestPriorityValue,
                qtestRootCauseValue,
                qtestDefectTypeValue,
                qtestStatusValue,
                qtestProposedFixValue,
                qtestClosedDateValue,
                qtestResolvedReasonValue,
                adoAssignedToIdentity
            );
        }

        
        function extractAdoAssignedToIdentity(adoAssignedTo) {
            if (!adoAssignedTo) return null;
            if (typeof adoAssignedTo === "string") return adoAssignedTo.trim() || null;
            // ADO often provides an identity object
            const candidate =
                adoAssignedTo.uniqueName ||
                adoAssignedTo.mail ||
                adoAssignedTo.email ||
                adoAssignedTo.userPrincipalName ||
                adoAssignedTo.displayName ||
                null;
            return (candidate && typeof candidate === "string" && candidate.trim().length > 0) ? candidate.trim() : null;
        }

        async function resolveQtestUserIdByUsernameOrEmail(identity) {
            if (!identity) return null;
            try {
                // qTest user search endpoint (queries by username); BP SSO users commonly store UPN/email in username.
                const url = `https://${constants.ManagerURL}/api/v3/users/search?username=${encodeURIComponent(identity)}`;
                const result = await get(url);

                // Result shape can vary by deployment; handle common patterns
                if (Array.isArray(result) && result.length > 0 && result[0]?.id) return result[0].id;
                if (result?.items && Array.isArray(result.items) && result.items.length > 0 && result.items[0]?.id) return result.items[0].id;
                if (result?.id) return result.id;

                return null;
            } catch (err) {
                console.log(`[Warn] Failed to resolve qTest user for '${identity}'. Leaving assignment blank. ${err.message}`);
                return null;
            }
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

        async function getDefectByWorkItemId(workItemId) {
            const prefix = getNamePrefix(workItemId);
            const url = "https://" + constants.ManagerURL + "/api/v3/projects/" + constants.ProjectID + "/search";
            const requestBody = {
                object_type: "defects",
                fields: ["*"],
                query: "Summary ~ '" + prefix + "'",
            };

            console.log(`[Info] Get existing defect for 'WI${workItemId}'`);
            let failed = false;
            let defect = undefined;

            try {
                const response = await post(url, requestBody);
                console.log(response);

                if (!response || response.total === 0) {
                    console.log("[Info] Defect not found by work item id.");
                } else {
                    if (response.total === 1) {
                        defect = response.items[0];
                    } else {
                        failed = true;
                        console.log("[Warn] Multiple Defects found by work item id.");
                    }
                }
            } catch (error) {
                console.log("[Error] Failed to get defect by work item id.", error);
                failed = true;
            }

            return { failed: failed, defect: defect };
        }

        async function updateDefect(
            defectToUpdate,
            summary,
            description,
            severityValue,
            priorityValue,
            rootCauseValue,
            defectTypeValue,
            statusValue,
            proposedFixValue,
            closedDateValue,
            resolvedReasonValue,
            assignedToUserId
        ) 
            {
            const defectId = defectToUpdate.id;
            const defectPid = defectToUpdate.pid; // PID

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

            // Add Severity to qTest update payload if mapped
            if (severityValue) {
                requestBody.properties.push({
                    field_id: constants.DefectSeverityFieldID,
                    field_value: parseInt(severityValue),
                });
            }

            // Add Priority to qTest update payload if mapped
            if (priorityValue) {
                requestBody.properties.push({
                    field_id: constants.DefectPriorityFieldID,
                    field_value: parseInt(priorityValue),
                });
            }

            // Add Root Cause (text field in qTest)
            if (constants.DefectRootCauseFieldID) {
                requestBody.properties.push({
                    field_id: constants.DefectRootCauseFieldID,
                    field_value: rootCauseValue || "", // clears if ADO Root Cause is empty
                });
            }

            // Add Defect Type to qTest update payload if mapped
            if (defectTypeValue) {
                requestBody.properties.push({
                    field_id: constants.DefectTypeFieldID,
                    field_value: parseInt(defectTypeValue),
                });
                console.log(`[Info] Added Defect Type '${defectTypeValue}' to qTest update payload.`);
            } else {
                console.log(`[Warn] No Defect Type mapping found or field is empty in ADO.`);
            }

            // Add Status to qTest update payload if mapped
            if (statusValue) {
                requestBody.properties.push({
                    field_id: constants.DefectStatusFieldID,
                    field_value: parseInt(statusValue),
                });
                console.log(`[Info] Added Status '${statusValue}' to qTest update payload.`);
            } else {
                console.log(`[Warn] No Status mapping found or ADO state '${adoState}' not mapped.`);
            }

            // Add Proposed Fix (LongText field in qTest)
            if (constants.DefectProposedFixFieldID) {
                const formattedProposedFix = proposedFixValue
                    ? `<p>${proposedFixValue}</p>` // wrap with <p> for LongText field
                    : "";
                requestBody.properties.push({
                    field_id: constants.DefectProposedFixFieldID,
                    //field_value: proposedFixValue || "",
                    field_value: formattedProposedFix,
                });
                console.log(`[Info] Added Proposed Fix to qTest update payload.`);
            }

            // Add External Reference (text field in qTest)
            if (constants.DefectExternalReferenceFieldID) {
                requestBody.properties.push({
                    field_id: constants.DefectExternalReferenceFieldID,
                    field_value: adoExternalReference || "",
                });
                console.log(`[Info] Added External Reference to qTest update payload.`);
            }

            // Add Closed Date (Date field in qTest)
            if (constants.DefectClosedDateFieldID && closedDateValue) {
                requestBody.properties.push({
                    field_id: constants.DefectClosedDateFieldID,
                    field_value: closedDateValue,
                });
                console.log(`[Info] Added Closed Date '${closedDateValue}' to qTest update payload.`);
            }

            // Add Resolved Reason to qTest update payload
            if (constants.DefectResolvedReasonFieldID && resolvedReasonValue) {
                requestBody.properties.push({
                    field_id: constants.DefectResolvedReasonFieldID,
                    field_value: parseInt(resolvedReasonValue),
                });
                console.log(`[Info] Added Resolved Reason '${resolvedReasonValue}' to qTest update payload.`);
            } else {
                console.log(`[Warn] No Resolved Reason provided or mapping not found`);
            }

            console.log(`[Info] Updating defect '${defectId}' (${defectPid}).`);
            console.log('[Debug] Final qTest Update Payload:', JSON.stringify(requestBody, null, 2));


            try {
                const response = await put(url, requestBody);
                console.log(`[Info] Defect '${defectId}' (${defectPid}) updated.`);
            } catch (error) {
                if (error.response) {
                    console.log(`[Error] Failed to update defect '${defectId}'.`, error);
                    //console.log(`[Debug] qTest API Response: ${JSON.stringify(response, null, 2)}`);
                } else {
                    console.log(`[Error] Failed to update defect '${defectId}' — ${error.message}`);
                    //console.log(`[Error] Response Data: ${JSON.stringify(error.response.data, null, 2)}`);
                    //console.log(`[Error] HTTP ${error.response.status}: ${JSON.stringify(error.response.data, null, 2)}`);
                }
            }
        }

        function post(url, requestBody) {
            return doqTestRequest(url, "POST", requestBody);
        }

        function put(url, requestBody) {
            return doqTestRequest(url, "PUT", requestBody);
        }

        
        function get(url) {
            return doqTestRequest(url, "GET");
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
                console.log(`[Error] URL: ${url}`);
                console.log(`[Error] HTTP Status: ${status}`);
                console.log(`[Error] Message: ${message}`);
                throw new Error(`qTest API ${method} ${url} failed with ${status || "Unknown"}: ${message}`);
                //throw error;
            }
        }
    };
 

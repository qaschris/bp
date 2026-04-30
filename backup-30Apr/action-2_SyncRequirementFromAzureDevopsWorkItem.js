const axios = require("axios");

exports.handler = async function ({ event, constants, triggers }, context, callback) {
    function buildRequirementDescription(eventData) {
        const fields = getFields(eventData);
        return `<a href="${eventData.resource._links.html.href}" target="_blank">Open in Azure DevOps</a><br>
<b>Type:</b> ${fields["System.WorkItemType"]}<br>
<b>Area:</b> ${fields["System.AreaPath"]}<br>
<b>Iteration:</b> ${fields["System.IterationPath"]}<br>
<b>State:</b> ${fields["System.State"]}<br>
<b>Reason:</b> ${fields["System.Reason"]}<br>
<b>Complexity:</b> ${fields["Custom.Complexity"] || ""}<br>
<b>Acceptance Criteria:</b> ${fields["Microsoft.VSTS.Common.AcceptanceCriteria"]}<br>
<b>Description:</b> ${fields["System.Description"] || ""}`;
    }

    function buildRequirementName(namePrefix, eventData) {
        const fields = getFields(eventData);
        return `${namePrefix}${fields["System.Title"]}`;
    }

    function getFields(eventData) {
        // In case of update the fields can be taken from the revision, in case of create from the resource directly
        return eventData.resource.revision ? eventData.resource.revision.fields : eventData.resource.fields;
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
    let requirementToUpdate = undefined;
    switch (event.eventType) {
        case eventType.CREATED: {
            workItemId = event.resource.id;
            console.log(`[Info] Create workitem event received for 'WI${workItemId}'`);
            break;
        }
        case eventType.UPDATED: {
            workItemId = event.resource.workItemId;
            console.log(`[Info] Update workitem event received for 'WI${workItemId}'`);
            const getReqResult = await getRequirementByWorkItemId(workItemId);
            if (getReqResult.failed) {
                return;
            }
            if (getReqResult.requirement === undefined && !constants.AllowCreationOnUpdate) {
                console.log("[Info] Creation of Requirement on update event not enabled. Exiting.");
                return;
            }
            requirementToUpdate = getReqResult.requirement;
            break;
        }
        case eventType.DELETED: {
            workItemId = event.resource.id;
            console.log(`[Info] Delete workitem event received for 'WI${workItemId}'`);
            const getReqResult = await getRequirementByWorkItemId(workItemId);
            if (getReqResult.failed) {
                return;
            }
            if (getReqResult.requirement === undefined) {
                console.log(`[Info] Requirement not found to delete. Exiting.`);
                return;
            }
            // Delete requirement and finish
            await deleteRequirement(getReqResult.requirement);
            return;
        }
        default:
            console.log(`[Error] Unknown workitem event type '${event.eventType}' for 'WI${workitemId}'`);
            return;
    }

    // Prepare data to create/update requirement
    const namePrefix = getNamePrefix(workItemId);
    const requirementDescription = buildRequirementDescription(event);
    const requirementName = buildRequirementName(namePrefix, event);

    // Get Area Path from ADO and map to qTest Stream/Squad field
    const fields = getFields(event);
    const adoAreaPath = fields["System.AreaPath"];
    console.log(`[Info] Mapped ADO Area Path '${adoAreaPath}' to qTest Stream/Squad field.`);

    // Get Complexity from ADO and map to qTest allowed value IDs
    const adoComplexity = fields["Custom.Complexity"];
    const complexityMap = {
        "1 - Very High": 1541,
        "2 - High": 1542,
        "3 - Medium": 1543,
        "4 - Low": 1544,
        "5 - Very Low": 1545,
    };
    const qtestComplexityValue = complexityMap[adoComplexity] || null;
    console.log(`[Info] Mapped ADO Complexity '${adoComplexity}' to qTest Complexity.`);

    if (requirementToUpdate) {
        await updateRequirement(requirementToUpdate, requirementName, requirementDescription, adoAreaPath, qtestComplexityValue);
    } else {
        createRequirement(requirementName, requirementDescription, adoAreaPath, qtestComplexityValue);
    }

    function getNamePrefix(workItemId) {
        return `WI${workItemId}: `;
    }

    async function getRequirementByWorkItemId(workItemId) {
        const prefix = getNamePrefix(workItemId);
        const url = "https://" + constants.ManagerURL + "/api/v3/projects/" + constants.ProjectID + "/search";
        const requestBody = {
            object_type: "requirements",
            fields: ["*"],
            query: "Name ~ '" + prefix + "'",
        };

        console.log(`[Info] Get existing requirement for 'WI${workItemId}'`);
        let failed = false;
        let requirement = undefined;

        try {
            const response = await post(url, requestBody);
            console.log(response);

            if (!response || response.total === 0) {
                console.log("[Info] Requirement not found by work item id.");
            } else {
                if (response.total === 1) {
                    requirement = response.items[0];
                } else {
                    failed = true;
                    console.log("[Warn] Multiple Requirements found by work item id.");
                }
            }
        } catch (error) {
            console.log("[Error] Failed to get requirement by work item id.", error);
            failed = true;
        }

        return { failed: failed, requirement: requirement };
    }

    async function updateRequirement(requirementToUpdate, name, description, areaPath, complexityValue) {
        const requirementId = requirementToUpdate.id;
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/requirements/${requirementId}`;
        const requestBody = {
            name: name,
            properties: [
                {
                    field_id: constants.RequirementDescriptionFieldID,
                    field_value: description,
                },
                {
                    field_id: constants.RequirementStreamSquadFieldID, // Stream/Squad constant
                    field_value: areaPath,
                },
            ],
        };

        if (complexityValue) {
            requestBody.properties.push({
                field_id: constants.RequirementComplexityFieldID, // Complexity constant
                field_value: complexityValue,
            });
        }

        console.log(`[Info] Updating requirement '${requirementId}'.`);

        try {
            await put(url, requestBody);
            console.log(`[Info] Requirement '${requirementId}' updated.`);
        } catch (error) {
            console.log(`[Error] Failed to update requirement '${requirementId}'.`, error);
        }
    }

    async function createRequirement(name, description, areaPath, complexityValue) {
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/requirements`;
        const requestBody = {
            name: name,
            parent_id: constants.RequirementParentID,
            properties: [
                {
                    field_id: constants.RequirementDescriptionFieldID,
                    field_value: description,
                },
                {
                    field_id: constants.RequirementStreamSquadFieldID, // Stream/Squad constant
                    field_value: areaPath,
                },
            ],
        };

        if (complexityValue) {
            requestBody.properties.push({
                field_id: constants.RequirementComplexityFieldID, // complexity constant
                field_value: complexityValue,
            });
        }

        console.log(`[Info] Creating requirement.`);

        try {
            await post(url, requestBody);
            console.log(`[Info] Requirement created.`);
        } catch (error) {
            console.log(`[Error] Failed to create requirement`, error);
        }
    }

    async function deleteRequirement(requirementToDelete) {
        const requirementId = requirementToDelete.id;
        const url = `https://${constants.ManagerURL}/api/v3/projects/${constants.ProjectID}/requirements/${requirementId}`;

        console.log(`[Info] Deleting requirement '${requirementId}'.`);

        try {
            await doRequest(url, "DELETE", null);
            console.log(`[Info] Requirement '${requirementId}' deleted.`);
        } catch (error) {
            console.log(`[Error] Failed to delete requirement '${requirementId}'.`, error);
        }
    }

    function post(url, requestBody) {
        return doRequest(url, "POST", requestBody);
    }

    function put(url, requestBody) {
        return doRequest(url, "PUT", requestBody);
    }

    async function doRequest(url, method, requestBody) {
        const opts = {
            url: url,
            method: method,
            headers: standardHeaders,
            data: requestBody,
        };

        try {
            const response = await axios(opts);
            return response.data;
        } catch (error) {
            throw new Error(`Failed to ${method} ${url}. ${error.message}`);
        }
    }
};

# ADO to qTest Focused Customizations

## Purpose

This document is the customer-facing technical delivery reference for the six live BP Azure DevOps and qTest integration items in this repository.

It is organized around the delivered items the customer asked to see, and each item includes the same decision-useful sections:

- Pre-requisite
- URLs
- Sample Code
- Impacts
- How to Use
- Outcomes

This document describes the live Pulse rule behavior as it exists today. It also calls out known operational limitations where they affect support, audit review, or end-user expectations.

## Solution Context

- Integration platform: qTest Pulse
- Runtime: Node.js-based Pulse webhook rules
- Systems in scope: Azure DevOps, qTest Manager, qTest Pulse
- Delivery shape: six live integration items, plus separate migration utilities for older qTest requirements
- Main base URLs:
  - qTest Manager base: `ManagerURL`
  - Azure DevOps project base: `AzDoProjectURL`
- Shared credentials:
  - `QTEST_TOKEN`
  - `AZDO_TOKEN`
- Shared support trigger:
  - `ChatOpsEvent`

## Delivered Items Summary

1. `CreateDefectInAzureDevops.P1.js`
2. `SyncDefectFromAzureDevopsWorkItem.P1.js`
3. `UpdateDefectInAzureDevops.P1.js`
4. `SyncRequirementFromAzureDevopsWorkItem.P1.js`
5. `UpdateRequirementStatusInAzureDevops.P1.js`
6. `SyncRICEFWFeatureFromAzureDevops.js`

## Shared Pre-requisites

- qTest Pulse access with the BP project rules deployed and enabled
- qTest Manager project access and Azure DevOps project access
- qTest custom fields created and their qTest field ids configured in Pulse constants
- Azure DevOps custom fields created and their field references configured in Pulse constants
- ADO service hooks or event sources wired so Pulse receives work item create, update, and delete events where required
- qTest event rules wired so Pulse receives defect-update and requirement-update events where required
- `SyncUserRegex` configured where loop prevention depends on identifying the integration service user
- The `WI<id>:` naming convention preserved on synced records so the linked Azure DevOps work item can be recovered later
- qTest comments API and Azure DevOps comments API available for the rules that synchronize comments

## Shared Integration Behaviors

- qTest constrained dropdown values are resolved dynamically from qTest field metadata instead of relying on hardcoded numeric option ids.
- Friendly warnings and failures are sent through `ChatOpsEvent` so support receives business-readable messages instead of only console output.
- Defect comments are synchronized in both directions.
- Standard Requirement and RICEFW comments currently flow one way from ADO to qTest.
- Defect field synchronization is asynchronous across qTest, Pulse, and Azure DevOps, so short delays are expected after create and update activity.
- Origin markers such as `[From ADO]`, `[From qTest]`, and CID markers are used where available to reduce comment duplication and echo loops.

## Delivered Item 1

### 1. CreateDefectInAzureDevops.P1.js

**Pre-requisite**

- In addition to the shared pre-requisites, this rule needs the qTest defect create event for the target BP project.
- Required qTest constants:
  - `DefectSummaryFieldID`
  - `DefectDescriptionFieldID`
  - `DefectSeverityFieldID`
  - `DefectPriorityFieldID`
  - `DefectTypeFieldID`
  - `DefectStatusFieldID`
  - `DefectAffectedReleaseFieldID`
  - `DefectCreatedByFieldID`
  - `DefectExternalReferenceFieldID`
  - `DefectRootCauseFieldID`
  - `DefectAssignedToFieldID`
  - `DefectAssignedToTeamFieldID`
  - `DefectTargetDateFieldID`
- Required ADO field references:
  - `title`
  - `reproSteps`
  - `tags`
  - `state`
  - `severity`
  - `priority`
  - `areaPath`
  - `assignedTo`
  - `defectType`
  - `bugStage`
  - `createdBy`
  - `externalReference`
  - `rootCause`
  - `proposedFix`
  - `targetDate`
- The Pulse trigger `qTestDefectSubmitted` must exist because the rule re-emits the event when it retries an incomplete qTest save.

**URLs**

- qTest defect field metadata:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/settings/defects/fields`
- qTest defect detail read:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/defects/{defectId}`
- qTest user lookup:
  - `{ManagerURL}/api/v3/users/{userId}`
- Azure DevOps classification lookup:
  - `{AzDoProjectURL}/_apis/wit/classificationnodes/{areas|iterations}?$depth=10&api-version=6.0`
- Azure DevOps bug create:
  - `{AzDoProjectURL}/_apis/wit/workitems/$Bug?api-version=6.0`
- qTest defect summary update after successful create:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/defects/{defectId}`

**Sample Code**

```
const url = `${baseUrl}/_apis/wit/workitems/$Bug?api-version=6.0`;
const requestBody = [
    { op: "add", path: "/fields/System.Title", value: adoTitle },
    { op: "add", path: "/fields/System.State", value: mappedStatus },
    { op: "add", path: "/relations/-", value: { rel: "Hyperlink", url: qTestLink } }
];

const response = await axios.post(url, requestBody, {
    headers: {
        "Content-Type": "application/json-patch+json",
        "Authorization": `basic ${Buffer.from(`:${constants.AZDO_TOKEN}`).toString("base64")}`
    }
});
```

**Impacts**

- Creates the linked Azure DevOps `Bug` from a qTest defect.
- Adds a backlink from the ADO bug to the originating qTest defect.
- Updates the qTest defect name so it contains the `WI<id>:` prefix used by later sync rules.
- Uses configured fallbacks when Area Path, Iteration Path, or assignee values cannot be resolved cleanly.
- Intentionally does not send `Closed Date` or `Resolved Reason` during initial bug create. Those values are handled by later update flows.

**How to Use**

- End user creates the defect in qTest and completes the save.
- If qTest initially returns an incomplete record, the rule retries until the record has enough detail to create the ADO bug.
- Support reviews any `ChatOpsEvent` warning when the rule had to default Area Path, Iteration Path, or assignee values.

**Outcomes**

- A linked ADO bug is created.
- The qTest defect is updated with the ADO work item id for downstream synchronization.
- The defect is ready for later two-way defect field sync and two-way comment sync.

## Delivered Item 2

### 2. SyncDefectFromAzureDevopsWorkItem.P1.js

**Pre-requisite**

- In addition to the shared pre-requisites, the qTest defect must already be linked through the `WI<id>:` naming convention.
- Azure DevOps work item events must be supplied to Pulse for:
  - `workitem.updated`
  - `workitem.created`
  - `workitem.deleted`
- Required qTest constants:
  - `DefectSummaryFieldID`
  - `DefectDescriptionFieldID`
  - `DefectSeverityFieldID`
  - `DefectPriorityFieldID`
  - `DefectTypeFieldID`
  - `DefectStatusFieldID`
  - `DefectRootCauseFieldID`
  - `DefectExternalReferenceFieldID`
  - `DefectResolvedReasonFieldID`
  - `DefectAssignedToFieldID`
  - `DefectAssignedToTeamFieldID`
  - `DefectTargetDateFieldID`
- Required ADO field references:
  - `title`
  - `reproSteps`
  - `state`
  - `severity`
  - `priority`
  - `defectType`
  - `rootCause`
  - `proposedFix`
  - `targetDate`
  - `externalReference`
  - `resolvedReason`
  - `areaPath`
  - `assignedTo`
- qTest project-user lookup must be available because inbound ADO assignee resolution depends on active qTest project users.

**URLs**

- qTest field metadata:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/settings/defects/fields`
- qTest project users:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/users?inactive=false`
- qTest defect search by `WI<id>`:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/search`
- qTest defect update:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/defects/{defectId}`
- qTest defect comments:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/defects/{defectId}/comments`
- Azure DevOps comments:
  - `{AzDoProjectURL}/_apis/wit/workitems/{workItemId}/comments?api-version=7.0-preview.3`

**Sample Code**

```
const requestBody = {
    properties: [
        { field_id: constants.DefectSummaryFieldID, field_value: summary },
        { field_id: constants.DefectDescriptionFieldID, field_value: description },
        { field_id: constants.DefectStatusFieldID, field_value: parseInt(statusValue, 10) }
    ]
};

await put(url, requestBody);
```

**Impacts**

- Updates the linked qTest defect from Azure DevOps field values.
- Adds qTest comments from ADO comment activity.
- Uses qTest dropdown resolution for inbound ADO values instead of hardcoded mappings where labels align.
- Falls back to configured defaults for unresolved inbound assignee or team values instead of failing the whole sync.
- Create and delete events are intentionally informational only for defects; this rule does not create new qTest defects from ADO and does not delete qTest defects from ADO delete events.
- Current timing limitation: this inbound rule still writes mapped ADO values back to qTest on qualifying updates, so rapid successive qTest edits can still be overwritten later by stale ADO-backed callbacks.

**How to Use**

- Azure DevOps user updates the linked bug.
- Pulse receives the work item event and resolves the matching qTest defect through the `WI<id>:` prefix.
- Support should expect comments and field updates to arrive asynchronously.
- Current best practice is to let one defect save and sync settle before making another broad round of defect field changes in qTest.

**Outcomes**

- qTest defect fields reflect the current ADO bug values for the mapped fields.
- ADO comments appear as qTest comments with `[From ADO]` markers.
- Support receives friendly warnings when an inbound value could not be mapped cleanly.

## Delivered Item 3

### 3. UpdateDefectInAzureDevops.P1.js

**Pre-requisite**

- In addition to the shared pre-requisites, the qTest defect must already contain the linked `WI<id>` reference.
- qTest defect update events must be enabled in Pulse for the target project.
- Required qTest constants:
  - `DefectSummaryFieldID`
  - `DefectDescriptionFieldID`
  - `DefectSeverityFieldID`
  - `DefectPriorityFieldID`
  - `DefectTypeFieldID`
  - `DefectStatusFieldID`
  - `DefectAffectedReleaseFieldID`
  - `DefectExternalReferenceFieldID`
  - `DefectRootCauseFieldID`
  - `DefectAssignedToFieldID`
  - `DefectAssignedToTeamFieldID`
  - `DefectTargetDateFieldID`
- Required ADO field references:
  - `title`
  - `reproSteps`
  - `state`
  - `severity`
  - `priority`
  - `defectType`
  - `externalReference`
  - `bugStage`
  - `rootCause`
  - `proposedFix`
  - `resolvedReason`
  - `areaPath`
  - `assignedTo`
  - `targetDate`
- `SyncUserRegex` is strongly recommended because this rule already uses it to skip qTest updates made by the integration user and prevent echo loops.

**URLs**

- qTest defect read:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/defects/{defectId}`
- qTest user lookup:
  - `{ManagerURL}/api/v3/users/{userId}`
- qTest comments:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/defects/{defectId}/comments`
- Azure DevOps work item read:
  - `{AzDoProjectURL}/_apis/wit/workitems/{workItemId}?api-version=6.0&$expand=Relations`
- Azure DevOps comments:
  - `{AzDoProjectURL}/_apis/wit/workItems/{workItemId}/comments?api-version=6.0-preview.3`
- Azure DevOps work item patch:
  - `{AzDoProjectURL}/_apis/wit/workitems/{workItemId}?api-version=6.0`

**Sample Code**

```
const patchData = [];

if (adoTitle && curTitle !== adoTitle) {
    patchData.push(buildFieldPatchOperation(adoFieldRefs.title, adoTitle));
}

if (patchData.length === 0) {
    console.log("[Info] No ADO changes detected; skipping patch (prevents loops).");
    return;
}

await axios.patch(adoPatchUrl, patchData, {
    headers: {
        Authorization: `Basic ${encodedToken}`,
        "Content-Type": "application/json-patch+json"
    }
});
```

**Impacts**

- Updates only the changed ADO fields instead of rewriting the full work item.
- Synchronizes qTest defect comments to ADO comments even when no field patch is needed.
- Skips qTest updates made by the integration user when `SyncUserRegex` matches.
- Suppresses no-op ADO field patches, which reduces unnecessary audit noise on the ADO side.
- Sanitizes qTest descriptions before writing them back into ADO so embedded link text does not keep growing.

**How to Use**

- End user edits a linked defect in qTest.
- The rule reads the current qTest defect, extracts the linked ADO work item id, reads the current ADO work item, and patches only the changed fields.
- If qTest comments were added, they are pushed to the ADO comments API even when no other field changed.
- Support should review warnings when assignee, Area Path, or Iteration Path values are defaulted.

**Outcomes**

- Linked ADO bug fields stay aligned to the qTest defect for the mapped outbound fields.
- qTest-origin comments appear in ADO with `[From qTest]` markers.
- Outbound defect sync already has better loop protection than the inbound defect field path because it skips sync-user changes and skips no-op patches.

## Delivered Item 4

### 4. SyncRequirementFromAzureDevopsWorkItem.P1.js

**Pre-requisite**

- In addition to the shared pre-requisites, Azure DevOps standard Requirement work item events must be supplied to Pulse.
- Required qTest constants:
  - `RequirementDescriptionFieldID`
  - `RequirementStreamSquadFieldID`
  - `RequirementComplexityFieldID`
  - `RequirementWorkItemTypeFieldID`
  - `RequirementPriorityFieldID`
  - `RequirementTypeFieldID`
  - `RequirementAssignedToFieldID`
  - `RequirementIterationPathFieldID`
- Optional but supported qTest constants:
  - `RequirementApplicationNameFieldID`
  - `RequirementFitGapFieldID`
  - `RequirementBPEntityFieldID`
  - `RequirementParentID`
- Required ADO field references:
  - `title`
  - `workItemType`
  - `areaPath`
  - `iterationPath`
  - `assignedTo`
  - `description`
  - `acceptanceCriteria`
  - `priority`
  - `complexity`
  - `requirementCategory`
  - optional when those source fields are used:
    - `applicationName`
    - `fitGap`
    - `entity`

**URLs**

- qTest requirement field metadata:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/settings/requirements/fields`
- qTest module tree read:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/modules/{parentId}?expand=descendants`
- qTest module create:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/modules`
- qTest requirement search:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/search`
- qTest requirement create:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/requirements`
- qTest requirement update:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/requirements/{requirementId}`
- qTest requirement comments:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/requirements/{requirementId}/comments`
- Azure DevOps work item comments:
  - `{AzDoProjectURL}/_apis/wit/workItems/{workItemId}/comments?api-version=7.0-preview.3`

**Sample Code**

```
const evaluation = evaluateRequirementUpdate(requirementDetails, desiredState);

if (!evaluation.needsUpdate) {
    console.log(`[Info] Requirement '${requirementToUpdate.id}' is already in sync. Skipping update.`);
    return requirementDetails;
}

const updated = await put(url, evaluation.requestBody);
```

**Impacts**

- Creates, updates, or deletes the matching qTest Requirement for standard ADO Requirement items.
- Builds the qTest module path from release plus ADO Area Path.
- Creates missing qTest modules when required.
- Uses desired-state comparison so no-op ADO events do not rewrite qTest unnecessarily.
- Supports comment-only ADO updates and keeps requirement comments flowing from ADO to qTest.
- Optional fields such as Application Name, Fit Gap, and BP Entity are left unchanged with warnings if their qTest field ids are not configured or their values cannot be resolved.

**How to Use**

- Azure DevOps Requirement is created, updated, or deleted.
- Pulse uses the work item event payload to build the desired qTest state, find or create the correct module path, and compare the desired state with the current qTest requirement.
- If nothing changed, the rule exits cleanly without rewriting qTest.
- If the item is new, the rule creates it under `RequirementParentID`.

**Outcomes**

- qTest Requirement records are created and maintained from Azure DevOps as the source of truth.
- The qTest requirement description contains the ADO link plus a structured snapshot of the ADO content.
- Requirement comments appear in qTest with `[From ADO]` and `[CID:<comment id>]` markers.

## Delivered Item 5

### 5. UpdateRequirementStatusInAzureDevops.P1.js

**Pre-requisite**

- In addition to the shared pre-requisites, this item needs qTest requirement update events for the target project.
- Required constants:
  - `ProjectID`
  - `RequirementStatusFieldID`
  - `AzDoTestingStatusFieldRef`
  - `AZDO_TOKEN`
  - `ManagerURL`
  - `AzDoProjectURL`
  - `QTEST_TOKEN`
- `SyncUserRegex` is recommended because the rule uses it to skip qTest updates made by the integration user.
- The qTest requirement must already contain the linked `WI<id>` reference so the ADO work item can be identified.

**URLs**

- qTest requirement read:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/requirements/{requirementId}`
- Azure DevOps work item read:
  - `{AzDoProjectURL}/_apis/wit/workitems/{workItemId}?api-version=6.0`
- Azure DevOps work item patch:
  - `{AzDoProjectURL}/_apis/wit/workitems/{workItemId}?api-version=6.0`

**Sample Code**

```
const requestBody = [
    {
        op: "add",
        path: `/fields/${adoTestingStatusFieldRef}`,
        value: adoStatus
    }
];

await axios.patch(adoUrl, requestBody, {
    headers: {
        "Content-Type": "application/json-patch+json",
        Authorization: `basic ${Buffer.from(`:${adoToken}`).toString("base64")}`
    }
});
```

**Impacts**

- Updates only one ADO field: the configured Testing Status field.
- Does not attempt full Requirement field synchronization from qTest back into ADO.
- Skips no-op patches when the ADO testing status already matches.
- Skips events when the changed qTest field set does not include the configured status field id.

**How to Use**

- qTest user changes the configured requirement status field.
- Pulse reads the linked qTest Requirement, derives the linked ADO work item id from the `WI<id>` name, reads the current ADO status, and patches ADO only when the value changed.
- Support should use this rule only for the agreed BP testing status process, not as a general-purpose Requirement update bridge.

**Outcomes**

- The ADO testing status field is kept aligned with the qTest requirement status process.
- The rule avoids broad side effects because it is intentionally narrow.
- Audit noise is limited because unchanged statuses are not repatched.

## Delivered Item 6

### 6. SyncRICEFWFeatureFromAzureDevops.js

**Pre-requisite**

- In addition to the shared pre-requisites, Azure DevOps Feature work item events must be supplied to Pulse.
- Required qTest constants:
  - `RequirementDescriptionFieldID`
  - `RequirementStreamSquadFieldID`
  - `RequirementWorkItemTypeFieldID`
  - `RequirementAssignedToFieldID`
  - `RequirementIterationPathFieldID`
  - `RequirementRICEFWConfigurationFieldID`
  - `FeatureParentID`
- Additional supported qTest constants used when configured:
  - `RequirementStateFieldID`
  - `RequirementReasonFieldID`
  - `RequirementAcceptanceCriteriaFieldID`
  - `RequirementComplexityFieldID`
  - `RequirementPriorityFieldID`
  - `RequirementTypeFieldID`
  - `RequirementTestingStatusFieldID`
- Required ADO field references:
  - `title`
  - `workItemType`
  - `areaPath`
  - `iterationPath`
  - `state`
  - `reason`
  - `assignedTo`
  - `description`
  - `acceptanceCriteria`
  - `priority`
  - `complexity`
  - `ricefwId`
  - `featureType`
  - `ricefwConfiguration`
- Qualification rule:
  - Work Item Type must be `Feature`
  - Feature Type must be `RICEFW` or `Change Request`
  - RICEFW Configuration must be one of `Enhancement`, `Form`, `Interface`, `Report`, or `Workflow`
  - State must not be `Rejected` or `Cancelled`

**URLs**

- qTest requirement field metadata:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/settings/requirements/fields`
- qTest module tree read:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/modules/{parentId}?expand=descendants`
- qTest module create:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/modules`
- qTest requirement search:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/search`
- qTest requirement create:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/requirements`
- qTest requirement update:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/requirements/{requirementId}`
- qTest requirement delete:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/requirements/{requirementId}`
- qTest requirement comments:
  - `{ManagerURL}/api/v3/projects/{ProjectID}/requirements/{requirementId}/comments`
- Azure DevOps work item comments:
  - `{AzDoProjectURL}/_apis/wit/workItems/{workItemId}/comments?api-version=7.0-preview.3`

**Sample Code**

```
const isRicefwFeature =
    workItemType.toLowerCase() === "feature" &&
    (featureType.toLowerCase() === "ricefw" || featureType.toLowerCase() === "change request") &&
    (
        ricefwConfiguration.toLowerCase() === "enhancement" ||
        ricefwConfiguration.toLowerCase() === "form" ||
        ricefwConfiguration.toLowerCase() === "interface" ||
        ricefwConfiguration.toLowerCase() === "report" ||
        ricefwConfiguration.toLowerCase() === "workflow"
    ) &&
    (featureState.toLowerCase() !== "rejected" && featureState.toLowerCase() !== "cancelled");
```

**Impacts**

- Creates, updates, or deletes qTest requirements for qualifying RICEFW or Change Request features in ADO.
- Builds module placement under `FeatureParentID` from release plus ADO Area Path.
- Synchronizes one-way comments from ADO to qTest.
- Uses no-op evaluation so unchanged ADO events do not keep rewriting qTest.
- If an already-synced item later falls out of scope, the rule leaves the existing qTest item unchanged and emits a warning instead of silently deleting or repurposing it.

**How to Use**

- ADO Feature is created, updated, or deleted.
- Pulse first checks whether the work item still qualifies as a RICEFW item under the configured business rules.
- If it qualifies, the rule creates or updates the qTest requirement under the RICEFW root.
- If it no longer qualifies and a qTest item already exists, support receives a warning and the qTest item is left unchanged pending business decision.

**Outcomes**

- qTest receives dedicated RICEFW requirements only for qualified feature items.
- The description block contains the ADO link plus RICEFW-specific details such as RICEFW ID, State, Reason, Acceptance Criteria, and Description.
- Comment synchronization for RICEFW follows the same `[From ADO]` and CID approach as the standard Requirement rule.

## Known Operational Notes

- Defect create and update processing is asynchronous across qTest, Pulse, and ADO. A short delay between user action and downstream visibility is normal.
- The current stale-update risk is concentrated in the inbound defect field-sync rule because it does not yet apply the same no-op evaluation pattern already used by the standard Requirement and RICEFW rules.
- The outbound defect update rule already provides stronger loop protection because it skips qTest changes made by the sync user and suppresses no-op ADO patches.
- Requirement and RICEFW flows are currently more mature on no-op detection than the defect inbound field-sync path.

## Appendix A

### Internal Utilities Not Counted in the Six Delivered Items

- `QueueRequirementMigration.P1.js`
- `ProcessRequirementMigrationBatch.P1.js`

These two Pulse rules support one-time migration of older qTest Requirements into the newer release-and-area-based module structure. They are operational support utilities rather than part of the six live customer-facing integration items listed above.

## Appendix B

### Expected Customer-Facing Outcomes Across the Full Delivery

- Defects start in qTest and are represented in Azure DevOps as linked Bugs.
- Linked defect comments move in both directions.
- Standard Requirements and qualified RICEFW items start in Azure DevOps and are represented in qTest as linked Requirements.
- Requirement comments and RICEFW comments currently move from ADO to qTest.
- qTest requirement status can update the agreed ADO testing status field without opening a full reverse Requirement sync.
- Friendly support warnings are available through `ChatOpsEvent` whenever the integration uses a default, skips an optional unresolved field, or hits a recoverable issue.

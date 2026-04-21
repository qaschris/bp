# ADO to qTest Focused Customizations

## Purpose

This document summarizes the customer-specific enhancements added on top of the stock Azure DevOps to qTest integration. It is intended as a handover reference for support, future maintenance, and customer walkthrough sessions.

The six sections below are based on the customer enhancement list and map each requested capability to the current implementation in this repository.

## Solution Context

- Integration platform: qTest Pulse
- Runtime: Node.js webhook handlers
- Systems in scope: Azure DevOps, qTest Manager, qTest Pulse
- Additional utility: standalone Node.js migration script for pre-existing qTest requirements

## General Prerequisites

- qTest API token with access to the target qTest project
- Azure DevOps PAT with work item read/update permissions
- qTest Pulse rules configured against the correct webhook events
- qTest project field IDs configured in Pulse constants
- qTest project and Azure DevOps project structures aligned where required
- Consistent naming convention for linked records using the `WI<id>:` prefix
- qTest users provisioned in the target project when assignment sync is required

## Focused Customizations

### 1. Defect Error Handling

**Purpose**

Provide business-readable failure messages when qTest-to-ADO defect creation or subsequent sync actions fail, instead of relying only on console logging.

**Direction / Event**

- Bidirectional
- Most visible during defect creation in qTest and defect update processing across both systems

**What Changed from Stock**

- Added a shared `emitFriendlyFailure` pattern to publish simplified failure messages through `ChatOpsEvent`
- Added retry logic during qTest defect creation flow so incomplete qTest saves do not immediately fail the integration
- Added user-facing warnings when defaults are applied, such as unmapped team-to-area-path scenarios

**Pre-Requisites**

- A Pulse trigger named `ChatOpsEvent`
- qTest and ADO connection constants configured correctly
- Defect field IDs configured in Pulse constants

**Primary Files**

- `CreateDefectInAzureDevops.P1.js`
- `SyncDefectFromAzureDevopsWorkItem.P1.js`
- `UpdateDefectInAzureDevops.P1.js`

**Representative Code Snippet**

```js
function emitFriendlyFailure(details = {}) {
    const message =
        `Sync failed. Platform: ${details.platform}. ` +
        `Object Type: ${details.objectType}. ` +
        `Object ID: ${details.objectId}. ` +
        `Detail: ${details.detail || "Sync failed."}`;

    console.error(`[Error] ${message}`);
    emitEvent('ChatOpsEvent', { message });
}
```

### 2. Defect Assigned To Synchronization

**Purpose**

Keep the defect owner aligned across qTest and Azure DevOps, while respecting the realities of the customer's SSO model.

**Direction / Event**

- Bidirectional
- qTest defect created / updated
- Azure DevOps defect updated

**What Changed from Stock**

- qTest to ADO: resolve the qTest user ID to a usable identity before setting `System.AssignedTo`
- ADO to qTest: resolve the assigned user only within the target qTest project, rather than searching the full qTest tenant
- Active project user filtering now includes `inactive=false`
- Matching is limited to `username`, `ldap_username`, and `external_user_name`

**Pre-Requisites**

- Users must exist in both systems in a comparable form
- qTest users must be members of the target qTest project
- Pulse constants must include the assigned-to field IDs and tokens

**Primary Files**

- `CreateDefectInAzureDevops.P1.js`
- `SyncDefectFromAzureDevopsWorkItem.P1.js`
- `UpdateDefectInAzureDevops.P1.js`

**Representative Code Snippet**

```js
const response = await axios.get(url, {
    headers: standardHeaders,
    params: { inactive: false },
});

const user = users.find(candidate => {
    const keys = [
        candidate?.username,
        candidate?.ldap_username,
        candidate?.external_user_name,
    ];

    return keys.some(value => normalizeLabel(value) === normalizedIdentity);
});
```

### 3. Defect Assigned to Team / ADO Area Path Mapping

**Purpose**

Translate organizational ownership between qTest and Azure DevOps, where qTest uses a controlled dropdown and Azure DevOps uses `System.AreaPath`.

**Direction / Event**

- Bidirectional
- qTest defect created / updated
- Azure DevOps defect updated

**What Changed from Stock**

- Added mapping between qTest "Assigned to Team" and ADO `System.AreaPath`
- Added default fallback to `bp_Quantum\\Technical\\Testing` when no valid mapping is available
- qTest write path now resolves constrained field values dynamically using the qTest Fields API instead of hardcoded option IDs

**Pre-Requisites**

- qTest "Assigned to Team" field configured in the project
- ADO area paths aligned with the agreed naming convention
- Default area path agreed between project and support teams

**Primary Files**

- `CreateDefectInAzureDevops.P1.js`
- `SyncDefectFromAzureDevopsWorkItem.P1.js`
- `UpdateDefectInAzureDevops.P1.js`

**Representative Code Snippet**

```js
const qtestAssignedToTeamValue = await mapAreaPathToQtestTeamValue(adoAreaPath);

patchData.push({
    op: "add",
    path: `/fields/${constants.AzDoAreaPathFieldRef || "System.AreaPath"}`,
    value: assignedToTeamLabel
});
```

### 4. Requirement Area Path to qTest Module Placement

**Purpose**

Place new and updated qTest requirements into the correct qTest module structure based on Azure DevOps `AreaPath` and `IterationPath`.

**Direction / Event**

- Azure DevOps to qTest
- Requirement created / updated / deleted

**What Changed from Stock**

- Derived qTest folder path from ADO release and area path segments
- Automatically created missing qTest modules when the path did not already exist
- Added update filtering so non-relevant ADO field changes do not re-trigger unnecessary qTest rewrites
- Replaced hardcoded qTest dropdown value IDs with dynamic field-resolution logic

**Pre-Requisites**

- `RequirementParentID` configured in Pulse constants
- qTest module tree permissions available to the integration user
- ADO `AreaPath` and `IterationPath` populated on the work item

**Primary Files**

- `SyncRequirementFromAzureDevopsWorkItem.P1.js`

**Representative Code Snippet**

```js
const targetModuleId = await ensureModulePath(adoAreaPath, iterationPath);

if (requirementToUpdate) {
    await updateRequirement(..., targetModuleId);
} else {
    await createRequirement(..., targetModuleId);
}
```

### 5. RICEFW Feature Mapping with Requirement Test Cases

**Purpose**

Separate RICEFW / Change Request features from standard requirements and place them under a dedicated qTest parent module while still preserving area-path-based organization beneath that root.

**Direction / Event**

- Azure DevOps to qTest
- Feature created / updated / deleted

**What Changed from Stock**

- Added explicit RICEFW feature eligibility rules based on work item type, feature type, configuration, and state
- Added dedicated `FeatureParentID` root in qTest
- Added sync of RICEFW-specific fields such as Process Release, RICEFW ID, Testing Status, and related metadata
- Added update filtering to reduce the chance of event churn during UAT and support

**Pre-Requisites**

- `FeatureParentID` configured in Pulse constants
- ADO features use the expected `Custom.BPFeatureType` and `BP.ERP.RICEFW` values
- qTest requirement fields for RICEFW metadata are created and mapped

**Primary Files**

- `SyncRICEFWFeatureFromAzureDevops.js`

**Representative Code Snippet**

```js
if (!isRicefwFeature(event)) {
    console.log("[Info] Work item is not a RICEFW Feature. Exiting.");
    return;
}

const rootModuleId = constants.FeatureParentID;
const targetModuleId = await ensureModulePath(rootModuleId, adoAreaPath, iterationPath);
```

### 6. Migration of Existing qTest Requirements into New Folder Structure

**Purpose**

Bring pre-existing qTest requirements forward into the new module and metadata standards used by the customized integration.

**Direction / Event**

- qTest Pulse migration workflow
- Triggered manually or by internal Pulse webhook chaining

**What Changed from Stock**

- Added a dedicated Pulse queue rule for already-existing qTest requirements
- Added a dedicated Pulse worker rule that updates batches of existing qTest requirements
- Reads the linked Azure DevOps work item for each qTest requirement
- Recomputes the target qTest module path using the same rules as the live integration
- Uses internal trigger chaining so the queue rule can page through the old root folder and the worker can recall itself with remaining ids when needed
- Retains single-requirement test mode by allowing the queue rule to emit one explicit qTest requirement id
- Uses dynamic qTest field-value resolution for constrained fields

**Pre-Requisites**

- qTest and ADO Pulse constants populated
- Internal Pulse triggers created for the queue rule and the worker rule
- qTest migration source parent/root supplied in the kickoff payload for full-root runs
- target parent may be supplied in the kickoff payload, but normally reuses `RequirementParentID`
- Existing qTest requirements already use the `WI<id>:` naming convention
- Single-requirement test path validated before full-root execution

**Primary Files**

- `QueueRequirementMigration.P1.js`
- `ProcessRequirementMigrationBatch.P1.js`
- `bp_requirement_migration.js` as the local reference baseline

**Representative Code Snippet**

```js
await emitEvent("ProcessRequirementMigrationBatch.P1", {
    runId,
    requirementIds: batchIds,
    targetParentId,
});

if (remainingIds.length) {
    await emitEvent("ProcessRequirementMigrationBatch.P1", {
        runId,
        requirementIds: remainingIds,
        continuationCount: continuationCount + 1,
    });
}
```

## Current Production-Readiness Notes

- qTest constrained field values are now resolved dynamically instead of depending on fixed numeric IDs in the requirement and defect write paths
- Requirement sync includes stronger loop-prevention filtering on both the ADO-to-qTest and qTest-to-ADO status flows
- qTest project-scoped user resolution is now used for ADO-to-qTest assignment sync, with inactive users excluded
- ChatOps wording and formatting updates are still planned as a separate follow-up

## Repository Reference

- Defect creation from qTest: `CreateDefectInAzureDevops.P1.js`
- Defect sync from ADO to qTest: `SyncDefectFromAzureDevopsWorkItem.P1.js`
- Defect sync from qTest to ADO: `UpdateDefectInAzureDevops.P1.js`
- Requirement sync from ADO to qTest: `SyncRequirementFromAzureDevopsWorkItem.P1.js`
- RICEFW feature sync from ADO to qTest: `SyncRICEFWFeatureFromAzureDevops.js`
- Requirement status sync from qTest to ADO: `UpdateRequirementStatusInAzureDevops.P1.js`
- Requirement migration queue: `QueueRequirementMigration.P1.js`
- Requirement migration worker: `ProcessRequirementMigrationBatch.P1.js`
- Existing requirement migration reference script: `bp_requirement_migration.js`

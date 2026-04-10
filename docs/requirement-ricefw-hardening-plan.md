# Requirement And RICEFW Workflow Hardening Plan

This note captures the current strategy for improving the requirement-related workflows in a similar way to the recent defect workflow hardening pass.

## Scope

Files in scope:

- `SyncRequirementFromAzureDevopsWorkItem.P1.js`
- `SyncRICEFWFeatureFromAzureDevops.js`
- `UpdateRequirementStatusInAzureDevops.P1.js`

Primary goals:

- move ADO field access to Pulse constant-based refs
- improve qTest field-value handling using the Fields API where labels align
- keep explicit mappings only where ADO and qTest business values truly differ
- improve loop prevention and no-op detection
- standardize ChatOps warnings/failures and avoid duplicate reporting
- correct requirement assignee handling to use active project users only

## Current Findings

### Standard Requirement Sync

`SyncRequirementFromAzureDevopsWorkItem.P1.js` still contains:

- inline ADO field refs throughout the script
- direct string-based qTest `AssignedTo` handling instead of project-scoped user resolution
- only partial qTest field resolution through the Fields API
- duplicated create/update payload logic
- no true no-op detection before updating qTest
- failure-only ChatOps with no dedupe or warning helper

### RICEFW / Feature Sync

`SyncRICEFWFeatureFromAzureDevops.js` has the same structural issues as the standard requirement sync, plus:

- immediate exit when the work item no longer qualifies as a RICEFW Feature
- no explicit handling for previously-synced qTest requirements that later fall out of scope
- many likely-constrained qTest fields written as raw values

### qTest -> ADO Requirement Status Sync

`UpdateRequirementStatusInAzureDevops.P1.js` is narrower, but still needs:

- constant-based ADO field ref validation
- deduped ChatOps helpers
- better label-based status handling instead of relying only on hard-coded qTest ids
- the same cleanup/validation style used in the defect workflows

## Implementation Strategy

### Phase 1: Shared Foundation

Apply the same helper pattern used in the defect workflows:

- `normalizeText` with Unicode-safe normalization
- deduped `emitFriendlyFailure`
- deduped `emitFriendlyWarning`
- `buildAdoFieldRefs()` from Pulse constants only
- `validateRequiredConfiguration()` up front
- `getAdoFieldValue()` for ADO reads
- qTest field resolution helpers with optional inactive-value support
- qTest project-user lookup using `/api/v3/projects/{projectId}/users?inactive=false`

User matching rules:

- match only `username`
- match only `ldap_username`
- match only `external_user_name`
- do not match on `email`

### Phase 2: Standard Requirement Sync

Refactor `SyncRequirementFromAzureDevopsWorkItem.P1.js` to:

- build one normalized desired qTest requirement state
- dynamically resolve all constrained qTest fields where labels should match
- warn and skip optional unresolved fields rather than failing the whole sync
- fetch the current qTest requirement on update and compare before writing
- include `parentId` only when the target module actually changes
- keep the ADO updated-field filter, but add true no-op detection

Fields that likely need dynamic qTest resolution:

- complexity
- work item type
- priority
- requirement category / type
- application name
- fit gap
- BP entity

Assigned To should stop using free-text normalization and instead:

- resolve against active project users
- fall back to the configured service identity if needed
- emit a warning only if the overall qTest update succeeds

### Phase 3: RICEFW / Feature Sync

Refactor `SyncRICEFWFeatureFromAzureDevops.js` using the same structure as the standard requirement sync.

Additional RICEFW-specific work:

- separate “is this event in scope for sync” from “what should happen to an already-synced qTest requirement”
- define behavior for items that were synced previously but later become rejected, cancelled, or otherwise out of scope
- dynamically resolve constrained qTest values for:
  - testing status
  - feature type
  - area
  - RICEFW configuration
  - any other constrained custom fields

### Phase 4: qTest -> ADO Requirement Status Sync

Refactor `UpdateRequirementStatusInAzureDevops.P1.js` to:

- validate required constants early
- use deduped ChatOps helpers
- rely on the constant-based ADO testing-status field ref only
- prefer label-based qTest status handling where possible
- keep explicit status mapping only where business values truly differ
- retain the current loop-prevention checks and tighten them where helpful

## ADO Field Ref Constants Needed

Reuse existing constants where already defined:

- `AzDoTitleFieldRef`
- `AzDoAreaPathFieldRef`
- `AzDoAssignedToFieldRef`
- `AzDoStateFieldRef`
- `AzDoPriorityFieldRef`

Add or verify these requirement-related ADO field refs:

- `AzDoWorkItemTypeFieldRef` = `System.WorkItemType`
- `AzDoIterationPathFieldRef` = `System.IterationPath`
- `AzDoReasonFieldRef` = `System.Reason`
- `AzDoDescriptionFieldRef` = `System.Description`
- `AzDoAcceptanceCriteriaFieldRef` = `Microsoft.VSTS.Common.AcceptanceCriteria`
- `AzDoComplexityFieldRef` = `Custom.Complexity`
- `AzDoRequirementCategoryFieldRef` = `Custom.RequirementCategory`
- `AzDoApplicationNameFieldRef` = `Custom.ApplicationName`
- `AzDoFitGapFieldRef` = `BP.ERP.FitGap`
- `AzDoEntityFieldRef` = `Custom.Entity`
- `AzDoProcessReleaseFieldRef` = `Custom.ProcessRelease`
- `AzDoRICEFWIdFieldRef` = `Custom.RICEFWID`
- `AzDoTestingStatusFieldRef` = `Custom.TestingStatus`
- `AzDoBPFeatureTypeFieldRef` = `Custom.BPFeatureType`
- `AzDoAreaFieldRef` = `Custom.Area`
- `AzDoRICEFWConfigurationFieldRef` = `BP.ERP.RICEFW`

## Recommended Execution Order

1. `SyncRequirementFromAzureDevopsWorkItem.P1.js`
2. `SyncRICEFWFeatureFromAzureDevops.js`
3. `UpdateRequirementStatusInAzureDevops.P1.js`

## Open Questions For Later

- What should happen when a previously-synced RICEFW Feature no longer qualifies as RICEFW?
- Which requirement-side qTest fields are truly constrained in the customer project versus free text?
- Should requirement-side `AssignedTo` also use a specific qTest service-account fallback identity, matching the defect workflow approach?

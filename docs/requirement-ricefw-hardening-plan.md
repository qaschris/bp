# Requirement And RICEFW Workflow Hardening Plan

Updated: 2026-04-14

This document is the working plan for bringing the Requirement and
RICEFW/Feature workflows up to the same standard as the recent Defect workflow
hardening pass. It is meant to be actionable, not just descriptive.

## Scope

Files in scope:

- `SyncRequirementFromAzureDevopsWorkItem.P1.js`
- `SyncRICEFWFeatureFromAzureDevops.js`
- `UpdateRequirementStatusInAzureDevops.P1.js`

Primary goals:

- move ADO field access to Pulse constant-based refs
- dynamically resolve qTest constrained field values where labels should align
- keep explicit mappings only where business values truly differ
- improve requirement loop prevention and no-op detection
- standardize ChatOps warnings and failures
- correct qTest requirement assignee handling to use active project users only
- reduce unnecessary qTest writes and module moves

## Current Findings

### 1. Standard Requirement Sync

`SyncRequirementFromAzureDevopsWorkItem.P1.js` is functionally complete, but it
still has several structural gaps:

- ADO field refs are still hard-coded throughout the script.
- qTest `AssignedTo` is still handled as normalized free text rather than by
  resolving against active users in the project.
- Several qTest requirement fields are written as raw values instead of being
  resolved through the qTest Fields API.
- The script still rebuilds and sends broad qTest payloads instead of comparing
  current vs desired state first.
- The update URL always includes `parentId=${targetModuleId}`, which means the
  item can be re-parented even when the module does not actually need to change.
- ChatOps is still older style: mostly failure-oriented, no warning helper, and
  no dedupe pattern like the one now used in the Defect workflows.

Fields currently handled too loosely:

- `AssignedTo`
- `Iteration Path`
- `Application Name`
- `Fit Gap`
- `BP Entity`
- any other constrained requirement property whose value is currently sent
  directly from ADO to qTest

### 2. RICEFW / Feature Sync

`SyncRICEFWFeatureFromAzureDevops.js` has the same foundational issues as the
standard requirement sync, plus a few RICEFW-specific gaps:

- It writes many likely-constrained qTest fields as raw values:
  - `AssignedTo`
  - `Iteration Path`
  - `Application Name`
  - `Process Release`
  - `RICEFW ID`
  - `RICEFW Configuration`
  - `Testing Status`
  - `Feature Type`
  - `Area`
- It exits immediately if the work item no longer qualifies as a RICEFW Feature.
- That means a previously-synced qTest requirement can be left behind with no
  explicit lifecycle handling if the source item later becomes out-of-scope.
- Like the standard requirement sync, it does not yet use a clean desired-state
  plus no-op-detection pattern.

### 3. qTest -> ADO Requirement Status Sync

`UpdateRequirementStatusInAzureDevops.P1.js` is narrower, but still needs
cleanup:

- it uses a hard-coded fallback for the ADO testing-status field ref
- it still depends on a hard-coded qTest status-id map
- it does have useful loop-prevention checks already, but it has not yet been
  brought into the same constant-validation and ChatOps pattern as the Defect
  flows

Current strengths we should preserve:

- it checks changed field ids and skips when the qTest event does not mention
  the status field
- it skips updates from the configured sync user
- it compares current ADO testing status before patching

## Hardening Strategy

## Phase 1: Shared Foundation

The first pass should standardize the helper layer across all three scripts.

Core helpers to introduce or align:

- `normalizeText`
- Unicode-safe label normalization
- deduped `emitFriendlyFailure`
- deduped `emitFriendlyWarning`
- `buildAdoFieldRefs()` from Pulse constants only
- `validateRequiredConfiguration()` up front
- `getAdoFieldValue()` for ADO reads
- qTest field metadata lookup and option resolution helpers
- qTest project-user lookup via:
  - `/api/v3/projects/{projectId}/users?inactive=false`

Requirement user matching rules should mirror the Defect decision:

- match `username`
- match `ldap_username`
- match `external_user_name`
- do not match on `email`

Recommendation:

- add one configured qTest fallback identity for requirements as well, so the
  requirement workflows behave consistently with defects when assignee lookup
  fails

## Phase 2: Standard Requirement Sync

Refactor `SyncRequirementFromAzureDevopsWorkItem.P1.js` around one desired-state
model.

Target structure:

1. validate constants
2. build constant-based ADO field ref map
3. collect ADO source values
4. resolve qTest constrained values
5. determine target module id
6. fetch current qTest requirement when updating
7. compare current vs desired
8. write only when a real change exists

Required behavior changes:

- replace inline ADO refs with constants only
- resolve qTest constrained values dynamically instead of sending raw labels
- resolve `AssignedTo` through active project users
- treat optional unresolved values as warning-and-skip, not hard failure
- include `parentId` only when the target module actually changes
- suppress no-op writes so unchanged ADO updates do not rewrite qTest

Fields that should likely move to dynamic qTest resolution:

- `Complexity`
- `Work Item Type`
- `Priority`
- `Requirement Category`
- `Application Name`
- `Fit Gap`
- `BP Entity`
- `Iteration Path`

Module handling strategy:

- keep the current release-folder + area-path module derivation pattern
- continue deriving the release folder from ADO `IterationPath`
- only move the qTest requirement when the computed module id differs from the
  current parent

## Phase 3: RICEFW / Feature Sync

Refactor `SyncRICEFWFeatureFromAzureDevops.js` to match the standard
Requirement sync structure, then add RICEFW-specific lifecycle handling.

Core changes:

- move to constant-based ADO field refs
- resolve constrained qTest fields dynamically
- use active project-user lookup for `AssignedTo`
- add true no-op detection before qTest writes
- only move modules when the target parent actually changes

RICEFW-specific fields that should be treated as constrained unless proven
otherwise:

- `Testing Status`
- `Feature Type`
- `Area`
- `RICEFW Configuration`
- `Process Release`

Lifecycle decision that still needs to be made explicit:

- if a work item was previously synced as a RICEFW Feature but later becomes
  rejected, cancelled, or otherwise out-of-scope, should we:
  - delete the qTest requirement
  - leave it in place and stop syncing
  - mark it somehow and keep it visible

Current recommendation:

- do not silently exit without handling previously-synced items
- make the out-of-scope behavior explicit in code once the business rule is
  confirmed

## Phase 4: qTest -> ADO Requirement Status Sync

Refactor `UpdateRequirementStatusInAzureDevops.P1.js` with the same cleanup
style used in the Defect scripts.

Required changes:

- validate required constants early
- use deduped ChatOps helpers
- rely on the constant-based ADO testing-status field ref only
- tighten log messages and error classification

Status handling strategy:

- prefer label-aware handling where possible
- keep explicit mapping where the business values truly differ
- do not remove the changed-field loop protection
- do not remove the sync-user protection

Because this script is intentionally narrow, it does not need the full module
or requirement-payload machinery of the other two scripts.

## Field Handling Model

The requirement pass should follow the same hybrid model now used on defects.

Use explicit business mappings only when labels do not truly align:

- testing-status values, if business semantics differ
- any RICEFW field with customer-specific terminology differences

Use dynamic qTest field resolution when labels are expected to align:

- complexity
- requirement category
- application name
- fit gap
- entity
- testing status, if labels really do align
- feature type
- area
- iteration path

If an optional constrained qTest value cannot be resolved:

- warn
- skip that field
- continue the overall sync if everything else is valid

## Constants To Add Or Verify

### ADO Field Ref Constants

Reuse where already defined:

- `AzDoTitleFieldRef`
- `AzDoAreaPathFieldRef`
- `AzDoAssignedToFieldRef`
- `AzDoStateFieldRef`
- `AzDoPriorityFieldRef`

Add or verify for standard Requirements:

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

Add or verify for RICEFW / Feature:

- `AzDoProcessReleaseFieldRef` = `Custom.ProcessRelease`
- `AzDoRICEFWIdFieldRef` = `Custom.RICEFWID`
- `AzDoTestingStatusFieldRef` = `Custom.TestingStatus`
- `AzDoBPFeatureTypeFieldRef` = `Custom.BPFeatureType`
- `AzDoAreaFieldRef` = `Custom.Area`
- `AzDoRICEFWConfigurationFieldRef` = `BP.ERP.RICEFW`

### qTest Field ID Constants To Verify

Standard Requirement:

- `RequirementComplexityFieldID`
- `RequirementWorkItemTypeFieldID`
- `RequirementPriorityFieldID`
- `RequirementTypeFieldID`
- `RequirementAssignedToFieldID`
- `RequirementIterationPathFieldID`
- `RequirementApplicationNameFieldID`
- `RequirementFitGapFieldID`
- `RequirementBPEntityFieldID`
- `RequirementStreamSquadFieldID`

RICEFW / Feature:

- `RequirementProcessReleaseFieldID`
- `RequirementRicefwIdFieldID`
- `RequirementRICEFWConfigurationFieldID`
- `RequirementTestingStatusFieldID`
- `RequirementFeatureTypeFieldID`
- `RequirementAreaFieldID`

Status Bridge:

- `RequirementStatusFieldID`
- `AzDoTestingStatusFieldRef`

## ChatOps And Error Handling Rules

Requirement workflows should follow the same behavior standard now used in the
Defect scripts:

- configuration problems should fail early and once
- informational business-rule skips should not be failures
- optional unresolved values should be warnings only if the overall sync succeeds
- duplicate user-facing failure messages for the same root cause should be
  deduped

Object identity conventions:

- ADO messages: work item id
- qTest messages: requirement id plus pid when available

## Recommended Execution Order

1. `SyncRequirementFromAzureDevopsWorkItem.P1.js`
2. `SyncRICEFWFeatureFromAzureDevops.js`
3. `UpdateRequirementStatusInAzureDevops.P1.js`

This order keeps the broadest shared patterns first and leaves the narrow
status bridge last.

## Suggested Test Matrix

For the standard Requirement workflow:

- create from ADO into qTest
- update with no meaningful changes
- update with module move required
- update with unresolved optional constrained field
- update with unresolved assignee

For RICEFW / Feature:

- create qualifying feature
- update qualifying feature
- update feature that becomes out-of-scope
- delete event for previously-synced feature

For qTest -> ADO status:

- qTest event where only non-status fields changed
- qTest event from sync user
- qTest event with real status change
- no-op case where ADO already has the desired testing status

## Open Decisions

- What should happen when a previously-synced RICEFW item later becomes
  rejected, cancelled, or otherwise non-qualifying?
- What should the configured qTest fallback assignee identity be for
  Requirement workflows?
- Which requirement-side qTest fields in the customer project are definitely
  constrained vs free text?
- Should any requirement-side status values remain explicitly mapped for
  business reasons, or can more of them move to label-based matching?

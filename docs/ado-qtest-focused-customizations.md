# ADO to qTest Focused Customizations

## Purpose

This document is the current technical handover reference for the BP Azure DevOps and qTest integration rules in this repository.

It is organized by live rule and sync direction rather than by the original enhancement request list. The goal is to describe what the code does today, including the newer comment handling, the mapping overrides, the fallback rules, and the operational behaviors that matter during support and go-live.

## Solution Context

- Integration platform: qTest Pulse
- Runtime: Node.js-based Pulse webhook rules
- Systems in scope: Azure DevOps, qTest Manager, qTest Pulse
- Current deployment model: live sync rules in Pulse, plus a two-rule Pulse migration workflow for pre-existing qTest requirements
- Local reference only: `bp_requirement_migration.js` remains in the repository as a baseline, but it is no longer the deployment target

## Shared Integration Behaviors

- qTest constrained dropdowns are resolved dynamically through qTest field metadata instead of relying on hardcoded option ids.
- Friendly warnings and failures are emitted through `ChatOpsEvent` so support receives business-readable messages instead of only console output.
- Linked qTest requirements use the `WI<id>:` naming convention so the Azure DevOps work item id can be recovered later.
- Defects created from qTest are updated to include the linked ADO work item id in the qTest defect name after create succeeds.
- Defect field synchronization is asynchronous across qTest, Pulse, and Azure DevOps, so short delays are expected after create and update activity.
- The configured default ADO area path comes from `constants.AreaPath`.
- The configured default ADO iteration path comes from `constants.IterationPath` or `constants.AzDoDefaultIterationPath`.
- Requirement and RICEFW comments currently flow one way from ADO to qTest.
- Defect comments currently flow both directions between ADO and qTest.
- Comment sync relies on origin markers such as `[From ADO]` and `[From qTest]`, plus CID markers where available, to reduce echo loops and duplicates.

## Rule Reference

### 1. CreateDefectInAzureDevops.P1.js

**Direction / Trigger**

- qTest to ADO
- Triggered when a qTest defect is created in the configured BP project
- This is the authoritative creation path for defects; ADO-created defects are not used to create new qTest defects

**Core Behavior**

- Reads the full qTest defect using retry logic so partially saved defects do not fail immediately.
- Creates an ADO `Bug` and adds a hyperlink back to the originating qTest defect.
- Updates the qTest defect name after ADO creation so the linked work item id is embedded for later sync.
- Emits friendly warnings when a default is used or when a non-critical field could not be mapped cleanly.

**ADO Fields Written**

- `System.Title`
- `Microsoft.VSTS.TCM.ReproSteps`
- `System.Tags`
- `System.State`
- `Microsoft.VSTS.Common.Severity`
- `Microsoft.VSTS.Common.Priority`
- `System.AreaPath`
- `System.AssignedTo`
- `BP.ERP.DefectType`
- `Custom.BugStage`
- `Custom.bpCreatedBy`
- `BP.ERP.ExternalReference`
- `Custom.Application`
- `Custom.SiteName`
- `System.IterationPath`
- `Microsoft.VSTS.CMMI.RootCause`
- `Microsoft.VSTS.CMMI.ProposedFix`
- `Microsoft.VSTS.Scheduling.TargetDate`

**Mapping and Override Behavior**

- qTest severity ids map to ADO severity labels `1 - Critical` through `4 - Low`.
- qTest priority ids map to ADO priorities `1` through `4`.
- qTest defect type ids map to the customer ADO defect type values such as `Code`, `Data`, `Infrastructure`, `User Authorization`, and `Automation`.
- qTest status ids map to the BP ADO defect states such as `New`, `In Analysis`, `Awaiting Implementation`, `Resolved`, `Retest`, `Reopened`, `Closed`, `On Hold`, `Rejected`, and `Triage`.
- qTest affected release values map to the BP bug stage values such as `P&O_R1_SIT1`, `P&O_R1_DC1`, `P&O_R1_UAT`, and `Unit Testing`.
- qTest `Assigned to Team` is translated to ADO `System.AreaPath`. If the qTest value is blank or cannot be resolved, the rule defaults to `constants.AreaPath` and emits a warning.
- qTest `Iteration Path` is matched against ADO classification nodes. If no match is found, the rule defaults to the configured iteration fallback and emits a warning.
- qTest `Assigned To` is resolved to a usable identity before writing `System.AssignedTo`.
- qTest `Root Cause` is written using the configured ADO field reference after label normalization rather than by qTest numeric id.

**Operational Notes**

- The active creation path intentionally does not send `Closed Date` or `Resolved Reason` during the initial ADO bug create. Those values are handled on later update flows instead.
- `Target Date` is sent during create when present.
- Missing optional mapped values are usually skipped with warnings instead of causing the entire defect create to fail.

### 2. SyncDefectFromAzureDevopsWorkItem.P1.js

**Direction / Trigger**

- ADO to qTest
- Handles `workitem.updated` for linked defects
- ADO defect create events are logged and ignored because the BP defect lifecycle starts in qTest
- ADO defect delete events are logged and ignored; linked qTest defects are not automatically deleted

**Core Behavior**

- Finds the linked qTest defect using the stored `WI<id>:` convention.
- Updates the existing qTest defect with the current ADO field values.
- Converts ADO comment activity into qTest defect comments.
- Uses project-scoped qTest user lookup for inbound assignment handling.

**qTest Fields Updated**

- Defect title / linked name
- Description / repro details
- Status
- Severity
- Priority
- Defect Type
- Root Cause
- Proposed Fix
- External Reference
- Resolved Reason
- Application
- Source Team
- Site Name
- Assigned To
- Assigned to Team
- Affected Release
- Target Date
- Iteration Path
- qTest comments from ADO comments

**Mapping and Override Behavior**

- Most ADO states are used by label, but the rule explicitly overrides `Active` to qTest `In Analysis` and `Cancelled` to qTest `Rejected`.
- ADO bug stage values such as `P&O_R1_SIT1`, `P&O_R1_DC1`, and `Unit Testing` map back to the qTest `Affected Release` dropdown values.
- ADO severity and priority are converted back to qTest numeric option values.
- ADO `AreaPath` is translated to qTest `Assigned to Team` through dynamic qTest field resolution.
- If the ADO area path cannot be resolved in qTest, the rule falls back to the configured default area path label and emits a warning.
- ADO `Assigned To` is resolved only against active users in the target qTest project.
- If inbound user resolution fails, the rule falls back to the configured service identity behavior rather than failing the whole sync.
- ADO `Root Cause`, `Resolved Reason`, and `Iteration Path` are resolved dynamically against qTest field options. When optional values cannot be resolved, the field is left unchanged and a warning is emitted.
- ADO `Target Date` is converted to qTest date-time format before write.

**Comment Behavior**

- ADO comment changes are detected through `System.CommentCount`, `System.History`, and discussion metadata.
- qTest comments are created with a `[From ADO]` marker.
- Duplicate ADO comments are skipped by CID-style dedupe.
- The current live comment behavior uses the qTest comments API for user-visible comment sync.
- `DefectDiscussionFieldID` remains an optional configured field in the defect payload and force-change path, but the main end-user comment experience is now the qTest comments stream.

**Current Timing Limitation**

- The current inbound defect field-sync path still writes the mapped ADO state back to qTest on each qualifying defect update event.
- It does not yet use the same desired-state no-op evaluation pattern that the standard requirement sync uses.
- It also does not yet suppress inbound defect field sync when the ADO revision was authored by the integration sync user.
- Because of that, rapid successive qTest defect edits can still produce stale ADO-backed callbacks that overwrite a newer qTest field value.
- Current customer guidance is to let one defect save and sync settle before making another broad round of defect field edits.

### 3. UpdateDefectInAzureDevops.P1.js

**Direction / Trigger**

- qTest to ADO
- Triggered by qTest defect update events on linked defects

**Core Behavior**

- Reads the linked ADO work item from the qTest defect name.
- Syncs qTest comments to the ADO work item comments API before deciding whether a field patch is needed.
- Applies only the field changes that actually differ from the current ADO work item values.
- Uses classification path resolution for outbound area path and iteration path writes.

**ADO Fields Updated**

- `System.Title`
- `Microsoft.VSTS.TCM.ReproSteps`
- `Microsoft.VSTS.Common.Severity`
- `Microsoft.VSTS.Common.Priority`
- `System.State`
- `BP.ERP.DefectType`
- `Microsoft.VSTS.CMMI.RootCause`
- `Microsoft.VSTS.CMMI.ProposedFix`
- `Microsoft.VSTS.Common.ResolvedReason`
- `Custom.Application`
- `Custom.Source_Team`
- `Custom.SiteName`
- `System.AreaPath`
- `System.AssignedTo`
- `Custom.BugStage`
- `Microsoft.VSTS.Scheduling.TargetDate`
- `System.IterationPath`
- ADO work item comments

**Mapping and Override Behavior**

- Outbound status mapping is label-aware first, then falls back to qTest status ids.
- The current outbound status map explicitly includes qTest status label `Active` to ADO state `Active`.
- qTest affected release values map back to BP bug stage values such as `P&O_R1_SIT1`, `P&O_R1_DC1`, and `P&O_R1_UAT`.
- qTest `Assigned to Team` is translated to ADO `System.AreaPath`; if the qTest value is blank or cannot be resolved, the rule defaults to `constants.AreaPath` and emits a warning.
- qTest `Iteration Path` is resolved against the live ADO classification tree; if no match is found, the rule uses the configured iteration fallback and emits a warning.
- qTest `Assigned To` is translated to a usable ADO identity when possible; a default service identity path is available when it cannot be resolved.
- qTest `Root Cause` is normalized before being written to the configured ADO picklist field.
- `Resolved Reason` is intentionally not written when the outbound ADO state is `New` or `Active`.
- If qTest `Resolved Reason` is blank and the ADO state is not locked, the rule clears the ADO resolved reason field.
- qTest descriptions are sanitized to remove the embedded ADO link block before the description is pushed back into ADO.

**Comment Behavior**

- qTest defect comments are posted to the ADO work item comments API.
- Outbound comments are formatted with `[From qTest]`.
- CID markers are used when possible so duplicate outbound comment posts can be skipped.
- Comments are synced even when no field-level patch is required, so comment-only qTest activity is not lost.
- If `DefectDiscussionFieldID` is configured, the rule writes a timestamp back to qTest after outbound comment sync to force downstream change detection.

**Operational Notes**

- The outbound defect path already skips qTest updates from the configured sync user and suppresses no-op ADO field patches.
- Those outbound protections reduce echo churn, but they do not fully eliminate stale inbound ADO callbacks until the inbound defect path is hardened as well.

### 4. SyncRequirementFromAzureDevopsWorkItem.P1.js

**Direction / Trigger**

- ADO to qTest
- Handles requirement create, update, comment-only update, and delete flows

**Core Behavior**

- Creates or updates qTest requirements under the correct qTest module path.
- Deletes the linked qTest requirement when the ADO work item is deleted.
- Prevents most field-loop churn by processing only relevant updated fields on normal update events.
- Allows comment-only ADO updates to run even when no synced requirement field changed.
- Creates missing qTest modules on demand.

**Module Placement Behavior**

- The root for standard requirements is `RequirementParentID`.
- The release folder is derived from `System.IterationPath`, using the `P&O Release <number>` naming convention where available.
- Additional module levels are derived from the ADO `AreaPath`.
- The requirement is moved only when the computed target module differs from its current qTest parent.

**qTest Fields Updated**

- Requirement name using the `WI<id>:` prefix
- Requirement description
- Stream / Squad from ADO area path
- Complexity
- Work Item Type
- Priority
- Requirement Category
- Assigned To as normalized text
- Iteration Path
- Application Name
- Fit Gap
- BP Entity
- qTest comments copied from ADO comments

**Description Behavior**

- The qTest description includes an ADO hyperlink plus a structured snapshot of ADO metadata.
- The current description block includes Type, Area, Iteration, State, Reason, Complexity, Acceptance Criteria, and Description.
- Acceptance Criteria is embedded inside the description block in the standard requirement rule rather than being written to a separate qTest acceptance-criteria property.

**Mapping and Override Behavior**

- Complexity, Work Item Type, Priority, Requirement Category, Iteration Path, Application Name, Fit Gap, and BP Entity are resolved dynamically against qTest fields.
- Optional fields that cannot be resolved are left unchanged and reported as warnings rather than failing the whole requirement sync.
- `Assigned To` is currently stored as normalized display text in qTest. The standard requirement rule does not currently resolve the user to a qTest project user id.
- The rule compares the full desired state to the current qTest requirement before writing, so no-op ADO updates do not rewrite qTest unnecessarily.

**Comment Behavior**

- Requirement comments currently flow one way from ADO to qTest.
- The rule reads ADO comments from the work item comments API.
- qTest comments are created or updated with a `[From ADO]` prefix and a `[CID:<ADO comment id>]` marker.
- Comment-only ADO events are supported.

### 5. UpdateRequirementStatusInAzureDevops.P1.js

**Direction / Trigger**

- qTest to ADO
- Triggered by qTest requirement update events when the qTest status field changed

**Core Behavior**

- This is a narrow bridge rule, not a full requirement field sync.
- It reads the linked ADO work item id from the qTest requirement name.
- It updates only the configured ADO testing-status field.
- It skips work when the current ADO testing status already matches the qTest value.

**Status Handling**

- The qTest source field is `RequirementStatusFieldID`.
- The destination field is `AzDoTestingStatusFieldRef`.
- The rule currently recognizes the BP testing status values represented by the configured ids and labels such as `SIT Dry Run In Progress`, `SIT Dry Run Complete`, `SIT 1 In Progress`, `SIT 1 Complete`, `SIT 2 In Progress`, `SIT 2 Complete`, `UAT In Progress`, and `UAT Complete`.
- If the mapped value is `New`, the rule intentionally skips the ADO update by design.

**Loop Prevention**

- If the qTest event does not reference the configured status field id, the rule exits.
- If the qTest updater matches `SyncUserRegex`, the rule exits to avoid integration loops.

### 6. SyncRICEFWFeatureFromAzureDevops.js

**Direction / Trigger**

- ADO to qTest
- Handles feature create, update, comment-only update, and delete flows for qualifying RICEFW items

**Qualification Rules**

- ADO work item type must be `Feature`.
- The BP feature type must be `RICEFW` or `Change Request`.
- The configured RICEFW classification must be one of `Enhancement`, `Form`, `Interface`, `Report`, or `Workflow`.
- The ADO work item state must not be `Rejected` or `Cancelled`.

**Core Behavior**

- Creates or updates a qTest requirement under the configured `FeatureParentID`.
- Builds the module tree from the release folder plus the ADO area path.
- Deletes the linked qTest requirement when the ADO feature is deleted.
- If an existing linked item no longer qualifies as a RICEFW feature during an update, the rule leaves the qTest record unchanged and emits a warning instead of silently deleting or repurposing it.

**qTest Fields Updated**

- Requirement name using the `WI<id>:` prefix
- Requirement description
- Stream / Squad from ADO area path
- Work Item Type forced to `Feature`
- Assigned To as normalized text
- State
- Reason
- Acceptance Criteria
- Complexity
- Priority
- Requirement Category from the ADO feature type
- Iteration Path
- RICEFW Configuration
- Testing Status
- qTest comments copied from ADO comments

**Description Behavior**

- The qTest description includes the ADO hyperlink plus Type, Area Path, Iteration, State, Reason, RICEFW ID, Acceptance Criteria, and Description.
- The current live rule includes `RICEFW ID` inside the description block.
- The current live rule does not create separate qTest property writes for `RICEFW ID` or `Process Release`.

**Mapping and Override Behavior**

- Required qTest dropdown properties are dynamically resolved from labels rather than hardcoded numeric ids.
- Optional values that cannot be resolved are left unchanged with warnings.
- `Assigned To` is currently stored as normalized text in qTest, not resolved to a qTest project user id.
- Comment-only ADO updates are allowed to run even when no other synced field changed.

**Comment Behavior**

- RICEFW comments currently flow one way from ADO to qTest.
- The rule uses the same `[From ADO]` and `[CID:<ADO comment id>]` pattern as the standard requirement rule.

### 7. QueueRequirementMigration.P1.js

**Direction / Trigger**

- Pulse internal workflow
- Manual kickoff or chained queue execution

**Core Behavior**

- Starts a requirement migration run by reading existing qTest requirements under a supplied source parent.
- Can also run in single-item test mode using an explicit requirement id.
- Splits the source requirements into batches and invokes the worker rule for each batch.
- Re-invokes itself for additional pages when the source root contains more items.

**Kickoff Inputs**

- `sourceParentId`
- `targetParentId` or fallback to `RequirementParentID`
- `singleRequirementId`
- `page`
- `pageSize`
- `batchSize`
- `runId`

**Operational Notes**

- The queue rule name is `QueueRequirementMigration.P1`.
- The worker rule name is `ProcessRequirementMigrationBatch.P1`.
- Friendly informational progress messages are emitted through `ChatOpsEvent`.

### 8. ProcessRequirementMigrationBatch.P1.js

**Direction / Trigger**

- Pulse internal workflow
- Invoked by the queue rule with one batch of qTest requirement ids

**Core Behavior**

- Reads each qTest requirement in the batch.
- Extracts the linked ADO work item id from the qTest requirement name.
- Reads the live ADO work item and rebuilds the desired qTest state using the same field and module rules as the live standard requirement sync.
- Updates only requirements that are actually out of sync.
- Requeues remaining ids if the worker approaches its runtime budget.

**Field and Module Behavior**

- Uses the same description format, field set, dynamic dropdown resolution, and module-placement rules as the standard requirement sync.
- Moves the requirement only when the computed target parent differs from the current qTest parent.
- Leaves optional unresolved constrained fields unchanged and emits warnings rather than failing the entire run.

**Runtime Controls**

- Default `maxRunMs` is 240000.
- Remaining ids are requeued with `continuationCount + 1` when the worker yields.

## Current Codebase Notes

- The live requirement and RICEFW rules now include ADO to qTest comment sync.
- The live defect rules now include two-way comment sync between qTest and ADO.
- The live defect creation rule intentionally suppresses `Closed Date` and `Resolved Reason` during initial ADO bug create.
- The live standard requirement rule still stores `Assigned To` as normalized text rather than resolving it to a qTest project user id.
- The migration path now lives in Pulse rules; the standalone migration script remains only as a local reference.

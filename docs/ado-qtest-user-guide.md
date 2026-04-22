# BP ADO and qTest User Guide

## Purpose

This guide explains the BP Azure DevOps and qTest integration from an end-user point of view.

It is intentionally less technical than the support handover document. The focus here is what qTest users and Azure DevOps users should expect to happen, which fields are included, and where the important workflow boundaries are.

## What Users Should Expect Overall

- Defects are created in qTest first and then pushed into Azure DevOps.
- Standard requirements and RICEFW features start in Azure DevOps and are created or updated in qTest from there.
- Requirement comments currently flow from ADO to qTest only.
- Defect comments currently flow both directions.
- qTest requirement status can update the ADO testing status field, but other requirement fields do not flow from qTest back into ADO.

## 1. qTest Defects to Azure DevOps

**What happens**

- When a user creates a defect in qTest, the integration creates a linked ADO bug.
- After the ADO bug is created, the qTest defect is updated so it contains the linked work item id.
- The ADO bug also contains a hyperlink back to the qTest defect.

**Fields users will see carried into ADO**

- Summary
- Description / repro steps
- Status
- Severity
- Priority
- Defect Type
- Assigned To
- Assigned to Team / Area Path
- Affected Release / Bug Stage
- External Reference
- Application
- Site Name
- Iteration Path
- Root Cause
- Proposed Fix
- Target Date

**Important user-facing rules**

- If `Assigned to Team` is blank or does not match a valid ADO area path, the integration uses the configured default area path and logs a warning.
- If the qTest iteration path does not match a valid ADO iteration path, the integration uses the configured default iteration path and logs a warning.
- `Closed Date` and `Resolved Reason` are not part of the initial create into ADO. Those fields are handled later through update flows.

## 2. Azure DevOps Defect Updates Back to qTest

**What happens**

- Updates made in ADO to a linked bug are pushed back into the matching qTest defect.
- This does not create new qTest defects from ADO.
- Deleting a bug in ADO does not automatically delete the qTest defect.

**Fields users will see updated in qTest**

- Title / linked name
- Description
- Status
- Severity
- Priority
- Defect Type
- Assigned To
- Assigned to Team
- Affected Release
- External Reference
- Root Cause
- Proposed Fix
- Resolved Reason
- Application
- Source Team
- Site Name
- Target Date
- Iteration Path

**Important user-facing rules**

- ADO `Active` is treated as qTest `In Analysis`.
- ADO `Cancelled` is treated as qTest `Rejected`.
- If an inbound dropdown value cannot be resolved in qTest, that field is skipped and a warning is logged instead of breaking the whole sync.

## 3. Defect Comments

**What happens**

- ADO comments on linked defects are copied into qTest comments.
- qTest comments on linked defects are copied into ADO work item comments.

**What users will see**

- Comments coming from ADO are marked as coming from ADO.
- Comments coming from qTest are marked as coming from qTest.
- Duplicate comments are suppressed by the integration so the same note should not keep reappearing on replayed events.

**Important user-facing rules**

- Comment-only activity is supported. A user does not need to change another field for the comment sync to run.
- The integration adds origin markers behind the scenes to reduce echo loops.

## 4. Standard Requirements from Azure DevOps to qTest

**What happens**

- When a standard ADO requirement is created or updated, the integration creates or updates a qTest requirement.
- The qTest requirement is placed under the correct qTest module tree based on the ADO release and area path.
- If the ADO requirement is deleted, the matching qTest requirement is deleted.

**How qTest users will recognize the record**

- The qTest requirement name starts with `WI<id>:` so the linked ADO work item is easy to identify.
- The qTest description includes a direct link back to the ADO work item plus a structured summary of the ADO content.

**Fields users will see in qTest**

- Requirement name from the ADO title
- Description block containing ADO link, type, area, iteration, state, reason, complexity, acceptance criteria, and description
- Stream / Squad from ADO area path
- Complexity
- Work Item Type
- Priority
- Requirement Category
- Assigned To as text
- Iteration Path
- Application Name
- Fit Gap
- BP Entity

**Important user-facing rules**

- The requirement is moved only when its computed qTest module location really changes.
- If optional dropdown values such as Application Name, Fit Gap, or BP Entity cannot be matched in qTest, those fields are left unchanged and the overall requirement still syncs.
- Requirement comments currently flow from ADO to qTest only.

## 5. qTest Requirement Status Back to Azure DevOps

**What happens**

- When a qTest user updates the configured requirement status field, the integration can update the ADO testing status field on the linked work item.

**What this does not do**

- It does not push the full requirement back to ADO.
- It does not update arbitrary ADO fields from qTest.

**Status values users should expect**

- SIT Dry Run In Progress
- SIT Dry Run Complete
- SIT 1 In Progress
- SIT 1 Complete
- SIT 2 In Progress
- SIT 2 Complete
- UAT In Progress
- UAT Complete

**Important user-facing rules**

- The integration only reacts when the qTest status field changed.
- If the status already matches in ADO, no patch is sent.
- The integration intentionally does not push a `New` status back into ADO.

## 6. RICEFW Features from Azure DevOps to qTest

**What happens**

- Qualifying RICEFW or Change Request features in ADO create or update qTest requirements under the dedicated RICEFW qTest root.
- The qTest module path is still organized by release and area path beneath that root.
- If the ADO feature is deleted, the linked qTest requirement is deleted.

**Which ADO items qualify**

- Work Item Type must be `Feature`
- Feature Type must be `RICEFW` or `Change Request`
- RICEFW configuration must be one of the supported BP values
- State must not be `Rejected` or `Cancelled`

**Fields users will see in qTest**

- Requirement name from the ADO title
- Description block containing ADO link, type, area path, iteration, state, reason, RICEFW ID, acceptance criteria, and description
- Stream / Squad
- Work Item Type
- Assigned To as text
- State
- Reason
- Acceptance Criteria
- Complexity
- Priority
- Requirement Category
- Iteration Path
- RICEFW Configuration
- Testing Status

**Important user-facing rules**

- RICEFW comments currently flow from ADO to qTest only.
- If a previously synced item no longer qualifies as RICEFW during an update, the integration leaves the existing qTest item in place and logs a warning instead of silently deleting it.

## 7. One-Time Migration of Older qTest Requirements

**What happens**

- Older qTest requirements that existed before the new module structure can be migrated into the correct folders using the Pulse migration workflow.
- This is an admin/support activity, not a normal day-to-day end-user workflow.

**What users should expect**

- Existing requirements are moved into the newer release-and-area-based qTest folder structure.
- The migration uses the linked ADO work item to decide the correct destination.
- The migration uses the same field logic as the live standard requirement sync.

## Important Boundaries and Defaults

- New defects are expected to start in qTest, not in ADO.
- Requirement and RICEFW comments are one-way from ADO to qTest.
- Defect comments are two-way.
- Standard requirement `Assigned To` and RICEFW `Assigned To` are currently stored as text in qTest rather than as qTest user assignment objects.
- If an area path, iteration path, or optional dropdown value cannot be matched, the integration prefers a warning and a safe fallback over a hard failure whenever that is practical.

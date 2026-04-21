# Comment Sync Update Plan

## Goal

Merge the customer's existing comment synchronization behavior into the current finished rules with the least possible reinterpretation:

- Requirements: one-way `ADO -> qTest`
- Defects: two-way `ADO <-> qTest`

The intent is to preserve the current finished field-sync logic and port the customer's comment behavior from the older variants in `../bp-changes` with only the minimum compatibility changes needed to fit the current files.

## What I Found

The older comment work exists in four files:

- `../bp-changes/SyncRequirementFromAzureDevopsWorkItem.P1.txt`
- `../bp-changes/SyncRICEFWFeatureFromAzureDevopsP1.txt`
- `../bp-changes/SyncDefectFromAzureDevopsWorkItem.p1.txt`
- `../bp-changes/UpdateDefectInAzureDevops.P1.txt`

The current rules have drifted a lot from those versions, so we should not merge the old files wholesale.

- Requirement sync diff is large: about `734 insertions / 780 deletions`
- RICEFW feature sync diff is moderate: about `45 insertions / 179 deletions`
- Defect inbound sync diff is very large: about `1389 insertions / 331 deletions`
- Defect outbound sync diff is very large: about `1163 insertions / 210 deletions`

That means the safest path is to transplant only the customer comment-specific behavior into the modern entry points.

## Current-State Notes

### Requirements

Current requirement rules do not sync comments yet.

- `SyncRequirementFromAzureDevopsWorkItem.P1.js`
- `SyncRICEFWFeatureFromAzureDevops.js`

Both rules also short-circuit update events based on a curated field list. Comment-only ADO updates are currently skipped because `System.CommentCount` is not part of the relevant change set.

### Defects

Current defect rules are only partially comment-aware.

- `SyncDefectFromAzureDevopsWorkItem.P1.js` already reads ADO comments and writes a rendered discussion snapshot into `DefectDiscussionFieldID`
- `UpdateDefectInAzureDevops.P1.js` does not currently post qTest comments back into ADO

The older customer logic uses the actual qTest comments endpoints for both directions on defects, which is closer to true comment sync than the current discussion-field snapshot.

## Second-Pass Findings

These are the items most likely to be missed if we move too quickly.

### 1. Requirement/RICEFW helper signature mismatch

In the customer requirement-style files, `doRequest(...)` accepts an optional headers argument and is reused for both qTest and ADO calls.

In the current finished files:

- `SyncRequirementFromAzureDevopsWorkItem.P1.js`
- `SyncRICEFWFeatureFromAzureDevops.js`

`doRequest(...)` is qTest-only and always uses `standardHeaders`.

So a direct merge still requires one compatibility change:

- either extend `doRequest(...)` to accept a headers override like the customer version
- or add a separate ADO request helper and point the merged comment logic at that

Without that, ADO comment reads will not be wired correctly.

### 2. Requirement comment-only ADO updates will currently be skipped

The customer requirement file did not have the same update-field guard that the current finished requirement rule now has.

Current file:

- `SyncRequirementFromAzureDevopsWorkItem.P1.js`

The current rule will return early on update events when only non-synced fields changed. A pure comment-only ADO event will hit that path today.

That means a faithful merge still needs one structural adjustment:

- if a comment was added, the comment sync path must be allowed to run even when the normal field-sync guard says "skip"

This is not redesign; it is required to preserve the customer comment behavior inside the newer rule structure.

### 3. Customer RICEFW comment logic has two mechanical issues

In `../bp-changes/SyncRICEFWFeatureFromAzureDevopsP1.txt`:

- comment sync on update is placed behind the current-style update guard, so pure comment-only events may still never reach it
- the create path calls `syncCommentsIfNeeded(..., response?.id)` even though `response` is not defined there

So a faithful merge still requires a minimal runtime fix for RICEFW:

- either return the created requirement from `createRequirement(...)`
- or do a post-create lookup by work item id before calling comment sync

### 4. Customer defect inbound logic is not the same shape as current defect discussion sync

In `../bp-changes/SyncDefectFromAzureDevopsWorkItem.p1.txt`, the customer comment logic uses:

- `System.CommentCount`
- `System.History`
- qTest defect comments API

It does **not** use the same ADO comments API flow the current finished rule uses for `DefectDiscussionFieldID`.

So for a faithful merge in `SyncDefectFromAzureDevopsWorkItem.P1.js`, we should treat these as two parallel behaviors:

- keep the existing discussion-field snapshot logic
- add the customer's separate comment-create logic from `System.History`

### 5. Customer defect CID markers appear to be commented out

Both customer defect files appear to have the CID marker write commented out:

- `../bp-changes/SyncDefectFromAzureDevopsWorkItem.p1.txt`
- `../bp-changes/UpdateDefectInAzureDevops.P1.txt`

That means a literal merge would preserve weaker defect dedupe behavior than the requirement rules.

Implication:

- requirements use explicit stored CID markers
- defects appear to rely more on text matching and origin markers than on persisted CID markers

If we want to be faithful, we should assume this was part of what they handed over unless directed otherwise.

### 6. Customer outbound defect force-change step adds a hidden config dependency

In `../bp-changes/UpdateDefectInAzureDevops.P1.txt`, the outbound comment flow updates `DefectDiscussionFieldID` after posting comments to ADO.

In the current finished `UpdateDefectInAzureDevops.P1.js`, `validateRequiredConfiguration()` does not currently require `DefectDiscussionFieldID`.

So a faithful merge needs one of these two outcomes:

- make `DefectDiscussionFieldID` a required constant for this rule
- or conditionally skip only that force-change step when the constant is missing

### 7. Requirement/RICEFW comment logic adds hidden ADO REST dependencies

The current requirement-style rules primarily use the webhook payload plus qTest APIs.

The customer comment logic adds direct ADO REST dependencies:

- `constants.AZDO_TOKEN`
- `constants.AzDoProjectURL`

Those are not currently enforced by the requirement/RICEFW configuration validation paths, so the merged plan should explicitly account for them.

### 8. Customer outbound defect comment formatting has an extra fallback dependency

In `../bp-changes/UpdateDefectInAzureDevops.P1.txt`, the outbound defect comment formatter uses:

- qTest comment `created_by` / `created_date` when present
- `lastModifiedUser` / `lastModifiedDate` from the defect as fallback

The current finished `UpdateDefectInAzureDevops.P1.js` does not currently carry that fallback block.

So if we want a faithful merge of the customer's outbound comment formatting, that helper context needs to come over too.

## Recommended Implementation

### 1. Treat this as a direct behavior merge, not a redesign

Do not refactor the main field-sync code paths. Add the customer's comment helpers adjacent to the current API helpers in each affected rule and call them from the existing event flow.

Only change what is required to make their logic run inside the current rules:

- rename or relocate helper functions to fit the current file structure
- wire their comment logic into the current entry points
- keep their markers, sync direction, and force-change behavior where they already use it

This keeps ownership boundaries clean because:

- the current finished field mapping stays intact
- the customer comment behavior is being carried forward, not redefined by us
- later cleanup can be treated as a customer enhancement, not part of this merge

### 2. Requirements: add one-way `ADO -> qTest` comment sync

Affected files:

- `SyncRequirementFromAzureDevopsWorkItem.P1.js`
- `SyncRICEFWFeatureFromAzureDevops.js`

Port the older requirement/feature pattern as literally as practical, adapting only enough for the current guard rails and helper layout.

Add helpers per file:

- `getADOHeaders()`
- `getQtestComments(requirementId)`
- `createQtestComment(requirementId, content)`
- `extractAllComments(workItemId)`
- `isCommentAdded(event)`
- `syncCommentsIfNeeded(event, workItemId, qtestRequirementId)`

Recommended merge behavior:

- Detect comment-only updates via `System.CommentCount`
- If the event is comment-only, do not force the normal requirement property update path
- Find the existing qTest requirement and run comment sync, then exit
- If the event includes both field changes and a new comment, run the normal create/update path first, then run comment sync

Why this is the lowest-responsibility merge path:

- it avoids unnecessary requirement updates on pure comment events
- it preserves the current loop-prevention logic for field updates
- it keeps comment handling as an additive path instead of changing existing evaluation logic any more than necessary

Implementation note:

- `SyncRequirementFromAzureDevopsWorkItem.P1.js` already returns the created/updated requirement object, so the synced requirement id can be reused directly
- `SyncRICEFWFeatureFromAzureDevops.js` currently does not return the created/updated requirement object; either return it or do a post-create lookup by work item id

### 3. Defects inbound: merge their `ADO -> qTest` comment creation without removing current discussion sync

Affected file:

- `SyncDefectFromAzureDevopsWorkItem.P1.js`

Keep the current `DefectDiscussionFieldID` behavior for now, and merge in the customer's actual qTest comment creation alongside it.

Add a second step that creates qTest comments from `System.History` / `System.CommentCount` using the customer's existing pattern.

Recommended behavior:

- Detect comment adds from `System.CommentCount`
- Read the latest comment text from `System.History`
- Skip qTest-origin comments such as `[From qTest]`
- Fetch existing qTest defect comments
- Preserve the customer dedupe/origin behavior as supplied
- Post the comment to `/defects/{id}/comments`

This preserves the current finished discussion-field snapshot while also carrying forward the customer's comment behavior.

### 4. Defects outbound: add `qTest -> ADO` comment sync as a separate path

Affected file:

- `UpdateDefectInAzureDevops.P1.js`

Port the older defect outbound pattern into the current rule near the existing qTest defect read / ADO work item read section with as little behavioral change as possible.

Add helpers:

- `getQtestComments(defectId)`
- `getAdoComments(workItemId)`
- `formatComment(...)`
- `syncQtestCommentsToAdo(...)`

Recommended behavior:

- Fetch qTest comments and existing ADO comments
- Ignore qTest comments already marked `[From ADO]`
- Format outbound comments with a clear origin marker like `[From qTest]`
- Include a stable dedupe marker like `[CID:<qtestCommentId>]`
- Create missing ADO comments through the work item comments API

Important placement detail:

- Run comment sync even when `patchData.length === 0`
- Otherwise comment-only qTest changes will be dropped because the current rule exits early on a no-op field patch

### 5. Carry forward the customer's "force change" step if it is part of their outbound comment flow

The older outbound defect rule updates `DefectDiscussionFieldID` to a timestamp after posting comments to ADO.

Under a direct-merge approach, this should be carried forward if needed to preserve their behavior, even though it is not something we would normally introduce ourselves.

Reason:

- the customer already chose this behavior
- removing it would be us taking ownership of redesigning their logic
- if it is part of what makes their comment path function, it belongs in a faithful merge

The only reason to omit it would be if it is technically incompatible with the current rule structure.

## Proposed Change Sequence

1. Implement requirement comment helpers and comment-only event handling in `SyncRequirementFromAzureDevopsWorkItem.P1.js`
2. Mirror the same pattern in `SyncRICEFWFeatureFromAzureDevops.js`
3. Add ADO-to-qTest defect comment creation in `SyncDefectFromAzureDevopsWorkItem.P1.js`
4. Add qTest-to-ADO defect comment creation in `UpdateDefectInAzureDevops.P1.js`
5. Verify comment-only events before touching any optional cleanup around the existing defect discussion field

This order delivers the customer ask incrementally while minimizing changes to the current finished rules.

## Validation Plan

### Requirements

- Create a new ADO comment on a linked requirement
- Confirm a qTest requirement comment is created once
- Re-run the same event or update the work item again and confirm no duplicate qTest comment is created
- Confirm a comment-only event does not rewrite unrelated requirement fields

### RICEFW Features

- Repeat the same checks for a linked RICEFW feature requirement

### Defects ADO -> qTest

- Add a new ADO comment to a linked defect
- Confirm a qTest defect comment is created once
- Confirm the existing discussion field behavior still works as before
- Confirm qTest does not get duplicate comments on replay

### Defects qTest -> ADO

- Add a qTest defect comment to a linked defect
- Confirm an ADO comment is created once
- Confirm comment-only qTest changes still sync when no defect field patch is needed
- Confirm the return ADO event does not echo the same comment back into qTest

## Main Risks To Watch

- qTest comment payload shape may vary between `content`, `text`, `id`, and `comment_id`
- ADO comment-only updates may not include the same metadata as field updates, so comment helpers should be defensive
- Requirement and feature rules currently do not require ADO REST configuration for their base sync path, so comment sync should fail softly if ADO comment config is incomplete
- If `SyncUserRegex` is missing or too broad, loop protection will be weaker; the explicit `[From ADO]` / `[From qTest]` markers from the customer logic should remain in place as a second guard

## Recommendation

Implement comments as additive helper flows inside the current files, but keep the customer logic semantically intact rather than trying to clean it up. The only deliberate deviations should be the minimum compatibility fixes needed for the code to run correctly inside the newer rule structure.

# Office reliability and post-edit format assessment

## Intent and baseline
User approved optimization of the reviewed Office issues and requires intelligent format assessment after edits. Preserve live-document routing, approval controls, existing action names, and result envelope. No fuzzy render selection, automatic missing-style creation, unrelated UI redesign, deployment, or modification of pre-existing untracked atlas scripts.

TaskStartSnapshot: HEAD d4ba9bc965d4ab0a7aeae862f6666ba3bc2a10ab. Existing untracked atlas scripts/notes retained. .codegraph created for investigation.
Baseline read: README.md, package.json, skills/ms-office/SKILL.md, server/actions/word.js, server/taskpane.js, server/capabilities.json, native renderer and focused tests. No AGENTS or CONTEXT file found in bounded path scan.
Decision: code-change; documentation alone cannot prevent object-to-string writes or resolve actual document styles. ArchitectureReviewRequired: yes, limited to existing action/normalization/native projection/skill boundaries. TDD Route: off; meaningful regression tests required, no RED-first gate.
PatchShape: owner-level validation and readback repairs. Canonical owners: Word action module, shared alignment mapper, native renderer, existing Office skills and localized capability metadata. No new runtime owner. Preserve bounded native matching and before/after checks.

## Accepted repair shape
- Validate HTML before Office mutation; diagnose misplaced/unknown parameters.
- Resolve actual paragraph styles, support built-in aliases without inventing localized names, preserve explicit-only creation policy.
- Normalize alignment inputs and outputs, report requested versus actual style patch fields.
- Read table values structurally; clarify deletion constraints; diagnose unsupported actions.
- Improve native projection using evidenced structural differences, retain exact unique matching.
- Document input/output structures and UTF-8 single-serialization example.
- Require post-edit format assessment, scoped improvements only where justified by task and existing document style.

## Verification and stop
Run focused regression tests, repository lint and suite; review final diff. Native live Office behavior and skill model behavior need explicit evidence or must remain disclosed limitations. Complete when implementation and available checks pass and residual native/host gaps are reported.

## Work split / checkpoint
Lead: word actions, shared alignment, capability metadata/docs, integration and verification.
Subagent 224bb681-7f61-4c81-a98b-b8cbac18dd25: native renderer and native tests only.
Subagent a78f58ba-c71e-4844-b037-e91b68a9fb4b: two Office skills and focused skill scenarios only.
Active: implement owner repairs. Next: inspect localized action metadata and test harness. Drift decision: continue.

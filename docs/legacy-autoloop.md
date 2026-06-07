---
doc-schema-version: 1
title: "Legacy Autoloop Workflow"
summary: "Legacy automation files retained for reference but not promoted as the default workflow."
read_when:
  - Touching legacy autoloop files
  - Explaining why autoloop is not the primary workflow
  - Cleaning up old automation documentation
doc_type: "guide"
---

# Legacy Autoloop Workflow

The repository still contains autoloop configuration:

- `.autoloop-preset/`
- `.autoloop/`

This workflow is retained for now. It is not the primary development path for
the modernized project documentation, package support contract, or normal
agent-driven maintenance.

## Current Policy

- Do not delete the autoloop files until the project owner decides what to do
  with them.
- Do not promote autoloop as the default contributor workflow.
- Do not update autoloop behavior as part of unrelated source or docs work.
- If a task explicitly asks to use or modernize autoloop, inspect the preset
  files first and keep changes scoped.

## Related Files

- `.autoloop-preset/README.md`
- `.autoloop-preset/autoloops.toml`
- `.autoloop-preset/miniloops.toml`
- `.autoloop-preset/topology.toml`
- `.autoloop-preset/harness.md`
- `.autoloop-preset/roles/`

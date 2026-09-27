# Repository handoff workflow

Read `PROJECT_CONTEXT.md` before making implementation changes. It is the persistent
project handoff and source of truth for future sessions; verify relevant details
against current code and reconcile discrepancies.

For every implementation task, determine whether the result changes information
future sessions need. If so, update `PROJECT_CONTEXT.md` in the same change. Update
existing sections, remove obsolete claims, and preserve architectural decisions,
capabilities, UI conventions, persistence, testing expectations and known constraints.
Keep it concise; do not append a change log or temporary debugging details.

Before finishing, verify the context document accurately describes the final repository
state. Do not change application behavior merely to make it agree with documentation.

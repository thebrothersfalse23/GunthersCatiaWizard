
---

**CATIA Macro Best Practices (CATVBA):**
- Always execute runAllGuards() when the user selects 'Run' in the UI. After this initial check, only re-run guard routines if the user changes the modification scope.
- Use camelCase for all variable, function, and subroutine names for consistency and readability.
- Encapsulate logic in modules and avoid duplicating code; use wrapper functions for traversal and property access.
- Access CATIA object properties safely using helper functions (e.g., SafeSet, GetPropStr) to prevent runtime errors.
- Keep UI logic (forms) separate from core logic (modules).
- Document all public API functions and usage patterns in a dedicated documentation module (e.g., docs.bas or docs.frm).
- When making changes to one module, ensure related modules and documentation are updated to reflect the current state.
- Avoid hard-coding paths or dependencies; rely on CATIA’s object model and user context.
- Use Option Explicit in all modules to enforce variable declaration and reduce runtime errors.
- Prefer early binding for CATIA objects when possible, but ensure late binding compatibility for distribution.
- Use error handling (On Error GoTo) in all public routines and log errors to a dedicated error handler module.
- Structure modules by responsibility: traversal, property access, guards, helpers, and UI dispatch.
- Use constants for all string literals and magic numbers; define them in a dedicated constants module (e.g., constants.bas).
- Write unit-testable functions; avoid side effects in core logic modules.

**For UI Development:**
- Keep UI forms lightweight and focused on user interaction.
- Use data-binding techniques to synchronize UI elements with underlying data models.
- Implement input validation and error handling at the UI level to improve user experience.
- Provide clear feedback to users for long-running operations (e.g., progress indicators).
- Use descriptive control names (e.g., btnRun, txtScope) and group related controls logically.
- Separate UI event handlers from business logic; UI events should call module functions.
- Support keyboard navigation and accessibility where possible.

**For Documentation:**
- Maintain a single source of truth for API documentation in docs.bas or docs.frm.
- Document all public types, constants, and enumerations.
- Include usage examples for each public API in the documentation module.
- Keep documentation synchronized with code changes; update immediately after edits.
- Use clear, concise language and consistent formatting for all documentation comments.

**When making any edits:**
- Always review and update related modules to maintain consistency across the codebase.
- Ensure that any changes to logic, function signatures, or workflows are reflected in all affected modules.
- Update the documentation module (e.g., docs.bas or docs.frm) immediately to accurately describe the current state, usage, and signatures of public APIs.
- Do not include change log or date-specific comments in the code; documentation and comments should remain relevant and timeless.
- After edits, verify that the UI, core logic, and documentation remain in sync, and that no hard-coded dependencies or assumptions have been introduced.
- Run all available tests (manual or automated) after making changes to verify correctness.
- Ensure that all modules compile without errors or warnings before committing changes.
- Remove unused variables, functions, and modules to keep the codebase clean.

**Your Copilot Instructions (Summary):**
- Use camelCase for all names unless otherwise specified.
- Ensure all code is compatible with CATVBA.
- When a change is made to one module, update other modules as needed to keep the project consistent.
- Any time code is changed or added, update docs.frm with the new information.
- Automatically edit code in the files if instructed to do so.
- UI ≠ core. Forms gather options and dispatch; traversal and helpers live in modules.
- Always use Option Explicit and error handling in all modules.
- Maintain a clean, organized, and well-documented codebase at all times.

---

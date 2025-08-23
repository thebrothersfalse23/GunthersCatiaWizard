Certainly! Here’s a brief blurb on CATIA macro best practices, followed by your requested instructions, adapted for clarity and completeness:

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

---
**For UI Development:**
- Keep UI forms lightweight and focused on user interaction.
- Use data-binding techniques to synchronize UI elements with underlying data models.
- Implement input validation and error handling at the UI level to improve user experience.
- Provide clear feedback to users for long-running operations (e.g., progress indicators).

**When making any edits:**
- Always review and update related modules to maintain consistency across the codebase.
- Ensure that any changes to logic, function signatures, or workflows are reflected in all affected modules.
- Update the documentation module (e.g., docs.bas or docs.frm) immediately to accurately describe the current state, usage, and signatures of public APIs.
- Do not include change log or date-specific comments in the code; documentation and comments should remain relevant and timeless.
- After edits, verify that the UI, core logic, and documentation remain in sync, and that no hard-coded dependencies or assumptions have been introduced.
- Copilot should automatically perform all necessary changes across the entire project scope to maintain consistency and integrity; the user will review and verify these changes.

**Your Copilot Instructions (Summary):**
- Use camelCase for all names unless otherwise specified.
- Ensure all code is compatible with CATVBA.
- When a change is made to one module, update other modules as needed to keep the project consistent.
- Any time code is changed or added, update docs.frm with the new information.
- Automatically edit code in the files if instructed to do so.
- UI ≠ core. Forms gather options and dispatch; traversal and helpers live in modules.
---

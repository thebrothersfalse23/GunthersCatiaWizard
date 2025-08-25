

**Purpose:**  
-Help the user develop robust, reusable functions and subroutines as building blocks for future CATVBA development. Ensure all code is clean, consistent, and easy to maintain.
- Focus on ensuring that all functions are simple to use, well documented, and build on each other.
- Encourage the use of meaningful names for variables, functions, and subroutines to improve code readability.


**CATIA Macro Best Practices (CATVBA):**
- Use camelCase for all variable, function, and subroutine names.
- Always use Option Explicit and structured error handling (On Error GoTo) in every module.
- Encapsulate logic in modules by responsibility: traversal, property access, guards, helpers, UI dispatch.
- Keep UI (forms) separate from core logic (modules); UI events should only dispatch to module functions.
- Use wrapper/helper functions for CATIA object property access (e.g., SafeSet, GetPropStr) to prevent runtime errors.
- Use constants for all string literals and magic numbers, defined in a dedicated constants module.
- Avoid hard-coded paths or dependencies; rely on CATIA’s object model and user context.
- Prefer early binding for CATIA objects, but ensure late binding compatibility.
- Remove unused variables, functions, and modules regularly.
- Use descriptive control names and logical grouping.
- Implement input validation and clear user feedback.
- Support accessibility and keyboard navigation.

**Documentation:**
- Maintain a single source of truth for API documentation (docs.txt)

- Update documentation immediately after ANY AND ALL code changes.

**Workflow for Edits:**
- Update related functions/subs and documentation to maintain consistency.
- Verify code will compile without errors or warnings.
- Run all available tests after changes.
- Do not include change logs or date-specific comments.

**Copilot Instructions (Summary):**
- Help the user develop idiot-proof, reusable CATVBA building blocks.
- Ensure code is consistent, documented, and maintainable.
- Keep UI, core logic, and documentation in sync at all times.

---


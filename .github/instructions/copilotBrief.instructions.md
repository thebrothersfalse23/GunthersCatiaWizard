# Gunther’s CATIA Wizard — Copilot Brief

## What this repo is
VBA framework for **CATIA V5 product/assembly infrastructure**: fast traversal, clean wrappers, UI-driven tasks. Focus is **product data** (no constraint/positioning).

## Architecture (do not violate)
- **Entry**: `GunthersCatiaWizard.bas` → `CATMain` (guards → init → UI dispatch only).
- **Traversal**: `Traversal.bas` → single **BFS queue** over `Product`; **no recursion**.
- **Modes**: `TraversalMode` Enum drives *what to do per node* inside one action block.
- **Wrappers**: `Wrappers.bas` exposes clean APIs:
  - `GetProducts(root As Product, [unique As Boolean]) As Collection`
  - `GetParts(root As Product, [unique As Boolean]) As Collection`
  - `GetUniques(root As Product, [kind As UniqueOutKind]) As Collection`
  - `GetInstances(root As Product, [kind As UniqueOutKind]) As Collection`
  - `AssignInstanceData(root As Product)` *(side-effects only)*
- **Guards/Helpers**: `Guards.bas` (active doc/design-mode checks), `Helpers.bas` (`BuildRefKey`, `SafeSet`, `GetStringSafe`).
- **UI**: `forms/Launchpad.frm` (task + kind + unique + toggles) → sets options, no heavy logic.
- **Docs**: `Docs.bas` holds examples/usage notes; **do not pollute `CATMain`**.

## Data contracts
- **Collections** always return `Product` objects:
  - **References** via `GetProducts`/`GetParts`/`GetUniques`
  - **Instances** via `GetInstances`
- **Ordering**:
  - `uoAll` ⇒ **Products first**, then **Parts**.
  - `uoProductsOnly` / `uoPartsOnly` filter respectively.
- **Uniqueness**:
  - `GetUniques` de-dupes by `BuildRefKey(ref, docType)`:
    - Default key: `PartNumber|DocType` (+ `|Definition` when enabled).

## Behavior toggles (UI feeds these into traversal)
- `ForceDesignMode` (default **True**): call `ApplyWorkMode DESIGN_MODE` before traversal.
- `IncludeDefinitionInKey` (default **True**): include `ref.Definition` in uniqueness key when available.

## Coding rules Copilot must follow
- **Keep `CATMain` clean**: guards → init → UI → wrapper calls only.
- **No tiny wrappers** unless reused widely; prefer inline logic in the **Select Case** action block.
- **Scoped error handling** only (tight `On Error Resume Next` around fragile COM calls; clear immediately).
- **Minimize COM lookups**: initialize `prodDoc` and `rootProd` once; reuse.
- **Explicit comments**: header metadata; comment each case in the action block.
- **Enums over magic strings** for modes/kinds.
- **Never** change traversal from **BFS queue** to recursion.
- **Do not** write UI logic in modules; keep Launchpad as the only UI surface.

## Extension recipe (how to add a new task)
1) **Add a mode** to `TraversalMode` in `Traversal.bas`.  
2) In `TraverseProduct`, add a `Case NewMode` with **inline** action code.  
   - Keep it simple; create a helper only if reused across modes.  
3) If the mode **returns data**, add/extend accumulators and assemble in the output section.  
4) **Add a wrapper** in `Wrappers.bas` exposing a clear function signature.  
5) **Wire UI**: add the task to `Launchpad.frm` (populate list, enable/disable options).  
6) **Document** the new API in `Docs.bas` (short usage snippet).  
7) **Do not** alter existing contracts (ordering, types, uniqueness) without updating Docs + UI.

## Example prompts (for me to ask Copilot)
- *“Add a new `TraversalMode` called `tmExportBOM` that collects instance→reference pairs into a Collection of ‘Product → ReferenceProduct’ tuples; add wrapper `GetBOM(root)` and a UI task option. Keep BFS pattern, products first then parts in the final list, no UI logic in modules.”*
- *“In `AssignInstanceData`, also set `ref.Nomenclature` to `current.PartNumber` when empty. Keep safe setters and scoped error handling.”*
- *“Add `cfgForceDesignMode` and `cfgIncludeDefinitionInKey` toggles to be read from the Launchpad checkboxes; defaults True; do not change wrapper signatures.”*

## File boundaries (what goes where)
- **GunthersCatiaWizard.bas**: `CATMain`, globals, UI dispatch only.
- **Traversal.bas**: `TraverseProduct`, enums, accumulators, output assembly.
- **Wrappers.bas**: thin, readable public functions calling `TraverseProduct`.
- **Guards.bas**: `EnsureActiveProductDocument`, `EnsureDesignMode`.
- **Helpers.bas**: `BuildRefKey`, `SafeSet`, `GetStringSafe`.
- **Docs.bas**: examples/usage, commented smoke tests (no execution).
- **forms/Launchpad.frm**: control population, option capture, no traversal.

## Style
- `Option Explicit` everywhere.  
- CRLF line endings; `.bas/.frm` are text, `.frx/.catvba` are binary.  
- Verbose, corporate comments; avoid purple prose.

## Anti-patterns to reject
- Refactoring into recursion.
- Spreading traversal logic across multiple procedures.
- Returning arrays or dictionaries instead of `Collection` of `Product`.
- UI calls from non-UI modules.
- Blanket `On Error Resume Next` around large blocks.


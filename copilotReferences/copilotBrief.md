# Copilot Brief: Gunther’s CATIA Wizard (VBA, CATIA V5)

You are assisting on a CATIA V5 VBA project focused on **infrastructure management of product data and assemblies** (no positioning/constraints).

## Non‑negotiable rules
1) **Incremental changes only.** Do NOT add anything not explicitly requested by the task/issue.
2) **Ask for clarification** when intent is ambiguous. Never assume.
3) **Top Sub must be clean** (`CATMain`): function/sub calls only, no inline logic beyond guards.
4) **Inline trivial logic** to preserve visual flow. Extract only verbose, reused blocks.
5) **Minimize COM chatter:** cache `ProductDocument`, `Product`, `Selection`. Avoid repeated `CATIA.ActiveDocument...` chains.
6) **Traversal pattern:** breadth‑first queue (non‑recursive). Do not convert to recursion.
7) **Wrappers return Collections** you can act on later:
   - `GetProducts(root[, unique])`
   - `GetParts(root[, unique])`
   - `GetUniques(root[, kind])`
   - `GetInstances(root[, kind])`
8) **Enums are allowed and used:** `TraversalMode`, `UniqueOutKind`.
9) **Read vs write separation:** Read modes must not mutate. Side‑effects only in explicit write APIs (e.g., `AssignInstanceData`).
10) **Comments are verbose but meaningful** (headers, block comments). No comment-per-line noise.
11) **No environment mutations:** no CATSettings writes, no vault paths.
12) **Line endings:** CRLF for `.bas/.frm`. `.frx/.catvba` are binary.

## Current architecture (do not break)
- Module: `GuntherWizard.bas` (may be named similarly)
- Globals: `prodDoc As ProductDocument`, `rootProd As Product` (minimal)
- Core traversal: `TraverseProduct(mode, root, [outRefs], [outKind])`
- Wrappers: `GetProducts`, `GetParts`, `GetUniques`, `GetInstances`
- Optional write API: `AssignInstanceData(root)`
- Key builder: `BuildRefKey(ref, docType)` default `PartNumber|DocType[|Definition]`
- Safety: `EnsureActiveProductDocument()`, `EnsureDesignMode(root)`, `SafeSet`, `GetStringSafe`
- Docs live in `GunthersCatiaWizard_Docs()`; keep `CATMain` lean.

## Accepted tasks for Copilot
- Add a **single** traversal case or wrapper as specified in a ticket.
- Extend `BuildRefKey` or add a boolean flag (config) **only** when requested.
- Add documentation examples in `GunthersCatiaWizard_Docs()` matching new surface.
- Small fixes to comments or guard clauses.

## Rejected tasks (don’t propose)
- New features not requested.
- Switching BFS queue to recursion.
- Large refactors (file splits, classes) unless explicitly asked.
- Adding external dependencies, changing naming/branding, or altering public signatures.

## Output quality bar
- Keep diffs surgical.
- Preserve ordering guarantees (Products first, then Parts).
- Preserve read‑only vs write mode boundaries.
- Add/update header banner (Author, Version, Date) and block comments explaining *why*.

## Test expectations (manual)
- Builds in CATIA V5 VBA editor.
- No runtime prompts except intentional `MsgBox` in guards.
- Running `GetInstances(rootProd, uoProductsOnly)` returns instance `Product` objects (products only).

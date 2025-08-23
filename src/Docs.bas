'===============================================================
' MODULE: Docs.bas
' PURPOSE: Documentation and usage examples for Gunther's Catia Wizard.
'          Keeps CATMain clean and provides reference for API usage.
'===============================================================
' -------------------------------------------------------------------------
' DOCUMENTATION FORMAT (for each public function/wrapper)
' -------------------------------------------------------------------------
' FunctionName([arg1 As Type][, arg2 As Type ...]) [As ReturnType]
'   Description:
'     [Short summary of what the function does and any important notes.]
'   Acceptable args:
'     arg1 [Type] - [Description of argument and valid values]
'     arg2 [Type, Optional] - [Description and default if any]
'     ...
'   Usage:
'     [Short code snippet showing typical usage]
'
' (Repeat for each function/wrapper)
'
' Enumerations:
'   EnumName: value1, value2, ...
'
' -------------------------------------------------------------------------
' UI USAGE (if applicable)
'   - [Describe how the UI interacts with this function/module]
'
' -------------------------------------------------------------------------
' UNDER THE HOOD (optional)
'   - [Describe internal mechanics, contracts, or design notes]
'
' -------------------------------------------------------------------------
' NOTES (optional)
'   - [Any additional notes, contracts, or conventions]
' -------------------------------------------------------------------------
Public Sub GunthersCatiaWizard_Docs()
    ' -------------------------------------------------------------------------
    ' PUBLIC API QUICK REFERENCE
    '   - GetProducts(rootProd As Product, [unique As Boolean = False]) As Collection
    '   - GetParts(rootProd As Product, [unique As Boolean = False]) As Collection
    '   - GetUniques(rootProd As Product, [kind As UniqueOutKind = uoAll]) As Collection
    '   - GetInstances(rootProd As Product, [kind As UniqueOutKind = uoAll]) As Collection
    '   - SafeSet(obj As Object, propName As String, value As String)
    '   - GetPropStr(obj As Object, propName As String) As String
    '   - BuildRefKey(ref As Product, docType As String) As String
    '   - EnsureActiveProductDocument() As Boolean
    '   - EnsureDesignMode(root As Product)
    '   - TraverseProduct(mode As TraversalMode, root As Product, [ByRef outRefs As Collection], [outKind As UniqueOutKind = uoAll])
    '
    ' ENUMS:
    '   TraversalMode: tmGetUniques, tmGetParts, tmAssignInstanceData, tmCollectRefsAll, tmGetInstances
    '   UniqueOutKind: uoAll, uoProductsOnly, uoPartsOnly
    ' -------------------------------------------------------------------------
    '' WRAPPER FUNCTION USAGE (all return Collection unless noted)
    ''
    '' GetProducts(rootProd As Product, [unique As Boolean = False])  → reference Products only
    ''   Description:
    ''     Returns a Collection of reference Products (Products only) in the assembly.
    ''     Optionally deduplicates by reference.
    ''   Acceptable args:
    ''     rootProd [Product] - root product to traverse
    ''     unique   [Boolean, Optional] - True for unique refs, False for all (default: False)
    ''   Usage:
    ''     Dim prodsAll As Collection
    ''     Set prodsAll = GetProducts(rootProd, False)
    ''     Dim prodsUniq As Collection
    ''     Set prodsUniq = GetProducts(rootProd, True)
    ''
    '' GetParts(rootProd As Product, [unique As Boolean = False])      → reference Parts only
    ''   Description:
    ''     Returns a Collection of reference Parts (Parts only) in the assembly.
    ''     Optionally deduplicates by reference.
    ''   Acceptable args:
    ''     rootProd [Product] - root product to traverse
    ''     unique   [Boolean, Optional] - True for unique refs, False for all (default: False)
    ''   Usage:
    ''     Dim partsAll As Collection
    ''     Set partsAll = GetParts(rootProd, False)
    ''     Dim partsUniq As Collection
    ''     Set partsUniq = GetParts(rootProd, True)
    ''
    '' GetUniques(rootProd As Product, [kind As UniqueOutKind = uoAll]) → unique refs (ordered)
    ''   Description:
    ''     Returns a Collection of unique reference Products and/or Parts, ordered with Products first.
    ''     Filtering by kind is available.
    ''   Acceptable args:
    ''     rootProd [Product] - root product to traverse
    ''     kind     [UniqueOutKind, Optional] - uoAll, uoProductsOnly, uoPartsOnly (default: uoAll)
    ''   Usage:
    ''     Dim uniqAll As Collection
    ''     Set uniqAll = GetUniques(rootProd, uoAll)
    ''     Dim uniqProds As Collection
    ''     Set uniqProds = GetUniques(rootProd, uoProductsOnly)
    ''     Dim uniqParts As Collection
    ''     Set uniqParts = GetUniques(rootProd, uoPartsOnly)
    ''
    '' GetInstances(rootProd As Product, [kind As UniqueOutKind = uoAll]) → instance Products
    ''   Description:
    ''     Returns a Collection of instance Products (not references) in the assembly.
    ''     Filtering by kind is available.
    ''   Acceptable args:
    ''     rootProd [Product] - root product to traverse
    ''     kind     [UniqueOutKind, Optional] - uoAll, uoProductsOnly, uoPartsOnly (default: uoAll)
    ''   Usage:
    ''     Dim instAll As Collection
    ''     Set instAll = GetInstances(rootProd, uoAll)
    ''     Dim instProds As Collection
    ''     Set instProds = GetInstances(rootProd, uoProductsOnly)
    ''     Dim instParts As Collection
    ''     Set instParts = GetInstances(rootProd, uoPartsOnly)
    ''
    '' SafeSet(obj As Object, propName As String, value As String)
    ''   Description:
    ''     Safely sets a property (e.g., "Description", "Name", etc.) on a CATIA object if it exists.
    ''     Ignores errors if the property is not present.
    ''   Acceptable args:
    ''     obj      [Object] - object to set property on
    ''     propName [String] - property name ("Nomenclature", "Name", "Description", "PartNumber", "Definition", "Revision", "ReferenceProduct")
    ''     value    [String] - value to assign
    ''   Usage:
    ''     SafeSet currentProduct, "Description", "MADE BY AMCO"
    ''
    '' GetPropStr(obj As Object, propName As String) As String
    ''   Description:
    ''     Safely retrieves a property value as a string from a CATIA object.
    ''     Returns "" if the property does not exist or on error.
    ''   Acceptable args:
    ''     obj      [Object] - object to get property from
    ''     propName [String] - property name ("Nomenclature", "Name", "Description", "PartNumber", "Definition", "Revision", "ReferenceProduct")
    ''   Usage:
    ''     Dim desc As String
    ''     desc = GetPropStr(currentProduct, "Description")
    ''
    '' BuildRefKey(ref As Product, docType As String) As String
    ''   Description:
    ''     Builds a stable, human-readable key for a reference Product.
    ''     Default: "PartNumber|DocType"; if Definition exists → "PartNumber|DocType|Definition"
    ''   Acceptable args:
    ''     ref     [Product] - reference Product object
    ''     docType [String]  - "ProductDocument" or "PartDocument"
    ''   Usage:
    ''     Dim key As String
    ''     key = BuildRefKey(refProduct, "ProductDocument")
    ''
    '' EnsureActiveProductDocument() As Boolean
    ''   Description:
    ''     Ensures a ProductDocument is open and active in CATIA.
    ''     Sets globals prodDoc/rootProd if successful.
    ''     Returns True if valid, False otherwise.
    ''   Acceptable args: none
    ''   Usage:
    ''     If Not EnsureActiveProductDocument() Then Exit Sub
    ''
    '' EnsureDesignMode(root As Product)
    ''   Description:
    ''     Applies Design Mode to a Product for consistent traversal.
    ''     No effect if already in Design Mode.
    ''   Acceptable args:
    ''     root [Product] - product to apply Design Mode to
    ''   Usage:
    ''     EnsureDesignMode rootProd
    ''
    '' TraverseProduct(mode As TraversalMode, root As Product, [ByRef outRefs As Collection], [outKind As UniqueOutKind = uoAll])
    ''   Description:
    ''     Core traversal logic for all wrappers.
    ''     Iterative BFS queue, not called directly by most users.
    ''   Acceptable args:
    ''     mode    [TraversalMode] - tmGetUniques, tmGetParts, tmAssignInstanceData, tmCollectRefsAll, tmGetInstances
    ''     root    [Product]       - root product to traverse
    ''     outRefs [Collection, Optional, ByRef] - receives output
    ''     outKind [UniqueOutKind, Optional]     - uoAll, uoProductsOnly, uoPartsOnly (default: uoAll)
    ''   Usage (advanced):
    ''     Dim outRefs As Collection
    ''     TraverseProduct tmGetUniques, rootProd, outRefs, uoAll
    ''
    '' Enumerations:
    ''   TraversalMode: tmGetUniques, tmGetParts, tmAssignInstanceData, tmCollectRefsAll, tmGetInstances
    ''   UniqueOutKind: uoAll, uoProductsOnly, uoPartsOnly
    ''
    '' -------------------------------------------------------------------------
    '' UI USAGE (Launchpad)
    ''
    '' - All UI is handled in a single form: Launchpad.frm.
    '' - Launchpad uses an SSTab control to switch between:
    ''     • Home tab (instructions, Run/Cancel)
    ''     • Single Tool Design tab
    ''     • Sequential Tool Design tab
    '' - Only the Home tab or the tool design tabs are visible at a time.
    ''     • After clicking "Run", only tool design tabs are shown.
    ''     • "Back" returns to Home and hides tool design tabs.
    '' - All navigation and tab visibility is managed in Launchpad.frm code.
    '' - No other forms are used; generateTD.frm is removed.
    ''
    '' -------------------------------------------------------------------------
    '' UNDER THE HOOD
    ''   - tmGetUniques      → BuildRefKey() + Dictionary to return unique refs,
    ''                         ordered Products then Parts (or filtered by kind).
    ''   - tmCollectRefsAll  → returns all refs (no dedupe), ordered Products then Parts.
    ''   - tmGetInstances    → returns instance Products, bucketed by ref doc type.
    ''   - SafeSet           → safely sets a property if it exists on the object.
    ''   - GetPropStr        → safely gets a property value as string.
    ''   - BuildRefKey       → builds a stable key for deduplication.
    ''   - EnsureDesignMode  → applies Design Mode for traversal.
    ''   - EnsureActiveProductDocument → ensures a ProductDocument is open and sets globals.
    ''   - TraverseProduct   → core traversal, not called directly by users.
    ''
    '' NOTES
    ''   - All wrappers pass ByRef collections through TraverseProduct and return them.
    ''   - No UI side effects here; add exporters or viewers only when you ask.
    ''   - BuildRefKey default is "PartNumber|DocType" and appends "|Definition"
    ''     if available to reduce collisions across libraries.
    ''   - Design Mode is applied up-front for consistent traversal depth.
    ''   - UniqueOutKind: uoAll, uoProductsOnly, uoPartsOnly (enum for filtering).
    ''   - All collections are 1-based (VBA default).
    '' -------------------------------------------------------------------------
End Sub

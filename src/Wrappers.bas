'===============================================================
' MODULE: wrappers.bas
' PURPOSE: Provides wrapper functions for traversing CATIA Product structures.
'          These functions return collections of reference or instance Products/Parts,
'          with options for deduplication and filtering.
'
' FUNCTIONS:
'   - getProducts:   Returns reference Products (Products only), deduped or all.
'   - getParts:      Returns reference Parts (Parts only), deduped or all.
'   - getUniques:    Returns unique references (Products and/or Parts).
'   - getInstances:  Returns instance Products (not references), filtered by kind.
'
' All functions accept a root Product and optional parameters for deduplication or filtering.
'===============================================================

' getProducts – returns reference Products (Products only)
' Parameters:
'   root   [Product]   - Root product to traverse.
'   unique [Boolean]   - If True, returns unique references; if False, returns all.
' Returns:
'   [Collection] of reference Products (Products only).
Public Function getProducts(ByVal root As Product, Optional ByVal unique As Boolean = False) As Collection
    Dim outRefs As Collection
    If unique Then
        traverseProduct tmGetUniques, root, outRefs, uoProductsOnly
    Else
        traverseProduct tmCollectRefsAll, root, outRefs, uoProductsOnly
    End If
    Set getProducts = outRefs
End Function

' getParts – returns reference Products (Parts only)
' Parameters:
'   root   [Product]   - Root product to traverse.
'   unique [Boolean]   - If True, returns unique references; if False, returns all.
' Returns:
'   [Collection] of reference Products (Parts only).
Public Function getParts(ByVal root As Product, Optional ByVal unique As Boolean = False) As Collection
    Dim outRefs As Collection
    If unique Then
        traverseProduct tmGetUniques, root, outRefs, uoPartsOnly
    Else
        traverseProduct tmCollectRefsAll, root, outRefs, uoPartsOnly
    End If
    Set getParts = outRefs
End Function

' getUniques – returns unique reference Products (ordered)
' Parameters:
'   root [Product]           - Root product to traverse.
'   kind [uniqueOutKind]     - Filter: uoAll, uoProductsOnly, or uoPartsOnly.
' Returns:
'   [Collection] of unique references (Products and/or Parts).
Public Function getUniques(ByVal root As Product, Optional ByVal kind As uniqueOutKind = uoAll) As Collection
    Dim outRefs As Collection
    traverseProduct tmGetUniques, root, outRefs, kind
    Set getUniques = outRefs
End Function

' getInstances – returns instance Products (not references)
' Parameters:
'   root [Product]           - Root product to traverse.
'   kind [uniqueOutKind]     - Filter: uoAll, uoProductsOnly, or uoPartsOnly.
' Returns:
'   [Collection] of instance Products.
Public Function getInstances(ByVal root As Product, Optional ByVal kind As uniqueOutKind = uoAll) As Collection
    Dim outInst As Collection
    traverseProduct tmGetInstances, root, outInst, kind
    Set getInstances = outInst
End Function



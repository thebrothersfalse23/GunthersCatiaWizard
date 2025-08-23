'===============================================================
' Macro:        Gunther's Catia Wizard
' Author:       [Your Name]
' Version:      v1.0 – 2025-08-22
' CATIA:        V5 (tested on V5-6R2020+; expected OK from V5R2016)
'
' DOCS INDEX
'   1) Quick Start ........................... See: GunthersCatiaWizard_Docs()
'   2) API surface (wrappers) ................ GetProducts/GetParts/GetUniques/GetInstances
'   3) Modes & contracts ..................... TraversalMode, UniqueOutKind, BuildRefKey()
'   4) Behavior guarantees ................... Ordering, uniqueness, instance/refs split
'   5) Optional write API .................... AssignInstanceData()
'   6) Full narrative & examples ............. Bottom section: "Documentation / Usage Examples"
'
' NOTE: This file is the source of truth for user docs. Do not strip comments.
'===============================================================

Option Explicit

'===============================================================
' Global Variables (kept minimal)
'===============================================================
'--- [SUGGESTED MODULE: Globals.bas] ---
Public prodDoc As ProductDocument     ' Active ProductDocument
Public rootProd As Product            ' Root Product of the assembly

'===============================================================
' Enumerations
'===============================================================
'--- [SUGGESTED MODULE: Enums.bas] ---
Public Enum TraversalMode
    tmGetUniques = 1            ' unique reference Products (deduped)
    tmGetParts = 2              ' reserved placeholder (not used by wrappers)
    tmAssignInstanceData = 3    ' explicit write API (separate from read traversals)
    tmCollectRefsAll = 4        ' all reference Products (not deduped)
    tmGetInstances = 5          ' instance Products by kind
End Enum

' Kind selector for refs/instances
Public Enum UniqueOutKind
    uoAll = 0            ' Products first, then Parts
    uoProductsOnly = 1   ' Products only
    uoPartsOnly = 2      ' Parts only
End Enum

'===============================================================
' Entry Point (guards → init → sample call)
'===============================================================
'--- [SUGGESTED MODULE: Main.bas] ---
Sub CATMain()

    If Not EnsureActiveProductDocument() Then Exit Sub

    ' Example: count unique reference Products+Parts
    Dim uniqAll As Collection
    Set uniqAll = GetUniques(rootProd, uoAll)

    MsgBox "Unique references found: " & CStr(uniqAll.Count), vbInformation, "Gunther's CATIA Wizard"

    ' Keep Main clean. See GunthersCatiaWizard_Docs for full examples & usage.

End Sub

'===============================================================
' WRAPPER FUNCTIONS (clean call surface for Main/other code)
'===============================================================
'--- [SUGGESTED MODULE: Wrappers.bas] ---
' GetProducts – returns reference Products (Products only)
' unique:=True  → deduped via tmGetUniques
' unique:=False → all refs via tmCollectRefsAll
Public Function GetProducts(ByVal root As Product, Optional ByVal unique As Boolean = False) As Collection
    Dim outRefs As Collection
    If unique Then
        TraverseProduct tmGetUniques, root, outRefs, uoProductsOnly
    Else
        TraverseProduct tmCollectRefsAll, root, outRefs, uoProductsOnly
    End If
    Set GetProducts = outRefs
End Function

' GetParts – returns reference Products (Parts only)
' unique:=True  → deduped via tmGetUniques
' unique:=False → all refs via tmCollectRefsAll
Public Function GetParts(ByVal root As Product, Optional ByVal unique As Boolean = False) As Collection
    Dim outRefs As Collection
    If unique Then
        TraverseProduct tmGetUniques, root, outRefs, uoPartsOnly
    Else
        TraverseProduct tmCollectRefsAll, root, outRefs, uoPartsOnly
    End If
    Set GetParts = outRefs
End Function

' GetUniques – returns unique reference Products (ordered)
' kind: uoAll (Products→Parts), uoProductsOnly, or uoPartsOnly
Public Function GetUniques(ByVal root As Product, Optional ByVal kind As UniqueOutKind = uoAll) As Collection
    Dim outRefs As Collection
    TraverseProduct tmGetUniques, root, outRefs, kind
    Set GetUniques = outRefs
End Function

' GetInstances – returns instance Products (not references)
' kind: uoAll (Products→Parts), uoProductsOnly, or uoPartsOnly
Public Function GetInstances(ByVal root As Product, Optional ByVal kind As UniqueOutKind = uoAll) As Collection
    Dim outInst As Collection
    TraverseProduct tmGetInstances, root, outInst, kind
    Set GetInstances = outInst
End Function

'===============================================================
' Optional write API (kept separate from read traversals)
'===============================================================
'--- [SUGGESTED MODULE: WriteAPI.bas] ---
Public Sub AssignInstanceData(ByVal root As Product)
    Dim unused As Collection
    TraverseProduct tmAssignInstanceData, root, unused, uoAll
End Sub


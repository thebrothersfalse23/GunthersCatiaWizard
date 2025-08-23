===============================================================
MODULE: GunthersCatiaWizard.bas
PURPOSE: Main entry point and orchestrator for Gunther's Catia Wizard macro.
         Initializes globals, exposes entry point, and demonstrates usage.
         See Docs.bas for API documentation and usage examples.
===============================================================

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
Sub CATMain()

    If Not EnsureActiveProductDocument() Then Exit Sub

    ' Example: count unique reference Products+Parts
    Dim uniqAll As Collection
    Set uniqAll = GetUniques(rootProd, uoAll)

    MsgBox "Unique references found: " & CStr(uniqAll.Count), vbInformation, "Gunther's CATIA Wizard"

    ' Keep Main clean. See GunthersCatiaWizard_Docs for full examples & usage.

End Sub


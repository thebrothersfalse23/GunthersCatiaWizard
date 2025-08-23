'===============================================================
' MODULE: Enums.bas
' PURPOSE: Enumerations for traversal and output kind.
'===============================================================

Option Explicit

'===============================================================
' Enumerations for traversal and output kind
'===============================================================

Public Enum TraversalMode
    tmGetUniques = 1            ' unique reference Products (deduped)
    tmGetParts = 2              ' reserved placeholder (not used by wrappers)
    tmAssignInstanceData = 3    ' explicit write API (separate from read traversals)
    tmCollectRefsAll = 4        ' all reference Products (not deduped)
    tmGetInstances = 5          ' instance Products by kind
End Enum

Public Enum UniqueOutKind
    uoAll = 0            ' Products first, then Parts
    uoProductsOnly = 1   ' Products only
    uoPartsOnly = 2      ' Parts only
End Enum
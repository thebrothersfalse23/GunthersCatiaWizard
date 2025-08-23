'===============================================================
' MODULE: Guards.bas
' PURPOSE: Provides document and work-mode guard routines for CATIA macros.
'          Ensures correct document type and design mode for safe traversal.
'===============================================================
Public Sub EnsureDesignMode(ByVal root As Product)
    On Error Resume Next
    root.ApplyWorkMode DESIGN_MODE
    Err.Clear
    On Error GoTo 0
End Sub

Public Function EnsureActiveProductDocument() As Boolean
    EnsureActiveProductDocument = False

    If CATIA.Documents.Count = 0 Then
        MsgBox "No valid CATProduct document is open. Please open a Product and try again.", vbExclamation, "Gunther's Catia Wizard"
        Exit Function
    End If

    If TypeName(CATIA.ActiveDocument) <> "ProductDocument" Then
        MsgBox "No valid CATProduct document is open. Please open a Product and try again.", vbExclamation, "Gunther's Catia Wizard"
        Exit Function
    End If

    Set prodDoc = CATIA.ActiveDocument
    Set rootProd = prodDoc.Product
    EnsureActiveProductDocument = True
End Function
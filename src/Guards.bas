'===============================================================
' MODULE: guards.bas
' PURPOSE: Provides document and work-mode guard routines for CATIA macros.
'          Ensures correct document type and design mode for safe traversal.
'===============================================================
Public Sub ensureDesignMode(ByVal root As Product)
    On Error Resume Next
    root.ApplyWorkMode DESIGN_MODE
    Err.Clear
    On Error GoTo 0
End Sub

Public Function ensureActiveProductDocument() As Boolean
    ensureActiveProductDocument = False

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
    ensureActiveProductDocument = True
End Function
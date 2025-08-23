'===============================================================
' Work-mode & document guards
'===============================================================
Private Sub EnsureDesignMode(ByVal root As Product)
    On Error Resume Next
    root.ApplyWorkMode DESIGN_MODE
    Err.Clear
    On Error GoTo 0
End Sub

Private Function EnsureActiveProductDocument() As Boolean
    EnsureActiveProductDocument = False

    If CATIA.Documents.Count = 0 Then
        MsgBox "No document open.", vbExclamation, "Gunther's CatIA Wizard"
        Exit Function
    End If

    If TypeName(CATIA.ActiveDocument) <> "ProductDocument" Then
        MsgBox "Active document is not a CATProduct.", vbExclamation, "Gunther's CatIA Wizard"
        Exit Function
    End If

    Set prodDoc = CATIA.ActiveDocument
    Set rootProd = prodDoc.Product
    EnsureActiveProductDocument = True
End Function
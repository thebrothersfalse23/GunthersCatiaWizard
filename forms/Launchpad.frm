
'===============================================================
' FORM: Launchpad.frm
' PURPOSE: Main GUI for Gunther's Catia Wizard. Provides buttons
'          to list unique parts/products and rename instances.
'===============================================================

Private Sub cmdListParts_Click()
    If Not EnsureActiveProductDocument() Then
        lblStatus.Caption = "No active CATProduct document."
        Exit Sub
    End If
    Dim parts As Collection
    Set parts = GetParts(rootProd, True)
    lstOutput.Clear
    Dim i As Long
    For i = 1 To parts.Count
        lstOutput.AddItem parts(i).PartNumber
    Next i
    lblStatus.Caption = "Listed unique parts."
End Sub

Private Sub cmdListProducts_Click()
    If Not EnsureActiveProductDocument() Then
        lblStatus.Caption = "No active CATProduct document."
        Exit Sub
    End If
    Dim prods As Collection
    Set prods = GetProducts(rootProd, True)
    lstOutput.Clear
    Dim i As Long
    For i = 1 To prods.Count
        lstOutput.AddItem prods(i).PartNumber
    Next i
    lblStatus.Caption = "Listed unique products."
End Sub

Private Sub cmdRenameInstances_Click()
    If Not EnsureActiveProductDocument() Then
        lblStatus.Caption = "No active CATProduct document."
        Exit Sub
    End If
    Dim prodsToRename As Collection
    Set prodsToRename = GetInstances(rootProd, uoProductsOnly)
    Dim i As Long, current As Product
    For i = 1 To prodsToRename.Count
        Set current = prodsToRename.Item(i)
        SafeSet current, "Description", "MADE BY AMCO"
    Next i
    lblStatus.Caption = "Renamed all product instances."
End Sub

Private Sub UserForm_Initialize()
    lstOutput.Clear
    lblStatus.Caption = "Ready."
End Sub
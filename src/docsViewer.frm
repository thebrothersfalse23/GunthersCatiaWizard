' filepath: c:\Users\TheFa\OneDrive\Documents\GitHub\GunthersCatiaWizard\forms\DocsViewer.frm
' =====================================================================
' Form: DocsViewer
' Purpose: Interactive documentation browser for Gunther's Catia Wizard.
' =====================================================================
VERSION 5.00
Begin VB.Form DocsViewer
    Caption         =   "Gunther's Catia Wizard – Docs"
    ClientHeight    =   4800
    ClientLeft      =   60
    ClientTop       =   345
    ClientWidth     =   8000
    StartUpPosition =   1  'CenterOwner
    Font.Name       =   "Segoe UI"
    Font.Size       =   10
    Begin VB.TextBox txtSearch
        Height          =   360
        Left            =   120
        Top             =   120
        Width           =   2200
        TabIndex        =   0
    End
    Begin VB.CommandButton btnSearch
        Caption         =   "Search"
        Height          =   360
        Left            =   2360
        Top             =   120
        Width           =   800
        TabIndex        =   1
    End
    Begin VB.ListBox lstTopics
        Height          =   3600
        Left            =   120
        Top             =   600
        Width           =   3200
        TabIndex        =   2
        ' The Click event is triggered by a single click on an item.
    End
    Begin VB.CommandButton btnPrev
        Caption         =   "< Prev"
        Height          =   360
        Left            =   120
        Top             =   4300
        Width           =   1000
        TabIndex        =   3
    End
    Begin VB.CommandButton btnNext
        Caption         =   "Next >"
        Height          =   360
        Left            =   2320
        Top             =   4300
        Width           =   1000
        TabIndex        =   4
    End
    Begin VB.CommandButton btnClose
        Caption         =   "Close"
        Height          =   360
        Left            =   6700
        Top             =   4300
        Width           =   1000
        TabIndex        =   5
    End
    Begin VB.Label lblDetails
        Caption         =   ""
        Height          =   4200
        Left            =   3500
        Top             =   120
        Width           =   4300
        TabIndex        =   6
        Alignment       =   0  'Left Justify
        WordWrap        =   -1 'True
    End
End

' --- Code section ---
Option Explicit

Private docsIndex As Collection
Private docsDetails As Object ' Scripting.Dictionary

Private Sub Form_Load()
    Me.Caption = "Gunther's Catia Wizard – Docs"
    loadDocsData
    populateTopics ""
    If lstTopics.ListCount > 0 Then
        lstTopics.ListIndex = 0
        showDetails lstTopics.List(0)
    End If
End Sub

Private Sub loadDocsData()
    ' Build the docs index and details dictionary
    Set docsIndex = New Collection
    Set docsDetails = CreateObject("Scripting.Dictionary")
    '---[getProducts]
    addDoc "getProducts", _
        "getProducts(rootProd As Product, [unique As Boolean = False]) As Collection" & vbCrLf & _
        "Returns a Collection of reference Products (Products only) in the assembly. Optionally deduplicates by reference." & vbCrLf & _
        "Args:" & vbCrLf & _
        "  rootProd [Product] - root product to traverse" & vbCrLf & _
        "  unique   [Boolean, Optional] - True for unique refs, False for all (default: False)" & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  Set prodsAll = getProducts(rootProd, False)" & vbCrLf & _
        "  Set prodsUniq = getProducts(rootProd, True)"
    '---[getParts]
    addDoc "getParts", _
        "getParts(rootProd As Product, [unique As Boolean = False]) As Collection" & vbCrLf & _
        "Returns a Collection of reference Parts (Parts only) in the assembly. Optionally deduplicates by reference." & vbCrLf & _
        "Args:" & vbCrLf & _
        "  rootProd [Product] - root product to traverse" & vbCrLf & _
        "  unique   [Boolean, Optional] - True for unique refs, False for all (default: False)" & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  Set partsAll = getParts(rootProd, False)" & vbCrLf & _
        "  Set partsUniq = getParts(rootProd, True)"
    '---[getUniques]
    addDoc "getUniques", _
        "getUniques(rootProd As Product, [kind As uniqueOutKind = uoAll]) As Collection" & vbCrLf & _
        "Returns unique references (Products and/or Parts)." & vbCrLf & _
        "Args:" & vbCrLf & _
        "  rootProd [Product] - root product to traverse" & vbCrLf & _
        "  kind     [uniqueOutKind, Optional] - uoAll, uoProductsOnly, uoPartsOnly (default: uoAll)" & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  Set uniqAll = getUniques(rootProd, uoAll)" & vbCrLf & _
        "  Set uniqProds = getUniques(rootProd, uoProductsOnly)" & vbCrLf & _
        "  Set uniqParts = getUniques(rootProd, uoPartsOnly)"
    '---[getInstances]
    addDoc "getInstances", _
        "getInstances(rootProd As Product, [kind As uniqueOutKind = uoAll]) As Collection" & vbCrLf & _
        "Returns a Collection of instance Products (not references) in the assembly. Filtering by kind is available." & vbCrLf & _
        "Args:" & vbCrLf & _
        "  rootProd [Product] - root product to traverse" & vbCrLf & _
        "  kind     [uniqueOutKind, Optional] - uoAll, uoProductsOnly, uoPartsOnly (default: uoAll)" & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  Set instAll = getInstances(rootProd, uoAll)" & vbCrLf & _
        "  Set instProds = getInstances(rootProd, uoProductsOnly)" & vbCrLf & _
        "  Set instParts = getInstances(rootProd, uoPartsOnly)"
     '---[safeSet]
      addDoc "safeSet", _
          "safeSet(obj As Object, propName As String, value As String)" & vbCrLf & _
          "Safely sets a property (e.g., 'Description', 'Name', 'PartNumber') on a CATIA object if it exists." & vbCrLf & _
          "Args:" & vbCrLf & _
          "  obj [Object] - object to set property on" & vbCrLf & _
          "  propName [String] - property name" & vbCrLf & _
          "  value [String] - value to assign" & vbCrLf & _
          "Usage:" & vbCrLf & _
          "  safeSet prod, 'Description', 'MADE BY AMCO'"
    '---[getPropStr]
    addDoc "getPropStr", _
        "getPropStr(obj As Object, propName As String) As String" & vbCrLf & _
        "Safely retrieves a property value as a string from a CATIA object." & vbCrLf & _
        "Args:" & vbCrLf & _
        "  obj [Object] - object to get property from" & vbCrLf & _
        "  propName [String] - property name" & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  desc = getPropStr(prod, 'Description')"
    '---[buildRefKey]
    addDoc "buildRefKey", _
        "buildRefKey(ref As Product, docType As String) As String" & vbCrLf & _
        "Builds a stable, human-readable key for a reference product." & vbCrLf & _
        "Args:" & vbCrLf & _
        "  ref [Product] - reference product" & vbCrLf & _
        "  docType [String] - document type ('ProductDocument' or 'PartDocument')" & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  key = buildRefKey(ref, 'ProductDocument')"
    '---[ensureActiveProductDocument]
    addDoc "ensureActiveProductDocument", _
        "ensureActiveProductDocument() As Boolean" & vbCrLf & _
        "Ensures a ProductDocument is open and active in CATIA. Sets globals if successful." & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  If Not ensureActiveProductDocument() Then Exit Sub"
    '---[ensureDesignMode]
    addDoc "ensureDesignMode", _
        "ensureDesignMode(root As Product)" & vbCrLf & _
        "Applies Design Mode to a product for consistent traversal." & vbCrLf & _
        "Args:" & vbCrLf & _
        "  root [Product] - product to apply Design Mode to" & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  ensureDesignMode rootProduct"
    '---[traverseProduct]
    addDoc "traverseProduct", _
        "traverseProduct(mode As traversalMode, root As Product, [ByRef outRefs As Collection], [outKind As uniqueOutKind = uoAll])" & vbCrLf & _
        "Core traversal logic for all wrappers. Iterative BFS queue." & vbCrLf & _
        "Args:" & vbCrLf & _
        "  mode [traversalMode] - traversal mode" & vbCrLf & _
        "  root [Product] - root product to traverse" & vbCrLf & _
        "  outRefs [Collection, Optional, ByRef] - receives output" & vbCrLf & _
        "  outKind [uniqueOutKind, Optional] - output kind (default: uoAll)" & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  traverseProduct tmGetUniques, rootProd, outRefs, uoAll"
    '---[getSelectedProducts]
    addDoc "getSelectedProducts", _
        "getSelectedProducts([firstSelection As Boolean = False]) As Variant" & vbCrLf & _
        "Returns either the first selected product or all selected products as a collection." & vbCrLf & _
        "Args:" & vbCrLf & _
        "  firstSelection [Boolean, Optional] - True for a single Product, False for all (default: False)" & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  Set prod = getSelectedProducts(True)" & vbCrLf & _
        "  Set prods = getSelectedProducts(False)"
    '---[generateTDSingle]
    addDoc "generateTDSingle", _
        "generateTDSingle(selectedProduct As Product, prefix As String, startOnSelected As Boolean, protectRefDocs As Boolean, createNewProduct As Boolean)" & vbCrLf & _
        "Renames and numbers all products and parts in the selected structure using the given prefix. Supports copy-on-write, reference protection, and flexible numbering. Throws error if selection is invalid or has no children." & vbCrLf & _
        "Args:" & vbCrLf & _
        "  selectedProduct [Product] - product to operate on" & vbCrLf & _
        "  prefix [String] - prefix for naming/numbering" & vbCrLf & _
        "  startOnSelected [Boolean] - numbering starts at selected product if True" & vbCrLf & _
        "  protectRefDocs [Boolean] - protect parts with 'REF' in name/part number if True" & vbCrLf & _
        "  createNewProduct [Boolean] - copy structure before renaming if True"
    '---[generateTDSequential]
    addDoc "generateTDSequential", _
        "generateTDSequential()" & vbCrLf & _
        "Executes the sequential tool design logic for the current selection/context."
    '---[traversalMode (enum)]
    addDoc "traversalMode (enum)", _
        "traversalMode: tmGetUniques, tmGetParts, tmAssignInstanceData, tmCollectRefsAll, tmGetInstances, tmDeepCopyStructure"
    '---[uniqueOutKind (enum)]
    addDoc "uniqueOutKind (enum)", _
        "uniqueOutKind: uoAll, uoProductsOnly, uoPartsOnly"
    '---[globals]
    addDoc "globals", _
        "prodDoc [ProductDocument]: active document" & vbCrLf & _
        "rootProd [Product]: root product of the assembly"
    '---[notes]
    addDoc "notes", _
        "All wrappers pass ByRef collections through traverseProduct and return them." & vbCrLf & _
        "Design Mode is applied up-front for consistent traversal depth."
    '---[guardCatiaRunning]
    addDoc "guardCatiaRunning", _
        "guardCatiaRunning() As Boolean" & vbCrLf & _
        "Returns True if CATIA is running, otherwise False." & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  If Not guardCatiaRunning() Then Exit Sub"
    '---[guardActiveDocument]
    addDoc "guardActiveDocument", _
        "guardActiveDocument() As Boolean" & vbCrLf & _
        "Returns True if there is an active document in CATIA, otherwise False." & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  If Not guardActiveDocument() Then Exit Sub"
    '---[guardProductDocument]
    addDoc "guardProductDocument", _
        "guardProductDocument() As Boolean" & vbCrLf & _
        "Returns True if the active document is a ProductDocument, otherwise False." & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  If Not guardProductDocument() Then Exit Sub"
    '---[guardProductSelection]
    addDoc "guardProductSelection", _
        "guardProductSelection() As Boolean" & vbCrLf & _
        "Returns True if the selection object exists and has at least one Product selected." & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  If Not guardProductSelection() Then Exit Sub"
    '---[guardDesignMode]
    addDoc "guardDesignMode", _
        "guardDesignMode() As Boolean" & vbCrLf & _
        "Returns True if the root product is in Design Mode (or can be set), otherwise False." & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  If Not guardDesignMode() Then Exit Sub"
    '---[runAllGuards]
    addDoc "runAllGuards", _
        "runAllGuards() As Boolean" & vbCrLf & _
        "Checks all preconditions for running CATIA macros safely. Returns True if all checks pass, False otherwise. Shows a message if any guard fails." & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  If Not runAllGuards() Then Exit Sub"
End Sub

Private Sub addDoc(topic As String, details As String)
    docsIndex.Add topic
    docsDetails.Add topic, details
End Sub

Private Sub populateTopics(filter As String)
    Dim i As Long
    lstTopics.Clear
    For i = 1 To docsIndex.Count
        If filter = "" Or InStr(1, docsIndex(i), filter, vbTextCompare) > 0 Then
            lstTopics.AddItem docsIndex(i)
        End If
    Next i
End Sub

Private Sub showDetails(topic As String)
    If docsDetails.Exists(topic) Then
        lblDetails.Caption = docsDetails(topic)
    Else
        lblDetails.Caption = ""
    End If
End Sub

Private Sub lstTopics_Click()
    If lstTopics.ListIndex >= 0 Then
        showDetails lstTopics.List(lstTopics.ListIndex)
    End If
End Sub

Private Sub btnSearch_Click()
    Dim filter As String
    filter = Trim$(txtSearch.Text)
    populateTopics filter
    If lstTopics.ListCount > 0 Then
        lstTopics.ListIndex = 0
        showDetails lstTopics.List(0)
    Else
        lblDetails.Caption = "No topics found."
    End If
End Sub

Private Sub btnPrev_Click()
    If lstTopics.ListIndex > 0 Then
        lstTopics.ListIndex = lstTopics.ListIndex - 1
        showDetails lstTopics.List(lstTopics.ListIndex)
    End If
End Sub

Private Sub btnNext_Click()
    If lstTopics.ListIndex < lstTopics.ListCount - 1 Then
        lstTopics.ListIndex = lstTopics.ListIndex + 1
        showDetails lstTopics.List(lstTopics.ListIndex)
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
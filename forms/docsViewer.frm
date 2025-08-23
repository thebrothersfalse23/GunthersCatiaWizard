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
        "Returns a collection of reference products (products only) in the assembly." & vbCrLf & _
        "Args: rootProduct [Product], unique [Boolean, Optional]" & vbCrLf & _
        "Usage: set prods = getProducts(rootProduct, true)"
 '---[getParts]
    addDoc "getParts", _
        "Returns a collection of reference parts (parts only) in the assembly." & vbCrLf & _
        "Args: rootProduct [Product], unique [Boolean, Optional]" & vbCrLf & _
        "Usage: set parts = getParts(rootProduct, false)"
 '---[getUniques]
    addDoc "getUniques", _
        "Returns unique reference products and/or parts, ordered with products first." & vbCrLf & _
        "Args: rootProduct [Product], kind [uniqueOutKind, Optional]" & vbCrLf & _
        "Usage: set uniqs = getUniques(rootProduct, uoAll)"
 '---[getInstances]
    addDoc "getInstances", _
        "Returns instance products (not references) in the assembly." & vbCrLf & _
        "Args: rootProduct [Product], kind [uniqueOutKind, Optional]" & vbCrLf & _
        "Usage: set insts = getInstances(rootProduct, uoProductsOnly)"
 '---[safeSet]
    addDoc "safeSet", _
        "Safely sets a property (e.g., 'Description', 'Name') on a CATIA object if it exists." & vbCrLf & _
        "Args: obj [Object], propName [String], value [String]" & vbCrLf & _
        "Usage: safeSet prod, 'Description', 'MADE BY AMCO'"
 '---[getPropStr]
    addDoc "getPropStr", _
        "Safely retrieves a property value as a string from a CATIA object." & vbCrLf & _
        "Args: obj [Object], propName [String]" & vbCrLf & _
        "Usage: desc = getPropStr(prod, 'Description')"
 '---[buildRefKey]
    addDoc "buildRefKey", _
        "Builds a stable, human-readable key for a reference product." & vbCrLf & _
        "Args: ref [Product], docType [String]" & vbCrLf & _
        "Usage: key = buildRefKey(ref, 'ProductDocument')"
 '---[ensureActiveProductDocument]
    addDoc "ensureActiveProductDocument", _
        "Ensures a ProductDocument is open and active in CATIA. Sets globals if successful." & vbCrLf & _
        "Usage: if not ensureActiveProductDocument() then exit sub"
 '---[ensureDesignMode]
    addDoc "ensureDesignMode", _
        "Applies Design Mode to a product for consistent traversal." & vbCrLf & _
        "Args: root [Product]" & vbCrLf & _
        "Usage: ensureDesignMode rootProduct"
 '---[traverseProduct]
    addDoc "traverseProduct", _
        "Core traversal logic for all wrappers. Iterative BFS queue." & vbCrLf & _
        "Args: mode [traversalMode], root [Product], outRefs [Collection, Optional], outKind [uniqueOutKind, Optional]"
 '---[getSelectedProducts]
    addDoc "getSelectedProducts", _
        "Returns either the first selected product or all selected products as a collection." & vbCrLf & _
        "Args: firstSelection [Boolean, Optional]" & vbCrLf & _
        "Usage: set prod = getSelectedProducts(true)"
 '---[generateTDSingle]
    addDoc "generateTDSingle", _
        "Executes the single tool design logic on the selected product." & vbCrLf & _
        "Args: selectedProduct [Product], prefix [String], startOnSelected [Boolean], protectRefDocs [Boolean]"
 '---[generateTDSequential]
    addDoc "generateTDSequential", _
        "Executes the sequential tool design logic for the current selection/context."
 '---[traversalMode (enum)]
    addDoc "traversalMode (enum)", _
        "Enum: tmGetUniques, tmGetParts, tmAssignInstanceData, tmCollectRefsAll, tmGetInstances"
 '---[uniqueOutKind (enum)]
    addDoc "uniqueOutKind (enum)", _
        "Enum: uoAll, uoProductsOnly, uoPartsOnly"
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
        "Returns True if CATIA is running, otherwise False." & vbCrLf & _
        "Usage: If Not guardCatiaRunning() Then Exit Sub"
 '---[guardActiveDocument]
    addDoc "guardActiveDocument", _
        "Returns True if there is an active document in CATIA, otherwise False." & vbCrLf & _
        "Usage: If Not guardActiveDocument() Then Exit Sub"
 '---[guardProductDocument]
    addDoc "guardProductDocument", _
        "Returns True if the active document is a ProductDocument, otherwise False." & vbCrLf & _
        "Usage: If Not guardProductDocument() Then Exit Sub"
 '---[guardProductSelection]
    addDoc "guardProductSelection", _
        "Returns True if the selection object exists and has at least one Product selected." & vbCrLf & _
        "Usage: If Not guardProductSelection() Then Exit Sub"
 '---[guardDesignMode]
    addDoc "guardDesignMode", _
        "Returns True if the root product is in Design Mode (or can be set), otherwise False." & vbCrLf & _
        "Usage: If Not guardDesignMode() Then Exit Sub"
 '---[runAllGuards]
    addDoc "runAllGuards", _
        "Checks all preconditions for running CATIA macros safely. Returns True if all checks pass, False otherwise. Shows a message if any guard fails." & vbCrLf & _
        "Usage: If Not runAllGuards() Then Exit Sub"
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
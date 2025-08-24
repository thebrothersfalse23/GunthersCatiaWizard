VERSION 5.00
Begin VB.Form DocsViewer
   Caption         =   "Gunther's Catia Wizard – Docs"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8000
   Font.Name       =   "Segoe UI"
   Font.Size       =   10
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSearch
      Height       =   360
      Left         =   120
      Top          =   120
      Width        =   2200
      TabIndex     =   0
   End
   Begin VB.CommandButton btnSearch
      Caption      =   "Search"
      Height       =   360
      Left         =   2360
      Top          =   120
      Width        =   800
      TabIndex     =   1
   End
   Begin VB.Label lblSort
      Caption      =   "Sort:"
      Height       =   240
      Left         =   3280
      Top          =   180
      Width        =   480
      TabIndex     =   7
   End
   Begin VB.ComboBox cboSort
      Height       =   315
      Left         =   3720
      Top          =   120
      Width        =   1200
      Style        =   2  'Dropdown List
      TabIndex     =   8
   End
   Begin VB.ListBox lstTopics
      Height       =   3600
      Left         =   120
      Top          =   600
      Width        =   3200
      TabIndex     =   2
   End
   Begin VB.CommandButton btnPrev
      Caption      =   "< Prev"
      Height       =   360
      Left         =   120
      Top          =   4300
      Width        =   1000
      TabIndex     =   3
   End
   Begin VB.CommandButton btnNext
      Caption      =   "Next >"
      Height       =   360
      Left         =   2320
      Top          =   4300
      Width        =   1000
      TabIndex     =   4
   End
   Begin VB.CommandButton btnClose
      Caption      =   "Close"
      Height       =   360
      Left         =   6700
      Top          =   4300
      Width        =   1000
      TabIndex     =   5
   End
   Begin VB.Label lblDetails
      Caption      =   ""
      Height       =   4200
      Left         =   3500
      Top          =   120
      Width        =   4300
      TabIndex     =   6
      Alignment    =   0  'Left Justify
      WordWrap     =   -1 'True
   End
End
Attribute VB_Name = "DocsViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =========================
' Module-scope state
' =========================
Private docsIndex As Collection                   ' ordered topic names
Private docsDetails As Object                     ' late-bound Scripting.Dictionary
Private topicCategory As Object                   ' topic -> category name
Private lastFilter As String

Private Enum SortMode
    smGrouped = 0
    smAZ = 1
End Enum

Private Const BULLET As String = "  · "

' Display order for grouped view
Private catOrder() As String

' =========================
' Lifecycle
' =========================
Private Sub Form_Load()
    Set docsIndex = New Collection
    Set docsDetails = CreateObject("Scripting.Dictionary")
    Set topicCategory = CreateObject("Scripting.Dictionary")

    catOrder = Split("UI & Navigation|Configurator|Traversal Core|Wrappers & Queries|Utilities|Guards & Preconditions|Enums|Globals|Notes", "|")

    cboSort.AddItem "Grouped"
    cboSort.AddItem "A–Z"
    cboSort.ListIndex = 0

    loadDocsData
    populateTopics ""
    SelectFirstTopic
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CleanupDocsViewer
End Sub

' =========================
' UI Events
' =========================
Private Sub btnSearch_Click()
    Dim filter As String
    filter = Trim$(txtSearch.Text)
    populateTopics filter
    SelectFirstTopic
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        btnSearch_Click
    End If
End Sub

Private Sub cboSort_Click()
    populateTopics lastFilter
    SelectFirstTopic
End Sub

Private Sub lstTopics_Click()
    If lstTopics.ListIndex < 0 Then Exit Sub
    If lstTopics.ItemData(lstTopics.ListIndex) = -1 Then Exit Sub  ' header row
    Dim disp As String, key As String
    disp = lstTopics.List(lstTopics.ListIndex)
    key = DisplayToKey(disp)
    showDetails key
End Sub

Private Sub btnPrev_Click()
    Dim i As Long
    i = lstTopics.ListIndex
    Do While i > 0
        i = i - 1
        If lstTopics.ItemData(i) <> -1 Then
            lstTopics.ListIndex = i
            lstTopics_Click
            Exit Do
        End If
    Loop
End Sub

Private Sub btnNext_Click()
    Dim i As Long
    i = lstTopics.ListIndex
    Do While i < lstTopics.ListCount - 1
        i = i + 1
        If lstTopics.ItemData(i) <> -1 Then
            lstTopics.ListIndex = i
            lstTopics_Click
            Exit Do
        End If
    Loop
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' =========================
' Core helpers
' =========================
Private Sub loadDocsData()
    ' --- Wrappers & Queries
    addDoc "getProducts", _
        "getProducts(rootProd As Product, [unique As Boolean = False]) As Collection" & vbCrLf & _
        "Returns a Collection of reference Products (Products only) in the assembly. Optionally deduplicates by reference." & vbCrLf & _
        "Args:" & vbCrLf & _
        "  rootProd [Product] - root product to traverse" & vbCrLf & _
        "  unique   [Boolean, Optional] - True for unique refs, False for all (default: False)" & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  Set prodsAll = getProducts(rootProd, False)" & vbCrLf & _
        "  Set prodsUniq = getProducts(rootProd, True)", "Wrappers & Queries"

    addDoc "getParts", _
        "getParts(rootProd As Product, [unique As Boolean = False]) As Collection" & vbCrLf & _
        "Returns a Collection of reference Parts (Parts only) in the assembly. Optionally deduplicates by reference." & vbCrLf & _
        "Args:" & vbCrLf & _
        "  rootProd [Product] - root product to traverse" & vbCrLf & _
        "  unique   [Boolean, Optional] - True for unique refs, False for all (default: False)" & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  Set partsAll = getParts(rootProd, False)" & vbCrLf & _
        "  Set partsUniq = getParts(rootProd, True)", "Wrappers & Queries"

    addDoc "getUniques", _
        "getUniques(rootProd As Product, [kind As uniqueOutKind = uoAll]) As Collection" & vbCrLf & _
        "Returns unique references (Products and/or Parts)." & vbCrLf & _
        "Args:" & vbCrLf & _
        "  rootProd [Product] - root product to traverse" & vbCrLf & _
        "  kind     [uniqueOutKind, Optional] - uoAll, uoProductsOnly, uoPartsOnly (default: uoAll)" & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  Set uniqAll = getUniques(rootProd, uoAll)" & vbCrLf & _
        "  Set uniqProds = getUniques(rootProd, uoProductsOnly)" & vbCrLf & _
        "  Set uniqParts = getUniques(rootProd, uoPartsOnly)", "Wrappers & Queries"

    addDoc "getInstances", _
        "getInstances(rootProd As Product, [kind As uniqueOutKind = uoAll]) As Collection" & vbCrLf & _
        "Returns a Collection of instance Products (not references) in the assembly. Filtering by kind is available." & vbCrLf & _
        "Args:" & vbCrLf & _
        "  rootProd [Product] - root product to traverse" & vbCrLf & _
        "  kind     [uniqueOutKind, Optional] - uoAll, uoProductsOnly, uoPartsOnly (default: uoAll)" & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  Set instAll = getInstances(rootProd, uoAll)" & vbCrLf & _
        "  Set instProds = getInstances(rootProd, uoProductsOnly)" & vbCrLf & _
        "  Set instParts = getInstances(rootProd, uoPartsOnly)", "Wrappers & Queries"

    addDoc "getSelectedProducts", _
        "getSelectedProducts([firstSelection As Boolean = False]) As Variant" & vbCrLf & _
        "Returns either the first selected product or all selected products as a collection." & vbCrLf & _
        "Args:" & vbCrLf & _
        "  firstSelection [Boolean, Optional] - True for a single Product, False for all (default: False)" & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  Set prod = getSelectedProducts(True)" & vbCrLf & _
        "  Set prods = getSelectedProducts(False)", "Wrappers & Queries"

    addDoc "generateTDSingle", _
        "generateTDSingle(selectedProduct As Product, prefix As String, startOnSelected As Boolean, protectRefDocs As Boolean, createNewProduct As Boolean)" & vbCrLf & _
        "Renames and numbers all products and parts in the selected structure using the given prefix. Supports copy-on-write, reference protection, and flexible numbering. Throws error if selection is invalid or has no children.", "Wrappers & Queries"

    addDoc "generateTDSequential", _
        "generateTDSequential()" & vbCrLf & _
        "Executes the sequential tool design logic for the current selection/context.", "Wrappers & Queries"

    ' --- Utilities
    addDoc "safeSet", _
        "safeSet(obj As Object, propName As String, value As String)" & vbCrLf & _
        "Safely sets a property (e.g., 'Description', 'Name', 'PartNumber') on a CATIA object if it exists." & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  safeSet prod, ""Description"", ""MADE BY AMCO""", "Utilities"

    addDoc "getPropStr", _
        "getPropStr(obj As Object, propName As String) As String" & vbCrLf & _
        "Safely retrieves a property value as a string from a CATIA object." & vbCrLf & _
        "Usage:" & vbCrLf & _
        "  desc = getPropStr(prod, ""Description"")", "Utilities"

    addDoc "buildRefKey", _
        "buildRefKey(ref As Product, docType As String) As String" & vbCrLf & _
        "Builds a stable, human-readable key for a reference product.", "Utilities"

    ' --- Guards & Preconditions
    addDoc "ensureActiveProductDocument", _
        "ensureActiveProductDocument() As Boolean" & vbCrLf & _
        "Ensures a ProductDocument is open and active in CATIA.", "Guards & Preconditions"

    addDoc "ensureDesignMode", _
        "ensureDesignMode(root As Product)" & vbCrLf & _
        "Applies Design Mode to a product for consistent traversal.", "Guards & Preconditions"

    addDoc "guardCatiaRunning", _
        "guardCatiaRunning() As Boolean" & vbCrLf & _
        "Returns True if CATIA is running, otherwise False.", "Guards & Preconditions"

    addDoc "guardActiveDocument", _
        "guardActiveDocument() As Boolean" & vbCrLf & _
        "Returns True if there is an active document in CATIA, otherwise False.", "Guards & Preconditions"

    addDoc "guardProductDocument", _
        "guardProductDocument() As Boolean" & vbCrLf & _
        "Returns True if the active document is a ProductDocument, otherwise False.", "Guards & Preconditions"

    addDoc "guardProductSelection", _
        "guardProductSelection() As Boolean" & vbCrLf & _
        "Returns True if the selection has at least one Product.", "Guards & Preconditions"

    addDoc "guardDesignMode", _
        "guardDesignMode() As Boolean" & vbCrLf & _
        "Returns True if root is in (or can be set to) Design Mode.", "Guards & Preconditions"

    addDoc "runAllGuards", _
        "runAllGuards() As Boolean" & vbCrLf & _
        "Checks all preconditions to run CATIA macros safely.", "Guards & Preconditions"

    ' --- Traversal Core
    addDoc "traverseProduct", _
        "traverseProduct(mode As traversalMode, root As Product, [ByRef outRefs As Collection], [outKind As uniqueOutKind = uoAll])" & vbCrLf & _
        "Core traversal logic for all wrappers. Iterative BFS queue.", "Traversal Core"

    ' --- Configurator
    addDoc "setupConfiguratorFields", _
        "setupConfiguratorFields(frame As Frame, configFields As Variant)" & vbCrLf & _
        "Dynamically creates and arranges configurator input fields.", "Configurator"

    addDoc "getConfiguratorValues", _
        "getConfiguratorValues(frame As Frame) As Dictionary" & vbCrLf & _
        "Returns a dictionary of configurator field values keyed by field name.", "Configurator"

    ' --- UI & Navigation
    addDoc "Launchpad Navigation", _
        "Wizard-style navigation via showPage(pageEnum).", "UI & Navigation"

    addDoc "Back/Docs Button Logic", _
        "Always-visible Back and Docs buttons. btnBack_Click is context-aware; btnDocs_Click opens DocsViewer.", "UI & Navigation"

    ' --- Enums / Globals / Notes
    addDoc "traversalMode (enum)", _
        "traversalMode: tmGetUniques, tmGetParts, tmAssignInstanceData, tmCollectRefsAll, tmGetInstances, tmDeepCopyStructure", "Enums"

    addDoc "uniqueOutKind (enum)", _
        "uniqueOutKind: uoAll, uoProductsOnly, uoPartsOnly", "Enums"

    addDoc "globals", _
        "prodDoc [ProductDocument]: active document" & vbCrLf & _
        "rootProd [Product]: root product of the assembly", "Globals"

    addDoc "notes", _
        "All wrappers pass ByRef collections through traverseProduct and return them." & vbCrLf & _
        "Design Mode is applied up-front for consistent traversal depth.", "Notes"
End Sub

Private Sub addDoc(ByVal topic As String, ByVal details As String, Optional ByVal category As String = "Uncategorized")
    Dim i As Long, exists As Boolean

    If docsIndex Is Nothing Then Set docsIndex = New Collection
    If docsDetails Is Nothing Then Set docsDetails = CreateObject("Scripting.Dictionary")
    If topicCategory Is Nothing Then Set topicCategory = CreateObject("Scripting.Dictionary")

    exists = False
    For i = 1 To docsIndex.Count
        If StrComp(docsIndex(i), topic, vbTextCompare) = 0 Then
            exists = True
            Exit For
        End If
    Next i
    If Not exists Then docsIndex.Add topic

    If Not docsDetails.Exists(topic) Then
        docsDetails.Add topic, details
    Else
        docsDetails(topic) = details
    End If

    topicCategory(topic) = category
End Sub

Private Sub populateTopics(ByVal filter As String)
    Dim mode As SortMode
    mode = CurrentSortMode()

    lstTopics.Clear

    If mode = smGrouped Then
        PopulateGrouped filter
    Else
        PopulateAZ filter
    End If

    lastFilter = filter
End Sub

Private Sub PopulateGrouped(ByVal filter As String)
    Dim cat As Variant, topics() As String
    Dim buf() As String, n As Long, i As Long

    For Each cat In catOrder
        buf = FilteredTopicsForCategory(filter, CStr(cat))
        If UBoundSafe(buf) >= 0 Then
            ' Header row (non-selectable)
            lstTopics.AddItem "— " & CStr(cat) & " —"
            lstTopics.ItemData(lstTopics.NewIndex) = -1

            SortStrings buf
            For i = LBound(buf) To UBound(buf)
                lstTopics.AddItem BULLET & buf(i)
                ' ItemData defaults to 0 for normal items
            Next i
        End If
    Next cat

    ' Also show any topics in categories not listed in catOrder
    buf = FilteredTopicsForCategory(filter, "Uncategorized")
    If UBoundSafe(buf) >= 0 Then
        lstTopics.AddItem "— Uncategorized —": lstTopics.ItemData(lstTopics.NewIndex) = -1
        SortStrings buf
        For i = LBound(buf) To UBound(buf)
            lstTopics.AddItem BULLET & buf(i)
        Next i
    End If
End Sub

Private Sub PopulateAZ(ByVal filter As String)
    Dim allTopics() As String, i As Long
    allTopics = FilteredTopics(filter)
    If UBoundSafe(allTopics) < 0 Then Exit Sub
    SortStrings allTopics
    For i = LBound(allTopics) To UBound(allTopics)
        lstTopics.AddItem allTopics(i)
    Next i
End Sub

Private Function FilteredTopics(ByVal filter As String) As String()
    Dim i As Long, t As String, d As String
    Dim arr() As String, n As Long
    n = -1
    For i = 1 To docsIndex.Count
        t = docsIndex(i)
        d = docsDetails(t)
        If Len(filter) = 0 _
           Or InStr(1, t, filter, vbTextCompare) > 0 _
           Or InStr(1, d, filter, vbTextCompare) > 0 Then
            n = n + 1
            ReDim Preserve arr(0 To n)
            arr(n) = t
        End If
    Next i
    FilteredTopics = arr
End Function

Private Function FilteredTopicsForCategory(ByVal filter As String, ByVal cat As String) As String()
    Dim i As Long, t As String, d As String, c As String
    Dim arr() As String, n As Long
    n = -1
    For i = 1 To docsIndex.Count
        t = docsIndex(i)
        c = CStr(topicCategory(t))
        If StrComp(c, cat, vbTextCompare) <> 0 Then
            If Not (StrComp(cat, "Uncategorized", vbTextCompare) = 0 And Not HasKnownCategory(c)) Then
                GoTo ContinueNext
            End If
        End If
        d = docsDetails(t)
        If Len(filter) = 0 _
           Or InStr(1, t, filter, vbTextCompare) > 0 _
           Or InStr(1, d, filter, vbTextCompare) > 0 Then
            n = n + 1
            ReDim Preserve arr(0 To n)
            arr(n) = t
        End If
ContinueNext:
    Next i
    FilteredTopicsForCategory = arr
End Function

Private Function HasKnownCategory(ByVal c As String) As Boolean
    Dim x As Variant
    For Each x In catOrder
        If StrComp(CStr(x), c, vbTextCompare) = 0 Then HasKnownCategory = True: Exit Function
    Next x
End Function

Private Function UBoundSafe(ByRef arr() As String) As Long
    On Error GoTo EmptyArr
    UBoundSafe = UBound(arr)
    Exit Function
EmptyArr:
    UBoundSafe = -1
End Function

Private Sub SortStrings(ByRef a() As String)
    If UBoundSafe(a) < 1 Then Exit Sub
    QuickSort a, LBound(a), UBound(a)
End Sub

Private Sub QuickSort(ByRef a() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim p As String, tmp As String
    i = lo: j = hi
    p = a((lo + hi) \ 2)
    Do While i <= j
        Do While StrComp(a(i), p, vbTextCompare) < 0: i = i + 1: Loop
        Do While StrComp(a(j), p, vbTextCompare) > 0: j = j - 1: Loop
        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSort a, lo, j
    If i < hi Then QuickSort a, i, hi
End Sub

Private Function CurrentSortMode() As SortMode
    If cboSort.ListIndex = 1 Then
        CurrentSortMode = smAZ
    Else
        CurrentSortMode = smGrouped
    End If
End Function

Private Function DisplayToKey(ByVal display As String) As String
    If Left$(display, Len(BULLET)) = BULLET Then
        DisplayToKey = Trim$(Mid$(display, Len(BULLET) + 1))
    Else
        DisplayToKey = display
    End If
End Function

Private Sub SelectFirstTopic()
    Dim i As Long
    If lstTopics.ListCount = 0 Then
        lblDetails.Caption = "No topics found."
        Exit Sub
    End If
    For i = 0 To lstTopics.ListCount - 1
        If lstTopics.ItemData(i) <> -1 Then
            lstTopics.ListIndex = i
            lstTopics_Click
            Exit For
        End If
    Next i
End Sub

Private Sub showDetails(ByVal topic As String)
    If Not docsDetails Is Nothing And docsDetails.Exists(topic) Then
        lblDetails.Caption = docsDetails(topic)
    Else
        lblDetails.Caption = ""
    End If
End Sub

Private Sub CleanupDocsViewer()
    On Error Resume Next
    Set docsIndex = Nothing
    Set docsDetails = Nothing
    Set topicCategory = Nothing
    lastFilter = vbNullString
End Sub

' ==============================================================================
' Form: Launchpad
' Description:
'   Main user interface for "Gunther's Catia Wizard".
'   Provides a tabbed interface for launching CATIA automation tools:
'     - Home
'     - Single Tool Design
'     - Sequential Tool Design
'
' Controls:
'   - SSTab tabMain: Main tab control with three tabs:
'       0. Home: Welcome message and instructions, Run, Cancel, Docs buttons.
'       1. Single Tool Design: Serial number prefix input, Run, Back buttons.
'       2. Sequential Tool Design: Run and Back buttons.
'
'   - Labels, TextBoxes, and CommandButtons are shown/hidden depending on the
'     selected tab and workflow state.
'
' Main Procedures:
'   - formLoad: Initializes the form, showing only the Home tab.
'   - btnRun_Click: Validates CATIA state, switches to tool design tabs.
'   - btnCancel_Click: Closes the form.
'   - btnDocs_Click: Opens the interactive documentation viewer.
'   - cmdBackSingle_Click, cmdBackSeq_Click: Return to Home tab.
'   - cmdRunSingle_Click: Validates input and triggers single tool design logic.
'   - cmdRunSeq_Click: Triggers sequential tool design logic.
'   - tabMain_Click: Handles tab switching and visibility of controls.
'
' Helper Procedures:
'   - showHomeTabsOnly: Restricts tab control to Home tab only.
'   - showToolDesignTabsOnly: Shows only tool design tabs.
'   - showHomeTab, showSingleToolTab, showSeqToolTab: Manage visibility of controls
'     for each tab context.
'
' Usage Notes:
'   - The form expects a CATProduct document to be open and active in CATIA.
'   - The user must provide a serial number prefix for single tool design.
'   - The actual tool design logic should be implemented where indicated.
' ==============================================================================

VERSION 5.00
Begin VB.Form Launchpad 
    Caption         =   "Gunther's Catia Wizard"
    ClientHeight    =   4200
    ClientLeft      =   60
    ClientTop       =   345
    ClientWidth     =   6000
    LinkTopic       =   "Launchpad"
    ScaleHeight     =   4200
    ScaleWidth      =   6000
    StartUpPosition =   1  'CenterOwner
    Font.Name       =   "Segoe UI"
    Font.Size       =   10
    Begin VB.SSTab tabMain
        Height          =   4000
        Left            =   0
        Top             =   0
        Width           =   6000
        TabHeight       =   420
        Tabs            =   3
        Tab             =   0
        TabCaption(0)   =   "Home"
        TabCaption(1)   =   "Single Tool Design"
        TabCaption(2)   =   "Sequential Tool Design"
        ' --- Tab 0: Home ---
        Begin VB.Label lblTitle
            Caption         =   "Gunther's Catia Wizard"
            Font.Size       =   16
            Font.Bold       =   -1  'True
            Height          =   480
            Left            =   0
            Top             =   120
            Width           =   6000
            Alignment       =   2  'Center
            TabIndex        =   0
        End
        Begin VB.Label lblInstructions
            Caption         =   "Welcome! To use this wizard:" & vbCrLf & _
                                 "• Open a CATProduct document in CATIA." & vbCrLf & _
                                 "• Ensure the assembly is fully loaded (Design Mode recommended)." & vbCrLf & _
                                 "• Ensure You have selected the top product." & vbCrLf & _
                                 "• Save your work before running tools." & vbCrLf & _
                                 "• Click 'Run' to verify prerequisites or 'Cancel' to exit."
            Height          =   900
            Left            =   360
            Top             =   720
            Width           =   5280
            Alignment       =   0  'Left Justify
            TabIndex        =   1
        End
        Begin VB.CommandButton btnRun
            Caption         =   "Run"
            Height          =   420
            Left            =   1800
            Top             =   2000
            Width           =   1000
            Enabled         =   True
            TabIndex        =   2
        End
        Begin VB.CommandButton btnCancel
            Caption         =   "Cancel"
            Height          =   420
            Left            =   3200
            Top             =   2000
            Width           =   1000
            Enabled         =   True
            TabIndex        =   3
        End
        Begin VB.CommandButton btnDocs
            Caption         =   "Docs"
            Height          =   420
            Left            =   4600
            Top             =   2000
            Width           =   1000
            Enabled         =   True
            TabIndex        =   14
        End
        ' --- Tab 1: Single Tool Design ---
        Begin VB.Label lblSingleInstructions
            Caption         =   "Select the product to act on. This can be any product with children."
            Height          =   300
            Left            =   360
            Top             =   300
            Width           =   5280
            TabIndex        =   11
            Alignment       =   0  'Left Justify
        End
        Begin VB.Label lblSerialPrefix
            Caption         =   "Serial Number Prefix:"
            Height          =   300
            Left            =   360
            Top             =   600
            Width           =   1800
            TabIndex        =   4
            Alignment       =   0  'Left Justify
        End
        Begin VB.TextBox txtSerialPrefix
            Height          =   300
            Left            =   2200
            Top             =   600
            Width           =   1800
            TabIndex        =   5
        End
        Begin VB.Label lblSerialExample
            Caption         =   "Example: VG201144"
            Height          =   300
            Left            =   2200
            Top             =   960
            Width           =   1800
            TabIndex        =   6
            ForeColor       =   &H00808080&
            Alignment       =   0  'Left Justify
        End
        ' --- StartOnSelected Option Group ---
        Begin VB.Label lblStartOnSelected
            Caption         =   "Suffix for selected product:"
            Height          =   300
            Left            =   360
            Top             =   1320
            Width           =   1800
            TabIndex        =   15
            Alignment       =   0  'Left Justify
        End
        Begin VB.OptionButton optStartNoSuffix
            Caption         =   "No suffix for selected product"
            Height          =   300
            Left            =   2200
            Top             =   1320
            Width           =   2200
            TabIndex        =   16
            Value           =   True
        End
        Begin VB.OptionButton optStartWithSuffix
            Caption         =   "Begin suffix at selected product (-0001)"
            Height          =   300
            Left            =   2200
            Top             =   1620
            Width           =   3000
            TabIndex        =   17
        End
        ' --- ProtectRefDocs Option Group ---
        Begin VB.Label lblProtectRefDocs
            Caption         =   "Reference document renaming:"
            Height          =   300
            Left            =   360
            Top             =   1980
            Width           =   2200
            TabIndex        =   18
            Alignment       =   0  'Left Justify
        End
        Begin VB.OptionButton optProtectRefDocs
            Caption         =   "Protect reference docs (replace prefix only)"
            Height          =   300
            Left            =   2600
            Top             =   1980
            Width           =   3000
            TabIndex        =   19
            Value           =   True
        End
        Begin VB.OptionButton optOverwriteRefDocs
            Caption         =   "Overwrite reference doc names"
            Height          =   300
            Left            =   2600
            Top             =   2280
            Width           =   3000
            TabIndex        =   20
        End
        ' --- CreateOrRename Option Group ---
        Begin VB.Label lblCreateOrRename
            Caption         =   "Operation:"
            Height          =   300
            Left            =   360
            Top             =   2640
            Width           =   1800
            TabIndex        =   21
            Alignment       =   0  'Left Justify
        End
        Begin VB.OptionButton optRenameFiles
            Caption         =   "Rename existing files"
            Height          =   300
            Left            =   2200
            Top             =   2640
            Width           =   2200
            TabIndex        =   22
            Value           =   True
        End
        Begin VB.OptionButton optCreateNewProduct
            Caption         =   "Create new product"
            Height          =   300
            Left            =   2200
            Top             =   2940
            Width           =   2200
            TabIndex        =   23
        End
        ' --- Tab 2: Sequential Tool Design ---
        Begin VB.CommandButton cmdRunSeq
            Caption         =   "Run"
            Height          =   360
            Left            =   2200
            Top             =   600
            Width           =   1000
            TabIndex        =   9
        End
        Begin VB.CommandButton cmdBackSeq
            Caption         =   "Back"
            Height          =   360
            Left            =   3400
            Top             =   600
            Width           =   1000
            TabIndex        =   10
        End
    End
End

' --- Code section ---
Option Explicit

Private Sub formLoad()
    showHomeTabsOnly
    tabMain.Tab = 0
    showHomeTab
End Sub

Private Sub btnRun_Click()
    ' Ensure a valid ProductDocument is active before running any tool
    If Not ensureActiveProductDocument() Then
        Exit Sub
    End If
    ' Ensure Design Mode is applied to root product
    ensureDesignMode rootProd
    ' Switch to Single Tool Design tab
    showToolDesignTabsOnly
    tabMain.Tab = 0 ' Show Single Tool Design tab (now index 0)
    showSingleToolTab
End Sub

Private Sub btnCancel_Click()
    Unload Me
    End
End Sub

Private Sub btnDocs_Click()
    showDocsViewer
End Sub

Private Sub cmdBackSingle_Click()
    tabMain.Tab = 0
    showHomeTab
End Sub

Private Sub cmdBackSeq_Click()
    tabMain.Tab = 0
    showHomeTab
End Sub

Private Sub cmdRunSingle_Click()
    If Trim(txtSerialPrefix.Text) = "" Then
        MsgBox "Serial Number Prefix cannot be empty.", vbExclamation, "Input Error"
        txtSerialPrefix.SetFocus
        Exit Sub
    End If
    Dim prefix As String
    Dim startOnSelected As Boolean
    Dim protectRefDocs As Boolean
    Dim createNewProduct As Boolean
    Dim selectedProduct As Product

    prefix = Trim(txtSerialPrefix.Text)
    startOnSelected = optStartWithSuffix.Value
    protectRefDocs = optProtectRefDocs.Value
    createNewProduct = optCreateNewProduct.Value
    Set selectedProduct = getSelectedProducts(True) ' Returns a Product or Nothing

    generateTDSingle selectedProduct, prefix, startOnSelected, protectRefDocs, createNewProduct
End Sub

Private Sub cmdRunSeq_Click()
    generateTDSequential
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
    If tabMain.Tabs = 1 Then
        showHomeTab
    Else
        Select Case tabMain.Tab
            Case 0
                showSingleToolTab
            Case 1
                showSeqToolTab
        End Select
    End If
End Sub

' Restrict tab control to Home tab only
Private Sub showHomeTabsOnly()
    tabMain.Tabs = 1
    tabMain.TabCaption(0) = "Home"
    tabMain.Tab = 0
End Sub

' Show only tool design tabs
Private Sub showToolDesignTabsOnly()
    tabMain.Tabs = 2
    tabMain.TabCaption(0) = "Single Tool Design"
    tabMain.TabCaption(1) = "Sequential Tool Design"
End Sub

' Show only home controls
Private Sub showHomeTab()
    lblTitle.Visible = True
    lblInstructions.Visible = True
    btnRun.Visible = True
    btnCancel.Visible = True
    btnDocs.Visible = True

    lblSingleInstructions.Visible = False
    lblSerialPrefix.Visible = False
    txtSerialPrefix.Visible = False
    lblSerialExample.Visible = False
    lblStartOnSelected.Visible = False
    optStartNoSuffix.Visible = False
    optStartWithSuffix.Visible = False
    lblProtectRefDocs.Visible = False
    optProtectRefDocs.Visible = False
    optOverwriteRefDocs.Visible = False
    lblCreateOrRename.Visible = False
    optRenameFiles.Visible = False
    optCreateNewProduct.Visible = False
    cmdRunSingle.Visible = False
    cmdBackSingle.Visible = False
    cmdRunSeq.Visible = False
    cmdBackSeq.Visible = False
End Sub

' Show only single tool design controls
Private Sub showSingleToolTab()
    lblTitle.Visible = False
    lblInstructions.Visible = False
    btnRun.Visible = False
    btnCancel.Visible = False
    btnDocs.Visible = False

    lblSingleInstructions.Visible = True
    lblSerialPrefix.Visible = True
    txtSerialPrefix.Visible = True
    lblSerialExample.Visible = True
    lblStartOnSelected.Visible = True
    optStartNoSuffix.Visible = True
    optStartWithSuffix.Visible = True
    lblProtectRefDocs.Visible = True
    optProtectRefDocs.Visible = True
    optOverwriteRefDocs.Visible = True
    lblCreateOrRename.Visible = True
    optRenameFiles.Visible = True
    optCreateNewProduct.Visible = True
    cmdRunSingle.Visible = True
    cmdBackSingle.Visible = True
    cmdRunSeq.Visible = False
    cmdBackSeq.Visible = False
End Sub

' Show only sequential tool design controls
Private Sub showSeqToolTab()
    lblTitle.Visible = False
    lblInstructions.Visible = False
    btnRun.Visible = False
    btnCancel.Visible = False
    btnDocs.Visible = False

    lblSingleInstructions.Visible = False
    lblSerialPrefix.Visible = False
    txtSerialPrefix.Visible = False
    lblSerialExample.Visible = False
    lblStartOnSelected.Visible = False
    optStartNoSuffix.Visible = False
    optStartWithSuffix.Visible = False
    lblProtectRefDocs.Visible = False
    optProtectRefDocs.Visible = False
    optOverwriteRefDocs.Visible = False
    lblCreateOrRename.Visible = False
    optRenameFiles.Visible = False
    optCreateNewProduct.Visible = False
    cmdRunSingle.Visible = False
    cmdBackSingle.Visible = False
    cmdRunSeq.Visible = True
    cmdBackSeq.Visible = True
End Sub

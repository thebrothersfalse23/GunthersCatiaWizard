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
    ' Remove old Home controls (lblTitle, lblInstructions, btnRun, btnCancel)
    ' Add SSTab with three tabs: Home, Single Tool Design, Sequential Tool Design
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
            Alignment       =   2  'Center (keep title centered)
            TabIndex        =   0
        End
        Begin VB.Label lblInstructions
            Caption         =   "Welcome! To use this wizard:" & vbCrLf & _
                                 "• Open a CATProduct document in CATIA." & vbCrLf & _
                                 "• Ensure the assembly is fully loaded (Design Mode recommended)." & vbCrLf & _
                                 "• Ensure You have selected the top product you wish to modify." & vbCrLf & _
                                 "• Save your work before running tools." & vbCrLf & _
                                 "• Click 'Run' to begin or 'Cancel' to exit."
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
            ' Alignment property does not apply to CommandButton captions
        End
        Begin VB.CommandButton btnCancel
            Caption         =   "Cancel"
            Height          =   420
            Left            =   3200
            Top             =   2000
            Width           =   1000
            Enabled         =   True
            TabIndex        =   3
            ' Alignment property does not apply to CommandButton captions
        End
        ' --- Tab 1: Single Tool Design ---
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
        Begin VB.CommandButton cmdRunSingle
            Caption         =   "Run"
            Height          =   360
            Left            =   2200
            Top             =   1400
            Width           =   1000
            TabIndex        =   7
            ' Alignment property does not apply to CommandButton captions
        End
        Begin VB.CommandButton cmdBackSingle
            Caption         =   "Back"
            Height          =   360
            Left            =   3400
            Top             =   1400
            Width           =   1000
            TabIndex        =   8
            ' Alignment property does not apply to CommandButton captions
        End
        ' --- Tab 2: Sequential Tool Design ---
        Begin VB.CommandButton cmdRunSeq
            Caption         =   "Run"
            Height          =   360
            Left            =   2200
            Top             =   600
            Width           =   1000
            TabIndex        =   9
            ' Alignment property does not apply to CommandButton captions
        End
        Begin VB.CommandButton cmdBackSeq
            Caption         =   "Back"
            Height          =   360
            Left            =   3400
            Top             =   600
            Width           =   1000
            TabIndex        =   10
            ' Alignment property does not apply to CommandButton captions
        End
    End
End

' --- Code section ---
Option Explicit

Private Sub Form_Load()
    ShowHomeTabsOnly
    tabMain.Tab = 0
    ShowHomeTab
End Sub

Private Sub btnRun_Click()
    ' Guard: Ensure a valid ProductDocument is active before running any tool
    If Not EnsureActiveProductDocument() Then
        Exit Sub
    End If
    ' Guard: Ensure Design Mode is applied to root product
    EnsureDesignMode rootProd
    ' Switch to Single Tool Design tab
    ShowToolDesignTabsOnly
    tabMain.Tab = 0 ' Show Single Tool Design tab (now index 0)
    ShowSingleToolTab
End Sub

Private Sub btnCancel_Click()
    Unload Me
    End
End Sub

Private Sub cmdBackSingle_Click()
    ShowHomeTabsOnly
    tabMain.Tab = 0
    ShowHomeTab
End Sub

Private Sub cmdBackSeq_Click()
    ShowHomeTabsOnly
    tabMain.Tab = 0
    ShowHomeTab
End Sub

Private Sub cmdRunSingle_Click()
    If Trim(txtSerialPrefix.Text) = "" Then
        MsgBox "Serial Number Prefix cannot be empty.", vbExclamation, "Input Error"
        txtSerialPrefix.SetFocus
        Exit Sub
    End If
    ' Place your single tool design logic here
End Sub

Private Sub cmdRunSeq_Click()
    ' Place your sequential tool design logic here
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
    If tabMain.Tabs = 1 Then
        ShowHomeTab
    Else
        Select Case tabMain.Tab
            Case 0
                ShowSingleToolTab
            Case 1
                ShowSeqToolTab
        End Select
    End If
End Sub

' Helper to show only the Home tab
Private Sub ShowHomeTabsOnly()
    tabMain.Tabs = 1
    tabMain.TabCaption(0) = "Home"
    tabMain.Tab = 0
End Sub

' Helper to show only the tool design tabs
Private Sub ShowToolDesignTabsOnly()
    tabMain.Tabs = 2
    tabMain.TabCaption(0) = "Single Tool Design"
    tabMain.TabCaption(1) = "Sequential Tool Design"
End Sub

Private Sub ShowHomeTab()
    ' Home tab: show only home controls
    lblTitle.Visible = True
    lblInstructions.Visible = True
    btnRun.Visible = True
    btnCancel.Visible = True

    lblSerialPrefix.Visible = False
    txtSerialPrefix.Visible = False
    lblSerialExample.Visible = False
    cmdRunSingle.Visible = False
    cmdBackSingle.Visible = False
    cmdRunSeq.Visible = False
    cmdBackSeq.Visible = False
End Sub

Private Sub ShowSingleToolTab()
    lblTitle.Visible = False
    lblInstructions.Visible = False
    btnRun.Visible = False
    btnCancel.Visible = False

    lblSerialPrefix.Visible = True
    txtSerialPrefix.Visible = True
    lblSerialExample.Visible = True
    cmdRunSingle.Visible = True
    cmdBackSingle.Visible = True
    cmdRunSeq.Visible = False
    cmdBackSeq.Visible = False
End Sub

Private Sub ShowSeqToolTab()
    lblTitle.Visible = False
    lblInstructions.Visible = False
    btnRun.Visible = False
    btnCancel.Visible = False

    lblSerialPrefix.Visible = False
    txtSerialPrefix.Visible = False
    lblSerialExample.Visible = False
    cmdRunSingle.Visible = False
    cmdBackSingle.Visible = False
    cmdRunSeq.Visible = True
    cmdBackSeq.Visible = True
End Sub

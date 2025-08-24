' ===============================================================================
' TODO: Launchpad Form Completion Checklist
' -------------------------------------------------------------------------------
' [ ] 1. UI Designer: Ensure all frames (frameWelcome, frameNavigation, frameSingleTool, frameSeqTool, frameConfigurator) and controls exist and are named as referenced in code.
' [ ] 2. UI Designer: Place and size Back and Docs buttons so they are always visible and not duplicated on any page.
' [ ] 3. UI Designer: Add/verify all controls referenced in showHomeTab, showNavigationTab, showSingleToolTab, showSeqToolTab routines.
' [ ] 4. Configurator: Test dynamic field creation in setupConfiguratorFields; ensure fields are cleared/created correctly and are accessible.
' [ ] 5. Configurator: Implement getConfiguratorValues to return a dictionary/object with all configurator field values.
' [ ] 6. Configurator: Wire up btnConfigRun_Click to collect values and pass them to the appropriate tool design logic.
' [ ] 7. Navigation: Test all navigation flows (including Back/Docs) for correct page/frame visibility and state.
' [ ] 8. Error Handling: Add user feedback for invalid/missing input on all relevant pages (e.g., configurator, single tool design).
' [ ] 9. Documentation: Update docsViewer and docs.frm to reflect any new/changed public APIs or UI flows.
' [ ] 10. Code Cleanup: Remove any obsolete code, comments, or unused controls after finalizing UI and logic.
' [ ] 11. Testing: Manually test all user flows in CATIA to ensure robust error handling and correct macro execution.
' -------------------------------------------------------------------------------
' ===============================================================================
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
    Tabs            =   4
    Tab             =   0
    TabCaption(0)   =   "Home"
    TabCaption(1)   =   "Navigation"

    ' --- Wizard Page Navigation ---

    Private Enum WizardPage
        PageWelcome = 0
        PageNavigation = 1
        PageSingleTool = 2
        PageSequentialTool = 3
        PageConfigurator = 4
    End Enum
    Private currentPage As WizardPage

    Private Sub showPage(page As WizardPage)

    ' --- Wizard Page Navigation ---
    Option Explicit

    Private Enum WizardPage
        PageWelcome = 0
        PageNavigation = 1
        PageSingleTool = 2
        PageSeqTool = 3
        PageConfigurator = 4
    End Enum
    Private currentPage As WizardPage

    Private Sub Form_Load()
        currentPage = PageWelcome
        showPage currentPage
    End Sub

    Private Sub showPage(page As WizardPage)
        currentPage = page
        ' Hide all frames/pages first
        frameWelcome.Visible = False
        frameNavigation.Visible = False
        frameSingleTool.Visible = False
        frameSeqTool.Visible = False
        frameConfigurator.Visible = False
        ' Show only the relevant frame
        Select Case page
            Case PageWelcome
                frameWelcome.Visible = True
                btnBack.Visible = False
            Case PageNavigation
                frameNavigation.Visible = True
                btnBack.Visible = True
            Case PageSingleTool
                frameSingleTool.Visible = True
                btnBack.Visible = True
            Case PageSeqTool
                frameSeqTool.Visible = True
                btnBack.Visible = True
            Case PageConfigurator
                frameConfigurator.Visible = True
                btnBack.Visible = True
        End Select
        btnDocs.Visible = True
    End Sub

    ' --- Welcome Page Events ---
    Private Sub btnWelcomeRun_Click()
        If Not runAllGuards() Then Exit Sub
        showPage PageNavigation
    End Sub

    Private Sub btnWelcomeCancel_Click()
        Unload Me
    End Sub

    ' --- Navigation Page Events ---
    Private Sub optNavSingle_Click()
        btnNavNext.Enabled = True
    End Sub

    Private Sub optNavSeq_Click()
        btnNavNext.Enabled = True
    End Sub

    Private Sub btnNavNext_Click()
        If optNavSingle.Value Or optNavSingleAdv.Value Then
            showPage PageSingleTool
        ElseIf optNavSeq.Value Then
            showPage PageSeqTool
        End If
    End Sub

    ' --- Single Tool Design Page Events ---
    Private Sub chkAdvancedOptions_Click()
        showAdvancedOptions chkAdvancedOptions.Value
    End Sub

    Private Sub showAdvancedOptions(show As Boolean)
        lblOpt4Right.Visible = show
    End Sub

    Private Sub btnSingleRun_Click()
        If Trim(txtSerialPrefix.Text) = "" Then Exit Sub
        Dim prefix As String
        Dim startOnSelected As Boolean
        Dim protectRefDocs As Boolean
        Dim createNewProduct As Boolean
        prefix = txtSerialPrefix.Text
        startOnSelected = chkStartOnSelected.Value
        protectRefDocs = chkProtectRefDocs.Value
        createNewProduct = chkCreateNewProduct.Value
        generateTDSingle selectedProduct, prefix, startOnSelected, protectRefDocs, createNewProduct
    End Sub

    ' --- Configurator Dynamic Fields ---
    Private configFields As Collection

    Private Sub setupConfiguratorFields()
        ' Clear previous dynamic fields
        Dim ctl As Control
        For Each ctl In frameConfigurator.Controls
            If Left(ctl.Name, 8) = "txtConfig" Or Left(ctl.Name, 8) = "lblConfig" Then
                frameConfigurator.Controls.Remove ctl.Name
            End If
        Next ctl

        Set configFields = New Collection

        ' Define product data fields: Definition, Revision, Nomenclature, Source, Description
        Dim fieldList As Variant
        fieldList = Array(_
            Array("Definition", "Product definition (e.g. part number, type, etc.)"), _
            Array("Revision", "Revision code or version (e.g. A, B, 01)"), _
            Array("Nomenclature", "Nomenclature or short name (e.g. bracket, housing)"), _
            Array("Source", "Source or origin (e.g. Make, Buy, Supplier)"), _
            Array("Description", "Detailed description of the product") _
        )

        Dim i As Integer
        Dim topPos As Integer: topPos = 24
        For i = LBound(fieldList) To UBound(fieldList)
            Dim lbl As Control, txt As Control
            Set lbl = frameConfigurator.Controls.Add("VB.Label", "lblConfig" & i)
            lbl.Caption = fieldList(i)(0)
            lbl.Top = topPos
            lbl.Left = 12
            lbl.Width = 1200
            lbl.ToolTipText = fieldList(i)(1)
            Set txt = frameConfigurator.Controls.Add("VB.TextBox", "txtConfig" & i)
            txt.Top = topPos
            txt.Left = 1300
            txt.Width = 1800
            txt.Text = ""
            txt.ToolTipText = fieldList(i)(1)
            configFields.Add txt, fieldList(i)(0)
            topPos = topPos + 360
        Next i
    End Sub

    Private Function getConfiguratorValues() As Object
        ' ...retrieve configurator values...
        Set getConfiguratorValues = Nothing
    End Function

    ' --- Configurator Page Events ---
    Private Sub btnConfigRun_Click()
        ' ...run configurator logic...
        showPage PageSingleTool
    End Sub

    ' --- Sequential Tool Design Page Events ---
    Private Sub btnSeqRun_Click()
        generateTDSequential
    End Sub

    ' --- Consolidated Back and Docs Buttons ---
    Private Sub btnBack_Click()
        Select Case currentPage
            Case PageNavigation
                showPage PageWelcome
            Case PageSingleTool, PageSeqTool, PageConfigurator
                showPage PageNavigation
            Case Else
                showPage PageWelcome
        End Select
    End Sub

    Private Sub btnDocs_Click()
        showDocsViewer
    End Sub
        Dim ctl As Control
        For Each ctl In frameConfigurator.Controls
            If Left(ctl.Name, 8) = "txtConfig" Or Left(ctl.Name, 8) = "lblConfig" Then
                '' All navigation and docs logic is now handled above. Obsolete routines removed.
            TabIndex        =   10
        End
    End
End

' --- Code section ---
Option Explicit

Private Sub Form_Load()
    showHomeTabsOnly
    tabMain.Tab = 0
    showHomeTab
End Sub

Private Sub btnRun_Click()
    ' Run all guards before proceeding
    If Not runAllGuards() Then
        Exit Sub
    End If
    ' Show navigation tab
    showNavigationTabOnly
    tabMain.Tab = 1 ' Navigation tab
    showNavigationTab
End Sub

Private Sub btnNavBack_Click()
    showHomeTabsOnly
    tabMain.Tab = 0
    showHomeTab
End Sub

Private Sub btnNavNext_Click()
    If optNavSingle.Value Then
        showToolDesignTabsOnly
        tabMain.Tab = 2 ' Single Tool Design
        showSingleToolTab
    ElseIf optNavSingleAdv.Value Then
        ' For now, route to Single Tool Design (Advanced tab can be implemented later)
        showToolDesignTabsOnly
        tabMain.Tab = 2
        showSingleToolTab
    ElseIf optNavSeq.Value Then
        showToolDesignTabsOnly
        tabMain.Tab = 3 ' Sequential Tool Design
        showSeqToolTab
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
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

    ' Build a template CATIA Product with configurator values
    Dim templateProd As Product
    Set templateProd = Nothing
    On Error Resume Next
    Set templateProd = CATIA.ActiveDocument.Products.AddNewComponent("Part")
    On Error GoTo 0
    If Not templateProd Is Nothing Then
        Dim configVals As Object
        Set configVals = getConfiguratorValues()
        Dim propName As Variant
        For Each propName In Array("Definition", "Revision", "Nomenclature", "Source", "Description")
            If Not configVals Is Nothing Then
                If configVals.Exists(propName) Then
                    If Len(Trim$(configVals(propName))) > 0 Then
                        safeSet templateProd, propName, configVals(propName)
                    End If
                End If
            End If
        Next propName
    End If

    generateTDSingle selectedProduct, prefix, startOnSelected, protectRefDocs, createNewProduct, templateProd

    ' Optionally remove the temporary product from the document
    If Not templateProd Is Nothing Then
        On Error Resume Next
        CATIA.ActiveDocument.Products.Remove templateProd
        On Error GoTo 0
    End If
End Sub

Private Sub cmdRunSeq_Click()
    generateTDSequential
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
    If tabMain.Tabs = 1 Then
        showHomeTab
    ElseIf tabMain.Tabs = 2 Then
        Select Case tabMain.Tab
            Case 0
                showSingleToolTab
            Case 1
                showSeqToolTab
        End Select
    ElseIf tabMain.Tabs = 4 Then
        Select Case tabMain.Tab
            Case 0
                showHomeTab
            Case 1
                showNavigationTab
            Case 2
                showSingleToolTab
            Case 3
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

Private Sub showNavigationTabOnly()
    tabMain.Tabs = 2
    tabMain.TabCaption(0) = "Home"
    tabMain.TabCaption(1) = "Navigation"
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
    ' Hide navigation controls
    lblNavTitle.Visible = False
    optNavSingle.Visible = False
    optNavSingleAdv.Visible = False
    optNavSeq.Visible = False
    btnNavNext.Visible = False
    btnNavBack.Visible = False
End Sub

Private Sub showNavigationTab()
    lblNavTitle.Visible = True
    optNavSingle.Visible = True
    optNavSingleAdv.Visible = True
    optNavSeq.Visible = True
    btnNavNext.Visible = True
    btnNavBack.Visible = True

    ' Hide all other controls
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

' --- Consolidated Back and Docs Buttons ---
Private Sub btnBack_Click()
    Select Case currentPage
        Case PageNavigation
            showPage PageWelcome
        Case PageSingleTool
            showPage PageNavigation
        Case PageSequentialTool
            showPage PageNavigation
        Case PageConfigurator
            showPage PageSingleTool
    End Select
End Sub

Private Sub btnDocs_Click()
    showDocsViewer
End Sub

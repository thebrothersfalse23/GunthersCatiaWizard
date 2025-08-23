'===============================================================
' MODULE: guards.bas
' PURPOSE: Provides discrete guard routines for CATIA macros.
'          Ensures correct environment and selection for safe traversal.
'
' LOGIC FLOW:
'   - Each guard function checks a single precondition for safe macro execution.
'   - runAllGuards() calls each guard in sequence:
'       1. guardCatiaRunning:      CATIA application is running.
'       2. guardActiveDocument:    There is an active document in CATIA.
'       3. guardProductDocument:   The active document is a ProductDocument.
'       4. guardProductSelection:  At least one Product is selected in the tree.
'       5. guardDesignMode:        The root product is in Design Mode (or can be set).
'   - If any guard fails, runAllGuards shows a message and returns False.
'   - Only if all guards pass, runAllGuards sets global variables and returns True.
'   - Use these guards at the start of any macro to prevent invalid state or user error.
'===============================================================
Option Explicit

'---------------------------------------------------------------
' Function: guardCatiaRunning
'   Returns True if CATIA is running, otherwise False.
'---------------------------------------------------------------
Public Function guardCatiaRunning() As Boolean
    guardCatiaRunning = Not (CATIA Is Nothing)
End Function

'---------------------------------------------------------------
' Function: guardActiveDocument
'   Returns True if there is an active document in CATIA, otherwise False.
'---------------------------------------------------------------
Public Function guardActiveDocument() As Boolean
    guardActiveDocument = False
    If CATIA Is Nothing Then Exit Function
    If CATIA.Documents.Count = 0 Then Exit Function
    If CATIA.ActiveDocument Is Nothing Then Exit Function
    guardActiveDocument = True
End Function

'---------------------------------------------------------------
' Function: guardProductDocument
'   Returns True if the active document is a ProductDocument, otherwise False.
'---------------------------------------------------------------
Public Function guardProductDocument() As Boolean
    guardProductDocument = False
    If Not guardActiveDocument() Then Exit Function
    If TypeName(CATIA.ActiveDocument) <> "ProductDocument" Then Exit Function
    guardProductDocument = True
End Function

'---------------------------------------------------------------
' Function: guardProductSelection
'   Returns True if the selection object exists and has at least one Product selected.
'---------------------------------------------------------------
Public Function guardProductSelection() As Boolean
    guardProductSelection = False
    If Not guardProductDocument() Then Exit Function
    Dim sel As Selection
    Set sel = CATIA.ActiveDocument.Selection
    If sel Is Nothing Then Exit Function
    If sel.Count = 0 Then Exit Function
    Dim i As Integer
    For i = 1 To sel.Count
        If TypeName(sel.Item(i).Value) = "Product" Then
            guardProductSelection = True
            Exit Function
        End If
    Next i
End Function

'---------------------------------------------------------------
' Function: guardDesignMode
'   Returns True if the root product is in Design Mode (or can be set), otherwise False.
'---------------------------------------------------------------
Public Function guardDesignMode() As Boolean
    guardDesignMode = False
    If Not guardProductDocument() Then Exit Function
    On Error Resume Next
    CATIA.ActiveDocument.Product.ApplyWorkMode DESIGN_MODE
    guardDesignMode = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

'---------------------------------------------------------------
' Function: runAllGuards
'   Checks all preconditions for running CATIA macros safely.
'
'   Returns:
'     Boolean - True if all checks pass, False otherwise.
'   Side effects:
'     Shows a MsgBox and exits if any guard fails.
'---------------------------------------------------------------
Public Function runAllGuards() As Boolean
    runAllGuards = False

    If Not guardCatiaRunning() Then
        MsgBox "CATIA is not running.", vbExclamation, "Gunther's Catia Wizard"
        Exit Function
    End If

    If Not guardActiveDocument() Then
        MsgBox "No document is open in CATIA. Please open a Product document and try again.", vbExclamation, "Gunther's Catia Wizard"
        Exit Function
    End If

    If Not guardProductDocument() Then
        MsgBox "The active document is not a ProductDocument. Please activate a Product document and try again.", vbExclamation, "Gunther's Catia Wizard"
        Exit Function
    End If

    If Not guardProductSelection() Then
        MsgBox "No Product is selected. Please select at least one Product in the tree and try again.", vbExclamation, "Gunther's Catia Wizard"
        Exit Function
    End If

    If Not guardDesignMode() Then
        MsgBox "Could not set Design Mode on the root product.", vbExclamation, "Gunther's Catia Wizard"
        Exit Function
    End If

    ' Set globals
    Set prodDoc = CATIA.ActiveDocument
    Set rootProd = prodDoc.Product

    runAllGuards = True
End Function
===============================================================
MODULE: GunthersCatiaWizard.bas
PURPOSE: Main entry point and orchestrator for Gunther's Catia Wizard macro.
         Initializes globals, exposes entry point, and demonstrates usage.
         See Docs.bas for API documentation and usage examples.
===============================================================

Option Explicit

'===============================================================
' Global Variables (kept minimal)
'===============================================================
'--- [SUGGESTED MODULE: Globals.bas] ---
Public prodDoc As ProductDocument     ' Active ProductDocument
Public rootProd As Product            ' Root Product of the assembly


'===============================================================
' Entry Point (guards → init → UI dispatch only)
'===============================================================
Sub CATMain()

    If Not EnsureActiveProductDocument() Then Exit Sub

    ' Show the Launchpad UI for user to select and run tools
    Launchpad.Show

    ' Keep Main clean. See GunthersCatiaWizard_Docs for full examples & usage.

End Sub

'===============================================================
' Launchpad button handlers (called from form events)
'===============================================================
Public Sub Launchpad_Run()
    ' Guard: Ensure a valid ProductDocument is active before running any tool
    If Not EnsureActiveProductDocument() Then
        MsgBox "No valid CATProduct document is open. Please open a Product and try again.", vbExclamation, "Gunther's Catia Wizard"
        Exit Sub
    End If
    ' TODO: Dispatch selected tool based on UI (placeholder)
    ' All UI except errors should be handled in a form
    ' Unload Launchpad after running (optional)
    Unload Launchpad
End Sub

Public Sub Launchpad_Cancel()
    Unload Launchpad
    End ' Terminates macro execution
End Sub


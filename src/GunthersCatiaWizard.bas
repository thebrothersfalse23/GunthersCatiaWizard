'===============================================================
' MODULE: GunthersCatiaWizard.bas
' PURPOSE: Main entry point and orchestrator for Gunther's Catia Wizard macro.
'          Initializes globals, exposes entry point, and demonstrates usage.
'          See Docs.bas for API documentation and usage examples.
'===============================================================

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
    Launchpad.Show
    ' Keep Main clean. See Docs for full examples & usage.
End Sub

'===============================================================
' Launchpad button handlers (called from form events)
'===============================================================
Public Sub Launchpad_Run()
'Need to populate this. something went wrong and its not handling launchpad_run correctly. 
'this should run the safeguards.

End Sub

Public Sub Launchpad_Cancel()
    Unload Launchpad
    End ' Terminates macro execution
End Sub


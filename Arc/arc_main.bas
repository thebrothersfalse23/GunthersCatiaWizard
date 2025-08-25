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

End Sub

'===============================================================
' Launchpad button handlers (called from form events)
'===============================================================
' filepath: c:\Users\TheFa\OneDrive\Documents\GitHub\GunthersCatiaWizard\src\main.bas
' ...existing code...

' Show the interactive docs viewer form
Public Sub showDocsViewer()
    DocsViewer.Show
End Sub

' ...existing code...

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
    Begin VB.Label lblTitle
        Caption         =   "Gunther's Catia Wizard"
        Font.Size       =   16
        Font.Bold       =   -1  'True
        Height          =   480
        Left            =   0
        Top             =   120
        Width           =   6000
        Alignment       =   2  'Center
    End
    Begin VB.Label lblInstructions
        Caption         =   "Welcome! To use this wizard:" & vbCrLf & _
                                 "• Open a CATProduct document in CATIA." & vbCrLf & _
                                 "• Ensure the assembly is fully loaded (Design Mode recommended)." & vbCrLf & _
                                 "• Save your work before running tools." & vbCrLf & _
                                 "• Click 'Run' to begin or 'Cancel' to exit."
        Height          =   900
        Left            =   360
        Top             =   720
        Width           =   5280
        Alignment       =   1  'Right Justify
    End
    Begin VB.CommandButton btnRun
        Caption         =   "Run"
        Height          =   420
        Left            =   1800
        Top             =   2000
        Width           =   1000
        Enabled         =   True
    End
    Begin VB.CommandButton btnCancel
        Caption         =   "Cancel"
        Height          =   420
        Left            =   3200
        Top             =   2000
        Width           =   1000
        Enabled         =   True
    End
End

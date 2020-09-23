VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLoading 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading Saved..."
   ClientHeight    =   450
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgLoading 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

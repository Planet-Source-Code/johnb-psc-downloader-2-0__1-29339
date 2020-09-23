VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmViewSavedCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Saved Code"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frmViewSavedCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   3420
      Width           =   2415
   End
   Begin VB.CommandButton cmdSaveFile 
      Caption         =   "&Save code to file..."
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3420
      Width           =   2415
   End
   Begin VB.CommandButton cmdViewImage 
      Caption         =   "&View associated image..."
      Height          =   375
      Left            =   2460
      TabIndex        =   2
      Top             =   3420
      Width           =   2415
   End
   Begin RichTextLib.RichTextBox rtfDescription 
      Height          =   3375
      Left            =   3060
      TabIndex        =   1
      Top             =   0
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   5953
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmViewSavedCode.frx":0442
   End
   Begin VB.ListBox lstNames 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmViewSavedCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oList As colCodeList
Private oCurrRec As clsPSCdlRecord
Private lCurrListPos As Long
Public Starting As Boolean

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSaveFile_Click()
    Dim fSaveFile As frmSaveCodeFile
    
    Set fSaveFile = New frmSaveCodeFile
    
    Load fSaveFile
    Set fSaveFile.oPSCdlRecord = oCurrRec
    fSaveFile.Caption = fSaveFile.Caption & fSaveFile.oPSCdlRecord.sName & ".ZIP"
    fSaveFile.chkAboutFile.Enabled = oList.Item(lCurrListPos).bDescription
    fSaveFile.chkImageFile.Enabled = oList.Item(lCurrListPos).bImage
    
    fSaveFile.Show vbModal
    lstNames.ListIndex = lCurrListPos - 1
    
    Set fSaveFile = Nothing
End Sub

Private Sub cmdViewImage_Click()
    Dim fView As frmViewImage
    
    Set fView = New frmViewImage
    
    Load fView
    Set fView.oPSCdlRecord = oCurrRec
    fView.LoadPic
    
    fView.Show vbModal
    lstNames.ListIndex = lCurrListPos - 1
    
    Set fView = Nothing
End Sub

Private Sub Form_Load()
    Set oList = New colCodeList
    Starting = True
End Sub

Private Sub Form_Activate()
    If Starting Then
        lstNames.ListIndex = 0
        lCurrListPos = 1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oList = Nothing
End Sub

Private Sub lstNames_Click()
    Dim lIndex As Long
    
    Set oCurrRec = New clsPSCdlRecord
    
    rtfDescription.TextRTF = ""
    With oList
        For lIndex = 1 To .Count
            If .Item(lIndex).lListPos = lstNames.ListIndex + 1 Then
                Set oCurrRec = oRec.GetRecord(.Item(lIndex).lID, .Item(lIndex).sName)
                If .Item(lIndex).bDescription Then
                    rtfDescription.TextRTF = oCurrRec.sDesc
                Else
                    rtfDescription.TextRTF = "No Description Available"
                End If
                cmdViewImage.Enabled = .Item(lIndex).bImage
                cmdSaveFile.Enabled = .Item(lIndex).bZip
                lCurrListPos = lIndex
                Exit For
            End If
        Next lIndex
    End With
End Sub

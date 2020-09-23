VERSION 5.00
Begin VB.Form frmSaveCodeFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saving: "
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   Icon            =   "frmSaveCodeFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4035
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdChangePath 
      Caption         =   "..."
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   840
      Width           =   315
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   435
      Left            =   2040
      TabIndex        =   6
      Top             =   1500
      Width           =   1995
   End
   Begin VB.TextBox txtSavePath 
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   840
      Width           =   3615
   End
   Begin VB.CheckBox chkSavePath 
      Caption         =   "Save path as default"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   1140
      Width           =   1815
   End
   Begin VB.CheckBox chkImageFile 
      Caption         =   "Save image file"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   300
      Width           =   3915
   End
   Begin VB.CheckBox chkAboutFile 
      Caption         =   "Save ""About"" file"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   3915
   End
   Begin VB.CommandButton cmdSaveFile 
      Caption         =   "&Save "
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   1500
      Width           =   1935
   End
   Begin VB.Label lblCurrentSavePath 
      Caption         =   "Current save path:"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frmSaveCodeFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oPSCdlRecord As clsPSCdlRecord

Private Sub cmdChangePath_Click()
    txtSavePath.Text = BrowseForFolder(Me.hwnd, "Please select save folder:")
    If chkSavePath.Value = 1 Then oSettings.sPathSave = txtSavePath.Text
    oLog.Save nINFO, "Save.path: " & oSettings.sPathDB
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSaveFile_Click()
    Dim sFullPathAndName As String
    Dim lFile As Long
    
    sFullPathAndName = txtSavePath.Text & "\" & oPSCdlRecord.sName & "."
    
    If chkAboutFile.Value = 1 Then
        lFile = FreeFile
        Open sFullPathAndName & "TXT" For Output As lFile
        Print #lFile, oPSCdlRecord.sDesc
        Close lFile
    End If
    If chkImageFile.Value = 1 Then oBase64.DecodeToFile oPSCdlRecord.b64ImageFile, sFullPathAndName & oPSCdlRecord.GetImageExt(oPSCdlRecord.lImageType)
    oBase64.DecodeToFile oPSCdlRecord.b64ZipFile, sFullPathAndName & "ZIP"
    
    MsgBox oPSCdlRecord.sName & ".ZIP saved to:" & vbCrLf & txtSavePath.Text & "\", vbInformation & vbOKOnly, "File Saved"
End Sub

Private Sub Form_Load()
    Set oPSCdlRecord = New clsPSCdlRecord
    txtSavePath.Text = oSettings.sPathSave
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If chkSavePath.Value = 1 Then oSettings.sPathSave = txtSavePath.Text
    Set oPSCdlRecord = Nothing
End Sub

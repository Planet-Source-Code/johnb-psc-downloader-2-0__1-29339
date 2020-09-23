VERSION 5.00
Begin VB.Form frmViewImage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Image:"
   ClientHeight    =   690
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   1335
   Icon            =   "frmViewImage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   46
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   89
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picTemp 
      AutoSize        =   -1  'True
      Height          =   795
      Left            =   4680
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Menu mnuClose 
      Caption         =   "&Close"
   End
End
Attribute VB_Name = "frmViewImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oPSCdlRecord As clsPSCdlRecord
Private strTempName As String

Public Sub LoadPic()
    strTempName = oSettings.sPathTemp & "\DOWNLOAD_IMG." & oPSCdlRecord.GetImageExt(oPSCdlRecord.lImageType)
    
    oBase64.DecodeToFile oPSCdlRecord.b64ImageFile, strTempName
       
    picTemp.Picture = LoadPicture(strTempName)
    Me.Width = (picTemp.ScaleWidth * Screen.TwipsPerPixelX) + 92
    Me.Height = (picTemp.ScaleHeight * Screen.TwipsPerPixelY) + 640
    Me.Picture = picTemp.Picture
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oPSCdlRecord = New clsPSCdlRecord
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Kill strTempName
    Set oPSCdlRecord = Nothing
End Sub

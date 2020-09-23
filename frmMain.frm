VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   870
   ClientLeft      =   75
   ClientTop       =   -210
   ClientWidth     =   2550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Zip"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   2003
      TabIndex        =   3
      Top             =   600
      Width           =   435
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Image"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   1363
      TabIndex        =   2
      Top             =   600
      Width           =   435
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Desc"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   723
      TabIndex        =   1
      Top             =   600
      Width           =   435
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   83
      TabIndex        =   0
      Top             =   600
      Width           =   435
   End
   Begin VB.Image imgDownload 
      Height          =   480
      Index           =   3
      Left            =   1980
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgDownload 
      Height          =   480
      Index           =   2
      Left            =   1340
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgDownload 
      Height          =   480
      Index           =   1
      Left            =   700
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgDownload 
      Height          =   480
      Index           =   0
      Left            =   60
      OLEDropMode     =   1  'Manual
      Top             =   60
      Width           =   480
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuViewSaved 
         Caption         =   "View Saved Code"
      End
      Begin VB.Menu mnuSepTest 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "Help/About"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bMoveMe As Boolean
Private m_bShowLog As Boolean
Private m_lX As Long
Private m_lY As Long
Private m_lImageIndex As Long
Private m_bImg1OK As Boolean
Private m_bImg2OK As Boolean
Private m_bImg3OK As Boolean
Private m_bImg4OK As Boolean

Public Sub Reset()
    imgDownload_DblClick 0
End Sub

Private Sub cmd_Click(index As Integer)
    Dim OldPath As String
    Dim newPath As String
        
    Select Case index
        Case 0
            frmAbout.Show 1, Me
        Case 1
            oSettings.sPathDB = BrowseForFolder(Me.hwnd, "Please select download folder:")
            oLog.Save nINFO, "Dest.path: " & oSettings.sPathDB
        Case 2
            If m_bShowLog Then
                m_bShowLog = False
                Me.Height = 2020
            Else
                m_bShowLog = True
                Me.Height = 3800
            End If
        Case 3
            oLog.Save nINFO, "Exiting"
            Unload Me
            End
    End Select
End Sub

Private Sub Form_Load()
    With Me
        .Left = oSettings.lPosLeft
        .Top = oSettings.lPosTop
    End With
    putMeOnTop Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_bMoveMe = True
    m_lX = X
    m_lY = Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_bMoveMe = False
    If Button = vbRightButton Then
        PopupMenu mnuMain
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim MoveX As Integer
    Dim MoveY As Integer
    
    If m_bMoveMe Then
        MoveX = X - m_lX
        MoveY = Y - m_lY
        Me.Left = Me.Left + MoveX
        Me.Top = Me.Top + MoveY
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    With oSettings
        .lPosLeft = CLng(Me.Left)
        .lPosTop = CLng(Me.Top)
        .SaveSettings
    End With
    End
End Sub

Private Sub imgDownload_OLEDragDrop(index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim r As Boolean
    Dim s As String
    Dim f As String
    Dim sfile As String
    Dim fn As String
    Dim i As Integer
    Dim numFiles As Integer

    Select Case index
        Case 0
            If oRec.sName = "" And InStr(1, Trim(Data.GetData(vbCFText)), vbCrLf) = 0 Then
                s = Trim(Data.GetData(vbCFText))
                oRec.sName = ExtTrim(s, False, True, "_")
                Do While Asc(Right(oRec.sName, 1)) < 40
                    oRec.sName = Left(oRec.sName, Len(oRec.sName) - 1)
                Loop
                ChangeIcon nNAME, Me, False
                
                bDLSaved = False
                oLog.Save nDL_OK, "<" & oRec.sName & ">"
            End If
        Case 1
            'if an url/text is drag and dropped from browser

            If Data.GetFormat(vbCFText) Then
                m_lImageIndex = 2
                oRec.sDesc = Data.GetData(vbCFText)
                         
                ChangeIcon nDESC, Me, False
                
                oLog.Save nDL_OK, "Captured description"
            End If
        Case 2
            If oSettings.sPathDB = "" Then MsgBox "No destination path available..."
            If Data.GetFormat(vbCFFiles) Then
                m_lImageIndex = 3
                numFiles = Data.Files.Count
                For i = 1 To numFiles
                    If (GetAttr(Data.Files(i)) And vbDirectory) <> vbDirectory Then
                        sfile = Data.Files(i) 'Skip directory and get the first file
                        Exit For
                    Else
                        sfile = Data.Files(i) 'Skip directory and get the first file
                        Exit For
                    End If
                Next i
                If oRec.sName = "" Then
                    fn = sfile
                Else
                    fn = oRec.sName & Right(sfile, 4)
                    oRec.lImageType = oRec.GetImageType(Right(sfile, 3))
                End If

                r = DL(sfile, oSettings.sPathDB & "\" & fn)
                
                If r Then
                    oRec.b64ImageFile = oBase64.EncodeFromFile(oSettings.sPathDB & "\" & fn)
                    Kill oSettings.sPathDB & "\" & fn
                    ChangeIcon nIMAGE, Me, False
                    oLog.Save nDL_OK, "Captured image file"
                Else
                    ChangeIcon nIMAGE, Me, True
                    oLog.Save nERROR, "Failed - image file"
                End If
                
            End If
        Case 3
            'if an url/text is drag and dropped from browser
            If oSettings.sPathDB = "" Then MsgBox "No destination path available..."
            If Data.GetFormat(vbCFText) Then
                m_lImageIndex = 4

                sfile = Data.GetData(vbCFText)

                If oRec.sName = "" Then
                    fn = sfile
                Else
                    fn = oRec.sName & ".zip"
                End If
                r = DL(sfile, oSettings.sPathDB & "\" & fn)

                If r Then
                    oRec.b64ZipFile = oBase64.EncodeFromFile(oSettings.sPathDB & "\" & fn)
                    Kill oSettings.sPathDB & "\" & fn
                    ChangeIcon nZIP, Me, False
                    oLog.Save nDL_OK, "Captured zip file"
                Else
                    ChangeIcon nZIP, Me, True
                    oLog.Save nERROR, "Failed - zip file"
                End If
            End If
    End Select
End Sub

Private Sub imgDownload_DblClick(index As Integer)
    If index = 0 Then
        If oRec.Save Then
            oRec.Reset
            oRec.sName = ""
            LoadIcons Me
            bDLSaved = True
        Else
            MsgBox "Error saving download" & vbCrLf & "Exiting.", vbCritical + vbOKOnly, "PSCdl - Download Error"
            Unload Me
        End If
    End If
End Sub

Private Sub imgDownload_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        m_bMoveMe = True
        m_lX = X
        m_lY = Y
    End If
End Sub

Private Sub imgDownload_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveFrm X, Y
End Sub

Private Sub imgDownload_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_bMoveMe = False
    If Button = vbRightButton Then
        Form_MouseUp vbRightButton, 0, X, Y
    End If
End Sub

Private Sub lblCaption_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_bMoveMe = True
    m_lX = X
    m_lY = Y
End Sub

Private Sub lblCaption_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveFrm X, Y
End Sub

Private Sub lblCaption_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_bMoveMe = False
    If Button = vbRightButton Then
        Form_MouseUp vbRightButton, 0, X, Y
    End If
End Sub

Private Sub MoveFrm(ByVal X As Integer, ByVal Y As Integer)
    Dim MoveX As Integer
    Dim MoveY As Integer
    
    If m_bMoveMe Then
        MoveX = X - m_lX
        MoveY = Y - m_lY
        Me.Left = Me.Left + MoveX
        Me.Top = Me.Top + MoveY
    End If
End Sub

Private Sub mnuExit_Click()
    If Not bDLSaved Then
        If MsgBox("The last download was not saved." & vbCrLf & vbCrLf & "Do you want to save it now?", vbExclamation + vbYesNo, "PSCdl - Last download") = vbYes Then
            oRec.Save
        End If
    End If
    Unload Me
End Sub

Private Sub mnuSettings_Click()
    ShowSettings
End Sub

Private Sub mnuViewSaved_Click()
    ShowSaved
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Upload files to db after editting constant and moving files."
      Height          =   495
      Left            =   660
      TabIndex        =   1
      Top             =   60
      Width           =   2955
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   660
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==============================
' Run PSCdl First!
' Registry settings will not be set otherwise.
'==============================

' Change this to a folder under where the database is located.
' \ <- is required on both sides!
' Put all current PSCdl downloads in here, along with any other downloads you wnat in the database.
Private Const FIXLOAD As String = "\Extracted\"

Private Sub Command2_Click()
    Dim oFileSys As New Scripting.FileSystemObject
    Dim oFolder As Scripting.Folder
    Dim oFile As Scripting.File
    Dim oBase64 As New Base64Lib.Base64
    Dim oADO As New AdoUtils
    Dim oRec As ADODB.Recordset
    Dim oSet As New clsSettings
    Dim oPSCdlRec As New clsPSCdlRecord
    Dim oFileCol As New colFiles
    Dim lIndex As Long
    Dim sFixedName As String
    
    oSet.GetSettings
    oADO.ConnectionString = oSet.sDataConnect
    Set oFolder = oFileSys.GetFolder(oSet.sPathDB & FIXLOAD)
    
    For Each oFile In oFolder.Files
        oFileCol.Add oFile.name
    Next
    
    Debug.Print "Files count: " & oFileCol.Count
    
    ProgressBar1.Max = oFileCol.Count

    For lIndex = 1 To oFileCol.Count
        Set oRec = New ADODB.Recordset
        ProgressBar1.Value = lIndex
        
            sFixedName = FixName(oFileCol.Item(lIndex).sFileName)
            Set oRec = oADO.GetRS("SELECT * FROM PSCdl_Table WHERE sFileName='" & Mid(sFixedName, 1, InStr(sFixedName, ".") - 1) & "'")
            With oRec
                If oADO.EmptyRS(oRec) Then
                    .AddNew
                    !sFileName = Mid(sFixedName, 1, InStr(sFixedName, ".") - 1)
                End If
                
                If Len(!b64Download) < 1 Or IsNull(!b64Download) Then !b64Download = ""
                If Len(!b64Image) < 1 Or IsNull(!b64Image) Then !b64Image = ""
                If Len(!sDescription) < 1 Or IsNull(!sDescription) Then !sDescription = ""
                If Len(!sImageType) < 1 Or IsNull(!sImageType) Then !sImageType = ""
                
                If UCase(Right(sFixedName, 3)) = "ZIP" Then !b64Download = oBase64.EncodeFromFile(oSet.sPathDB & FIXLOAD & sFixedName)
                If UCase(Right(sFixedName, 3)) = "TXT" Then !sDescription = ReadAbout(oSet.sPathDB & FIXLOAD & oFileCol.Item(lIndex).sFileName)
                If oPSCdlRec.GetImageType(Right(sFixedName, 3)) <> NONE Then
                   !b64Image = oBase64.EncodeFromFile(oSet.sPathDB & FIXLOAD & sFixedName)
                   !sImageType = UCase(Right(sFixedName, 3))
                End If
            End With
            Debug.Print oADO.PutRS(oRec)
    Next lIndex
    
    Set oFileSys = Nothing
    Set oFolder = Nothing
    Set oFile = Nothing
    Set oBase64 = Nothing
    Set oADO = Nothing
    Set oSet = Nothing
    Set oPSCdlRec = Nothing
    Set oFileCol = Nothing
End Sub

Private Function ReadAbout(ByVal sFileName As String) As String
    Dim oFileSys As New Scripting.FileSystemObject
    Dim oTS As Scripting.TextStream
    Set oFileSys = New Scripting.FileSystemObject
    Set oTS = oFileSys.OpenTextFile(sFileName, ForReading)
    
    ReadAbout = oTS.ReadAll
    oTS.Close
    
    Set oTS = Nothing
    Set oFileSys = Nothing
End Function

Private Function FixName(ByVal sName As String) As String
    FixName = sName
    If InStr(UCase(FixName), "_ABOUT") > 0 Then FixName = Mid(FixName, 1, InStr(UCase(FixName), "_ABOUT") - 1) & Right(FixName, 4)
End Function

Attribute VB_Name = "modMain"
Option Explicit

Public oSettings As clsSettings
Public oLog As clsLog
Public oBase64 As Base64Lib.Base64
Public oRec As clsPSCdlRecord
Public bDLSaved As Boolean

Private lGotAll As Long
Private Const lTOTAL_TO_GET As Long = 3

Public Enum enumStage
    nNAME = 1
    nDESC = 2
    nIMAGE = 3
    nZIP = 4
End Enum

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer

Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const INTERNET_FLAG_RELOAD = &H80000000

Public Sub Main()
    Dim fMain As New frmMain
        
    bDLSaved = True
    
    Set oSettings = New clsSettings
    If Not oSettings.GetSettings() Then MsgBox "Error"
    
    Set oLog = New clsLog
    oLog.bEnabled = oSettings.bLogging
    
    Load fMain
    LoadIcons fMain
       
    Set oBase64 = New Base64Lib.Base64
    Set oRec = New clsPSCdlRecord
    oRec.ConnectString = oSettings.sDataConnect
    
    fMain.Show
End Sub

Public Sub LoadIcons(fMain As frmMain)
    With fMain
        Set .imgDownload(0).Picture = LoadResPicture("open_wait_arrow", vbResIcon)
        Set .imgDownload(1).Picture = LoadResPicture("closed_wait", vbResIcon)
        Set .imgDownload(2).Picture = LoadResPicture("closed_wait", vbResIcon)
        Set .imgDownload(3).Picture = LoadResPicture("closed_wait", vbResIcon)
        .imgDownload(0).OLEDropMode = vbOLEDropManual
        .imgDownload(1).OLEDropMode = vbOLEDropNone
        .imgDownload(2).OLEDropMode = vbOLEDropNone
        .imgDownload(3).OLEDropMode = vbOLEDropNone
    End With
    lGotAll = 0
End Sub

Public Sub ChangeIcon(ByVal Stage As enumStage, fMain As frmMain, ByVal bError As Boolean)
    Dim lIndex As Long
    With fMain
        Select Case Stage
            Case nNAME
                If Not bError Then
                    Set .imgDownload(0).Picture = LoadResPicture("closed_ok", vbResIcon)
                    For lIndex = 1 To 3
                        Set .imgDownload(lIndex).Picture = LoadResPicture("open_wait", vbResIcon)
                        .imgDownload(lIndex).OLEDropMode = vbOLEDropManual
                    Next lIndex
                Else
                    Set .imgDownload(0).Picture = LoadResPicture("open_error", vbResIcon)
                End If
            Case nDESC
                If Not bError Then
                    Set .imgDownload(1).Picture = LoadResPicture("closed_ok", vbResIcon)
                    If CheckAll(True) Then .Reset
                Else
                    Set .imgDownload(1).Picture = LoadResPicture("open_error", vbResIcon)
                End If
            Case nIMAGE
                If Not bError Then
                    Set .imgDownload(2).Picture = LoadResPicture("closed_ok", vbResIcon)
                    If CheckAll(True) Then .Reset
                Else
                    Set .imgDownload(2).Picture = LoadResPicture("open_error", vbResIcon)
                End If
            Case nZIP
                If Not bError Then
                    Set .imgDownload(3).Picture = LoadResPicture("closed_ok", vbResIcon)
                    If CheckAll(True) Then .Reset
                Else
                    Set .imgDownload(3).Picture = LoadResPicture("open_error", vbResIcon)
                End If
        End Select
    End With
End Sub

Private Function CheckAll(ByVal bAdd As Boolean) As Boolean
    CheckAll = False
    If bAdd Then lGotAll = lGotAll + 1
    If lGotAll = lTOTAL_TO_GET Then CheckAll = True
End Function

Public Function DL(ByVal SourceFile As String, ByVal DestinationFile As String) As Boolean
    Screen.MousePointer = vbHourglass
    DL = Get_File(SourceFile, DestinationFile)
    Screen.MousePointer = vbDefault
End Function

Public Function Get_File(sURLFileName As String, sSaveFileName As String) As Boolean
    Dim lRet As Long
    On Error GoTo err_Fix
  
    lRet = InternetOpen("", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    lRet = URLDownloadToFile(0, sURLFileName, sSaveFileName, 0, 0)
    Get_File = True
    Exit Function
err_Fix:
    Debug.Print Err.LastDllError, lRet
    Err.Clear
    Get_File = False
End Function

Public Sub putMeOnTop(Form As Form)
    SetWindowPos Form.hwnd, -1, 0, 0, 0, 0, 1 Or 2
End Sub

Public Sub takeMeDown(Form As Form)
    SetWindowPos Form.hwnd, -2, 0, 0, 0, 0, 1 Or 2
End Sub

Public Sub ShowSaved()
    Dim fSaved As frmViewSavedCode
    Dim fLoad As frmLoading
    
    Set fSaved = New frmViewSavedCode
    Set fLoad = New frmLoading
    
    fLoad.Show
    fLoad.MousePointer = vbHourglass
    fLoad.Refresh
         
    Load fSaved
    InitSaved fSaved, fLoad
    fSaved.Starting = False
    
    fLoad.MousePointer = vbNormal
    fLoad.Hide
    Set fLoad = Nothing
    
    fSaved.Show vbModal
    
    Set fSaved = Nothing
End Sub

Public Sub ShowSettings()
    Dim fSet As frmOptions
    
    Set fSet = New frmOptions
    
    Load fSet
    fSet.Show vbModal
    
    Set fSet = Nothing
End Sub

Public Sub InitSaved(ByVal fLoaded As frmViewSavedCode, ByVal fLoading As frmLoading)
    Dim lIndex As Long
    
    Set fLoaded.oList = oRec.GetList()
    fLoading.prgLoading.Max = fLoaded.oList.Count
    
    With fLoaded.oList
        If .Count > 0 Then
            For lIndex = 1 To .Count
                fLoaded.lstNames.AddItem Replace(.Item(lIndex).sName, "_", " ")
                .Item(lIndex).lListPos = lIndex
                fLoading.prgLoading.Value = lIndex
                fLoading.prgLoading.Refresh
                fLoading.Refresh
            Next lIndex
        Else
            MsgBox "No saved code to view!", vbInformation + vbOKOnly, "View Saved Code"
        End If
    End With
End Sub

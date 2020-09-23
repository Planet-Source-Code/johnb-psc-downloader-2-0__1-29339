VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const sBASE_REG_KEY As String = "Software\PSCdl"

Private oReg As Registry

Private mvarlPosLeft As Long
Private mvarlPosTop As Long
Private mvarsDataConnect As String
Private mvarsPathDB As String
Private mvarsPathTemp As String
Private mvarsPathSave As String
Private mvarbLogging As Boolean

Public Property Let bLogging(ByVal vData As Boolean)
    mvarbLogging = vData
End Property

Public Property Get bLogging() As Boolean
    bLogging = mvarbLogging
End Property

Public Property Let sPathDB(ByVal vData As String)
    mvarsPathDB = vData
End Property

Public Property Get sPathDB() As String
    sPathDB = mvarsPathDB
End Property

Public Property Let sPathTemp(ByVal vData As String)
    mvarsPathTemp = vData
End Property

Public Property Get sPathTemp() As String
    sPathTemp = mvarsPathTemp
End Property

Public Property Let sPathSave(ByVal vData As String)
    mvarsPathSave = vData
End Property

Public Property Get sPathSave() As String
    sPathSave = mvarsPathSave
End Property

Public Property Let sDataConnect(ByVal vData As String)
    mvarsDataConnect = vData
End Property

Public Property Get sDataConnect() As String
    sDataConnect = mvarsDataConnect
End Property

Public Property Let lPosTop(ByVal vData As Long)
    mvarlPosTop = vData
End Property

Public Property Get lPosTop() As Long
    lPosTop = mvarlPosTop
End Property

Public Property Let lPosLeft(ByVal vData As Long)
    mvarlPosLeft = vData
End Property

Public Property Get lPosLeft() As Long
    lPosLeft = mvarlPosLeft
End Property

Public Function SaveSettings() As Boolean
    oReg.SetRegistryValue HKEY_CURRENT_USER, sBASE_REG_KEY & "\Data", "Connect", mvarsDataConnect
    oReg.SetRegistryValue HKEY_CURRENT_USER, sBASE_REG_KEY & "\Position", "Left", mvarlPosLeft
    oReg.SetRegistryValue HKEY_CURRENT_USER, sBASE_REG_KEY & "\Position", "Top", mvarlPosTop
    oReg.SetRegistryValue HKEY_CURRENT_USER, sBASE_REG_KEY & "\Paths", "DB", mvarsPathDB
    oReg.SetRegistryValue HKEY_CURRENT_USER, sBASE_REG_KEY & "\Paths", "Save", mvarsPathSave
    oReg.SetRegistryValue HKEY_CURRENT_USER, sBASE_REG_KEY & "\Paths", "Temp", mvarsPathTemp
    oReg.SetRegistryValue HKEY_CURRENT_USER, sBASE_REG_KEY & "\Logging", "0", mvarbLogging
End Function

Public Function GetSettings() As Boolean
    GetSettings = True
    If oReg.CheckRegistryKey(HKEY_CURRENT_USER, sBASE_REG_KEY) Then
        mvarsDataConnect = oReg.GetRegistryValue(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Data", "Connect")
        mvarlPosLeft = oReg.GetRegistryValue(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Position", "Left")
        mvarlPosTop = oReg.GetRegistryValue(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Position", "Top")
        mvarsPathDB = oReg.GetRegistryValue(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Paths", "DB")
        mvarsPathSave = oReg.GetRegistryValue(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Paths", "Save")
        mvarsPathTemp = oReg.GetRegistryValue(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Paths", "Temp")
        mvarbLogging = oReg.GetRegistryValue(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Logging", "0")
    Else
        GetSettings = oReg.CreateRegistryKey(HKEY_CURRENT_USER, sBASE_REG_KEY)
        GetSettings = oReg.CreateRegistryKey(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Data")
        GetSettings = oReg.CreateRegistryKey(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Position")
        GetSettings = oReg.CreateRegistryKey(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Paths")
        GetSettings = oReg.CreateRegistryKey(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Logging")
        GetSettings = oReg.SetRegistryValue(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Data", "Connect", mvarsDataConnect)
        GetSettings = oReg.SetRegistryValue(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Position", "Left", mvarlPosLeft)
        GetSettings = oReg.SetRegistryValue(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Position", "Top", mvarlPosTop)
        GetSettings = oReg.SetRegistryValue(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Paths", "DB", mvarsPathDB)
        GetSettings = oReg.SetRegistryValue(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Paths", "Save", mvarsPathSave)
        GetSettings = oReg.SetRegistryValue(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Paths", "Temp", mvarsPathTemp)
        GetSettings = oReg.SetRegistryValue(HKEY_CURRENT_USER, sBASE_REG_KEY & "\Logging", "0", mvarbLogging)
    End If
End Function

Private Sub Class_Initialize()
    Set oReg = New Registry
    mvarsDataConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PSCdl.mdb;Persist Security Info=False"
    mvarsPathDB = App.Path
    mvarsPathSave = App.Path
    mvarsPathTemp = App.Path & "\Download_Temp"
    mvarlPosLeft = 700
    mvarlPosTop = 1500
    mvarbLogging = False
End Sub

Private Sub Class_Terminate()
    Set oReg = Nothing
End Sub

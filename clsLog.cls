VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarsType As String
Private mvarsDescription As String
Private mvarlID As Long
Private mvarbEnabled As Boolean

Public Enum enumLogType
    nERROR = 1
    nDL_OK = 2
    nINFO = 3
End Enum

Public Function Reset() As Boolean
    If mvarbEnabled Then
        mvarsType = ""
        mvarsDescription = ""
        mvarlID = vbNull
    End If
End Function

Public Function GetLog() As Boolean
    ' Not implemented yet
End Function

Public Function Save(ByVal lLogType As enumLogType, ByVal sDescription As String) As Boolean
    If mvarbEnabled Then
        Select Case lLogType
            Case nERROR: mvarsType = "ERROR"
            Case nDL_OK: mvarsType = "DL_OK"
            Case nINFO: mvarsType = "INFO"
        End Select
        mvarsDescription = sDescription
        ' Stick in DB here
        ' Get new ID here
        ' Need an event to say pick up the entry for display.
    End If
End Function

Public Property Let bEnabled(ByVal vData As Boolean)
    mvarbEnabled = vData
End Property

Public Property Get bEnabled() As Boolean
    bEnabled = mvarbEnabled
End Property

Public Property Let lID(ByVal vData As Long)
    mvarlID = vData
End Property

Public Property Get lID() As Long
    lID = mvarlID
End Property

Public Property Let sDescription(ByVal vData As String)
    mvarsDescription = vData
End Property

Public Property Get sDescription() As String
    sDescription = mvarsDescription
End Property

Public Property Let sType(ByVal vData As String)
    mvarsType = vData
End Property

Public Property Get sType() As String
    sType = mvarsType
End Property

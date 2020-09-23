VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colCodeList"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarlID As Long
Private mvarsName As String
Private mvarbImage As Boolean
Private mvarbDescription As Boolean
Private mvarbZip As Boolean
Private mvarlListPos As Long

Public Property Let lListPos(ByVal vData As Long)
    mvarlListPos = vData
End Property

Public Property Get lListPos() As Long
    lListPos = mvarlListPos
End Property

Public Property Let bZip(ByVal vData As Boolean)
    mvarbZip = vData
End Property

Public Property Get bZip() As Boolean
    bZip = mvarbZip
End Property

Public Property Let bDescription(ByVal vData As Boolean)
    mvarbDescription = vData
End Property

Public Property Get bDescription() As Boolean
    bDescription = mvarbDescription
End Property

Public Property Let bImage(ByVal vData As Boolean)
    mvarbImage = vData
End Property

Public Property Get bImage() As Boolean
    bImage = mvarbImage
End Property

Public Property Let sName(ByVal vData As String)
    mvarsName = vData
End Property

Public Property Get sName() As String
    sName = mvarsName
End Property

Public Property Let lID(ByVal vData As Long)
    mvarlID = vData
End Property

Public Property Get lID() As Long
    lID = mvarlID
End Property




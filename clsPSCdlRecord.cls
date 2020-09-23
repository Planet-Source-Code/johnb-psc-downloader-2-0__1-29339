VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPSCdlRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private oADO As AdoUtils

Private mvarsName As String
Private mvarsDesc As String
Private mvarsImageName As String
Private mvarsImagePath As String
Private mvarsZipName As String
Private mvarsZipPath As String
Private mvarID As Long
Private mvarb64ZipFile As String
Private mvarb64ImageFile As String
Private mvarsImageType As String

Public Enum enum_ImageType
    NONE = 0
    BMP = 1
    JPG = 2
    GIF = 3
    PNG = 4
    TIF = 5
    PCX = 6
End Enum

Public Function GetImageExt(ByVal lImageType As enum_ImageType) As String
    Select Case lImageType
        Case BMP
            GetImageExt = "BMP"
        Case JPG
            GetImageExt = "JPG"
        Case GIF
            GetImageExt = "GIF"
        Case PNG
            GetImageExt = "PNG"
        Case TIF
            GetImageExt = "TIF"
        Case PCX
            GetImageExt = "PCX"
    End Select
End Function

Public Function GetImageType(ByVal sImageType As String) As enum_ImageType
    Select Case UCase(sImageType)
        Case "BMP"
            GetImageType = BMP
        Case "JPG"
            GetImageType = JPG
        Case "GIF"
            GetImageType = GIF
        Case "PNG"
            GetImageType = PNG
        Case "TIF"
            GetImageType = TIF
        Case "PCX"
            GetImageType = PCX
        Case Else
            GetImageType = NONE
    End Select
End Function

Public Function Reset() As Boolean
    mvarsName = ""
    mvarsDesc = ""
    mvarsImageName = ""
    mvarsImagePath = ""
    mvarsImageType = ""
    mvarsZipName = ""
    mvarsZipPath = ""
    mvarID = vbNull
    mvarb64ZipFile = ""
    mvarb64ImageFile = ""
End Function

Public Function GetRecord(ByVal lID As Long, ByVal sFileName As String) As clsPSCdlRecord
    Dim oTempRec As New clsPSCdlRecord
    Dim oADORec As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT * FROM PSCdl_Table WHERE ID=" & lID & " AND sFileName='" & sFileName & "'"
    
    Set oADORec = oADO.GetRS(sSQL)
    With oADORec
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                oTempRec.ID = !ID
                oTempRec.sName = !sFileName
                oTempRec.sDesc = !sDescription
                oTempRec.b64ImageFile = !b64Image
                oTempRec.lImageType = GetImageType(!sImageType)
                oTempRec.b64ZipFile = !b64Download
                .MoveNext
            Loop
        End If
    End With
    
    Set GetRecord = oTempRec
    
    Set oTempRec = Nothing
    Set oADORec = Nothing
End Function

Public Function GetList() As colCodeList
    Dim oList As New colCodeList
    Dim oADORec As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT * FROM PSCdl_Table ORDER BY ID"
    
    Set oADORec = oADO.GetRS(sSQL)
    
    With oADORec
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                oList.Add !ID, !sFileName, _
                    IIf(Len(!sDescription) > 0, True, False), _
                    IIf(Len(!b64Image) > 0, True, False), _
                    IIf(Len(!b64Download) > 0, True, False), 0, CStr(!ID) & !sFileName
                .MoveNext
            Loop
        End If
    End With
    
    Set GetList = oList
    
    Set oList = Nothing
    Set oADORec = Nothing
End Function

Public Function Save() As Boolean
    Dim sSQL As String
    Dim oADORec As New ADODB.Recordset
    
'    sSQL = "INSERT INTO PSCdl_Table(sFileName,sDescription,b64Image,sImageType,b64Download) "
'    sSQL = sSQL & "VALUES('" & mvarsName & "','" & mvarsDesc & "','" & mvarb64ImageFile & "','" & mvarsImageType & "','" & mvarb64ZipFile & "')"
'    Save = oADO.Execute(sSQL)
    
    sSQL = "SELECT * FROM PSCdl_Table WHERE ID < 0 "
    Set oADORec = oADO.GetRS(sSQL)
    
    If oADO.EmptyRS(oADORec) Then
        With oADORec
            .AddNew
            !sFileName = mvarsName
            !sDescription = mvarsDesc
            !b64Image = mvarb64ImageFile
            !sImageType = mvarsImageType
            !b64Download = mvarb64ZipFile
            oADO.PutRS oADORec
        End With
    End If
    
    Set oADORec = Nothing
End Function

Public Property Let lImageType(ByVal vData As enum_ImageType)
    Select Case vData
        Case BMP
            mvarsImageType = "BMP"
        Case JPG
            mvarsImageType = "JPG"
        Case GIF
            mvarsImageType = "GIF"
        Case PNG
            mvarsImageType = "PNG"
        Case TIF
            mvarsImageType = "TIF"
        Case PCX
            mvarsImageType = "PCX"
        Case Else
            mvarsImageType = "ERR"
    End Select
End Property

Public Property Get lImageType() As enum_ImageType
    Select Case mvarsImageType
        Case "BMP"
            lImageType = BMP
        Case "JPG"
            lImageType = JPG
        Case "GIF"
            lImageType = GIF
        Case "PNG"
            lImageType = PNG
        Case "TIF"
            lImageType = TIF
        Case "PCX"
            lImageType = PCX
        Case Else
            lImageType = NONE
    End Select
End Property

Public Property Let ConnectString(ByVal vData As String)
    oADO.ConnectionString = vData
End Property

Public Property Get ConnectString() As String
    ConnectString = oADO.ConnectionString
End Property

Public Property Let ID(ByVal vData As Long)
    mvarID = vData
End Property

Public Property Get ID() As Long
    ID = mvarID
End Property

Public Property Let sZipPath(ByVal vData As String)
    mvarsZipPath = vData
End Property

Public Property Get sZipPath() As String
    sZipPath = mvarsZipPath
End Property

Public Property Let sZipName(ByVal vData As String)
    mvarsZipName = vData
End Property

Public Property Get sZipName() As String
    sZipName = mvarsZipName
End Property

Public Property Let sImagePath(ByVal vData As String)
    mvarsImagePath = vData
End Property

Public Property Get sImagePath() As String
    sImagePath = mvarsImagePath
End Property

Public Property Let sImageName(ByVal vData As String)
    mvarsImageName = vData
End Property

Public Property Get sImageName() As String
    sImageName = mvarsImageName
End Property

Public Property Let sDesc(ByVal vData As String)
    mvarsDesc = Replace(vData, "'", "&quot;")
End Property

Public Property Get sDesc() As String
    sDesc = Replace(mvarsDesc, "&quot;", "'")
End Property

Public Property Let sName(ByVal vData As String)
    mvarsName = vData
End Property

Public Property Get sName() As String
    sName = mvarsName
End Property

Public Property Let b64ImageFile(ByVal vData As String)
    mvarb64ImageFile = vData
End Property

Public Property Get b64ImageFile() As String
    b64ImageFile = mvarb64ImageFile
End Property

Public Property Let b64ZipFile(ByVal vData As String)
    mvarb64ZipFile = vData
End Property

Public Property Get b64ZipFile() As String
    b64ZipFile = mvarb64ZipFile
End Property

Private Sub Class_Initialize()
    Set oADO = New AdoUtils
End Sub

Private Sub Class_Terminate()
    Set oADO = Nothing
End Sub

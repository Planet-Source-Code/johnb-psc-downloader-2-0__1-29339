VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdoUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'INCLUDE A REFERENCE TO MICROSOFT ACTIVE DATA OBJECTS
'IN ORDER TO USE THIS CLASS
Option Explicit

Private m_sConnectionString As String
Private m_sLastError As String
Private m_bIsSQL As Boolean
Private m_sDBPath As String
Private m_DBType As enumDBType 'local copy

Public Enum enumDBType
    ' Currently only Access 97/2000
    Access2000 = 0
    Access97
    Other
End Enum

Public Property Let DBType(ByVal vData As enumDBType)
    m_DBType = vData
    Select Case m_DBType
        Case Access2000
        Case Access97
        Case Other
    End Select
End Property

Public Property Get DBType() As enumDBType
    DBType = m_DBType
End Property

Public Property Get ConnectionString() As String
    ConnectionString = m_sConnectionString
End Property

Public Property Let ConnectionString(ByVal NewValue As String)
    'for some of the functions in this class
    'you will need to first
    'set this property to a
    'valid connection string
    
    'or reset it when you need to change the connection
    m_sConnectionString = NewValue
End Property

Public Property Get isSQL() As Boolean
    isSQL = m_bIsSQL
End Property

Public Property Let isSQL(ByVal NewValue As Boolean)
    'SET TO TRUE IF YOU ARE USING SQL SERVER
    'DEFAULT IS ACCESS
   m_bIsSQL = NewValue
End Property

Public Function Clone(ByVal objRecordset As ADODB.Recordset, Optional ByVal LockType As ADODB.LockTypeEnum = adLockBatchOptimistic) As ADODB.Recordset
    'RETURNS A CLONE (COPY OF AN EXISTING RECORDSET)
        
    Dim objNewRS As ADODB.Recordset
    Dim objField As Object
    Dim lngCnt As Long
    On Error GoTo LocalError
    
    Set objNewRS = New ADODB.Recordset
    objNewRS.CursorLocation = adUseClient
    objNewRS.LockType = LockType

    For Each objField In objRecordset.Fields
            objNewRS.Fields.Append objField.name, objField.Type, objField.DefinedSize, objField.Attributes
    Next objField

    If Not objRecordset.RecordCount = 0 Then
            Set objNewRS.ActiveConnection = objRecordset.ActiveConnection
            objNewRS.Open
          
        objRecordset.MoveFirst
        While Not objRecordset.EOF
              objNewRS.AddNew
            For lngCnt = 0 To objRecordset.Fields.Count - 1
                objNewRS.Fields(lngCnt).Value = objRecordset.Fields(lngCnt).Value
            Next lngCnt
            objRecordset.MoveNext
        Wend
    objNewRS.MoveFirst
    End If
    
    Set Clone = objNewRS
    Exit Function
LocalError:
    m_sLastError = Err.Number & " - " & Err.Description
    If objNewRS.State = adStateOpen Then
        objNewRS.Close
    End If
    Set objNewRS = Nothing
End Function

Function Datashape(ByVal tblParent As String, _
                   ByVal tblChild As String, _
                   ByVal fldParent As String, _
                   ByVal fldChild As String, _
                   Optional ordParent As String = "", _
                   Optional ordChild As String = "", _
                   Optional WhereParent As String = "", _
                   Optional WhereChild As String = "", _
                   Optional ParentFields As String = "*", _
                   Optional ChildFields As String = "*", _
                   Optional MaxRecords As Long = 0) As ADODB.Recordset
    '=========================================================
    'This function will return a DisConnected SHAPEd RecordSet
    'Assumptions:
    '
    'tblParent      = Valid Table in the Database   - String \ Parent Table
    'tblChild       = Valid Table in the Database   - String / Child  Table
    '
    'fldParent      = Valid Field in Parent Table   - String \ relate this field
    'fldChild       = Valid Field in Child Table    - String / ..to this field
    '
    'ordParent      = Valid Field in Parent Table   - String \ ordering
    'ordChild       = Valid Field in Child Table    - String /
    '
    'WhereParent    = Valid SQL Where Clause        - Variant (Optional)
    'WhereChild     = Valid SQL Where Clause        - Variant (Optional)
    '
    'ParentFields   = Specific Fields to return     - String (pipe delimitered)
    'ChildFields    = Specific Fields to return     - String (pipe delimitered)
    'MaxRecords     = Specify Maximum Child Records - Long (0 = ALL)
    
    'NOTE: You may have to change connection string:  Normal Connection Strings
    'Begin with "Provider=". For the MsDataShape Provider, the connection string
    'begins with "Data Provider = "
    
    'EXAMPLE: THIS RETURNS A HYPOTHETICAL RECORDSET OF CUSTOMERS,
    'WHERE ONE OF THE MEMBERS IS A HYPOTHETICAL CHILD RECORDSET
    'OF THE CUSTOMERS' ORDERS
    
    'Dim sShapeConnectionString As String
    'Dim oCustRs As ADODB.Recordset
    'Dim oOrderRs As ADODB.Recordset
    'Dim oADO As New AdoUtils
    'Dim sSQL As String

    'sShapeConnectionString = "Data Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyBusiness.mdb"
    'sSQL = "SELECT * FROM CUSTOMERS"
    'With oTest
    '   .ConnectionString = sShapeConnectionString
    '   Set oCustRs = .Datashape("Customers", "Orders", "ID", "CustomerID")
    '   Set oOrderRs = ors.Fields(ors.Fields.Count - 1).Value
    'End With
    
    
    '=========================================================
    On Error GoTo ErrHandler

    Dim cn        As ADODB.Connection
    Dim rS        As ADODB.Recordset
    Dim lSQL      As String
    Dim pSQL      As String
    Dim cSQL      As String
    Dim pWhere    As String
    Dim cWhere    As String
    Dim pOrder    As String
    Dim cOrder    As String

    'Define the SQL Statement
    lSQL = ""
    ParentFields = Replace(ParentFields, "|", ", ")
    ChildFields = Replace(ChildFields, "|", ", ")
    pWhere = WhereParent
    cWhere = WhereChild
    pOrder = ordParent
    cOrder = ordChild

    If WhereParent <> "" Then WhereParent = " WHERE " & WhereParent
    If WhereChild <> "" Then WhereChild = " WHERE " & WhereChild
    If pOrder <> "" Then pOrder = " ORDER By " & pOrder
    If cOrder <> "" Then cOrder = " ORDER By " & cOrder
    'Define Parent SQL Statement
    pSQL = ""
    If MaxRecords > 0 Then
        If isSQL Then
            pSQL = pSQL & "{SET ROWCOUNT " & MaxRecords & " SELECT [@PARENTFIELDS]"
        Else
            pSQL = pSQL & "{SELECT TOP " & MaxRecords & " [@PARENTFIELDS]"
        End If
    Else
        pSQL = pSQL & "{SELECT " & "[@PARENTFIELDS]"
    End If
    pSQL = pSQL & " FROM [@PARENT]"
    pSQL = pSQL & " [@WHEREPARENT]"
    pSQL = pSQL & " [@ORDPARENT]} "
    'Substitute for actual values
    pSQL = Replace(pSQL, "[@PARENTFIELDS]", ParentFields)
    pSQL = Replace(pSQL, "[@PARENT]", tblParent)
    pSQL = Replace(pSQL, "[@WHEREPARENT]", pWhere)
    pSQL = Replace(pSQL, "[@ORDPARENT]", pOrder)
    pSQL = Trim(pSQL)
    'Define Child SQL Statement
    cSQL = ""
    cSQL = cSQL & "{SELECT " & "[@CHILDFIELDS]"
    cSQL = cSQL & " FROM [@CHILD]"
    cSQL = cSQL & " [@WHERECHILD]"
    cSQL = cSQL & " [@ORDCHILD]} "
    'Substitute for actual values
    cSQL = Replace(cSQL, "[@CHILDFIELDS]", ChildFields)
    cSQL = Replace(cSQL, "[@CHILD]", tblChild)
    cSQL = Replace(cSQL, "[@WHERECHILD]", cWhere)
    cSQL = Replace(cSQL, "[@ORDCHILD]", cOrder)
    cSQL = Trim(cSQL)

    'Define Parent Properties
    lSQL = "SHAPE " & pSQL & vbCrLf
    'Define Child Properties
    lSQL = lSQL & "APPEND (" & cSQL & " RELATE " & fldParent & " TO " & fldChild & ") AS ChildItems"

    'Get the data
    Set cn = New ADODB.Connection
    With cn
        .ConnectionString = ConnectionString
        .CursorLocation = adUseServer
        .Provider = "MSDataShape"
        .Open
    End With

    Set rS = New ADODB.Recordset
    With rS
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .source = lSQL
        .ActiveConnection = cn
        .Open
    End With
    Set rS.ActiveConnection = Nothing
    cn.Close
    Set cn = Nothing
    Set Datashape = rS
    Set rS = Nothing
Exit Function
ErrHandler:
    If Not cn Is Nothing Then
            
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    m_sLastError = Err.Number & " - " & Err.Description

End Function

Public Function EmptyRS(ByVal adoRS As ADODB.Recordset) As Boolean
    'Checks for an EMPTY RecordSet
    On Error GoTo ErrHandler
    EmptyRS = True
    If Not adoRS Is Nothing Then
        EmptyRS = ((adoRS.BOF = True) And (adoRS.EOF = True))
    End If
Exit Function
ErrHandler:
    m_sLastError = Err.Number & " - " & Err.Description
    EmptyRS = True
End Function

Public Function Execute(SQL As String) As Boolean
    'TO DIRECTLY EXECUTE AN INSERT, UPDATE, OR DELETE
    'SQL STATMENT. SET THE CONNECTION STRING PROPERTY
    'TO A VALID CONNECTION STRING FIRST
    
    On Error GoTo LocalError
    Dim cn As New ADODB.Connection
    With cn
        .ConnectionString = ConnectionString
        .CursorLocation = adUseServer
        .Open
        .BeginTrans
        .Execute SQL
        .CommitTrans
        .Close
    End With
    Set cn = Nothing
    Execute = True
Exit Function
LocalError:
    m_sLastError = Err.Number & " - " & Err.Description
    Debug.Print Err.Description
    If cn.State = adStateOpen Then
        cn.RollbackTrans
        cn.Close
    End If
    Set cn = Nothing
    Execute = False
End Function

Public Function GetRS(SQL As String) As ADODB.Recordset
    'SET THE CONNECTION STRING PROPERTY TO A VALID CONNECTION STRING
    'PASS AN SQL STATEMENT TO THIS FUNCTION
    'THE RETURN VALUE WILL BE AN ADODB RECORDSET
    
    Dim rS As New ADODB.Recordset
    On Error GoTo LocalError
    With rS
        .ActiveConnection = ConnectionString
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenKeyset
        .source = SQL
        .Open
        Set .ActiveConnection = Nothing
    End With
    Set GetRS = rS
    Set rS = Nothing
Exit Function
LocalError:
    m_sLastError = Err.Number & " - " & Err.Description
    Set rS = Nothing
End Function

Public Function GetCount(TableName As String, Optional WhereClause As String = "") As Long
    
    'RETURNS COUNT OF RECORDS WITHIN A TABLE, WITH OPTIONAL WHERE CLAUSE
    
    
    
    On Error GoTo LocalError
    Dim rS  As New ADODB.Recordset
    Dim SQL As String
    GetCount = 0
    If WhereClause <> "" Then
        SQL = "Select COUNT (*) FROM " & TableName & " WHERE " & WhereClause
    Else
        SQL = "Select COUNT (*) FROM " & TableName
    End If
    With rS
        .ActiveConnection = ConnectionString
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .CursorType = adOpenKeyset
        .source = SQL
        .Open
        Set .ActiveConnection = Nothing
    End With
    GetCount = rS.Fields(0).Value
    Set rS = Nothing
Exit Function
LocalError:
    m_sLastError = Err.Number & " - " & Err.Description
    If rS.State = adStateOpen Then
        rS.Close
    End If
    Set rS = Nothing
    GetCount = -1
End Function

Public Function PutRS(rS As ADODB.Recordset) As Boolean
'USE THIS TO UPDATE A RECORDSET IN BATCH (TRANSACTIONAL) MODE
'IF CHANGES TO THE RECORDSET'S WERE MADE PRIOR TO THIS CALL
'THIS FUNCTION WILL COMMIT THEM TO THE UNDERYLING DATABASE


    On Error GoTo LocalError
    PutRS = False
    If EmptyRS(rS) Then
        Exit Function
    ElseIf rS.LockType = adLockReadOnly Then
        Exit Function
    Else
        Dim cn As New ADODB.Connection
        With cn
            .ConnectionString = ConnectionString
            .CursorLocation = adUseServer
            .Open
            .BeginTrans
        End With
        With rS
            .ActiveConnection = cn
            .UpdateBatch
            cn.CommitTrans
            Set .ActiveConnection = Nothing
        End With
        cn.Close
        Set cn = Nothing
    End If
    PutRS = True
Exit Function
LocalError:
    m_sLastError = Err.Number & " - " & Err.Description
    If cn.State = adStateOpen Then
        cn.RollbackTrans
        cn.Close
    End If
    Set cn = Nothing
    PutRS = False
End Function

Public Function sqlBoolean(TrueFalse As Boolean) As Integer
    'CONVERTS BIT RETURN VALUE FROM SQL SERVER
    
    'This is because SQL True = 1
    'VB True = -1
    sqlBoolean = TrueFalse
    If isSQL Then
        If TrueFalse = True Then sqlBoolean = TrueFalse * TrueFalse
    End If
End Function

Public Function sqlDate(ByVal vDate As Variant) As String

'THIS FUNCTION TAKES VALUES THAT ARE POSSIBLE
'DATES AND FORMATS THEM PROPERFOR INSERTION INTO
'DATABASE COLUMNS DEFINED AS DATES

    On Error GoTo LocalError
    'Remove all invalid characters
    vDate = Trim(CStr(vDate))
    vDate = Replace(vDate, "#", "")
    vDate = Replace(vDate, "'", "")
    vDate = Replace(vDate, Chr(34), "")
    '--------------------------------------
    'Convert the Date to a Double Precision
    '   for international compatability
    '--------------------------------------
    sqlDate = ""
    'First see in what format the data came
    ' Validate parameter
    If Not IsDate(vDate) Or IsNull(vDate) Then
        'Maybe it is a number
        If IsNumeric(vDate) Then
            vDate = CDate(vDate)
        End If
        If Not IsDate(vDate) Then
            'Still not a date
            Exit Function
        End If
    End If
    If isSQL Then
        'Format is MM/DD/??YY
        sqlDate = Format(vDate, "mm\/dd\/yyyy")
        sqlDate = "'" & sqlDate & "'"
    Else
        'Format by Regional Settings
        sqlDate = FormatDateTime(vDate, vbShortDate)
        sqlDate = "#" & sqlDate & "#"
    End If
Exit Function
LocalError:
    m_sLastError = Err.Number & " - " & Err.Description
    sqlDate = ""
End Function

Public Function sqlDateTime(ByVal vDate As Variant) As String

'THIS FUNCTION TAKES VALUES THAT ARE POSSIBLE
'DATES AND FORMATS THEM PROPERFOR INSERTION INTO
'DATABASE COLUMNS DEFINED AS DATE/TIMES

    On Error GoTo LocalError
    'Remove all invalid characters
    vDate = Trim(CStr(vDate))
    vDate = Replace(vDate, "#", "")
    vDate = Replace(vDate, "'", "")
    vDate = Replace(vDate, Chr(34), "")
    '--------------------------------------
    'Convert the Date to a Double Precision
    '   for international compatability
    '--------------------------------------
    sqlDateTime = ""
    'First see in what format the data came
    ' Validate parameter
    If Not IsDate(vDate) Or IsNull(vDate) Then
        'Maybe it is a number
        If IsNumeric(vDate) Then
            vDate = CDate(vDate)
        End If
        If Not IsDate(vDate) Then
            'Still not a date
            Exit Function
        End If
    End If
    If isSQL Then
        'Format is MM/DD/??YY HH:MM:SS
        sqlDateTime = Format(vDate, "mm\/dd\/yyyy hh\:mm\:ss")
        sqlDateTime = "'" & sqlDateTime & "'"
    Else
        'Format by Regional Settings
        sqlDateTime = FormatDateTime(vDate, vbShortDate) & " " & Format(vDate, "hh\:mm\:ss")
        sqlDateTime = "#" & sqlDateTime & "#"
    End If
Exit Function
LocalError:
    m_sLastError = Err.Number & " - " & Err.Description
    sqlDateTime = ""
End Function

Public Function sqlTime(ByVal vDate As Variant) As String
'THIS FUNCTION TAKES VALUES THAT ARE POSSIBLE
'DATES AND FORMATS THEM PROPERFOR INSERTION INTO
'DATABASE COLUMNS DEFINED AS TIME ONLY
    
    On Error GoTo LocalError
    'Remove all invalid characters
    vDate = Trim(CStr(vDate))
    vDate = Replace(vDate, "#", "")
    vDate = Replace(vDate, "'", "")
    vDate = Replace(vDate, Chr(34), "")
    '--------------------------------------
    'Convert the Date to a Double Precision
    '   for international compatability
    '--------------------------------------
    sqlTime = ""
    'First see in what format the data came
    ' Validate parameter
    If Not IsDate(vDate) Or IsNull(vDate) Then
        'Maybe it is a number
        If IsNumeric(vDate) Then
            vDate = CDate(vDate)
        End If
        If Not IsDate(vDate) Then
            'Still not a date
            Exit Function
        End If
    End If
    If isSQL Then
        'Format is MM/DD/??YY HH:MM:SS
        sqlTime = FormatDateTime(vDate, vbLongTime)
        sqlTime = "'" & sqlTime & "'"
    Else
        'Format by Regional Settings
        sqlTime = FormatDateTime(vDate, vbLongTime)
        sqlTime = "#" & sqlTime & "#"
    End If
Exit Function
LocalError:
    m_sLastError = Err.Number & " - " & Err.Description
    sqlTime = ""
End Function

Public Function sqlEncode(sqlValue) As String

    'IF A STRING VALUE IN AN SQL STATMENT HAS A ' CHARACTER,
    'USE THIS FUNCTION SO THE STRING CAN BE USED IN THE STATEMENT
     sqlEncode = Replace(sqlValue, "'", "''")
End Function


Public Property Get LastError() As String
    'IF AN ERROR OCCURS IN CALLING ONE OF THE FUNCTIONS IN THIS CLASS
    'READ THIS PROPERTY TO SEE WHAT THE ERROR WAS

    LastError = m_sLastError
    m_sLastError = ""
End Property

Public Function ExecuteID(SQL As String) As Long
'PURPOSE: RETURN VALUE OF IDENTITY COLUMN
'OF A NEWLY INSERTED RECORD

'SQL is a valid Insert statement.
'ConnetionString properyt has been set to a valid Connection String
'Tested on SQL7 as well as ACCESS 2000 using Jet4

    On Error GoTo LocalError
    Dim cn As New ADODB.Connection
    Dim rS As New ADODB.Recordset
    Dim AutoID As Long
    
    With rS
        'Prepare the RecordSet
        .CursorLocation = adUseServer
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .source = "SELECT @@IDENTITY"
    End With
    
    With cn
        .ConnectionString = ConnectionString
        .CursorLocation = adUseServer
        .Open
        .BeginTrans
        .Execute SQL, , adCmdText + adExecuteNoRecords
        With rS
            .ActiveConnection = cn
            .Open , , , , adCmdText
            AutoID = rS(0).Value
            .Close
        End With
        .CommitTrans
        .Close
    End With

    Set rS = Nothing
    Set cn = Nothing

    'If we get here ALL was Okay
    ExecuteID = AutoID

    Exit Function
    
LocalError:
    m_sLastError = Err.Number & " - " & Err.Description
    ExecuteID = 0
End Function

Private Sub Class_Initialize()
    m_DBType = Other
End Sub

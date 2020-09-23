VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal RESERVED As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal RESERVED As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal numBytes As Long)

Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_MULTI_SZ = 7
Private Const ERROR_MORE_DATA = 234
Private Const REG_OPENED_EXISTING_KEY = &H2
Private Const KEY_READ = &H20019  ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
                          ' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
                          ' SYNCHRONIZE))

Private Const KEY_WRITE = &H20006  '((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or
                           ' KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Public Enum HKEYS
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
  HKEY_PERFORMANCE_DATA = &H80000004
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_DYN_DATA = &H80000006
End Enum

Public Function IsWinNT() As Boolean
    ' Returns True is executed under Windows NT,
    ' False if executed under Windows 95/98
    
    IsWinNT = (GetVersion() And &H80000000) = 0
End Function

' Read a Registry value
'
' Use KeyName = "" for the default value
' If the value isn't there, it returns the DefaultValue
' argument, or Empty if the argument has been omitted
'
' Supports DWORD, REG_SZ, REG_EXPAND_SZ, REG_BINARY and REG_MULTI_SZ
' REG_MULTI_SZ values are returned as a null-delimited stream of strings
' (VB6 users can use Split to convert to an array of string)
Public Function GetRegistryValue(ByVal hKey As HKEYS, ByVal KeyName As String, ByVal ValueName As String, Optional DefaultValue As Variant) As Variant
    Dim handle As Long
    Dim resLong As Long
    Dim resString As String
    Dim resBinary() As Byte
    Dim length As Long
    Dim retval As Long
    Dim valueType As Long
    
    ' Prepare the default result
    GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)
    
    ' Open the key, exit if not found.
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
        
    ' prepare a 1K receiving resBinary
    length = 1024
    ReDim resBinary(0 To length - 1) As Byte
    
    ' read the registry key
    retval = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), length)
    
    ' if resBinary was too small, try again
    If retval = ERROR_MORE_DATA Then
        ' enlarge the resBinary, and read the value again
        ReDim resBinary(0 To length - 1) As Byte
        retval = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), length)
    End If
    
    ' return a value corresponding to the value type
    Select Case valueType
        Case REG_DWORD
            CopyMemory resLong, resBinary(0), 4
            GetRegistryValue = resLong
        Case REG_SZ, REG_EXPAND_SZ
            ' copy everything but the trailing null char
            resString = Space$(length - 1)
            CopyMemory ByVal resString, resBinary(0), length - 1
            GetRegistryValue = resString
        Case REG_BINARY
            ' resize the result resBinary
            If length <> UBound(resBinary) + 1 Then
                ReDim Preserve resBinary(0 To length - 1) As Byte
            End If
            GetRegistryValue = resBinary()
        Case REG_MULTI_SZ
            ' copy everything but the 2 trailing null chars
            resString = Space$(length - 2)
            CopyMemory ByVal resString, resBinary(0), length - 2
            GetRegistryValue = resString
        Case Else
            RegCloseKey handle
            '  Err.Raise 1001, , "Unsupported value type"
    End Select
    
    ' close the registry key
    RegCloseKey handle
End Function

' Write or Create a Registry value
' returns True if successful
'
' Use KeyName = "" for the default value
'
' Value can be an integer value (REG_DWORD), a string (REG_SZ)
' or an array of binary (REG_BINARY). Raises an error otherwise.
Public Function SetRegistryValue(ByVal hKey As HKEYS, ByVal KeyName As String, ByVal ValueName As String, Value As Variant) As Boolean
    Dim handle As Long
    Dim lngValue As Long
    Dim strValue As String
    Dim binValue() As Byte
    Dim length As Long
    Dim retval As Long
    
    ' Open the key, exit if not found
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, handle) Then Exit Function
        
    ' three cases, according to the data type in Value
    Select Case VarType(Value)
        Case vbInteger, vbLong
            lngValue = Value
            retval = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
        Case vbString
            strValue = Value
            retval = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, _
            Len(strValue))
        Case vbBoolean
            lngValue = Abs(Value)
            retval = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
        Case vbArray + vbByte
            binValue = Value
            length = UBound(binValue) - LBound(binValue) + 1
            retval = RegSetValueEx(handle, ValueName, 0, REG_BINARY, binValue(LBound(binValue)), length)
        Case Else
        RegCloseKey handle
        '  Err.Raise 1001, , "Unsupported value type"
    End Select
    
    ' Close the key and signal success
    RegCloseKey handle
    
    ' signal success if the value was written correctly
    SetRegistryValue = (retval = 0)
End Function

' Enumerate values under a given registry key returns a collection,
' where each element of the collection is a 2-element array of Variants:
' element(0) is the value name, element(1) is the value's value
Public Function EnumRegistryValues(ByVal hKey As HKEYS, ByVal KeyName As String) As Collection
    Dim handle As Long
    Dim index As Long
    Dim valueType As Long
    Dim name As String
    Dim nameLen As Long
    Dim resLong As Long
    Dim resString As String
    Dim dataLen As Long
    Dim valueInfo(0 To 1) As Variant
    Dim retval As Long
    
    ' initialize the result
    Set EnumRegistryValues = New Collection
    
    ' Open the key, exit if not found.
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all cases, subsequent functions use hKey
        hKey = handle
    End If
    
    Do
        ' this is the max length for a key name
        nameLen = 260
        name = Space$(nameLen)
        ' prepare the receiving buffer for the value
        dataLen = 4096
        ReDim resBinary(0 To dataLen - 1) As Byte
        
        ' read the value's name and data exit the loop if not found
        retval = RegEnumValue(hKey, index, name, nameLen, ByVal 0&, valueType, resBinary(0), dataLen)
        
        ' enlarge the buffer if you need more space
        If retval = ERROR_MORE_DATA Then
            ReDim resBinary(0 To dataLen - 1) As Byte
            retval = RegEnumValue(hKey, index, name, nameLen, ByVal 0&, _
            valueType, resBinary(0), dataLen)
        End If
        
        ' exit the loop if any other error (typically, no more values)
        If retval Then Exit Do
        
        ' retrieve the value's name
        valueInfo(0) = Left$(name, nameLen)
        
        ' return a value corresponding to the value type
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong
            Case REG_SZ, REG_EXPAND_SZ
                ' copy everything but the trailing null char
                resString = Space$(dataLen - 1)
                CopyMemory ByVal resString, resBinary(0), dataLen - 1
                valueInfo(1) = resString
            Case REG_BINARY
                ' shrink the buffer if necessary
                If dataLen < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To dataLen - 1) As Byte
                End If
                valueInfo(1) = resBinary()
            Case REG_MULTI_SZ
                ' copy everything but the 2 trailing null chars
                resString = Space$(dataLen - 2)
                CopyMemory ByVal resString, resBinary(0), dataLen - 2
                valueInfo(1) = resString
            Case Else
            ' Unsupported value type - do nothing
        End Select
        
        ' add the array to the result collection the element's key is the value's name
        ' EnumRegistryValues.Add valueInfo(1), valueInfo(0)
        EnumRegistryValues.Add valueInfo, valueInfo(0)
        index = index + 1
    Loop
    
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
End Function

' Enumerate registry keys under a given key returns a collection of strings
Public Function EnumRegistryKeys(ByVal hKey As Long, ByVal KeyName As String) As Collection
    Dim handle As Long
    Dim length As Long
    Dim index As Long
    Dim subkeyName As String
    
    ' initialize the result collection
    Set EnumRegistryKeys = New Collection
    
    ' Open the key, exit if not found
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all case the subsequent functions use hKey
        hKey = handle
    End If
    
    Do
        ' this is the max length for a key name
        length = 260
        subkeyName = Space$(length)
        ' get the N-th key, exit the loop if not found
        If RegEnumKey(hKey, index, subkeyName, length) Then Exit Do
        
        ' add to the result collection
        subkeyName = Left$(subkeyName, InStr(subkeyName, vbNullChar) - 1)
        EnumRegistryKeys.Add subkeyName, subkeyName
        
        ' prepare to query for next key
        index = index + 1
    Loop
    
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
End Function

' ===[ DELETE ] ===

' Delete a registry key
' Under Windows NT it doesn't work if the key contains subkeys
Public Sub DeleteRegistryKey(ByVal hKey As HKEYS, ByVal KeyName As String)
    RegDeleteKey hKey, KeyName
End Sub

' Delete a registry value
' Return True if successful, False if the value hasn't been found
Public Function DeleteRegistryValue(ByVal hKey As HKEYS, ByVal KeyName As String, ByVal ValueName As String) As Boolean
    Dim handle As Long
    
    ' Open the key, exit if not found
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, handle) Then Exit Function
    
    ' Delete the value (returns 0 if success)
    DeleteRegistryValue = (RegDeleteValue(handle, ValueName) = 0)
    
    ' Close the handle
    RegCloseKey handle
End Function

' Return True if a Registry key exists
Public Function CheckRegistryKey(ByVal hKey As HKEYS, ByVal KeyName As String) As Boolean
    Dim handle As Long
    ' Try to open the key
    CheckRegistryKey = (RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) = 0)
    
    ' Close it before exiting
    RegCloseKey handle
End Function

' Create a registry key, then close it
Public Function CreateRegistryKey(ByVal hKey As HKEYS, ByVal KeyName As String) As Boolean
    ' Pre  :
    ' Post : Returns True if the key already existed, False if it was created.
    Dim handle As Long, disposition As Long
    
    If RegCreateKeyEx(hKey, KeyName, 0, 0, 0, 0, 0, handle, disposition) Then
        Err.Raise 1001, , "Unable to create the registry key"
    Else
        ' Return True if the key already existed.
        CreateRegistryKey = (disposition = REG_OPENED_EXISTING_KEY)
        ' Close the key.
        RegCloseKey handle
    End If
End Function


' JMB ===================================================================================
' JMB
' JMB   Examine below and use parts for revision
' JMB
' JMB   Immediate notices:
' JMB       Error handling

'    '***********************************************************************************'
'    ' Programmer : Patrick S. Seymour '
'    ' '
'    ' Created : January 2, 1999 '
'    ' Last Modified : June 5, 1999 (to make it NT-compliant) '
'    ' '
'    ' Purpose : Provides functions that interact with the Windows registry. '
'    '***********************************************************************************'
'    '
'    ' require that all variables be declared
'    Option Explicit
'
'    ' Type to store information about a key (see the GetKeyInfo function)
'    Public Type KeyInfo
'        Subkeys As Long ' number of subkeys
'        LenSubkeys As Long ' length of the longest subkey names
'        Values As Long ' number of value entries in the key
'        LenValueNames As Long ' length of the longest value names
'        LenValues As Long ' length (bytes) of the longest value
'    End Type ' end KeyInfo type declaration
'
'    ' Type to store access time information for a file (used by API functions)
'    Public Type FILETIME
'        lLowDateTime As Long
'        lHighDateTime As Long
'    End Type ' end FILETIME Type declaration
'
'    ' Type to store security information for a registry key
'    Public Type SECURITY_ATTRIBUTES
'        lLength As Long
'        lSecurityDescriptor As Long
'        lInheritHandle As Long
'    End Type ' end SECURITY_ATTRIBUTES Type declaration
'
'    Public Type ACL
'        AclRevision As Byte
'        Sbz1 As Byte
'        AclSize As Integer
'        AceCount As Integer
'        Sbz2 As Integer
'    End Type
'
'    ' Type to handle the security descriptor for a key
'    Public Type SECURITY_DESCRIPTOR
'        Revision As Byte
'        Sbz1 As Byte
'        Control As Long
'        Owner As Long
'        Group As Long
'        Sacl As ACL
'        Dacl As ACL
'    End Type ' end SECURITY_DESCRIPTOR Type declaration
'
'    ' RegQueryMultipleValues is not included in this module because it uses the VALENT
'    ' data structure, which contains pointers. Since Visual Basic does not support
'    ' pointers, the function is incompatible with Visual Basic.
'
'    ' RegRemapPreDefKey is not included either, mainly because its usefulness is limited
'    ' and its use dangerous. This function can change which registry key is pointed to
'    ' by HKEY_CURRENT_USER and/or HKEY_CURRENT_CONFIG. If you feel a need to use this
'    ' function, feel free to Declare it and use it. However, we will all know that you
'    ' are insane.
'
'    ' RegSaveKey, RegLoadKey, RegUnloadKey, and RegReplaceKey are not included here
'    ' because they deal with importing and exporting registry sections in their binary
'    ' format. These functions do have their uses, but I had no use for them when
'    ' creating this module for its target audience.
'
'    ' API registry functions... declared non-Private so programmers can call the functions
'    ' by their Declare names if desired
'    Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'    Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal sMachineName As String, ByVal hKey As Long, hResult As Long) As Long
'    Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal sSubkey As String, ByVal lReserved As Long, ByVal sClass As String, ByVal lOptions As Long, ByVal lSAMDesired As Long, lSecurityAttributes As Long, lResult As Long, lDisposition As Long) As Long
'    Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal sSubkey As String) As Long
'    Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal sValueName As String) As Long
'    Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal lIndex As Long, ByVal sName As String, lName As Long, ByVal lReserved As Long, ByVal sClass As String, lClass As Long, ftLastWriteTime As FILETIME) As Long
'    Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal lIndex As Long, ByVal sValueName As String, lValueName As Long, ByVal lReserved As Long, lType As Long, ByVal sData As String, lData As Long) As Long
'    Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'    Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal sSubkey As String, ByVal lOptions As Long, ByVal lSAMDesired As Long, hResult As Long) As Long
'    Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal sClass As String, lClass As Long, ByVal lReserved As Long, lSubKeys As Long, lMaxSubKeyLen As Long, lMaxClassLen As Long, lValues As Long, lMaxValueNameLen As Long, lMaxValueLen As Long, lSecurityDescriptor As Long, ftLastWriteTime As FILETIME) As Long
'    Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal sValueName As String, ByVal lReserved As Long, lType As Long, aData As Any, lData As Long) As Long
'    Declare Function RegQueryStringEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
'    Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal sValueName As String, ByVal lReserved As Long, ByVal lType As Long, aData As Any, ByVal lData As Long) As Long
'    Declare Function RegSetStringEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal RESERVED As Long, ByVal dwType As Long, lpData As String, ByVal cbData As Long) As Long
'
'    ' API registry functions supported by Windows NT and not by Windows 9x
'    Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal sFile As String, ByVal lFlags As Long) As Long
'    Declare Function RegSetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal lSecurityInformation As Long, sdDescriptor As SECURITY_DESCRIPTOR) As Long
'    Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hKey As Long, ByVal lWatchSubtree As Long, ByVal lNotifyFilter As Long, ByVal hEvent As Long, ByVal lAsynchronus As Long) As Long
'    Declare Function RegGetKeySecurity Lib "advapi32.dll" (ByVal hKey As Long, ByVal lSecurityInformation As Long, sdDescriptor As SECURITY_DESCRIPTOR, lSecurityDescriptor As Long) As Long
'
'    ' values for the top-level keys
'    Public Const HKEY_CLASSES_ROOT = &H80000000
'    Public Const HKEY_CURRENT_USER = &H80000001
'    Public Const HKEY_LOCAL_MACHINE = &H80000002
'    Public Const HKEY_USERS = &H80000003
'    Public Const HKEY_CURRENT_CONFIG = &H80000005
'    Public Const HKEY_DYN_DATA = &H80000006
'
'    ' Abbreviations for top-level keys. These abbreviations are not official, but many
'    ' developers utilizing the registry have come to use this same set of abbreviations.
'    Public Const HKCR = HKEY_CLASSES_ROOT
'    Public Const HKCU = HKEY_CURRENT_USER
'    Public Const HKLM = HKEY_LOCAL_MACHINE
'    Public Const HKU = HKEY_USERS
'    Public Const HKCC = HKEY_CURRENT_CONFIG
'    Public Const HKDD = HKEY_DYN_DATA
'
'    ' Error codes returned by registry API functions. Note that, in the API, these
'    ' constant names are prefaced by ERROR_. For example, SUCCESS is actually
'    ' ERROR_SUCCESS in the API. I've shortened the names for easier typing. See the
'    ' ShowErr Sub (below) for an explanation of each of these constants.
'    Public Const ACCESS_DENIED = 5
'    Public Const BAD_NETPATH = 53
'    Public Const BAD_PATHNAME = 161
'    Public Const BADDB = 1009
'    Public Const BADKEY = 1010
'    Public Const CALL_NOT_IMPLEMENTED = 120
'    Public Const CANTOPEN = 1011
'    Public Const CANTREAD = 1012
'    Public Const CANTWRITE = 1013
'    Public Const DLL_INIT_FAILED = 1114
'    Public Const FILE_NOT_FOUND = 2
'    Public Const INSUFFICIENT_BUFFER = 122
'    Public Const INVALID_HANDLE = 6
'    Public Const INVALID_PARAMETER = 87
'    Public Const KEY_DELETED = 1018
'    Public Const KEY_NOT_FOUND = FILE_NOT_FOUND ' not a standard error code constant
'    Public Const LOCK_FAILED = 167
'    Public Const MORE_DATA = 234
'    Public Const NO_MORE_ITEMS = 259
'    Public Const NOT_REGISTRY_FILE = 1017
'    Public Const REGISTRY_CORRUPT = 1015
'    Public Const REGISTRY_IO_FAILED = 1016
'    Public Const REGISTRY_RECOVERED = 1014
'    Public Const SUCCESS = 0
'    Public Const TRANSFER_TOO_LONG = 222
'
'    ' Options (Note: In the API, these constant names start with REG_OPTION_. For example,
'    ' VOLATILE below is REG_OPTION_VOLATILE in the API. The names have been truncated to
'    ' ease typing and protect the innocent.
'    Public Const BACKUP_RESTORE = 4
'    Public Const CREATE_LINK = 2
'    Public Const NON_VOLATILE = 0
'    Public Const RESERVED = 0
'    Public Const VOLATILE = 1
'
'    ' Data type constants for registry data. Note that in the API declarations for
'    ' these constants, the constant names are prefaced with REG_. For example,
'    ' the BINARY constant listed below is REG_BINARY in the API. As with the error codes,
'    ' I have shortened the constant names for easier typing.
'    Public Const BINARY = 3 ' binary data in any form
'    Public Const DWORD = 4 ' 4-byte number stored in
'    ' DWORD_LITTLE_ENDIAN format
'    Public Const DWORD_LITTLE_ENDIAN = DWORD ' see DWORD
'    Public Const DWORD_BIG_ENDIAN = 5 ' 4-byte number in big-endian format
'    Public Const EXPAND_SZ = 2 ' null-terminated string that
'    ' contains unexpanded references to
'    ' environment variables that should
'    ' be replaced with their values
'    Public Const MULTI_SZ = 7 ' an array of null-terminated
'    ' strings, terminated by an extra
'    ' null to mark the end of the array
'    Public Const NONE = 0 ' no defined value type
'    Public Const SZ = 1 ' null-terminated string
'
'    Dim lErrorCode As Long ' error code returned by a registry function; this variable
'    ' is used in the majority of the functions below, so I just
'    ' made it global
'
'    ' security access mask constants
'    ' Windows 95 and 98 do not store security information in the registry, so these
'    ' constants are ignored when calling API registry functions under those operating
'    ' systems. Windows NT, however, does store security information with its keys, so
'    ' these constants are necessary when accessing the registry under Windows NT.
'    Public Const KEY_ALL_ACCESS = &HF003F
'    Public Const KEY_CREATE_LINK = &H20
'    Public Const KEY_CREATE_SUB_KEY = &H4
'    Public Const KEY_ENUMERATE_SUB_KEYS = &H8
'    Public Const KEY_EXECUTE = &H20019
'    Public Const KEY_NOTIFY = &H10
'    Public Const KEY_QUERY_VALUE = &H1
'    Public Const KEY_READ = &H20019
'    Public Const KEY_SET_VALUE = &H2
'    Public Const KEY_WRITE = &H20006
'
'    ' Some registry functions have parameters that are not currently in use. For these
'    ' parameters, the value passed must be zero (0). To make the code below more readable,
'    ' the following constant will be used in place of zero (0).
'    Public Const NOT_USED = 0
'
'
'    Module
'
'     '*** OpenKey *************************************************************************
'    ' Purpose
'    ' opens a registry key
'    ' Required Arguments
'    ' hKey - handle of the key whose subkey will be opened
'    ' sSubkey - path of the subkey to open
'    ' hNewHandle - handle of the newly opened key (if successful)
'    ' Optional Arguments
'    ' bCreate - whether to create a new key if sSubkey doesn't exist (default is
'    ' False: no key will be created)
'    ' Return Value
'    ' a long containing the error code returned by the registry API function
'    '*************************************************************************************
'    Public Function OpenKey(hKey As Long, sSubkey As String, ByRef hNewHandle As Long, sClassName As String, lOptions, lSAMDesired, lSecurityAttributes, Optional bCreate = False) As Long
'
'    Dim lDisposition As Long ' whether RegCreateKeyEx opened an existing key or
'    ' created a new one
'
'        Do While Left(sSubkey, 1) = "\"
'            sSubkey = Right(sSubkey, Len(sSubkey) - 1)
'        Loop
'        If bCreate Then
'            lErrorCode = RegCreateKeyEx(hKey, sSubkey, NOT_USED, sClassName, lOptions, lSAMDesired, lSecurityAttributes, hNewHandle, lDisposition)
'        Else
'            lErrorCode = RegOpenKeyEx(hKey, sSubkey, NOT_USED, lSAMDesired, hNewHandle)
'        End If
'        OpenKey = lErrorCode
'    End Function
'
'    '*** OpenRemoteKey *******************************************************************
'    ' Purpose
'    ' opens a top-level key on a remote computer
'    ' Required Arguments
'    ' sMachineName - UNC name of the machine to be accessed
'    ' hTopLevelKey - top-level key to access on the machine (see top-level constants)
'    ' hRemoteKey - handle of the top-level key on the remote machine (if successful)
'    ' Optional Arguments
'    ' none
'    ' Return Value
'    ' a long containing the error code returned by the registry API function
'    '*************************************************************************************
'    Public Function OpenRemoteKey(sMachineName As String, hTopLevelKey As Long, ByRef hRemoteKey As Long) As Long
'        OpenRemoteKey = RegConnectRegistry(sMachineName, hTopLevelKey, hRemoteKey)
'    End Function
'
'    '*** GetKeyInfo **********************************************************************
'    ' Purpose
'    ' gathers information about a key
'    ' Required Arguments
'    ' hKey - handle of the key about which information is to be gathered
'    ' kiKeyInfo - a KeyInfo variable containing the information about the key (see the
'    ' Type statement above)
'    ' Optional Arguments
'    ' none
'    ' Return Value
'    ' a long containing the error code returned by the registry API function
'    '*************************************************************************************
'    Public Function GetKeyInfo(hKey As Long, ByRef kiKeyInfo As KeyInfo) As Long
'
'    ' These variables are declared merely to make the function work. Windows 9x
'    ' does not use the values that are passed in these variables in calls to
'    ' RegQueryInfoKey.
'        Dim sClassName As String ' class name of the key
'        Dim lLenClassName As Long ' length of the key's class name
'        Dim lMaxLenClass As Long ' length of the longest class name of the key's
'    ' subkeys
'        Dim lDescriptor As Long ' security descriptor
'        Dim ftWriteTime As FILETIME ' last time this key was written
'
'        lErrorCode = RegQueryInfoKey(hKey, sClassName, lLenClassName, NOT_USED, kiKeyInfo.Subkeys, kiKeyInfo.LenSubkeys, lMaxLenClass, kiKeyInfo.Values, kiKeyInfo.LenValueNames, kiKeyInfo.LenValues, lDescriptor, ftWriteTime)
'    GetKeyInfo = lErrorCode
'    End Function
'
'    '*** GetKeyName **********************************************************************
'    ' Purpose
'    ' gets the name of a key
'    ' Required Arguments
'    ' hKey - handle of the key containing the key whose name is sought
'    ' lKeyNum - number of the key to be enumerated
'    ' sKeyName - name of the key
'    ' Optional Arguments
'    ' none
'    ' Return Value
'    ' a long containing the error code returned by the registry API function
'    '*************************************************************************************
'    Public Function GetKeyName(hKey As Long, ByVal lKeyNum As Long, sKeyName As String) As Long
'        Dim sClassName As String
'        Dim lLenClassName As Long
'        Dim ftWriteTime As FILETIME
'        Dim lNameLen As Long
'        lNameLen = Len(sKeyName)
'        lErrorCode = RegEnumKeyEx(hKey, lKeyNum, sKeyName, lNameLen, NOT_USED, sClassName, lLenClassName, ftWriteTime)
'    sKeyName = Left(sKeyName, lNameLen)
'        GetKeyName = lErrorCode
'    End Function
'
'    '*** GetValueName ********************************************************************
'    ' Purpose
'    ' gets the name of a value
'    ' Required Arguments
'    ' hKey - handle of the key containing the value to be enumerated
'    ' lValueNum - number of the value to be enumerated
'    ' sValueName - name of the value
'    ' Optional Arguments
'    ' none
'    ' Return Value
'    ' a long containing the error code returned by the registry API function
'    '*************************************************************************************
'    Public Function GetValueName(hKey As Long, ByVal lValueNum As Long, ByRef sValueName As String) As Long
'        Dim LenValueName As Long
'        LenValueName = Len(sValueName)
'        lErrorCode = RegEnumValue(hKey, lValueNum, sValueName, LenValueName, ByVal NOT_USED, NOT_USED, vbNullString, NOT_USED)
'    ' The last three parameters in the call to RegEnumValue are not reserved. They
'    ' do actually serve a purpose, but not for this function. See the documentation
'    ' for RegEnumValue for more information.
'        sValueName = Left(sValueName, LenValueName)
'        GetValueName = lErrorCode
'    End Function
'
'    '*** GetStringValue ******************************************************************
'    ' Purpose
'    ' retrieves a String variable from the registry
'    ' Required Arguments
'    ' hKey - handle of the key containing the String data
'    ' sValueName - name of the String data in the registry
'    ' sData - the String retrieved from the registry
'    ' Optional Arguments
'    ' none
'    ' Return Value
'    ' a long containing the error code returned by the registry API function
'    '*************************************************************************************
'    Public Function GetStringValue(hKey As Long, sValueName As String, ByRef sData As String) As Long
'        Dim lStringLen As Long
'        lStringLen = Len(sData)
'        lErrorCode = RegQueryStringEx(hKey, sValueName, NOT_USED, SZ, sData, lStringLen)
'        If lStringLen > 0 Then sData = Left(sData, lStringLen - 1)
'        GetStringValue = lErrorCode
'    End Function
'
'    '*** SetStringValue ******************************************************************
'    ' Purpose
'    ' saves a String variable to the registry
'    ' Required Arguments
'    ' hKey - handle of the key to which the String will be saved
'    ' sValueName - name given to the String data in the registry
'    ' sData - the String to be saved in the registry
'    ' Optional Arguments
'    ' none
'    ' Return Value
'    ' a long containing the error code returned by the registry API function
'    '*************************************************************************************
'    Public Function SetStringValue(hKey As Long, sValueName As String, ByRef sData As String) As Long
'        lErrorCode = RegSetStringEx(hKey, sValueName, NOT_USED, SZ, ByVal sData, Len(sData))
'        SetStringValue = lErrorCode
'    End Function
'
'    '*** GetLongValue ********************************************************************
'    ' Purpose
'    ' retrieves a Long variable from the registry
'    ' Required Arguments
'    ' hKey - handle of the key containing the Long data
'    ' sValueName - name of the Long data in the registry
'    ' lData - the Long retrieved from the registry
'    ' Optional Arguments
'    ' none
'    ' Return Value
'    ' a long containing the error code returned by the registry API function
'    '*************************************************************************************
'    Public Function GetLongValue(hKey As Long, sValueName As String, lData As Long) As Long
'        lErrorCode = RegQueryValueEx(hKey, sValueName, NOT_USED, Binary, lData, Len(lData))
'        GetLongValue = lErrorCode
'    End Function
'
'    Public Function GetDWordValue(hKey As Long, sValueName As String, lData As Long) As Long
'        lErrorCode = RegQueryValueEx(hKey, sValueName, NOT_USED, DWORD, lData, Len(lData))
'        GetDWordValue = lErrorCode
'    End Function
'
'    '*** SetLongValue ********************************************************************
'    ' Purpose
'    ' saves a Long variable to the registry
'    ' Required Arguments
'    ' hKey - handle of the key to which the Long will be saved
'    ' sValueName - name given to the Long data in the registry
'    ' lData - the Long to be saved in the registry
'    ' Optional Arguments
'    ' none
'    ' Return Value
'    ' a long containing the error code returned by the registry API function
'    '*************************************************************************************
'    Public Function SetLongValue(hKey As Long, sValueName As String, lData As Long) As Long
'        lErrorCode = RegSetValueEx(hKey, sValueName, NOT_USED, Binary, lData, Len(lData))
'        SetLongValue = lErrorCode
'    End Function
'
'    Public Function SetDWordValue(hKey As Long, sValueName As String, lData As Long) As Long
'        lErrorCode = RegSetValueEx(hKey, sValueName, NOT_USED, DWORD, lData, Len(lData))
'        SetDWordValue = lErrorCode
'    End Function
'
'    Public Function GetIntegerValue(hKey As Long, sValueName As String, nData As Integer) As Long
'        lErrorCode = RegQueryValueEx(hKey, sValueName, NOT_USED, Binary, nData, Len(nData))
'        GetIntegerValue = lErrorCode
'    End Function
'
'    Public Function SetIntegerValue(hKey As Long, sValueName As String, nData As Integer) As Long
'        lErrorCode = RegSetValueEx(hKey, sValueName, NOT_USED, Binary, nData, Len(nData))
'        SetIntegerValue = lErrorCode
'    End Function
'    '*** SaveToDisk **********************************************************************
'    ' Purpose
'    ' forces saving of the key whose handle is hKey to be saved to disk
'    ' Required Arguments
'    ' hKey - handle of the key to be saved to disk
'    ' Optional Arguments
'    ' none
'    ' Return Value
'    ' a long containing the error code returned by the registry API function
'    '*************************************************************************************
'    Public Function SaveToDisk(hKey As Long) As Long
'        lErrorCode = RegFlushKey(hKey)
'        SaveToDisk = lErrorCode
'    End Function
'
'    '*** DeleteKey ***********************************************************************
'    ' Purpose
'    ' deletes a registry key
'    ' Required Arguments
'    ' hKey - handle of the key containing the key to be deleted
'    ' Optional Arguments
'    ' sSubkey - path of the subkey to be deleted (default is an empty string, which will
'    ' cause the key whose handle is hKey to be deleted)
'    ' Return Value
'    ' a long containing the error code returned by the registry API function
'    ' Notes
'    ' Windows 95 allows a key to be deleted even if it has subkeys. Windows NT, on the
'    ' other hand, does not. If you attempt to delete a key with subkeys in NT,
'    ' RegDeleteKey returns ACCESS_DENIED. In cases like this, this function will
'    ' perform a recursive delete transparent to the application. So whether coding for
'    ' 9x or NT, no special programming tricks need to be performed.
'    '*************************************************************************************
'    Public Function DeleteKey(hKey As Long, Optional sSubkey = "") As Long
'        Dim kiInfo As KeyInfo
'        Dim hSubkey As Long
'        Dim sPath As String
'        Dim sKeyName As String
'        Dim X As Integer
'
'        sPath = sSubkey
'        lErrorCode = RegDeleteKey(hKey, sSubkey)
'        If lErrorCode = ACCESS_DENIED Then
'    ' The key probably has subkeys. Windows NT will not allow a key to be deleted if
'    ' that key has subkeys. Another possibility is that the SAM is incorrect for the
'    ' key. Attempting to open the key will KEY_ALL_ACCESS will indicate which is the
'    ' case.
'            lErrorCode = OpenKey(hKey, sPath, hSubkey, "", 0, KEY_ALL_ACCESS, 0)
'            If lErrorCode = ACCESS_DENIED Then
'                MsgBox "Access denied to delete registry key.", vbOKOnly + vbCritical, "Access Denied"
'            Else
'                GetKeyInfo hSubkey, kiInfo
'                For X = kiInfo.Subkeys To 1 Step -1
'                    sKeyName = Space(40)
'                    GetKeyName hSubkey, X - 1, sKeyName
'                    DeleteKey hSubkey, sKeyName
'                Next X
'                DeleteKey hSubkey
'            End If
'            CloseKey hSubkey
'        End If
'        DeleteKey = lErrorCode
'    End Function
'
'    '*** DeleteValue *********************************************************************
'    ' Purpose
'    ' deletes a value
'    ' Required Arguments
'    ' hKey - handle of the key containing the value to be deleted
'    ' Optional Arguments
'    ' sValueName - name of the value to be deleted (default is an empty string, which
'    ' will cause the key's (Default) value to be deleted)
'    ' Return Value
'    ' a long containing the error code returned by the registry API function
'    '*************************************************************************************
'    Public Function DeleteValue(hKey As Long, Optional sValueName = "") As Long
'        lErrorCode = RegDeleteValue(hKey, sValueName)
'        DeleteValue = lErrorCode
'    End Function
'
'    '*** CloseKey ************************************************************************
'    ' Purpose
'    ' closes a registry key
'    ' Required Arguments
'    ' hKey - handle of the key to be closed
'    ' Optional Arguments
'    ' none
'    ' Return Value
'    ' a long containing the error code returned by the registry API function
'    '*************************************************************************************
'    Public Function CloseKey(hKey As Long) As Long
'        lErrorCode = RegCloseKey(hKey)
'        CloseKey = lErrorCode
'    End Function
'
'    '*** ShowErr *************************************************************************
'    ' Purpose
'    ' displays a message box describing the error code passed in lCode; does not display
'    ' a message box if lCode is 0 (a success error code)
'    ' Required Arguments
'    ' lCode - code of the error being displayed
'    ' Optional Arguments
'    ' none
'    ' Return Value
'    ' none
'    '*************************************************************************************
'    Public Sub ShowErr(lCode As Long)
'        If lCode = SUCCESS Then Exit Sub
'        Select Case lCode
'        Case Is = ACCESS_DENIED
'            MsgBox "Security access mask insufficient for operation (WinNT), or operation not permitted on key.", vbCritical + vbOKOnly, "Access Denied"
'        Case Is = BAD_NETPATH
'            MsgBox "The network path is invalid.", vbCritical + vbOKOnly, "Bad Net Path"
'        Case Is = BAD_PATHNAME
'            MsgBox "Invalid registry path.", vbCritical + vbOKOnly, "Bad Path Name"
'        Case Is = BADDB
'            MsgBox "The registry is damaged.", vbCritical + vbOKOnly, "Bad Database"
'        Case Is = BADKEY
'            MsgBox "Invalid handle to a registry key.", vbCritical + vbOKOnly, "Bad Key"
'        Case Is = CALL_NOT_IMPLEMENTED
'            MsgBox "Function not implemented for Win9x in the Win32 API.", vbCritical + vbOKOnly, "Call Not Implemented"
'        Case Is = CANTOPEN
'            MsgBox "Unable to open a registry key.", vbCritical + vbOKOnly, "Can't Open"
'        Case Is = CANTREAD
'            MsgBox "Unable to access a dynamic key.", vbCritical + vbOKOnly, "Can't Read"
'        Case Is = CANTWRITE
'            MsgBox "Unable to write to a dynamic registry key.", vbCritical + vbOKOnly, "Can't Write"
'        Case Is = DLL_INIT_FAILED
'            MsgBox "Unable to establish RPC connection. Make sure Remote Registry Services are installed.", vbCritical + vbOKOnly, "DLL Initialization Failed"
'        Case Is = FILE_NOT_FOUND
'            MsgBox "The named value or key does not exist.", vbCritical + vbOKOnly, "Key or Value Not Found"
'        Case Is = INSUFFICIENT_BUFFER
'            MsgBox "The value written by the function during remote registry access exceeds the size of the buffer.", vbCritical + vbOKOnly, "Insufficient Buffer"
'        Case Is = INVALID_HANDLE
'            MsgBox "The handle passed to the function is invalid.", vbCritical + vbOKOnly, "Invalid Handle"
'        Case Is = INVALID_PARAMETER
'            MsgBox "There is an extra parameter, a missing parameter, or a parameter that doesn't correspond to its declared length, its expected data type, or its expected value.", vbCritical + vbOKOnly, "Invalid Parameter"
'        Case Is = KEY_DELETED
'            MsgBox "The open key has been deleted.", vbCritical + vbOKOnly, "Key Deleted"
'        Case Is = LOCK_FAILED
'            MsgBox "The registry's internal read/write locking scheme failed.", vbCritical + vbOKOnly, "Lock Failed"
'        Case Is = MORE_DATA
'            MsgBox "The value written by the function to a buffer exceeds the size of the buffer.", vbCritical + vbOKOnly, "More Data"
'        Case Is = NO_MORE_ITEMS
'            MsgBox "There are no additional items to enumerate.", vbCritical + vbOKOnly, "No More Items"
'        Case Is = NOT_REGISTRY_FILE
'            MsgBox "Unable to recognize the format of the file to be loaded or imported.", vbCritical + vbOKOnly, "Not Registry File"
'        Case Is = REGISTRY_CORRUPT
'            MsgBox "Unable to access some portion of the registry.", vbCritical + vbOKOnly, "Registry Corrupt"
'        Case Is = REGISTRY_IO_FAILED
'            MsgBox "Unsuccessful attempt to read from or write to an external file.", vbCritical + vbOKOnly, "Registry I/O Failed"
'        Case Is = REGISTRY_RECOVERED
'            MsgBox "One of the registry's files had to be recovered by using an alternate copy.", vbCritical + vbOKOnly, "Registry Recovered"
'        Case Is = TRANSFER_TOO_LONG
'            MsgBox "The length of the requested data exceeds the system's limit of one megabyte.", vbCritical + vbOKOnly, "Transfer Too Long"
'        Case Else
'            MsgBox "Unrecognized Error", vbCritical + vbOKOnly, "Error"
'        End Select
'    End Sub
'
'
'    Usage
'
'     '********************************************************************
'    ' Programmer : Patrick S. Seymour
'    '
'    ' Created : January 24, 1999
'    ' Last Modified : January 24, 1999
'    '
'    ' Purpose : This function uses the Windows API to access the registry's Shell Folders key
'    ' for the current user. It returns the path that's stored in the value named
'    ' sFolder.
'    '********************************************************************
'    Public Function GetShellFolder(sFolder As String) As String
'
'    ' handle of the registry key to be accessed
'        Dim hKey As Long
'    ' path of the folder being sought (returned by function)
'        Dim sFolderName As String
'    ' path of the Shell Folders key in the registry
'        Dim sShellFoldersPath As String
'    ' define the Shell Folders path
'        sShellFoldersPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
'    ' attempt to open the key; on failure, exit the function
'        If (OpenKey(HKCU, sShellFoldersPath, hKey, vbNullString, NOT_USED, KEY_READ, NOT_USED)) <> SUCCESS Then Exit Function
'    ' declare a buffer large enough to hold the data being accessed
'        sFolderName = Space(255)
'    ' query the registry for the desired value
'        GetStringValue hKey, sFolder, sFolderName
'    ' close the key
'        CloseKey hKey
'    ' trim the spaces from the left and right of the path value
'        sFolderName = Trim(sFolderName)
'    ' remove the null character from the end of the string
'        sFolderName = Left(sFolderName, Len(sFolderName) - 1)
'    'if there is a backslash on the end of the path, remove it
'        If Right(sFolderName, 1) = "\" Then sFolderName = Left(sFolderName, Len(sFolderName) - 1)
'    ' return the value found in the registry
'        GetShellFolder = sFolderName
'
'    End Function
'


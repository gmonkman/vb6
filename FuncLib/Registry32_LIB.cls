VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long
 
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
    "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
    As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
    As Long, phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
    "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
    Long) As Long

Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
    As String, lpcbData As Long) As Long

Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, lpData As _
    Long, lpcbData As Long) As Long

Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
    As Long, lpcbData As Long) As Long

Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
    String, ByVal cbData As Long) As Long

Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
    ByVal cbData As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias _
    "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
    "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" Alias _
    "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
    lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, _
    lpcbData As Long) As Long

Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)

Public Enum Reg32
    REG_SZ = 1
    REG_DWORD = 4
    REG_BINARY = 3
    HKEY_CLASSES_ROOT = -2147483648#
    HKEY_CURRENT_USER = -2147483647
    HKEY_LOCAL_MACHINE = -2147483646
    HKEY_USERS = -2147483645
    REG_OPTION_NON_VOLATILE = 0
    
    ERROR_NONE = 0
    ERROR_BADDB = 1
    ERROR_BADKEY = 2
    ERROR_CANTOPEN = 3
    ERROR_CANTREAD = 4
    ERROR_CANTWRITE = 5
    ERROR_OUTOFMEMORY = 6
    ERROR_INVALID_PARAMETER = 7
    ERROR_ACCESS_DENIED = 8
    ERROR_INVALID_PARAMETERS = 87
    ERROR_NO_MORE_ITEMS = 259

    KEY_SET_VALUE = &H2
    KEY_QUERY_VALUE = &H1
    KEY_NOTIFY = &H10
    KEY_EVENT = &H1
    KEY_ENUMERATE_SUB_KEYS = &H8
    KEY_CREATE_SUB_KEY = &H4
    KEY_CREATE_LINK = &H20
    KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
    KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
    KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
End Enum

Public Enum eBaseKey
    HKEY_CLASSES_ROOT = -2147483648#
    HKEY_CURRENT_USER = -2147483647
    HKEY_LOCAL_MACHINE = -2147483646
    HKEY_USERS = -2147483645
End Enum

Public Enum eValueType
    REG_SZ = 1
    REG_DWORD = 4
    REG_BINARY = 3
End Enum

Public Enum eBK
    bkHKEY_CLASSES_ROOT = -2147483648#
    bkHKEY_CURRENT_USER = -2147483647
    bkHKEY_LOCAL_MACHINE = -2147483646
    bkHKEY_USERS = -2147483645
End Enum

Public Enum eVT
    vtREG_SZ = 1
    vtREG_DWORD = 4
    vtREG_BINARY = 3
End Enum


Public Function GetValueNumric(ByVal lBaseKey As eBK, ByVal sKeyName As String, ByVal sValueName As String) As Long
Dim l&
Dim vIn As Variant
l = Me.QueryValue(lBaseKey, sKeyName, sValueName, vIn)
If l <> 0 Then 'fail
    GetValueNumric = 0
Else
    GetValueNumric = CLng(Val(vIn))
End If
End Function

'Query registry value
'QueryValueEx only supports REG_SZ and REG_DWORD types
Public Function QueryValue(lBaseKey As eBaseKey, sKeyName As String, sValueName As String, ByRef vKeyValue As Variant) As Long
Dim lRetVal As Long 'result of the API functions
Dim hKey As Long 'handle of opened key
Dim vValue As Variant 'setting of queried value

On Error GoTo errh

RemoveLeftSlash sKeyName
lRetVal = RegOpenKeyEx(lBaseKey, sKeyName, 0, KEY_READ, hKey)
If lRetVal = 0 Then
    QueryValue = QueryValueEx(hKey, sValueName, vKeyValue)
Else
    QueryValue = lRetVal
End If
    
On Error Resume Next
RegCloseKey (hKey)

Exit Function

errh:
QueryValue = vbObjectError
Err.Clear
Exit Function
End Function


Public Function SetKeyValue(lBaseKey As eBaseKey, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As eValueType) As Long
Dim lRetVal As Long         'result of the SetValueEx function
Dim hKey As Long         'handle of open key
On Error GoTo errh
'open the specified key

lRetVal = RegOpenKeyEx(lBaseKey, sKeyName, 0, KEY_WRITE, hKey)
If lRetVal = 0 Then
    SetKeyValue = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
Else
    SetKeyValue = lRetVal
End If

On Error Resume Next
RegCloseKey (hKey)

Exit Function

errh:
SetKeyValue = vbObjectError
Err.Clear
Exit Function

End Function


Friend Function SetValueEx(ByVal hKey As Long, ByVal sValueName As String, ByVal lType As Long, ByVal vValue As Variant) As Long
Dim lValue As Long
Dim sValue As String
Select Case lType
    Case Reg32.REG_SZ
        sValue = vValue & Chr$(0)
        SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
    Case Reg32.REG_DWORD
        lValue = vValue
        SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
End Select
End Function


Public Function CreateNewKey(ByVal sNewKeyName As String, ByVal lPredefinedKey As eBaseKey) As Long
Dim hNewKey As Long         'handle to the new key
Dim lRetVal As Long         'result of the RegCreateKeyEx function
On Error GoTo errh

RemoveLeftSlash sNewKeyName

CreateNewKey = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, Reg32.REG_OPTION_NON_VOLATILE, Reg32.KEY_WRITE, 0&, hNewKey, lRetVal)

On Error Resume Next
RegCloseKey (hNewKey)

Exit Function

errh:
CreateNewKey = vbObjectError
Err.Clear
Exit Function
End Function

Public Function ValueExists(ByVal BaseKey As eBaseKey, ByVal keyname As String, ByVal vCheckValue As Variant, ByRef lOutErr As Long) As Boolean
Dim l&, cnt&
Dim Key$
Dim arrIn() As String
Dim vIn As Variant
On Error GoTo errh

l = EnumValues(BaseKey, keyname, arrIn)
If l <> 0 Then
    lOutErr = vbObjectError
    ValueExists = False
    Exit Function
End If

If UBound(arrIn) = 0 And arrIn(0) = "" Then
    ValueExists = False
    Exit Function
End If

For cnt = 0 To UBound(arrIn)
    Key = arrIn(cnt)
    l = Me.QueryValue(BaseKey, keyname, Key, vIn)
    If VarType(vCheckValue) = vbLong Or VarType(vIn) = vbInteger Then
        If VarType(vIn) = vbLong Or VarType(vIn) = vbInteger Then
            If CLng(vIn) = CLng(vCheckValue) Then
                ValueExists = True
                Exit Function
            End If
        Else
            If LCase$(CStr(vIn)) = LCase$(CStr(vCheckValue)) Then
                ValueExists = True
                Exit Function
            End If
        End If
    Else
        If CStr(vCheckValue) = CStr(vIn) Then
            ValueExists = True
            Exit Function
        End If
    End If
Next cnt
ValueExists = False
Exit Function

errh:
lOutErr = Err.Number
ValueExists = False
Exit Function
End Function

Friend Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, ByRef vValue As Variant) As Long
Dim cch As Long
Dim lrc As Long
Dim lType As Long
Dim lValue As Long
Dim sValue As String

On Error GoTo QueryValueExError

' Determine the size and type of data to be read
lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
If lrc <> ERROR_NONE Then
       QueryValueEx = lrc
       Exit Function
End If

Select Case lType
    ' For strings
    Case Reg32.REG_SZ:
        sValue = String(cch, 0)
        lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
        If lrc = ERROR_NONE Then
            vValue = Left$(sValue, cch - 1)
        Else
            vValue = Empty
        End If
    ' For Numerical
    Case Reg32.REG_DWORD:
        lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
        If lrc = ERROR_NONE Then vValue = lValue
    Case Else
        'all other data types not supported
        lrc = -1
End Select

QueryValueExExit:
       QueryValueEx = lrc
       Exit Function
QueryValueExError:
       Resume QueryValueExExit
End Function


Public Function DeleteKey(lBaseName As eBaseKey, sKeyName As String) As Long
On Error GoTo errh
DeleteKey = RegDeleteKey(lBaseName, sKeyName)
Exit Function
errh:
DeleteKey = vbObjectError
Err.Clear
Exit Function
End Function


Public Function DeleteValue(lBaseKey As eBaseKey, sKeyName As String, sValueName As String) As Long
Dim lRetVal&, hKey&
On Error GoTo errh
lRetVal = RegOpenKeyEx(lBaseKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)

If lRetVal = 0 Then
    DeleteValue = RegDeleteValue(hKey, sValueName)
Else
    DeleteValue = lRetVal
End If

Exit Function
errh:
DeleteValue = vbObjectError
Err.Clear
Exit Function
End Function


Public Function EnumValues(lBaseKey As eBaseKey, sKeyName As String, ByRef arrReturn() As String) As Long
Dim lngResult&, kKey&, lngCurIdx&, lngDataLen&, lngValueLen&, lngData&
Dim strValue As String
Dim booFirstPass As Boolean
On Error GoTo errh

ReDim arrReturn(0)
lngCurIdx = 0: booFirstPass = True

lngResult = RegOpenKeyEx(lBaseKey, sKeyName, 0&, KEY_READ, kKey)

If lngResult > 0 Then
    EnumValues = lngResult
    Exit Function
End If

Do
    lngValueLen = 2000
    strValue = String(lngValueLen, 0)
    lngDataLen = 2000

    lngResult = RegEnumValue(kKey, lngCurIdx, ByVal strValue, lngValueLen, 0&, _
                  Reg32.REG_SZ, ByVal lngData, lngDataLen)
          
    If Not booFirstPass Then
        ReDim Preserve arrReturn(UBound(arrReturn) + 1)
    Else
        booFirstPass = False
    End If

    If lngResult = Reg32.ERROR_NONE Then
        arrReturn(lngCurIdx) = Left$(strValue, lngValueLen)
    End If
    
    lngCurIdx = lngCurIdx + 1

Loop While lngResult = ERROR_NONE

On Error Resume Next
Call RegCloseKey(kKey)

EnumValues = False

Exit Function
errh:
EnumValues = vbObjectError
Err.Clear
Exit Function
End Function


Friend Sub RemoveLeftSlash(ByRef strTemp)
On Error Resume Next
If Left$(strTemp, 1) = "\" Then strTemp = Right$(strTemp, Len(strTemp) - 1)
End Sub

'fills outsystemdsnlist with datasource name<seperator>odbc driver list
Public Function ODBCGetSystemDSNList(ByRef OutSystemDSNList() As String, Optional ByVal Seperator = ":") As Boolean
Dim l&, cnt&
Dim vVal As Variant
Dim s$, sTmp$
Const keyname = "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources"
Dim sKeys() As String
On Error GoTo errh
ReDim OutSystemDSNList(0)
l = Me.EnumValues(Reg32.HKEY_LOCAL_MACHINE, keyname, sKeys())

If l > 0 Then: ODBCGetSystemDSNList = False: Exit Function

For cnt = 0 To UBound(sKeys)
    s = sKeys(cnt)
    l = Me.QueryValue(Reg32.HKEY_LOCAL_MACHINE, keyname, s, vVal)
    If l = 0 Then
        sTmp = s & Seperator & CStr(vVal)
        ReDim Preserve OutSystemDSNList(cnt)
        OutSystemDSNList(cnt) = sTmp
    End If
Next cnt

ODBCGetSystemDSNList = True

Exit Function

errh:
ODBCGetSystemDSNList = False
Err.Clear
Exit Function
End Function


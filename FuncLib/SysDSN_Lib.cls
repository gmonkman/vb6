VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SysDSN_Lib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
    ByVal cbData As Long) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
    "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
    
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
    String, ByVal cbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long
    
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias _
    "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
    "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
    ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
    Long) As Long
    
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
    As Long, lpcbData As Long) As Long

Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
    As String, lpcbData As Long) As Long

Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
    String, ByVal lpReserved As Long, lpType As Long, lpData As _
    Long, lpcbData As Long) As Long
    
Private Declare Function SQLInstallerError Lib "ODBCCP32.DLL" _
        (ByVal iError As Integer, ByRef pfErrorCode As Long, _
        ByVal lpszErrorMsg As String, ByVal bufErrorMsgMax As Integer, _
        ByRef pcbErrorMessage As Integer) As Long
        
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
          (ByVal hwndParent As Long, ByVal fRequest As Long, _
          ByVal lpszDriver As String, ByVal lpszAttributes As String) _
          As Long
          
Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
          
          
Public Enum eDSNAction
    ODBC_ADD_DSN = 1
    ODBC_CONFIG_DSN = 2
    ODBC_REMOVE_DSN = 3
    ODBC_ADD_SYS_DSN = 4       'Add data source
    ODBC_CONFIG_SYS_DSN = 5    'Configure (edit) data source
    ODBC_REMOVE_SYS_DSN = 6
End Enum

Public Enum eDSNDriver
    SQLServer_
    Access97_
End Enum

Public Enum eDSNSQLSecurity
    NT
    SQLServer
End Enum

'keep for backward compat
Public Enum eSQLNetLibs
    TCPIP
    IPX_SPX
    Vines
    NamedPipes
    MultiProt
    Default
    DO_NOT_CONFIG
End Enum

Public Enum eSQLNetworkLibs
    nlTCPIP
    nlIPX_SPX
    nlVines
    nlNamedPipes
    nlMultiProt
    nlDefault
    nlDO_NOT_CONFIG
End Enum

Private Enum Reg32
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

Private Enum eBK
    HKEY_CLASSES_ROOT = -2147483648#
    HKEY_CURRENT_USER = -2147483647
    HKEY_LOCAL_MACHINE = -2147483646
    HKEY_USERS = -2147483645
End Enum

Private Enum eVT
    REG_SZ = 1
    REG_DWORD = 4
    REG_BINARY = 3
End Enum

Private Const SQL_MAX_MESSAGE_LENGTH As Long = 512

Public Enum eSQLInstCodes
    SQL_SUCCESS = 0
    SQL_SUCCESS_WITH_INFO = 1
End Enum

Public Function GetNetworkLibEnumFromString(ByVal sLib As String) As eSQLNetLibs
On Error GoTo errh
Select Case sLib
    Case "tcpip", "tcp", "ip", "tcp/ip"
        GetNetworkLibEnumFromString = TCPIP
    Case "ipx", "ipxspx", "spx", "ipx/spx", "novell"
        GetNetworkLibEnumFromString = IPX_SPX
    Case "vines", "banyan", "banyanvines", "banyan vines"
        GetNetworkLibEnumFromString = Vines
    Case "pipes", "named", "named pipes", "namedpipes"
        GetNetworkLibEnumFromString = NamedPipes
    Case "multi", "mixed", "multiprotocol", "multi protocol"
        GetNetworkLibEnumFromString = MultiProt
    Case "defalt"
        GetNetworkLibEnumFromString = NamedPipes
    Case Else
        GetNetworkLibEnumFromString = NamedPipes
End Select
Exit Function
errh:
 On Error Resume Next
 GetNetworkLibEnumFromString = Default
End Function

Public Function GetNetworkLibName(ByVal SQLNetLib As eSQLNetLibs) As String
On Error GoTo errh
Select Case SQLNetLib
    Case eSQLNetLibs.IPX_SPX
        GetNetworkLibName = "DBMSSPXN"  'ok
    Case eSQLNetLibs.NamedPipes
        GetNetworkLibName = "DBNMPNTW"  'ok, need to also specify the connection details otherwise defaults to TCPIP
    Case eSQLNetLibs.MultiProt
        GetNetworkLibName = "DBMSRPCN"  'ok
    Case eSQLNetLibs.TCPIP
        GetNetworkLibName = "DBMSSOCN"  'ok
    Case eSQLNetLibs.Vines
        GetNetworkLibName = "DBMSVINN"  'ok
    Case eSQLNetLibs.Default
        GetNetworkLibName = "DBMSSOCN"  'ok
End Select
Exit Function
errh:
GetNetworkLibName = "DBMSSOCN"
End Function

Private Function DeleteValue(lBaseKey As eBK, sKeyName As String, sValueName As String) As Long
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

Private Sub RemoveLeftSlash(ByRef strTemp)
On Error Resume Next
If Left$(strTemp, 1) = "\" Then strTemp = Right$(strTemp, Len(strTemp) - 1)
End Sub

Private Function SetKeyValue(lBaseKey As eBK, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As eVT) As Long
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


Private Function SetValueEx(ByVal hKey As Long, ByVal sValueName As String, ByVal lType As Long, ByVal vValue As Variant) As Long
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

'Query registry value
'QueryValueEx only supports REG_SZ and REG_DWORD types
Private Function QueryValue(lBaseKey As eBK, sKeyName As String, sValueName As String, ByRef vKeyValue As Variant) As Long
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

Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, ByRef vValue As Variant) As Long
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


Private Function DSNDriverGetString(ByVal DSNDrv As eDSNDriver) As String
On Error GoTo errh
Select Case DSNDrv
    Case eDSNDriver.SQLServer_
        DSNDriverGetString = "SQL Server"
    Case eDSNDriver.Access97_
        DSNDriverGetString = "Microsoft Access Driver (*.mdb)"
End Select
Exit Function
errh:
DSNDriverGetString = ""
End Function

Public Function DSNExists(ByVal DSNName As String, ByRef lOutErr As Long) As Boolean
Dim l&
Dim vIn As Variant
Const DSNKEY = "software\odbc\odbc.ini\odbc data sources"
On Error GoTo errh
l = QueryValue(eBK.HKEY_LOCAL_MACHINE, DSNKEY, DSNName, vIn) 'use driver value to check existance
If l <> 0 Then
    DSNExists = False
Else
    DSNExists = True
End If
Exit Function
errh:
lOutErr = Err.Number
End Function

Public Function ErrorsGetList(ByRef ErrList() As String) As Boolean
Dim cnt%, iSizeOfStr%, i%
Dim lInErrCode&
Dim lIn
Dim l&
Dim bWasErr As Boolean
Dim ErrMsg As String * SQL_MAX_MESSAGE_LENGTH
Dim ErrTrim$, sTmp$
On Error GoTo errh
ReDim ErrList(0)
For i = 1 To 8 'looks odd but apparently stores only the last 8 errors!
    l = SQLInstallerError(i, lInErrCode, ErrMsg, SQL_MAX_MESSAGE_LENGTH, iSizeOfStr)
    If l = 0 Then
        ReDim Preserve ErrList(cnt)
        ErrTrim = Trim$(Left$(ErrMsg, iSizeOfStr))
        ErrList(cnt) = ErrTrim
        cnt = cnt + 1
        bWasErr = True
    End If
Next i
If bWasErr Then
    ErrorsGetList = True
Else
    ErrorsGetList = False
End If
Exit Function
errh:
ErrorsGetList = False
End Function


'note that this could fail (return false) on nt/2000 machines where the user does not
'have the permissions to ammend odbc entries
Public Function DSNConfigSQLServer(ByVal DSNName$, ByVal ServerName$, ByVal DatabaseName$, _
 Optional ByVal description = "APD Technical (01928 739808) - Establish Enterprise Datasource.", _
 Optional ByVal Security As eDSNSQLSecurity = eDSNSQLSecurity.SQLServer, _
 Optional ByVal UID As String = "sa", Optional ByVal NetLib As eSQLNetLibs = eSQLNetLibs.NamedPipes, _
 Optional ParentHwnd As Long = 0&) As Boolean
Dim Action As eDSNAction
Dim lAct&, lRet&, lInErr&
Dim strDriver$, sAtr$, sKey$, sLib$, sAdd$
Dim b As Boolean
Const ODBC_KEY = "software\odbc\odbc.ini"
On Error GoTo errh
If NetLib = DO_NOT_CONFIG Then Exit Function
sLib = GetNetworkLibName(NetLib)

b = Me.DSNExists(DSNName, lInErr)
If lInErr <> 0 Then Err.Raise lInErr
If b Then
    Action = ODBC_CONFIG_SYS_DSN
Else
    Action = ODBC_ADD_SYS_DSN
End If
lAct = CLng(Action)
strDriver = DSNDriverGetString(eDSNDriver.SQLServer_)

If GetNetworkLibName(eSQLNetLibs.NamedPipes) = sLib Then
    sAdd = "\\" & ServerName & "\pipe\sql\query"
End If

sAtr = "SERVER=" & ServerName & Chr$(0)
sAtr = sAtr & "DSN=" & DSNName & Chr$(0)
sAtr = sAtr & "DATABASE=" & DatabaseName & Chr$(0)
sAtr = sAtr & "Network=" & sLib & Chr$(0)
If sAdd <> "" Then sAtr = sAtr & "Address=" & sAdd & Chr$(0)
sAtr = sAtr & "DESCRIPTION=" & description
sAtr = sAtr & Chr$(0) & Chr$(0)
lRet = SQLConfigDataSource(ParentHwnd, lAct&, strDriver, sAtr)
If lRet = 0 Then GoTo errh

If Security = eDSNSQLSecurity.NT Then
    sKey = ODBC_KEY & "\" & DSNName
    lRet = SetKeyValue(eBK.HKEY_LOCAL_MACHINE, sKey, "Trusted_Connection", "Yes", eVT.REG_SZ)
    If lRet <> 0 Then GoTo errh
Else
    sKey = ODBC_KEY & "\" & DSNName
    lRet = DeleteValue(eBK.HKEY_LOCAL_MACHINE, sKey, "Trusted_Connection")
    lRet = SetKeyValue(eBK.HKEY_LOCAL_MACHINE, sKey, "LastUser", UID, eVT.REG_SZ)
End If

DSNConfigSQLServer = True
Exit Function

errh:
DSNConfigSQLServer = False
End Function



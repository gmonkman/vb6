VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SysInfo_Lib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type udtDiskInformation
    SectorsPerCluster As Double
    BytesPerSector As Double
    NumberOfFreeClusters As Double
    TotalNumberOfClusters As Double
    BytesPerCluster As Double
    NumFreeBytes As Double
    NumFreeKiloBytes As Double
    FreeMegabytes As Double
End Type

Public Type udtDisk
    Name As String
    uDiskInfo As udtDiskInformation
End Type

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" _
   Alias "GetLocaleInfoA" (ByVal Locale As Long, _
   ByVal LCType As Long, ByVal lpLCData As String, _
   ByVal cchData As Long) As Long
Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()

Private Const LOCALE_SCURRENCY = &H14
Private Const LOCALE_IDATE = &H21        '  short date format ordering
Private Const LOCALE_ILDATE = &H22

Public Enum eDateFormat
    MDY = 0
    DMY = 1
    YMD = 2
End Enum

Public Enum eOS
    osUnknown
    osWin9x
    osWinNT
End Enum

Public Function SystemLocaleDateIsDMY(ByRef lErrOut As Long) As Boolean
Dim lRet As Long
Dim lpLCDataVar As String * 2
Dim l&

Dim Locale As Long
Dim lRet1 As Long

On Error GoTo errh
Locale = GetUserDefaultLCID()

lRet = GetLocaleInfo(Locale, LOCALE_IDATE, lpLCDataVar, 0) 'short date format
lRet1 = GetLocaleInfo(Locale, LOCALE_IDATE, lpLCDataVar, lRet)
l = Val(lpLCDataVar)
If l = eDateFormat.DMY Then
    lRet = GetLocaleInfo(Locale, LOCALE_ILDATE, lpLCDataVar, 0) 'long date format
    lRet1 = GetLocaleInfo(Locale, LOCALE_ILDATE, lpLCDataVar, lRet)
    l = Val(lpLCDataVar)
    If l = eDateFormat.DMY Then
        SystemLocaleDateIsDMY = True
    Else
        SystemLocaleDateIsDMY = False
    End If
Else
    SystemLocaleDateIsDMY = False
End If

lErrOut = 0
Exit Function
errh:
lErrOut = Err.Number
Err.Clear
Exit Function
End Function


Public Function EnvironGetUserName() As String
Dim sTmp$
Dim sOut As String * 255
Dim l&
Dim cStrings As New Strings_LIB
Dim b As Boolean
On Error GoTo errh
l = 255
b = GetUserName(sOut, l)
If l > 0 Then
    sTmp = Trim$(Left$(sOut, l))
End If
cStrings.CleanString sTmp
EnvironGetUserName = sTmp
GoSub LC
Exit Function

errh:
Resume e
e:
On Error Resume Next
GoSub LC
EnvironGetUserName = ""
Err.Clear
Exit Function

LC:
On Error Resume Next
D cStrings
Return
End Function

Private Sub D(obj As Object)
On Error Resume Next
If TypeOf obj Is ADODB.Recordset Or TypeOf obj Is ADODB.Connection Then
    obj.Close
End If
If Not obj Is Nothing Then Set obj = Nothing
Err.Clear
End Sub


Public Function EnvironGetTempDir() As String
Dim strTmp$
On Error GoTo errh
strTmp = Environ$("tmp")
If strTmp = "" Then strTmp = Environ$("temp")
EnvironGetTempDir = strTmp
Exit Function
errh:
Resume e
e:
On Error Resume Next
EnvironGetTempDir = ""
Err.Clear
Exit Function
End Function

Public Function EnvironGetOS() As String
Dim sTmp$
On Error GoTo errh
sTmp = Environ$("OS")
If sTmp = "" Then sTmp = "Win9x"
EnvironGetOS = sTmp
Exit Function
errh:
Resume e
e:
On Error Resume Next
Err.Clear
EnvironGetOS = ""
Exit Function
End Function

Public Function EnvironGetOSEnum() As eOS
Dim sTmp$
On Error GoTo errh
sTmp = Environ$("OS")
If sTmp = "" Then
    EnvironGetOSEnum = osWin9x
Else
    EnvironGetOSEnum = osWinNT
End If
Exit Function
errh:
 On Error Resume Next
 EnvironGetOSEnum = osUnknown
 Err.Clear
End Function

Public Function EnvironGetSystemDir() As String
Dim sDIR As String * 255
Dim sOut$
Dim l&
On Error GoTo errh
l = GetSystemDirectory(sDIR, 255)
If l > 0 Then
    sOut = Trim$(Left$(sDIR, l))
End If
EnvironGetSystemDir = sOut
Exit Function
errh:
Resume e
e:
On Error Resume Next
EnvironGetSystemDir = ""
Err.Clear
Exit Function
End Function

Public Function EnvironGetLogonServer() As String
Dim sOut$
On Error GoTo errh
sOut = Environ$("logonserver")
EnvironGetLogonServer = sOut
Exit Function
errh:
 On Error Resume Next
 EnvironGetLogonServer = ""
 Err.Clear
Exit Function
End Function

Public Function EnvironGetComputerName() As String
Dim sDIR As String * 255
Dim sOut$
Dim l&
Dim b As Boolean
On Error GoTo errh
l = 255
b = GetComputerName(sDIR, l)
If b Then
    sOut = Trim$(Left$(sDIR, l))
End If
EnvironGetComputerName = sOut
Exit Function
errh:
Resume e
e:
On Error Resume Next
EnvironGetComputerName = ""
Err.Clear
Exit Function
End Function

Public Function SystemCurrencyGetSymbol() As String
On Error GoTo errh
Dim Symbol As String
Dim iRet1 As Long
Dim iRet2 As Long
Dim lpLCDataVar As String
Dim Pos As Integer
Dim Locale As Long
On Error GoTo errh
Locale = GetUserDefaultLCID()
iRet1 = GetLocaleInfo(Locale, LOCALE_SCURRENCY, lpLCDataVar, 0)
Symbol = String$(iRet1, 0)
iRet2 = GetLocaleInfo(Locale, LOCALE_SCURRENCY, Symbol, iRet1)
Pos = InStr(Symbol, Chr$(0))
If Pos > 0 Then
   Symbol = Left$(Symbol, Pos - 1)
End If
SystemCurrencyGetSymbol = Symbol
Exit Function

errh:
 On Error Resume Next
 SystemCurrencyGetSymbol = "£"
End Function

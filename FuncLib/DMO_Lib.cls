VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DMO_Lib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eSQLEngineVer
    SQL_7
    SQL_2000
End Enum

Public Enum eSQLEngineVerBitwise
    bwNone = 0
    bwSQL_7 = 1
    bwSQL_2000 = 2
End Enum

Public Function SQLServer2000ServerList(OutList() As String) As Boolean
Dim dmoNL As SQLDMO.NameList
Dim cnt&
Dim i&
On Error GoTo errh
Set dmoNL = SQLDMO.Application.ListAvailableSQLServers
ReDim OutList(0)
If dmoNL.Count = 0 Then GoTo SKIP
For cnt = 0 To dmoNL.Count
    If Len(dmoNL.Item(cnt)) > 0 And dmoNL.Item(cnt) <> "." Then
        ReDim Preserve OutList(i)
        OutList(i) = dmoNL.Item(cnt)
        i = i + 1
    End If
Next cnt
SKIP:
SQLServer2000ServerList = True
Exit Function
errh:
SQLServer2000ServerList = False
End Function

Public Function ServerConnect(Optional ByVal sName As String = "(local)", Optional ByVal UseNTSecurity As Boolean = True, Optional LoginTimeout As Long = 30, Optional ByVal UID As String = "sa", Optional ByVal PWD As String = "", Optional ByVal SQLVer As eSQLEngineVer = SQL_7) As Object
Dim dmoServ As Object
On Error GoTo errh
If SQLVer = SQL_2000 Then
    Set dmoServ = CreateObject("SQLDMO.SQLServer2")
ElseIf SQLVer = SQL_7 Then
    Set dmoServ = CreateObject("SQLDMO.SQLServer")
Else
    Err.Raise vbObjectError
End If
dmoServ.LoginTimeout = LoginTimeout
If UseNTSecurity Then
    dmoServ.LoginSecure = True
    dmoServ.Connect sName
Else:
    dmoServ.LoginSecure = False
    dmoServ.Connect sName, UID, PWD
End If
Set ServerConnect = dmoServ
Exit Function

errh:
 On Error Resume Next
 Set ServerConnect = Nothing
 Err.Clear
Exit Function

End Function


Public Function DatabasesEnum(ByRef dmoServ As Object, ByRef OutDBList() As String, Optional ByVal SQLVer As eSQLEngineVer = SQL_7) As Boolean
Dim cnt&
Dim dmoDB As Object
On Error GoTo errh
If SQLVer = SQL_2000 Then
    Set dmoDB = CreateObject("SQLDMO.Database2")
ElseIf SQLVer = SQL_7 Then
    Set dmoDB = CreateObject("SQLDMO.Database")
Else
    Err.Raise vbObjectError
End If
ReDim OutDBList(0)
For Each dmoDB In dmoServ.Databases
    If Not dmoDB.SystemObject Then
        ReDim Preserve OutDBList(cnt)
        OutDBList(cnt) = dmoDB.Name
        cnt = cnt + 1
    End If
Next
DatabasesEnum = True
Exit Function

errh:
Resume e
e:
On Error Resume Next
DatabasesEnum = False
Err.Clear
Exit Function

LC:
On Error Resume Next
D dmoDB
Return

End Function

Public Function DMOVersionsAvailable() As eSQLEngineVerBitwise
Dim lRet&
Dim obj As Object
On Error Resume Next
Set obj = CreateObject("sqldmo.sqlserver")
If Not Err.Number Then
    lRet = lRet + eSQLEngineVerBitwise.bwSQL_7
End If
Set obj = Nothing
Err.Clear

Set obj = CreateObject("sqldmo.sqlserver2")
If Not Err.Number Then
    lRet = lRet + eSQLEngineVerBitwise.bwSQL_2000
End If
Set obj = Nothing
Err.Clear
DMOVersionsAvailable = lRet
End Function

Private Sub D(obj As Object)
On Error Resume Next
If TypeOf obj Is ADODB.Recordset Or TypeOf obj Is ADODB.Connection Then
    obj.Close
End If
If Not obj Is Nothing Then Set obj = Nothing
Err.Clear
End Sub

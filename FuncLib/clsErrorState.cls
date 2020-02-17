VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Error_Lib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Const resizevalue = 10
Private len_errorstk As Long
Private mstrErrorLog As String

Private Type TError
    Number As Long
    Description As String
    Source As String
End Type

Private errorstk() As TError

Public Property Get ErrorLogName() As String
ErrorLogName = mstrErrorLog
End Property


Public Function SeverityGetString(uSev As udtErrorSeverity) As String
On Error GoTo errh
Select Case uSev
    Case udtErrorSeverity.Critical
        SeverityGetString = "critical"
    Case udtErrorSeverity.Fatal
        SeverityGetString = "fatal"
    Case udtErrorSeverity.Major
        SeverityGetString = "major"
    Case udtErrorSeverity.Minor
        SeverityGetString = "minor"
    Case udtErrorSeverity.Recoverable
        SeverityGetString = "recoverable"
    Case udtErrorSeverity.Unrecoverable
        SeverityGetString = "unrecoverable"
    Case Else
        SeverityGetString = "unknown"
End Select
Exit Function
errh:
Resume e
e:
On Error Resume Next
SeverityGetString = "Unknown"
Err.Clear
Exit Function
End Function

Public Sub ErrorsClear(ByRef uErr() As udtErrorObject)
On Error Resume Next
ReDim uErr(0)
End Sub

Private Function ArrayIsEmpty(arr) As Boolean
Dim lUB&
On Error Resume Next
lUB = UBound(arr)
If Err.Number = 9 Or lUB = -1 Then 'subscript out of range
    ArrayIsEmpty = True
Else
    ArrayIsEmpty = False
End If
Err.Clear
End Function

Public Sub ErrorAdd(uErrs() As udtErrorObject, uErr As udtErrorObject, Optional ClearErrors As Boolean = False, Optional ByVal LogError As Boolean = True)
On Error GoTo errh
Dim b As Boolean
Dim lUB As Long
If ArrayIsEmpty(uErrs) Then ReDim uErrs(0)
If ClearErrors Then
    ReDim uErrs(0)
Else
    If uErrs(0).Number = 0 Then
    Else
        ReDim Preserve uErrs(UBound(uErrs) + 1)
    End If
End If
lUB = UBound(uErrs)
uErrs(lUB) = uErr
If LogError Then b = LogErrorToFile(uErr)
Exit Sub

errh:
Err.Clear
Exit Sub

End Sub

Public Sub ErrorDisplayUserFriendly(uarrErrs() As udtErrorObject, Optional ByVal Index As Long, Optional ByVal ExeName As String = "Current Application")
Dim str$
On Error Resume Next
If ArrayIsEmpty(uarrErrs) Then Exit Sub
If UBound(uarrErrs) > Index Then Index = UBound(uarrErrs) 'retrieve last if non specified
str = "An error has occured in " & uarrErrs(Index).Location & vbCrLf & vbCrLf & _
      uarrErrs(Index).Description & vbCrLf & vbCrLf & uarrErrs(Index).Advice
MsgBox str, vbOKOnly, ExeName
End Sub

Public Function LogErrorToFile(Optional uErrorObjectOrSimpleString As Variant) As Boolean
Dim cFSO As New FSO_LIB
Dim str$, sPCName$, sUser$, sOS$, sLog$, sDate$
Dim cSys As New SysInfo_Lib
Dim b As Boolean
On Error GoTo errh

sPCName = cSys.EnvironGetComputerName
sUser = cSys.EnvironGetUserName
sOS = cSys.EnvironGetOS
sDate = Format$(Now(), "dd/mmm/yyyy hh:mm:ss")
str = sDate & "," & sPCName & "," & sUser & "," & sOS & ","
If VarType(uErrorObjectOrSimpleString) = vbString Then
    sLog = str & uErrorObjectOrSimpleString
ElseIf VarType(uErrorObjectOrSimpleString) = vbUserDefinedType Then
    sLog = str & uErrorObjectOrSimpleString.Number & _
    "," & uErrorObjectOrSimpleString.VBDescription & _
    "," & uErrorObjectOrSimpleString.Description & _
    "," & uErrorObjectOrSimpleString.Location
Else
    sLog = str & "Undescribed Error"
End If
b = cFSO.WriteText(Me.ErrorLogName, sLog)
GoSub LC
Exit Function

errh:
Resume e
e:
On Error Resume Next
GoSub LC
Err.Clear
Exit Function


LC:
On Error Resume Next
D cFSO
D cSys
Err.Clear
Return

End Function

Private Sub D(obj As Object)
On Error Resume Next
If Not obj Is Nothing Then
    If TypeOf obj Is ADODB.Recordset Or TypeOf obj Is ADODB.Connection Then
        obj.Close
    End If
    Set obj = Nothing
End If
Err.Clear
End Sub

Public Sub PopError()
    If len_errorstk = 0 Then Exit Sub
    len_errorstk = len_errorstk - 1
    If errorstk(len_errorstk).Number = 0 Then
        Err.Clear
        Exit Sub
    End If
    If errorstk(len_errorstk).Source = "" Then
        If errorstk(len_errorstk).Description = "" Then
            Err.Raise errorstk(len_errorstk).Number
        Else
            Err.Raise errorstk(len_errorstk).Number, , errorstk(len_errorstk).Description
        End If
    Else
        If errorstk(len_errorstk).Description = "" Then
            Err.Raise errorstk(len_errorstk).Number, , errorstk(len_errorstk).Description
        Else
            Err.Raise errorstk(len_errorstk).Number, errorstk(len_errorstk).Source, errorstk(len_errorstk).Description
        End If
    End If
End Sub

Public Sub PushError(Optional ByVal Number As Long = 0, Optional ByVal Description As String = "", Optional ByVal Source = "")

    If UBound(errorstk) < len_errorstk Then
        ReDim Preserve errorstk(len_errorstk + resizevalue - 1)
    End If
    If Number = 0 Then
        errorstk(len_errorstk).Number = Err.Number
        errorstk(len_errorstk).Description = Err.Description
        errorstk(len_errorstk).Source = Err.Source
    Else
        errorstk(len_errorstk).Number = Number
        errorstk(len_errorstk).Description = Description
        errorstk(len_errorstk).Source = Source
    End If
    len_errorstk = len_errorstk + 1
End Sub

Private Sub Class_Initialize()
On Error Resume Next
Dim cFSO As New FSO_LIB
Dim cSt As New Strings_LIB
Dim cSysInfo As New SysInfo_Lib
Dim tmp As String
Dim b As Boolean
len_errorstk = 0
ReDim errorstk(0) As TError
tmp = cSysInfo.EnvironGetTempDir
mstrErrorLog = cSt.AddRightSlash_F(tmp) & cSt.AddRight_F(App.ExeName, ".") & "log"
D cSt
D cSysInfo
D cFSO
End Sub



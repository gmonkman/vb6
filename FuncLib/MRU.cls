VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MRU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mlShowTopN&
Private mrs As New ADODB.Recordset
Private mcADO As New ADO_LIB
Private mcReg As New Registry32
Private mcStrL As New Strings_LIB
Private mbooErrorsOccured As Boolean
Private mstrRegKey$
Private mstrLeftPadName$
Public Event RSRefreshed(Recordset As ADODB.Recordset)

Public Property Let LeftPadSimpleName(ByVal Pad$)
mstrLeftPadName$ = Pad
End Property
Public Property Get LeftPadSimpleName() As String
LeftPadSimpleName = mstrLeftPadName$
End Property

Public Property Get RegKey() As String
RegKey = mstrRegKey
End Property
Public Property Let RegKey(ByVal Key As String)
mstrRegKey = Key
End Property

Public Property Get ErrorsOccured() As Boolean
ErrorsOccured = mbooErrorsOccured
End Property

Public Property Get ShowTopN() As Long
ShowTopN = mlShowTopN
End Property
Public Property Let ShowTopN(ByVal TopN&)
mlShowTopN = TopN
End Property

Public Function MRUItemTruncate(ByVal sItem As String, Optional ByVal Threshhold As Long = 35) As String
Dim sFile$
Dim sSubDir$
Dim sPath$, sTmp$
Dim l&, cnt&
On Error GoTo errh
If Len(sItem) > Threshhold Then
    sFile = mcStrL.GetFileOf(sItem)
    sPath = mcStrL.GetPathOf_F(sItem)
    sSubDir = mcStrL.GetFileOf(sPath)
    l = mcStrL.SubStringCountOf(sItem, "\")
    For l = 1 To cnt
        sTmp = sTmp & "..\"
    Next
    sTmp = sTmp & sPath & "\" & sFile
Else
    sTmp = sItem
End If
MRUItemTruncate = sTmp
Exit Function
errh:
MRUItemTruncate = sItem
End Function

Public Function TopN(Optional ByVal NameOnly As Boolean = False) As String()
Dim arr() As String
Dim cnt&
Dim sTmp$
On Error GoTo errh
If Not RP(mrs) Then GoTo errh
mrs.Sort = "ind ASC"
mrs.MoveFirst
ReDim arr(0)
Do
    If mrs("ind").Value <= Me.ShowTopN Then
        ReDim Preserve arr(cnt)
        sTmp = CStr(mrs("Document").Value)
        If NameOnly Then sTmp = mcStrL.GetFileOf(sTmp)
        arr(cnt) = sTmp
        cnt = cnt + 1
    End If
Loop Until mrs.EOF
TopN = arr()
Exit Function
errh:
TopN = arr()
End Function

Public Function TopN_Truncated() As String()
Dim arr() As String
Dim cnt&
Dim sTmp$
On Error GoTo errh
If Not RP(mrs) Then GoTo errh
mrs.Sort = "ind ASC"
mrs.MoveFirst
ReDim arr(0)
Do
    If mrs("ind").Value <= Me.ShowTopN Then
        ReDim Preserve arr(cnt)
        sTmp = CStr(mrs("Document").Value)
        sTmp = Me.MRUItemTruncate(sTmp)
        arr(cnt) = sTmp
        cnt = cnt + 1
    End If
Loop Until mrs.EOF
TopN_Truncated = arr()
Exit Function
errh:
TopN_Truncated = arr()
End Function

Public Function MRU_FullList(Optional ByVal NameOnly As Boolean = False) As String()
Dim arr() As String
Dim cnt&
Dim sTmp$
On Error GoTo errh
If Not RP(mrs) Then GoTo errh
mrs.Sort = "ind ASC"
mrs.MoveFirst
ReDim arr(0)
Do
    ReDim Preserve arr(cnt)
    sTmp = CStr(mrs("Document").Value)
    If NameOnly Then sTmp = mcStrL.GetFileOf(sTmp)
    arr(cnt) = sTmp
    cnt = cnt + 1
    mrs.MoveNext
Loop Until mrs.EOF
MRU_FullList = arr()
Exit Function
errh:
MRU_FullList = arr()
End Function

Public Function AmmendMRU(ByVal Document As String) As Boolean
Dim sql$
Dim l&
Dim rsTmp As New ADODB.Recordset
Dim bk As Variant
On Error GoTo errh
If RP(mrs) Then
    mrs.Sort = "document ASC"
    If mcADO.ADOFind(SearchFromStart, mrs, "Document='" & Document & "'") Then 'existing item, move it!
        l = mrs("Ind").Value
        mrs("Ind").Value = 9999
        mrs.Update
        mrs.MoveFirst
        Do
            If mrs("Ind").Value < l And mrs("ind").Value <> 9999 Then
                mrs("ind").Value = mrs("ind").Value + 1
                mrs.Update
            ElseIf mrs("ind").Value = 9999 Then
                mrs("ind").Value = 0 '0 based order
                mrs.Update
            End If
            mrs.MoveNext
        Loop Until mrs.EOF
    Else 'new item
        mrs.MoveFirst
        Do 'move em all down 1!
            l = mrs("ind").Value
            mrs("Ind").Value = l + 1
            mrs.Update
            mrs.MoveNext
        Loop Until mrs.EOF
        mrs.AddNew
        mrs("Document").Value = Document
        mrs("ind").Value = 0
        mrs.Update
    End If
Else
    mrs.AddNew
    mrs("Document").Value = Document
    mrs("ind").Value = 0
    mrs.Update
End If
RaiseEvent RSRefreshed(mrs)
AmmendMRU = True
Exit Function
errh:
AmmendMRU = False
End Function

Private Function RSInit() As Boolean
On Error GoTo errh
If mrs.State <> ObjectStateEnum.adStateClosed Then mrs.Close
With mrs.Fields
    .Append "Document", adVarChar, 255
    .Append "Ind", adInteger, , adFldIsNullable
    .Append "Name", adVarChar, 255
End With
mrs.Open
RSInit = True
Exit Function
errh:
RSInit = False
End Function

Public Property Get Recordset() As ADODB.Recordset
Set Recordset = mrs
End Property

Public Function RSRefresh() As Boolean
Dim b As Boolean
On Error GoTo errh
If Exist_(mrs) Then
    If mrs.State <> ObjectStateEnum.adStateClosed Then mrs.Close
Else
    Set mrs = New ADODB.Recordset
End If
b = RSInit
If b Then b = RegReadMRU
If b Then RaiseEvent RSRefreshed(mrs)
RSRefresh = b
mbooErrorsOccured = Not b
Exit Function
errh:
RSRefresh = False
mbooErrorsOccured = True
End Function

Private Function RegReadMRU() As Boolean
Dim l&
Dim cnt&
Dim vIn As Variant
Dim arrIn() As String
Dim arrSplit() As String
Dim sVal$, s$
On Error GoTo errh
l = mcReg.EnumValues(eBaseKey.HKEY_LOCAL_MACHINE, Me.RegKey, arrIn)
If l <> 0 Or UBound(arrIn) = -1 Then GoTo errh
For cnt = 0 To UBound(arrIn)
    l = mcReg.QueryValue(eBaseKey.HKEY_LOCAL_MACHINE, Me.RegKey, arrIn(cnt), vIn)
    If l = 0 Then
        If arrIn(cnt) <> "" Then
            sVal = CStr(vIn)
            mrs.AddNew
            s = CStr(sVal)
            mrs("Document").Value = s
            mrs("Ind").Value = CLng(Val(arrIn(cnt)))
            s = mcStrL.GetFileOf(s)
            mrs("name").Value = mstrLeftPadName & s
            mrs.Update
            mrs.Sort = "IND ASC"
        End If
    End If
Next cnt
RegReadMRU = True
Exit Function
errh:
RegReadMRU = False
End Function

Public Function RegClearMRU() As Boolean
Dim l&
Dim cnt&
Dim vIn As Variant
Dim arrIn() As String
On Error GoTo errh
l = mcReg.EnumValues(eBaseKey.HKEY_LOCAL_MACHINE, Me.RegKey, arrIn)
If l <> 0 Or UBound(arrIn) = -1 Then GoTo errh
For cnt = 0 To UBound(arrIn)
    If arrIn(cnt) <> "" Then
        l = mcReg.DeleteValue(eBaseKey.HKEY_LOCAL_MACHINE, Me.RegKey, arrIn(cnt))
    End If
Next cnt
RegClearMRU = True
Exit Function
errh:
RegClearMRU = False
End Function

Public Function RSClear() As Boolean
Dim b As Boolean
On Error GoTo errh
b = RSInit
RSClear = True
Exit Function
errh:
RSClear = False
End Function

Public Function RegWriteMRU() As Boolean
Dim b As Boolean
Dim l&, cnt&
Dim keyval$, sVal$
On Error GoTo errh
b = RegClearMRU
If Not b Then GoTo errh
If Not RP(mrs) Then
    RegWriteMRU = True
    Exit Function
End If
mrs.MoveFirst
Do
    keyval = CStr(mrs("Ind").Value)
    sVal = mrs("Document").Value
    l = mcReg.SetKeyValue(eBaseKey.HKEY_LOCAL_MACHINE, Me.RegKey, keyval, sVal, eValueType.REG_SZ)
    mrs.MoveNext
Loop Until mrs.EOF
RegWriteMRU = True
Exit Function
errh:
RegWriteMRU = False
End Function

Private Function Exist_(obj As Object) As Boolean
On Error GoTo errh
Exist_ = IIf(obj Is Nothing, False, True)
Exit Function
errh:
Exist_ = False
Exit Function
End Function

Private Function RP(rs As Object) As Boolean
On Error GoTo errh
If rs.EOF And rs.BOF Then
    RP = False
Else
    RP = True
End If
Exit Function
errh:
 On Error Resume Next
 RP = False
 Err.Clear
End Function


Private Function ReadField(ByRef vIn As Variant, ByVal vRetVal As Variant) As Variant
On Error GoTo errh
If IsNull(vIn) Then
    ReadField = vRetVal
Else
    ReadField = vIn
End If
Exit Function
errh:
ReadField = vIn
Exit Function
End Function


Private Sub D(obj As Object)
On Error Resume Next
If Not obj Is Nothing Then
    If TypeOf obj Is VB.Form Then Unload obj
    If TypeOf obj Is ADODB.Recordset Or TypeOf obj Is ADODB.Connection Then obj.Close
    Set obj = Nothing
End If
Err.Clear
End Sub

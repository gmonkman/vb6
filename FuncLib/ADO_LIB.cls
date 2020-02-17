VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ADO_LIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event ConnectedSQL(DBName As String)
Public Event ConnectedJet(DBName As String)
Public Event ConnectedText(DBName As String)

Private mBookMark As Variant

Public Enum eiBrowse
    NextPage = 0
    PreviousPage = 1
    SamePage = 2
End Enum

Public Enum eADOFind
    SearchForward = 0
    SearchBackward = 1
    SearchFromStart = 2
    SearchFromEnd = 3
End Enum

Public Enum eADOTimeOuts
    Short = 5
    Medium = 45
    Long_ = 180
End Enum

Private Const BLOCK_SIZE = 32768

Public Enum eSQLObjectType
    sotCheckConstraint
    sotDefaultConstraint
    sotForeignKeyConstraint
    sotLog
    sotScalarFunction
    sotInlinedTableFunction
    sotStoredProcedure
    sotPrimaryKeyConstraint
    sotReplicationFilterStoredProcedure
    sotSystemTable
    sotTableFunction
    sotTrigger
    sotUserTable
    sotUniqueConstraint
    sotView
    sotExtendedStoredProcedure
End Enum


Private Function SQLObjectTypeGetTextFromEnum(ByVal eSOT As eSQLObjectType) As String
Dim s$
Select Case True
    Case eSOT = sotCheckConstraint
        s = "C"
    Case eSOT = sotDefaultConstraint
        s = "D"
    Case eSOT = sotExtendedStoredProcedure
        s = "X"
    Case eSOT = sotForeignKeyConstraint
        s = "F"
    Case eSOT = sotInlinedTableFunction
        s = "IF"
    Case eSOT = sotLog
        s = "L"
    Case eSOT = sotPrimaryKeyConstraint
        s = "PK"
    Case eSOT = sotReplicationFilterStoredProcedure
        s = "RF"
    Case eSOT = sotScalarFunction
        s = "FN"
    Case eSOT = sotStoredProcedure
        s = "P"
    Case eSOT = sotSystemTable
        s = "S"
    Case eSOT = sotTableFunction
        s = "TF"
    Case eSOT = sotTrigger
        s = "TR"
    Case eSOT = sotUniqueConstraint
        s = "UQ"
    Case eSOT = sotUserTable
        s = "U"
    Case eSOT = sotView
        s = "V"
    Case Else
        s = ""
End Select
SQLObjectTypeGetTextFromEnum = s
End Function

Public Function CreateTempTable(ByRef Cnn As ADODB.Connection, table_name As String, Parent_Table_Name As String, _
Optional Deleted_Field As Boolean = True, Optional Populate_Where_SQL As _
String = "") As Boolean

Dim sql As String
Dim rs As New ADODB.Recordset
Dim SelectSQL As String

If Left(table_name, 1) <> "#" Then
    table_name = "#" & table_name
End If

sql = "SELECT"
sql = sql & _
    " COLUMN_NAME , COLUMN_DEFAULT, IS_NULLABLE, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH" & _
    " FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & Parent_Table_Name & _
    "'"
rs.Open sql, Cnn, adOpenStatic, adLockReadOnly
If Not rs.EOF Then
    sql = ""
    sql = "USE TEMPDB IF EXISTS (SELECT * FROM tempdb..sysobjects where id =object_id(N'tempdb.." & table_name & "') and OBJECTPROPERTY(id,n 'IsUserTable') = 1) "
    sql = sql & "DROP TABLE " & table_name
    sql = sql & " USE " & Cnn.DatabaseName & " CREATE TABLE " & table_name & " ("
    While Not rs.EOF
        sql = sql & " " & rs("COLUMN_NAME") & " " & rs("DATA_TYPE") & " "

        If ReadField(rs("CHARACTER_MAXIMUM_LENGTH"), 0) <> 0 And ReadField(rs("CHARACTER_MAXIMUM_LENGTH"), 0) < 8000 Then
            sql = sql & "(" & rs("CHARACTER_MAXIMUM_LENGTH") & ") "
        End If

        If UCase(rs("IS_NULLABLE")) = "YES" Then
            sql = sql & "NULL "
        Else
            sql = sql & "NOT NULL "
        End If

        If Not IsNull(rs("COLUMN_DEFAULT")) Then
            sql = sql & "DEFAULT " & rs("COLUMN_DEFAULT") & " "
        End If

        SelectSQL = SelectSQL & rs("COLUMN_NAME")

        rs.MoveNext

        If Not rs.EOF Then
            sql = sql & ","
            SelectSQL = SelectSQL & ","
        End If
    Wend

    If Deleted_Field = True Then
        sql = sql & ", DELETED BIT NOT NULL DEFAULT 0"
        SelectSQL = SelectSQL & ",0"
    End If

    sql = sql & ")"

    Cnn.Execute sql
    If Populate_Where_SQL <> "" Then
        sql = "INSERT INTO " & table_name & " " & "SELECT " & SelectSQL & " FROM " & Parent_Table_Name & " " & Populate_Where_SQL
        Cnn.Execute sql
    End If
    CreateTempTable = True
Else
    CreateTempTable = False
End If
End Function

Public Property Get Bookmark() As Variant
If VarType(mBookMark) = vbObject Then
    Set Bookmark = mBookMark
Else
    Bookmark = mBookMark
End If
End Property

Public Sub BookmarkSave(ByRef rs As ADODB.Recordset)
If Not rs Is Nothing Then
    If RP(rs) Then
        mBookMark = rs.Bookmark
        GoTo ENDBS
    End If
End If
BookMarkClear
ENDBS:
End Sub

Public Sub BookmarkSet(ByRef rs As ADODB.Recordset)
On Error Resume Next
If VarType(mBookMark) = vbObject Then If mBookMark Is Nothing Then GoTo ENDBSET
If Not rs Is Nothing Then
    If RP(rs) Then
        rs.Bookmark = mBookMark
        GoTo ENDBSET
    End If
End If
BookMarkClear
ENDBSET:
End Sub

Public Function AccessTablesEnum(AccessDBCnn As ADODB.Connection) As String()
Dim a() As String
Dim sql$
Dim cnt&
Dim rs As New ADODB.Recordset
sql = ""
sql = "SELECT name FROM MSysObjects WHERE Type=" & CStr(edaoSysObjTypes.sotStandard) & _
    " AND Flag=" & CStr(edaoSysObjFlags.sotUserTable)
rs.Open sql, AccessDBCnn, adOpenForwardOnly, adLockReadOnly, adCmdText
cnt = 0
Erase a
If RP(rs) Then
    ReDim Preserve a(cnt)
    a(cnt) = rs("name").Value
    cnt = cnt + 1
End If
AccessTablesEnum = a()
End Function

Public Sub BookMarkClear()
On Error Resume Next
Set mBookMark = Nothing
Err.Clear
End Sub

Public Function FieldLength(ByRef Cnn As ADODB.Connection, ByVal TableName$, ByVal FieldName$, Optional ByVal lOutErr As Long = vbObjectError) As Long
Dim rs As New ADODB.Recordset
On Error GoTo errh
rs.Open "select " & FieldName & " from " & TableName & " where (1=2);", Cnn
FieldLength = rs.Fields(0).DefinedSize
GoSub LC
Exit Function

errh:
 On Error Resume Next
 GoSub LC
 FieldLength = lOutErr
Exit Function

LC:
 On Error Resume Next
 D rs
Return
End Function

Public Function JetGetOLEDBConnectionString(ByVal Database$, Optional ByVal UID$ = "admin", Optional ByVal PWD$ = "") As String
Dim prov$
Dim lErr&
On Error GoTo errh
prov = Me.JetProviderGetLatest(lErr, True)
If lErr <> 0 Then GoTo errh
JetGetOLEDBConnectionString = "Provider=" & prov & ";Data Source=" & Database & ";User ID=" & UID & ";Password=" & PWD
Exit Function
errh:
JetGetOLEDBConnectionString = ""
End Function

Public Function TextGetODBCConnectionString(ByVal DirName As String) As String
Dim s$
s = ""
s = s & _
    "DRIVER={Microsoft Text Driver (*.txt; *.csv)}" & _
";DBQ=" & DirName & _
";DefaultDir=" & DirName & _
";Uid=Admin;Pwd=;"
TextGetODBCConnectionString = s
End Function

Public Function JetProviderGetLatest(ByRef lOutErr As Long, Optional ByVal ReturnFull As Boolean = False) As String
Dim cReg As New Registry32
Dim sKeyName As String, s$
Dim b As Boolean
Dim vIn As Variant
Dim l&

Const JET350 As String = "Microsoft.Jet.OLEDB.3.50"
Const JET351 As String = "Microsoft.Jet.OLEDB.3.51"
Const JET40 As String = "Microsoft.Jet.OLEDB.4.0"
Const TEXTLENGTH = 20
On Error GoTo errh

l = cReg.QueryValue(Reg32.HKEY_CLASSES_ROOT, JET40, "", vIn)
If l <> 0 Then
    l = cReg.QueryValue(Reg32.HKEY_CLASSES_ROOT, JET351, "", vIn)
    If l <> 0 Then
        l = cReg.QueryValue(Reg32.HKEY_CLASSES_ROOT, JET350, "", vIn)
    End If
End If
s = CStr(vIn)
If Not ReturnFull Then
    s = Trim$(Right$(s, Len(s) - TEXTLENGTH))
Else: s = Trim$(s)
End If
If s = "" Then lOutErr = vbObjectError
JetProviderGetLatest = s
GoSub LC
Exit Function

errh:
JetProviderGetLatest = ""
lOutErr = Err.Number
GoSub LC
Exit Function

LC:
D cReg
Return
End Function


Public Function GetRefFromFieldString(ByVal strRefToGet$, ByVal strFieldWithKnownValue$, ByVal strValue$, ByVal strTableName$, ByRef varReturn As Variant, ByRef Cnn As ADODB.Connection) As Boolean
Dim strSQL$
Dim rsTarget As New ADODB.Recordset
On Error GoTo errh

strSQL = "Select " & strRefToGet & ", " & strFieldWithKnownValue
strSQL = strSQL & " FROM " & strTableName
strSQL = strSQL & " WHERE " & strFieldWithKnownValue & "='" & strValue & "';"

rsTarget.Open strSQL, Cnn, adOpenForwardOnly, adLockOptimistic, adCmdText


If Not RP(rsTarget) Then
    GetRefFromFieldString = False: varReturn = 0
    GoSub LC
    Exit Function
End If

varReturn = rsTarget(strRefToGet)

GoSub LC
GetRefFromFieldString = True
Exit Function

errh:
 On Error Resume Next
 GoSub LC
 GetRefFromFieldString = False
Exit Function

LC:
On Error Resume Next
D rsTarget
Return

End Function

Public Function GetRefFromKnownField(ByVal strRefToGet$, ByVal strFieldWithKnownValue$, vValue As Variant, ByVal strTableName$, ByRef varReturn As Variant, ByRef Cnn As ADODB.Connection, Optional IsJet As Boolean = False) As Boolean
Dim strSQL$
Dim rsTarget As New ADODB.Recordset
Dim sDate$
Dim dblLower#, dblUpper#
On Error GoTo errh

If VarType(vValue) = vbString Then
    strSQL = "Select DISTINCT " & strRefToGet & ", " & strFieldWithKnownValue & _
        " FROM " & strTableName & _
        " WHERE " & strFieldWithKnownValue & "='" & CStr(vValue) & "';"
ElseIf VarType(vValue) = vbLong Or VarType(vValue) = vbInteger Then
    strSQL = "Select DISTINCT " & strRefToGet & ", " & strFieldWithKnownValue & _
        " FROM " & strTableName & _
        " WHERE " & strFieldWithKnownValue & "=" & CLng(vValue) & ";"
ElseIf VarType(vValue) = vbDate Then
    If IsJet Then
        sDate = "#" & Format$(vValue, "dd mmm yyyy") & "#"
    Else
        sDate = "'" & Format$(vValue, "dd mmm yyyy") & "'"
    End If
    strSQL = "Select DISTINCT " & strRefToGet & ", " & strFieldWithKnownValue & _
        " FROM " & strTableName & _
        " WHERE " & strFieldWithKnownValue & "=" & sDate & ";"
ElseIf VarType(vValue) = vbDouble Or VarType(vValue) = vbSingle Then
    dblLower = CDbl(vValue) - 0.001
    dblUpper = CDbl(vValue) + 0.001
    strSQL = "Select DISTINCT " & strRefToGet & ", " & strFieldWithKnownValue & _
        " FROM " & strTableName & _
        " WHERE " & strFieldWithKnownValue & ">" & CStr(dblLower) & " AND " & _
        strFieldWithKnownValue & "<" & CStr(dblUpper) & ";"
ElseIf VarType(vValue) = vbCurrency Then
    strSQL = "Select DISTINCT " & strRefToGet & ", " & strFieldWithKnownValue & _
        " FROM " & strTableName & _
        " WHERE " & strFieldWithKnownValue & "=" & CCur(vValue) & ";"
Else
    GetRefFromKnownField = False
    GoSub LC
    Exit Function
End If

rsTarget.Open strSQL, Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText

If Not RP(rsTarget) Then
    GetRefFromKnownField = False: varReturn = 0
    GoSub LC
    Exit Function
End If

varReturn = rsTarget(strRefToGet)

GoSub LC
GetRefFromKnownField = True
Exit Function

errh:
 On Error Resume Next
 GoSub LC
 GetRefFromKnownField = False
Exit Function

LC:
On Error Resume Next
D rsTarget
Return

End Function


Public Function GetRefFromFieldWithFK(ByVal strRefToGet$, ByVal strFieldWithKnownValue$, ByVal strValue$, ByVal strFKFieldName$, ByVal lFKValue As Long, ByVal strTableName$, ByRef varReturn As Variant, ByRef Cnn As ADODB.Connection) As Boolean
Dim strSQL$
Dim rsTarget As New ADODB.Recordset
On Error GoTo errh

strValue = VBA.Replace(strValue, "'", "''", , , vbTextCompare)
strSQL = "Select DISTINCT " & strRefToGet & ", " & strFieldWithKnownValue & _
 " FROM " & strTableName & _
 " WHERE " & strFieldWithKnownValue & "='" & strValue & "'" & _
 " AND " & strFKFieldName & "=" & lFKValue & ";"
rsTarget.Open strSQL, Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText

If Not RP(rsTarget) Then
    GetRefFromFieldWithFK = False: varReturn = 0
    GoSub LC
    Exit Function
End If

varReturn = rsTarget(strRefToGet)

GoSub LC
GetRefFromFieldWithFK = True
Exit Function

errh:
Resume e
e:
GoSub LC
GetRefFromFieldWithFK = False
Exit Function

LC:
On Error Resume Next
D rsTarget
Return

End Function

Private Function Exist_(ByRef obj As Object) As Boolean
On Error GoTo errh
Exist_ = IIf(obj Is Nothing, False, True)
Exit Function
errh:
Exist_ = False
Exit Function
End Function

Public Function SQL_nVarCharColumnResize(ByRef Cnn As ADODB.Connection, ByVal TableName As String, ByVal ColName As String, ByVal NewSize As Long) As Boolean
Dim sql$
On Error GoTo errh
sql = "ALTER TABLE " & TableName & _
  " ALTER COLUMN " & ColName & _
    " nvarchar(" & NewSize & ")"
Cnn.Execute sql
SQL_nVarCharColumnResize = True
Exit Function
errh:
SQL_nVarCharColumnResize = False
End Function

Public Function SQL_IndexAdd(ByRef cn As ADODB.Connection, ByVal TableName$, ByVal IndexName$, ByVal IsUnique As Boolean, ByVal IsClustured As Boolean, ByVal NoCheck As Boolean, ByVal IsConstraint As Boolean, ByVal IgnoreDupKey As Boolean, ByRef lOurErr&, ParamArray Cols()) As Boolean
Dim sql$
Dim s$, sNC$, sCols$, sDup$, sWith$
Dim cnt&

On Error GoTo errh
If UBound(Cols()) = -1 Then Err.Raise vbObjectError
If Me.SQL_IndexExists(cn, IndexName, cnt) Then Exit Function
For cnt = 0 To UBound(Cols)
    sCols = Trim$(sCols & " [" & Cols(cnt) & "], ")
Next cnt
sCols = IIf(Right$(sCols, 1) = ",", Left$(sCols, Len(sCols) - 1), sCols)

If IsConstraint Then
    s = IIf(IsUnique, " UNIQUE ", "")
    s = s & IIf(IsClustured, " CLUSTERED ", " NONCLUSTERED ")
    sNC = IIf(NoCheck, " WITH NOCHECK ", "")
    sql = "ALTER TABLE [dbo].[" & TableName & "] " & sNC & " ADD " & vbCrLf & _
        "CONSTRAINT [" & IndexName & "] " & s & " ( " & sCols & " ) ON [PRIMARY]"
Else
    s = IIf(IsUnique, " UNIQUE ", "")
    s = s & IIf(IsClustured, " CLUSTERED ", " NONCLUSTERED ")
    sDup = IIf(IgnoreDupKey, " IGNORE_DUP_KEY ", "")
    If Len(sDup) > 0 Then sWith = " WITH " & sDup
     sql = "CREATE  " & s & " INDEX [" & IndexName & "] ON [dbo].[" & TableName & "]" & _
        "( " & sCols & " ) " & sWith & " ON [PRIMARY]"

End If

sql = Replace(sql, "  ", " ")
sql = Replace(sql, "  ", " ")
sql = Replace(sql, "  ", " ")
    
cn.Execute sql
SQL_IndexAdd = True
Exit Function
errh:
 lOurErr = Err.Number
 On Error Resume Next
 SQL_IndexAdd = False
End Function

Public Function SQL_ConstraintRelationAdd(ByRef cn As ADODB.Connection, ByVal KeyTable As String, ByVal FKTable As String, ByVal KeyTableColName$, ByVal FKTableColName As String, ByVal relname As String, ByRef lOurErr&) As Boolean
Dim sql$
Dim l&
On Error GoTo errh
If Me.SQL_ForeignKeyExists(cn, relname, FKTable, l) Then
    SQL_ConstraintRelationAdd = True
    Exit Function
End If
sql = "ALTER TABLE [dbo].[" & FKTable & "] ADD" & vbCrLf & _
    "CONSTRAINT [" & relname & "] FOREIGN KEY" & vbCrLf & _
    "([" & FKTableColName & "]) REFERENCES [dbo].[" & KeyTable & "] ([" & KeyTableColName & "])"
cn.Execute sql
SQL_ConstraintRelationAdd = True
Exit Function
errh:
 SQL_ConstraintRelationAdd = False
End Function

Public Function SQL_ConstraintDefaultAdd(ByRef cn As ADODB.Connection, ByVal TableName$, ByVal ConstraintName$, ByVal ColName$, ByVal Value As Variant, Optional NoCheck As Boolean = True) As Boolean
Dim sql$
Dim s$
On Error GoTo errh
s = Me.SQL_ColumnGetConstraintName(cn, TableName, ColName)
If s = ConstraintName Then SQL_ConstraintDefaultAdd = True: Exit Function
If s <> "" Then
    Me.SQL_ConstraintDrop cn, TableName, s
End If
If NoCheck Then
    sql = "ALTER TABLE [dbo].[" & TableName & "] WITH NOCHECK ADD"
Else
    sql = "ALTER TABLE [dbo].[" & TableName & "]"
End If

If VarType(Value) = vbString Or VarType(Value) = vbDate Then
    If VarType(Value) = vbDate Then
        s = Format$(Value, "yyyy mmm dd")
    Else
        s = Value
    End If
    sql = sql & vbCrLf & " CONSTRAINT [" & ConstraintName & "] DEFAULT ('" & s & "') FOR [" & ColName & "]"
ElseIf VarType(Value) = vbLong Or VarType(Value) = vbByte Or VarType(Value) = vbInteger Or VarType(Value) Then
    sql = sql & vbCrLf & " CONSTRAINT [" & ConstraintName & "] DEFAULT (" & CStr(Value) & ") FOR [" & ColName & "]"
Else
    SQL_ConstraintDefaultAdd = False
    Exit Function
End If
cn.Execute sql
SQL_ConstraintDefaultAdd = True
Exit Function
errh:
SQL_ConstraintDefaultAdd = False
End Function

Public Function SQL_IndexExists(ByRef cn As ADODB.Connection, ByVal IndName$, ByRef lOutErr&) As Boolean
Dim sql$
Dim l&
Dim rs As New ADODB.Recordset
On Error GoTo errh
sql = "select name, indid from dbo.sysindexes where name = '" & IndName & "'"
rs.Open sql, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
If RP(rs) Then
    SQL_IndexExists = CBool(rs("indid").Value)
Else
    SQL_IndexExists = False
End If
GoSub LC
Exit Function

errh:
 lOutErr = Err.Number
 On Error Resume Next
 GoSub LC
 SQL_IndexExists = False
Exit Function

LC:
 On Error Resume Next
 D rs
Return
End Function

Public Function SQL_TableDrop(ByRef cn As ADODB.Connection, ByVal TableName As String) As Boolean
Dim sql$
On Error GoTo errh
sql = "IF EXISTS (select * from dbo.sysobjects where id = object_id(N'[dbo].[" & TableName & _
    "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
    " DROP Table " & TableName
cn.Execute sql
SQL_TableDrop = True
Exit Function

errh:
 SQL_TableDrop = False
End Function

Public Function SQL_ForeignKeyExists(ByRef cn As ADODB.Connection, ByVal FK_RelationName As String, ByVal TableName As String, ByRef lOutErr As Long) As Boolean
Dim sql$
Dim rs As New ADODB.Recordset
On Error GoTo errh
sql = "IF EXISTS (select * from dbo.sysobjects where id = object_id(N'[dbo].[" & FK_RelationName & "]')" & _
    " AND OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & _
    " SELECT 1 as Ret Else SELECT 0 as Ret"
Set rs = cn.Execute(sql)
If RP(rs) Then
    If rs("ret").Value = 1 Then
        SQL_ForeignKeyExists = True
    Else
        SQL_ForeignKeyExists = False
    End If
Else
    lOutErr = vbObjectError
    SQL_ForeignKeyExists = False
End If
GoSub LC
Exit Function

errh:
 lOutErr = Err.Number
 On Error Resume Next
 GoSub LC
 SQL_ForeignKeyExists = False
Exit Function

LC:
 On Error Resume Next
 D rs
Return
End Function

Public Function SQL_ForeignKeyDrop(ByRef cn As ADODB.Connection, ByVal FK_RelationName As String, ByVal TableName As String) As Boolean
Dim sql$
On Error GoTo errh
sql = "IF EXISTS (select * from dbo.sysobjects where id = object_id(N'[dbo].[" & FK_RelationName & "]')" & _
    " AND OBJECTPROPERTY(id, N'IsForeignKey') = 1)" & _
    " ALTER TABLE [dbo].[" & TableName & "] DROP CONSTRAINT " & FK_RelationName
cn.Execute sql
SQL_ForeignKeyDrop = True
Exit Function
errh:
SQL_ForeignKeyDrop = False
End Function

Public Function SQL_IndexDrop(ByRef cn As ADODB.Connection, ByVal IndexName As String) As Boolean
Dim sql$
On Error GoTo errh
sql = "IF EXISTS (SELECT name FROM sysindexes WHERE name = '" & IndexName & "')" & _
        " DROP INDEX " & IndexName
cn.Execute sql
SQL_IndexDrop = True
Exit Function
errh:
 SQL_IndexDrop = False
End Function

Public Function SQL_TableExists(ByRef cn As ADODB.Connection, ByVal TableName As String, ByRef lOutErr&) As Boolean
Dim rs As New ADODB.Recordset
Dim sql$
On Error GoTo errh
sql = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[" & TableName & "]')" & _
    " and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & _
    " Select 1 as Ret else Select 0 as Ret"
Set rs = cn.Execute(sql)
If RP(rs) Then
    If rs("Ret").Value = 1 Then
        SQL_TableExists = True
    Else
        SQL_TableExists = False
    End If
Else
    SQL_TableExists = False
End If
GoSub LC
Exit Function

errh:
 lOutErr = Err.Number
 On Error Resume Next
 GoSub LC
 SQL_TableExists = False
Exit Function

LC:
 On Error Resume Next
 D rs
Return
End Function

Public Sub ArrayFillFromField(ByRef Cnn As ADODB.Connection, ByVal sql As String, _
    FieldName As String, ByRef vOut() As Variant)
Dim cnt&
Dim rs As New ADODB.Recordset
rs.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly
If Not RP(rs) Then Exit Sub
Erase vOut
For cnt = 0 To rs.RecordCount - 1
    ReDim Preserve vOut(cnt)
    vOut(cnt) = rs(FieldName).Value
Next cnt
End Sub

Public Function SQL_SysObjectExists(ByRef Cnn As ADODB.Connection, ByVal ObjectName$, ByVal eSOT As eSQLObjectType) As Boolean
Dim sql$
Dim rs As New ADODB.Recordset
sql = "select 1 as one from sysobjects where name='" & FQ(ObjectName) & "' AND xtype='" & SQLObjectTypeGetTextFromEnum(eSOT) & "'"
rs.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
SQL_SysObjectExists = RP(rs)
End Function

Private Function FQ(ByVal s As String) As String
FQ = Replace(s, "'", "''")
End Function

Public Function SQL_FieldExists(ByRef Cnn As ADODB.Connection, ByVal TableName$, ByVal FieldName$, ByRef lOutErr As Long) As Boolean
Dim sql$
Dim rs As New ADODB.Recordset
On Error GoTo errh
sql = "if exists (select * from sysobjects where id = object_id(N'[dbo].[" & TableName & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1 and COLUMNPROPERTY(id, N'" & FieldName & "', N'IsIdentity') is not null)" & _
        " SELECT -1 as Ret Else SELECT 0 as Ret"
rs.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
SQL_FieldExists = rs("ret").Value
GoSub LC
Exit Function

LC:
 On Error Resume Next
 D rs
 Err.Clear
Return

errh:
 lOutErr = Err.Number
 SQL_FieldExists = False
 Err.Clear
End Function

Public Function SQL_FieldAdd(ByRef Cnn As ADODB.Connection, ByVal TableName$, ByVal fld$, ByVal attrs$, Optional ByVal TimeOut& = 10) As Boolean
Dim sql$
Dim cm As New ADODB.Command
On Error GoTo errh
sql = sql + " if exists (select * from sysobjects where id = object_id(N'[dbo].[" & TableName & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1 and COLUMNPROPERTY(id, N'" & fld & "', N'IsIdentity') is null)"
sql = sql + " alter table " & TableName
sql = sql + " add " & fld & " " & attrs  'DISABLED bit not null default 0"
cm.CommandText = sql
cm.ActiveConnection = Cnn
cm.CommandTimeout = TimeOut
cm.CommandType = adCmdText
cm.Execute
SQL_FieldAdd = True
Exit Function
errh:
SQL_FieldAdd = False
Err.Clear
End Function

Public Function ADORecordsPresent(ByRef adoRS As ADODB.Recordset) As Boolean
If adoRS.EOF And adoRS.BOF Then
    ADORecordsPresent = False
Else
    ADORecordsPresent = True
End If
End Function

Public Function CmdGetDefaultSP(ByRef Cnn As ADODB.Connection, ByRef adoCmd As ADODB.Command, ByVal sprocName As String, Optional ByVal TimeOut As Long = 30) As Boolean
On Error GoTo errh
If Exist_(adoCmd) Then Set adoCmd = Nothing
Set adoCmd = New ADODB.Command
With adoCmd
   ' .NamedParameters = True
    .CommandText = sprocName
    .CommandType = adCmdStoredProc
    .ActiveConnection = Cnn
    .Parameters.Refresh
End With
CmdGetDefaultSP = True
Exit Function

errh:
Resume e
e:
On Error Resume Next
CmdGetDefaultSP = False
Set adoCmd = Nothing
Exit Function
End Function

Public Function GetRecordCount(ByVal str_aColName As String, ByVal strWhere As String, ByVal strTable As String, adoConn As ADODB.Connection) As Long
Dim sql$
Dim adoRS As ADODB.Recordset
On Error GoTo errh
Set adoRS = New ADODB.Recordset
sql = "SELECT Count(" & str_aColName & ") as cnt FROM " & strTable & " " & strWhere
adoRS.Open sql, adoConn, adOpenForwardOnly, adLockReadOnly, adCmdText
GetRecordCount = CLng(ReadField(adoRS.Fields("cnt").Value, -1))
GoSub LocalClean
Exit Function
errh:
GoSub LocalClean
GetRecordCount = -1
Exit Function

LocalClean:
On Error Resume Next
DestroyObj adoRS
Return

End Function

Private Sub DestroyObj(obj As Object)
On Error Resume Next
obj.Close
Set obj = Nothing
Err.Clear
End Sub

Private Function ReadField(ByVal Arg As Variant, DefVal As Variant) As Variant
On Error Resume Next
If IsNull(Arg) Then
    ReadField = DefVal
Else
    ReadField = Arg
End If
Exit Function
errh:
Err.Clear
ReadField = Null
Exit Function
End Function

Public Function GetOLEDBADOConn(ByVal ServerName As String, ByVal DatabaseName As String, Optional ByVal UID As String = "sa", Optional ByVal PWD As String = "", Optional ByVal CommandTimeout As Long = 30, Optional ByVal CursorLocation As ADODB.CursorLocationEnum = adUseClient, Optional ByVal bIntegrated As Boolean = True) As ADODB.Connection
Dim adocnn As ADODB.Connection
Dim str$
Set adocnn = New ADODB.Connection

adocnn.Provider = "SQLOLEDB"
adocnn.CommandTimeout = IIf(CommandTimeout = 0, 30, CommandTimeout)
adocnn.CursorLocation = IIf(CursorLocation = 0, adUseClient, CursorLocation)
adocnn.Properties("Data Source").Value = ServerName
adocnn.Properties("Initial Catalog").Value = DatabaseName

If bIntegrated = True Then
    adocnn.Properties("Integrated Security").Value = "SSPI"
Else
    adocnn.Properties("User ID").Value = UID
    adocnn.Properties("Password").Value = PWD
End If
adocnn.Open
If Len(DatabaseName) > 0 Then adocnn.DefaultDatabase = DatabaseName
Set GetOLEDBADOConn = adocnn
RaiseEvent ConnectedSQL(DatabaseName)
End Function

Public Function GetOLEDBADOConn2(ByVal ServerName As String, ByVal DatabaseName As String, Optional ByVal UID As String = "sa", Optional ByVal PWD As String = "", Optional ByVal CommandTimeout As Long = 0, Optional ByVal CursorLocation As ADODB.CursorLocationEnum = adUseClient, Optional ByVal bIntegrated As Boolean = True, Optional ByVal ConnectionTimeout As Long = 60) As ADODB.Connection
Dim adocnn As ADODB.Connection
Dim str$
Set adocnn = New ADODB.Connection

adocnn.Provider = "SQLOLEDB"
adocnn.CommandTimeout = CommandTimeout
adocnn.ConnectionTimeout = CommandTimeout
adocnn.CursorLocation = IIf(CursorLocation = 0, adUseClient, CursorLocation)
adocnn.Properties("Data Source").Value = ServerName
adocnn.Properties("Initial Catalog").Value = DatabaseName

If bIntegrated = True Then
    adocnn.Properties("Integrated Security").Value = "SSPI"
Else
    adocnn.Properties("User ID").Value = UID
    adocnn.Properties("Password").Value = PWD
End If
adocnn.Open
If Len(DatabaseName) > 0 Then adocnn.DefaultDatabase = DatabaseName
Set GetOLEDBADOConn2 = adocnn
RaiseEvent ConnectedSQL(DatabaseName)
End Function

Public Function GetODBCADOTextConn(ByVal Directory As String, Optional ByVal CommandTimeout As Long = 0, Optional ByVal CursorLocation As ADODB.CursorLocationEnum = adUseClient, Optional ByVal ConnectionTimeout As Long = 60) As ADODB.Connection
Dim adocnn As ADODB.Connection
Dim str$
Set adocnn = New ADODB.Connection
str = Me.TextGetODBCConnectionString(Directory)
adocnn.Open str
Set GetODBCADOTextConn = adocnn
RaiseEvent ConnectedText(Directory)
End Function

Public Function GetJETADOConn(ByVal Database As String, Optional ByVal UID As String = "admin", Optional ByVal PWD As String = "", Optional ByVal CommandTimeout As Long = 30, Optional ByVal CursorLocation As ADODB.CursorLocationEnum = adUseClient, Optional ByVal ProviderVer = "3.51", Optional ByVal UseLatestProvider As Boolean = True) As ADODB.Connection
Dim adocnn As ADODB.Connection
Dim str$, s$
Dim lErr&
Set adocnn = New ADODB.Connection
If UseLatestProvider Then
    ProviderVer = Me.JetProviderGetLatest(lErr)
End If
If lErr <> 0 Then Err.Raise lErr
If Left$(ProviderVer, 1) <> "." Then ProviderVer = "." & ProviderVer
adocnn.Provider = "Microsoft.Jet.OLEDB" & ProviderVer
adocnn.CommandTimeout = IIf(CommandTimeout = 0, 30, CommandTimeout)
adocnn.CursorLocation = IIf(CursorLocation = 0, adUseClient, CursorLocation)
str = IIf(UID = "", "admin", UID)

adocnn.Open Database, str, PWD
Set GetJETADOConn = adocnn
RaiseEvent ConnectedJet(Database)
End Function

Public Function GetDSNADOConn(ByVal DSNName As String, Optional ByVal UID As String = "", Optional ByVal PWD As String = "", Optional ByVal adoCursorLocation As ADODB.CursorLocationEnum = adUseClient, Optional ByVal adoCommandTimeout As Long = 30, Optional ByRef lOutErr As Long = 0) As ADODB.Connection
Dim strCnn$
Dim adocnn As New ADODB.Connection
On Error GoTo errh
strCnn = "DSN=" & DSNName & IIf(UID <> "", ";UID=" & UID, "") & IIf(PWD <> "", ";PWD=" & PWD, "")
With adocnn
    .ConnectionString = strCnn
    .CursorLocation = adoCursorLocation
    .CommandTimeout = adoCommandTimeout
End With
adocnn.Open strCnn
Set GetDSNADOConn = adocnn
Exit Function

errh:
lOutErr = Err.Number
Err.Clear
Exit Function
End Function

Public Function SetadoRSPage(ByRef adoRS As ADODB.Recordset, Optional ByVal lPageNo = 0, Optional ByVal lPageSize As Long = 10, Optional ByVal Browse As eiBrowse = NextPage) As Long 'return the pagenumber
On Error GoTo errh
adoRS.PageSize = lPageSize
If lPageNo > 0 Then
ElseIf Browse = eiBrowse.PreviousPage Then
    If lPageNo > 1 Then
        lPageNo = lPageNo - 1
    Else
        lPageNo = 1
    End If
Else
    If adoRS.AbsolutePage < adoRS.PageCount Then
        lPageNo = lPageNo + 1
    Else
        lPageNo = adoRS.PageCount
    End If
End If
adoRS.AbsolutePage = lPageNo
SetadoRSPage = lPageNo
Exit Function

errh:
SetadoRSPage = 0
Exit Function

End Function


Public Function WriteBLOBToDB(ByVal strSourceFile As String, ByVal strSQL As String, ByVal strFieldName As String, cnnADO As ADODB.Connection) As Long
Dim rsADO As ADODB.Recordset
Dim intNumBlocks As Integer, intSourceFile As Integer, i As Integer
Dim lngFileLength As Long, lngLeftOver As Long
Dim bytByteData() As Byte
On Error GoTo errh

Set rsADO = New ADODB.Recordset
' Open the strSourceFile file.
intSourceFile = FreeFile
Open strSourceFile For Binary Access Read As intSourceFile

' Get the length of the file.
lngFileLength = LOF(intSourceFile)
If lngFileLength = 0 Then
    WriteBLOBToDB = 0
    GoSub LocalClean
    Exit Function
End If

' Calculate the number of blocks to read and lngLeftOver bytes.
intNumBlocks = lngFileLength \ BLOCK_SIZE
lngLeftOver = lngFileLength Mod BLOCK_SIZE

rsADO.Open strSQL, cnnADO, adOpenKeyset, adLockOptimistic, adCmdText
rsADO.MoveFirst

' Read the lngLeftOver data, writing it to the table.
ReDim bytByteData(lngLeftOver)
Get intSourceFile, , bytByteData()
rsADO(strFieldName).AppendChunk bytByteData()

' Read the remaining blocks of data, writing them to the table.
ReDim bytByteData(BLOCK_SIZE)
For i = 1 To intNumBlocks
    Get intSourceFile, , bytByteData()
    rsADO(strFieldName).AppendChunk bytByteData()
Next i

rsADO.Update

GoSub LocalClean

Exit Function

errh:
Err.Raise Err.Number, Err.Source, Err.description
lngFileLength = 0
GoSub LocalClean
Exit Function

LocalClean:
On Error Resume Next
WriteBLOBToDB = lngFileLength
Close intSourceFile
rsADO.Close
Set rsADO = Nothing
Return

End Function


Public Function WriteBLOBToDisk(ByVal strSQL As String, ByVal strFieldName As String, ByVal strFileName As String, cnnADO As ADODB.Connection) As Long
Dim rsADO As ADODB.Recordset
Dim intNumBlocks As Integer, intDestFile As Integer, i As Integer
Dim lngFileLength As Long, lngLeftOver As Long
Dim bytByteData() As Byte

On Error GoTo errh

rsADO.Open strSQL, cnnADO, adOpenForwardOnly, adLockReadOnly, adCmdText

lngFileLength = rsADO(strFieldName).ActualSize

If lngFileLength < 0 Then
    Err.Raise vbObjectError + 512, "ADO_OLE#WriteBLOBToDisk", "Invalid OLE field size"
ElseIf lngFileLength = 0 Then
    WriteBLOBToDisk = 0
    GoSub LocalClean
End If

' Calculate number of blocks to write and LeftOver bytes.
intNumBlocks = lngFileLength \ BLOCK_SIZE
lngLeftOver = lngFileLength Mod BLOCK_SIZE

intDestFile = FreeFile
Open strFileName For Binary Access Write As intDestFile

' Write the lngLeftOver data to the output file.
bytByteData() = rsADO(strFieldName).GetChunk(lngLeftOver)
Put intDestFile, , bytByteData()

' Write the remaining blocks of data to the output file.
For i = 1 To intNumBlocks
    ' Reads a chunk and writes it to output file.
    bytByteData() = rsADO(strFieldName).GetChunk(BLOCK_SIZE)
    Put intDestFile, , bytByteData()
Next i

rsADO.Update

GoSub LocalClean

Exit Function

errh:
Err.Raise Err.Number, Err.Source, Err.description
lngFileLength = 0
GoSub LocalClean
Exit Function

LocalClean:
On Error Resume Next
WriteBLOBToDisk = lngFileLength
Close intDestFile
rsADO.Close
Set rsADO = Nothing
Return

End Function


Public Function ADOFind(ByVal eADOFind As eADOFind, ByRef rstTemp As ADODB.Recordset, ByVal strFind As String) As Boolean    ' Use Find method based on user input.
Dim vBkStart As Variant
On Error GoTo errh

If strFind = "" Then
    ADOFind = False
    Exit Function
End If

If Not ADORecordsPresent(rstTemp) Then
    ADOFind = False
    Exit Function
End If
If rstTemp.EOF Or rstTemp.BOF Then rstTemp.MoveFirst
vBkStart = rstTemp.Bookmark
Select Case eADOFind
    Case SearchFromStart
        rstTemp.MoveFirst
        rstTemp.Find strFind, 0, adSearchForward
    Case SearchFromEnd
        rstTemp.MoveLast
        rstTemp.Find strFind, 0, adSearchBackward
    Case SearchForward
         rstTemp.Find strFind, 0, adSearchForward
    Case SearchBackward
         rstTemp.Find strFind, 0, adSearchBackward
End Select

If Not (rstTemp.BOF Or rstTemp.EOF) Then
    ADOFind = True
Else
    ADOFind = False
    rstTemp.Bookmark = vBkStart
End If

Exit Function
errh:
On Error Resume Next
ADOFind = False
Err.Clear
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

Public Function SQL_ColumnnVarCharResize(ByRef Cnn As ADODB.Connection, ByVal TableName As String, ByVal ColName As String, ByVal NewSize As Long) As Boolean
Dim sql$
On Error GoTo errh
sql = "ALTER TABLE " & TableName & _
  " ALTER COLUMN " & ColName & _
    " nvarchar(" & NewSize & ")"
Cnn.Execute sql
SQL_ColumnnVarCharResize = True
Exit Function
errh:
SQL_ColumnnVarCharResize = False
End Function

Public Function SQL_ColumnGetConstraintName(ByRef Cnn As ADODB.Connection, ByVal TableName As String, ByVal ColumnName As String) As String
Dim sql$
Dim rs As New ADODB.Recordset
On Error GoTo errh
sql = "select name from sysobjects where id IN  (select cdefault from syscolumns sc" & _
    " inner join sysobjects so on sc.id=so.id" & _
    " where sc.name='" & ColumnName & "' AND so.name='" & TableName & "' and so.xtype='u')"
rs.Open sql, Cnn, adOpenForwardOnly, adLockReadOnly, adCmdText
If RP(rs) Then
    SQL_ColumnGetConstraintName = rs("name").Value
Else
    SQL_ColumnGetConstraintName = ""
End If
GoSub LC
Exit Function
LC:
On Error Resume Next
 D rs
Return

errh:
 Resume e
e:
 On Error Resume Next
 GoSub LC
 SQL_ColumnGetConstraintName = ""
End Function

Public Function SQL_ConstraintDrop(ByRef Cnn As ADODB.Connection, ByVal TableName As String, ByVal ConstraintName As String) As Boolean
Dim sql$
On Error GoTo errh
If ConstraintName = "" Then SQL_ConstraintDrop = True: Exit Function
sql = "IF EXISTS(SELECT * FROM sysobjects WHERE name='" & ConstraintName & "' AND xtype='D') "
sql = sql & " ALTER TABLE " & TableName & " DROP CONSTRAINT " & ConstraintName
Cnn.Execute sql
SQL_ConstraintDrop = True
Exit Function
errh:
 Resume e
e:
 SQL_ConstraintDrop = False
End Function

Public Function SQL_ColumnNumericRetype(ByRef Cnn As ADODB.Connection, ByVal TableName As String, ByVal ColName As String, ByVal NewType As String) As Boolean
Dim sql$
On Error GoTo errh
sql = "ALTER TABLE " & TableName & _
  " ALTER COLUMN " & ColName & _
    " " & NewType
Cnn.Execute sql
SQL_ColumnNumericRetype = True
Exit Function
errh:
SQL_ColumnNumericRetype = False
End Function

Public Function SQL_GetIdentity(ByRef Cnn As ADODB.Connection) As Long
Dim rs As ADODB.Recordset
Dim cdoGetIdentity As New ADODB.Command
Dim sql$
On Error GoTo errh

sql = ""
sql = sql & "Select id=@@identity"
With cdoGetIdentity
    Set .ActiveConnection = Cnn
    .CommandType = adCmdText
    .CommandText = sql
    .Prepared = True
End With

With cdoGetIdentity
    Set .ActiveConnection = Cnn
    Set rs = .Execute
    SQL_GetIdentity = rs(0).Value
End With
GoSub LC
Exit Function

LC:
 On Error Resume Next
 If rs.State <> ObjectStateEnum.adStateClosed Then rs.Close
 Set rs = Nothing
Return

errh:
 On Error Resume Next
 GoSub LC
 SQL_GetIdentity = vbObjectError
End Function

Public Function Insert_Across_Connection(src_cn As ADODB.Connection, ByVal sqlselect As String, dst_cn As ADODB.Connection, ByVal dst_table As String, ParamArray dst_fields()) As Boolean
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim arr
Dim i&, j&
Dim sql$
Dim savetimeout&
Dim savetimeout2&
On Error GoTo errh
    
savetimeout = dst_cn.CommandTimeout
dst_cn.CommandTimeout = 300
savetimeout2 = src_cn.CommandTimeout
src_cn.CommandTimeout = 300
rs.Open sqlselect, src_cn, adOpenForwardOnly, adLockReadOnly
sql = "select all * from " & dst_table & " where (1=2)"
rs2.Open sql, dst_cn, adOpenKeyset, adLockBatchOptimistic

If Not RP(rs) Then
    GoSub LC
    Insert_Across_Connection = True 'doesnt represent a failure
    Exit Function
End If

Do While rs.EOF = False
    If UBound(dst_fields) < 0 Then
        rs2.AddNew
        For i = 0 To rs.Fields.Count - 1
            If StrComp(rs.Fields(i).Name, "new_record", vbTextCompare) = 0 Then
            Else
                rs2(rs.Fields(i).Name).Value = rs.Fields(i).Value
            End If
        Next i
        rs2.Update
    ElseIf (UBound(dst_fields) = 1) And IsNull(dst_fields(0)) Then
        rs2.AddNew
        For i = 0 To rs.Fields.Count - 1
            If StrComp(rs.Fields(i).Name, "new_record", vbTextCompare) = 0 Then
            ElseIf StrComp(dst_fields(1), rs.Fields(i).Name, vbTextCompare) = 0 Then
            Else
                rs2(rs.Fields(i).Name).Value = rs.Fields(i).Value
            End If
        Next i
        rs2.Update
    ElseIf (UBound(dst_fields) = 2) And IsNull(dst_fields(0)) Then
        rs2.AddNew
        For i = 0 To rs.Fields.Count - 1
            If StrComp(dst_fields(1), rs.Fields(i).Name, vbTextCompare) = 0 Then
                rs2(dst_fields(2)).Value = rs.Fields(i).Value
            ElseIf StrComp(rs.Fields(i).Name, "new_record", vbTextCompare) = 0 Then
            Else
                rs2(rs.Fields(i).Name).Value = rs.Fields(i).Value
            End If
        Next i
        rs2.Update
    Else
        If VarType(dst_fields(0)) And vbArray Then
            arr = dst_fields(0)
        Else
            arr = dst_fields
        End If
        rs2.AddNew
        j = 0
        For i = LBound(arr) To UBound(arr)
            rs2(arr(i)).Value = rs.Fields(j).Value
            j = j + 1
        Next i
        rs2.Update
    End If
    rs.MoveNext
Loop

rs2.UpdateBatch
Insert_Across_Connection = True
GoSub LC
Exit Function

errh:
Resume e
e:
On Error Resume Next
GoSub LC
Insert_Across_Connection = False
Err.Clear
Exit Function

LC:
On Error Resume Next
rs2.CancelBatch
D rs
D rs2
dst_cn.CommandTimeout = savetimeout
src_cn.CommandTimeout = savetimeout2
Err.Clear
Return

End Function

Public Function TableIsEmpty(Cnn As ADODB.Connection, ByVal tbl As String, Optional ByRef lRetErr As Long) As Boolean
Dim rs As New ADODB.Recordset
On Error GoTo errh
rs.Open "select top 1 * from " + tbl, Cnn, adOpenForwardOnly, adLockOptimistic
If rs.RecordCount = 0 Then
    TableIsEmpty = True
Else
    TableIsEmpty = False
End If
GoSub LC
Exit Function

errh:
lRetErr = Err.Number
TableIsEmpty = False
GoSub LC
Exit Function

LC:
On Error Resume Next
D rs
Return
End Function

Private Sub D(obj As Object)
On Error Resume Next
If Not obj Is Nothing Then
    If TypeOf obj Is ADODB.Recordset Or TypeOf obj Is ADODB.Connection Then
        obj.Close
    ElseIf TypeOf obj Is VB.Form Then
        Unload obj
        DoEvents
    End If
    Set obj = Nothing
End If
Err.Clear
End Sub


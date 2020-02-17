VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Strings_LIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Public Function FixLike(ByVal s As String) As String
Dim c&
Dim ch$
Dim t$

t = ""
For c = 1 To Len(s)
    ch = Mid$(s, c, 1)
    Select Case ch
    Case "%", "[", "_"
        t = t & "[" & ch & "]"
    Case Else
        t = t & ch
    End Select
Next c
FixLike = t$
End Function

Public Function ArrayIsEmpty(arr) As Boolean
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

Public Function FileNameRemoveInvalidChars(ByVal sIn As String, Optional ByVal Substitute = "+", Optional ByVal FileNameIncludesPath As Boolean = False) As String
Dim s$
If FileNameIncludesPath Then
    s = Me.GetFileOf(sIn)
Else
    s = sIn
End If
s = VBA.Replace(s, "/", Substitute)
s = VBA.Replace(s, "\", Substitute)
s = VBA.Replace(s, ":", Substitute)
s = VBA.Replace(s, "*", Substitute)
s = VBA.Replace(s, "?", Substitute)
s = VBA.Replace(s, """", Substitute)
s = VBA.Replace(s, "<", Substitute)
s = VBA.Replace(s, ">", Substitute)
s = VBA.Replace(s, "|", Substitute)
If FileNameIncludesPath Then
    FileNameRemoveInvalidChars = Me.GetPathOf_F(sIn) & "\" & s
Else
    FileNameRemoveInvalidChars = s
End If
End Function

Public Function PadSpaceLR(ByVal Targ As String) As String
Call Me.AddLeftSpace(Targ)
Call Me.AddRight(Targ, " ")
PadSpaceLR = Targ
End Function

Public Function ASCII_IsNumeric(ByVal KeyCode As Integer) As Boolean
On Error GoTo errh
ASCII_IsNumeric = IIf(KeyCode >= 48 And KeyCode <= 57, True, False)
Exit Function
errh:
 Err.Clear
 ASCII_IsNumeric = False
End Function

Public Function NumbersGetFromMixedString(ByVal Targ As String) As Long()
Dim cnt&
Dim s$, s1$
Dim l&
Dim arr() As Long
Dim arrS() As String
On Error GoTo errh
s = Me.TransformAlphas(Targ, "$", True)
If s <> "" Then
    Do
        s1 = s
        s = VBA.Replace(s1, "$$", "$")
    Loop Until s = s1
    arrS = Split(s1, "$")
    l = 0
    For cnt = 0 To UBound(arrS)
        If IsNumeric(arrS(cnt)) Then
            ReDim Preserve arr(l)
            arr(l) = CLng(Val(arrS(cnt)))
            l = l + 1
        End If
    Next cnt
    NumbersGetFromMixedString = arr
End If
Exit Function
errh:
 On Error Resume Next
 Erase arr
 Exit Function
End Function

Public Function StringM(ByVal cnt&, ByVal CharM As String) As String
Dim n&
Dim s$
On Error GoTo errh
For n = 1 To cnt
    s = s & CharM
Next n
StringM = s
Exit Function
errh:
 On Error Resume Next
 StringM = ""
End Function

Public Function CharMaxRepeatCount(ByVal StringToSearch As String, ByVal TestChar As String, Optional ByVal lRetOnErr = -1) As Long
Dim l&, cnt&
Dim s$
On Error GoTo errh
For cnt = Len(StringToSearch) To 1 Step -1
    s = Me.StringM(cnt, TestChar)
    l = cnt
    If InStr(1, StringToSearch, s, vbTextCompare) > 0 Then Exit Function
    l = 0
Next cnt
CharMaxRepeatCount = l
Exit Function
errh:
 On Error Resume Next
 CharMaxRepeatCount = lRetOnErr
End Function

Public Function ASCII_IsPrintableNonNumeric(ByVal KeyCode As Integer) As Boolean
On Error GoTo errh
If (KeyCode >= 32 And KeyCode <= 47) Or (KeyCode >= 58 And KeyCode <= 126) Then
    ASCII_IsPrintableNonNumeric = True
Else: ASCII_IsPrintableNonNumeric = False
End If
Exit Function
errh:
 Err.Clear
 ASCII_IsPrintableNonNumeric = False
End Function

'same As VB 's instr except backwards
Public Function InStrL(ByVal InString As String, ByVal srchString As String, Optional ByVal Compare As VBA.VbCompareMethod = vbTextCompare) As Integer
  Dim iLastPos As Integer   'Set to 0 on initialization
  ' Check srchString -- a 0-length string will match every time
  If Len(srchString) Then
    ' Set iLastPos to the last matching position
    Dim iCurPos As Integer
    Do
      iLastPos = iCurPos
      iCurPos = InStr(iCurPos + 1, InString, srchString, Compare)
    Loop Until iCurPos = 0
  End If
  
  InStrL = iLastPos
End Function

Public Function StringAddSep(ByVal str$, ByVal StringToAdd$, Optional ByVal Seperator$ = ",") As String
On Error GoTo errh
str = Trim$(str)

If StringToAdd = "" Then
    If Right$(str, 1) = Seperator Then str = Left$(str, Len(str) - 1)
    StringAddSep = str
End If

If str = "" Then
    StringAddSep = StringToAdd
    Exit Function
End If

If Right$(str, 1) = Seperator Then
    str = str & " " & StringToAdd
    Exit Function
Else
    str = str & Seperator & " " & StringToAdd
End If

StringAddSep = str
Exit Function
errh:
StringAddSep = str
End Function

Public Function SingleQuoteDoubleUp(ByVal InString$) As String
On Error GoTo errh
InString = VBA.Replace(InString, "'", "''", , , vbTextCompare)
SingleQuoteDoubleUp = InString
Exit Function
errh:
SingleQuoteDoubleUp = ""
End Function

Public Function TransformAlphas(ByVal str, Optional Substitute As String = "$", Optional SpaceIsAlpha As Boolean = True) As String
Dim x&
Dim concat$, s$
If str = "" Then
    TransformAlphas = ""
    Exit Function
End If
For x = 1 To Len(str)
    s = Mid$(str, x, 1)
    If IsAlpha(s) Then
        Mid$(str, x, 1) = Substitute
    End If
Next x
TransformAlphas = str
End Function


Public Function TransformCustom(ByVal str, Substitute As String, ParamArray CharsToReplace()) As String
Dim x&, y&
Dim s$
If str = "" Or UBound(CharsToReplace) < 0 Then
    TransformCustom = ""
    Exit Function
End If
For x = 1 To Len(str)
    s = Mid$(str, x, 1)
    For y = 0 To UBound(CharsToReplace)
        If LCase$(CharsToReplace(y)) = LCase$(s) Then
            Mid$(str, x, 1) = Substitute
            Exit For
        End If
    Next y
Next x
TransformCustom = str
End Function


Public Function TransformNumerics(ByVal str, Optional Substitute As String = "#") As String
Dim x&
Dim s$
If str = "" Then
    TransformNumerics = ""
    Exit Function
End If
For x = 1 To Len(str)
    s = Mid$(str, x, 1)
    If IsNumeric(s) Then
        Mid$(str, x, 1) = Substitute
    End If
Next x
TransformNumerics = str
End Function

Public Function RemoveAlpha(ByVal str As String, SpaceIsAlpha As Boolean, ParamArray Ignore()) As String
Dim x&, y&
Dim bAdded As Boolean
Dim s$, concat$
On Error GoTo errh


For x = 1 To Len(str)
    bAdded = False
    s = Mid$(str, x, 1)
    If UBound(Ignore()) > -1 Then
        For y = 0 To UBound(Ignore())
            If VarType(Ignore(y)) = vbString Then
                If LCase$(CStr(Ignore(y))) = LCase$(s) Then
                    concat = concat + s
                    bAdded = True
                    Exit For
                End If
            End If
        Next y
    End If
    
    If Not bAdded Then
        If Not IsAlpha(s, SpaceIsAlpha) Then
            concat = concat + s
        End If
    End If
Next x

RemoveAlpha = concat
Exit Function
errh:
On Error Resume Next
RemoveAlpha = ""
Err.Clear
Exit Function
End Function

Public Function RemoveAlphaAll(ByVal str As String) As String
Dim x&
Dim iAsc%
Dim concat$, s$
On Error GoTo errh
For x = 1 To Len(str)
    s = Mid$(str, x, 1)
    iAsc = Asc(s)
    If Me.ASCII_IsNumeric(iAsc) Then
        concat = concat + s
    End If
Next x
RemoveAlphaAll = concat
Exit Function
errh:
 On Error Resume Next
 RemoveAlphaAll = ""
End Function

Public Function RemoveCustom(ByVal str As String, ParamArray Ignore()) As String
Dim x&, y&
Dim bAdded As Boolean
Dim s$, concat$
On Error GoTo errh


For x = 1 To Len(str)
    bAdded = False
    s = Mid$(str, x, 1)
    If UBound(Ignore()) > -1 Then
        For y = 0 To UBound(Ignore())
            If VarType(Ignore(y)) = vbString Then
                If LCase$(CStr(Ignore(y))) = LCase$(s) Then
                    concat = concat + s
                    bAdded = True
                    Exit For
                End If
            End If
        Next y
    End If
Next x

RemoveCustom = concat
Exit Function
errh:
On Error Resume Next
RemoveCustom = ""
Err.Clear
Exit Function
End Function


Public Function RemoveNumeric(ByVal str As String, ParamArray Ignore()) As String
Dim x&, y&
Dim bAdded As Boolean
Dim s$, concat$
On Error GoTo errh

For x = 1 To Len(str)
    bAdded = False
    s = Mid$(str, x, 1)
    If UBound(Ignore()) > -1 Then
        For y = 0 To UBound(Ignore())
            If VarType(Ignore(y)) = vbString Then
                If LCase$(CStr(Ignore(y))) = LCase$(s) Then
                    concat = concat + s
                    bAdded = True
                    Exit For
                End If
            End If
        Next y
    End If
    
    If Not bAdded Then
        If Not IsNumeric(s) Then
            concat = concat + s
        End If
    End If
Next x
RemoveNumeric = concat

Exit Function
errh:
On Error Resume Next
RemoveNumeric = ""
Err.Clear
Exit Function
End Function


'vb6 now has a function to do this
Public Function Replace(ByVal SubStringTarg As String, ByVal ReplaceWith As String, ByVal Target As String, Optional Start As Long = 1) As Variant
Dim x&, lSST_Len&, max&
On Error GoTo errh
If Start < 1 Then
    Replace = Null
    Exit Function
End If
lSST_Len = Len(SubStringTarg)
If IsNull(InStr(1, Target, SubStringTarg, vbTextCompare)) Or InStr(1, Target, SubStringTarg, vbTextCompare) < 1 Then 'dont bother if no occurences
    Replace = Null
    Exit Function
End If
x = 1
Do
    If LCase$(Mid$(Target, x, lSST_Len)) = LCase$(SubStringTarg) Then
        Target = Left$(Target, x - 1) & ReplaceWith & Right$(Target, Len(Target) - lSST_Len - x + 1)
        x = x + Len(ReplaceWith)
    Else
        x = x + 1
    End If
    max = Len(Target)
Loop Until x - 1 = max
Replace = Target
Exit Function

errh:
On Error Resume Next
Replace = Null
Err.Clear
Exit Function
End Function


Public Function IsAlpha(ByVal str As String, Optional SpaceIsAlpha As Boolean = False) As Boolean
Dim a&
If str = "" Then
    IsAlpha = False
    Exit Function
End If
    
str = Left$(str, 1)
a = Asc(str)
If (a >= 65 And a <= 90) Or (a >= 97 And a <= 122) Or (a = 32 And SpaceIsAlpha) Then
    IsAlpha = True
Else
    IsAlpha = False
End If
End Function


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


'now a function in vb6 that will do this!
Public Function Parse_(ByVal sParse As String, sSeperator As String, Optional IgnoreEmptyStrings As Boolean = True) As Variant
Dim lSep&, lCnt&, lStart&, x&, lLast&
Dim vRet As Variant
Dim ParseOut() As String

On Error GoTo errh

If sSeperator = "" Or sParse = "" Then
    ReDim ParseOut(0)
    ParseOut(0) = sParse
    Parse_ = ParseOut()
    Exit Function
End If
lStart = 1
vRet = InStr(lStart, sParse, sSeperator, vbTextCompare)
vRet = ReadField(vRet, 0)
If vRet = 0 Then
    ReDim ParseOut(0)
    ParseOut(0) = sParse
    Parse_ = ParseOut()
    Exit Function
End If

lSep = Len(sSeperator)

x = 1
Do
    vRet = InStr(lStart, sParse, sSeperator, vbTextCompare)
    vRet = ReadField(vRet, 0)
    If CLng(vRet) > 0 Then
        ReDim Preserve ParseOut(lCnt)
        If Mid$(sParse, x, vRet - x) = "" And IgnoreEmptyStrings Then 'eg ,123, would be "123" rather than "","123",""
        Else
            ParseOut(lCnt) = Mid$(sParse, x, vRet - x)
            lCnt = lCnt + 1
        End If
        lStart = CLng(vRet) + lSep
        x = lStart
        lLast = CLng(vRet)
    Else
        If Right$(sParse, Len(sParse) - (lLast + lSep - 1)) = "" And IgnoreEmptyStrings Then
        Else
            ReDim Preserve ParseOut(lCnt)
            ParseOut(lCnt) = Right$(sParse, Len(sParse) - (lLast + lSep - 1))
        End If
    End If
Loop While vRet > 0
Parse_ = ParseOut() '  8)
Exit Function
errh:
Parse_ = ""
Err.Clear
Exit Function
End Function


Public Sub PadString(ByRef strToPad$, Optional ByVal n As Integer = 1, Optional ByVal strPadChar$ = "\")
Dim i As Integer, j As Integer
j = n - Len(strToPad)
If strPadChar = "" Then strPadChar = " "
For i = 1 To j
    strToPad = strPadChar + strToPad
Next
End Sub

Public Function PadString_F(ByVal strToPad$, Optional ByVal TargetLength As Long = 1, Optional ByVal strPadChar$ = "\") As String
Dim i As Integer, j As Integer
j = TargetLength - Len(strToPad)
If strPadChar = "" Then strPadChar = " "
For i = 1 To j
    strToPad = strPadChar + strToPad
Next
PadString_F = strToPad
End Function


Public Sub AddLeftSpace(ByRef strPad$)
If strPad <> "" Then
    If Left$(strPad, 1) <> " " Then
        strPad = " " & strPad
    End If
End If
End Sub


'add the \ character to the end of a string
Public Sub AddRightSlash(ByRef str$)
If VBA.Right$(str, 1) <> "\" Then str = str + "\"
End Sub
Public Function AddRightSlash_F(ByVal str$) As String
If VBA.Right$(str, 1) <> "\" Then str = str + "\"
AddRightSlash_F = str
End Function


'remove left slash from a string
Public Sub CropLeftSlash(ByRef str$)
If VBA.Left$(str, 1) = "\" Then
    str = VBA.Right$(str, Len(str) - 1)
End If
End Sub
Public Function CropLeftSlash_F(ByVal str$) As String
If VBA.Left$(str, 1) = "\" Then
    str = VBA.Right$(str, Len(str) - 1)
End If
CropLeftSlash_F = str
End Function

Public Function CropRightSlash_F(ByVal str$) As String
If VBA.Right$(str, 1) = "\" Then
    str = VBA.Left$(str, Len(str) - 1)
End If
CropRightSlash_F = str
End Function
'remove the \ from a path string
Public Sub CropRightSlash(ByRef str$)
If VBA.Right$(str, 1) = "\" Then
    str = VBA.Left$(str, Len(str) - 1)
End If
End Sub



'add ; to end of of a string
Public Sub AddSemiColon(ByRef strTemp$)
If VBA.Right$(strTemp, 1) <> ";" Then strTemp = strTemp + ";"
End Sub

Public Sub AddRight(ByRef strTarget$, ByVal strToAdd$)
If VBA.Right$(strTarget, Len(strToAdd)) <> strToAdd Then strTarget = strTarget & strToAdd
End Sub
Public Function AddRight_F(ByVal strTarget$, ByVal strToAdd$) As String
If VBA.Right$(strTarget, Len(strToAdd)) <> strToAdd Then strTarget = strTarget & strToAdd
AddRight_F = strTarget
End Function

Public Sub AddLeft(ByRef strTarget$, ByVal strToAdd$)
If VBA.Left$(strTarget, Len(strToAdd)) <> strToAdd Then strTarget = strToAdd & strTarget
End Sub
Public Function AddLeft_F(ByVal strTarget$, ByVal strToAdd$) As String
If VBA.Left$(strTarget, Len(strToAdd)) <> strToAdd Then strTarget = strToAdd & strTarget
AddLeft_F = strTarget
End Function

'remove chr$(0) character from a string, good for some API's
Public Sub CleanString(ByRef strArg$)
Dim intArgLength%, intI%
Dim strTemp$, strSingleChar$
intArgLength = Len(strArg)
If intArgLength > 0 Then
    For intI = 1 To intArgLength
        strSingleChar = Mid$(strArg, intI, 1)
        If (strSingleChar Like "?") = True Then
            If Not Asc(strSingleChar) = 0 Then strTemp = strTemp & strSingleChar
        End If
    Next intI
    strArg = strTemp
End If
End Sub

Public Function CleanString_F(ByVal strArg$) As String
Dim intArgLength%, intI%
Dim strTemp$, strSingleChar$
intArgLength = Len(strArg)
If intArgLength > 0 Then
    For intI = 1 To intArgLength
        strSingleChar = Mid$(strArg, intI, 1)
        If (strSingleChar Like "?") = True Then
            If Not Asc(strSingleChar) = 0 Then strTemp = strTemp & strSingleChar
        End If
    Next intI
    CleanString_F = strTemp
Else
    CleanString_F = ""
End If

End Function


'add intCRNumber carriage returns to a string
Public Sub AddCR(ByRef strCR As String, Optional intCRNumber As Integer)
Dim x As Byte
If intCRNumber = 0 Then intCRNumber = 1
For x = 1 To intCRNumber
    strCR = strCR & Chr$(13)
Next
End Sub
Public Function AddCR_F(ByVal strCR As String, Optional intCRNumber As Integer) As String
Dim x As Byte
If intCRNumber = 0 Then intCRNumber = 1
For x = 1 To intCRNumber
    strCR = strCR & Chr$(13)
Next
AddCR_F = strCR
End Function

'take the first letter of a string and make it a capital
Public Sub Capitalise(ByRef strTemp$)
If strTemp <> "" Then
    strTemp = UCase$(Left$(strTemp, 1)) & Right$(strTemp, Len(strTemp) - 1)
End If
End Sub
Public Function Capitalise_F(ByVal strTemp$) As String
If strTemp <> "" Then
    strTemp = UCase$(Left$(strTemp, 1)) & Right$(strTemp, Len(strTemp) - 1)
End If
Capitalise_F = strTemp
End Function

Public Sub AddSquareBrackets(ByRef strTemp$)
If Left$(strTemp$, 1) <> "[" Then strTemp = "[" & strTemp
If Right$(strTemp, 1) <> "]" Then strTemp = strTemp & "]"
End Sub
Public Function AddSquareBrackets_F(ByVal strTemp$) As String
If Left$(strTemp$, 1) <> "[" Then strTemp = "[" & strTemp
If Right$(strTemp, 1) <> "]" Then strTemp = strTemp & "]"
AddSquareBrackets_F = strTemp
End Function


Public Sub GetPathOf(ByRef strPath$, Optional strDivider$ = "\")
Dim i%
Dim strTmp$

For i = Len(strPath) To 1 Step -1
   If Mid$(strPath, i, 1) = strDivider Then
      strTmp = Trim$(Left$(strPath, i - 1))
      Exit For
   End If
Next
If strTmp = "" Then strPath = strTmp
strPath = strTmp
End Sub
Public Function GetPathOf_F(ByVal strPath$, Optional ByVal strDivider$ = "\") As String
Dim i%
Dim strTmp$

For i = Len(strPath) To 1 Step -1
   If Mid$(strPath, i, 1) = strDivider Then
      strTmp = Trim$(Left$(strPath, i - 1))
      Exit For
   End If
Next
If strTmp = "" Then strTmp = strPath
GetPathOf_F = strTmp
End Function


Public Sub RemoveOrderBy(ByRef strSQL$)
Dim lngPos&, lngStrLength
lngPos = InStr(strSQL, " Order By ", vbTextCompare)
If lngPos <> 0 Then
    strSQL = Trim$(Left$(strSQL, lngPos - 1))
End If
End Sub
Public Function RemoveOrderBy_F(ByVal strSQL$) As String
Dim lngPos&, lngStrLength
lngPos = InStr(strSQL, " Order By ", vbTextCompare)
If lngPos <> 0 Then
    strSQL = Trim$(Left$(strSQL, lngPos - 1))
End If
RemoveOrderBy_F = strSQL
End Function


Public Function GetFileOf(ByVal strTarg$, Optional strDivider$ = "\") As String
Dim strRet$
Dim i&

strRet = strTarg
For i = Len(strTarg) To 1 Step -1
    If Mid$(strTarg, i, 1) = strDivider Then
        strRet = Trim$(Right$(strTarg, Len(strTarg) - i))
        Exit For
    End If
Next
GetFileOf = strRet
End Function


Public Sub Str2ByteArray(ByVal StringIn As String, ByRef ByteArray() As Byte)
    ByteArray = StrConv(StringIn, vbFromUnicode)
End Sub
Public Sub ByteArray2Str(ByteArray() As Byte, ByRef StringIn As String)
    StringIn = StrConv(ByteArray, vbUnicode)
End Sub


Public Function SubStringCountOf(ByVal TargetStr As String, ByVal SubStr As String) As Long
Dim l&, lStart&, lCnt&
Dim vRet As Variant
Dim s$
lStart = 1
On Error GoTo errh
Do Until lStart > Len(TargetStr)
    vRet = InStr(lStart, TargetStr, SubStr, vbTextCompare)
    If VarType(vRet) = vbLong Or VarType(vRet) = vbInteger Then
        If CLng(vRet) > 0 Then
            lCnt = lCnt + 1
            lStart = CLng(vRet) + Len(SubStr)
        Else
            Exit Do
        End If
    End If
Loop
SubStringCountOf = lCnt
Exit Function
errh:
SubStringCountOf = 0
Err.Clear
Exit Function
End Function

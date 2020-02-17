VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FuncLib_Lib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type

Public Function CursorGetPosInTwips(ByRef x As Long, ByRef y As Long) As Boolean
Dim l&
Dim pt As POINTAPI
On Error GoTo errh
l = GetCursorPos(pt)
If l > 0 Then 'non zero is success!
    x = pt.x
    y = pt.y
    x = x * Screen.TwipsPerPixelX
    y = y * Screen.TwipsPerPixelY
    CursorGetPosInTwips = True
Else
    CursorGetPosInTwips = False
End If
Exit Function
errh:
Err.Clear
CursorGetPosInTwips = False
Exit Function
End Function

Friend Sub FormCenter(frm As Form, Optional mdi As MDIForm = Nothing, Optional SizeMDI As Boolean)
On Error Resume Next
If frm.MDIChild = True And frm.BorderStyle = 2 And Not mdi Is Nothing Then
    frm.Top = 100
    frm.Left = (mdi.Width - frm.Width) / 2
    If SizeMDI = True Then frm.Height = mdi.Height - 2000
ElseIf frm.MDIChild = True And frm.BorderStyle = 1 And Not mdi Is Nothing Then
    frm.Top = (mdi.ScaleHeight - frm.Height) / 2
    frm.Left = (mdi.ScaleWidth - frm.Width) / 2
Else
    frm.Top = (Screen.Height - frm.Height) / 2
    frm.Left = (Screen.Width - frm.Width) / 2
End If
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Launcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Sub WriteWindowed(ByVal windowed As Boolean)
Dim l&
pcIO.MoveFirst
l = pcIO.SearchRows(ioFromStart, "windowed=", True)
If l > 0 Then
    If windowed Then
        pcIO.WriteData ReplaceData, "windowed=1"
    Else
        pcIO.WriteData ReplaceData, "windowed=0"
    End If
End If
End Sub

Public Sub Launch()
Dim s$
s = App.Path & "\Camelot.exe"
Shell s, vbNormalFocus
End Sub

Public Sub WriteResolution(ByVal height&, ByVal width&)
Dim l&
pcIO.MoveFirst
l = pcIO.SearchRows(ioFromStart, "screen_height", True)
If l > 0 Then
    pcIO.WriteData ReplaceData, "screen_height=" & CStr(height)
End If
l = pcIO.SearchRows(ioFromStart, "screen_width", True)
If l > 0 Then
    pcIO.WriteData ReplaceData, "screen_width=" & CStr(width)
End If
End Sub


Attribute VB_Name = "modMain"
Option Explicit
Option Compare Text
Public pcIO As IOFileMan
Public pcLauncher As Launcher
Public Sub ShowError()
MsgBox Err.Description, vbCritical
End Sub

Public Sub Main()
Dim l&
Dim s$
On Error GoTo errh
s = App.Path & "\camelot.exe"
If Dir$(s, vbNormal) = "" Then
    MsgBox "Cannot find camelot.exe"
    Set pcIO = Nothing
    Exit Sub
End If
Set pcIO = New IOFileMan
pcIO.OpenFile App.Path, "user.dat"
pcIO.MoveFirst
Set pcLauncher = New Launcher
frmLauncher.Show vbModal
Exit Sub
errh:
 MsgBox Err.Description
 Set pcIO = Nothing
End Sub


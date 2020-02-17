VERSION 5.00
Begin VB.Form frmLauncher 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Camelot Launcher"
   ClientHeight    =   1635
   ClientLeft      =   6225
   ClientTop       =   4335
   ClientWidth     =   3030
   Icon            =   "frmLauncher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3030
   Begin VB.CheckBox chkUnloadMe 
      Caption         =   "Unload me on launch"
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   840
      Value           =   1  'Checked
      Width           =   1965
   End
   Begin VB.ComboBox cmbResolution 
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   90
      Width           =   2835
   End
   Begin VB.CommandButton cmdLaunch 
      Caption         =   "&Launch"
      Default         =   -1  'True
      Height          =   315
      Left            =   1260
      TabIndex        =   2
      Top             =   1260
      Width           =   795
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   315
      Left            =   2130
      TabIndex        =   1
      Top             =   1260
      Width           =   795
   End
   Begin VB.CheckBox chkWindowed 
      Caption         =   "Launch windowed"
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
End
Attribute VB_Name = "frmLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Sub cmdExit_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cmdLaunch_Click()
Dim s$
Dim h&, w&
On Error GoTo errh

Call pcLauncher.WriteWindowed(CBool(Me.chkWindowed.Value))

s = Me.cmbResolution.Text
If s <> "current" Then
    Select Case True
        Case s = "800x600"
            w = 800
            h = 600
        Case s = "1024x768"
            w = 1024
            h = 768
        Case s = "1280x1024"
            w = 1280
            h = 1024
    End Select
    Call pcLauncher.WriteResolution(h, w)
End If
Call pcLauncher.Launch
If Me.chkUnloadMe > 0 Then Unload Me
Exit Sub
errh:
 ShowError
End Sub

Private Sub Form_Load()
On Error GoTo errh
Me.Top = (Screen.height - Me.height) / 2
Me.Left = (Screen.width - Me.width) / 2
Call x_ComboFill
Me.cmbResolution.ListIndex = 0
Exit Sub
errh:
 ShowError
End Sub

Public Sub x_ComboFill()
With Me.cmbResolution
    .AddItem "Current"
    .AddItem "800x600"
    .AddItem "1024x768"
    .AddItem "1280x1024"
End With
End Sub

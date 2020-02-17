VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   2370
   ClientLeft      =   3120
   ClientTop       =   4260
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   6585
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   570
      TabIndex        =   0
      Top             =   570
      Width           =   2025
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim cDMO As New FuncLib.DMO_Lib

Dim List_() As String
Dim b As Boolean
b = cDMO.SQLServer2000ServerList(List_)
End Sub


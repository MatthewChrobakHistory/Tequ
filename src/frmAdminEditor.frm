VERSION 5.00
Begin VB.Form frmAdminEditor 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtMap 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Map:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmAdminEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
    Player(MyIndex).Map = txtMap.text
    Call WarpPlayer(MyIndex, txtMap.text, Player(MyIndex).X, Player(MyIndex).Y)
End Sub

Private Sub txtMap_Change()
    If Not IsNumeric(txtMap.text) Then txtMap.text = "1"
End Sub

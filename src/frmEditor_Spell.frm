VERSION 5.00
Begin VB.Form frmEditor_Spell 
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   6120
      TabIndex        =   22
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   6120
      TabIndex        =   21
      Top             =   240
      Width           =   375
   End
   Begin VB.HScrollBar scrlMap 
      Height          =   255
      Left            =   4080
      TabIndex        =   20
      Top             =   600
      Width           =   1575
   End
   Begin VB.HScrollBar scrlAOE 
      Height          =   255
      Left            =   4440
      TabIndex        =   18
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtStunDuration 
      Height          =   285
      Left            =   3480
      TabIndex        =   16
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtCoolDown 
      Height          =   285
      Left            =   3000
      TabIndex        =   13
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txtVital 
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   10
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtVital 
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   9
      Top             =   2400
      Width           =   1935
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmEditor_Spell.frx":0000
      Left            =   3720
      List            =   "frmEditor_Spell.frx":000D
      TabIndex        =   8
      Text            =   "Vital Affect"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.HScrollBar scrlRange 
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.PictureBox picPicture 
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   3000
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   5
      Top             =   480
      Width           =   1020
   End
   Begin VB.HScrollBar scrlPicture 
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox lstIndex 
      Height          =   5325
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "X:        Y:"
      Height          =   495
      Left            =   5880
      TabIndex        =   23
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblMap 
      Caption         =   "Map: 0"
      Height          =   255
      Left            =   4080
      TabIndex        =   19
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblAOE 
      Caption         =   "AOE: 0"
      Height          =   255
      Left            =   3600
      TabIndex        =   17
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Stun Duration:"
      Height          =   255
      Left            =   2400
      TabIndex        =   15
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Cooldown:"
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Spirit"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Health"
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblRange 
      Caption         =   "Self-Cast"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblPicture 
      Caption         =   "Picture: 0"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmEditor_Spell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbType_Click()
    Spell(ESpellNum).Type = cmbType.ListIndex
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim I As Long
    For I = 1 To MAX_SPELLS
        Call SaveSpell(I)
    Next
End Sub

Private Sub lstIndex_Click()
    ESpellNum = lstIndex.ListIndex + 1
    Call NewSpellIndex(ESpellNum)
End Sub

Private Sub scrlAOE_Change()
    Spell(ESpellNum).AOE = scrlAOE.Value
End Sub

Private Sub scrlMap_Change()
    Spell(ESpellNum).Map = scrlMap.Value
    lblMap.Caption = "Map: " & scrlMap.Value
End Sub

Private Sub scrlPicture_Change()
    Spell(ESpellNum).Picture = scrlPicture.Value
    lblPicture.Caption = "Picture: " & scrlPicture.Value
    picPicture.Picture = Nothing
    If FileExist(App.Path & "\graphics\spells\" & scrlPicture.Value & ".bmp") Then
        picPicture.Picture = LoadPicture(App.Path & "\graphics\spells\" & scrlPicture.Value & ".bmp")
    End If
End Sub

Private Sub scrlRange_Change()
    Spell(ESpellNum).Range = scrlRange.Value
    If Spell(ESpellNum).Range = 0 Then
        lblRange.Caption = "Self-Cast"
    Else
        lblRange.Caption = "Range: " & scrlRange.Value
    End If
End Sub

Private Sub txtCoolDown_Change()
    If Not IsNumeric(txtCoolDown.text) Then
        txtCoolDown.text = Spell(ESpellNum).CoolDown
    End If
    Spell(ESpellNum).CoolDown = txtCoolDown.text
End Sub

Private Sub txtName_Change()
    Spell(ESpellNum).name = txtName.text
End Sub

Private Sub txtStunDuration_Change()
    If Not IsNumeric(txtStunDuration.text) Then
        txtStunDuration.text = Spell(ESpellNum).StunDuration
    End If
    Spell(ESpellNum).StunDuration = txtStunDuration.text
End Sub

Private Sub txtVital_Change(Index As Integer)
    If Not IsNumeric(txtVital(Index).text) Then
        txtVital(Index).text = Spell(ESpellNum).VitalAffect(Index)
    End If
    Spell(ESpellNum).VitalAffect(Index) = txtVital(Index).text
End Sub

Private Sub txtX_Change()
    If Not IsNumeric(txtX.text) Then
        txtX.text = Spell(ESpellNum).X
    End If
    Spell(ESpellNum).X = txtX.text
End Sub

Private Sub txtY_Change()
    If Not IsNumeric(txtY.text) Then
        txtY.text = Spell(ESpellNum).Y
    End If
    Spell(ESpellNum).Y = txtY.text
End Sub

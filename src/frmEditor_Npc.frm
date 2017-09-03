VERSION 5.00
Begin VB.Form frmEditor_Npc 
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   8400
      TabIndex        =   48
      Top             =   840
      Width           =   2055
      Begin VB.TextBox txtChance 
         Height          =   285
         Left            =   820
         TabIndex        =   55
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtValue 
         Height          =   285
         Left            =   600
         TabIndex        =   53
         Top             =   1080
         Width           =   1335
      End
      Begin VB.HScrollBar scrlDrop 
         Height          =   255
         Left            =   120
         Min             =   1
         TabIndex        =   51
         Top             =   0
         Value           =   1
         Width           =   1815
      End
      Begin VB.HScrollBar scrlItem 
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Value           =   1
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Chance:"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblValue 
         Caption         =   "Value: "
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblItem 
         Caption         =   "Item:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.TextBox txtDamage 
      Height          =   285
      Left            =   8160
      TabIndex        =   47
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox cmbAttackType 
      Height          =   315
      ItemData        =   "frmEditor_Npc.frx":0000
      Left            =   5280
      List            =   "frmEditor_Npc.frx":0010
      TabIndex        =   45
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stat Bonus"
      Height          =   1695
      Left            =   5400
      TabIndex        =   34
      Top             =   3720
      Width           =   4455
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   1215
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   1215
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   37
         Top             =   600
         Width           =   1215
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   36
         Top             =   1200
         Width           =   1215
      End
      Begin VB.HScrollBar scrlStat 
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   35
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblStat 
         Caption         =   "Stat: 0"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblStat 
         Caption         =   "Stat: 0"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   43
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblStat 
         Caption         =   "Stat: 0"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   42
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblStat 
         Caption         =   "Stat: 0"
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   41
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblStat 
         Caption         =   "Stat: 0"
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   40
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.TextBox txtXP 
      Height          =   285
      Left            =   6840
      TabIndex        =   32
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Offense"
      Height          =   1455
      Left            =   2760
      TabIndex        =   25
      Top             =   4800
      Width           =   2295
      Begin VB.TextBox txtOffense 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtOffense 
         Height          =   285
         Index           =   2
         Left            =   720
         TabIndex        =   27
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtOffense 
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   26
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Melee:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Range:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Magic:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Defense"
      Height          =   1455
      Left            =   2760
      TabIndex        =   18
      Top             =   3240
      Width           =   2295
      Begin VB.TextBox txtDefense 
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   21
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtDefense 
         Height          =   285
         Index           =   2
         Left            =   720
         TabIndex        =   20
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtDefense 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Magic:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Range:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Melee:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.HScrollBar scrlAttackSpeed 
      Height          =   255
      Left            =   6360
      Max             =   30
      Min             =   1
      TabIndex        =   17
      Top             =   3000
      Value           =   1
      Width           =   1815
   End
   Begin VB.TextBox txtRange 
      Height          =   285
      Left            =   6960
      TabIndex        =   15
      Top             =   2280
      Width           =   1335
   End
   Begin VB.HScrollBar scrlSpeed 
      Height          =   255
      Left            =   3720
      Max             =   30
      Min             =   1
      TabIndex        =   13
      Top             =   2880
      Value           =   1
      Width           =   1575
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmEditor_Npc.frx":0030
      Left            =   5280
      List            =   "frmEditor_Npc.frx":0040
      TabIndex        =   11
      Text            =   "cmbType"
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtSpirit 
      Height          =   285
      Left            =   4080
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtHealth 
      Height          =   285
      Left            =   4080
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtRespawn 
      Height          =   285
      Left            =   3480
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.HScrollBar scrlSprite 
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox lstIndex 
      Height          =   5910
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Damage:"
      Height          =   255
      Left            =   7440
      TabIndex        =   46
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "XP:"
      Height          =   255
      Left            =   6480
      TabIndex        =   33
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label lblAttackSpeed 
      Caption         =   "Attack Speed: 0"
      Height          =   255
      Left            =   6360
      TabIndex        =   16
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Range:"
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblSpeed 
      Caption         =   "Speed: 1"
      Height          =   375
      Left            =   3720
      TabIndex        =   12
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Spirit:"
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblHealth 
      Caption         =   "Health:"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblRespawn 
      Caption         =   "Respawn:"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblSprite 
      Caption         =   "Sprite: 0"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmEditor_Npc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbAttackType_Change()
    Npc(ENpcNum).AttackType = cmbAttackType.ListIndex
End Sub

Private Sub cmbType_Click()

    Npc(ENpcNum).Type = cmbType.ListIndex

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim I As Long
    For I = 1 To MAX_NPCS
        Call SaveNpc(I)
    Next
End Sub

Private Sub lstIndex_Click()

    ENpcNum = lstIndex.ListIndex + 1
    Call NewNpcIndex(ENpcNum)

End Sub

Private Sub scrlRespawn_Change()

    Npc(ENpcNum).Respawn = scrlRespawn.Value
    lblRespawn.Caption = "Respawn: " & scrlRespawn.Value
    
End Sub

Private Sub scrlAttackSpeed_Change()

    Npc(ENpcNum).AttackSpeed = scrlAttackSpeed.Value / 10
    lblAttackSpeed.Caption = "Attack Speed: " & Npc(ENpcNum).AttackSpeed

End Sub

Private Sub scrlDrop_Change()
    With Npc(ENpcNum).Drop(scrlDrop.Value)
        txtValue.text = .Value
        txtChance.text = .Chance
        scrlItem.Value = .Item
    End With
End Sub

Private Sub scrlItem_Change()
    Npc(ENpcNum).Drop(scrlDrop.Value).Item = scrlItem.Value
    If scrlItem.Value = 0 Then
        lblItem.Caption = "Item: None"
    Else
        lblItem.Caption = "Item: " + Trim$(Item(scrlItem.Value).name)
    End If
End Sub

Private Sub scrlSpeed_Change()

    Npc(ENpcNum).Speed = scrlSpeed.Value / 10
    lblSpeed.Caption = "Speed: " & Npc(ENpcNum).Speed
    
End Sub

Private Sub scrlSprite_Change()

    Npc(ENpcNum).Sprite = scrlSprite.Value
    lblSprite.Caption = "Sprite: " & scrlSprite.Value

End Sub

Private Sub scrlStat_Change(Index As Integer)
    Npc(ENpcNum).Stat(Index) = scrlStat(Index).Value

    Select Case Index
        Case Stats.Attack
            lblStat(Index).Caption = "Att: " & scrlStat(Index).Value
        Case Stats.Strength
            lblStat(Index).Caption = "Str: " & scrlStat(Index).Value
        Case Stats.Defense
            lblStat(Index).Caption = "Def: " & scrlStat(Index).Value
        Case Stats.Agility
            lblStat(Index).Caption = "Agi: " & scrlStat(Index).Value
        Case Stats.Sagacity
            lblStat(Index).Caption = "Sag: " & scrlStat(Index).Value
    End Select
End Sub

Private Sub txtChance_Change()
    If Not IsNumeric(txtChance.text) Then
        txtChance.text = Npc(ENpcNum).Drop(scrlDrop.Value).Chance
    ElseIf txtChance.text > 100 Then
        txtChance.text = "100"
    End If
    
    Npc(ENpcNum).Drop(scrlDrop.Value).Chance = txtChance.text
End Sub

Private Sub txtDamage_Change()
    If Not IsNumeric(txtDamage.text) Then
        txtDamage.text = Npc(ENpcNum).Damage
    End If
    Npc(ENpcNum).Damage = txtDamage.text
End Sub

Private Sub txtDefense_Change(Index As Integer)
    If Not IsNumeric(txtDefense(Index).text) Then
        txtDefense(Index).text = Npc(ENpcNum).Defense(Index)
    End If
    Npc(ENpcNum).Defense(Index) = txtDefense(Index).text
End Sub

Private Sub txtHealth_Change()

    If Not IsNumeric(txtHealth.text) Then
        If Npc(ENpcNum).Vital(Vitals.Health) = 0 Then
            txtHealth.text = "1"
        Else
            txtHealth.text = Npc(ENpcNum).Vital(Vitals.Health)
        End If
    End If
    
    Npc(ENpcNum).Vital(Vitals.Health) = txtHealth.text

End Sub

Private Sub txtName_Change()

    Npc(ENpcNum).name = txtName.text

End Sub

Private Sub txtOffense_Change(Index As Integer)
    If Not IsNumeric(txtOffense(Index).text) Then
        txtOffense(Index).text = Npc(ENpcNum).Offense(Index)
    End If
    Npc(ENpcNum).Offense(Index) = txtOffense(Index).text
End Sub

Private Sub txtRange_Change()

    If Not IsNumeric(txtRange.text) Then
        txtRange.text = "1"
    End If
    
    If txtRange.text < 0 Or txtRange.text > 255 Then txtRange.text = "1"
    Npc(ENpcNum).Range = txtRange.text

End Sub

Private Sub txtRespawn_Change()

    If IsNumeric(txtRespawn.text) = False Then
        txtRespawn.text = Npc(ENpcNum).Respawn
    End If
    
    Npc(ENpcNum).Respawn = txtRespawn.text

End Sub

Private Sub txtSpirit_Change()

    If Not IsNumeric(txtSpirit.text) Then
        If Npc(ENpcNum).Vital(Vitals.Spirit) = 0 Then
            txtSpirit.text = "1"
        Else
            txtSpirit.text = Npc(ENpcNum).Vital(Vitals.Spirit)
        End If
    End If
    
    Npc(ENpcNum).Vital(Vitals.Spirit) = txtSpirit.text

End Sub

Private Sub txtValue_Change()
    If Not IsNumeric(txtValue.text) Then
        txtValue.text = Npc(ENpcNum).Drop(scrlDrop.Value).Value
    End If
    
    Npc(ENpcNum).Drop(scrlDrop.Value).Value = txtValue.text
End Sub

Private Sub txtXP_Change()
    If Not IsNumeric(txtXP.text) Then
        txtXP.text = Npc(ENpcNum).XP
    End If
    Npc(ENpcNum).XP = txtXP.text
End Sub

VERSION 5.00
Begin VB.Form frmEditor_Resource 
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbEquipmentType 
      Height          =   315
      ItemData        =   "frmEditor_Resource.frx":0000
      Left            =   2280
      List            =   "frmEditor_Resource.frx":0010
      TabIndex        =   49
      Text            =   "None"
      Top             =   480
      Width           =   2055
   End
   Begin VB.HScrollBar scrlRewardValue 
      Height          =   255
      Left            =   5520
      TabIndex        =   48
      Top             =   1200
      Width           =   1455
   End
   Begin VB.HScrollBar scrlItem 
      Height          =   255
      Left            =   5520
      TabIndex        =   46
      Top             =   840
      Width           =   1455
   End
   Begin VB.HScrollBar scrlDead 
      Height          =   255
      Left            =   5520
      TabIndex        =   44
      Top             =   480
      Width           =   1455
   End
   Begin VB.HScrollBar scrlAlive 
      Height          =   255
      Left            =   5520
      TabIndex        =   43
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame7 
      Caption         =   "Skill Reward"
      Height          =   3375
      Left            =   4920
      TabIndex        =   24
      Top             =   1800
      Width           =   2535
      Begin VB.TextBox txtSkillReward 
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   32
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReward 
         Height          =   285
         Index           =   7
         Left            =   1200
         TabIndex        =   31
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReward 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   30
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReward 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   29
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReward 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   28
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReward 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   27
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReward 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReward 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblSkill 
         Caption         =   "PotionBrew:"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   40
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Crafting:"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   39
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Fletching:"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   38
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Cooking:"
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   37
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Smithing:"
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Fishing:"
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Mining:"
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Woodcutting:"
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Skill Requirements"
      Height          =   3375
      Left            =   2280
      TabIndex        =   7
      Top             =   1800
      Width           =   2535
      Begin VB.TextBox txtSkillReq 
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   15
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReq 
         Height          =   285
         Index           =   7
         Left            =   1200
         TabIndex        =   14
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReq 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   13
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReq 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReq 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReq 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReq 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReq 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblSkill 
         Caption         =   "PotionBrew:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   23
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Crafting:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Fletching:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Cooking:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Smithing:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Fishing:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Mining:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Woodcutting:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.TextBox txtHealth 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtRespawn 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstIndex 
      Height          =   5130
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblRewardValue 
      Caption         =   "Value: 0"
      Height          =   255
      Left            =   4680
      TabIndex        =   47
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblItem 
      Caption         =   "Item: 0"
      Height          =   255
      Left            =   4680
      TabIndex        =   45
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblDeadGFX 
      Caption         =   "Dead: 0"
      Height          =   255
      Left            =   4680
      TabIndex        =   42
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblAliveGFX 
      Caption         =   "Alive: 0"
      Height          =   255
      Left            =   4680
      TabIndex        =   41
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Health:"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblRespawn 
      Caption         =   "Respawn:"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmEditor_Resource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbEquipmentType_Click()

    Resource(EResourceNum).EquipmentType = cmbEquipmentType.ListIndex

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim I As Long
    For I = 1 To MAX_RESOURCES
        Call SaveResource(I)
    Next
End Sub

Private Sub lstIndex_Click()

    EResourceNum = lstIndex.ListIndex + 1
    Call NewResourceIndex(EResourceNum)

End Sub

Private Sub scrlAlive_Change()

    Resource(EResourceNum).AliveGFX = scrlAlive.Value
    lblAliveGFX.Caption = "Alive: " & scrlAlive.Value

End Sub

Private Sub scrlDead_Change()

    Resource(EResourceNum).DeadGFX = scrlDead.Value
    lblDeadGFX.Caption = "Dead: " & scrlDead.Value

End Sub

Private Sub scrlItem_Change()

    Resource(EResourceNum).Reward = scrlItem.Value
    lblItem.Caption = "Item: " & scrlItem.Value

End Sub

Private Sub scrlRewardValue_Change()

    Resource(EResourceNum).RewardValue = scrlRewardValue.Value
    lblRewardValue.Caption = "Amount: " & scrlRewardValue.Value

End Sub

Private Sub txtHealth_Change()

    If Not IsNumeric(txtHealth.text) Then
        txtHealth.text = Resource(EResourceNum).Health
    End If
    
    Resource(EResourceNum).Health = txtHealth.text

End Sub

Private Sub txtName_Change()

    Resource(EResourceNum).name = txtName.text

End Sub

Private Sub txtRespawn_Change()

    If IsNumeric(txtRespawn.text) = False Then
        txtRespawn.text = Resource(EResourceNum).RespawnRate
    End If
    
    Resource(EResourceNum).RespawnRate = txtRespawn.text

End Sub

Private Sub txtSkillReq_Change(Index As Integer)

    If Not IsNumeric(txtSkillReq(Index).text) Then
        txtSkillReq(Index).text = Resource(EResourceNum).RequiredXP(Index)
    End If
    
    Resource(EResourceNum).RequiredXP(Index) = txtSkillReq(Index).text

End Sub

Private Sub txtSkillReward_Change(Index As Integer)

    If Not IsNumeric(txtSkillReward(Index).text) Then
        txtSkillReward(Index).text = Resource(EResourceNum).RewardXP(Index)
    End If
    
    Resource(EResourceNum).RewardXP(Index) = txtSkillReward(Index).text

End Sub

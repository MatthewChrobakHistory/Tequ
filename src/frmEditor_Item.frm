VERSION 5.00
Begin VB.Form frmEditor_Item 
   Caption         =   "Form1"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   ScaleHeight     =   610
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   831
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar scrlPSpeed 
      Height          =   255
      Left            =   10680
      TabIndex        =   135
      Top             =   7320
      Width           =   1575
   End
   Begin VB.HScrollBar scrlPRange 
      Height          =   255
      Left            =   10680
      TabIndex        =   133
      Top             =   6720
      Width           =   1695
   End
   Begin VB.HScrollBar scrlPImage 
      Height          =   255
      Left            =   10680
      TabIndex        =   131
      Top             =   6120
      Width           =   1695
   End
   Begin VB.ComboBox cmbEquipmentType 
      Height          =   315
      ItemData        =   "frmEditor_Item.frx":0000
      Left            =   3720
      List            =   "frmEditor_Item.frx":0010
      TabIndex        =   127
      Text            =   "None"
      Top             =   1080
      Width           =   855
   End
   Begin VB.Frame Frame7 
      Caption         =   "Making Reward"
      Height          =   3375
      Left            =   8040
      TabIndex        =   110
      Top             =   5640
      Width           =   2535
      Begin VB.TextBox txtSkillReward 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   118
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReward 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   117
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReward 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   116
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReward 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   115
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReward 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   114
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReward 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   113
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReward 
         Height          =   285
         Index           =   7
         Left            =   1200
         TabIndex        =   112
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReward 
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   111
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblSkill 
         Caption         =   "Woodcutting:"
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   126
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Mining:"
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   125
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Fishing:"
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   124
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Smithing:"
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   123
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Cooking:"
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   122
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Fletching:"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   121
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Crafting:"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   120
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "PotionBrew:"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   119
         Top             =   2880
         Width           =   1455
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Making Requirements"
      Height          =   3375
      Left            =   5400
      TabIndex        =   93
      Top             =   5640
      Width           =   2535
      Begin VB.TextBox txtSkillMake 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   101
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSkillMake 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   100
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtSkillMake 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   99
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtSkillMake 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   98
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtSkillMake 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   97
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtSkillMake 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   96
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtSkillMake 
         Height          =   285
         Index           =   7
         Left            =   1200
         TabIndex        =   95
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtSkillMake 
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   94
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblSkill 
         Caption         =   "Woodcutting:"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   109
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Mining:"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   108
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Fishing:"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   107
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Smithing:"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   106
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Cooking:"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   105
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Fletching:"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   104
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Crafting:"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   103
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "PotionBrew:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   102
         Top             =   2880
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Wielding Requirements"
      Height          =   3375
      Left            =   2760
      TabIndex        =   76
      Top             =   5640
      Width           =   2535
      Begin VB.TextBox txtSkillReq 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   84
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReq 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   83
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReq 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   82
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReq 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   81
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReq 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   80
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReq 
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   79
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReq 
         Height          =   285
         Index           =   7
         Left            =   1200
         TabIndex        =   78
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtSkillReq 
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   77
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblSkill 
         Caption         =   "Woodcutting:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   92
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Mining:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   91
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Fishing:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   90
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Smithing:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   89
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Cooking:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   88
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Fletching:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   87
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "Crafting:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   86
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblSkill 
         Caption         =   "PotionBrew:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   85
         Top             =   2880
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Defense"
      Height          =   1455
      Left            =   7560
      TabIndex        =   65
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txtDefense 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   68
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtDefense 
         Height          =   285
         Index           =   2
         Left            =   720
         TabIndex        =   67
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtDefense 
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   66
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Melee:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   71
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Range:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   70
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Magic:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   69
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Offense"
      Height          =   1455
      Left            =   7560
      TabIndex        =   58
      Top             =   1680
      Width           =   2295
      Begin VB.TextBox txtOffense 
         Height          =   285
         Index           =   3
         Left            =   720
         TabIndex        =   64
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtOffense 
         Height          =   285
         Index           =   2
         Left            =   720
         TabIndex        =   62
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtOffense 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   60
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Magic:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   63
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Range:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   61
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Melee:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CheckBox chkStackable 
      Caption         =   "Stackable?"
      Height          =   255
      Left            =   2760
      TabIndex        =   57
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame fraCombat 
      Caption         =   "Combat Equipment"
      Height          =   3735
      Left            =   2760
      TabIndex        =   34
      Top             =   1800
      Width           =   4695
      Begin VB.Frame Frame2 
         Caption         =   "Stat Requirements"
         Height          =   1695
         Left            =   120
         TabIndex        =   46
         Top             =   1920
         Width           =   4455
         Begin VB.HScrollBar scrlStatReq 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   600
            Width           =   1215
         End
         Begin VB.HScrollBar scrlStatReq 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   50
            Top             =   1200
            Width           =   1215
         End
         Begin VB.HScrollBar scrlStatReq 
            Height          =   255
            Index           =   3
            Left            =   1440
            TabIndex        =   49
            Top             =   600
            Width           =   1215
         End
         Begin VB.HScrollBar scrlStatReq 
            Height          =   255
            Index           =   4
            Left            =   1440
            TabIndex        =   48
            Top             =   1200
            Width           =   1215
         End
         Begin VB.HScrollBar scrlStatReq 
            Height          =   255
            Index           =   5
            Left            =   2880
            TabIndex        =   47
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblStatReq 
            Caption         =   "Stat: 0"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblStatReq 
            Caption         =   "Stat: 0"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   55
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblStatReq 
            Caption         =   "Stat: 0"
            Height          =   255
            Index           =   3
            Left            =   1440
            TabIndex        =   54
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblStatReq 
            Caption         =   "Stat: 0"
            Height          =   255
            Index           =   4
            Left            =   1440
            TabIndex        =   53
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblStatReq 
            Caption         =   "Stat: 0"
            Height          =   255
            Index           =   5
            Left            =   2880
            TabIndex        =   52
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Stat Bonus"
         Height          =   1695
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   4455
         Begin VB.ComboBox cmbCombatType 
            Height          =   315
            ItemData        =   "frmEditor_Item.frx":0039
            Left            =   3480
            List            =   "frmEditor_Item.frx":0049
            TabIndex        =   136
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox txtSpeed 
            Height          =   285
            Left            =   3600
            TabIndex        =   75
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtDamage 
            Height          =   285
            Left            =   3600
            TabIndex        =   73
            Top             =   960
            Width           =   735
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   5
            Left            =   2880
            TabIndex        =   45
            Top             =   600
            Width           =   1215
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   4
            Left            =   1440
            TabIndex        =   44
            Top             =   1200
            Width           =   1215
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   3
            Left            =   1440
            TabIndex        =   43
            Top             =   600
            Width           =   1215
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   42
            Top             =   1200
            Width           =   1215
         End
         Begin VB.HScrollBar scrlStat 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   41
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Speed:"
            Height          =   255
            Left            =   2880
            TabIndex        =   74
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Damage:"
            Height          =   255
            Left            =   2880
            TabIndex        =   72
            Top             =   960
            Width           =   1335
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
         Begin VB.Label lblStat 
            Caption         =   "Stat: 0"
            Height          =   255
            Index           =   4
            Left            =   1440
            TabIndex        =   39
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblStat 
            Caption         =   "Stat: 0"
            Height          =   255
            Index           =   3
            Left            =   1440
            TabIndex        =   38
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblStat 
            Caption         =   "Stat: 0"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblStat 
            Caption         =   "Stat: 0"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   1215
         End
      End
   End
   Begin VB.HScrollBar scrlGiveBack 
      Height          =   255
      Left            =   10200
      TabIndex        =   33
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Frame fraConsume 
      Caption         =   "Consume Item Data"
      Height          =   2295
      Left            =   7560
      TabIndex        =   27
      Top             =   3240
      Visible         =   0   'False
      Width           =   2295
      Begin VB.HScrollBar scrlLearnSpell 
         Height          =   255
         Left            =   120
         TabIndex        =   128
         Top             =   1920
         Width           =   2055
      End
      Begin VB.HScrollBar scrladdSP 
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   2055
      End
      Begin VB.HScrollBar scrladdHP 
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblLearnSpell 
         Caption         =   "Spell: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lbladdSP 
         Caption         =   "Spirit: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lbladdHP 
         Caption         =   "Health: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.CheckBox chkBltGraphics 
      Caption         =   "blt Player Graphic?"
      Height          =   255
      Left            =   5520
      TabIndex        =   26
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtInfo 
      Height          =   855
      Left            =   5160
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton cmdRefreshList 
      Caption         =   "Refresh List"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   60
      Width           =   2535
   End
   Begin VB.CheckBox chkTwoHanded 
      Caption         =   "Two Handed?"
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Frame fraPaperdolls 
      Caption         =   "Paperdolls"
      Height          =   4455
      Left            =   9960
      TabIndex        =   9
      Top             =   120
      Width           =   2415
      Begin VB.HScrollBar scrlStance 
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   4080
         Width           =   2175
      End
      Begin VB.HScrollBar scrlStance 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   3480
         Width           =   2175
      End
      Begin VB.HScrollBar scrlStance 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   2880
         Width           =   2175
      End
      Begin VB.HScrollBar scrlStance 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   2175
      End
      Begin VB.HScrollBar scrlStance 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   2175
      End
      Begin VB.HScrollBar scrlStance 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblStance 
         Caption         =   "Female Two Hand: 0"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label lblStance 
         Caption         =   "Female Shield: 0"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label lblStance 
         Caption         =   "Female Normal: 0"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label lblStance 
         Caption         =   "Male Two Hand: 0"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lblStance 
         Caption         =   "Male Shield: 0"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblStance 
         Caption         =   "Male Normal: 0"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   5760
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox lstType 
      Height          =   315
      ItemData        =   "frmEditor_Item.frx":0069
      Left            =   2760
      List            =   "frmEditor_Item.frx":008E
      TabIndex        =   6
      Text            =   "None"
      Top             =   1080
      Width           =   855
   End
   Begin VB.HScrollBar scrlPicNum 
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.PictureBox picItem 
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   4080
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   3
      Top             =   480
      Width           =   1020
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   3360
      MaxLength       =   25
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox lstIndex 
      Height          =   8445
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lblPSpeed 
      Caption         =   "Projectile Speed: 0"
      Height          =   255
      Left            =   10680
      TabIndex        =   134
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label lblPRange 
      Caption         =   "Projectile Range: 0"
      Height          =   255
      Left            =   10680
      TabIndex        =   132
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label lblPImage 
      Caption         =   "Projectile Image: 0"
      Height          =   255
      Left            =   10680
      TabIndex        =   130
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label lblGiveBack 
      Caption         =   "Give Back Item: 0"
      Height          =   255
      Left            =   10080
      TabIndex        =   32
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Info:"
      Height          =   255
      Left            =   4320
      TabIndex        =   25
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "price:"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblPicNum 
      Caption         =   "Num: 0"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkBltGraphics_Click()

    If chkBltGraphics.Value = 1 Then
        Item(EItemNum).BltPlayerGraphics = True
    Else
        Item(EItemNum).BltPlayerGraphics = False
    End If
    
End Sub

Private Sub chkStackable_Click()

    If chkStackable.Value = 1 Then
        Item(EItemNum).Stackable = True
    Else
        Item(EItemNum).Stackable = False
    End If

End Sub

Private Sub chkTwoHanded_Click()

    If chkTwoHanded.Value = 1 Then
        Item(EItemNum).IsTwoHanded = True
    Else
        Item(EItemNum).IsTwoHanded = False
    End If

End Sub

Private Sub cmbCombatType_Click()
    Item(EItemNum).CombatType = cmbCombatType.ListIndex
End Sub

Private Sub cmbEquipmentType_Click()

    Item(EItemNum).EquipmentType = cmbEquipmentType.ListIndex

End Sub

Private Sub cmdRefreshList_Click()

    Call InitItemEditor

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim I As Long
    For I = 1 To MAX_ITEMS
        Call SaveItem(I)
    Next
End Sub

Private Sub Label5_Click()

    MsgBox GetPlayerDamage(EItemNum)

End Sub
Private Sub lstIndex_Click()
    
    EItemNum = lstIndex.ListIndex + 1
    Call NewItemIndex(EItemNum)

End Sub

Private Sub lstType_Click()

    Item(EItemNum).Type = lstType.ListIndex
    
With frmEditor_Item
    .fraConsume.Visible = False
    Select Case lstType.ListIndex
        Case ITEM_TYPE_CONSUME
            .fraConsume.Visible = True
    End Select
End With

End Sub

Private Sub scrladdHP_Change()

    Item(EItemNum).addHP = scrladdHP.Value
    lbladdHP.Caption = "Health: " & scrladdHP.Value

End Sub

Private Sub scrladdSP_Change()

    Item(EItemNum).addSP = scrladdSP.Value
    lbladdSP.Caption = "Spirit: " & scrladdSP.Value

End Sub

Private Sub scrlGiveBack_Change()

    Item(EItemNum).GiveBack = scrlGiveBack.Value
    lblGiveBack.Caption = "Give Back Item: " & scrlGiveBack.Value

End Sub

Private Sub scrlLearnSpell_Change()
    Item(EItemNum).Spell = scrlLearnSpell.Value
    lblLearnSpell.Caption = "Spell: " & scrlLearnSpell.Value
End Sub

Private Sub scrlPicNum_Change()

    Item(EItemNum).Picture = scrlPicNum.Value
    
    picItem.Picture = Nothing
    If FileExist(App.Path & "\graphics\items\" & Item(EItemNum).Picture & ".bmp") = True Then
        picItem.Picture = LoadPicture(App.Path & "\graphics\items\" & Item(EItemNum).Picture & ".bmp")
    End If
    
    lblPicNum.Caption = "Num: " & scrlPicNum.Value

End Sub

Private Sub scrlPImage_Change()
    Item(EItemNum).Projectile.Image = scrlPImage.Value
    lblPImage.Caption = "Projectile Image: " & scrlPImage.Value
End Sub

Private Sub scrlPRange_Change()
    Item(EItemNum).Projectile.Range = scrlPRange.Value
    lblPRange.Caption = "Projectile Range: " & scrlPRange.Value
End Sub

Private Sub scrlPSpeed_Change()
    Item(EItemNum).Projectile.Speed = scrlPSpeed.Value / 10
    lblPSpeed.Caption = "Projectile Speed: " & scrlPSpeed.Value / 10
End Sub

Private Sub scrlStance_Change(Index As Integer)
Dim text As String

    Select Case Index
        Case Stance.MNorm
            text = "Male Normal: "
        Case Stance.MShield
            text = "Male Shield: "
        Case Stance.MTwoHand
            text = "Male Two Hand: "
        Case Stance.FNorm
            text = "Female Norm: "
        Case Stance.FShield
            text = "Female Shield: "
        Case Stance.FTwoHand
            text = "Female Two Hand: "
    End Select
    
    lblStance(Index).Caption = text & scrlStance(Index).Value
    Item(EItemNum).Paperdoll(Index) = scrlStance(Index).Value
    
End Sub

Private Sub scrlStat_Change(Index As Integer)

    Item(EItemNum).Stat(Index) = scrlStat(Index).Value

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

Private Sub scrlStatReq_Change(Index As Integer)

    Item(EItemNum).StatReq(Index) = scrlStatReq(Index).Value

    Select Case Index
        Case Stats.Attack
            lblStatReq(Index).Caption = "Att: " & scrlStatReq(Index).Value
        Case Stats.Strength
            lblStatReq(Index).Caption = "Str: " & scrlStatReq(Index).Value
        Case Stats.Defense
            lblStatReq(Index).Caption = "Def: " & scrlStatReq(Index).Value
        Case Stats.Agility
            lblStatReq(Index).Caption = "Agi: " & scrlStatReq(Index).Value
        Case Stats.Sagacity
            lblStatReq(Index).Caption = "Sag: " & scrlStatReq(Index).Value
    End Select

End Sub

Private Sub txtDamage_Change()

    If Not IsNumeric(txtDamage.text) Then
        txtDamage.text = Item(EItemNum).Damage
    End If
    
    Item(EItemNum).Damage = txtDamage.text

End Sub

Private Sub txtDefense_Change(Index As Integer)

    If Not IsNumeric(txtDefense(Index).text) Then
        txtDefense(Index).text = Item(EItemNum).Defense(Index)
    End If
    
    Item(EItemNum).Defense(Index) = txtDefense(Index).text

End Sub

Private Sub txtInfo_Change()

    Item(EItemNum).info = txtInfo.text

End Sub

Private Sub txtName_Change()

    Item(EItemNum).name = txtName.text

End Sub

Private Sub txtOffense_Change(Index As Integer)

    If Not IsNumeric(txtOffense(Index).text) Then
        txtOffense(Index).text = Item(EItemNum).Offense(Index)
    End If
    
    Item(EItemNum).Offense(Index) = txtOffense(Index).text

End Sub

Private Sub txtPrice_Change()

    If IsNumeric(txtPrice.text) = False Then
        txtPrice.text = Item(EItemNum).Price
    End If
    
    Item(EItemNum).Price = Trim$(txtPrice.text)

End Sub

Private Sub txtSkillMake_Change(Index As Integer)

    If Not IsNumeric(txtSkillReq(Index).text) Then
        txtSkillMake(Index).text = Item(EItemNum).ReqXP(Index)
    End If
    
    If txtSkillMake(Index).text > 100 Then
        txtSkillMake(Index).text = "100"
    End If
    
    Item(EItemNum).ReqXP(Index) = txtSkillMake(Index).text

End Sub

Private Sub txtSkillReq_Change(Index As Integer)

    If Not IsNumeric(txtSkillReq(Index).text) Then
        txtSkillReq(Index).text = Item(EItemNum).WReqXP(Index)
    End If
    
    If txtSkillReq(Index).text > 100 Then
        txtSkillReq(Index).text = "100"
    End If
    
    Item(EItemNum).WReqXP(Index) = txtSkillReq(Index).text

End Sub

Private Sub txtSkillReward_Change(Index As Integer)

    If Not IsNumeric(txtSkillReward(Index).text) Then
        txtSkillReward(Index).text = Item(EItemNum).RewXP(Index)
    End If
    
    Item(EItemNum).RewXP(Index) = txtSkillReward(Index).text

End Sub

Private Sub txtSpeed_Change()

    If Not IsNumeric(txtSpeed.text) Then
        txtSpeed.text = Item(EItemNum).Speed
    End If
    
    Item(EItemNum).Speed = txtSpeed.text

End Sub

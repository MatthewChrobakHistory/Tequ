VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000003&
   Caption         =   "Tequ"
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11760
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   551
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   784
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer tmrError 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   10440
      Top             =   1440
   End
   Begin VB.PictureBox picSpells 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3705
      Left            =   8115
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   70
      Top             =   3930
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox picSkills 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3705
      Left            =   8115
      Picture         =   "frmMain.frx":014A
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   52
      Top             =   3930
      Visible         =   0   'False
      Width           =   2895
      Begin VB.Label lblMax 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   8
         Left            =   2430
         TabIndex        =   69
         Top             =   3225
         Width           =   375
      End
      Begin VB.Label lblSkillLevel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   8
         Left            =   2100
         TabIndex        =   68
         Top             =   2985
         Width           =   375
      End
      Begin VB.Label lblMax 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   5
         Left            =   2430
         TabIndex        =   67
         Top             =   2310
         Width           =   375
      End
      Begin VB.Label lblSkillLevel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   5
         Left            =   2100
         TabIndex        =   66
         Top             =   2070
         Width           =   375
      End
      Begin VB.Label lblMax 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   4
         Left            =   2430
         TabIndex        =   65
         Top             =   1395
         Width           =   375
      End
      Begin VB.Label lblSkillLevel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   4
         Left            =   2100
         TabIndex        =   64
         Top             =   1155
         Width           =   375
      End
      Begin VB.Label lblMax 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   6
         Left            =   2430
         TabIndex        =   63
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblSkillLevel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   6
         Left            =   2100
         TabIndex        =   62
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblMax 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   7
         Left            =   1050
         TabIndex        =   61
         Top             =   3225
         Width           =   375
      End
      Begin VB.Label lblSkillLevel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   7
         Left            =   720
         TabIndex        =   60
         Top             =   2985
         Width           =   375
      End
      Begin VB.Label lblMax 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   3
         Left            =   1050
         TabIndex        =   59
         Top             =   2310
         Width           =   375
      End
      Begin VB.Label lblSkillLevel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   58
         Top             =   2070
         Width           =   375
      End
      Begin VB.Label lblMax 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   1050
         TabIndex        =   57
         Top             =   1395
         Width           =   375
      End
      Begin VB.Label lblSkillLevel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   56
         Top             =   1155
         Width           =   375
      End
      Begin VB.Label lblMax 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   1050
         TabIndex        =   55
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblSkillLevel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   54
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.PictureBox picChest 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   1350
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   51
      Top             =   2070
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.PictureBox picShop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5115
      Left            =   1755
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   49
      Top             =   480
      Visible         =   0   'False
      Width           =   4125
      Begin VB.PictureBox picShopItems 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3330
         Left            =   615
         ScaleHeight     =   222
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   50
         Top             =   630
         Width           =   2895
      End
      Begin VB.Image imgShopSell 
         Height          =   435
         Left            =   2475
         Top             =   4350
         Width           =   1035
      End
      Begin VB.Image imgShopBuy 
         Height          =   435
         Left            =   615
         Top             =   4350
         Width           =   1035
      End
   End
   Begin VB.PictureBox picBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   150
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   48
      Top             =   150
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1215
      Left            =   7440
      Picture         =   "frmMain.frx":23128
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   26
      Top             =   1680
      Visible         =   0   'False
      Width           =   4200
      Begin VB.Label lblPrice 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   29
         Top             =   180
         Width           =   1935
      End
      Begin VB.Label lblInfoText 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   285
         TabIndex        =   28
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label lblInfoName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   180
         Width           =   1335
      End
   End
   Begin VB.PictureBox picCharacter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3705
      Left            =   8115
      Picture         =   "frmMain.frx":33B32
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   13
      Top             =   3930
      Visible         =   0   'False
      Width           =   2895
      Begin VB.Label lblPoints 
         BackStyle       =   0  'Transparent
         Caption         =   "Points: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   1560
         TabIndex        =   25
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   210
         Width           =   2895
      End
      Begin VB.Label lblStat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   18
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblStat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   17
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblStat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   16
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblStat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   15
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblStat 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   14
         Top             =   2760
         Width           =   1215
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   6
      Left            =   10860
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   12
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   5
      Left            =   10245
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   11
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   4
      Left            =   9630
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   10
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   9015
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   9
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   8400
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   7785
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   7
      Top             =   3000
      Width           =   480
   End
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3705
      Left            =   8115
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   4
      Top             =   3930
      Visible         =   0   'False
      Width           =   2895
      Begin VB.HScrollBar scrlSFX 
         Height          =   255
         Left            =   480
         Max             =   10
         Min             =   1
         TabIndex        =   21
         Top             =   1440
         Value           =   7
         Width           =   1935
      End
      Begin VB.HScrollBar scrlVolume 
         Height          =   255
         Left            =   480
         Max             =   10
         Min             =   1
         TabIndex        =   6
         Top             =   720
         Value           =   7
         Width           =   1935
      End
      Begin VB.Label lblSmallScreen 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SmallScreen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   270
         Left            =   1560
         TabIndex        =   47
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblFullScreen 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fullscreen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   270
         Left            =   360
         TabIndex        =   46
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label chkDIOC 
         BackStyle       =   0  'Transparent
         Caption         =   "Display info on click?: False"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   480
         TabIndex        =   30
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label lblSound 
         BackStyle       =   0  'Transparent
         Caption         =   "Sound: False at 0.0 Volume"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblMusic 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Music: False at 0.0 Volume"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.PictureBox picInventory 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3705
      Left            =   8115
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   3
      Top             =   3930
      Width           =   2895
   End
   Begin VB.TextBox txtMyChat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   180
      MaxLength       =   90
      TabIndex        =   1
      Top             =   7815
      Visible         =   0   'False
      Width           =   7140
   End
   Begin MSWinsockLib.Winsock socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   5760
      Left            =   150
      ScaleHeight     =   384
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   150
      Width           =   7200
      Begin VB.Frame fraPlayerCreate 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   990
         TabIndex        =   31
         Top             =   3480
         Visible         =   0   'False
         Width           =   5175
         Begin VB.ComboBox cmbLegs 
            Height          =   315
            ItemData        =   "frmMain.frx":56B10
            Left            =   3840
            List            =   "frmMain.frx":56B17
            TabIndex        =   45
            Text            =   "Normal"
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cmbBody 
            Height          =   315
            ItemData        =   "frmMain.frx":56B23
            Left            =   2640
            List            =   "frmMain.frx":56B2A
            TabIndex        =   44
            Text            =   "Normal"
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cmbHair 
            Height          =   315
            ItemData        =   "frmMain.frx":56B36
            Left            =   1440
            List            =   "frmMain.frx":56B3D
            TabIndex        =   43
            Text            =   "Normal"
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton optGender 
            BackColor       =   &H80000012&
            Caption         =   "Female"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   375
            Index           =   2
            Left            =   3480
            TabIndex        =   41
            Top             =   1560
            Width           =   945
         End
         Begin VB.OptionButton optGender 
            BackColor       =   &H80000012&
            Caption         =   "Male"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   375
            Index           =   1
            Left            =   1080
            TabIndex        =   40
            Top             =   1560
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.HScrollBar scrlLegs 
            Height          =   255
            Left            =   3840
            Max             =   4
            Min             =   1
            TabIndex        =   39
            Top             =   1080
            Value           =   1
            Width           =   1095
         End
         Begin VB.HScrollBar scrlBody 
            Height          =   255
            Left            =   2640
            Max             =   4
            Min             =   1
            TabIndex        =   38
            Top             =   1080
            Value           =   1
            Width           =   1095
         End
         Begin VB.HScrollBar scrlHair 
            Height          =   255
            Left            =   1440
            Max             =   4
            Min             =   1
            TabIndex        =   37
            Top             =   1080
            Value           =   1
            Width           =   1095
         End
         Begin VB.HScrollBar scrlSkin 
            Height          =   255
            Left            =   240
            Max             =   4
            Min             =   1
            TabIndex        =   36
            Top             =   1080
            Value           =   1
            Width           =   1095
         End
         Begin VB.Label lblDone 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "[ DONE ]"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   1920
            TabIndex        =   42
            Top             =   1600
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Legs"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   375
            Left            =   3840
            TabIndex        =   35
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Body"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   2640
            TabIndex        =   34
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Hair"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   1440
            TabIndex        =   33
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Skin"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000B&
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   1800
      Left            =   180
      TabIndex        =   2
      Top             =   6030
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   3175
      _Version        =   393217
      BackColor       =   790032
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":56B49
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblError 
      BackStyle       =   0  'Transparent
      Caption         =   "An error occured. Click here to report it!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7440
      TabIndex        =   71
      Top             =   -240
      Width           =   4335
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   8040
      TabIndex        =   53
      Top             =   7680
      Width           =   3015
   End
   Begin VB.Label lblXP 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   7875
      TabIndex        =   24
      Top             =   1110
      Width           =   3375
   End
   Begin VB.Label lblMP 
      BackStyle       =   0  'Transparent
      Caption         =   "TEST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   7875
      TabIndex        =   23
      Top             =   780
      Width           =   3375
   End
   Begin VB.Label lblHp 
      BackStyle       =   0  'Transparent
      Caption         =   "TEST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   7875
      TabIndex        =   22
      Top             =   435
      Width           =   3375
   End
   Begin VB.Image imgXp 
      Height          =   240
      Left            =   7770
      Picture         =   "frmMain.frx":56BC4
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Image imgSp 
      Height          =   240
      Left            =   7770
      Picture         =   "frmMain.frx":5A09A
      Top             =   750
      Width           =   3615
   End
   Begin VB.Image imgHp 
      Height          =   240
      Left            =   7770
      Picture         =   "frmMain.frx":5D777
      Top             =   420
      Width           =   3615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MX As Long
Private MY As Long

Private Sub chkDIOC_Click()

    If Options.DIOC = False Then
        Options.DIOC = True
        chkDIOC.Caption = "Display info on click?: True"
    Else
        Options.DIOC = False
        picInfo.Visible = False
        chkDIOC.Caption = "Display info on click?: False"
    End If

End Sub

Private Sub cmbBody_Click()

    Select Case cmbBody.ListIndex
        Case 0
            Player(MyIndex).Graphics.BodyDir = "norm"
    End Select
    
End Sub

Private Sub cmbHair_Click()

    Select Case cmbHair.ListIndex
        Case 0
            Player(MyIndex).Graphics.HairDir = "norm"
    End Select

End Sub

Private Sub cmbLegs_Click()

    Select Case cmbLegs.ListIndex
        Case 0
            Player(MyIndex).Graphics.LegsDir = "norm"
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DestroyGame
End Sub

Private Sub imgShopBuy_Click()

    If Game.ShopState = SHOP_STATE_BUY Then Exit Sub
    Game.ShopState = SHOP_STATE_BUY
    Call AddText("Click on an item you wish to buy.", BrightGreen)

End Sub

Private Sub imgShopSell_Click()

    If Game.ShopState = SHOP_STATE_SELL Then Exit Sub
    Game.ShopState = SHOP_STATE_SELL
    Call AddText("Click on an item in your inventory to sell.", BrightGreen)

End Sub
Private Sub lblDone_Click()

    CreatingCharacter = False
    fraPlayerCreate.Visible = False

End Sub

Private Sub lblError_Click()
    Call InitDebugGUI
    frmDebugger.fraReport.Visible = True
    With frmDebugger
        .txtDescription = LastExceptionDescription
        .txtErrorType = LastExceptionNum
        .txtLine = LastExceptionLine
        .txtSource = LastExceptionSource
        .txtDescription.Enabled = False
        .txtErrorType.Enabled = False
        .txtLine.Enabled = False
        .txtSource.Enabled = False
    End With
End Sub

Private Sub lblFullScreen_Click()
    
    'Call AddText("Temporarily disbaled...", BrightRed)
    'Exit Sub
    
    If Screen.Height < 11520 Or Screen.Width < 20490 Then
        Call AddText("Your screen size is too small. Sorry.", BrightRed)
        Exit Sub
    End If
    
    Options.FullScreen = True
    Call SetupGUI

End Sub

Private Sub lblInfoName_Click()

    picInfo.Visible = False

End Sub

Private Sub lblInfoText_Click()

    picInfo.Visible = False

End Sub

Private Sub lblMusic_Click()

    If Options.Music = True Then
        Options.Music = False
        Call StopMusic
    Else
        Options.Music = True
        Call StopMusic
        If Map(Player(MyIndex).Map).Music <> vbNullString Then
            Call PlayMusic(Map(Player(MyIndex).Map).Music)
        End If
    End If
    
    lblMusic.Caption = "Music: " & Options.Music & " at " & scrlVolume.Value / 10 & " Volume"
    
End Sub

Private Sub lblPrice_Click()

    picInfo.Visible = False

End Sub

Private Sub lblSmallScreen_Click()

    Options.FullScreen = False
    Call SetupGUI
    
End Sub

Private Sub lblSound_Click()

    If Options.Sound = True Then
        Options.Sound = False
    Else
        Options.Sound = True
    End If
    
    lblSound.Caption = "Sound: " & Options.Sound & " at " & scrlSFX.Value / 10 & " Volume"
    
End Sub

Private Sub lblStat_Click(Index As Integer)

    Call TrainStat(MyIndex, Index)

End Sub

Private Sub optGender_Click(Index As Integer)

    Player(MyIndex).Graphics.Gender = Index
    Select Case Player(MyIndex).Graphics.Gender
        Case GENDER_MALE
            Player(MyIndex).Stance = Stance.MNorm
        Case GENDER_FEMALE
            Player(MyIndex).Stance = Stance.FNorm
    End Select

End Sub

Private Sub picBank_DblClick()
Dim BankSlot
    
    BankSlot = IsBankItem(BankX, BankY)
    
    If BankSlot = 0 Then Exit Sub
    
    If Bank(MyIndex).BankTab(CurTab).BankItem(BankSlot).Num = 0 Then Exit Sub

    If BankSlot > 0 Then
        Call WithdrawBankItem(BankSlot)
    End If

End Sub

Private Sub picBank_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim bankNum As Long
Dim tabNum As Long

    bankNum = IsBankItem(X, Y)
    
    If bankNum <> 0 Then
        If Button = 1 Then DragBankSlotNum = bankNum
        Exit Sub
    End If
    
    tabNum = IsTabNum(X, Y)
    If tabNum <> 0 Then
        CurTab = tabNum
        Call RenderBank
    End If
    
End Sub

Private Sub picBank_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    BankX = X
    BankY = Y

End Sub

Private Sub picBank_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim recPos As RECT
Dim OldItemNum As Long, NewItemNum As Long
Dim OldItemValue As Long, NewItemValue As Long
Dim OldTab As Byte

    If DragBankSlotNum > 0 Then
        For i = 1 To MAX_BANK_ITEMS
            With recPos
                .top = 38 + (35 * ((i - 1) \ 11))
                .Bottom = .top + 32
                .Left = 42 + (36 * (((i - 1) Mod 11)))
                .Right = .Left + 32
            End With

            If X >= recPos.Left And X <= recPos.Right Then
                If Y >= recPos.top And Y <= recPos.Bottom Then
                    If DragBankSlotNum <> i Then
                        With Bank(MyIndex).BankTab(CurTab).BankItem(i)
                            OldItemNum = .Num
                            OldItemValue = .Value
                        End With
                        With Bank(MyIndex).BankTab(CurTab).BankItem(DragBankSlotNum)
                            NewItemNum = .Num
                            NewItemValue = .Value
                        End With
                        Bank(MyIndex).BankTab(CurTab).BankItem(i).Num = NewItemNum
                        Bank(MyIndex).BankTab(CurTab).BankItem(i).Value = NewItemValue
                        Bank(MyIndex).BankTab(CurTab).BankItem(DragBankSlotNum).Num = OldItemNum
                        Bank(MyIndex).BankTab(CurTab).BankItem(DragBankSlotNum).Value = OldItemValue
                        Call RenderBank
                        Exit For
                    End If
                End If
            End If
        Next
        
        If IsTabNum(X, Y) Then
            If FindBankSlot(Bank(MyIndex).BankTab(IsTabNum(X, Y)).BankItem(DragBankSlotNum).Num) > 0 Then
                OldTab = CurTab
                With Bank(MyIndex).BankTab(CurTab).BankItem(DragBankSlotNum)
                    OldItemNum = .Num
                    OldItemValue = .Value
                    .Num = 0
                    .Value = 0
                End With
                CurTab = IsTabNum(X, Y)
                With Bank(MyIndex).BankTab(CurTab).BankItem(FindBankSlot(OldItemNum))
                    .Num = OldItemNum
                    .Value = .Value + OldItemValue
                End With
                CurTab = OldTab
                Call RenderBank
            End If
        End If
    End If

    DragBankSlotNum = 0
End Sub

Private Sub picCharacter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If IsEquipmentItem(X, Y) > 0 Then
        Call UnequipItem(IsEquipmentItem(X, Y))
    End If

End Sub

Private Sub picChest_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ChestItem As Byte
Dim Itemnum As Long, ItemValue As Long
Dim i As Long

    If Button = 1 Then
        ChestItem = IsChestItem(X, Y)
        If ChestItem > 0 Then
            Itemnum = MapChest(Player(MyIndex).Map).Tile(Game.ChestX, Game.ChestY).ChestItem(ChestItem).Itemnum
            ItemValue = MapChest(Player(MyIndex).Map).Tile(Game.ChestX, Game.ChestY).ChestItem(ChestItem).ItemValue
            If Itemnum > 0 Then
                If Item(Itemnum).Stackable = True Then
                    If FindOpenInvSlot(Itemnum) > 0 Then
                        Call GivePlayerItem(Itemnum, ItemValue)
                        MapChest(Player(MyIndex).Map).Tile(Game.ChestX, Game.ChestY).ChestItem(ChestItem).Itemnum = 0
                        MapChest(Player(MyIndex).Map).Tile(Game.ChestX, Game.ChestY).ChestItem(ChestItem).ItemValue = 0
                    End If
                Else
                    For i = 1 To ItemValue
                        If FindOpenInvSlot(Itemnum) > 0 Then
                            Call GivePlayerItem(Itemnum, 1)
                            MapChest(Player(MyIndex).Map).Tile(Game.ChestX, Game.ChestY).ChestItem(ChestItem).ItemValue = MapChest(Player(MyIndex).Map).Tile(Game.ChestX, Game.ChestY).ChestItem(ChestItem).ItemValue - 1
                            If i = ItemValue Then
                                MapChest(Player(MyIndex).Map).Tile(Game.ChestX, Game.ChestY).ChestItem(ChestItem).Itemnum = 0
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
    End If
    
    Call RenderChest(Map(Player(MyIndex).Map).Tile(Game.ChestX, Game.ChestY).LongValue(1), Game.ChestX, Game.ChestY)
    Call BltInventory
    
End Sub

Private Sub picInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

picInfo.Visible = False

Exit Sub

' clicking on the image
If Y <= 32 Then
    If X > 10 And X < 42 Then
    End If
End If

End Sub

Private Sub picInventory_DblClick()
Dim InvNum As Long

    InvNum = IsInvItem(InvX, InvY)

    If Game.InShop = True Then
        If Game.ShopState = SHOP_STATE_SELL Then
            If FindOpenInvSlot(1) Then
                If InvNum > 0 Then
                    If Player(MyIndex).Inv(InvNum).Num = 0 Then Exit Sub
                    Call GivePlayerItem(1, Item(Player(MyIndex).Inv(InvNum).Num).Price)
                    Call TakeInvItem(MyIndex, InvNum, 1)
                End If
            End If
        Else
            Exit Sub
        End If
    End If

    If InvNum > 0 And InvNum <= MAX_INV Then
        If Player(MyIndex).Inv(InvNum).Num > 0 Then
            Call UseItem(InvNum)
            picInfo.Visible = False
        End If
    End If

End Sub

Private Sub picInventory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim InvNum As Long

    InvNum = IsInvItem(X, Y)
    
    If Game.InShop = True Then
        If Button = 2 Then
            If InvNum <> 0 Then
                If Player(MyIndex).Inv(InvNum).Num > 0 Then
                    Call RenderInfo(Player(MyIndex).Inv(InvNum).Num)
                End If
            End If
        End If
        Exit Sub
    End If

    If Button = 1 Then
        If InvNum <> 0 Then
            If Player(MyIndex).Inv(InvNum).Num > 0 Then
                DragInvSlotNum = InvNum
                If Item(Player(MyIndex).Inv(InvNum).Num).Picture > 0 Then
                    Call RenderInfo(Player(MyIndex).Inv(InvNum).Num)
                End If
            End If
        End If
    ElseIf Button = 2 Then
        If InvNum <> 0 Then
            If Options.OnlineMode = True Then
                Call SendDropItem(InvNum)
            Else
                Call DropItem(InvNum)
            End If
        End If
    End If

End Sub

Private Sub picInventory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    InvX = X
    InvY = Y
    
    Exit Sub
    
    If IsInvItem(X, Y) Then
        If Player(MyIndex).Inv(IsInvItem(X, Y)).Num > 0 Then
            Call RenderInfo(Player(MyIndex).Inv(IsInvItem(X, Y)).Num)
            Exit Sub
        End If
    End If
    
    picInfo.Visible = False

End Sub

Private Sub picInventory_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim InvNum As Long
Dim Num(1 To 2) As Long, Amount(1 To 2) As Long

    InvNum = IsInvItem(X, Y)
    
    If DragInvSlotNum > 0 Then
        If InvNum <> DragInvSlotNum Then
            If InvNum = 0 Then Exit Sub
            
            Num(1) = Player(MyIndex).Inv(DragInvSlotNum).Num
            Num(2) = Player(MyIndex).Inv(InvNum).Num
            Amount(1) = Player(MyIndex).Inv(DragInvSlotNum).Value
            Amount(2) = Player(MyIndex).Inv(InvNum).Value
            Player(MyIndex).Inv(DragInvSlotNum).Num = Num(2)
            Player(MyIndex).Inv(DragInvSlotNum).Value = Amount(2)
            Player(MyIndex).Inv(InvNum).Num = Num(1)
            Player(MyIndex).Inv(InvNum).Value = Amount(1)
            Call BltInventory
        End If
    End If
    
    InvNum = 0
    DragInvSlotNum = 0

End Sub

Private Sub picPlayer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If IsEquipmentItem(X, Y) > 0 Then
        Call UnequipItem(IsEquipmentItem(X, Y))
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Call HandleKeyPresses(KeyCode)

    ' prevents textbox on error ding sound
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
        KeyCode = 0
    End If

End Sub

Private Sub picScreen_KeyDown(KeyCode As Integer, Shift As Integer)
    Call HandleKeyPresses(KeyCode)

    ' prevents textbox on error ding sound
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
        KeyCode = 0
    End If
End Sub

Private Sub picScreen_KeyPress(KeyAscii As Integer)
    Call HandleKeyPresses(KeyAscii)

    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        KeyAscii = 0
    End If
End Sub

Private Sub picScreen_KeyUp(KeyCode As Integer, Shift As Integer)

    Call HandleKeyReleases(KeyCode)

End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MouseX As Single, MouseY As Single
Dim ConvX As Byte, ConvY As Byte
Dim i As Long

    If frmEditor_Map.Visible = True Then
        If Button = vbRightButton Then
            Call EditMap(X, Y, True)
            Exit Sub
        Else
            Call EditMap(X, Y, False)
            Exit Sub
        End If
    End If
    
    If Button = vbRightButton Then
        If Shift = 1 Then
            If Player(MyIndex).Access = ACCESS_ADMIN Then
            Dim NewX As Long
            Dim NewY As Long
            NewX = X / 32
            NewY = Y / 32
            If NewX > MAX_MAP_X Then NewX = MAX_MAP_X
            If NewY > MAX_MAP_Y Then NewY = MAX_MAP_Y
            If NewX < 1 Then NewX = 1
            If NewY < 1 Then NewY = 1
            Player(MyIndex).X = NewX
            Player(MyIndex).Y = NewY
            End If
        End If
    ElseIf Button = vbLeftButton Then
        
    MouseX = X / 32
    MouseY = Y / 32
    If MouseX < 0 Or MouseX > 480 / 32 Then Exit Sub
    If MouseY < 0 Or MouseY > 384 / 32 Then Exit Sub
    ConvX = MouseX
    ConvY = MouseY
    ' If the rounded number is bigger than the original number, we must have rounded up. Deduct one
    If ConvX - MouseX > 0 Then ConvX = ConvX - 1
    If ConvY - MouseY > 0 Then ConvY = ConvY - 1
    ConvX = ConvX + 1
    ConvY = ConvY + 1
    
    For i = 1 To MAX_MAP_NPCS
        If TempNpc(Player(MyIndex).Map).NpcNum(i).X = ConvX And TempNpc(Player(MyIndex).Map).NpcNum(i).Y = ConvY Then
            If TempPlayer(MyIndex).Target = i Then
                TempPlayer(MyIndex).Target = 0
            Else
                TempPlayer(MyIndex).Target = i
            End If
        End If
    Next
    
    End If
    
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MouseX As Single, MouseY As Single
Dim ConvX As Byte, ConvY As Byte
Dim i As Long

    CurX = X
    CurY = Y

    If Button = vbLeftButton And frmEditor_Map.Visible = True Then
        Call EditMap(X, Y, False)
        Exit Sub
    ElseIf Button = vbRightButton And frmEditor_Map.Visible = True Then
        Call EditMap(X, Y, True)
        Exit Sub
    End If
    
    MouseX = X / 32
    MouseY = Y / 32
    If MouseX < 0 Or MouseX > 480 / 32 Then Exit Sub
    If MouseY < 0 Or MouseY > 384 / 32 Then Exit Sub
    ConvX = MouseX
    ConvY = MouseY
    ' If the rounded number is bigger than the original number, we must have rounded up. Deduct one
    If ConvX - MouseX > 0 Then ConvX = ConvX - 1
    If ConvY - MouseY > 0 Then ConvY = ConvY - 1
    ConvX = ConvX + 1
    ConvY = ConvY + 1
    
    For i = 1 To MAX_MAP_NPCS
        If Map(Player(MyIndex).Map).MapNpc(i).Num > 0 Then
            If TempNpc(Player(MyIndex).Map).NpcNum(i).X = ConvX And TempNpc(Player(MyIndex).Map).NpcNum(i).Y = ConvY Then
                TempPlayer(MyIndex).HoverTarget = i
                Exit For
            Else
                If i = MAX_MAP_NPCS Then
                    TempPlayer(MyIndex).HoverTarget = 0
                    Exit For
                End If
            End If
        Else
            If i = MAX_MAP_NPCS Then
                TempPlayer(MyIndex).HoverTarget = 0
                Exit For
            End If
        End If
    Next

End Sub

Private Sub picShopItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim text As String

    If Button = 1 Then
        If Game.ShopState <> SHOP_STATE_BUY Then Exit Sub
        If IsShopItem(X, Y) = 0 Then Exit Sub
        Call BuyItem(IsShopItem(X, Y))
    Else
        If IsShopItem(X, Y) > 0 Then
            If Shop(Game.ShopNum).ShopItem(IsShopItem(X, Y)).StockItem > 0 Then
                Call RenderInfo(Shop(Game.ShopNum).ShopItem(IsShopItem(X, Y)).StockItem)
                With Shop(Game.ShopNum).ShopItem(IsShopItem(X, Y))
                    If .NumberofCosts = 0 Then
                        Call AddText("This item is free.", BrightGreen)
                    Else
                        text = "You need " & .ItemCost(1).ItemCostValue & " " & Trim$(Item(.ItemCost(1).ItemCostNum).name)
                        For i = 1 To .NumberofCosts
                            If i <> 1 Then text = text + "and " & .ItemCost(i).ItemCostValue & " " & Trim$(Item(.ItemCost(i).ItemCostNum).name) & " "
                            If i = .NumberofCosts Then text = text & " to " & Trim$(.Verb) & " this item."
                        Next
                        Call AddText(text, BrightGreen)
                    End If
                End With
            End If
        End If
    End If

End Sub

Private Sub picSkills_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim XP As String

    If X > 97 Then
        X = 2
    ElseIf X <= 96 Then
        X = 1
    End If
    
    If Y >= 184 Then
        Y = 4
    ElseIf Y >= 124 Then
        Y = 3
    ElseIf Y >= 64 Then
        Y = 2
    ElseIf Y < 44 Or Y >= 44 Then
        Y = 1
    End If
    
    Select Case X
        Case 1
            Select Case Y
                Case 1
                    XP = Format$(Player(MyIndex).Skill(Skills.Woodcutting).XP, "#,###,###,###") & " / " & Format$(GetPlayerNextLevelXP(Player(MyIndex).Skill(Skills.Woodcutting).level), "#,###,###,###")
                Case 2
                    XP = Format$(Player(MyIndex).Skill(Skills.Mining).XP, "#,###,###,###") & " / " & Format$(GetPlayerNextLevelXP(Player(MyIndex).Skill(Skills.Mining).level), "#,###,###,###")
                Case 3
                    XP = Format$(Player(MyIndex).Skill(Skills.Fishing).XP, "#,###,###,###") & " / " & Format$(GetPlayerNextLevelXP(Player(MyIndex).Skill(Skills.Fishing).level), "#,###,###,###")
                Case 4
                    XP = Format$(Player(MyIndex).Skill(Skills.Crafting).XP, "#,###,###,###") & " / " & Format$(GetPlayerNextLevelXP(Player(MyIndex).Skill(Skills.Crafting).level), "#,###,###,###")
            End Select
        Case 2
            Select Case Y
                Case 1
                    XP = Format$(Player(MyIndex).Skill(Skills.Fletching).XP, "#,###,###,###") & " / " & Format$(GetPlayerNextLevelXP(Player(MyIndex).Skill(Skills.Fletching).level), "#,###,###,###")
                Case 2
                    XP = Format$(Player(MyIndex).Skill(Skills.Smithing).XP, "#,###,###,###") & " / " & Format$(GetPlayerNextLevelXP(Player(MyIndex).Skill(Skills.Smithing).level), "#,###,###,###")
                Case 3
                    XP = Format$(Player(MyIndex).Skill(Skills.Cooking).XP, "#,###,###,###") & " / " & Format$(GetPlayerNextLevelXP(Player(MyIndex).Skill(Skills.Cooking).level), "#,###,###,###")
                Case 4
                    XP = Format$(Player(MyIndex).Skill(Skills.PotionBrewing).XP, "#,###,###,###") & " / " & Format$(GetPlayerNextLevelXP(Player(MyIndex).Skill(Skills.PotionBrewing).level), "#,###,###,###")
            End Select
    End Select
    
    lblInfo.Caption = XP
    
End Sub

Private Sub picSpells_DblClick()
    If IsSpell(MX, MY) Then Call CastSpell(MyIndex, IsSpell(MX, MY))
End Sub

Private Sub picSpells_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If IsSpell(X, Y) Then
        Select Case Button
            Case vbLeftButton ' Info
            
            Case vbRightButton ' forgetting
                Call ForgetSpell(IsSpell(X, Y))
        End Select
    End If
End Sub

Private Sub picSpells_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Setting variables for double clicking on picSpells
    MY = Y
    MX = X
End Sub

Private Sub picTab_Click(Index As Integer)

If picInventory.Visible = True Then
    If Index = Tabs.Inventory Then
        picInventory.Visible = False
        Exit Sub
    End If
End If

If picOptions.Visible = True Then
    If Index = Tabs.MyOptions Then
        picOptions.Visible = False
        Exit Sub
    End If
End If

If picCharacter.Visible = True Then
    If Index = Tabs.Character Then
        picCharacter.Visible = False
        Exit Sub
    End If
End If

If picSkills.Visible = True Then
    If Index = Tabs.Skills Then
        picSkills.Visible = False
        Exit Sub
    End If
End If

If picSpells.Visible = True Then
    If Index = Tabs.Spells Then
        picSpells.Visible = False
        Exit Sub
    End If
End If


picInventory.Visible = False
picOptions.Visible = False
picCharacter.Visible = False
picSkills.Visible = False
picSpells.Visible = False

    Select Case Index
        Case Tabs.Inventory
            picInventory.Visible = True
            Call BltInventory
        Case Tabs.MyOptions
            picOptions.Visible = True
        Case Tabs.Character
            picCharacter.Visible = True
            Call TabCharacterInit
        Case Tabs.Spells
            picSpells.Visible = True
            Call BltSpells
        Case Tabs.Skills
            picSkills.Visible = True
            Call TabSkillsInit
            lblInfo.Visible = True
            lblInfo.Caption = vbNullString
        Case Tabs.Quit
            Call DestroyGame
    End Select

End Sub

Public Sub TabSkillsInit()
Dim i As Byte

    Me.lblInfo.Caption = vbNullString
    
    For i = 1 To Skills.Skill_Count - 1
        lblSkillLevel(i).Caption = Player(MyIndex).Skill(i).level
    Next
    
End Sub

Public Sub TabCharacterInit()
    lblName.Caption = Trim$(Player(MyIndex).name) & ": Level " & Player(MyIndex).Combat.level
    lblStat(Stats.Attack).Caption = "Attack: " & Player(MyIndex).Stat(Stats.Attack)
    lblStat(Stats.Strength).Caption = "Strength: " & Player(MyIndex).Stat(Stats.Strength)
    lblStat(Stats.Defense).Caption = "Defense: " & Player(MyIndex).Stat(Stats.Defense)
    lblStat(Stats.Agility).Caption = "Agility: " & Player(MyIndex).Stat(Stats.Agility)
    lblStat(Stats.Sagacity).Caption = "Sagacity: " & Player(MyIndex).Stat(Stats.Sagacity)
    lblPoints.Caption = "Points: " & Player(MyIndex).Points
End Sub
Private Sub scrlBody_Change()

    Player(MyIndex).Graphics.Body = scrlBody.Value

End Sub

Private Sub scrlHair_Change()

    Player(MyIndex).Graphics.Hair = scrlHair.Value

End Sub

Private Sub scrlLegs_Change()

    Player(MyIndex).Graphics.Legs = scrlLegs.Value

End Sub

Private Sub scrlSFX_Change()

    Call SetVolume(SoundIndex, scrlSFX.Value / 10)
    
    lblSound.Caption = "Music: " & Options.Sound & " at " & scrlSFX.Value / 10 & " Volume"
    
End Sub

Private Sub scrlSkin_Change()

    Player(MyIndex).Graphics.Skin = scrlSkin.Value

End Sub

Private Sub scrlVolume_Change()

    Call SetVolume(MusicIndex, scrlVolume.Value / 10)
    
    lblMusic.Caption = "Music: " & Options.Music & " at " & scrlVolume.Value / 10 & " Volume"

End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

End Sub

Private Sub tmrError_Timer()

tickError = tickError + 0.5

Select Case tickError
    Case Is <= 16
        lblError.top = lblError.top + 1
    Case Is <= 31
    Case Is < 50
        lblError.top = lblError.top - 1
    Case 50
        tmrError.Enabled = False
        lblError.top = -16
        tickError = 0
End Select

End Sub

Private Sub txtMyChat_KeyDown(KeyCode As Integer, Shift As Integer)

    Call HandleKeyPresses(KeyCode)

End Sub

Public Function IsInvItem(ByVal X As Single, ByVal Y As Single) As Long
Dim TempRec As RECT
Dim i As Long

    IsInvItem = 0

    For i = 1 To MAX_INV
        'If Player(MyIndex).Inv(i).Num > 0 Then
        
            With TempRec
                .top = 1 + (35 * ((i - 1) \ 5))
                .Bottom = .top + 32
                .Left = 12 + (35 * (((i - 1) Mod 5)))
                .Right = .Left + 32
            End With

            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.top And Y <= TempRec.Bottom Then
                    IsInvItem = i
                    Exit Function
                End If
            End If
        'End If
    Next

End Function

Public Function IsSpell(ByVal X As Single, ByVal Y As Single) As Long
Dim TempRec As RECT
Dim i As Long

    IsSpell = 0
    
    For i = 1 To MAX_PLAYER_SPELLS
        If Player(MyIndex).PlayerSpell(i).Num > 0 Then
            
            With TempRec
                .top = 1 + (35 * ((i - 1) \ 5))
                .Bottom = .top + 32
                .Left = 12 + (35 * ((i - 1) Mod 5))
                .Right = .Left + 32
            End With
            
            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.top And Y <= TempRec.Bottom Then
                    IsSpell = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function


Private Function IsEquipmentItem(ByVal X As Single, ByVal Y As Single) As Long
Dim TempRec As RECT
Dim i As Long
Dim LeftOffset As Single, TopOffset As Single

    IsEquipmentItem = 0
    
    For i = 1 To Equipment.Equipment_Count - 1
    
        TopOffset = i / 3
        If TopOffset <= 1 Then
            TopOffset = 48
        ElseIf TopOffset <= 2 Then
            TopOffset = 88
        ElseIf TopOffset <= 3 Then
            TopOffset = 128
        End If
        If i = 1 Or i = 4 Or i = 7 Then
            LeftOffset = 32
        ElseIf i = 2 Or i = 5 Or i = 8 Then
            LeftOffset = 80
        ElseIf i = 3 Or i = 6 Or i = 9 Then
            LeftOffset = 128
        End If
                    
        With TempRec
            .top = TopOffset
            .Bottom = .top + 32
            .Left = LeftOffset
            .Right = .Left + 32
        End With
        
        If X >= TempRec.Left And X <= TempRec.Right Then
            If Y >= TempRec.top And Y <= TempRec.Bottom Then
                IsEquipmentItem = i
                Exit Function
            End If
        End If
    Next
    
End Function

Public Function IsBankItem(ByVal X As Long, ByVal Y As Long) As Long
Dim TempRec As RECT
Dim i As Long

    IsBankItem = 0
    
    For i = 1 To MAX_BANK_ITEMS
        With TempRec
            .top = 38 + (35 * ((i - 1) \ 11))
            .Bottom = .top + 32
            .Left = 42 + (36 * (((i - 1) Mod 11)))
            .Right = .Left + 32
        End With
            
        If X >= TempRec.Left And X <= TempRec.Right Then
            If Y >= TempRec.top And Y <= TempRec.Bottom Then
                IsBankItem = i
                Exit Function
            End If
        End If
    Next
End Function

Public Function IsTabNum(ByVal X As Long, ByVal Y As Long) As Byte
Dim TempRec As RECT
Dim i As Byte

    IsTabNum = 0
    
    For i = 1 To MAX_BANK_TABS
        With TempRec
            .Left = 34 + (i * 8) + (i - 1) + ((i - 1) * 32) - 4
            .Right = .Left + 32
            .top = 7
            .Bottom = .top + 32
        End With
        
        If X >= TempRec.Left And X <= TempRec.Right Then
            If Y >= TempRec.top And Y <= TempRec.Bottom Then
                IsTabNum = i
                Exit Function
            End If
        End If
    Next
    
End Function

Public Function IsShopItem(ByVal X As Long, ByVal Y As Long) As Byte
Dim TempRec As RECT
Dim i As Byte

    IsShopItem = 0
    
    For i = 1 To MAX_SHOP_ITEMS
        With TempRec
            .top = 5 + ((4 + 32) * ((i - 1) \ 5))
            .Bottom = .top + 32
            .Left = 10 + ((4 + 32) * (((i - 1) Mod 5)))
            .Right = .Left + 32
        End With
        
        If X >= TempRec.Left And X <= TempRec.Right Then
            If Y >= TempRec.top And Y <= TempRec.Bottom Then
                IsShopItem = i
                Exit Function
            End If
        End If
    Next
End Function

Public Function IsChestItem(ByVal X As Long, ByVal Y As Long) As Byte
Dim RECT As RECT
Dim i As Byte

    IsChestItem = 0
    
    For i = 1 To MAX_CHEST_ITEMS
        With RECT
            .Left = 32 + ((32) * (((i - 1) Mod 8)))
            .top = (Int((i - 1) / 8) * 32) + 28 + (Int((i - 1) / 8) * 6)
            .Bottom = .top + 32
            .Right = .Left + 32
        End With
        
        If X >= RECT.Left And X <= RECT.Right Then
            If Y >= RECT.top And Y <= RECT.Bottom Then
                IsChestItem = i
                Exit Function
            End If
        End If
    Next
End Function

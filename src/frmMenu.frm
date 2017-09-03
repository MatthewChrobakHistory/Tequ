VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tequ"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13485
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   496
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   899
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrEnterGame 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9960
      Top             =   3000
   End
   Begin VB.PictureBox picOptions 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4185
      Left            =   5520
      ScaleHeight     =   279
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   541
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   8115
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   3720
         TabIndex        =   14
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   3600
         TabIndex        =   12
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblDebugMode 
         BackStyle       =   0  'Transparent
         Caption         =   "[ Debug Mode ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   15
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lblPort 
         BackStyle       =   0  'Transparent
         Caption         =   "[ Port ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lblIP 
         BackStyle       =   0  'Transparent
         Caption         =   "[ IP ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   1920
         Width           =   855
      End
   End
   Begin VB.PictureBox picPlayer 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   4185
      Left            =   1920
      ScaleHeight     =   279
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   541
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   8115
      Begin VB.TextBox txtPass2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3240
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtPass1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3240
         MaxLength       =   12
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   3240
         MaxLength       =   12
         TabIndex        =   2
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblPass2 
         BackStyle       =   0  'Transparent
         Caption         =   "Retype:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblPass1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblUsername 
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblLogin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[ Login ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4680
         TabIndex        =   6
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label lblCreate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "[ Create ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   2760
         Width           =   855
      End
   End
   Begin VB.Label lblErrorNotification 
      BackStyle       =   0  'Transparent
      Caption         =   "An error just occured. Click here to view it."
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Image imgMenuButton 
      Height          =   435
      Index           =   1
      Left            =   960
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Image imgMenuButton 
      Height          =   435
      Index           =   2
      Left            =   2460
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Image imgMenuButton 
      Height          =   435
      Index           =   3
      Left            =   3960
      Top             =   4305
      Width           =   1335
   End
   Begin VB.Image imgMenuButton 
      Height          =   435
      Index           =   4
      Left            =   5460
      Top             =   4305
      Width           =   1335
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    If picOptions.Visible = True Then picOptions.Visible = False
    If picPlayer.Visible = True Then picPlayer.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Byte
    
    For i = 1 To 4
        If Hovering(i) = True Then
            imgMenuButton(i).Picture = Nothing
            imgMenuButton(i).Picture = LoadPicture(Norm(i))
            Hovering(i) = False
        End If
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call DestroyClient

End Sub

Private Sub imgMenuButton_Click(Index As Integer)
    
    Select Case Index
        Case 1 'Singleplayer
            Options.OnlineMode = False
            Call SetupPlayerScreen
        Case 2 'Multiplayer
            Options.OnlineMode = True
            Call SetupPlayerScreen
        Case 3 'Options
            picPlayer.Visible = False
            picOptions.Visible = True
            Call ShowOptions
        Case 4 'Exit
            Call DestroyClient
    End Select
End Sub

Private Sub SetupPlayerScreen()
Dim Resolved As Boolean
Dim Tick As Long, tmr3000 As Long
Dim State As Long

    Select Case Options.OnlineMode
        Case False
            txtPass1.Visible = False
            txtPass2.Visible = False
            lblPass1.Visible = False
            lblPass2.Visible = False
        Case True
            ' setup the socket and hide the menu so we can try to connect
            frmMenu.Hide
            With frmMain.socket
                .close
                .RemoteHost = Options.IP
                .RemotePort = Options.Port
                .close
                .Connect
            End With

            'connection checking loop
            tmr3000 = timeGetTime + 5000
            Do While Resolved = False
                Tick = timeGetTime
                If tmr3000 < Tick Then
                    If frmMain.socket.State <> 7 Then
                        frmMain.socket.close
                        picPlayer.Visible = False
                        picOptions.Visible = False
                        frmMenu.Show
                        MsgBox "Connection failed!", vbCritical, "Failed to connect"
                        Resolved = True
                        Exit Sub
                        
                    End If
                ElseIf frmMain.socket.State = 7 Then
                    Resolved = True
                    frmMenu.Show
                    Call InitMessages
                End If
                DoEvents
                Sleep 1
            Loop
            
            txtPass1.Visible = True
            txtPass2.Visible = True
            lblPass1.Visible = True
            lblPass2.Visible = True
            txtPass1.text = Options.Password
            txtPass2.text = Options.Password
            
            ' Try to connect. If you can't, then show the homescreen and shoot an error message.
    End Select
    
    txtUsername.text = Options.Username
    
    picOptions.Visible = False
    picPlayer.Visible = True
    
End Sub

Private Sub imgMenuButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Byte
    
    For i = 1 To 4
        If Hovering(i) = True And i <> Index Then
            Hovering(i) = False
            imgMenuButton(i).Picture = LoadPicture(Norm(i))
        End If
    Next
    
    If Hovering(Index) = False Then
        Hovering(Index) = True
        imgMenuButton(Index).Picture = LoadPicture(Hover(Index))
    End If
    
End Sub

Private Sub lblCreate_Click()

    If Options.OnlineMode = True Then
        If Trim$(txtPass1.text) <> vbNullString Then
            If Trim$(txtPass1.text) = Trim$(txtPass2.text) Then
                Call SendCreatePlayer(txtUsername.text, txtPass1.text)
            Else
                'passwords didn't match
                MsgBox "Passwords don't match!", vbCritical
            End If
        End If
    ElseIf Options.OnlineMode = False Then
        If Len(Trim$(txtUsername.text)) > 0 Then Call MakeAccount(Trim$(txtUsername.text))
    End If

End Sub

Private Sub lblDebugMode_Click()

If Options.Debug = True Then
    Options.Debug = False
Else
    Options.Debug = True
End If

Call ShowOptions

End Sub

Private Sub lblLogin_Click()

    Select Case Options.OnlineMode
        Case True
            If frmMain.socket.State <> 7 Then
                Call DestroyGame
                Call MsgBox("Disconnected from server.", vbCritical)
                Exit Sub
            End If
            If txtPass1.text = txtPass2.text Then
                If Len(txtUsername.text) > 0 Then
                    Call SendRequestLogin(txtUsername.text, txtPass1.text)
                    Options.Password = txtPass1.text
                End If
            Else
                Call MsgBox("Passwords don't match!", vbCritical)
            End If
            
        Case False
            If Len(txtUsername.text) > 0 Then
                If FileExist(App.Path & "\data\players\" & Trim$(txtUsername.text) & ".bin") = True Then
                    MyIndex = 1
                    Call LoadPlayer(Trim$(txtUsername.text))
                    Call EnterGame
                Else
                    Call MsgBox("Player doesn't exist!", vbCritical)
                End If
            End If
    End Select

End Sub

Private Sub ShowOptions()
Dim DBValue As String

If Options.Debug = True Then
    DBValue = "True"
Else
    DBValue = "False"
End If

    lblDebugMode.Caption = "[ Debug Mode ]  " & DBValue
    txtIP.text = Options.IP
    txtPort.text = Options.Port
End Sub

Private Sub tmrEnterGame_Timer()

tmrEnterGame.Enabled = False
Call EnterGame

End Sub

Private Sub txtIP_Change()
    Options.IP = txtIP.text
End Sub

Private Sub txtPort_Change()
    If IsNumeric(txtPort.text) = False Then txtPort.text = "7001"
    Options.Port = txtPort.text
End Sub

Private Sub txtUsername_Change()

    Options.Username = txtUsername.text

End Sub


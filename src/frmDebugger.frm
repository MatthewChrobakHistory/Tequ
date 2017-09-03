VERSION 5.00
Begin VB.Form frmDebugger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraReport 
      Caption         =   "Reporting a bug"
      Height          =   4575
      Left            =   3600
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox txtReport 
         Height          =   1455
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   2520
         Width           =   3015
      End
      Begin VB.TextBox txtSource 
         Height          =   285
         Left            =   1200
         TabIndex        =   35
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtLine 
         Height          =   285
         Left            =   1200
         TabIndex        =   33
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   1200
         TabIndex        =   31
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtErrorType 
         Height          =   285
         Left            =   1200
         TabIndex        =   29
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmdReport 
         Caption         =   "Report!"
         Height          =   375
         Left            =   1920
         TabIndex        =   27
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Player written report: (optional)"
         Height          =   495
         Left            =   240
         TabIndex        =   36
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label5 
         Caption         =   "Source:"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Line:"
         Height          =   255
         Left            =   480
         TabIndex        =   32
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Error Type:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdReportCustom 
      Caption         =   "Report a bug"
      Height          =   495
      Left            =   1920
      TabIndex        =   24
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "Set Common RTE's"
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CheckBox chkPassAll 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Caption         =   "Pass all?"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   20
         Left            =   1800
         TabIndex        =   20
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   19
         Left            =   1800
         TabIndex        =   19
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   18
         Left            =   1800
         TabIndex        =   18
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   17
         Left            =   1800
         TabIndex        =   17
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   16
         Left            =   1800
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   15
         Left            =   1800
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   14
         Left            =   1800
         TabIndex        =   14
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   13
         Left            =   1800
         TabIndex        =   13
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   12
         Left            =   1800
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   11
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox chkPass 
         BackColor       =   &H000000FF&
         Caption         =   "Pass"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Common RTE's include 6, 7, 9, 11, 13, 28, 35. 52, 53, or 76. "
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Width           =   3375
   End
End
Attribute VB_Name = "frmDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkPass_Click(Index As Integer)
Dim ErrNum As String

    If LoadedGUI = False Then Exit Sub

    If chkPass(Index).HelpContextID = 1 Then
        chkPass(Index).HelpContextID = 0
        Exit Sub
    End If
    
    If ErrOptions.PassException(Index) = 0 And chkPass(Index).Value = 1 Then
        ErrNum = InputBox("Type an error to bypass, if one is produced.", "Error Handling")
        If Not IsNumeric(ErrNum) Then
            chkPass(Index).HelpContextID = 1
            chkPass(Index).Value = 0
            Exit Sub
        Else
            For i = 1 To MAX_PASS
                If ErrOptions.PassException(i) = ErrNum Then
                    chkPass(Index).HelpContextID = 1
                    chkPass(Index).Value = 0
                    Exit Sub
                ElseIf i = MAX_PASS Then
                    ErrOptions.PassException(Index) = ErrNum
                    chkPass(Index).Caption = "Pass RTE " & ErrNum
                    chkPass(Index).BackColor = DBGR_GREEN
                End If
            Next
        End If
    Else
        If chkPass(Index).Value = 0 Then
            ErrOptions.PassException(Index) = 0
            chkPass(Index).Caption = "Pass RTE 0"
            chkPass(Index).BackColor = DBGR_RED
        End If
    End If
    
End Sub

Private Sub chkPassAll_Click()
Dim i As Long
    ErrOptions.PassAll = chkPassAll.Value
    If ErrOptions.PassAll = True Then
        For i = 1 To MAX_PASS
            chkPass(i).Visible = False
        Next
        chkPassAll.BackColor = DBGR_GREEN
    Else
        For i = 1 To MAX_PASS
            chkPass(i).Visible = True
        Next
        chkPassAll.BackColor = DBGR_RED
    End If
End Sub

Private Sub cmdCancel_Click()
    With Me
        .txtDescription.Enabled = True
        .txtDescription.text = vbNullString
        .txtErrorType.Enabled = True
        .txtErrorType.text = vbNullString
        .txtLine.Enabled = True
        .txtLine.text = vbNullString
        .txtSource.Enabled = True
        .txtSource.text = vbNullString
        .txtReport.Enabled = True
        .txtReport.text = vbNullString
        .fraReport.Visible = False
    End With
End Sub

Private Sub cmdDefaults_Click()

    With ErrOptions
        .PassException(1) = 6
        .PassException(2) = 7
        .PassException(3) = 9
        .PassException(4) = 11
        .PassException(5) = 13
        .PassException(6) = 28
        .PassException(7) = 35
        .PassException(8) = 52
        .PassException(9) = 53
        .PassException(10) = 76
    End With
    
    Call InitDebugGUI
End Sub

Private Sub cmdReport_Click()
    If Trim$(txtDescription.text) = vbNullString Then
        Call MsgBox("Please add a short description in the second textbox!", vbCritical)
        Exit Sub
    End If
    Call PrintErrorReport(txtErrorType.text, txtDescription.text, txtReport.text, txtLine.text, txtSource.text)
    With Me
        .txtDescription.Enabled = True
        .txtDescription.text = vbNullString
        .txtErrorType.Enabled = True
        .txtErrorType.text = vbNullString
        .txtLine.Enabled = True
        .txtLine.text = vbNullString
        .txtSource.Enabled = True
        .txtSource.text = vbNullString
        .txtReport.Enabled = True
        .txtReport.text = vbNullString
        .fraReport.Visible = False
    End With
End Sub

Private Sub cmdReportCustom_Click()
    fraReport.Visible = True
    With Me
        .txtErrorType.text = "CUSTOM"
        .txtErrorType.Enabled = False
        .txtLine.text = "CUSTOM"
        .txtLine.Enabled = False
        .txtSource = "CUSTOM"
        .txtSource.Enabled = False
    End With
End Sub


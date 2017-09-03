VERSION 5.00
Begin VB.Form frmEditor_Interface 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   764
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picScreen 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   5760
      Left            =   120
      ScaleHeight     =   5760
      ScaleWidth      =   7200
      TabIndex        =   21
      Top             =   120
      Width           =   7200
      Begin VB.PictureBox picBackground 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   5760
         Left            =   0
         ScaleHeight     =   5760
         ScaleWidth      =   7200
         TabIndex        =   22
         Top             =   0
         Width           =   7200
         Begin VB.Label Label 
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   41
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   40
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   10
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   11
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   12
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   13
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   14
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   15
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   16
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   17
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   18
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   19
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label 
            Height          =   255
            Index           =   20
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
      End
   End
   Begin VB.Frame fraObject 
      Height          =   2175
      Left            =   7440
      TabIndex        =   7
      Top             =   4080
      Width           =   3975
      Begin VB.TextBox txtEventData 
         Height          =   855
         Left            =   2400
         TabIndex        =   19
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtEvent 
         Height          =   285
         Left            =   2400
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   1935
      End
      Begin VB.HScrollBar scrlTop 
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.HScrollBar scrlLeft 
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.HScrollBar scrlWidth 
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.HScrollBar scrlHeight 
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Event Data"
         Height          =   255
         Left            =   2400
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Event:"
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Height"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Width"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Top"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Left"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.HScrollBar scrlPage 
      Height          =   255
      Left            =   120
      Min             =   1
      TabIndex        =   1
      Top             =   6000
      Value           =   1
      Width           =   7095
   End
   Begin VB.Frame fraLabel 
      Height          =   3975
      Left            =   7440
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.HScrollBar scrlSize 
         Height          =   255
         Left            =   1800
         Max             =   24
         Min             =   8
         TabIndex        =   46
         Top             =   2640
         Value           =   8
         Width           =   1695
      End
      Begin VB.TextBox txtCaption 
         Height          =   525
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CheckBox chkOpactiy 
         Caption         =   "Opacity?"
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtForeColor_Label 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtBackColor_Label 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblSize 
         Caption         =   "Font Size: 8"
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Caption:"
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "ForeColor"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "BackColor"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Menu New 
      Caption         =   "New"
      Begin VB.Menu New_Label 
         Caption         =   "New Label"
      End
      Begin VB.Menu New_Picturebox 
         Caption         =   "New Picturebox"
      End
   End
End
Attribute VB_Name = "frmEditor_Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkOpactiy_Click()
    If chkOpactiy.Value = 1 Then
        UserInterface(EInterfaceNum).Page(EditingPage).Label(Editing_index).Opacity = True
    Else
        UserInterface(EInterfaceNum).Page(EditingPage).Label(Editing_index).Opacity = False
    End If
    Me.Label(Editing_index).BackStyle = chkOpactiy.Value
End Sub

Private Sub Label_Click(Index As Integer)
    Editing = EDITING_LABEL
    Editing_index = Index
    Call NewEditing
End Sub

Private Sub New_Label_Click()
Dim i As Byte
    With UserInterface(EInterfaceNum).Page(EditingPage)
        For i = 1 To MAX_INTERFACE_LABELS
            If .Label(i).Object.Visible = False Then
                With .Label(i)
                    .Object.Visible = True
                    Me.Label(i).Visible = True
                    Editing = EDITING_LABEL
                    Editing_index = i
                    Call NewEditing
                    Exit For
                End With
            End If
        Next
    End With
End Sub

Private Sub picBackground_Click()
    Editing = EDITING_BACKGROUND
    Editing_index = 0
End Sub

Private Sub scrlHeight_Change()
    Select Case Editing
        Case EDITING_LABEL
            UserInterface(EInterfaceNum).Page(EditingPage).Label(Editing_index).Object.Height = scrlHeight.Value
            Me.Label(Editing_index).Height = scrlHeight.Value
    End Select
End Sub

Private Sub scrlPage_Change()
    EditingPage = scrlPage.Value
    Me.Caption = "Editing " & Trim$(UserInterface(EInterfaceNum).Name) + ": " & scrlPage.Value
    Call NewPage(scrlPage.Value)
End Sub

Public Sub NewPage(ByVal Index As Long)
Dim i As Long

    With frmEditor_Interface
        For i = 1 To MAX_INTERFACE_LABELS
            .Label(i).Caption = ""
            .Label(i).BackColor = &H8000000F
            .Label(i).BackStyle = 1
            .Label(i).Left = 0
            .Label(i).Top = 0
            .Label(i).Width = 1215
            .Label(i).Height = 255
            .Label(i).ForeColor = &H80000012
            .Label(i).Visible = False
        Next
    End With

    With UserInterface(EInterfaceNum).Page(Index)
        For i = 1 To MAX_INTERFACE_LABELS
            frmEditor_Interface.Label(i).Caption = .Label(i).Caption
            frmEditor_Interface.Label(i).BackColor = .Label(i).BackColor
            frmEditor_Interface.Label(i).BackStyle = .Label(i).Opacity
            frmEditor_Interface.Label(i).Left = .Label(i).Object.Left
            frmEditor_Interface.Label(i).Top = .Label(i).Object.Top
            frmEditor_Interface.Label(i).Width = .Label(i).Object.Width
            frmEditor_Interface.Label(i).Height = .Label(i).Object.Height
            frmEditor_Interface.Label(i).Visible = .Label(i).Object.Visible
        Next
    End With
End Sub

Public Sub NewEditing()

    Me.fraLabel.Visible = False
    Me.fraObject.Visible = False
    
    Select Case Editing
        Case EDITING_BACKGROUND
        Case EDITING_LABEL
            fraLabel.Visible = True
            fraObject.Visible = True
            
            With UserInterface(EInterfaceNum).Page(scrlPage.Value).Label(Editing_index)
                    Me.txtCaption.text = .Caption
                    Me.txtBackColor_Label.text = .BackColor
                    Me.txtForeColor_Label.text = .ForeColor
                    Me.chkOpactiy.Value = .Opacity
                    Me.scrlLeft.Value = .Object.Left
                    Me.scrlTop.Value = .Object.Top
                    Me.scrlWidth.Value = .Object.Width
                    Me.scrlHeight.Value = .Object.Height
            End With
            
            
        Case EDITING_PICTUREBOX
    End Select
End Sub

Private Sub scrlWidth_Change()
    Select Case Editing
        Case EDITING_LABEL
            UserInterface(EInterfaceNum).Page(EditingPage).Label(Editing_index).Object.Width = scrlWidth.Value
            Me.Label(Editing_index).Width = scrlWidth.Value
    End Select
End Sub

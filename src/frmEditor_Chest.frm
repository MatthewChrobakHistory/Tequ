VERSION 5.00
Begin VB.Form frmEditor_Chest 
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   3735
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   3495
         Begin VB.HScrollBar scrlValue 
            Height          =   255
            Left            =   1560
            TabIndex        =   14
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtChance 
            Height          =   285
            Left            =   1800
            TabIndex        =   13
            Top             =   1080
            Width           =   1575
         End
         Begin VB.HScrollBar scrlItem 
            Height          =   255
            Left            =   1560
            TabIndex        =   10
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblPercentile 
            Caption         =   "Chance: (value/100)"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblValue 
            Caption         =   "Value: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lblItem 
            Caption         =   "Item: None"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.HScrollBar scrlChestItem 
         Height          =   255
         Left            =   240
         Min             =   1
         TabIndex        =   7
         Top             =   470
         Value           =   1
         Width           =   3255
      End
      Begin VB.Label lblChestItem 
         Caption         =   "Chest Item: 1"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.HScrollBar scrlPicture 
      Height          =   255
      Left            =   2880
      Min             =   1
      TabIndex        =   4
      Top             =   480
      Value           =   1
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox lstIndex 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblPicture 
      Caption         =   "Picture: 0"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmEditor_Chest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
Dim I As Long
    For I = 1 To MAX_CHESTS
        Call SaveChest(I)
    Next
End Sub

Private Sub lstIndex_Click()

    EChestNum = lstIndex.ListIndex + 1
    Call NewChestIndex(EChestNum)

End Sub

Private Sub scrlChestItem_Change()

    scrlItem.Value = Chest(EChestNum).ChestItem(scrlChestItem.Value).Itemnum
    scrlValue.Value = Chest(EChestNum).ChestItem(scrlChestItem.Value).ItemValue
    txtChance.text = Chest(EChestNum).ChestItem(scrlChestItem.Value).Chance
    lblChestItem.Caption = "Chest Item: " & scrlChestItem.Value

End Sub

Private Sub scrlItem_Change()

    Chest(EChestNum).ChestItem(scrlChestItem.Value).Itemnum = scrlItem.Value

    If scrlItem.Value = 0 Then
        lblItem.Caption = "Item: None"
        Exit Sub
    End If
    lblItem.Caption = "Item: " & Trim$(Item(scrlItem.Value).name)

End Sub

Private Sub scrlPicture_Change()

    Chest(EChestNum).Picture = scrlPicture.Value
    lblPicture.Caption = "Picture: " & scrlPicture.Value

End Sub

Private Sub scrlValue_Change()

    Chest(EChestNum).ChestItem(scrlChestItem.Value).ItemValue = scrlValue.Value
    lblValue.Caption = "Value: " & scrlValue.Value

End Sub

Private Sub txtChance_Change()

    If Not IsNumeric(txtChance.text) Then
        txtChance.text = Chest(EChestNum).ChestItem(scrlChestItem.Value).Chance
        If Chest(EChestNum).ChestItem(scrlChestItem.Value).Chance = 0 Then
            Chest(EChestNum).ChestItem(scrlChestItem.Value).Chance = 1
            txtChance.text = "1"
        End If
    End If
    If txtChance.text > 100 Then
        txtChance.text = "100"
    End If
    
    Chest(EChestNum).ChestItem(scrlChestItem.Value).Chance = txtChance.text
    
End Sub

Private Sub txtName_Change()

    Chest(EChestNum).name = Trim$(txtName.text)

End Sub

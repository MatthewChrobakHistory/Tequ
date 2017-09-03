VERSION 5.00
Begin VB.Form frmEditor_Shop 
   Caption         =   "Form1"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   11250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   9375
      Begin VB.Frame Frame2 
         Caption         =   "Stock"
         Height          =   6735
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   9135
         Begin VB.HScrollBar scrlStockValue 
            Height          =   255
            Left            =   240
            TabIndex        =   69
            Top             =   3960
            Width           =   1455
         End
         Begin VB.Frame Frame4 
            Height          =   975
            Left            =   4560
            TabIndex        =   62
            Top             =   600
            Width           =   4455
            Begin VB.ComboBox cmbItemCosts 
               Height          =   315
               ItemData        =   "frmEditor_Shop.frx":0000
               Left            =   3000
               List            =   "frmEditor_Shop.frx":0025
               TabIndex        =   66
               Text            =   "Free"
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox txtVerb 
               Height          =   285
               Left            =   3000
               TabIndex        =   64
               Top             =   240
               Width           =   1335
            End
            Begin VB.CheckBox chkXP 
               Caption         =   "Enable XP?"
               Height          =   255
               Left            =   240
               TabIndex        =   63
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "Number of Costs:"
               Height          =   255
               Left            =   1680
               TabIndex        =   67
               Top             =   600
               Width           =   2655
            End
            Begin VB.Label Label6 
               Caption         =   "Verb to use:"
               Height          =   255
               Left            =   1680
               TabIndex        =   65
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.HScrollBar scrlPicture 
            Height          =   255
            Left            =   7320
            Max             =   10
            TabIndex        =   61
            Top             =   240
            Width           =   1575
         End
         Begin VB.Frame Frame3 
            Height          =   5295
            Left            =   4560
            TabIndex        =   9
            Top             =   1320
            Width           =   4455
            Begin VB.HScrollBar scrlAmount 
               Height          =   225
               Index           =   10
               Left            =   2775
               TabIndex        =   59
               Top             =   4800
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlItem 
               Height          =   225
               Index           =   10
               Left            =   495
               TabIndex        =   58
               Top             =   4800
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.CheckBox chkTakeAway 
               Caption         =   "Use up "
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   55
               Top             =   4560
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.HScrollBar scrlAmount 
               Height          =   225
               Index           =   9
               Left            =   2775
               TabIndex        =   54
               Top             =   4320
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlItem 
               Height          =   225
               Index           =   9
               Left            =   495
               TabIndex        =   53
               Top             =   4320
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlAmount 
               Height          =   225
               Index           =   8
               Left            =   2775
               TabIndex        =   50
               Top             =   3840
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlItem 
               Height          =   225
               Index           =   8
               Left            =   495
               TabIndex        =   49
               Top             =   3840
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlAmount 
               Height          =   225
               Index           =   7
               Left            =   2775
               TabIndex        =   46
               Top             =   3360
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlItem 
               Height          =   225
               Index           =   7
               Left            =   495
               TabIndex        =   45
               Top             =   3360
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlAmount 
               Height          =   225
               Index           =   6
               Left            =   2775
               TabIndex        =   42
               Top             =   2880
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlItem 
               Height          =   225
               Index           =   6
               Left            =   495
               TabIndex        =   41
               Top             =   2880
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlAmount 
               Height          =   225
               Index           =   5
               Left            =   2775
               TabIndex        =   38
               Top             =   2400
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlItem 
               Height          =   225
               Index           =   5
               Left            =   495
               TabIndex        =   37
               Top             =   2400
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlAmount 
               Height          =   225
               Index           =   4
               Left            =   2775
               TabIndex        =   34
               Top             =   1920
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlItem 
               Height          =   225
               Index           =   4
               Left            =   495
               TabIndex        =   33
               Top             =   1920
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlAmount 
               Height          =   225
               Index           =   3
               Left            =   2775
               TabIndex        =   30
               Top             =   1440
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlItem 
               Height          =   225
               Index           =   3
               Left            =   495
               TabIndex        =   29
               Top             =   1440
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlAmount 
               Height          =   225
               Index           =   2
               Left            =   2775
               TabIndex        =   26
               Top             =   960
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlItem 
               Height          =   225
               Index           =   2
               Left            =   495
               TabIndex        =   25
               Top             =   960
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.CheckBox chkTakeAway 
               Caption         =   "Use up "
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   22
               Top             =   4080
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.CheckBox chkTakeAway 
               Caption         =   "Use up "
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   21
               Top             =   3600
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.CheckBox chkTakeAway 
               Caption         =   "Use up "
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   20
               Top             =   3120
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.CheckBox chkTakeAway 
               Caption         =   "Use up "
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   19
               Top             =   2640
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.CheckBox chkTakeAway 
               Caption         =   "Use up "
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   18
               Top             =   2160
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.CheckBox chkTakeAway 
               Caption         =   "Use up "
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   17
               Top             =   1680
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.CheckBox chkTakeAway 
               Caption         =   "Use up "
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   16
               Top             =   1200
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.CheckBox chkTakeAway 
               Caption         =   "Use up "
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   15
               Top             =   720
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.HScrollBar scrlAmount 
               Height          =   225
               Index           =   1
               Left            =   2780
               TabIndex        =   14
               Top             =   480
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.HScrollBar scrlItem 
               Height          =   225
               Index           =   1
               Left            =   500
               TabIndex        =   13
               Top             =   480
               Value           =   1
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.CheckBox chkTakeAway 
               Caption         =   "Use up "
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   10
               Top             =   240
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.Label Label4 
               Caption         =   "Amount:"
               Height          =   255
               Index           =   10
               Left            =   2160
               TabIndex        =   57
               Top             =   4800
               Visible         =   0   'False
               Width           =   660
            End
            Begin VB.Label Label3 
               Caption         =   "Item:"
               Height          =   255
               Index           =   10
               Left            =   120
               TabIndex        =   56
               Top             =   4800
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label Label4 
               Caption         =   "Amount:"
               Height          =   255
               Index           =   9
               Left            =   2160
               TabIndex        =   52
               Top             =   4320
               Visible         =   0   'False
               Width           =   660
            End
            Begin VB.Label Label3 
               Caption         =   "Item:"
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   51
               Top             =   4320
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label Label4 
               Caption         =   "Amount:"
               Height          =   255
               Index           =   8
               Left            =   2160
               TabIndex        =   48
               Top             =   3840
               Visible         =   0   'False
               Width           =   660
            End
            Begin VB.Label Label3 
               Caption         =   "Item:"
               Height          =   255
               Index           =   8
               Left            =   120
               TabIndex        =   47
               Top             =   3840
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label Label4 
               Caption         =   "Amount:"
               Height          =   255
               Index           =   7
               Left            =   2160
               TabIndex        =   44
               Top             =   3360
               Visible         =   0   'False
               Width           =   660
            End
            Begin VB.Label Label3 
               Caption         =   "Item:"
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   43
               Top             =   3360
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label Label4 
               Caption         =   "Amount:"
               Height          =   255
               Index           =   6
               Left            =   2160
               TabIndex        =   40
               Top             =   2880
               Visible         =   0   'False
               Width           =   660
            End
            Begin VB.Label Label3 
               Caption         =   "Item:"
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   39
               Top             =   2880
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label Label4 
               Caption         =   "Amount:"
               Height          =   255
               Index           =   5
               Left            =   2160
               TabIndex        =   36
               Top             =   2400
               Visible         =   0   'False
               Width           =   660
            End
            Begin VB.Label Label3 
               Caption         =   "Item:"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   35
               Top             =   2400
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label Label4 
               Caption         =   "Amount:"
               Height          =   255
               Index           =   4
               Left            =   2160
               TabIndex        =   32
               Top             =   1920
               Visible         =   0   'False
               Width           =   660
            End
            Begin VB.Label Label3 
               Caption         =   "Item:"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   31
               Top             =   1920
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label Label4 
               Caption         =   "Amount:"
               Height          =   255
               Index           =   3
               Left            =   2160
               TabIndex        =   28
               Top             =   1440
               Visible         =   0   'False
               Width           =   660
            End
            Begin VB.Label Label3 
               Caption         =   "Item:"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   27
               Top             =   1440
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label Label4 
               Caption         =   "Amount:"
               Height          =   255
               Index           =   2
               Left            =   2160
               TabIndex        =   24
               Top             =   960
               Visible         =   0   'False
               Width           =   660
            End
            Begin VB.Label Label3 
               Caption         =   "Item:"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   23
               Top             =   960
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label Label4 
               Caption         =   "Amount:"
               Height          =   255
               Index           =   1
               Left            =   2160
               TabIndex        =   12
               Top             =   480
               Visible         =   0   'False
               Width           =   660
            End
            Begin VB.Label Label3 
               Caption         =   "Item:"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   11
               Top             =   480
               Visible         =   0   'False
               Width           =   540
            End
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   5040
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   1800
            TabIndex        =   6
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Height          =   375
            Left            =   1800
            TabIndex        =   5
            Top             =   2400
            Width           =   1095
         End
         Begin VB.ListBox lstItems 
            Height          =   3375
            Left            =   3000
            TabIndex        =   4
            Top             =   240
            Width           =   1455
         End
         Begin VB.ListBox lstStock 
            Height          =   3375
            Left            =   240
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblStockValue 
            Caption         =   "Amount: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   3720
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Picture: 0"
            Height          =   255
            Left            =   6480
            TabIndex        =   60
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Name:"
            Height          =   255
            Left            =   4560
            TabIndex        =   7
            Top             =   240
            Width           =   735
         End
      End
   End
   Begin VB.ListBox lstIndex 
      Height          =   6885
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmEditor_Shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkTakeAway_Click(Index As Integer)

    If chkTakeAway(Index).Value = 1 Then
        Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).ItemCost(Index).UseUpItem = True
    Else
        Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).ItemCost(Index).UseUpItem = False
    End If

End Sub

Private Sub chkXP_Click()

    If chkXP.Value = 1 Then
        Shop(EShopNum).ShopItem(frmEditor_Shop.lstStock.ListIndex + 1).AddXP = True
    Else
        Shop(EShopNum).ShopItem(frmEditor_Shop.lstStock.ListIndex + 1).AddXP = False
    End If
    
End Sub

Private Sub cmbItemCosts_Click()
Dim I As Long

    For I = 1 To 10
        Me.scrlAmount(I).Visible = False
        Me.scrlItem(I).Visible = False
        Me.Label3(I).Visible = False
        Me.Label4(I).Visible = False
        Me.chkTakeAway(I).Visible = False
    Next

    If cmbItemCosts.ListIndex > 0 Then
        For I = 1 To cmbItemCosts.ListIndex
            Me.scrlAmount(I).Visible = True
            Me.scrlItem(I).Visible = True
            Me.Label3(I).Visible = True
            Me.Label4(I).Visible = True
            Me.chkTakeAway(I).Visible = True
        Next
    End If
    
    If lstStock.ListIndex < 0 Then
        lstStock.ListIndex = 0
    End If
    
    Shop(EShopNum).ShopItem(frmEditor_Shop.lstStock.ListIndex + 1).NumberofCosts = cmbItemCosts.ListIndex
    
    For I = 1 To cmbItemCosts.ListIndex
        With Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).ItemCost(I)
            scrlAmount(I).Value = .ItemCostValue
            scrlItem(I).Value = .ItemCostNum
            If .UseUpItem = True Then
                chkTakeAway(I).Value = 1
            Else
                chkTakeAway(I).Value = 0
            End If
            If .ItemCostNum > 0 Then
                chkTakeAway(I).Caption = "Take away item " & .ItemCostValue & "x  " & Trim$(Item(.ItemCostNum).name)
            Else
                chkTakeAway(I).Caption = "Take away item " & .ItemCostValue & "x None"
            End If
        End With
    Next
        
End Sub

Private Sub cmdAdd_Click()
Dim Store As Long

    Store = lstStock.ListIndex
    Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).StockItem = lstItems.ListIndex + 1
    
    lstStock.Clear
    For I = 1 To MAX_SHOP_ITEMS
        If Shop(EShopNum).ShopItem(I).StockItem = 0 Then
            lstStock.AddItem (I & ": ")
        Else
            lstStock.AddItem (I & ": " & Trim$(Item(Shop(EShopNum).ShopItem(I).StockItem).name))
        End If
    Next
    
    lstStock.ListIndex = Store
    
    If Item(Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).StockItem).Stackable = False Then
        scrlStockValue.Visible = False
        lblStockValue.Visible = False
        Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).StockValue = 1
    Else
        scrlStockValue.Visible = True
        lblStockValue.Visible = True
    End If

End Sub

Private Sub cmdRemove_Click()
Dim Store As Long

    Store = lstStock.ListIndex
    Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).StockItem = 0
    
    lstStock.Clear
    For I = 1 To MAX_SHOP_ITEMS
        If Shop(EShopNum).ShopItem(I).StockItem = 0 Then
            lstStock.AddItem (I & ": ")
        Else
            lstStock.AddItem (I & ": " & Trim$(Item(Shop(EShopNum).ShopItem(I).StockItem).name))
        End If
    Next
    
    lstStock.ListIndex = Store
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim I As Long
    For I = 1 To MAX_SHOPS
        Call SaveShop(I)
    Next
End Sub

Private Sub lstIndex_Click()

    EShopNum = lstIndex.ListIndex + 1
    Call NewShopIndex(EShopNum)

End Sub

Private Sub lstStock_Click()
Dim I As Byte

    cmbItemCosts.ListIndex = Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).NumberofCosts
    txtVerb.text = Trim$(Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).Verb)
    For I = 1 To Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).NumberofCosts
        If Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).ItemCost(I).UseUpItem = True Then
            chkTakeAway(I).Value = 1
        Else
            chkTakeAway(I).Value = 0
        End If
    Next
    scrlStockValue.Value = Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).StockValue
    If Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).AddXP = True Then
        chkXP.Value = 1
    Else
        chkXP.Value = 0
    End If
    
    scrlStockValue.Value = Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).StockValue
    scrlStockValue.Visible = False
    lblStockValue.Visible = False
    If Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).StockItem > 0 Then
        If Item(Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).StockItem).Stackable = True Then
            scrlStockValue.Visible = True
            lblStockValue.Visible = True
        Else
            Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).StockValue = 1
        End If
    End If
    
End Sub

Private Sub scrlAmount_Change(Index As Integer)
Dim Itemnum As Long
Dim ItemValue As Long

    Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).ItemCost(Index).ItemCostValue = scrlAmount(Index).Value
    
    Itemnum = Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).ItemCost(Index).ItemCostNum
    ItemValue = Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).ItemCost(Index).ItemCostValue
    If Itemnum > 0 Then
        chkTakeAway(Index).Caption = "Take away item " & ItemValue & "x  " & Trim$(Item(Itemnum).name)
    Else
        chkTakeAway(Index).Caption = "Take away item " & ItemValue & "x None"
    End If
End Sub

Private Sub scrlItem_Change(Index As Integer)
Dim Itemnum As Long
Dim ItemValue As Long

    Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).ItemCost(Index).ItemCostNum = scrlItem(Index).Value
    
    Itemnum = Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).ItemCost(Index).ItemCostNum
    ItemValue = Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).ItemCost(Index).ItemCostValue
    If Itemnum > 0 Then
        chkTakeAway(Index).Caption = "Take away item " & ItemValue & "x  " & Trim$(Item(Itemnum).name)
    Else
        chkTakeAway(Index).Caption = "Take away item " & ItemValue & "x None"
    End If
End Sub

Private Sub scrlPicture_Change()

    Shop(EShopNum).Picture = scrlPicture.Value

End Sub

Private Sub scrlStockValue_Change()

    If lstStock.ListIndex < 0 Then lstStock.ListIndex = 0
    Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).StockValue = scrlStockValue.Value
    
    lblStockValue.Caption = "Value: " & scrlStockValue.Value

End Sub

Private Sub txtName_Change()

    Shop(EShopNum).name = txtName.text

End Sub

Private Sub txtVerb_Change()

    If lstStock.ListIndex < 0 Then lstStock.ListIndex = 0
    Shop(EShopNum).ShopItem(lstStock.ListIndex + 1).Verb = txtVerb.text

End Sub

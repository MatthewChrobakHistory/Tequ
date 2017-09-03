VERSION 5.00
Begin VB.Form frmAdminPanel 
   Caption         =   "Admin Panel"
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   2880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox lstIndex 
      Height          =   3570
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Menu Data 
      Caption         =   "Data"
      Begin VB.Menu Npc 
         Caption         =   "Npc"
      End
      Begin VB.Menu Map 
         Caption         =   "Map"
      End
      Begin VB.Menu Item 
         Caption         =   "Item"
      End
      Begin VB.Menu Resource 
         Caption         =   "Resource"
      End
      Begin VB.Menu Shop 
         Caption         =   "Shop"
      End
      Begin VB.Menu Chest 
         Caption         =   "Chest"
      End
      Begin VB.Menu Spell 
         Caption         =   "Spell"
      End
      Begin VB.Menu Interface 
         Caption         =   "Interface"
      End
   End
   Begin VB.Menu Tools 
      Caption         =   "Tools"
      Begin VB.Menu PacketViewer 
         Caption         =   "Packet Viewer"
      End
      Begin VB.Menu LoadedMapList 
         Caption         =   "Loaded Map List"
      End
   End
   Begin VB.Menu UsefulExtras 
      Caption         =   "Useful Extras"
      Begin VB.Menu Debugger 
         Caption         =   "Debugger"
      End
      Begin VB.Menu RespawnMap 
         Caption         =   "Respawn Map"
      End
      Begin VB.Menu UnAdminMe 
         Caption         =   "Un-Admin Me"
      End
      Begin VB.Menu ChangeMyLook 
         Caption         =   "Change My Look"
      End
      Begin VB.Menu PlayerEditor 
         Caption         =   "Player Editor"
      End
   End
End
Attribute VB_Name = "frmAdminPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChangeMyLook_Click()
    CreatingCharacter = True
    frmMain.fraPlayerCreate.Visible = True
End Sub

Public Sub PanelInit()
    Me.Show
End Sub

Private Sub Chest_Click()
    Call InitChestEditor
    Call Unload(Me)
End Sub

Private Sub cmdEdit_Click()
    If lstIndex.ListIndex >= 0 Then
        EInterfaceNum = lstIndex.ListIndex + 1
        Call InitInterfaceEditor
    End If
End Sub

Private Sub Debugger_Click()
    'Call InitDebugGUI
    'Call Unload(Me)
End Sub

Private Sub Interface_Click()
Dim i As Long

    lstIndex.Clear
    lstIndex.Visible = True
    cmdEdit.Visible = True
    For i = 1 To MAX_INTERFACES
        lstIndex.AddItem i & ": " & Trim$(UserInterface(i).Name)
    Next
End Sub

Private Sub Item_Click()
    Call InitItemEditor
    Call Unload(Me)
End Sub

Private Sub Items_Click()
Dim i As Long
    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next
End Sub

Private Sub LoadedMapList_Click()
Dim i As Long

    lstIndex.Clear
    lstIndex.Visible = True
    For i = 1 To MAX_MAPS
        lstIndex.AddItem i & ": " & LoadedMap(i)
    Next

End Sub

Private Sub lstIndex_DblClick()

    lstIndex.Visible = False

End Sub

Private Sub Map_Click()
    Call InitMapEditor
    Call Unload(Me)
End Sub

Private Sub Maps_Click()
Dim i As Long
    For i = 1 To MAX_MAPS
        If LoadedMap(i) = True Then
            SaveMap (i)
        End If
    Next
End Sub

Private Sub Npc_Click()
    Call InitNpcEditor
    Call Unload(Me)
End Sub

Private Sub Npcs_Click()
Dim i As Long
    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next
End Sub

Private Sub PlayerEditor_Click()
    frmAdminEditor.Show
End Sub

Private Sub Resource_Click()
    Call InitResourceEditor
    Call Unload(Me)
End Sub

Private Sub Resources_Click()
Dim i As Long
    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next
End Sub

Private Sub RespawnMap_Click()
    Call LoadMap(Player(MyIndex).Map)
    Call AddText("Map reloaded.", Blue)
End Sub

Private Sub Shop_Click()
    Call InitShopEditor
    Call Unload(Me)
End Sub

Private Sub Spell_Click()
    Call InitSpellEditor
    Call Unload(Me)
End Sub

Private Sub Spells_Click()
Dim i As Long
    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next
End Sub

Private Sub UnAdminMe_Click()
    Player(MyIndex).Access = ACCESS_PLAYER
    Call UnloadEditors
End Sub

Public Sub UnloadEditors()
    Call Unload(frmAdminPanel)
    Call Unload(frmEditor_Item)
    Call Unload(frmEditor_Map)
    Call Unload(frmEditor_Npc)
End Sub

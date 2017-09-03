Attribute VB_Name = "modEnumerations"
Option Explicit

Public Enum ServerPackets
    SMsgBox = 1
    SEnterGame
    SPlayerData
    SClientAddText ' adds text to txtChat on frmMain
    SMapData
    SPlayerMove
    SCanStop
    SDropItem
    SClearItem
    SClearPlayer
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

Public Enum ClientPackets
    CRequestLogin = 1
    CCreatePlayer
    CServerMessage ' Sending the text from txtMyChat
    CPlayerMove
    CCanStop
    CDropItem
    CPickUpItem
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum

' Map Layers
Public Enum Layers
    Ground = 1
    Mask1
    Mask2
    Mask3
    MaskAnim
    Fringe1
    Fringe2
    Fringe3
    FringeAnim
    Layer_Count
End Enum

' Vitals
Public Enum Vitals
    Health = 1
    Spirit
    Vital_Count
End Enum

' Stats
Public Enum Stats
    Attack = 1
    Strength
    Defense
    Agility
    Sagacity
    Stat_Count
End Enum

Public Enum Equipment
    Null1 = 1
    Head
    Null2
    Weapon
    Body
    Shield
    Hands
    Legs
    Boots
    Equipment_Count
End Enum

Public Enum Skills
    Woodcutting = 1
    Mining
    Fishing
    Smithing
    Cooking
    Fletching
    Crafting
    PotionBrewing
    Skill_Count
End Enum

Public Enum Attributes
    BlockedTile = 1
    WarpTile
    SoundTile
    ItemTile
    HealTile
    TrapTile
    NpcSpawnTile
    NpcAvoidTile
    KeyTile
    ResourceTile
    BankTile
    ShopTile
    ChestTile
End Enum
    
Public Enum Tabs
    Inventory = 1
    MyOptions
    Character
    Spells
    Skills
    Quit
    Tabs_Count
End Enum

Public Enum Combat
    Melee = 1
    Ranged
    Magic
    Combat_Count
End Enum

Public Enum Stance
    MNorm = 1
    MShield
    MTwoHand
    FNorm
    FShield
    FTwoHand
    Stance_Count
End Enum
Public HandleDataSub(SMSG_COUNT) As Long

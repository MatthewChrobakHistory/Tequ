Attribute VB_Name = "modConstants"
' DX7
Public Const NumTilesets As Long = 4
Public Const NumItems As Long = 13
Public Const NumSprites As Long = 5
Public Const NumPaperdolls As Long = 6
Public Const NumChests As Long = 1
Public Const NumResources As Long = 2
Public Const NumSpells As Long = 2
Public Const NumProjectiles As Long = 2
' Sprites
Public Const NumMaleSkin As Byte = 12
Public Const NumMaleNormHair As Byte = 12
Public Const NumMaleNormBody As Byte = 12
Public Const NumMaleNormLegs As Byte = 12
Public Const NumFemaleSkin As Byte = 12
Public Const NumFemaleNormHair As Byte = 12
Public Const NumFemaleNormBody As Byte = 12
Public Const NumFemaleNormLegs As Byte = 12

Public Const MAX_PLAYERS As Long = 10
Public Const MAX_MAPS As Long = 100
Public Const FIRST_DUNGEON_MAP As Byte = 51
Public Const MAX_NPCS As Long = 2
Public Const MAX_MAP_NPCS As Long = 2
Public Const MAX_NPC_DROPS As Byte = 10
Public Const MAX_ITEMS As Long = 10
Public Const MAX_MAP_ITEM_LAYERS As Byte = 21
Public Const Default_Map_Item_Appear As Long = 3000 '3seconds
Public Const Default_Map_Item_Despawn As Long = 30000 '30seconds
Public Const MAPITEMSTATE_Me As Byte = 1
Public Const MAPITEMSTATE_All As Byte = 2
Public Const MAX_COMBAT_LEVEL As Byte = 100
Public Const MAX_RESOURCES As Long = 100
Public Const MAX_SHOPS As Long = 10
Public Const MAX_SHOP_ITEMS As Byte = 30
Public Const MAX_CHESTS As Byte = 10
Public Const MAX_CHEST_ITEMS As Byte = 16
Public Const MAX_SPELLS As Byte = 10
Public Const MAX_PLAYER_SPELLS As Byte = 35
Public Const MAX_MAP_PROJECTILES As Byte = 255
Public Const MAX_DROPS As Byte = 10

Public Const MIN_DUNGEON_MAP As Long = 50
Public Const MAX_DUNGEON_MAP As Long = 60

Public Const ACTIONMSG_STATIC As Long = 0
Public Const ACTIONMSG_SCROLL As Long = 1
Public Const ACTIONMSG_SCREEN As Long = 2

Public Const SHOP_STATE_NONE As Byte = 0
Public Const SHOP_STATE_BUY As Byte = 1
Public Const SHOP_STATE_SELL As Byte = 2

Public Const MAX_BANK_TABS As Byte = 10
Public Const MAX_BANK_ITEMS As Byte = 99

Public Const GENDER_MALE As Byte = 1
Public Const GENDER_FEMALE As Byte = 2

Public Const Vital_Bar_Width As Byte = 241

Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_DUNGEON As Byte = 1

Public Const MAX_INV As Byte = 35

Public Const MAX_MAP_X As Byte = 15
Public Const MAX_MAP_Y As Byte = 12
Public Const NAME_LENGTH As Byte = 25
Public Const INFO_LENGTH As Byte = 100

Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3
Public Const DIR_UP As Byte = 4

' text color constants
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15

' Font variables
Public Const FONT_SIZE As Byte = 14

Public Const ACCESS_PLAYER As Byte = 1
Public Const ACCESS_MEMBER As Byte = 2
Public Const ACCESS_MODERATOR As Byte = 3
Public Const ACCESS_ADMIN As Byte = 4
Public Const ACCESS_OWNER As Byte = 5

' Equipment
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_NULL1 As Byte = 1
Public Const ITEM_TYPE_HEAD As Byte = 2
Public Const ITEM_TYPE_NULL2 As Byte = 3
Public Const ITEM_TYPE_WEAPON As Byte = 4
Public Const ITEM_TYPE_BODY As Byte = 5
Public Const ITEM_TYPE_SHIELD As Byte = 6
Public Const ITEM_TYPE_HANDS As Byte = 7
Public Const ITEM_TYPE_LEGS As Byte = 8
Public Const ITEM_TYPE_BOOTS As Byte = 9
Public Const ITEM_TYPE_CONSUME As Byte = 10

' Spells
Public Const SPELL_TYPE_VITAL_AFFECT As Byte = 0
Public Const SPELL_TYPE_WARP As Byte = 1
Public Const SPELL_TYPE_CUSTOM As Byte = 2

' Npc Types
Public Const NPC_TYPE_FRIENDLY As Byte = 0
Public Const NPC_TYPE_STATIONARY As Byte = 1
Public Const NPC_TYPE_ATTACK_WHEN_ATTACKED As Byte = 2
Public Const NPC_TYPE_ATTACK_ON_SIGHT As Byte = 3

Public Const START_MAP As Byte = 1
Public Const START_X As Byte = 5
Public Const START_Y As Byte = 5

' Key constants
Public Const VK_UP As Long = &H26
Public Const VK_DOWN As Long = &H28
Public Const VK_LEFT As Long = &H25
Public Const VK_RIGHT As Long = &H27
Public Const VK_SHIFT As Long = &H10
Public Const VK_RETURN As Long = &HD
Public Const VK_CONTROL As Long = &H11

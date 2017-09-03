Attribute VB_Name = "modTypes"

Public Options As OptionRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public TempNpc(1 To MAX_MAPS) As TempNpcRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Map(1 To MAX_MAPS) As MapRec
Public MapItem(1 To MAX_MAPS) As MapItemRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public Npc(1 To MAX_NPCS) As NpcRec
Public MapResource(1 To MAX_MAPS) As MapResourceRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Bank(1 To MAX_PLAYERS) As BankTabRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Chest(1 To MAX_CHESTS) As ChestRec
Public MapChest(1 To MAX_MAPS) As MapChestRec
Public TempActionMsg(1 To 255) As ActionMsgRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public MapProjectile(1 To MAX_MAPS) As MapProjectileRec
Public Game As TempGameRec

Private Type TileProjectileRec
    X As Long
    Y As Long
    XOffset As Long
    YOffset As Long
    Dir As Byte
    Picture As Long
    Range As Long
    Distance As Long
    Speed As Double
End Type

Private Type MapProjectileRec
    MapProjectile(1 To MAX_MAP_PROJECTILES) As TileProjectileRec
End Type

Private Type TempGameRec
    InBank As Boolean
    InShop As Boolean
    ChestX As Long
    ChestY As Long
    ShopNum As Long
    ShopState As Byte
End Type

Private Type OptionRec
    Debug As Boolean
    OnlineMode As Boolean
    InGame As Boolean
    IP As String
    Port As String
    InstallRuntimes As Boolean
    GameFont As String
    Username As String
    Password As String
    Music As Boolean
    Sound As Boolean
    Voices As Boolean
    DIOC As Boolean ' Display Info On Click
    FullScreen As Boolean
End Type

Private Type TempPlayerRec
    Buffer As clsBuffer
    InGame As Boolean
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    
    Moving As Byte
    CanStop As Boolean
    XOffset As Long
    YOffset As Long
    X As Long
    Y As Long
    Step As Single
    CombatTimer As Byte
    AttackTimer As Long
    
    'Graphical Stuff
    DrawCoords As Boolean
    DrawCPS As Boolean
    ' Key System
    UnlockedTile(1 To MAX_MAP_X, 1 To MAX_MAP_Y) As Boolean
    Target As Long
    HoverTarget As Long
    StunDuration As Long
End Type

Private Type NpcNumRec
    SpawnX As Long
    SpawnY As Long
    Num As Long
    Dir As Byte
    'Combat
    Vital(1 To Vitals.Vital_Count - 1) As Long
End Type

Private Type TempMapNpcNumRec
    Moving As Byte
    CanStop As Boolean
    XOffset As Long
    YOffset As Long
    X As Long
    Y As Long
    Step As Single
    Alive As Boolean
    CombatTimer As Byte
    AttackTimer As Long
    Target As Long
    StunDuration As Long
End Type

Private Type TempNpcRec
    NpcNum(1 To MAX_MAP_NPCS) As TempMapNpcNumRec
End Type

Private Type SkillRec
    level As Long
    XP As Long
End Type

Private Type PlayerInvRec
    Num As Long
    Value As Long
End Type

Private Type EquipmentRec
    Num As Long
    Value As Long
End Type

Private Type PlayerGraphicsRec
    Gender As Byte
    Skin As Byte
    BodyDir As String
    Body As Byte
    HairDir As String
    Hair As Byte
    LegsDir As String
    Legs As Byte
End Type

Private Type PlayerSpellRec
    Num As Long
    CoolDownTimer As Long
End Type

Private Type PlayerRec
    Name As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    Access As Byte
    Graphics As PlayerGraphicsRec
    Combat As SkillRec
    Vital(1 To Vitals.Vital_Count - 1) As Long
    Stat(1 To Stats.Stat_Count - 1) As Long
    Points As Long
    Equipment(1 To Equipment.Equipment_Count - 1) As EquipmentRec
    Inv(1 To MAX_INV) As PlayerInvRec
    Skill(1 To Skills.Skill_Count - 1) As SkillRec
    PlayerSpell(1 To MAX_PLAYER_SPELLS) As PlayerSpellRec
    Map As Long
    X As Long
    Y As Long
    Dir As Byte
    Stance As Byte
End Type

Private Type LayerRec
    Tileset As Byte
    X As Byte
    Y As Byte
End Type

Private Type TileRec
    Layer(1 To Layers.Layer_Count) As LayerRec
    Attribute As Long
    LongValue(1 To 4) As Long
    StringValue(1 To 4) As String * NAME_LENGTH
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Tile(0 To MAX_MAP_X, 0 To MAX_MAP_Y) As TileRec
    MapPlayer(1 To MAX_PLAYERS) As Long
    Music As String
    Moral As Byte
    UpWarp As Long
    DownWarp As Long
    LeftWarp As Long
    RightWarp As Long
    MapNpc(1 To MAX_MAP_NPCS) As NpcNumRec
End Type

Private Type DropRec
    Chance As Single
    Item As Long
    Value As Long
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    Type As Byte
    Sprite As Long
    Respawn As Long
    Vital(1 To Vitals.Vital_Count - 1) As Long
    Stat(1 To Stats.Stat_Count - 1) As Long
    Speed As Single
    Range As Byte
    AttackSpeed As Single
    Offense(1 To Combat.Combat_Count - 1) As Long
    Defense(1 To Combat.Combat_Count - 1) As Long
    XP As Long
    AttackType As Byte
    Damage As Long
    Drop(1 To MAX_DROPS) As DropRec
End Type

Private Type ProjectileRec
    Image As Long
    Speed As Double
    Range As Long
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    Stackable As Boolean
    Picture As Long
    Type As Byte
    info As String * INFO_LENGTH
    Price As Long
    IsTwoHanded As Boolean
    Paperdoll(1 To Stance.Stance_Count - 1) As Long
    CustomScript As Long
    Stat(1 To Stats.Stat_Count - 1) As Long
    StatReq(1 To Stats.Stat_Count - 1) As Long
    BltPlayerGraphics As Boolean
    addHP As Long
    addSP As Long
    GiveBack As Long
    Offense(1 To Combat.Combat_Count - 1) As Long
    Defense(1 To Combat.Combat_Count - 1) As Long
    Damage As Long
    Speed As Long
    ReqXP(1 To Skills.Skill_Count - 1) As Long
    RewXP(1 To Skills.Skill_Count - 1) As Long
    WReqXP(1 To Skills.Skill_Count - 1) As Long
    EquipmentType As Byte
    Spell As Long
    Projectile As ProjectileRec
    CombatType As Byte
End Type

Private Type MapItemLayerRec
    Num As Long
    Value As Long
    Tick As Long
    MapItemState As Byte
End Type

Private Type MapItemTileRec
    Layer(0 To MAX_MAP_ITEM_LAYERS) As MapItemLayerRec ' 0 = map item
End Type

Private Type MapItemRec
    Tile(1 To MAX_MAP_X, 1 To MAX_MAP_Y) As MapItemTileRec
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    AliveGFX As Long
    DeadGFX As Long
    RespawnRate As Long
    Health As Long
    Reward As Long
    RewardValue As Long
    RewardXP(1 To Skills.Skill_Count - 1) As Long
    RequiredXP(1 To Skills.Skill_Count - 1) As Long
    EquipmentType As Byte
End Type

Private Type MapResourceTileRec
    Num As Long
    Health As Long
    Alive As Boolean
End Type
    
Private Type MapResourceRec
    Tile(1 To MAX_MAP_X, 1 To MAX_MAP_Y) As MapResourceTileRec
End Type

Private Type BankItemRec
    Num As Long
    Value As Long
End Type

Private Type BankRec
    BankItem(1 To MAX_BANK_ITEMS) As BankItemRec
End Type

Private Type BankTabRec
    BankTab(1 To MAX_BANK_TABS) As BankRec
End Type

Private Type ShopItemCostRec
    ItemCostNum As Long
    ItemCostValue As Long
    UseUpItem As Boolean
End Type

Private Type ShopItemRec
    NumberofCosts As Byte
    StockItem As Long
    StockValue As Long
    AddXP As Boolean
    Verb As String * NAME_LENGTH
    ItemCost(1 To 10) As ShopItemCostRec
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    Picture As Long
    ShopItem(1 To MAX_SHOP_ITEMS) As ShopItemRec
End Type

Private Type ChestItemRec
    Itemnum As Long
    ItemValue As Long
    Chance As Long
End Type

Private Type ChestRec
    Name As String * NAME_LENGTH
    Picture As Long
    ChestItem(1 To MAX_CHEST_ITEMS) As ChestItemRec
End Type

Private Type MapChestRec
    Tile(1 To MAX_MAP_X, 1 To MAX_MAP_Y) As ChestRec
End Type

Private Type ActionMsgRec
    message As String
    Created As Long
    color As Long
    Scroll As Long
    X As Long
    Y As Single
    Timer As Long
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Picture As Long
    Type As Byte
    Range As Byte
    VitalAffect(1 To Vitals.Vital_Count - 1) As Long
    AOE As Byte
    CoolDown As Long
    StunDuration As Long
    Map As Long
    X As Long
    Y As Long
    Custom As Long
End Type

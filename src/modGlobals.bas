Attribute VB_Name = "modGlobals"
Public tickError As Double

Public Hovering(1 To 4) As Boolean
Public Norm(1 To 4) As String
Public Hover(1 To 4) As String

' Game Loop Variable
Public Running As Boolean
Public CPS As Long

Public CreatingCharacter As Boolean

' Text variables
Public TexthDC As Long
Public GameFont As Long

Public CurX As Double
Public CurY As Double

Public MyIndex As Long

Public FontStyle(1 To 10) As String
Public FontNumber As Byte

Public LoadedMap(1 To MAX_MAPS) As Boolean
Public OpenChest(1 To MAX_MAP_X, 1 To MAX_MAP_Y) As Boolean


' Globals
Public ChatFocus As Boolean

Public InvX As Long
Public InvY As Long
Public DragInvSlotNum As Long

Public WIMultiplier As Long
Public CurTab As Long
Public BankX As Long
Public BankY As Long
Public DragBankSlotNum As Long

Public IsLegitAdmin As Boolean
Public Const AdminPassword As String = "admin"

Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

Attribute VB_Name = "modInterfaces"
Public Editing_index As Byte
Public Editing As Byte
Public Const EDITING_BACKGROUND As Byte = 0
Public Const EDITING_LABEL As Byte = 1
Public Const EDITING_PICTUREBOX As Byte = 2

Public Const MAX_INTERFACES As Byte = 10
Public Const MAX_INTERFACE_PAGES As Byte = 10
Public Const MAX_INTERFACE_LABELS As Byte = 20
Public Const MAX_INTERFACE_PICTUREBOXES As Byte = 20
Public UserInterface(1 To MAX_INTERFACES) As InterfaceRec

Private Type ObjectRec
    Height As Integer
    Width As Integer
    Left As Integer
    Top As Integer
    Visible As Boolean
    Event As String
    EventData As String
End Type

Private Type BackgroundRec
    BackColor As Long
    PictureDir As String
    Object As ObjectRec
End Type

Private Type LabelRec
    BackColor As Long
    ForeColor As Long
    Opacity As Boolean
    Caption As String
    CaptionSize As Byte
    Object As ObjectRec
End Type

Private Type PictureboxRec
    Picture As String
    Object As ObjectRec
End Type

Private Type PageRec
    Background As BackgroundRec
    Label(1 To MAX_INTERFACE_LABELS) As LabelRec
    Picturebox(1 To MAX_INTERFACE_PICTUREBOXES) As PictureboxRec
    GoBack_Page As Byte
End Type

Private Type InterfaceRec
    Name As String * 12
    Page(1 To MAX_INTERFACE_PAGES) As PageRec
End Type


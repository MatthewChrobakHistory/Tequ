Attribute VB_Name = "modClientTCP"
Option Explicit

Private PlayerBuffer As clsBuffer

Public Sub IncomingData(ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long
Set PlayerBuffer = New clsBuffer

    frmMain.socket.GetData Buffer, vbUnicode, DataLength
    
    PlayerBuffer.WriteBytes Buffer()
    
    If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= PlayerBuffer.Length - 4
        If pLength <= PlayerBuffer.Length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If
    
        pLength = 0
        If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop
    PlayerBuffer.Trim
    DoEvents
    
End Sub

Function IsConnected() As Boolean
    
    If frmMain.socket.State = sckConnected Then
        IsConnected = True
    End If

End Function

Sub SendData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
    
    If IsConnected Then
        Set Buffer = New clsBuffer
                
        Buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        Buffer.WriteBytes Data()
        frmMain.socket.SendData Buffer.ToArray()
    End If
    
End Sub

Public Sub SendRequestLogin(ByVal name As String, ByVal Password As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestLogin
    Buffer.WriteString name
    Buffer.WriteString Password
    SendData Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Public Sub SendCreatePlayer(ByVal name As String, ByVal Password As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CCreatePlayer
    Buffer.WriteString name
    Buffer.WriteString Password
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

Public Sub SendServerMessage(ByVal Index As Long, ByVal text As String)
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CServerMessage
    Buffer.WriteLong Index
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

Public Sub SendPlayerMove()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerMove
    Buffer.WriteLong MyIndex
    Buffer.WriteLong TempPlayer(MyIndex).Moving
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

Public Sub SendPlayerStop()
Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCanStop
    Buffer.WriteLong MyIndex
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

Public Sub SendDropItem(ByVal Slot As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CDropItem
    Buffer.WriteLong MyIndex
    Buffer.WriteLong Slot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
End Sub

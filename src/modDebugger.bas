Attribute VB_Name = "modDebugger"
Public ErrOptions As ErrRec

' Exception values
Public Const Pass As Byte = 0
Public Const Catch As Byte = 1
Public Const MAX_PASS As Byte = 20

Public Const DBGR_GREEN As Long = &HFF00&
Public Const DBGR_RED As Long = &HFF&

Public LastExceptionNum As Long
Public LastExceptionTime As Long
Public LastExceptionDescription As String
Public LastExceptionLine As Long
Public LastExceptionSource As String

Public LoadedGUI As Boolean

Private Type ErrRec
    PassException(1 To MAX_PASS) As Long
    PassAll As Boolean
End Type

Public Sub PrintErrorReport(ByVal Num As String, ByVal Description As String, ByVal PlayerReport As String, ByVal Line As String, ByVal Source As String)
Dim FilePath As String
Dim F As Long
Dim Report As String

    Report = "Problem: " & Num & vbCrLf & "Desc: " & Description & vbCrLf & "Src: " & Source & vbCrLf & "Erl: " & Line & vbCrLf & vbCrLf & PlayerReport
    If Num = "CUSTOM" Then
        FilePath = App.Path & "\debug\customReport_" & Description & ".BUG_DATA"
    Else
        FilePath = App.Path & "\debug\" & Num & "_LINE" & Line & ".BUG_DATA"
    End If
    F = FreeFile
    Open FilePath For Output As #F
    Print #F, Report
    Close #F
End Sub

Public Sub HandleError(ByVal Num As String, ByVal Description As String, ByVal Line As String, ByVal Source As String)
Dim i As Long

    If LastExceptionNum = Num And timeGetTime - LastExceptionTime < 5000 Then
        Call MsgBox("A critical error occured. Make sure to report this error.", vbCritical, "Critical Error")
        frmMain.tmrError.Enabled = True
            LastExceptionNum = Num
            LastExceptionTime = timeGetTime
            LastExceptionDescription = Description
            LastExceptionLine = Line
            LastExceptionSource = Source
    Else ' it's not so critical

        ' Are we passing all?
        If ErrOptions.PassAll = True Then
            LastExceptionNum = Num
            LastExceptionTime = timeGetTime
            LastExceptionDescription = Description
            LastExceptionLine = Line
            LastExceptionSource = Source
            frmMain.tmrError.Enabled = True
            Exit Sub
        End If
            
        ' Nope
        For i = 1 To MAX_PASS
            ' Is it on the list?
            If ErrOptions.PassException(i) = Num Then
                LastExceptionNum = Num
                LastExceptionTime = timeGetTime
                LastExceptionDescription = Description
                LastExceptionLine = Line
                LastExceptionSource = Source
                frmMain.tmrError.Enabled = True
                Exit Sub
            End If
        Next
    End If

End Sub

Public Sub LoadDebugOptions()
Dim i As Long
Dim FileName As String

    ' Get the filename
    FileName = App.Path & "\debug\options.ini"
    
    ' If the file doesn't exist, save it and then it will continue as normal.
    If FileExist(FileName) = False Then
        Call SaveDebugOptions(True)
    End If
    
    With ErrOptions
        .PassAll = GetVar(FileName, "Options", "PassAll")
        For i = 1 To MAX_PASS
            .PassException(i) = GetVar(FileName, "Options", "Pass_#" & i)
        Next
    End With
End Sub

Public Sub SaveDebugOptions(Optional ByVal NewFile As Boolean = False)
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\debug\options.ini"
    
    If NewFile = True Then
        With ErrOptions
            .PassAll = False
            For i = 1 To MAX_PASS
                .PassException(i) = 0
            Next
        End With
    End If
    
    Call PutVar(FileName, "Options", "PassAll", Str(ErrOptions.PassAll))
    For i = 1 To MAX_PASS
        Call PutVar(FileName, "Options", "Pass_#" & i, Str(ErrOptions.PassException(i)))
    Next
End Sub

Public Sub InitDebugGUI()
Dim i As Long

    LoadedGUI = False

    With frmDebugger
        .Width = 3675
        .fraReport.Left = 120
        For i = 1 To MAX_PASS
            
            If ErrOptions.PassException(i) > 0 Or ErrOptions.PassException(i) < 0 Then
                .chkPass(i).Value = 1
                .chkPass(i).Caption = "Pass RTE " & ErrOptions.PassException(i)
                .chkPass(i).BackColor = DBGR_GREEN
            Else
                .chkPass(i).Value = False
                .chkPass(i).Caption = "Pass RTE: 0"
                .chkPass(i).BackColor = DBGR_RED
            End If
        Next
        If ErrOptions.PassAll = True Then
            .chkPassAll.Value = 1
            .chkPassAll.BackColor = DBGR_GREEN
            For i = 1 To MAX_PASS
                .chkPass(i).Visible = False
            Next
        Else
            .chkPassAll.Value = 0
            .chkPassAll.BackColor = DBGR_RED
            For i = 1 To MAX_PASS
                .chkPass(i).Visible = True
            Next
        End If
        .Visible = True
    End With
    
    LoadedGUI = True
End Sub

Option Explicit

Public Sub NewGame()
    InitState
    InitMines
    InitBoard
End Sub

Public Sub SetOpenMode()
    CurrentMode = MODE_OPEN
    WriteStatus
End Sub

Public Sub SetFlagMode()
    CurrentMode = MODE_FLAG
    WriteStatus
End Sub

Public Function HandleCellSelect(ByVal Target As Range) As Boolean
    Dim r As Long
    Dim c As Long

    If Not GameRunning Then Exit Function
    If Not IsBoardCell(Target) Then Exit Function

    CellToBoardPos Target, r, c

    If CurrentMode = MODE_FLAG Then
        ToggleFlag r, c
    Else
        OpenCell r, c
    End If

    HandleCellSelect = True
End Function

Public Sub OpenCell(ByVal r As Long, ByVal c As Long)
    If Not GameRunning Then Exit Sub
    If Opened(r, c) Then Exit Sub
    If Flagged(r, c) Then Exit Sub

    Opened(r, c) = True

    If Mine(r, c) Then
        GameRunning = False
        RenderCell r, c, False
        Cells(BOARD_TOP - 1, BOARD_LEFT).Value = "상태: 게임 오버"
        Exit Sub
    End If

    RenderCell r, c, False
End Sub

Public Sub ToggleFlag(ByVal r As Long, ByVal c As Long)
    If Not GameRunning Then Exit Sub
    If Opened(r, c) Then Exit Sub

    Flagged(r, c) = Not Flagged(r, c)

    RenderCell r, c, True
End Sub

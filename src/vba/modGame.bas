Option Explicit

Public Sub NewGame()
    Application.ScreenUpdating = False
    InitState
    InitBoard
    Application.ScreenUpdating = True
End Sub

Private Sub StartGame(ByVal firstR As Long, ByVal firstC As Long)
    InitMines firstR, firstC
    GameStatus = GAME_ONGOING
    GameStartTime = Now
    WriteStatus
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

    If GameStatus = GAME_WIN Or GameStatus = GAME_OVER Then Exit Function
    If Not IsBoardCell(Target) Then Exit Function

    CellToBoardPos Target, r, c

    If GameStatus = GAME_READY Then
        StartGame r, c
    End If

    If CurrentMode = MODE_FLAG Then
        ToggleFlag r, c
    Else
        OpenCell r, c
    End If

    HandleCellSelect = True
End Function

Public Sub OpenCell(ByVal r As Long, ByVal c As Long)
    If GameStatus = GAME_WIN Or GameStatus = GAME_OVER Then Exit Sub
    If Opened(r, c) Then Exit Sub
    If Flagged(r, c) Then Exit Sub

    If Mine(r, c) Then
        GameStatus = GAME_OVER
        RenderCell r, c, False
        RenderBoard
        WriteStatus
        Exit Sub
    End If

    RevealArea r, c
    
    If OpenedCount = BOARD_ROWS * BOARD_COLS - MINE_TOTAL Then
        GameStatus = GAME_WIN
        RenderBoard
        WriteStatus
    End If
End Sub

Private Sub RevealArea(ByVal startR As Long, ByVal startC As Long)
    Dim qR() As Long
    Dim qC() As Long
    Dim s As Long
    Dim e As Long
    Dim maxCells As Long

    Dim r As Long, c As Long
    Dim nr As Long, nc As Long
    Dim dr As Long, dc As Long

    If Opened(startR, startC) Then Exit Sub
    If Flagged(startR, startC) Then Exit Sub
    If Mine(startR, startC) Then Exit Sub

    Opened(startR, startC) = True
    OpenedCount = OpenedCount + 1

    If MineCount(startR, startC) <> 0 Then
        RenderCell startR, startC, False
        Exit Sub
    End If

    maxCells = BOARD_ROWS * BOARD_COLS
    ReDim qR(1 To maxCells)
    ReDim qC(1 To maxCells)
    s = 1
    e = 1

    qR(e) = startR
    qC(e) = startC

    Do While s <= e
        r = qR(s)
        c = qC(s)
        s = s + 1

        For dr = -1 To 1
            For dc = -1 To 1
                If Not (dr = 0 And dc = 0) Then
                    nr = r + dr
                    nc = c + dc

                    If IsBoardCell(Range(Cells(nr + BOARD_TOP - 1, nc + BOARD_LEFT - 1))) Then
                        If Not Opened(nr, nc) Then
                            Opened(nr, nc) = True
                            OpenedCount = OpenedCount + 1

                            If MineCount(nr, nc) = 0 Then
                                e = e + 1
                                qR(e) = nr
                                qC(e) = nc
                            End If
                        End If
                    End If
                End If
            Next dc
        Next dr
    Loop
    RenderBoard
End Sub

Public Sub ToggleFlag(ByVal r As Long, ByVal c As Long)
    If GameStatus = GAME_WIN Or GameStatus = GAME_OVER Then Exit Sub
    If Opened(r, c) Then Exit Sub

    Flagged(r, c) = Not Flagged(r, c)

    If Flagged(r, c) Then
        FlaggedCount = FlaggedCount + 1
    Else
        FlaggedCount = FlaggedCount - 1
    End If
    WriteStatus

    RenderCell r, c, True
End Sub

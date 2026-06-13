Option Explicit

Public Sub NewGame()
    If IsResetting Then Exit Sub

    On Error GoTo CleanUp

    IsResetting = True
    StopTimer
    RenderFacePressed
    WaitSeconds 0.2
    Application.ScreenUpdating = False
    InitState
    InitBoard

CleanUp:
    Application.ScreenUpdating = True
    IsResetting = False

    If Err.Number <> 0 Then
        MsgBox "새 게임을 시작하는 중 오류가 발생했습니다." & vbCrLf & Err.Description
    End If
End Sub

Private Sub StartGame(ByVal firstR As Long, ByVal firstC As Long)
    InitMines firstR, firstC
    GameStatus = GAME_ONGOING
    RenderFace
    StartTimer
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

    If Not IsBoardCell(Target) Then Exit Function

    CellToBoardPos Target, r, c

    HandleCellSelect = HandleBoardClick(r, c)
End Function

Public Function HandleBoardClick(ByVal r As Long, ByVal c As Long) As Boolean
    If IsResetting Then Exit Function
    If GameStatus = GAME_WIN Or GameStatus = GAME_OVER Then Exit Function
    If Not IsInsideBoard(r, c) Then Exit Function

    If GameStatus = GAME_READY Then
        StartGame r, c
    End If

    If Opened(r, c) Then
        ChordCell r, c
    ElseIf CurrentMode = MODE_FLAG Then
        ToggleFlag r, c
    Else
        OpenCell r, c
    End If

    HandleBoardClick = True
End Function

Public Sub TileClick()
    Dim caller As String
    Dim parts() As String
    Dim r As Long
    Dim c As Long

    caller = CStr(Application.Caller)

    parts = Split(caller, "_")

    If UBound(parts) <> 2 Then Exit Sub
    If parts(0) <> "tile" Then Exit Sub

    r = CLng(parts(1))
    c = CLng(parts(2))

    HandleBoardClick r, c
End Sub

Public Sub OpenCell(ByVal r As Long, ByVal c As Long)
    RevealCell r, c
    FinishTurn
End Sub

Public Sub ChordCell(ByVal r As Long, ByVal c As Long)
    Dim dr As Long, dc As Long
    Dim nr As Long, nc As Long

    If GameStatus <> GAME_ONGOING Then Exit Sub
    If Not Opened(r, c) Then Exit Sub
    If MineCount(r, c) <= 0 Then Exit Sub

    If CountFlags(r, c) <> MineCount(r, c) Then
        Exit Sub
    End If

    For dr = -1 To 1
        For dc = -1 To 1
            If Not (dr = 0 And dc = 0) Then
                nr = r + dr
                nc = c + dc
                If IsInsideBoard(nr, nc) Then
                    If Not Opened(nr, nc) Then
                        If Not Flagged(nr, nc) Then
                            RevealCell nr, nc
                        End If
                    End If
                End If
            End If
        Next dc
    Next dr

    FinishTurn
End Sub

Private Sub RevealCell(ByVal startR As Long, ByVal startC As Long)
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
    If Mine(startR, startC) Then
        Opened(startR, startC) = True
        If HitMineRow = 0 And HitMineCol = 0 Then
            HitMineRow = startR
            HitMineCol = startC
        End If
        Exit Sub
    End If

    MarkOpened startR, startC

    If MineCount(startR, startC) <> 0 Then
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
                    If IsInsideBoard(nr, nc) Then
                        If Not Opened(nr, nc) Then
                            MarkOpened nr, nc

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
End Sub

Private Sub FinishTurn()
    If HitMineRow > 0 And HitMineCol > 0 Then
        GameEndTick = Timer
        GameStatus = GAME_OVER
        StopTimer
        RenderCell HitMineRow, HitMineCol, False
    ElseIf OpenedCount = BOARD_ROWS * BOARD_COLS - MINE_TOTAL Then
        GameEndTick = Timer
        GameStatus = GAME_WIN
        StopTimer
    End If
    RenderFace
    RenderBoard
    WriteStatus
    WriteTimer
End Sub

Private Sub MarkOpened(ByVal r As Long, ByVal c As Long)
    If Opened(r, c) Then Exit Sub

    If Flagged(r, c) Then
        Flagged(r, c) = False
        FlaggedCount = FlaggedCount - 1
    End If

    Opened(r, c) = True
    OpenedCount = OpenedCount + 1
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

Private Sub WaitSeconds(ByVal seconds As Double)
    Dim startTick As Double
    Dim elapsed As Double

    startTick = Timer

    Do
        elapsed = Timer - startTick

        If elapsed < 0 Then
            elapsed = elapsed + 86400
        End If

        DoEvents
    Loop While elapsed < seconds
End Sub

Public Sub RenameSelectedShape()
    Dim newName As String

    newName = InputBox("새 Shape 이름을 입력하세요.")
    If newName = "" Then Exit Sub

    Selection.ShapeRange(1).Name = newName
End Sub

Public Sub CheckTiles()
    Dim names As Variant
    Dim i As Long
    Dim missing As String

    names = Array( _
        TILE_CLOSED, _
        TILE_TYPE0, _
        TILE_TYPE1, _
        TILE_TYPE2, _
        TILE_TYPE3, _
        TILE_TYPE4, _
        TILE_TYPE5, _
        TILE_TYPE6, _
        TILE_TYPE7, _
        TILE_TYPE8, _
        TILE_FLAG, _
        TILE_MINE, _
        TILE_MINE_RED, _
        TILE_MINE_WRONG _
    )

    For i = LBound(names) To UBound(names)
        If Not HasShape(AssetSheet, CStr(names(i))) Then
            missing = missing & CStr(names(i)) & vbCrLf
        End If
    Next i

    If missing = "" Then
        MsgBox "타일 이미지가 모두 확인되었습니다."
    Else
        MsgBox "누락된 타일 이미지가 있습니다:" & vbCrLf & missing
    End If
End Sub

Private Function HasShape(ByVal ws As Worksheet, ByVal shpName As String) As Boolean
    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shpName)
    On Error GoTo 0

    HasShape = Not shp Is Nothing
End Function

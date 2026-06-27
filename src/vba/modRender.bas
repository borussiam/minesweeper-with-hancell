Option Explicit
Option Private Module

Public Sub InitBoard()
    ClearSheet
    FormatCanvas
    HideSheetChrome

    InitTileLayer
    InitFaceLayout
    InitCountersLayout
    InitModeButtonsLayout

    RenderHud
    RenderModeButtons
    ParkSelection
End Sub

Private Sub ClearSheet()
    With Cells
        .ClearContents
        .ClearFormats
    End With
End Sub

Private Sub FormatCanvas()
    With Cells
        .ClearContents
        .ClearFormats
        .Interior.Color = RGB(160, 160, 160)
    End With
End Sub

Private Sub HideSheetChrome()
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    On Error GoTo 0
End Sub

Private Sub FormatBoard()
    Columns.ColumnWidth = 2.78
    Rows.RowHeight = 24

    Dim br As Range
    Set br = Cells(BOARD_TOP, BOARD_LEFT).Resize(BOARD_ROWS, BOARD_COLS)

    With br
        .NumberFormat = "0;;;@"
        .Value = CELL_CLOSED
        .Interior.Color = RGB(230, 230, 230)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Font.Name = "Segoe UI Emoji"
        .Font.Bold = True
        .FormatConditions.Delete
    End With

    ApplyCondFormatting br
End Sub

Private Sub ApplyCondFormatting(ByVal br As Range)
    Dim n As Long
    Dim fc As FormatCondition

    Set fc = br.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlBetween, _
        Formula1:="=0", _
        Formula2:="=8" _
    )
    fc.Interior.Color = RGB(198, 198, 198)

    For n = 1 To 8
        Set fc = br.FormatConditions.Add( _
            Type:=xlCellValue, _
            Operator:=xlEqual, _
            Formula1:="=" & n _
        )

        Select Case n
            Case 1
                fc.Font.Color = RGB(0, 0, 247)
            Case 2
                fc.Font.Color = RGB(0, 119, 0)
            Case 3
                fc.Font.Color = RGB(236, 0, 0)
            Case 4
                fc.Font.Color = RGB(0, 0, 128)
            Case 5
                fc.Font.Color = RGB(128, 0, 0)
            Case 6
                fc.Font.Color = RGB(0, 128, 128)
            Case 7
                fc.Font.Color = RGB(0, 0, 0)
            Case 8
                fc.Font.Color = RGB(112, 112, 112)
        End Select
    Next n
End Sub

Public Sub WriteStatus()
    Dim statusText As String
    Dim modeText As String
    Dim MinesLeft As Long

    MinesLeft = MINE_TOTAL - FlaggedCount

    Select Case GameStatus
        Case GAME_READY
            statusText = "준비"
        Case GAME_ONGOING
            statusText = "진행 중"
        Case GAME_WIN
            statusText = "승리!"
            MinesLeft = 0
        Case GAME_OVER
            statusText = "게임 오버"
    End Select

    If CurrentMode = MODE_FLAG Then
        modeText = "깃발"
    Else
        modeText = "열기"
    End If

    Cells(BOARD_TOP - 2, BOARD_LEFT).Value = "상태: " & statusText 
    Cells(BOARD_TOP - 2, BOARD_LEFT + BOARD_COLS - 2).Value = "모드: " & modeText
    Cells(BOARD_TOP - 1, BOARD_LEFT).Value = "남은 지뢰 수: " & MinesLeft
End Sub

Public Sub WriteTimer()
    Dim elapsed As Double
    Dim timeText As String

    elapsed = GetElapsedSeconds()

    If GameStatus = GAME_WIN Or GameStatus = GAME_OVER Then
        timeText = Format(elapsed, "0.000")
    Else
        timeText = CStr(Int(elapsed + 0.5))
    End If

    Cells(BOARD_TOP - 1, BOARD_LEFT + BOARD_COLS - 2).Value = "시간: " & timeText
End Sub

Public Sub RenderCell(ByVal r As Long, ByVal c As Long, isFlag as Boolean)
    Dim curr As Range
    Set curr = GameSheet.Cells(r + BOARD_TOP - 1, c + BOARD_LEFT - 1)
    If isFlag Then
        If Flagged(r, c) Then
            curr.Value = FlagText
        Else
            curr.Value = CELL_CLOSED
        End If
    ElseIf Mine(r, c) Then
        curr.Value = MineText
        curr.Interior.Color = RGB(255, 0, 0)
    ElseIf MineCount(r, c) > 0 Then
        curr.Value = MineCount(r, c)
    End If

    RenderTile r, c
End Sub

Public Sub RenderBoard()
    Dim outArr() As Variant
    Dim r as Long, c as Long
    ReDim outArr(1 To BOARD_ROWS, 1 To BOARD_COLS)

    For r = 1 To BOARD_ROWS
        For c = 1 To BOARD_COLS
            If GameStatus = GAME_WIN Then
                If Mine(r, c) Then
                    outArr(r, c) = FlagText
                Else
                    outArr(r, c) = MineCount(r, c)
                End If
            ElseIf GameStatus = GAME_OVER Then
                If Mine(r, c) Then
                    If Flagged(r, c) Then
                        outArr(r, c) = FlagText
                    Else
                        outArr(r, c) = MineText
                    End If
                ElseIf Flagged(r, c) Then
                    outArr(r, c) = FlagText
                    GameSheet.Cells(r + BOARD_TOP - 1, c + BOARD_LEFT - 1).Interior.Color = RGB(255, 160, 160)
                ElseIf Opened(r, c) Then
                    outArr(r, c) = MineCount(r, c)
                Else
                    outArr(r, c) = CELL_CLOSED
                End If
            ElseIf Not Opened(r, c) Then
                If Flagged(r, c) Then
                    outArr(r, c) = FlagText
                Else
                    outArr(r, c) = CELL_CLOSED
                End If
            ElseIf Mine(r, c) Then
                outArr(r, c) = MineText
            Else
                outArr(r, c) = MineCount(r, c)
            End If
        Next c
    Next r

    Cells(BOARD_TOP, BOARD_LEFT).Resize(BOARD_ROWS, BOARD_COLS).Value2 = outArr

    RenderAllTiles
End Sub

Public Sub RenderHud()
    RenderMineCounter
    RenderTimeCounter
    RenderFace
    InitModeButtonsLayout
End Sub

Private Sub InitCountersLayout()
    Dim mineLeft As Double
    Dim timeLeft As Double
    Dim topPos As Double

    mineLeft = BOARD_X + HUD_PADDING_X
    timeLeft = BOARD_X + BoardW() - HUD_PADDING_X - COUNTER_BG_WIDTH
    topPos = HUD_Y + (HUD_HEIGHT - COUNTER_BG_HEIGHT) / 2

    SetCounterBg COUNTER_MINE_BG, mineLeft, topPos
    SetCounterBg COUNTER_TIME_BG, timeLeft, topPos
End Sub

Private Sub SetCounterBg( _
    ByVal bgName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double _
)
    Dim beforeCount As Long
    Dim shp As Shape

    If Not HasShape(GameSheet, NUMS_BG) Then
        MsgBox "원본 숫자 배경 Shape를 찾을 수 없습니다: " & NUMS_BG
        Exit Sub
    End If

    DeleteShapeIfExists GameSheet, bgName

    beforeCount = GameSheet.Shapes.Count

    GameSheet.Shapes(NUMS_BG).Duplicate

    If GameSheet.Shapes.Count <= beforeCount Then
        MsgBox "숫자 배경 복제에 실패했습니다: " & bgName
        Exit Sub
    End If

    Set shp = GameSheet.Shapes(GameSheet.Shapes.Count)

    With shp
        .Name = bgName
        .Visible = msoTrue
        .LockAspectRatio = msoFalse

        .Left = leftPos
        .Top = topPos
        .Width = COUNTER_BG_WIDTH
        .Height = COUNTER_BG_HEIGHT

        .Line.Visible = msoFalse
        .Placement = xlFreeFloating
        .OnAction = ""
        .ZOrder msoBringToFront
    End With
End Sub

Public Sub RenderMineCounter()
    Dim minesLeft As Long
    Dim leftPos As Double
    Dim topPos As Double

    If GameStatus = GAME_WIN Then
        minesLeft = 0
    Else
        minesLeft = MINE_TOTAL - FlaggedCount
    End If

    leftPos = BOARD_X + HUD_PADDING_X
    topPos = HUD_Y + (HUD_HEIGHT - DIGIT_HEIGHT) / 2

    RenderCounter COUNTER_MINE_PREFIX, minesLeft, leftPos, topPos
End Sub

Public Sub RenderTimeCounter()
    Dim elapsed As Long
    Dim leftPos As Double
    Dim topPos As Double

    elapsed = CLng(Int(GetElapsedSeconds()))

    If elapsed > 999 Then elapsed = 999

    leftPos = BOARD_X + BoardW() - HUD_PADDING_X - CounterW()
    topPos = HUD_Y + (HUD_HEIGHT - DIGIT_HEIGHT) / 2

    RenderCounter COUNTER_TIME_PREFIX, elapsed, leftPos, topPos
End Sub

Private Sub RenderCounter( _
    ByVal prefix As String, _
    ByVal value As Long, _
    ByVal leftPos As Double, _
    ByVal topPos As Double _
)
    Dim textValue As String
    Dim i As Long
    Dim ch As String
    Dim srcName As String
    Dim digitLeft As Double
    Dim digitTop As Double

    textValue = CounterText(value)

    For i = 1 To COUNTER_DIGITS
        ch = Mid$(textValue, i, 1)
        srcName = DigitSource(ch)

        digitLeft = leftPos + DIGIT_OFFSET_X + (i - 1) * DIGIT_STEP
        digitTop = topPos + DIGIT_OFFSET_Y

        SetDigitShape _
            prefix, _
            i, _
            srcName, _
            digitLeft, _
            digitTop
    Next i
End Sub

Private Function CounterText(ByVal value As Long) As String
    Dim absValue As Long

    If value > 999 Then value = 999
    If value < -99 Then value = -99

    If value < 0 Then
        absValue = Abs(value)

        If absValue < 10 Then
            CounterText = "0-" & CStr(absValue)
        Else
            CounterText = "-" & Right$("00" & CStr(absValue), 2)
        End If
    Else
        CounterText = Right$("000" & CStr(value), 3)
    End If
End Function

Private Function DigitSource(ByVal ch As String) As String
    If ch = "-" Then
        DigitSource = "d-"
    Else
        DigitSource = "d" & ch
    End If
End Function

Private Sub SetDigitShape( _
    ByVal prefix As String, _
    ByVal idx As Long, _
    ByVal srcName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double _
)
    Dim shpName As String
    Dim beforeCount As Long
    Dim shp As Shape

    shpName = prefix & idx

    If Not HasShape(GameSheet, srcName) Then
        MsgBox "원본 숫자 Shape를 찾을 수 없습니다: " & srcName
        Exit Sub
    End If

    DeleteShapeIfExists GameSheet, shpName

    beforeCount = GameSheet.Shapes.Count

    GameSheet.Shapes(srcName).Duplicate

    If GameSheet.Shapes.Count <= beforeCount Then
        MsgBox "숫자 복제에 실패했습니다: " & srcName
        Exit Sub
    End If

    Set shp = GameSheet.Shapes(GameSheet.Shapes.Count)

    With shp
        .Name = shpName
        .Visible = msoTrue
        .LockAspectRatio = msoFalse

        .Left = leftPos
        .Top = topPos
        .Width = DIGIT_WIDTH
        .Height = DIGIT_HEIGHT

        .Line.Visible = msoFalse
        .Placement = xlFreeFloating
        .OnAction = ""
        .ZOrder msoBringToFront
    End With
End Sub

Private Sub InitFaceLayout()
    Dim leftPos As Double
    Dim topPos As Double

    leftPos = BOARD_X + (BoardW() - FACE_SIZE) / 2
    topPos = HUD_Y + (HUD_HEIGHT - FACE_SIZE) / 2

    SetupFaceShape FACE_UNPRESSED, leftPos, topPos
    SetupFaceShape FACE_PRESSED, leftPos, topPos
    SetupFaceShape FACE_WIN, leftPos, topPos
    SetupFaceShape FACE_LOSE, leftPos, topPos
End Sub

Private Sub SetupFaceShape( _
    ByVal shpName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double _
)
    Dim shp As Shape

    On Error Resume Next
    Set shp = GameSheet.Shapes(shpName)
    On Error GoTo 0

    If shp Is Nothing Then Exit Sub

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = FACE_SIZE
        .Height = FACE_SIZE
        .Placement = xlFreeFloating
        .OnAction = "NewGame"
        .Visible = msoFalse
    End With
End Sub

Public Sub RenderFace()
    Select Case GameStatus
        Case GAME_WIN
            ShowFace FACE_WIN

        Case GAME_OVER
            ShowFace FACE_LOSE

        Case Else
            ShowFace FACE_UNPRESSED
    End Select
End Sub

Public Sub RenderFacePressed()
    ShowFace FACE_PRESSED
    DoEvents
End Sub

Private Sub ShowFace(ByVal visibleName As String)
    SetFaceVisible FACE_UNPRESSED, visibleName
    SetFaceVisible FACE_PRESSED, visibleName
    SetFaceVisible FACE_WIN, visibleName
    SetFaceVisible FACE_LOSE, visibleName
End Sub

Private Sub SetFaceVisible(ByVal shpName As String, ByVal visibleName As String)
    Dim shp As Shape

    On Error Resume Next
    Set shp = GameSheet.Shapes(shpName)
    On Error GoTo 0

    If shp Is Nothing Then Exit Sub

    If shpName = visibleName Then
        shp.Visible = msoTrue
        shp.ZOrder msoBringToFront
    Else
        shp.Visible = msoFalse
    End If
End Sub

Private Sub InitModeButtonsLayout()
    Dim centerX As Double
    Dim topPos As Double
    Dim openLeft As Double
    Dim flagLeft As Double

    centerX = BOARD_X + BoardW() / 2
    topPos = BoardY() + BoardH() + MODE_BTN_TOP_GAP

    openLeft = centerX - MODE_BTN_GAP / 2 - MODE_BTN_WIDTH
    flagLeft = centerX + MODE_BTN_GAP / 2

    SetupModeButton BTN_OPEN, openLeft, topPos, MODE_BTN_WIDTH, MODE_BTN_HEIGHT, "SetOpenMode"
    SetupModeButton BTN_FLAG, flagLeft, topPos, MODE_BTN_WIDTH, MODE_BTN_HEIGHT, "SetFlagMode"
End Sub

Private Sub SetupModeButton( _
    ByVal shpName As String, _
    ByVal leftPos As Double, _
    ByVal topPos As Double, _
    ByVal w As Double, _
    ByVal h As Double, _
    ByVal macroName As String _
)
    Dim shp As Shape

    If Not HasShape(GameSheet, shpName) Then Exit Sub

    Set shp = GameSheet.Shapes(shpName)

    With shp
        .Visible = msoTrue
        .LockAspectRatio = msoFalse

        .Left = leftPos
        .Top = topPos
        .Width = w
        .Height = h

        .Placement = xlFreeFloating
        .OnAction = macroName
        .ZOrder msoBringToFront
    End With
End Sub

Public Sub RenderModeButtons()
    ApplyModeButtonStyle BTN_OPEN, (CurrentMode = MODE_OPEN)
    ApplyModeButtonStyle BTN_FLAG, (CurrentMode = MODE_FLAG)
End Sub

Private Sub ApplyModeButtonStyle(ByVal shpName As String, ByVal isActive As Boolean)
    Dim shp As Shape

    If Not HasShape(GameSheet, shpName) Then Exit Sub
    Set shp = GameSheet.Shapes(shpName)

    With shp
        .Visible = msoTrue
        .Placement = xlFreeFloating
        .Line.Visible = msoTrue
    End With

    Select Case shpName
        Case BTN_OPEN
            If isActive Then
                With shp
                    .Fill.ForeColor.RGB = RGB(68, 114, 196)
                    .Line.ForeColor.RGB = RGB(0, 0, 0)
                    .Line.Weight = 3
                End With
            Else
                With shp
                    .Fill.ForeColor.RGB = RGB(150, 175, 220)
                    .Line.ForeColor.RGB = RGB(90, 90, 90)
                    .Line.Weight = 1
                End With
            End If

        Case BTN_FLAG
            If isActive Then
                With shp
                    .Fill.ForeColor.RGB = RGB(237, 125, 49)
                    .Line.ForeColor.RGB = RGB(0, 0, 0)
                    .Line.Weight = 3
                End With
            Else
                With shp
                    .Fill.ForeColor.RGB = RGB(230, 160, 105)
                    .Line.ForeColor.RGB = RGB(90, 90, 90)
                    .Line.Weight = 1
                End With
            End If
    End Select
End Sub

Public Sub InitTileLayer()
    Dim r As Long
    Dim c As Long

    On Error GoTo CleanUp

    Application.ScreenUpdating = False

    ClearTileLayer

    For r = 1 To BOARD_ROWS
        For c = 1 To BOARD_COLS
            SetTileShape r, c, TILE_CLOSED
        Next c
    Next r

CleanUp:
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "타일 레이어 초기화 중 오류가 발생했습니다." & vbCrLf & Err.Description
    End If
End Sub

Public Sub ClearTileLayer()
    Dim i As Long
    Dim shp As Shape

    For i = GameSheet.Shapes.Count To 1 Step -1
        Set shp = GameSheet.Shapes(i)

        If Left$(shp.Name, 5) = "tile_" Then
            shp.Delete
        End If
    Next i
End Sub

Public Sub SetTileShape(ByVal r As Long, ByVal c As Long, ByVal tile As String)
    Dim shpName As String
    Dim shp As Shape
    Dim beforeCount As Long
    Dim size As Double

    If DrawnTile(r, c) = tile Then Exit Sub

    If Not HasShape(GameSheet, tile) Then
        MsgBox "원본 타일 Shape를 찾을 수 없습니다: " & tile
        Exit Sub
    End If

    shpName = TileId(r, c)

    DeleteShapeIfExists GameSheet, shpName

    beforeCount = GameSheet.Shapes.Count

    GameSheet.Shapes(tile).Duplicate

    If GameSheet.Shapes.Count <= beforeCount Then
        MsgBox "타일 복제에 실패했습니다: " & tile
        Exit Sub
    End If

    Set shp = GameSheet.Shapes(GameSheet.Shapes.Count)

    size = TILE_SIZE + TILE_OVERLAP


    With shp
        .Name = shpName
        .Visible = msoTrue
        .LockAspectRatio = msoFalse

        .Left = TileX(c) - TILE_OVERLAP / 2
        .Top = TileY(r) - TILE_OVERLAP / 2
        .Width = size
        .Height = size

        .Line.Visible = msoFalse
        .Placement = xlFreeFloating
        .OnAction = "TileClick"
        .ZOrder msoBringToFront
    End With

    DrawnTile(r, c) = tile
End Sub

Public Sub MarkChanged(ByVal r As Long, ByVal c As Long)
    If Not IsInsideBoard(r, c) Then Exit Sub
    If ChangedCount >= BOARD_ROWS * BOARD_COLS Then Exit Sub

    ChangedCount = ChangedCount + 1
    ChangedR(ChangedCount) = r
    ChangedC(ChangedCount) = c
End Sub

Public Sub RenderChangedTiles()
    Dim i As Long

    If ChangedCount = 0 Then Exit Sub

    On Error GoTo CleanUp

    Application.ScreenUpdating = False

    For i = 1 To ChangedCount
        RenderTile ChangedR(i), ChangedC(i)
    Next i

CleanUp:
    Application.ScreenUpdating = True
    ChangedCount = 0

    If Err.Number <> 0 Then
        MsgBox "변경 타일 렌더링 중 오류가 발생했습니다." & vbCrLf & Err.Description
    End If
End Sub

Public Sub ClearChangedTiles()
    ChangedCount = 0
End Sub

Public Sub RenderTile(ByVal r As Long, ByVal c As Long)
    SetTileShape r, c, GetTile(r, c)
End Sub

Public Sub RenderAllTiles()
    Dim r As Long
    Dim c As Long

    On Error GoTo CleanUp

    Application.ScreenUpdating = False

    For r = 1 To BOARD_ROWS
        For c = 1 To BOARD_COLS
            RenderTile r, c
        Next c
    Next r

CleanUp:
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "타일 렌더링 중 오류가 발생했습니다." & vbCrLf & Err.Description
    End If
End Sub

Private Function BoardY() As Double
    BoardY = HUD_Y + HUD_HEIGHT + HUD_TO_BOARD_GAP
End Function

Private Function BoardW() As Double
    BoardW = BOARD_COLS * TILE_SIZE
End Function

Private Function BoardH() As Double
    BoardH = BOARD_ROWS * TILE_SIZE
End Function

Private Function CounterW() As Double
    CounterW = COUNTER_BG_WIDTH
End Function

Private Function TileX(ByVal c As Long) As Double
    TileX = BOARD_X + (c - 1) * TILE_SIZE
End Function

Private Function TileY(ByVal r As Long) As Double
    TileY = BoardY() + (r - 1) * TILE_SIZE
End Function

Private Function TileId(ByVal r As Long, ByVal c As Long) As String
    TileId = "tile_" & r & "_" & c
End Function

Public Function GetTile(ByVal r As Long, ByVal c As Long) As String
    If GameStatus = GAME_WIN Then
        GetTile = WinTile(r, c)
        Exit Function
    End If

    If GameStatus = GAME_OVER Then
        GetTile = GameOverTile(r, c)
        Exit Function
    End If

    GetTile = NormalTile(r, c)
End Function

Private Function NormalTile(ByVal r As Long, ByVal c As Long) As String
    If Not Opened(r, c) Then
        If Flagged(r, c) Then
            NormalTile = TILE_FLAG
        Else
            NormalTile = TILE_CLOSED
        End If
        Exit Function
    End If

    NormalTile = OpenTile(r, c)
End Function

Private Function WinTile(ByVal r As Long, ByVal c As Long) As String
    If Mine(r, c) Then
        WinTile = TILE_FLAG
    Else
        WinTile = OpenTile(r, c)
    End If
End Function

Private Function GameOverTile(ByVal r As Long, ByVal c As Long) As String
    If Mine(r, c) Then
        If r = HitMineRow And c = HitMineCol Then
            GameOverTile = TILE_MINE_RED
        ElseIf Flagged(r, c) Then
            GameOverTile = TILE_FLAG
        Else
            GameOverTile = TILE_MINE
        End If
        Exit Function
    End If

    If Flagged(r, c) Then
        GameOverTile = TILE_MINE_WRONG
        Exit Function
    End If

    If Opened(r, c) Then
        GameOverTile = OpenTile(r, c)
    Else
        GameOverTile = TILE_CLOSED
    End If
End Function

Private Function OpenTile(ByVal r As Long, ByVal c As Long) As String
    Select Case MineCount(r, c)
        Case 0
            OpenTile = TILE_TYPE0
        Case 1
            OpenTile = TILE_TYPE1
        Case 2
            OpenTile = TILE_TYPE2
        Case 3
            OpenTile = TILE_TYPE3
        Case 4
            OpenTile = TILE_TYPE4
        Case 5
            OpenTile = TILE_TYPE5
        Case 6
            OpenTile = TILE_TYPE6
        Case 7
            OpenTile = TILE_TYPE7
        Case 8
            OpenTile = TILE_TYPE8
        Case Else
            OpenTile = TILE_TYPE0
    End Select
End Function

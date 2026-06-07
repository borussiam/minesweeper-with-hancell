Option Explicit
Option Private Module

Public Sub InitBoard()
	ClearSheet
	FormatBoard
	WriteStatus
	WriteTimer
	RenderFace
End Sub

Private Sub ClearSheet()
    With Cells
        .ClearContents
        .ClearFormats
    End With
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
        timeText = CStr(Int(elapsed))
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
End Sub

Public Sub RenderBoard()
	Dim outArr() As Variant
	Dim r as Long, c as Long
	Dim curr As Range
	
	ReDim outArr(1 To BOARD_ROWS, 1 To BOARD_COLS)

	For r = 1 To BOARD_ROWS
		For c = 1 To BOARD_COLS
			Set curr = GameSheet.Cells(r + BOARD_TOP - 1, c + BOARD_LEFT - 1)
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
					curr.Interior.Color = RGB(255, 160, 160)
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
    SetupFaceShape FACE_UNPRESSED, visibleName
    SetupFaceShape FACE_PRESSED, visibleName
    SetupFaceShape FACE_WIN, visibleName
    SetupFaceShape FACE_LOSE, visibleName
End Sub

Private Sub SetupFaceShape(ByVal shpName As String, ByVal visibleName As String)
    Dim shp As Shape
    Dim br As Range
    Dim leftPos As Double
    Dim topPos As Double

    On Error Resume Next
    Set shp = GameSheet.Shapes(shpName)
    On Error GoTo 0

    If shp Is Nothing Then Exit Sub

    Set br = GameSheet.Cells(BOARD_TOP, BOARD_LEFT).Resize(BOARD_ROWS, BOARD_COLS)

    leftPos = br.Left + (br.Width - FACE_SIZE) / 2
    topPos = GameSheet.Cells(BOARD_TOP - 2, BOARD_LEFT).Top + 1

    With shp
        .Left = leftPos
        .Top = topPos
        .Width = FACE_SIZE
        .Height = FACE_SIZE
        .Placement = xlMoveAndSize
        .OnAction = "NewGame"

        If shpName = visibleName Then
            .Visible = msoTrue
            .ZOrder msoBringToFront
        Else
            .Visible = msoFalse
        End If
    End With
End Sub

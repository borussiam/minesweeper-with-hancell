Option Explicit

Public Sub InitBoard()
	FormatBoard
	WriteStatus
End Sub

Private Sub FormatBoard()
	With Cells(BOARD_TOP, BOARD_LEFT).Resize(BOARD_ROWS, BOARD_COLS)
        .Interior.Color = RGB(230, 230, 230)
        .ColumnWidth = 2.78
        .RowHeight = 24
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Font.Bold = True
    End With
End Sub

Private Sub WriteStatus()
    Cells(BOARD_TOP-1, BOARD_LEFT).Value = "상태: 준비 완료"
    Cells(BOARD_TOP-1, BOARD_LEFT+BOARD_COLS).Value = "지뢰 수: "&MINE_TOTAL
End Sub

' Debugger
Public Sub RenderBoard()
	NewGame
	Cells(BOARD_TOP, BOARD_LEFT).Resize(BOARD_ROWS*2, BOARD_COLS*2).Value = ""
	Dim r as Long, c as Long
	For r = 1 To BOARD_ROWS
		For c = 1 To BOARD_COLS
			If Mine(r, c) Then
				Cells(BOARD_TOP+r-1, BOARD_LEFT+c-1).Value = "M"
			ElseIf MineCount(r, c) > 0 Then
				Cells(BOARD_TOP+r-1, BOARD_LEFT+c-1).Value = MineCount(r, c)
			End If
			Cells(BOARD_TOP+r-1, BOARD_LEFT+c-1).Interior.Color = RGB(170, 170, 170)
		Next c
	Next r
End Sub

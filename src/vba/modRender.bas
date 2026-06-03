Option Explicit
'Option Private Module

Public Sub InitBoard()
	ClearSheet
	FormatBoard
	WriteStatus
End Sub

Private Sub ClearSheet()
    With Cells
        .ClearContents
        .ClearFormats
    End With

    Columns.ColumnWidth = 2.78
    Rows.RowHeight = 24
End Sub

Private Sub FormatBoard()
	With Cells(BOARD_TOP, BOARD_LEFT).Resize(BOARD_ROWS, BOARD_COLS)
        .Interior.Color = RGB(230, 230, 230)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Font.Bold = True
    End With
End Sub

Public Sub WriteStatus()
	Dim modeText As String

    If CurrentMode = MODE_FLAG Then
        modeText = "깃발"
    Else
        modeText = "열기"
    End If

    Cells(BOARD_TOP - 1, BOARD_LEFT).Value = "상태: 진행 중 / 모드: " & modeText
    Cells(BOARD_TOP - 1, BOARD_LEFT + BOARD_COLS - 1).Value = "지뢰 수: " & MINE_TOTAL
End Sub

Public Sub RenderCell(ByVal r As Long, ByVal c As Long, isFlag as Boolean)
	If isFlag Then
		If Flagged(r, c) Then
			Cells(BOARD_TOP+r-1, BOARD_LEFT+c-1).Value = "F"
			Cells(BOARD_TOP+r-1, BOARD_LEFT+c-1).Interior.Color = RGB(238, 201, 201)
		Else
			Cells(BOARD_TOP+r-1, BOARD_LEFT+c-1).Value = ""
			Cells(BOARD_TOP+r-1, BOARD_LEFT+c-1).Interior.Color = RGB(230, 230, 230)
		End If
	ElseIf Mine(r, c) Then
		Cells(BOARD_TOP+r-1, BOARD_LEFT+c-1).Value = "M"
		Cells(BOARD_TOP+r-1, BOARD_LEFT+c-1).Interior.Color = RGB(255, 0, 0)
	ElseIf MineCount(r, c) > 0 Then
		Cells(BOARD_TOP+r-1, BOARD_LEFT+c-1).Value = MineCount(r, c)
		Cells(BOARD_TOP+r-1, BOARD_LEFT+c-1).Interior.Color = RGB(170, 170, 170)
	ElseIf MineCount(r, c) = 0 Then
		Cells(BOARD_TOP+r-1, BOARD_LEFT+c-1).Interior.Color = RGB(170, 170, 170)
	End If
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

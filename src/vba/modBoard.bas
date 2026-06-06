Option Explicit
Option Private Module

Public Sub InitMines()
	Randomize
	PlaceMines
	CountMines
End Sub

Private Sub PlaceMines()
	Dim placed As Long
	Dim r As Long
	Dim c As Long
	
	If MINE_TOTAL > BOARD_ROWS * BOARD_COLS Then
		MsgBox "지뢰 수가 전체 칸 수보다 많습니다."
		Exit Sub
	End If
	
	placed = 0
	
	Do While placed < MINE_TOTAL
		r = Int(Rnd() * BOARD_ROWS) + 1
		c = Int(Rnd() * BOARD_COLS) + 1
		
		If Not Mine(r, c) Then
			Mine(r, c) = True
			placed = placed + 1
		End If
	Loop
End Sub

Private Sub CountMines()
	Dim r As Long, c As Long
	Dim dr As Long, dc As Long
	Dim nr As Long, nc As Long
	Dim num As Long
	For r = 1 To BOARD_ROWS
		For c = 1 To BOARD_COLS
			num = 0
			If Not Mine(r, c) Then
				For dr = -1 To 1
					For dc = -1 To 1
						nr = r + dr
						nc = c + dc
						If Not (dr = 0 And dc = 0) _
						   And 1 <= nr And nr <= BOARD_ROWS _
						   And 1 <= nc And nc <= BOARD_COLS Then
							If Mine(nr, nc) Then
								num = num + 1
							End If
						End If
					Next dc
				Next dr
			End If
			MineCount(r, c) = num
		Next c
	Next r
End Sub

Public Function IsBoardCell(ByVal Target As Range) As Boolean
    Dim boardRange As Range
	Dim sr As Long, sc As Long

    If Target.Count > 1 Then
		sr = ActiveWindow.ScrollRow
		sc = ActiveWindow.ScrollColumn
		On Error GoTo CleanUp

		Application.EnableEvents = False
		Application.ScreenUpdating = False

		GameSheet.Range(PARK_CELL).Select

		ActiveWindow.ScrollRow = sr
    	ActiveWindow.ScrollColumn = sc

	CleanUp:
		Application.ScreenUpdating = True
		Application.EnableEvents = True

        IsBoardCell = False
        Exit Function
    End If

    Set boardRange = Target.Worksheet.Cells(BOARD_TOP, BOARD_LEFT).Resize(BOARD_ROWS, BOARD_COLS)

    IsBoardCell = Not Intersect(Target, boardRange) Is Nothing
End Function

Public Sub CellToBoardPos(ByVal Target As Range, ByRef r As Long, ByRef c As Long)
    r = Target.Row - BOARD_TOP + 1
    c = Target.Column - BOARD_LEFT + 1
End Sub

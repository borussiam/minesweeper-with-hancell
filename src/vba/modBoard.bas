Option Explicit

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

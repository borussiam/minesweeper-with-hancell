Option Explicit
Option Private Module

Public Sub InitMines(ByVal firstR As Long, ByVal firstC As Long)
	Randomize
	PlaceMines firstR, firstC
	CountMines
End Sub

Private Sub PlaceMines(ByVal firstR As Long, ByVal firstC As Long)
	Dim placed As Long
	Dim r As Long
	Dim c As Long
	
	If MINE_TOTAL > BOARD_ROWS * BOARD_COLS - 1 Then
		MsgBox "지뢰 수가 너무 많습니다."
		Exit Sub
	End If
	
	placed = 0
	
	Do While placed < MINE_TOTAL
		r = Int(Rnd() * BOARD_ROWS) + 1
		c = Int(Rnd() * BOARD_COLS) + 1
		
		If CurrentMode = MODE_OPEN And r = firstR And c = firstC Then
		ElseIf Not Mine(r, c) Then
			Mine(r, c) = True
			placed = placed + 1
		End If
	Loop
End Sub

Private Sub CountMines()
	Dim r As Long, c As Long
	Dim dr As Long, dc As Long
	Dim nr As Long, nc As Long
	Dim cnt As Long

	For r = 1 To BOARD_ROWS
		For c = 1 To BOARD_COLS
			cnt = 0
			If Not Mine(r, c) Then
				For dr = -1 To 1
					For dc = -1 To 1
						If Not (dr = 0 And dc = 0) Then
							nr = r + dr
							nc = c + dc
							If IsInsideBoard(nr, nc) Then
								If Mine(nr, nc) Then
									cnt = cnt + 1
								End If
							End If
						End If
					Next dc
				Next dr
			End If
			MineCount(r, c) = cnt
		Next c
	Next r
End Sub

Public Function CountFlags(ByVal r As Long, ByVal c As Long) As Long
    Dim dr As Long, dc As Long
    Dim nr As Long, nc As Long
    Dim cnt As Long

    For dr = -1 To 1
        For dc = -1 To 1
            If Not (dr = 0 And dc = 0) Then
                nr = r + dr
                nc = c + dc
                If IsInsideBoard(nr, nc) Then
                    If Flagged(nr, nc) Then
                        cnt = cnt + 1
                    End If
                End If
            End If
        Next dc
    Next dr
    CountFlags = cnt
End Function

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

Public Function IsInsideBoard(ByVal r As Long, ByVal c As Long) As Boolean
    IsInsideBoard = _
        (r >= 1 And r <= BOARD_ROWS And _
         c >= 1 And c <= BOARD_COLS)
End Function

Public Sub CellToBoardPos(ByVal Target As Range, ByRef r As Long, ByRef c As Long)
    r = Target.Row - BOARD_TOP + 1
    c = Target.Column - BOARD_LEFT + 1
End Sub

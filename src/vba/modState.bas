Option Explicit
Option Private Module

Public Mine() As Boolean
Public Opened() As Boolean
Public Flagged() As Boolean
Public MineCount() As Integer

Public GameRunning As Boolean
Public CurrentMode As Long

Public Sub InitState()
    ReDim Mine(1 To BOARD_ROWS, 1 To BOARD_COLS)
    ReDim Opened(1 To BOARD_ROWS, 1 To BOARD_COLS)
    ReDim Flagged(1 To BOARD_ROWS, 1 To BOARD_COLS)
    ReDim MineCount(1 To BOARD_ROWS, 1 To BOARD_COLS)

    GameRunning = True
    CurrentMode = MODE_OPEN
End Sub

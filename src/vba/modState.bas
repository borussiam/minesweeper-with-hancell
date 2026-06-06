Option Explicit
Option Private Module

Public Mine() As Boolean
Public Opened() As Boolean
Public Flagged() As Boolean
Public MineCount() As Integer

Public GameStatus As Long
Public OpenedCount As Long
Public FlaggedCount As Long
Public GameStartTime As Date
Public CurrentMode As Long

Public Sub InitState()
    ReDim Mine(1 To BOARD_ROWS, 1 To BOARD_COLS)
    ReDim Opened(1 To BOARD_ROWS, 1 To BOARD_COLS)
    ReDim Flagged(1 To BOARD_ROWS, 1 To BOARD_COLS)
    ReDim MineCount(1 To BOARD_ROWS, 1 To BOARD_COLS)

    GameStatus = GAME_READY
    CurrentMode = MODE_OPEN
    OpenedCount = 0
    FlaggedCount = 0
End Sub

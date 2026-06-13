Option Explicit
Option Private Module

Public Mine() As Boolean
Public Opened() As Boolean
Public Flagged() As Boolean
Public MineCount() As Integer
Public DrawnTile() As String

Public GameStatus As Long
Public OpenedCount As Long
Public FlaggedCount As Long
Public CurrentMode As Long

Public HitMineRow As Long
Public HitMineCol As Long

Public NextTimerSecond As Long
Public NextTimerTime As Date
Public TimerScheduled As Boolean

Public GameStartTick As Double
Public GameEndTick As Double

Public IsResetting As Boolean

Public Sub InitState()
    ReDim Mine(1 To BOARD_ROWS, 1 To BOARD_COLS)
    ReDim Opened(1 To BOARD_ROWS, 1 To BOARD_COLS)
    ReDim Flagged(1 To BOARD_ROWS, 1 To BOARD_COLS)
    ReDim MineCount(1 To BOARD_ROWS, 1 To BOARD_COLS)
    ReDim DrawnTile(1 To BOARD_ROWS, 1 To BOARD_COLS)

    GameStatus = GAME_READY
    CurrentMode = MODE_OPEN
    OpenedCount = 0
    FlaggedCount = 0

    HitMineRow = 0
    HitMineCol = 0

    NextTimerSecond = 0
    NextTimerTime = 0
    TimerScheduled = False

    GameStartTick = 0
    GameEndTick = 0
End Sub

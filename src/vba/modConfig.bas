Option Explicit
Option Private Module

Public Const SHEET_NAME As String = "Sheet1"

Public Const BOARD_ROWS As Long = 10
Public Const BOARD_COLS As Long = 10
Public Const MINE_TOTAL As Long = 10

Public Const BOARD_TOP As Long = 3
Public Const BOARD_LEFT As Long = 2

Public Const MODE_OPEN As Long = 1
Public Const MODE_FLAG As Long = 2

Public Const CELL_CLOSED As Long = -1

Public Const GAME_READY As Long = 0
Public Const GAME_ONGOING As Long = 1
Public Const GAME_WIN As Long = 2
Public Const GAME_OVER As Long = 3

Public Const PARK_CELL As String = "XFD1048576"

Public Function GameSheet() As Worksheet
    Set GameSheet = ThisWorkbook.Worksheets(SHEET_NAME)
End Function

Public Function FlagText() As String
    FlagText = ChrW$(&HD83D) & ChrW$(&HDEA9)   ' 🚩
End Function

Public Function MineText() As String
    MineText = ChrW(&HD83D) & ChrW(&HDCA3)   ' 💣
End Function
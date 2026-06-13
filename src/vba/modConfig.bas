Option Explicit
Option Private Module

Public Const SHEET_NAME As String = "Sheet1"
Public Const ASSET_SHEET_NAME As String = "_assets"

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

Public Const PARK_CELL As String = "EA105"

Public Const FACE_UNPRESSED As String = "face_unpressed"
Public Const FACE_PRESSED As String = "face_pressed"
Public Const FACE_WIN As String = "face_win"
Public Const FACE_LOSE As String = "face_lose"

Public Const FACE_SIZE As Double = 40

Public Const TILE_CLOSED As String = "closed"
Public Const TILE_TYPE0 As String = "type0"
Public Const TILE_TYPE1 As String = "type1"
Public Const TILE_TYPE2 As String = "type2"
Public Const TILE_TYPE3 As String = "type3"
Public Const TILE_TYPE4 As String = "type4"
Public Const TILE_TYPE5 As String = "type5"
Public Const TILE_TYPE6 As String = "type6"
Public Const TILE_TYPE7 As String = "type7"
Public Const TILE_TYPE8 As String = "type8"
Public Const TILE_FLAG As String = "flag"
Public Const TILE_MINE As String = "mine"
Public Const TILE_MINE_RED As String = "mine_red"
Public Const TILE_MINE_WRONG As String = "mine_wrong"

Public Function GameSheet() As Worksheet
    Set GameSheet = ThisWorkbook.Worksheets(SHEET_NAME)
End Function

Public Function AssetSheet() As Worksheet
    Set AssetSheet = ThisWorkbook.Worksheets(ASSET_SHEET_NAME)
End Function

Public Function FlagText() As String
    FlagText = ChrW$(&HD83D) & ChrW$(&HDEA9)   ' 🚩
End Function

Public Function MineText() As String
    MineText = ChrW(&HD83D) & ChrW(&HDCA3)   ' 💣
End Function

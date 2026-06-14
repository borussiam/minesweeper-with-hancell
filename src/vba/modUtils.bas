Option Explicit

Public Sub RenameSelectedShape()
    Dim newName As String

    newName = InputBox("새 Shape 이름을 입력하세요.")
    If newName = "" Then Exit Sub

    Selection.ShapeRange(1).Name = newName
End Sub

Public Sub CheckTiles()
    Dim names As Variant
    Dim i As Long
    Dim missing As String

    names = Array( _
        TILE_CLOSED, _
        TILE_TYPE0, _
        TILE_TYPE1, _
        TILE_TYPE2, _
        TILE_TYPE3, _
        TILE_TYPE4, _
        TILE_TYPE5, _
        TILE_TYPE6, _
        TILE_TYPE7, _
        TILE_TYPE8, _
        TILE_FLAG, _
        TILE_MINE, _
        TILE_MINE_RED, _
        TILE_MINE_WRONG _
    )

    For i = LBound(names) To UBound(names)
        If Not HasShape(AssetSheet, CStr(names(i))) Then
            missing = missing & CStr(names(i)) & vbCrLf
        End If
    Next i

    If missing = "" Then
        MsgBox "타일 이미지가 모두 확인되었습니다."
    Else
        MsgBox "누락된 타일 이미지가 있습니다:" & vbCrLf & missing
    End If
End Sub

Public Sub ArrangeTileSources()
    Dim names As Variant
    Dim i As Long
    Dim baseCell As Range
    Dim shp As Shape

    names = Array( _
        TILE_CLOSED, _
        TILE_TYPE0, _
        TILE_TYPE1, _
        TILE_TYPE2, _
        TILE_TYPE3, _
        TILE_TYPE4, _
        TILE_TYPE5, _
        TILE_TYPE6, _
        TILE_TYPE7, _
        TILE_TYPE8, _
        TILE_FLAG, _
        TILE_MINE, _
        TILE_MINE_RED, _
        TILE_MINE_WRONG _
    )

    Set baseCell = GameSheet.Range("EA1")

    For i = LBound(names) To UBound(names)
        Set shp = GameSheet.Shapes(CStr(names(i)))

        With shp
            .Left = baseCell.Offset(i, 0).Left
            .Top = baseCell.Offset(i, 0).Top
            .Width = 24
            .Height = 24
            .Placement = xlFreeFloating
            .OnAction = ""
            .Visible = msoTrue
        End With
    Next i

    MsgBox "원본 타일 이미지를 게임 시트 오른쪽 보관 구역으로 이동했습니다."
End Sub

Public Sub ArrangeCounterSources()
    Dim names As Variant
    Dim i As Long
    Dim baseCell As Range
    Dim shp As Shape

    names = Array( _
        DIGIT_0, _
        DIGIT_1, _
        DIGIT_2, _
        DIGIT_3, _
        DIGIT_4, _
        DIGIT_5, _
        DIGIT_6, _
        DIGIT_7, _
        DIGIT_8, _
        DIGIT_9, _
        DIGIT_MINUS _
    )

    Set baseCell = GameSheet.Range("EB1")

    For i = LBound(names) To UBound(names)
        Set shp = GameSheet.Shapes(CStr(names(i)))

        With shp
            .Left = baseCell.Offset(i, 0).Left
            .Top = baseCell.Offset(i * 2, 0).Top
            .Width = 20
            .Height = 37.5
            .Placement = xlFreeFloating
            .OnAction = ""
            .Visible = msoTrue
        End With
    Next i

    MsgBox "원본 숫자 이미지를 게임 시트 오른쪽 보관 구역으로 이동했습니다."
End Sub

Public Function HasShape(ByVal ws As Worksheet, ByVal shpName As String) As Boolean
    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(shpName)
    On Error GoTo 0

    HasShape = Not shp Is Nothing
End Function

Public Sub DeleteShapeIfExists(ByVal ws As Worksheet, ByVal shpName As String)
    On Error Resume Next
    ws.Shapes(shpName).Delete
    On Error GoTo 0
End Sub

Public Sub WaitSeconds(ByVal seconds As Double)
    Dim startTick As Double
    Dim elapsed As Double

    startTick = Timer

    Do
        elapsed = Timer - startTick

        If elapsed < 0 Then
            elapsed = elapsed + 86400
        End If

        DoEvents
    Loop While elapsed < seconds
End Sub

Public Sub ParkSelection()
    Dim sr As Long
    Dim sc As Long

    On Error GoTo CleanUp

    sr = ActiveWindow.ScrollRow
    sc = ActiveWindow.ScrollColumn

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    GameSheet.Activate
    GameSheet.Range(PARK_CELL).Select

    ActiveWindow.ScrollRow = sr
    ActiveWindow.ScrollColumn = sc

CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Sub TestOneTile()
    ReDim DrawnTile(1 To BOARD_ROWS, 1 To BOARD_COLS)

    SetTileShape 1, 1, TILE_CLOSED
End Sub

Public Sub TestSeveralTiles()
    ReDim DrawnTile(1 To BOARD_ROWS, 1 To BOARD_COLS)

    SetTileShape 1, 1, TILE_CLOSED
    SetTileShape 1, 2, TILE_TYPE1
    SetTileShape 1, 3, TILE_FLAG
    SetTileShape 1, 4, TILE_MINE
End Sub

Public Sub TestInitTileLayer()
    ReDim DrawnTile(1 To BOARD_ROWS, 1 To BOARD_COLS)
    InitTileLayer
End Sub

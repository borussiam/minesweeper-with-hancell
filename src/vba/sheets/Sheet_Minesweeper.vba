Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim sr As Long
    Dim sc As Long

    If Not HandleCellSelect(Target) Then Exit Sub

    sr = ActiveWindow.ScrollRow
    sc = ActiveWindow.ScrollColumn

    On Error GoTo CleanUp

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Me.Range("XFD1048576").Select

    ActiveWindow.ScrollRow = sr
    ActiveWindow.ScrollColumn = sc

CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

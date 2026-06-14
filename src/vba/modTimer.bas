Option Explicit

Public Sub StartTimer()
    StopTimer

    GameStartTick = Timer
    GameEndTick = 0
    NextTimerSecond = 1

    RenderTimeCounter
    ScheduleNextTimerTick
End Sub

Public Sub StopTimer()
    On Error Resume Next

    If TimerScheduled And NextTimerTime <> 0 Then
        Application.OnTime _
            EarliestTime:=NextTimerTime, _
            Procedure:=TimerProcName(), _
            Schedule:=False
    End If

    On Error GoTo 0

    TimerScheduled = False
    NextTimerTime = 0
End Sub

Public Sub TimerTick()
    TimerScheduled = False
    NextTimerTime = 0

    If GameStatus <> GAME_ONGOING Then Exit Sub

    RenderTimeCounter
    NextTimerSecond = NextTimerSecond + 1
    ScheduleNextTimerTick
End Sub

Private Sub ScheduleNextTimerTick()
    NextTimerTime = Now + TimeSerial(0, 0, 1)

    Application.OnTime _
        EarliestTime:=NextTimerTime, _
        Procedure:=TimerProcName(), _
        Schedule:=True

    TimerScheduled = True
End Sub

Private Function TimerProcName() As String
    TimerProcName = "TimerTick"
End Function

Public Function GetElapsedSeconds() As Double
    Dim endTick As Double
    Dim elapsed As Double

    If GameStartTick = 0 Then
        GetElapsedSeconds = 0
        Exit Function
    End If

    If GameStatus = GAME_WIN Or GameStatus = GAME_OVER Then
        endTick = GameEndTick
    Else
        endTick = Timer
    End If

    elapsed = endTick - GameStartTick

    If elapsed < 0 Then
        elapsed = elapsed + 86400
    End If

    GetElapsedSeconds = elapsed
End Function

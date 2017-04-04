Option Explicit On
'
Module Module3
    '--------------------------------------------------------------------------------
    'TA(백승효) 작성 시작. 170330
    Public Sub Main_10()
        Dim iHour As Integer                '현재 시간에 대한 변수.
        Dim iMin As Integer                 '현재 시간에 대한 변수.
        Dim iMovement As Integer            '몇번째 무브먼트 인지.
        '
        Dim nHour As Integer                '블라인드 제어시 사용하는 시간 변수.
        Dim nMin As Integer                 '블라인드 제어시 사용하는 시간 변수.
        '
        For iHour = 0 To 23
            '
            If iHour = 0 Then
                Calculate_Altitude()
                Calcuate_Azimuth()
                Calcuate_HSA()
                Calcuate_VSA()
                Calcuate_BlindPosition()
                '
                iMovement = 1
                nHour = gBlindOoperation(iMovement).Hour
                nMin = gBlindOoperation(iMovement).Minute
                '
                '
            End If
            '
            For iMin = 0 To 59
                If iHour = nHour And iMin = nMin Then
                    MoveBlind()
                    '
                    iMovement = iMovement + 1
                    nHour = gBlindOoperation(iMovement).Hour
                    nMin = gBlindOoperation(iMovement).Minute
                    '
                End If
                '
            Next iMin
            '
        Next iHour
        '
    End Sub
    '
    'TA(백승효) 작성 끝. 170330
    '--------------------------------------------------------------------------------
End Module

Option Explicit On
'
Module Module2
    '--------------------------------------------------------------------------------
    'TA(백승효) 작성 시작. 170330
    '
    Public Sub Calculate_Altitude()
        '설명 : 태양 고도각을 계산하는 sub
        'gThisMonth, gToday의 매 분에 대한 태양의 고도각을 계산하여 변수 gAltitude(23,59)에 저장.
        '담당 : team B
        '
        Dim i As Integer
        Dim j As Integer
        '
        For i = 0 To 23
            For j = 0 To 59
                'altitude 계산
                '계산된 altitude를 gAltitude(i,j)에 저장하기.
                '
            Next j
        Next i
        '
    End Sub
    '
    Public Sub Calcuate_Azimuth()
        '설명 : 태양 방위각을 계산하는 sub
        'gThisMonth, gToday의 매 분에 대한 태양의 고도각을 계산하여 변수 gAzimuth(23,59)에 저장.
        '담당 : team B
        '
    End Sub
    '
    Public Sub Calculate_SolarTime()
        '설명 : 해당 지방시의 태양시를 계산하는 sub
        'gThisMonth, gToday의 매 분에 대한 태양시를 계산하여,
        '변수 gSolarTime(, ).Hour, gSolarTime(, ).Minute에 변수에 저장.
        '담당 : team B
        '
    End Sub
    '
    Public Sub Calcuate_VSA()
        '설명 : 수직 음영각을 계산하는 sub
        'gThisMonth, gToday의 매 분에 대한 수직 음영각을 계산하여 변수 gVSA(23,59)에 저장.
        '담당 : team A
        '
    End Sub
    '
    Public Sub Calcuate_HSA()
        '설명 : 수평 음영각을 계산하는 sub
        'gThisMonth, gToday의 매 분에 대한 수평 음영각을 계산하여 변수 gHSA(23,59)에 저장.
        '담당 : team A
        '
    End Sub
    '
    Public Sub Calcuate_BlindPosition()
        '설명 : 블라인드 포지션을 계산하는 sub
        'gThisMonth, gToday의 매 분에 대한 블라인드 포지션을 계산하여 변수 gBlindPosition(23,59)에 저장.
        '담당 : team C
        '
        Dim i As Integer
        '
        'clear existing data
        For i = 1 To 50
            gBlindOoperation(i).Hour = 0
            gBlindOoperation(i).Minute = 0
            gBlindOoperation(i).Position = 0
            '
        Next i
        '
        'blind postion을 계산
        '여기서부터 코드 작성할 것.
        '
    End Sub
    '
    Public Sub BlindOperation()
        '설명 : 블라인드 운영에 대한 sub
        '운영을 위한 블라인드 포지션을 결정. gBlindOperation(50)에 대한 변수값을 결정하여 저장.
        '구해진 블라인드 포지션이 몇개인지(변수명 : gNofBlindOperationPostion)를 결정하여 저장.
        '담당 : Professor
        '
    End Sub
    '
    Public Sub MoveBlind()
        '설명 : gBlindCurrentPosition과 gBlindOperation( ).Position과 비교하여 얼마나 움직여야 되는지 거리를 계산하고,
        '그 거리만큼 가기 위해서 제어 신호를 얼마나(시간) 줘야 되는지를 결정.
        '담당 : Professor
        '
    End Sub
    '
    Public Sub CalculateBlindTime()
        '설명 : 블라인드가 완전히 내린 위치에서 완전히 올릴때까지 걸리는 시간 계산.
        '
        Move_FullyClosed()
        Move_FullyOpened0()
        '
        '
    End Sub
    '
    'TA(백승효) 작성 끝. 170330
    '--------------------------------------------------------------------------------
    ''
    Public Sub Move_FullyClosed()
        '블라인드의 현재위치에서 블라인드를 완전히 내리는 제어 sub
        '담당 : teamC 
        '
    End Sub
    '
    Public Sub Move_FullyOpened()
        '블라인드의 현재위치에서 블라인드를 완전히 올리는 제어 sub
        '담당 : teamC 
        '
    End Sub
    '
    Public Sub Move_FullyOpened0()
        '블라인드가 완전히 내린 위치에서 완전히 올릴때까지 걸리는 시간을 체크하기 위한 제어.
        'time setting을 0으로 하고,
        '블라인드를 완전히 끝까지 올리고,
        '그 때 까지 걸린 시간은 체크한다.
        '담당 : teamC     
        '
    End Sub
    '
    Public Sub Control_GlarePrevention()
        '현휘 차단을 위한 blind 제어.
        '현재 vertical position과 목표 vertical position을 비교하여 목표 vertical position으로 제어하기 위해서 어떻게 해야할지를 판단하고 제어 신호를 주는 sub
        '
    End Sub
    '
    Public Sub Control_Setting_UpperLimitCheck()
        '블라인드 셋팅 시 Upper Limit을 체크하는 sub
        '
    End Sub
    '
    Public Sub Control_Setting_LowerLimitCheck()
        '블라인드 셋팅 시 Lower Limit을 체크하는 sub
        '
    End Sub
    '
    Public Sub Point()
        'overhang 좌표 
        Po(1).X = -frame2wallin_x - 2 * frame2_x - frame2wallout_x - overhang_x
        Po(1).Y = -overhangleft_y
        Po(1).Z = window_z + 2 * frame2_z + overhang_z
        Po(2).X = -frame2wallin_x - 2 * frame2_x - frame2wallout_x - overhang_x
        Po(2).Y = overhangright_y
        Po(2).Z = window_z + 2 * frame2_z + overhang_z
        Po(3).X = -frame2wallin_x - 2 * frame2_x - frame2wallout_x
        Po(3).Y = overhangright_y
        Po(3).Z = window_z + 2 * frame2_z + overhang_z]
         Po(4).X = -frame2wallin_x - 2 * frame2_x - frame2wallout_x
        Po(4).Y = -overhangleft_y
        Po(4).Z = window_z + 2 * frame2_z + overhang_z]

    'desk 좌표

        Pd(1).X = deskloca_x
        Pd(1).Y = deskloca_y
        Pd(1).Z = -wallbottom_z
        Pd(2).X = deskloca_x
        Pd(2).Y = deskloca_y + desk_y
        Pd(2).Z = -wallbottom_z
        Pd(3).X = deskloca_x + desk_x
        Pd(3).Y = deskloca_y + desk_y
        Pd(3).Z = -wallbottom_z
        Pd(4).X = deskloca_x + desk_x
        Pd(4).Y = deskloca_y
        Pd(4).Z = -wallbottom_z
        Pd(5).X = deskloca_x
        Pd(5).Y = deskloca_y
        Pd(5).Z = desk_z - wallbottom_z
        Pd(6).X = deskloca_x
        Pd(6).Y = deskloca_y + desk_y
        Pd(6).Z = desk_z - wallbottom_z
        Pd(7).X = deskloca_x + desk_x
        Pd(7).Y = deskloca_y + desk_y
        Pd(7).Z = desk_z - wallbottom_z
        Pd(8).X = deskloca_x + desk_x
        Pd(8).Y = deskloca_y
        Pd(8).Z = desk_z - wallbottom_z

        '벽체 좌표

        Pz(1).X = 0
        Pz(1).Y = -wallleft_y
        Pz(1).Z = -wallbottom_z
        Pz(2).X = 0
        Pz(2).Y = -wallleft_y
        Pz(2).Z = room_z - wallbottom_z
        Pz(3).X = 0
        Pz(3).Y = 2 * frame2_y + window_y + wallright_y
        Pz(3).Z = room_z - wallbottom_z
        Pz(4).X = 0
        Pz(4).Y = 2 * frame2_y + window_y + wallright_y
        Pz(4).Z = -wallbottom_z

        '개구부 좌표 

        Pi(1).X = -frame2wallin_x - frame2_x - frame2wallout_x
        Pi(1).Y = 0
        Pi(1).Z = 0
        Pi(2).X = -frame2wallin_x - frame2_x - frame2wallout_x
        Pi(2).Y = frame2_y
        Pi(2).Z = 0
        Pi(3).X = 0
        Pi(3).Y = frame2_y
        Pi(3).Z = 0
        Pi(4).X = 0
        Pi(4).Y = 0
        Pi(4).Z = 0
        Pi(5).X = -frame2wallin_x - frame2_x - frame2wallout_x
        Pi(5).Y = 0
        Pi(5).Z = frame2_z
        Pi(6).X = -frame2wallin_x - frame2_x - frame2wallout_x
        Pi(6).Y = frame2_y
        Pi(6).Z = frame2_z
        Pi(7).X = 0
        Pi(7).Y = frame2_y
        Pi(7).Z = frame2_z
        Pi(8).X = 0
        Pi(8).Y = 0
        Pi(8).Z = frame2_z

        '내부 프레임 좌표

        Pa(1).X = -frame2wallin_x - frame2gapin_x - frame1_x
        Pa(1).Y = frame2_y
        Pa(1).Z = frame2_z
        Pa(2).X = -frame2wallin_x - frame2gapin_x - frame1_x
        Pa(2).Y = frame2_y + 2 * frame1_y + window_y
        Pa(2).Z = frame2_z
        Pa(3).X = -frame2wallin_x - frame2gapin_x
        Pa(3).Y = frame2_y + 2 * frame1_y + window_y
        Pa(3).Z = frame2_z
        Pa(4).X = -frame2wallin_x - frame2gapin_x
        Pa(4).Y = frame2_y
        Pa(4).Z = frame2_z
        Pa(5).X = -frame2wallin_x - frame2gapin_x - frame1_x
        Pa(5).Y = frame2_y
        Pa(5).Z = frame2_z + 2 * frame1_z + window_z
        Pa(6).X = -frame2wallin_x - frame2gapin_x - frame1_x
        Pa(6).Y = frame2_y + 2 * frame1_y + window_y
        Pa(6).Z = frame2_z + 2 * frame1_z + window_z
        Pa(7).X = -frame2wallin_x - frame2gapin_x
        Pa(7).Y = frame2_y + 2 * frame1_y + window_y
        Pa(7).Z = frame2_z + 2 * frame1_z + window_z
        Pa(8).X = -frame2wallin_x - frame2gapin_x
        Pa(8).Y = frame2_y
        Pa(8).Z = frame2_z + 2 * frame1_z + window_z

        Pm(1).X = -frame2wallin_x - frame2gapin_x - frame1_x
        Pm(1).Y = frame2_y + frame1_y
        Pm(1).Z = frame2_z + frame1_z
        Pm(2).X = -frame2wallin_x - frame2gapin_x - frame1_x
        Pm(2).Y = frame2_y + frame1_y + window_y
        Pm(2).Z = frame2_z + frame1_z
        Pm(3).X = -frame2wallin_x - frame2gapin_x
        Pm(3).Y = frame2_y + 2 * frame1_y + window_y
        Pm(3).Z = frame2_z + frame1_z
        Pm(4).X = -frame2wallin_x - frame2gapin_x
        Pm(4).Y = frame2_y + 2 * frame1_y
        Pm(4).Z = frame2_z + frame1_z
        Pm(5).X = -frame2wallin_x - frame2gapin_x - frame1_x
        Pm(5).Y = frame2_y
        Pm(5).Z = frame2_z + 2 * frame1_z + window_z
        Pm(6).X = -frame2wallin_x - frame2gapin_x - frame1_x
        Pm(6).Y = frame2_y + frame1_y + window_y
        Pm(6).Z = frame2_z + frame1_z + window_z
        Pm(7).X = -frame2wallin_x - frame2gapin_x
        Pm(7).Y = frame2_y + frame1_y + window_y
        Pm(7).Z = frame2_z + frame1_z + window_z
        Pm(8).X = -frame2wallin_x - frame2gapin_x
        Pm(8).Y = frame2_y + frame1_y
        Pm(8).Z = frame2_z + frame1_z + window_z

        '이중창 좌표

        Pe(1).X = -frame2wallin_x - frame2gapin_x - frame1gapin_x - window_x
        Pe(1).Y = frame2_y + frame1_y
        Pe(1).Z = frame2_z + frame1_z
        Pe(2).X = -frame2wallin_x - frame2gapin_x - frame1gapin_x - window_x
        Pe(2).Y = frame2_y + frame1_y + window_y
        Pe(2).Z = frame2_z + frame1_z
        Pe(3).X = -frame2wallin_x - frame2gapin_x - frame1gapin_x
        Pe(3).Y = frame2_y + frame1_y + window_y
        Pe(3).Z = frame2_z + frame1_z
        Pe(4).X = -frame2wallin_x - frame2gapin_x - frame1gapin_x
        Pe(4).Y = frame2_y + frame1_y
        Pe(4).Z = frame2_z + frame1_z
        Pe(5).X = -frame2wallin_x - frame2gapin_x - frame1gapin_x - window_x
        Pe(5).Y = frame2_y + frame1_y
        Pe(5).Z = frame2_z + frame1_z + window_z
        Pe(6).X = -frame2wallin_x - frame2gapin_x - frame1gapin_x - window_x
        Pe(6).Y = frame2_y + frame1_y + window_y
        Pe(6).Z = frame2_z + frame1_z + window_z
        Pe(7).X = -frame2wallin_x - frame2gapin_x - frame1gapin_x
        Pe(7).Y = frame2_y + frame1_y + window_y
        Pe(7).Z = frame2_z + frame1_z + window_z
        Pe(8).X = -frame2wallin_x - frame2gapin_x - frame1gapin_x - window_x
        Pe(8).Y = frame2_y + frame1_y
        Pe(8).Z = frame2_z + frame1_z + window_z

        '바깥 프레임 좌표

        Pf(1).X = -frame2wallin_x - frame2_x
        Pf(1).Y = 0
        Pf(1).Z = 0
        Pf(2).X = -frame2wallin_x - frame2_x
        Pf(2).Y = window_y + 2 * frame2_x
        Pf(2).Z = 0
        Pf(3).X = -frame2wallin_x
        Pf(3).Y = window_y + 2 * frame2_x
        Pf(3).Z = 0
        Pf(4).X = -frame2wallin_x
        Pf(4).Y = 0
        Pf(4).Z = 0
        Pf(5).X = -frame2wallin_x - frame2_x
        Pf(5).Y = 0
        Pf(5).Z = window_z + 2 * frame2_z
        Pf(6).X = -frame2wallin_x - frame2_x
        Pf(6).Y = window_y + 2 * frame2_x
        Pf(6).Z = window_z + 2 * frame2_z
        Pf(7).X = -frame2wallin_x
        Pf(7).Y = window_y + 2 * frame2_x
        Pf(7).Z = window_z + 2 * frame2_z
        Pf(8).X = -frame2wallin_x
        Pf(8).Y = 0
        Pf(8).Z = window_z + 2 * frame2_z

        Pr(1).X = -frame2wallin_x - frame2_x
        Pr(1).Y = frame2_y
        Pr(1).Z = frame2_z
        Pr(2).X = -frame2wallin_x - frame2_x
        Pr(2).Y = window_y + frame2_y
        Pr(2).Z = frame2_z
        Pr(3).X = -frame2wallin_x
        Pr(3).Y = window_y + frame2_y
        Pr(3).Z = frame2_z
        Pr(4).X = -frame2wallin_x
        Pr(4).Y = frame2_y
        Pr(4).Z = frame2_z
        Pr(5).X = -frame2wallin_x - frame2_x
        Pr(5).Y = frame2_y
        Pr(5).Z = window_z + frame2_z
        Pr(6).X = -frame2wallin_x - frame2_x
        Pr(6).Y = window_y + frame2_y
        Pr(6).Z = window_z + frame2_z
        Pr(7).X = -frame2wallin_x
        Pr(7).Y = window_y + frame2_y
        Pr(7).Z = window_z + frame2_z
        Pr(8).X = -frame2wallin_x
        Pr(8).Y = frame2_y
        Pr(8).Z = window_z + frame2_z


    End Sub
    '
End Module

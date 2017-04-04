Option Explicit On
Imports WindowsApp1
'
Module Module1
    '--------------------------------------------------------------------------------
    'TA(백승효) 작성 시작. 170330
    Public gThisMonth As Integer            '오늘의 날짜 (월)
    Public gToday As Integer                '오늘의 날짜 (일)
    '
    Public gBlindOoperation(50) As BlindOperationStructure
    '
    Structure BlindOperationStructure
        Dim Hour As Integer
        Dim Minute As Integer
        Dim Position As Single
        '
    End Structure
    '
    Public gNofBlindOperatinPosition As Integer             '운영에서 정해지는 블라인드 포지션의 개수, [-]
    '
    Public gBlindOperationInterval As Single                '블라인드 운전 간격, 단위는 분, [minute]
    Public gMinimumBlindMovement As Single                  '블라인드 최소 이동 거리, 단위는 mm, [mm]
    '
    Public gBlindCurrentPosition As Single                  '현재 블라인드의 포지션, 운영을 위해 사용되는 변수, [mm]
    '
    'TA(백승효) 작성 끝. 170330
    '--------------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------------
    'TA(백승효) 작성 시작. 170329.
    '
    Public gP(100) As PointStructure
    '주요 요소(포인트)에 대한 변수.
    '일단은 dimension을 100까지 잡았으나, Team A에서 진행하는 내용을 봐서 다시 dimensioning을 할 필요 있음.
    '
    Structure PointStructure
        Dim X As Single             '해당 포인트의 X좌표, [mm]
        Dim Y As Single             '해당 포인트의 Y좌표, [mm]
        Dim Z As Single             '해당 포인트의 Z좌표, [mm]
        '
    End Structure
    '
    Public gBlindPosition(23, 59)
    '블라인드 포지션에 대한 변수. 단위는 [mm]
    '포지션(길이)의 정의는 블라인드 상단에서부터의 거리. (이 정의에 대한 확인 필요 - teamC)
    '하루치의 블라인드 포지션을 미리 계산하여 변수에 저장.
    'dimension 의미 23 : 시간, 59 : 분
    '예를 들어, g_BlindPosition(15,30)은 오후 3시 30분의 블라인드 포지션을 의미.
    '여기서, 시간은 태양시를 기준으로 한다.
    '
    'TA(백승효) 작성 끝. 170329
    '--------------------------------------------------------------------------------
    '
    Public gAltitude(23, 59) As Single
    '태양 고도각에 대한 변수.
    '하루치의 태양 고도각 값을 미리 계산하여 변수에 저장.
    'dimension 의미 23 : 시간, 59 : 분
    '예를 들어, gAltitude(15,30)은 오후 3시 30분의 태양 고도각을 의미.
    '여기서, 시간은 태양시를 기준으로 한다. 
    '수평일때는 0도, 지표면과 수직을 이룰때는 90도.
    '
    Public gAzimuth(23, 59) As Single
    '태양 방위각에 대한 변수.
    '하루치의 태양 방위각 값을 미리 계산하여 변수에 저장.
    'dimension 의미 23 : 시간, 59 : 분
    '예를 들어, gAzimuth(15,30)은 오후 3시 30분의 태양 방위각을 의미.
    '여기서, 시간은 태양시를 기준으로 한다. 
    '방위각의 정의 (동쪽 -, 서쪽 +)
    '정남 : 0도
    '정동 : -90도
    '정서 : 90도
    '
    Public gVSA(23, 59) As Single
    '수직 음영각에 대한 변수
    '하루치의 수직 음영각 값을 미리 계산하여 변수에 저장.
    'dimension 의미 23 : 시간, 59 : 분
    '예를 들어, gVSA(15,30)은 오후 3시 30분의 수직 음영각을 의미.
    '여기서, 시간은 태양시를 기준으로 한다.
    '고도각과 마찬가지로, 수평일때는 0도, 지표면과 수직을 이룰때는 90도.
    '
    Public gHSA(23, 59) As Single
    '수평 음영각에 대한 변수
    '하루치의 수평 음영각 값을 미리 계산하여 변수에 저장.
    'dimension 의미 23 : 시간, 59 : 분
    '예를 들어, gHSA(15,30)은 오후 3시 30분의 수평 음영각을 의미.
    '태양 방위각과 같이, 동쪽은 -, 서쪽은 +
    '
    Public gSunRise_Hour As Integer                      '해당일의 일출 시각, 태양시 기준. [hour]
    '
    Public gSunRise_Minute As Single                     '해당일의 일출 시각, 태양시 기준. [minute]
    '
    Public gSUnSet_Hour As Integer                       '해당일의 일몰 시각, 태양시 기준. [hour]
    '
    Public gSunSet_Minute As Single                      '해당일의 일몰 시각, 태양시 기준, [minute]
    '
    Public gSolarTime(23, 59) As SolarTimeType
    '해당 지방시에 대응하는 태양시를 정의하기 위한 변수.
    '해당 지방시의 태양시를 알 수 있게 structure로 구성.
    '예를 들어, gSolarTime(6,30).Hour -> 지방시 오전 6시 30분 일 때 태양시의 시간을 의미.
    '예를 들어, gSolarTime(6,30).Minute -> 지방시 오전 6시 30분 일 때 태양시의 분을 의미.
    '
    Structure SolarTimeType
        Dim Hour As Integer        '해당 지방시의 태양시, [hour]
        Dim Minute As Single       '해당 지방시의 태양시, [min]
        '
    End Structure
    '
    Public gBlindHeight As Single
    '블라인드가 완전히 내려왔을 때 수직 길이, [mm]
    '
    Public gBlindTime As Single
    '블라인드가 완전히 내려가있는 상태에서 완전히 올라갔을때 까지 소요 시간, [sec]
    '


    ''박진영 작성 시작. 170403.'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '실내의 필요한 길이 정의
    '

    Public desk_x As Single, desk_y As Single, desk_z As Single, deskloca_x As Single, deskloca_y As Single
    Public window_x As Single, window_y As Single, window_z As Single, frame1_x As Single, frame1_y As Single, frame1_z As Single, frame2_x As Single, frame2_y As Single, frame2_z As Single
    Public frame1gapin_x As Single, frame1gapout_x As Single, frame2gapin_x As Single, frame2gapout_x As Single, frame2wallin_x As Single, frame2wallout_x As Single
    Public overhang_x As Single, overhangleft_y As Single, overhangright_y As Single, overhang_z As Single
    Public wallleft_y As Single, wallright_y As Single, wallbottom_z As Single, blind_x As Single, blind_y As Single, blindupper_z As Single
    Public room_z As Single
    Public face As Single, latitude As Single, longitude As Single

    '각각의 물체의 좌표를 하나의 structure 안에 생성

    Public Po(4) As PointStructure
    Public Pd(8) As PointStructure
    Public Pz(4) As PointStructure
    Public Pi(8) As PointStructure
    Public Pa(8) As PointStructure
    Public Pm(8) As PointStructure
    Public Pe(8) As PointStructure
    Public Pf(8) As PointStructure
    Public Pr(8) As PointStructure

    '박진영 작성 끝. 170403.
    Structure PointStructure
        Dim X As Single             '해당 포인트의 X좌표, [mm]
        Dim Y As Single             '해당 포인트의 Y좌표, [mm]
        Dim Z As Single             '해당 포인트의 Z좌표, [mm]
        '
    End Structure
    '박진영 작성 끝. 170403.

End Module

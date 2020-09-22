Attribute VB_Name = "modgeneral"
Option Explicit

Public Const PI                 As Double = 3.14159265358979
Public Const PI_OVER_180        As Double = PI / 180
Public Const C180_OVER_PI       As Double = 180 / PI

Public Const RIGHT_ANGLE        As Long = 90
Public Const HALF_CIRCLE        As Long = 180
Public Const FULL_CIRCLE        As Long = 360

Public Const CENTRE_SIZE        As Long = 250

Public Const RIGHT_EYE_BEARING  As Long = 40
Public Const LEFT_EYE_BEARING   As Long = 320
Public Const EYE_SCALE          As Double = 0.55
Public Const EYE_SIZE           As Long = 25

Public Const SMILE_SCALE        As Double = 0.7
Public Const SMILE_START        As Double = 200 * PI_OVER_180
Public Const SMILE_END          As Double = 340 * PI_OVER_180

Public Const DEFAULT_STALK      As Long = 50

Public BLACK                    As Long '= QBColor(0)
Public GREEN                    As Long '= QBColor(10)
Public YELLOW                   As Long '= QBColor(14)

'Forces a value into the 0-359 range
Public Function Force360(StartValue As Long)
    Dim lWhole360 As Long
    Dim lCalculator As Long
    
    lCalculator = StartValue
    If StartValue < 0 Then
        Do Until lCalculator > 0
            lCalculator = FULL_CIRCLE + lCalculator
        Loop
    Else
        lWhole360 = Int(Abs(StartValue / FULL_CIRCLE))
        lCalculator = Abs(StartValue) - lWhole360 * FULL_CIRCLE
        If StartValue < 0 Then lCalculator = FULL_CIRCLE - lCalculator
    End If
    Force360 = lCalculator
End Function

'This function give an 'X' value in a co-ordinate system, from any point, given
'the distance to the target and the bearing to the target (in degrees)
Public Function XChange(Distance As Long, Bearing As Long) As Long
    Dim lBear As Long
    
    lBear = Force360(Bearing)
    XChange = (Sin(lBear * PI_OVER_180) * Distance)
End Function

'This function give a 'Y' value in a co-ordinate system, from any point, given
'the distance to the target and the bearing to the target (in degrees)
Public Function YChange(Distance As Long, Bearing As Long) As Long
    Dim lBear As Long
    
    lBear = Force360(Bearing)
    YChange = Sin(Force360(lBear - RIGHT_ANGLE) * PI_OVER_180) * Distance
End Function

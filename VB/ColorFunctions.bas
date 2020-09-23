Attribute VB_Name = "ColorFunctions"
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum Style_
    [Longhorn]
    [Red_Hat]
    [XP_Blue]
    [XP_Green]
    [XP_Silver]
End Enum

Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Sub GetRGB(r As Integer, g As Integer, b As Integer, ByVal Color As Long)
    Dim TempValue As Long
    
    'First translate the color from a long v
    '     alue to a short value
    TranslateColor Color, 0, TempValue
    
    'Calculate the red, green, and blue valu
    '     es from the short value
    r = TempValue And &HFF&
    g = (TempValue And &HFF00&) / 2 ^ 8
    b = (TempValue And &HFF0000) / 2 ^ 16
End Sub

Public Function MakeGrey(ByVal Col As ColorConstants) As ColorConstants
    Dim r As Integer, g As Integer, b As Integer
    GetRGB r, g, b, Col 'EXTRACT COLOUR VARIABLES
    Dim x As Integer
    x = (r + g + b) / 3 'GET AVERAGE VALUE OF Each
    MakeGrey = RGB(x, x, x) 'Make the GREY colour
End Function


Public Function MakeBW(ByVal Col As ColorConstants) As ColorConstants
    Dim r As Integer, g As Integer, b As Integer
    GetRGB r, g, b, Col 'EXTRACT COLOUR VARIABLES
    Dim x As Integer
    x = (r + g + b) / 3 'GET AVERAGE VALUE OF Each


    If x < (255 / 2) Then x = 0 Else x = 255 'IF AVERAGE IS LESS THAN HALF OF MAX THEN
        'MAKE BLACK, ELSE MAKE WHITE
        MakeBW = RGB(x, x, x)
    End Function

Public Function AdjustBrightness(ByVal Color As Long, ByVal Amount As Single) As Long
    On Error Resume Next
    
    Dim r(1) As Integer, g(1) As Integer, b(1) As Integer
    
    'get red, green, and blue values
    GetRGB r(0), g(0), b(0), Color
    
    'add/subtract the amount to/from the ori
    '     ginal RGB values
    r(1) = SetBound(r(0) + Amount, 0, 255)
    g(1) = SetBound(g(0) + Amount, 0, 255)
    b(1) = SetBound(b(0) + Amount, 0, 255)
    
    'convert RGB back to Long value
    AdjustBrightness = RGB(r(1), g(1), b(1))
End Function

Private Function SetBound(ByVal Num As Single, ByVal MinNum As Single, ByVal MaxNum As Single) As Single
    If Num < MinNum Then
        SetBound = MinNum
    ElseIf Num > MaxNum Then
        SetBound = MaxNum
    Else
        SetBound = Num
    End If
End Function

Public Function InvertColor(ByVal hdc As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single)
    Dim hRect As RECT
    SetRect hRect, X1, Y1, X2, Y2
    InvertRect hdc, hRect
End Function


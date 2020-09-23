Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function GetPixel Lib "GDI32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "GDI32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Public Type UDTStar 'stars travel from offset 0,0 at an angle,
                    'speed is dependant on length!
    Speed As Long
    Angle As Single
    Len As Long
    Color As Long
End Type
Public Type UDTText
    Txt As String
    ColorDirection As Integer
    Color As Long
End Type

Public Stars() As UDTStar
Public lCentreOfPicture(2) As Long
Public lMaxLength As Long 'the shortest distance of width or height of the screen
Public Const lNumStars As Long = 200 'set number of stars here

Public Const Pi As Single = 3.14159265358979

Function GimmeX(ByVal aIn As Single, lIn As Long) As Integer
    'from an angle and length, give the x axis co'ordinate
    GimmeX = Sin(aIn * (Pi / 180)) * lIn
End Function
Function GimmeY(ByVal aIn As Single, lIn As Long) As Integer
    'from an angle and length, give the y axis co'ordinate
    GimmeY = Cos(aIn * (Pi / 180)) * lIn
End Function

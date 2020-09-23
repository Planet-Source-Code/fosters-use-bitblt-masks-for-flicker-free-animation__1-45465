VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picText 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   300
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   5
      Top             =   3660
      Width           =   915
   End
   Begin VB.PictureBox picTextM 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1140
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   6
      Top             =   3960
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Text            =   "Mike Toye"
      Top             =   3240
      Width           =   2115
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1500
      Top             =   1620
   End
   Begin VB.PictureBox picScr 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   180
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   421
      TabIndex        =   0
      Top             =   120
      Width           =   6315
   End
   Begin VB.PictureBox picBlank 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   1380
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   2
      Top             =   660
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   1020
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Text"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   3300
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub SetupText()
    'set the font attributes and size the text and mask pictures
    picText.Font = "Arial"
    picText.FontSize = 36
    
    picTextM.Font = picText.Font
    picTextM.FontSize = picText.FontSize
    
    picText.ForeColor = RGB(10, 10, 0)
    
    picText.Height = picText.TextHeight(Text1) + 10
    picText.Width = picText.TextWidth(Text1) + 10
    picTextM.Width = picText.Width
    picTextM.Height = picText.Height

End Sub

Private Function ShiftColor(ByVal Color As Long, ByVal Value As Long, Optional isXP As Boolean = False) As Long
'this function will add or remove a certain color
'quantity and return the result

Dim Red As Long
Dim Blue As Long
Dim Green As Long

    If isXP = False Then
        Blue = ((Color \ &H10000) Mod &H100) + Value
    Else
        Blue = ((Color \ &H10000) Mod &H100)
        Blue = Blue + ((Blue * Value) \ &HC0)
    End If
    
    Green = ((Color \ &H100) Mod &H100) + Value
    Red = (Color And &HFF) + Value
    
    If Red < 0 Then 'check red
        Red = 0
    ElseIf Red > 255 Then
        Red = 255
    End If
    
    If Green < 0 Then 'check green
        Green = 0
    ElseIf Green > 255 Then
        Green = 255
    End If
    
    If Blue < 0 Then 'check blue
        Blue = 0
    ElseIf Blue > 255 Then
        Blue = 255
    End If

ShiftColor = RGB(Red, Green, Blue)
End Function

Private Sub Form_Load()
    SetPicSizes
    SetupStarField
    SetupText
    DoEvents
    Timer1.Enabled = True
    
End Sub
Sub SetPicSizes()
    picBuffer.Width = picScr.Width
    picBuffer.Height = picScr.Height
    picBlank.Width = picScr.Width
    picBlank.Height = picScr.Height

End Sub
Sub SetupStarField()
Dim lX As Long

    ReDim Stars(lNumStars)
    Randomize Timer
    
    SetCentreOfPicture
    
    lMaxLength = IIf(lCentreOfPicture(0) < lCentreOfPicture(1), lCentreOfPicture(0), lCentreOfPicture(1))
    
    For lX = 0 To lNumStars - 1
        With Stars(lX)
            .Angle = Rnd * 360
            .Speed = Int(Rnd * 10) + 1
            .Len = Rnd * lMaxLength
            .Color = (255 / lMaxLength) * .Len
        End With
    Next lX
End Sub
Sub SetCentreOfPicture()
    lCentreOfPicture(0) = picBuffer.Width \ Screen.TwipsPerPixelX \ 2
    lCentreOfPicture(1) = picBuffer.Height \ Screen.TwipsPerPixelY \ 2
End Sub
Sub ClearBuffer()
    BitBlt picBuffer.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBlank.hdc, 0, 0, vbSrcCopy
End Sub
Sub PutBufferToScreen()
    BitBlt picScr.hdc, 0, 0, picScr.Width, picScr.Height, picBuffer.hdc, 0, 0, vbSrcCopy
End Sub
Sub PuStarsToBuffer()
Dim lX As Long
Dim lPosX As Long
Dim lPosY As Long

    For lX = 0 To lNumStars - 1
        lPosX = GimmeX(Stars(lX).Angle, Stars(lX).Len) + lCentreOfPicture(0)
        lPosY = GimmeY(Stars(lX).Angle, Stars(lX).Len) + lCentreOfPicture(1)
        SetPixel picBuffer.hdc, lPosX, lPosY, RGB(Stars(lX).Color, Stars(lX).Color, Stars(lX).Color)
    Next lX
    
End Sub

Sub MoveStars()
Dim lX As Long
Dim lPosX As Long
Dim lPosY As Long

    For lX = 0 To lNumStars - 1
        Stars(lX).Len = Stars(lX).Len + Stars(lX).Speed
        lPosX = GimmeX(Stars(lX).Angle, Stars(lX).Len) + lCentreOfPicture(0)
        lPosY = GimmeY(Stars(lX).Angle, Stars(lX).Len) + lCentreOfPicture(1)
        If (lPosX < 0 Or lPosX > (picBuffer.Width \ Screen.TwipsPerPixelX)) _
        Or (lPosY < 0 Or lPosY > (picBuffer.Height \ Screen.TwipsPerPixelY)) Then
            'if a star goes off screen, place it back in the middle
            Stars(lX).Len = Stars(lX).Speed
        End If
        Stars(lX).Color = (350 / lMaxLength) * Stars(lX).Len
        If Stars(lX).Color > 255 Then Stars(lX).Color = 255
    Next lX

End Sub

Private Sub Form_Resize()
Exit Sub 'remove this line and set the form to start Maximised for full screen!
picScr.Move 0, 0, Me.Width, Me.Height
SetPicSizes
DoEvents
SetupStarField
End Sub

Private Sub Text1_Change()
SetupText
End Sub

Private Sub Timer1_Timer()
    'does what it says!
    ClearBuffer
    PuStarsToBuffer
    PutTextToBuffer
    PutBufferToScreen
    MoveStars
End Sub

Sub PutTextToBuffer()
Dim lTX As Long
Dim lTY As Long
    If picText.ForeColor <> vbWhite Then
        picText.ForeColor = ShiftColor(picText.ForeColor, 2)
        picTextM.Cls
        picText.Cls
        
        picText.Print Text1
        picTextM.Print Text1
    End If
    
    lTX = (picBuffer.Width \ Screen.TwipsPerPixelX \ 2) - (picText.Width \ Screen.TwipsPerPixelX \ 2)
    lTY = (picBuffer.Height \ Screen.TwipsPerPixelY \ 2) - (picText.Height \ Screen.TwipsPerPixelY \ 2)

    BitBlt picBuffer.hdc, lTX, lTY, picTextM.Width, picTextM.Height, picTextM.hdc, 0, 0, vbSrcAnd
    BitBlt picBuffer.hdc, lTX, lTY, picText.Width, picText.Height, picText.hdc, 0, 0, vbSrcPaint
    
End Sub

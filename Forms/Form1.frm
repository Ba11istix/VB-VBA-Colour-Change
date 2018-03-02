VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll2 
      Height          =   1695
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   495
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1815
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function ColorHLSToRGB Lib "shlwapi.dll" (ByVal wHue As Long, ByVal wLuminance As Long, ByVal wSaturation As Long) As Long


Private Sub Form_Load()

Form1.Width = Form1.ScaleX(400, vbPixels, vbTwips)
Form1.Height = Form1.ScaleY(400, vbPixels, vbTwips)
Form1.ScaleMode = vbPixels
Form1.FillStyle = vbSolid
Form1.DrawWidth = 1.5
Form1.AutoRedraw = True
Form1.BackColor = RGB(0, 0, 0)

With VScroll1
    .Enabled = False
    .Top = 0
    .Left = 0
    .Height = Form1.ScaleHeight
    .Width = 13
    .Max = 240
    .Min = 0
    .value = 120
End With
With VScroll2
    .Enabled = False
    .Top = 0
    .Left = Form1.ScaleWidth - 13
    .Height = Form1.ScaleHeight
    .Width = 13
    .Max = 240
    .Min = 0
    .value = 240
End With
Form1.Show
Form1.Refresh
DoEvents
VScroll1.Enabled = True
VScroll2.Enabled = True
End Sub

Private Sub QuickWheel()
Dim AT As Single, R As Single, J As Single, x As Single, y As Single
Dim cx As Single, cy As Single
Const PI = 3.14159265358979 / 2
cx = ScaleWidth / 2
cy = ScaleHeight / 2

'upper left
For x = -125 To 0: For y = -125 To 0
       If Sqr(x * x + y * y) <= 125 Then
            If x Then
                AT = 120 - Atn(y / x) / PI * 60
            Else
                AT = 120 - Atn(y / 0.001) / PI * 60
            End If
            SetPixel Me.hdc, cx + x, cy + y, ColorHLSToRGB(AT, VScroll1, VScroll2)
       End If
Next: Next
'lower left
For x = -125 To 0: For y = 0 To 125
       If Sqr(x * x + y * y) <= 125 Then
            If x Then
                AT = 120 - Atn(y / x) / PI * 60
            Else
                AT = 120 - Atn(y / 0.001) / PI * 60
            End If
            SetPixel Me.hdc, cx + x, cy + y, ColorHLSToRGB(AT, VScroll1, VScroll2)
       End If
Next: Next
'lower right
For x = 0 To 125: For y = 0 To 125
       If Sqr(x * x + y * y) <= 125 Then
            If x Then
                AT = 240 - Atn(y / x) / PI * 60
            Else
                AT = 240 - Atn(y / 0.001) / PI * 60
            End If
'            Debug.Print Int(AT); " ";
            SetPixel Me.hdc, cx + x, cy + y, ColorHLSToRGB(AT, VScroll1, VScroll2)
       End If
Next: Next
'upper right
For x = 0 To 125: For y = -125 To 0
       If Sqr(x * x + y * y) <= 125 Then
            If x Then
                AT = 240 - Atn(y / x) / PI * 60
            Else
                AT = 240 - Atn(y / 0.001) / PI * 60
            End If
            
            SetPixel Me.hdc, cx + x, cy + y, ColorHLSToRGB(AT, VScroll1, VScroll2)
       End If
Next: Next
Form1.Refresh
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print GetPixel(Me.hdc, x, y)
End Sub

Private Sub VScroll1_Change()
QuickWheel
End Sub

Private Sub VScroll1_Scroll()
VScroll1_Change
End Sub

Private Sub VScroll2_Change()
QuickWheel
End Sub

Private Sub VScroll2_Scroll()
VScroll2_Change
End Sub

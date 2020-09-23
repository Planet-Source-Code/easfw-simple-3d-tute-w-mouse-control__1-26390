VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'fluoats@hotmail.com

'These will hold 'model' data
Dim x1(0 To 2860) As Single, y1(0 To 2860) As Single, z1(0 To 2860) As Single

'Used to center points on draw surface
Dim halfheight As Integer, halfwidth As Integer

Dim n As Integer 'all-purpose

'mouse control
Dim pressed As Boolean, gbye As Boolean
Dim xr As Single, yr As Single
Dim xr2 As Single, yr2 As Single
Dim axi As Single, ayi As Single

Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Sub Form_Load()
 Form1.BackColor = vbBlack
 Form1.ScaleMode = 3
 Randomize

'Filling coordinate array with random data
 For n = 0 To 2860 Step 1
  x1(n) = Rnd - 0.5
  y1(n) = Rnd - 0.5
  z1(n) = Rnd - 0.5
 Next n

 halfheight = Form1.ScaleHeight / 2
 halfwidth = Form1.ScaleWidth / 2
End Sub
Private Sub Form_Activate()
Do Until gbye
 Call plotRand3DPoints
 DoEvents
Loop
Unload Me
End Sub

Private Sub plotRand3DPoints()
'These store 3d projection computations based upon rotation
Dim X As Single, Y As Single, z As Single

'Will be used to erase old points just before drawing new
Static prevx(0 To 2860) As Integer, prevy(0 To 2860) As Integer

'ax looks down or up
'ay rotates like a watch display case
'az is like touch your ear to shoulder
Static az As Single, ax As Single, ay As Single

Dim vpd As Single  'vanishing-point distortion, or near-far distortion
Dim csay As Single, snay As Single 'csay = cos(ay)
Dim csaz As Single, snaz As Single
Dim csax As Single, snax As Single

Const eye = 200  'eye from screen
Const radius = 100  'point distance multiplier
Const pi As Single = 3.14159265
Const twopi = 2 * pi

Dim b As Byte 'brightness
Dim cL As Long 'color of pixel to be drawn

 
 'pre-calc sines and cosines
 
 snay = Sin(ay): csay = Cos(ay)
 snax = Sin(ax): csax = Cos(ax)
 snaz = Sin(az): csaz = Cos(az)

 'Rendering Loop
 For n = 1 To 160 Step 1
 
  'Euler transformations
  z = z1(n) * csay - x1(n) * snay
  X = x1(n) * csay + z1(n) * snay
  Y = y1(n) * csax - z * snax
  z = z * csax + y1(n) * snax
  X = X * csaz - Y * snaz
  Y = Y * csaz + X * snaz
  z = radius * z
  'vanishing-point distortion
  vpd = radius * eye / (eye - z)
  
  X = X * vpd + halfwidth 'halfwidth is horizontal center
  Y = Y * vpd + halfheight
  b = 170 + z 'more distant pixels will be darker
  cL = RGB(b, 80, b)
  
  'Erase old point before plot new
  SetPixelV Form1.hdc, prevx(n), prevy(n), vbBlack
  SetPixelV Form1.hdc, X, Y, cL
  
  'Write new point over previous
  prevx(n) = X: prevy(n) = Y
 Next n


 'Increment rotation
 ay = ay + ayi
 ax = ax + axi
 az = az + 0
 'Keep rotations inside first octave
 If ay > twopi Then ay = ay - twopi
 If ay < 0 Then ay = ay + twopi
 If ax > twopi Then ax = ax - twopi
 If ax < 0 Then ax = ax + twopi
 If az > twopi Then az = az - twopi
 If az < 0 Then az = az + twopi

End Sub
Private Sub Form_Unload(Cancel As Integer)
 gbye = True
End Sub

'Wasn't that easy?  Too bad I don't understand the rotation formulas :B
'If I did, horizontal mouse control would act 'normal'


'Everything below here is Non-3d, except axi and ayi, _
 which are my modular (under Option Explicit) rotation speed variables
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 pressed = 1
 xr = X
 ayi = 0
 yr = Y
 axi = 0
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case pressed
Case 1
 xr = xr - X
 If xr > 0 Then
  If xr > xr2 Then
   xr2 = xr
  End If
  ayi = ayi * 0.8
  If xr > 6 Then
   ayi = ayi + xr2 / 1000
  Else
   ayi = ayi + xr / 500: xr2 = xr2 - xr
  End If
 
 ElseIf xr < 0 Then
  If xr < xr2 Then
   xr2 = xr
  End If
  ayi = ayi * 0.8
  If xr < -6 Then
   ayi = ayi + xr2 / 1000
  Else
   ayi = ayi + xr / 500: xr2 = xr2 - xr
  End If
 End If
 
 yr = yr - Y
 If yr > 0 Then
  If yr > yr2 Then
   yr2 = yr
  End If
  axi = axi * 0.8
  If yr > 6 Then
   axi = axi + yr2 / 1000
  Else
   axi = axi + yr / 500: yr2 = yr2 - yr
  End If
 
 ElseIf yr < 0 Then
  If yr < yr2 Then
   yr2 = yr
  End If
  axi = axi * 0.8
  If yr < -6 Then
   axi = axi + yr2 / 1000
  Else
   axi = axi + yr / 500: yr2 = yr2 - yr
  End If
 End If
 
 xr = X
 yr = Y
End Select
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 pressed = 0
End Sub

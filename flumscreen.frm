VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Jeppe's Flum (C) 2000"
   ClientHeight    =   9405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14445
   Icon            =   "flumscreen.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   9405
   ScaleWidth      =   14445
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   7200
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   8280
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Worm (C) 2000"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r As Integer
Dim g As Integer
Dim b As Integer
Dim rto As Integer
Dim gto As Integer
Dim bto As Integer
Private Type posrad
    xpos As Integer
    ypos As Integer
    rad As Integer
    r As Integer
    g As Integer
    b As Integer
End Type
Dim Data(89) As posrad
Dim i As Integer
Dim xspeed As Integer
Dim yspeed As Integer
Dim xauto As Integer
Dim yauto As Integer
Dim antal As Integer
Private Declare Function ShowCursor Lib "User32" (ByVal Show As Integer) As Integer
Private mouseCount As Integer
'Declare the API to show/hide the mouse









Private Sub Form_KeyPress(KeyAscii As Integer)
'Unload Me
End Sub

Private Sub Form_Load()
mouseCount = 0
Form1.BackColor = RGB(0, 0, 0)

Form1.BorderStyle = 0
'Form1.WindowState = 2

ShowCursor 0
r = Rnd * 255
g = Rnd * 255
b = Rnd * 255
rto = Rnd * 255
gto = Rnd * 255
bto = Rnd * 255

moveme
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

mouseCount = mouseCount + 1
If mouseCount > 3 Then Unload Me
    'On the fourth mouse movement exit the SS.

'Unload Me
'Form1.BackColor = RGB(g, b, r)
'data(0).xpos = x + ((Rnd * 400) - 200)
'data(0).ypos = y + ((Rnd * 400) - 200)
'data(0).rad = Rnd * 50
'data(0).r = r
'data(0).g = g
'data(0).b = b
'i = 39
'Do While i > 1
'    data(i).r = data(i).r - 7
'    data(i).g = data(i).g - 7
'    data(i).b = data(i).b - 7
'    data(i).rad = data(i).rad + 5
'    i = i - 1
'Loop

'If r = rto Then
'    rto = Rnd * 255
'Else

End Sub

Private Sub Form_Unload(Cancel As Integer)
ShowCursor 1
End Sub

Private Sub Timer1_Timer()
mouseCount = 0
Call farg

'Dim j As Integer
'
'If data(0).r < 0 And data(0).g < 0 And data(0).b < 0 Then Exit Sub'
'
'i = 1
'Do While i < 39
'If data(i).r > 0 Or data(i).g > 0 Or data(i).b > 0 Then
'If data(i).r < 0 Then data(i).r = 0
'If data(i).g < 0 Then data(i).g = 0
'If data(i).b < 0 Then data(i).b = 0

''Form1.FillColor = RGB(data(i).r, data(i).g, data(i).b)
''Form1.FillStyle = 0
''Form1.Circle (data(i).xpos, data(i).ypos), data(i).rad, RGB(data(i).r, data(i).g, data(i).b)
'Call rita(data(i).xpos, data(i).ypos)


'End If
'i = i + 1
'Loop
''Form1.FillColor = RGB(0, 0, 0)
''Form1.FillStyle = 0
''Form1.Circle (data(39).xpos, data(39).ypos), data(39).rad, RGB(0, 0, 0)

'i = 39
'Do While i > 0
'    data(i).xpos = data(i - 1).xpos
'    data(i).ypos = data(i - 1).ypos
'    data(i).rad = data(i - 1).rad
'    data(i).r = data(i - 1).r
'    data(i).g = data(i - 1).g
'    data(i).b = data(i - 1).b
'    i = i - 1
'Loop
'data(0).rad = 0
'data(0).r = 0
'data(0).g = 0
'data(0).b = 0

'j = 39
'Do While j > 1
'    data(j).r = data(j).r - VScroll1.Value
'    data(j).g = data(j).g - VScroll1.Value
'    data(j).b = data(j).b - VScroll1.Value
'    data(j).rad = data(j).rad + VScroll2.Value
    
'    j = j - 1
'Loop

''Form1.FillColor = Form1.BackColor
''Form1.FillStyle = 0
''Form1.Circle (data(39).xpos, data(39).ypos), data(39).rad, Form1.BackColor

End Sub

Private Sub Timer2_Timer()

i = 0
Data(0).r = r
Data(0).g = g
Data(0).b = b
Data(0).rad = 1
Form1.FillColor = RGB(255, 255, 255)
'Form1.FillStyle = 0
'Form1.Circle (xauto, yauto), 100, RGB(255, 255, 255)
Call rita(xauto, yauto)
    xauto = xauto + xspeed + (100 - (Rnd * 200))
    yauto = yauto + yspeed + (100 - (Rnd * 200))
    If xauto > Form1.Width Then xspeed = -100
    If yauto > Form1.Height Then yspeed = -100
    If xauto < 0 Then xspeed = 100
    If yauto < 0 Then yspeed = 100
    xspeed = (100 - (Rnd * 200)) + xspeed + (0.03 * (Form1.Width / 2 - xauto))
    yspeed = (100 - (Rnd * 200)) + yspeed + (0.03 * (Form1.Height / 2 - yauto))
    
    Data(0).xpos = xauto
    Data(0).ypos = yauto
    
    
    If RGB(r, g, b) = RGB(rto, gto, bto) Then
    rto = Rnd * 255
    gto = Rnd * 255
    bto = Rnd * 255
End If
If r < rto Then
    r = r + 1
ElseIf r > rto Then
    r = r - 1
End If
'If g = gto Then
'    gto = Rnd * 255
'Else
If g < gto Then
    g = g + 1
ElseIf g > gto Then
    g = g - 1
End If
'If b = bto Then
'    bto = Rnd * 255
'Else
If b < bto Then
    b = b + 1
ElseIf b > bto Then
    b = b - 1
End If
    
    
    Call farg
    
    
    
    
    'If xauto > Form1.Width Then xspeed = -100
    'If yauto > Form1.Height Then yspeed = -100
    'If xauto < 0 Then xspeed = 100
    'If yauto < 0 Then yspeed = 100
    
End Sub


Sub moveme()



xspeed = 200 - (Rnd * 400)
yspeed = 200 - (Rnd * 400)

xauto = Rnd * Form1.Width
yauto = Rnd * Form1.Height

Timer2.Enabled = True

End Sub

Sub rita(xpos As Integer, ypos As Integer)

Form1.FillColor = RGB(Data(i).r, Data(i).g, Data(i).b)
Form1.FillStyle = 0
Form1.Circle (xpos, ypos), Data(i).rad, RGB(Data(i).r, Data(i).g, Data(i).b)

End Sub
Sub farg()

Label1.ForeColor = RGB(b, r, g)


Dim j As Integer

If Data(0).r < 0 And Data(0).g < 0 And Data(0).b < 0 Then Exit Sub

i = 1
Do While i < 89
'If data(i).r > 0 And data(i).g > 0 And data(i).b > 0 Then


If Data(i).r < 0 Then Data(i).r = 0
If Data(i).g < 0 Then Data(i).g = 0
If Data(i).b < 0 Then Data(i).b = 0


Call rita(Data(i).xpos, Data(i).ypos)

If Data(i).r = 0 And Data(i).g = 0 And Data(i).b = 0 Then Exit Do
'Form1.FillColor = RGB(data(i).r, data(i).g, data(i).b)
'Form1.FillStyle = 0
'Form1.Circle (data(i).xpos, data(i).ypos), data(i).rad, RGB(data(i).r, data(i).g, data(i).b)



'End If
i = i + 1
Loop
'Form1.FillColor = RGB(0, 0, 0)
'Form1.FillStyle = 0
'Form1.Circle (data(39).xpos, data(39).ypos), data(39).rad, RGB(0, 0, 0)

i = 89
Do While i > 0
    Data(i).xpos = Data(i - 1).xpos
    Data(i).ypos = Data(i - 1).ypos
    Data(i).rad = Data(i - 1).rad
    Data(i).r = Data(i - 1).r
    Data(i).g = Data(i - 1).g
    Data(i).b = Data(i - 1).b
    i = i - 1
Loop
'data(0).rad = 0
'data(0).r = 0
'data(0).g = 0
'data(0).b = 0

j = 89
Do While j > 1
    Data(j).r = Data(j).r - 5
    Data(j).g = Data(j).g - 5
    Data(j).b = Data(j).b - 5
    Data(j).rad = Data(j).rad + 20
    
    j = j - 1
Loop

'Form1.FillColor = Form1.BackColor
'Form1.FillStyle = 0
'Form1.Circle (data(39).xpos, data(39).ypos), data(39).rad, Form1.BackColor

End Sub

Private Sub Timer3_Timer()
Form1.BackColor = RGB(0, 0, 0)
moveme

End Sub

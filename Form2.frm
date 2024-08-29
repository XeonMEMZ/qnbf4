VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "启动窗口"
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8985
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "Form2.frx":6988A
   ScaleHeight     =   2085
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Interval        =   5
      Left            =   1320
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1800
      Top             =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "调试模式"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   360
      Left            =   2580
      TabIndex        =   0
      Top             =   1100
      Width           =   1035
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
Const LWA_COLORKEY = &H1
Dim tmd As String

Private Sub Form_Load()
Open "data\tmdu.txt" For Input As #1
Line Input #1, tmd
Close #1
tmdu = Val(tmd)
Open "themes\themes.txt" For Input As #1
Line Input #1, thm
Close #1
Open "data\aud.txt" For Input As #1
Line Input #1, aud
Close #1
Timer1.Enabled = False
Timer2.Enabled = True
toumdd = 0
Dim rtn As Long
Me.BackColor = RGB(0, 0, 0)
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 0, LWA_ALPHA
Form1.Show
Open "data\dt.txt" For Input As #1
Line Input #1, dt
Close #1
If exitproc("qnbf.exe") Then
 Label1.Visible = False
 If dt = "1" Then
  If Not exitproc("dual.exe") Then
   Call Shell("cmd /c start dual.exe")
  End If
 End If
Else
 Label1.Visible = True
End If
End Sub

Private Sub Timer1_Timer()
Dim fxs As String
Open "data\fx.txt" For Input As #1
Line Input #1, fxs
Close #1
If fxs = "5" Then
 If Weekday(Date, 2) = 6 Or Weekday(Date, 2) = 7 Then
  Form31.Show
 End If
ElseIf fxs = "6" Then
 If Weekday(Date, 2) = 7 Then
  Form31.Show
 End If
End If
Form2.Hide
Unload Me
End Sub

Private Sub Timer2_Timer()
SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), tomd, LWA_ALPHA
If tomd >= 255 Then
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 255, LWA_ALPHA
 Timer1.Enabled = True
 Timer2.Enabled = False
End If
End Sub

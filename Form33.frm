VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form33 
   BorderStyle     =   0  'None
   Caption         =   "Form33"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   LinkTopic       =   "Form33"
   Picture         =   "Form33.frx":0000
   ScaleHeight     =   3375
   ScaleMode       =   0  'User
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   3480
      Top             =   2040
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   3960
      Top             =   2040
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4920
      Top             =   2040
   End
   Begin VB.Timer Timer4 
      Interval        =   2
      Left            =   4440
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5400
      Top             =   2040
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D5D587&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "8"
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D5D587&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "888"
      Top             =   2520
      Width           =   975
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   873
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "节课上课还有"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   2610
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "距第"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2610
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "分钟"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   2610
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1665
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2000/01/01"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "星期一"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form33"
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

Private Sub Form_Load()
Text1.BackColor = RGB(collist("bg11r"), collist("bg11g"), collist("bg11b"))
Text4.BackColor = RGB(collist("bg11r"), collist("bg11g"), collist("bg11b"))
Form33.Picture = LoadPicture("themes\" & thm & "\bg1.jpg")
If aud = "1" Then
 WindowsMediaPlayer1.URL = "themes\" & thm & "\audio\qd.mp3"
End If
Dim rtn As Long
Me.BackColor = RGB(0, 0, 0)
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 255, LWA_ALPHA
toumd = 255
Timer5.Enabled = False
speed = 1
Form1.Top = Form1.Height * -1 - 100
Form1.Left = lleft
Form33.Left = Screen.Width
lleft = Screen.Width - Form33.Width
Timer3.Enabled = False
Timer4.Enabled = True
Open "data\exit.txt" For Output As #1
Print #1, "0"
Close #1
Dim fxs As String
Open "data\fx.txt" For Input As #1
Line Input #1, fxs
Close #1
If fxs = "5" Then
 If Weekday(Date, 2) = 5 Then
  fricls = 1
 Else
  fricls = 0
 End If
ElseIf fxs = "6" Then
 If Weekday(Date, 2) = 6 Then
  fricls = 1
 Else
  fricls = 0
 End If
ElseIf fxs = "7" Then
 If Weekday(Date, 2) = 7 Then
  fricls = 1
 Else
  fricls = 0
 End If
End If
If Len(Time) = 7 Then
 Label1.Caption = "0" & Time
Else
 Label1.Caption = Time
End If
Label2.Caption = Date
Label3.Caption = "星期" & " " & Weekday(Date, 2)
If Time < CDate(zxtime("zm1s")) Then
 Text1.Text = "1"
 If Hour(CDate(zxtime("zm1s")) - Time) > 0 Then
  Label4.Caption = "节课上课还有"
  Text4.Text = Minute(CDate(zxtime("zm1s")) - Time) + 61
 Else
  Label4.Caption = "节课上课还有"
  Text4.Text = Minute(CDate(zxtime("zm1s")) - Time) + 1
 End If
ElseIf CDate(zxtime("zm1s")) <= Time And Time < CDate(zxtime("zm1x")) Then
 Text1.Text = "1"
 Label4.Caption = "节课下课还有"
 Text4.Text = Minute(CDate(zxtime("zm1x")) - Time) + 1
ElseIf CDate(zxtime("zm1x")) <= Time And Time < CDate(zxtime("zm2s")) Then
 Text1.Text = "2"
 Label4.Caption = "节课上课还有"
 Text4.Text = Minute(CDate(zxtime("zm2s")) - Time) + 1
ElseIf CDate(zxtime("zm2s")) <= Time And Time < CDate(zxtime("zm2x")) Then
 Text1.Text = "2"
 Label4.Caption = "节课下课还有"
 Text4.Text = Minute(CDate(zxtime("zm2x")) - Time) + 1
ElseIf CDate(zxtime("zm2x")) <= Time And Time < CDate(zxtime("zm3s")) Then
 Text1.Text = "3"
 Label4.Caption = "节课上课还有"
 Text4.Text = Minute(CDate(zxtime("zm3s")) - Time) + 1
ElseIf CDate(zxtime("zm3s")) <= Time And Time < CDate(zxtime("zm3x")) Then
 Text1.Text = "3"
 Label4.Caption = "节课下课还有"
 Text4.Text = Minute(CDate(zxtime("zm3x")) - Time) + 1
ElseIf CDate(zxtime("zm3x")) <= Time And Time < CDate(zxtime("zm4s")) Then
 Text1.Text = "4"
 Label4.Caption = "节课上课还有"
 Text4.Text = Minute(CDate(zxtime("zm4s")) - Time) + 1
ElseIf CDate(zxtime("zm4s")) <= Time And Time < CDate(zxtime("zm4x")) Then
 Text1.Text = "4"
 Label4.Caption = "节课下课还有"
 Text4.Text = Minute(CDate(zxtime("zm4x")) - Time) + 1
Else
 If fricls = 0 Then
  If CDate(zxtime("zm4x")) <= Time And Time < CDate(zxtime("zm5s")) Then
   Text1.Text = "5"
   If Hour(CDate(zxtime("zm5s")) - Time) > 0 Then
    Label4.Caption = "节课上课还有"
    Text4.Text = Minute(CDate(zxtime("zm5s")) - Time) + 61
   Else
    Label4.Caption = "节课上课还有"
    Text4.Text = Minute(CDate(zxtime("zm5s")) - Time) + 1
   End If
  ElseIf CDate(zxtime("zm5s")) <= Time And Time < CDate(zxtime("zm5x")) Then
   Text1.Text = "5"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm5x")) - Time) + 1
  ElseIf CDate(zxtime("zm5x")) <= Time And Time < CDate(zxtime("zm6s")) Then
   Text1.Text = "6"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm6s")) - Time) + 1
  ElseIf CDate(zxtime("zm6s")) <= Time And Time < CDate(zxtime("zm6x")) Then
   Text1.Text = "6"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm6x")) - Time) + 1
  ElseIf CDate(zxtime("zm6x")) <= Time And Time < CDate(zxtime("zm7s")) Then
   Text1.Text = "7"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm7s")) - Time) + 1
  ElseIf CDate(zxtime("zm7s")) <= Time And Time < CDate(zxtime("zm7x")) Then
   Text1.Text = "7"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm7x")) - Time) + 1
  ElseIf CDate(zxtime("zm7x")) <= Time And Time < CDate(zxtime("zm8s")) Then
   Text1.Text = "8"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm8s")) - Time) + 1
  ElseIf CDate(zxtime("zm8s")) <= Time And Time < CDate(zxtime("zm8x")) Then
   Text1.Text = "8"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm8x")) - Time) + 1
  ElseIf CDate(zxtime("zm8x")) < Time Then
   Text1.Text = "8"
   Label4.Caption = "节课下课还有"
   Text4.Text = "0"
  End If
 Else
  If CDate(zxtime("zm4s")) <= Time And Time < CDate(zxtime("zf5s")) Then
   Text1.Text = "5"
   If Hour(CDate(zxtime("zf5s")) - Time) > 0 Then
    Label4.Caption = "节课上课还有"
    Text4.Text = Minute(CDate(zxtime("zf5s")) - Time) + 61
   Else
    Label4.Caption = "节课上课还有"
    Text4.Text = Minute(CDate(zxtime("zf5s")) - Time) + 1
   End If
  ElseIf CDate(zxtime("zf5s")) <= Time And Time < CDate(zxtime("zf5x")) Then
   Text1.Text = "5"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zf5x")) - Time) + 1
  ElseIf CDate(zxtime("zf5x")) <= Time And Time < CDate(zxtime("zf6s")) Then
   Text1.Text = "6"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zf6s")) - Time) + 1
  ElseIf CDate(zxtime("zf6s")) <= Time And Time < CDate(zxtime("zf6x")) Then
   Text1.Text = "6"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zf6x")) - Time) + 1
  ElseIf CDate(zxtime("zf6x")) < Time Then
   Text1.Text = "6"
   Label4.Caption = "节课下课还有"
   Text4.Text = "0"
  End If
 End If
End If
End Sub

Private Sub Timer1_Timer()
If Len(Time) = 7 Then
 Label1.Caption = "0" & Time
Else
 Label1.Caption = Time
End If
Label2.Caption = Date
If Weekday(Date, 2) = 1 Then
 Label3.Caption = "星期一"
ElseIf Weekday(Date, 2) = 2 Then
 Label3.Caption = "星期二"
ElseIf Weekday(Date, 2) = 3 Then
 Label3.Caption = "星期三"
ElseIf Weekday(Date, 2) = 4 Then
 Label3.Caption = "星期四"
ElseIf Weekday(Date, 2) = 5 Then
 Label3.Caption = "星期五"
ElseIf Weekday(Date, 2) = 6 Then
 Label3.Caption = "星期六"
ElseIf Weekday(Date, 2) = 7 Then
 Label3.Caption = "星期日"
End If
End Sub

Private Sub Timer2_Timer()
If Time < CDate(zxtime("zm1s")) Then
 Text1.Text = "1"
 If Hour(CDate(zxtime("zm1s")) - Time) > 0 Then
  Label4.Caption = "节课上课还有"
  Text4.Text = Minute(CDate(zxtime("zm1s")) - Time) + 61
 Else
  Label4.Caption = "节课上课还有"
  Text4.Text = Minute(CDate(zxtime("zm1s")) - Time) + 1
 End If
ElseIf CDate(zxtime("zm1s")) <= Time And Time < CDate(zxtime("zm1x")) Then
 Text1.Text = "1"
 Label4.Caption = "节课下课还有"
 Text4.Text = Minute(CDate(zxtime("zm1x")) - Time) + 1
ElseIf CDate(zxtime("zm1x")) <= Time And Time < CDate(zxtime("zm2s")) Then
 Text1.Text = "2"
 Label4.Caption = "节课上课还有"
 Text4.Text = Minute(CDate(zxtime("zm2s")) - Time) + 1
ElseIf CDate(zxtime("zm2s")) <= Time And Time < CDate(zxtime("zm2x")) Then
 Text1.Text = "2"
 Label4.Caption = "节课下课还有"
 Text4.Text = Minute(CDate(zxtime("zm2x")) - Time) + 1
ElseIf CDate(zxtime("zm2x")) <= Time And Time < CDate(zxtime("zm3s")) Then
 Text1.Text = "3"
 Label4.Caption = "节课上课还有"
 Text4.Text = Minute(CDate(zxtime("zm3s")) - Time) + 1
ElseIf CDate(zxtime("zm3s")) <= Time And Time < CDate(zxtime("zm3x")) Then
 Text1.Text = "3"
 Label4.Caption = "节课下课还有"
 Text4.Text = Minute(CDate(zxtime("zm3x")) - Time) + 1
ElseIf CDate(zxtime("zm3x")) <= Time And Time < CDate(zxtime("zm4s")) Then
 Text1.Text = "4"
 Label4.Caption = "节课上课还有"
 Text4.Text = Minute(CDate(zxtime("zm4s")) - Time) + 1
ElseIf CDate(zxtime("zm4s")) <= Time And Time < CDate(zxtime("zm4x")) Then
 Text1.Text = "4"
 Label4.Caption = "节课下课还有"
 Text4.Text = Minute(CDate(zxtime("zm4x")) - Time) + 1
Else
 If fricls = 0 Then
  If CDate(zxtime("zm4x")) <= Time And Time < CDate(zxtime("zm5s")) Then
   Text1.Text = "5"
   If Hour(CDate(zxtime("zm5s")) - Time) > 0 Then
    Label4.Caption = "节课上课还有"
    Text4.Text = Minute(CDate(zxtime("zm5s")) - Time) + 61
   Else
    Label4.Caption = "节课上课还有"
    Text4.Text = Minute(CDate(zxtime("zm5s")) - Time) + 1
   End If
  ElseIf CDate(zxtime("zm5s")) <= Time And Time < CDate(zxtime("zm5x")) Then
   Text1.Text = "5"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm5x")) - Time) + 1
  ElseIf CDate(zxtime("zm5x")) <= Time And Time < CDate(zxtime("zm6s")) Then
   Text1.Text = "6"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm6s")) - Time) + 1
  ElseIf CDate(zxtime("zm6s")) <= Time And Time < CDate(zxtime("zm6x")) Then
   Text1.Text = "6"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm6x")) - Time) + 1
  ElseIf CDate(zxtime("zm6x")) <= Time And Time < CDate(zxtime("zm7s")) Then
   Text1.Text = "7"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm7s")) - Time) + 1
  ElseIf CDate(zxtime("zm7s")) <= Time And Time < CDate(zxtime("zm7x")) Then
   Text1.Text = "7"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm7x")) - Time) + 1
  ElseIf CDate(zxtime("zm7x")) <= Time And Time < CDate(zxtime("zm8s")) Then
   Text1.Text = "8"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zm8s")) - Time) + 1
  ElseIf CDate(zxtime("zm8s")) <= Time And Time < CDate(zxtime("zm8x")) Then
   Text1.Text = "8"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zm8x")) - Time) + 1
  ElseIf CDate(zxtime("zm8x")) < Time Then
   Text1.Text = "8"
   Label4.Caption = "节课下课还有"
   Text4.Text = "0"
  End If
 Else
  If CDate(zxtime("zm4s")) <= Time And Time < CDate(zxtime("zf5s")) Then
   Text1.Text = "5"
   If Hour(CDate(zxtime("zf5s")) - Time) > 0 Then
    Label4.Caption = "节课上课还有"
    Text4.Text = Minute(CDate(zxtime("zf5s")) - Time) + 61
   Else
    Label4.Caption = "节课上课还有"
    Text4.Text = Minute(CDate(zxtime("zf5s")) - Time) + 1
   End If
  ElseIf CDate(zxtime("zf5s")) <= Time And Time < CDate(zxtime("zf5x")) Then
   Text1.Text = "5"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zf5x")) - Time) + 1
  ElseIf CDate(zxtime("zf5x")) <= Time And Time < CDate(zxtime("zf6s")) Then
   Text1.Text = "6"
   Label4.Caption = "节课上课还有"
   Text4.Text = Minute(CDate(zxtime("zf6s")) - Time) + 1
  ElseIf CDate(zxtime("zf6s")) <= Time And Time < CDate(zxtime("zf6x")) Then
   Text1.Text = "6"
   Label4.Caption = "节课下课还有"
   Text4.Text = Minute(CDate(zxtime("zf6x")) - Time) + 1
  ElseIf CDate(zxtime("zf6x")) < Time Then
   Text1.Text = "6"
   Label4.Caption = "节课下课还有"
   Text4.Text = "0"
  End If
 End If
End If
End Sub

Private Sub Timer3_Timer()
If speed > 0 Then
 If Not Form1.Top + speed >= Form33.Height / 3 + 140 Then
  Form1.Top = Form1.Top + speed
  speed = speed + 6
 Else
  Form1.Top = Form1.Top + speed
  speed = speed - 28
 End If
Else
 Form1.Top = Form33.Height
 speed = 1
 Timer4.Enabled = False
 Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
If speed > 0 Then
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), (Screen.Width - Form33.Left) * 0.0422, LWA_ALPHA
 tomd = (Screen.Width - Form33.Left) * 0.0422
 If Form33.Left - speed >= lleft + (Form33.Width / 3) Then
  Form33.Left = Form33.Left - speed
  speed = speed + 5
 Else
  Form33.Left = Form33.Left - speed
  speed = speed - 10.5
 End If
Else
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 255, LWA_ALPHA
 tomd = 255
 Form33.Left = lleft
 Form1.Top = Form1.Height * -1 - 100
 Form1.Left = lleft
 Form33.Show
 speed = 1
 Timer3.Enabled = True
 Timer5.Enabled = True
 Timer4.Enabled = False
End If
End Sub

Public Function zxtime(s$) As String
Dim timelist As String
Open "data\timelist.txt" For Input As #1
Line Input #1, timelist
Close #1
zxtime = Trim(Mid(timelist, InStr(timelist, CStr(s)) + 4, 9))
End Function

Private Sub Timer5_Timer()
If Form1.Top > 0 Then
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 255 - Form1.Top \ (Form33.Height / (256 - tmdu)), LWA_ALPHA
 toumd = 255 - Form1.Top \ (Form33.Height / (256 - tmdu))
Else
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 255, LWA_ALPHA
 toumd = 255
End If
End Sub

Public Function collist(c$) As String
Dim allcolor As String
Open "themes\" & thm & "\color.txt" For Input As #1
Line Input #1, allcolor
Close #1
collist = Trim(Mid(allcolor, InStr(allcolor, CStr(c)) + 5, 3))
End Function

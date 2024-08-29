VERSION 5.00
Begin VB.Form Form35 
   BackColor       =   &H004D5E2B&
   BorderStyle     =   0  'None
   Caption         =   "Form35"
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9720
   LinkTopic       =   "Form35"
   Moveable        =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004D5E2B&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   3720
      TabIndex        =   5
      Text            =   "Ê±¼ä"
      Top             =   4680
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004D5E2B&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   1200
      TabIndex        =   4
      Text            =   "¿ÆÄ¿"
      Top             =   4680
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004D5E2B&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   3720
      TabIndex        =   3
      Text            =   "Ê±¼ä"
      Top             =   3840
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004D5E2B&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   1200
      TabIndex        =   2
      Text            =   "¿ÆÄ¿"
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   2
      Left            =   600
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "¹Ø±Õ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   0
      Top             =   5640
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   99.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2265
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   8055
   End
End
Attribute VB_Name = "Form35"
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

Private Sub Command1_Click()
Dim gb As Integer
gb = MsgBox("ÊÇ·ñ¹Ø±Õ¿¼ÊÔÄ£Ê½?", 36, "¿¼ÊÔÄ£Ê½")
If gb = 6 Then
 speed = 1
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 255, LWA_ALPHA
 Timer2.Enabled = True
End If
End Sub

Private Sub Form_Load()
ksmss = True
Form35.Picture = LoadPicture("themes\" & thm & "\bg4.jpg")
Dim tmd As Integer
Dim rtn As Long
Me.BackColor = RGB(0, 0, 0)
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 0, LWA_ALPHA
speed = 1
Form35.Left = 0
Form35.Top = 0
Form35.Width = Screen.Width
Form35.Height = Screen.Height
Command1.Left = Form35.Width - 1535
Command1.Top = Form35.Height - 1535
Label1.Left = (Screen.Width - Label1.Width) / 2
Label1.Top = (Screen.Height - Label1.Height) / 2 - 1500
Text1.Left = Label1.Left + 360
Text1.Top = Label1.Top + 2520
Text2.Left = Label1.Left + 2880
Text2.Top = Label1.Top + 2520
Text3.Left = Label1.Left + 360
Text3.Top = Label1.Top + 3360
Text4.Left = Label1.Left + 2880
Text4.Top = Label1.Top + 3360
If Len(Time) = 7 Then
 Label1.Caption = "0" & Time
Else
 Label1.Caption = Time
End If
Timer1.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer1_Timer()
If Not tmd + speed >= 255 Then
 tmd = tmd + speed
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), tmd, LWA_ALPHA
 speed = speed + 5
Else
 speed = 1
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 255, LWA_ALPHA
 Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
If Not Form35.Top < Screen.Height * -1 - 100 Then
 Form35.Top = Form35.Top - speed
 speed = speed + 7
Else
 speed = 1
 SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), 0, LWA_ALPHA
 If zdd = True Then
  zdrw = True
 Else
  zdrw = False
 End If
 If audd = True Then
  aud = "1"
 Else
  aud = "0"
 End If
 ksmss = False
 Unload Me
 Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
If Len(Time) = 7 Then
 Label1.Caption = "0" & Time
Else
 Label1.Caption = Time
End If
End Sub

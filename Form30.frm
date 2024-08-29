VERSION 5.00
Begin VB.Form Form30 
   BorderStyle     =   0  'None
   Caption         =   "Form30"
   ClientHeight    =   11445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1095
   LinkTopic       =   "Form30"
   Picture         =   "Form30.frx":0000
   ScaleHeight     =   11445
   ScaleWidth      =   1095
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer6 
      Interval        =   2
      Left            =   600
      Top             =   4560
   End
   Begin VB.Timer Timer5 
      Interval        =   2
      Left            =   120
      Top             =   4560
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   248
      Picture         =   "Form30.frx":3882
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10600
      Width           =   615
   End
   Begin VB.Timer Timer4 
      Interval        =   2
      Left            =   600
      Top             =   4080
   End
   Begin VB.Timer Timer3 
      Interval        =   2
      Left            =   120
      Top             =   4080
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   600
      Top             =   3600
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   120
      Top             =   3600
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A2D581&
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
      Height          =   7095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3060
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   930
      Left            =   120
      TabIndex        =   5
      Top             =   780
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   930
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   930
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   840
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "¿Î±í"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   180
      TabIndex        =   2
      Top             =   2560
      Width           =   855
   End
End
Attribute VB_Name = "Form30"
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
Dim kb As String

Private Sub Command12_Click()
Timer5.Enabled = True
End Sub

Private Sub Form_Load()
Text2.BackColor = RGB(collist("bg21r"), collist("bg21g"), collist("bg21b"))
Form30.Picture = LoadPicture("themes\" & thm & "\bg3.jpg")
Dim rtn As Long
Me.BackColor = RGB(0, 0, 0)
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), tmdu, LWA_ALPHA
speedd = 1
Form30.Left = Screen.Width
Form30.Hide
Timer3.Enabled = True
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
If Len(Hour(Time)) = 1 Then
 Label1.Caption = "0" & Hour(Time)
Else
 Label1.Caption = Hour(Time)
End If
If Len(Minute(Time)) = 1 Then
 Label6.Caption = "0" & Minute(Time)
Else
 Label6.Caption = Minute(Time)
End If
If Len(Second(Time)) = 1 Then
 Label3.Caption = "0" & Second(Time)
Else
 Label3.Caption = Second(Time)
End If
If lskb = "" Then
 Dim kb As String
 If Weekday(Date, 2) = 1 Then
  Open "data\z1.txt" For Input As #1
  Line Input #1, kb
  Close #1
  Text2.Text = kb
 ElseIf Weekday(Date, 2) = 2 Then
  Open "data\z2.txt" For Input As #1
  Line Input #1, kb
  Close #1
  Text2.Text = kb
 ElseIf Weekday(Date, 2) = 3 Then
  Open "data\z3.txt" For Input As #1
  Line Input #1, kb
  Close #1
  Text2.Text = kb
 ElseIf Weekday(Date, 2) = 4 Then
  Open "data\z4.txt" For Input As #1
  Line Input #1, kb
  Close #1
  Text2.Text = kb
 ElseIf Weekday(Date, 2) = 5 Then
  Open "data\z5.txt" For Input As #1
  Line Input #1, kb
  Close #1
  Text2.Text = kb
 Else
  Text2.Text = "½ñÌì²»ÉÏ¿Î"
 End If
Else
 Text2.Text = lskb
End If
End Sub

Private Sub Timer1_Timer()
If lskb = "" Then
 If Weekday(Date, 2) = 1 Then
  Open "data\z1.txt" For Input As #1
  Line Input #1, kb
  Close #1
  Text2.Text = kb
 ElseIf Weekday(Date, 2) = 2 Then
  Open "data\z2.txt" For Input As #1
  Line Input #1, kb
  Close #1
  Text2.Text = kb
 ElseIf Weekday(Date, 2) = 3 Then
  Open "data\z3.txt" For Input As #1
  Line Input #1, kb
  Close #1
  Text2.Text = kb
 ElseIf Weekday(Date, 2) = 4 Then
  Open "data\z4.txt" For Input As #1
  Line Input #1, kb
  Close #1
  Text2.Text = kb
 ElseIf Weekday(Date, 2) = 5 Then
  Open "data\z5.txt" For Input As #1
  Line Input #1, kb
  Close #1
  Text2.Text = kb
 Else
  Text2.Text = "½ñÌì²»ÉÏ¿Î"
 End If
Else
 Text2.Text = lskb
End If
End Sub

Private Sub Timer2_Timer()
If Len(Hour(Time)) = 1 Then
 Label1.Caption = "0" & Hour(Time)
Else
 Label1.Caption = Hour(Time)
End If
If Len(Minute(Time)) = 1 Then
 Label6.Caption = "0" & Minute(Time)
Else
 Label6.Caption = Minute(Time)
End If
If Len(Second(Time)) = 1 Then
 Label3.Caption = "0" & Second(Time)
Else
 Label3.Caption = Second(Time)
End If
End Sub

Private Sub Timer3_Timer()
If Not Form1.Left > Screen.Width + 100 Then
 Form1.Left = Form1.Left + speed
 Form33.Left = Form33.Left + speed
 speed = speed + 7
Else
 speed = 1
 Form1.Hide
 Timer4.Enabled = True
 Timer5.Enabled = False
 Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
If speedd > 0 Then
 If Form30.Left - speedd >= Screen.Width - Form30.Width + (Form30.Width / 3) Then
  Form30.Left = Form30.Left - speedd
  speedd = speedd + 5
 Else
  Form30.Left = Form30.Left - speedd
  speedd = speedd - 10.5
 End If
Else
 Form30.Left = Screen.Width - Form30.Width
 Form1.Left = Screen.Width
 speedd = 1
 Timer3.Enabled = False
 Timer5.Enabled = False
 Timer4.Enabled = False
End If
End Sub

Private Sub Timer5_Timer()
If Not Form30.Left > Screen.Width + 100 Then
 Form30.Left = Form30.Left + speedd
 speedd = speedd + 7
Else
 speed = 1
 Timer6.Enabled = True
 Timer3.Enabled = False
 Timer4.Enabled = False
 Timer5.Enabled = False
 Form30.Hide
 Form1.Show
End If
End Sub

Private Sub Timer6_Timer()
If speed > 0 Then
 If Form1.Left - speed >= lleft + (Form1.Width / 3) Then
  Form1.Left = Form1.Left - speed
  Form33.Left = Form1.Left - speed
  speed = speed + 5
 Else
  Form1.Left = Form1.Left - speed
  Form33.Left = Form1.Left - speed
  speed = speed - 10.5
 End If
Else
 Form33.Left = lleft
 Form1.Left = lleft
 speed = 1
 Timer6.Enabled = False
 Unload Me
End If
End Sub

Public Function collist(c$) As String
Dim allcolor As String
Open "themes\" & thm & "\color.txt" For Input As #1
Line Input #1, allcolor
Close #1
collist = Trim(Mid(allcolor, InStr(allcolor, CStr(c)) + 5, 3))
End Function


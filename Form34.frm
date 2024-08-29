VERSION 5.00
Begin VB.Form Form34 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form34"
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   LinkTopic       =   "Form34"
   ScaleHeight     =   495
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer4 
      Interval        =   2
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   2
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   2
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2160
      Picture         =   "Form34.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   90
      Width           =   1695
   End
End
Attribute VB_Name = "Form34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form33.Show
Timer3.Enabled = True
End Sub

Private Sub Form_Load()
Form34.BackColor = RGB(collist("bg31r"), collist("bg31g"), collist("bg31b"))
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer1.Enabled = True
speedd = 1
Form33.Show
Form34.Left = Screen.Width - Form34.Width
Form34.Top = Form34.Height * -1
End Sub

Private Sub Timer1_Timer()
If Not Form1.Top < Form1.Height * -1 - 100 Then
 Form1.Top = Form1.Top - speed
 speed = speed + 7
Else
 speed = 1
 Form1.Hide
 Form34.Show
 Form33.Show
 Timer2.Enabled = True
 Timer1.Enabled = False
 Timer3.Enabled = False
 Timer4.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
If speedd > 0 Then
 If Form34.Top + speedd <= Form33.Height / 3 Then
  Form34.Top = Form34.Top + speedd
  speedd = speedd + 13
 Else
  Form34.Top = Form34.Top + speedd
  speedd = speedd - 10
 End If
Else
 speedd = 1
 Form34.Top = Form33.Height
 Form33.Show
 Timer1.Enabled = False
 Timer2.Enabled = False
 Timer3.Enabled = False
 Timer4.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
If Form34.Top > Form33.Height - Form34.Height Then
 Form34.Top = Form34.Top - speedd
 speedd = speedd + 8
Else
 speed = 1
 Form34.Top = -100
 Form34.Hide
 Form1.Show
 Form33.Show
 Timer4.Enabled = True
 Timer1.Enabled = False
 Timer2.Enabled = False
 Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
If speed > 0 Then
 If Not Form1.Top + speed >= Form33.Height / 3 + 140 Then
  Form1.Top = Form1.Top + speed
  speed = speed + 6
 Else
  Form1.Top = Form1.Top + speed
  speed = speed - 27
 End If
Else
 Form1.Top = Form33.Height
 speed = 1
 Form33.Show
 Timer1.Enabled = False
 Timer2.Enabled = False
 Timer3.Enabled = False
 Timer4.Enabled = False
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

VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form3 
   BackColor       =   &H00F9FFDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ºóÌ¨"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "È·¶¨"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   4095
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   120
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ÇëÊäÈëºóÌ¨ÃÜÂë"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim setpsw As String
Open "data\pswset.txt" For Input As #1
Line Input #1, setpsw
Close #1
Dim clspsw As String
Open "data\pswcls.txt" For Input As #1
Line Input #1, clspsw
Close #1
If Text1.Text = setpsw Then
 Form4.Show
 Unload Me
ElseIf Text1.Text = clspsw Then
 If aud = "1" Then
  WindowsMediaPlayer1.URL = "themes\" & thm & "\audio\tc.mp3"
 End If
 Open "data\exit.txt" For Output As #1
 Print #1, "1"
 Close #1
 Form3.Hide
 Timer1.Enabled = True
Else
 MsgBox "ºóÌ¨ÃÜÂë´íÎó", vbCritical, "ÌáÊ¾"
End If
End Sub

Private Sub Form_Load()
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
If Not Form1.Left > Screen.Width + 100 Then
 Form1.Left = Form1.Left + speed
 Form33.Left = Form33.Left + speed
 speed = speed + 7
Else
 Unload Form2
 Unload Form4
 Unload Form5
 Unload Form6
 Unload Form7
 Unload Form8
 Unload Form9
 Unload Form10
 Unload Form11
 Unload Form12
 Unload Form13
 Unload Form14
 Unload Form15
 Unload Form16
 Unload Form17
 Unload Form18
 Unload Form19
 Unload Form20
 Unload Form21
 Unload Form24
 Unload Form25
 Unload Form26
 Unload Form27
 Unload Form28
 Unload Form29
 Unload Form30
 Unload Form31
 Unload Form32
 Unload Form33
 Unload Form34
 Unload Form35
 Unload Form1
 Unload Me
End If
End Sub

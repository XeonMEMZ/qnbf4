VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "Form1.frx":6988A
   ScaleHeight     =   8070
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00A2D581&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   5520
      Width           =   3255
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A2D581&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   2700
      TabIndex        =   19
      Top             =   4800
      Width           =   1275
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A2D581&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   4080
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      Picture         =   "Form1.frx":71236
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6960
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   4080
      Top             =   120
   End
   Begin VB.CommandButton Command5 
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
      Height          =   375
      Left            =   4920
      Picture         =   "Form1.frx":7158B
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7440
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A2D581&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Timer Timer5 
      Interval        =   5000
      Left            =   4560
      Top             =   120
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
      Height          =   375
      Left            =   5400
      Picture         =   "Form1.frx":71912
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6960
      Width           =   375
   End
   Begin VB.DriveListBox Drive1 
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ëæ»úµãÃû"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A2D581&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3080
      Width           =   1935
   End
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   5040
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   5520
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      Picture         =   "Form1.frx":71C7A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00A2D581&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ò»¼ü´ò¿ªUÅÌ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   26.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6960
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D5D587&
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
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   120
      TabIndex        =   23
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "¼Ù"
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
      Height          =   735
      Left            =   4120
      TabIndex        =   22
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Çë"
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
      Height          =   735
      Left            =   4120
      TabIndex        =   21
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Êµµ½:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1515
      TabIndex        =   18
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ó¦µ½:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1520
      TabIndex        =   17
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¾àÔÂ¿¼»¹ÓÐ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1440
      TabIndex        =   13
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Ìì"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4200
      TabIndex        =   12
      Top             =   1940
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   855
      Left            =   4995
      Picture         =   "Form1.frx":720B0
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   4995
      Picture         =   "Form1.frx":DB93A
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   4995
      Picture         =   "Form1.frx":1451C4
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   4995
      Picture         =   "Form1.frx":1AEA4E
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "³£ÓÃÈí¼þË«»÷´ò¿ª"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   5010
      TabIndex        =   5
      Top             =   135
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "¿Î±í"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ÖµÈÕ:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1520
      TabIndex        =   0
      Top             =   200
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
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
Dim upf As String
Dim ksms As Integer
Dim cname As String
Dim fxs As String
Dim zrnext As String
Dim zname As String
Dim namect As String
Dim zrname1 As String
Dim zrname2 As String
Dim znext As String
Dim nzr As String
Dim nc As String
Dim atl1 As String
Dim att1 As String
Dim t As String
Dim kb As String
Dim fx As String
Dim cy1tt As String
Dim cy2tt As String
Dim cy3tt As String
Dim cy4tt As String
Dim djstext As String
Dim djstime As String
Dim cy1l As String
Dim cy1t As String
Dim cy2l As String
Dim cy2t As String
Dim cy3l As String
Dim cy3t As String
Dim cy4l As String
Dim cy4t As String
Dim lin1 As String
Dim tim1 As String
Dim oc1 As String
Dim lin2 As String
Dim tim2 As String
Dim oc2 As String
Dim lin3 As String
Dim tim3 As String
Dim oc3 As String
Dim lin4 As String
Dim tim4 As String
Dim oc4 As String
Dim sklin As String
Dim skoc As String
Dim xklin As String
Dim xkoc As String
Dim xkchb As String
Dim xkchbtm As String
Dim autostd As String
Dim zdgj As String
Dim tlino As String
Dim tlint As String
Dim ttimjy As String
Dim ttimjj As String
Dim tllq As String
Dim tmgr As String
Dim udatop As String

Private Sub Command1_Click()
Open "data\upf.txt" For Input As #1
Line Input #1, upf
Close #1
Call Shell("cmd /c start " & upf)
End Sub

Private Sub Command12_Click()
Form30.Show
End Sub

Private Sub Command2_Click()
Form3.Show
End Sub

Private Sub Command3_Click()
Unload Form35
ksms = MsgBox("ÊÇ·ñ´ò¿ª¿¼ÊÔÄ£Ê½?" & vbCrLf & "¸ÃÄ£Ê½ÏÂ½«¹Ø±Õ²¿·Ö×Ô¶¯»¯¹¦ÄÜ", 36, "¿¼ÊÔÄ£Ê½")
If ksms = 6 Then
 If zdrw = True Then
  zdd = True
 Else
  zdd = False
 End If
 zdrw = False
 If aud = "1" Then
  audd = True
  WindowsMediaPlayer1.URL = "themes\" & thm & "\audio\ks.mp3"
 Else
  audd = False
 End If
 aud = "0"
 Form35.Show
End If
End Sub

Private Sub Command4_Click()
Open "data\namec.txt" For Input As #1
Line Input #1, cname
Close #1
Dim sj As Integer
Randomize
sj = Int(Rnd * (CInt(cname) - 1 + 1) + 1)
Text3.Text = namelist(sj)
End Sub

Private Sub Command5_Click()
Form34.Show
End Sub

Private Sub Form_Load()
Text2.BackColor = RGB(collist("bg21r"), collist("bg21g"), collist("bg21b"))
Text1.BackColor = RGB(collist("bg22r"), collist("bg22g"), collist("bg22b"))
Text3.BackColor = RGB(collist("bg22r"), collist("bg22g"), collist("bg22b"))
Text5.BackColor = RGB(collist("bg23r"), collist("bg23g"), collist("bg23b"))
Text4.BackColor = RGB(collist("bg23r"), collist("bg23g"), collist("bg23b"))
Text6.BackColor = RGB(collist("bg23r"), collist("bg23g"), collist("bg23b"))
Text7.BackColor = RGB(collist("bg23r"), collist("bg23g"), collist("bg23b"))
Form1.Picture = LoadPicture("themes\" & thm & "\bg2.jpg")
Dim rtn As Long
Me.BackColor = RGB(0, 0, 0)
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), tmdu, LWA_ALPHA
Form33.Show
Timer5.Enabled = False
speed = 1
disk = Drive1.ListCount
dskop = 0
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
If fxs = "5" Then
 If Not Weekday(Date, 2) = 6 And Not Weekday(Date, 2) = 7 Then
  Open "data\nextzr.txt" For Input As #1
  Line Input #1, zrnext
  Close #1
  If Not CStr(Day(Date)) = zrnext Then
   Open "data\namezr.txt" For Input As #1
   Line Input #1, zname
   Close #1
   Open "data\namec.txt" For Input As #1
   Line Input #1, namect
   Close #1
   If Not Fix(zname) >= Fix(namect) Then
    Open "data\namezr.txt" For Output As #1
    zrname1 = CStr(Fix(Val(zname)) + 1)
    Print #1, zrname1
    Close #1
   Else
    Open "data\namezr.txt" For Output As #1
    zrname2 = "1"
    Print #1, zrname2
    Close #1
   End If
   Open "data\nextzr.txt" For Output As #1
   znext = CStr(Day(Date))
   Print #1, znext
   Close #1
  End If
 End If
ElseIf fxs = "6" Then
 If Not Weekday(Date, 2) = 7 Then
  Open "data\nextzr.txt" For Input As #1
  Line Input #1, zrnext
  Close #1
  If Not CStr(Day(Date)) = zrnext Then
   Open "data\namezr.txt" For Input As #1
   Line Input #1, zname
   Close #1
   Open "data\namec.txt" For Input As #1
   Line Input #1, namect
   Close #1
   If Not Fix(zname) >= Fix(namect) Then
    Open "data\namezr.txt" For Output As #1
    zrname1 = CStr(Fix(Val(zname)) + 1)
    Print #1, zrname1
    Close #1
   Else
    Open "data\namezr.txt" For Output As #1
    zrname2 = "1"
    Print #1, zrname2
    Close #1
   End If
   Open "data\nextzr.txt" For Output As #1
   znext = CStr(Day(Date))
   Print #1, znext
   Close #1
  End If
 End If
ElseIf fxs = "7" Then
 Open "data\nextzr.txt" For Input As #1
 Line Input #1, zrnext
 Close #1
 If Not CStr(Day(Date)) = zrnext Then
  Open "data\namezr.txt" For Input As #1
  Line Input #1, zname
  Close #1
  Open "data\namec.txt" For Input As #1
  Line Input #1, namect
  Close #1
  If Not Fix(zname) >= Fix(namect) Then
   Open "data\namezr.txt" For Output As #1
   zrname1 = CStr(Fix(Val(zname)) + 1)
   Print #1, zrname1
   Close #1
  Else
   Open "data\namezr.txt" For Output As #1
   zrname2 = "1"
   Print #1, zrname2
   Close #1
  End If
  Open "data\nextzr.txt" For Output As #1
  znext = CStr(Day(Date))
  Print #1, znext
  Close #1
 End If
End If
If lszr = "" Then
 Open "data\namezr.txt" For Input As #1
 Line Input #1, nzr
 Close #1
 Text1.Text = namelist(Int(nzr))
Else
 Text1.Text = lszr
End If
Open "data\namec.txt" For Input As #1
Line Input #1, nc
Close #1
Text4.Text = nc
Text6.Text = nc
Open "data\atlink1.txt" For Input As #1
Line Input #1, atl1
Close #1
Open "data\attime1.txt" For Input As #1
Line Input #1, att1
Close #1
If Not atl1 = "" And Not att1 = "" Then
 zdrw = True
Else
 zdrw = False
End If
Form1.Left = Screen.Width
lleft = Screen.Width - Form1.Width
Close #1
Open "data\top.txt" For Input As #1
Line Input #1, t
Form1.Top = t + Form33.Height
Close #1
If lskb = "" Then
 Open "data\fx.txt" For Input As #1
 Line Input #1, fx
 Close #1
 If fx = "5" Then
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
 ElseIf fx = "6" Then
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
  ElseIf Weekday(Date, 2) = 6 Then
   Open "data\z6.txt" For Input As #1
   Line Input #1, kb
   Close #1
   Text2.Text = kb
  Else
   Text2.Text = "½ñÌì²»ÉÏ¿Î"
  End If
 ElseIf fx = "7" Then
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
  ElseIf Weekday(Date, 2) = 6 Then
   Open "data\z6.txt" For Input As #1
   Line Input #1, kb
   Close #1
   Text2.Text = kb
  ElseIf Weekday(Date, 2) = 7 Then
   Open "data\z7.txt" For Input As #1
   Line Input #1, kb
   Close #1
   Text2.Text = kb
  Else
   Text2.Text = "½ñÌì²»ÉÏ¿Î"
  End If
 End If
Else
 Text2.Text = lskb
End If
Open "data\cy1t.txt" For Input As #1
Line Input #1, cy1tt
Close #1
Open "data\cy2t.txt" For Input As #1
Line Input #1, cy2tt
Close #1
Open "data\cy3t.txt" For Input As #1
Line Input #1, cy3tt
Close #1
Open "data\cy4t.txt" For Input As #1
Line Input #1, cy4tt
Close #1
Image1.Picture = LoadPicture(cy1tt)
Image2.Picture = LoadPicture(cy2tt)
Image3.Picture = LoadPicture(cy3tt)
Image4.Picture = LoadPicture(cy4tt)
Open "data\djstext.txt" For Input As #1
Line Input #1, djstext
Close #1
Open "data\djstime.txt" For Input As #1
Line Input #1, djstime
Close #1
Label11.Caption = "¾à" & djstext & "»¹ÓÐ"
Text5.Text = CDate(djstime) - Date
End Sub

Private Sub Image1_DblClick()
Open "data\cy1l.txt" For Input As #1
Line Input #1, cy1l
Close #1
Open "data\cy1t.txt" For Input As #1
Line Input #1, cy1t
Close #1
Image1.Picture = LoadPicture(cy1t)
Call Shell("cmd /c start " & cy1l)
End Sub

Private Sub Image2_DblClick()
Open "data\cy2l.txt" For Input As #1
Line Input #1, cy2l
Close #1
Open "data\cy2t.txt" For Input As #1
Line Input #1, cy2t
Close #1
Image2.Picture = LoadPicture(cy2t)
Call Shell("cmd /c start " & cy2l)
End Sub

Private Sub Image3_DblClick()
Open "data\cy3l.txt" For Input As #1
Line Input #1, cy3l
Close #1
Open "data\cy3t.txt" For Input As #1
Line Input #1, cy3t
Close #1
Image3.Picture = LoadPicture(cy3t)
Call Shell("cmd /c start " & cy3l)
End Sub

Private Sub Image4_DblClick()
Open "data\cy4l.txt" For Input As #1
Line Input #1, cy4l
Close #1
Open "data\cy4t.txt" For Input As #1
Line Input #1, cy4t
Close #1
Image4.Picture = LoadPicture(cy4t)
Call Shell("cmd /c start " & cy4l)
End Sub

Private Sub Label5_DblClick()
Form15.Show
End Sub

Private Sub Timer1_Timer()
SetLayeredWindowAttributes hwnd, RGB(0, 0, 0), toumd, LWA_ALPHA
End Sub

Private Sub Timer2_Timer()
If exitproc("qnbf.exe") Then
 If dt = "1" Then
  If Not exitproc("dual.exe") Then
   Call Shell("cmd /c start dual.exe")
  End If
 End If
End If
If zdrw = True Then
 Open "data\atlink1.txt" For Input As #1
 Line Input #1, lin1
 Open "data\attime1.txt" For Input As #2
 Line Input #2, tim1
 Open "data\atoc1.txt" For Input As #3
 Line Input #3, oc1
 Close #1
 Close #2
 Close #3
 If tim1 = CStr(Time) Then
  If oc1 = "1" Then
   Call Shell("cmd /c start " & lin1)
  Else
   Call Shell("cmd /c taskkill /f /im " & lin1)
  End If
 End If
 Open "data\atlink2.txt" For Input As #1
 Line Input #1, lin2
 Open "data\attime2.txt" For Input As #2
 Line Input #2, tim2
 Open "data\atoc2.txt" For Input As #3
 Line Input #3, oc2
 Close #1
 Close #2
 Close #3
 If tim2 = CStr(Time) Then
  If oc2 = "1" Then
   Call Shell("cmd /c start " & lin2)
  Else
   Call Shell("cmd /c taskkill /f /im " & lin2)
  End If
 End If
 Open "data\atlink3.txt" For Input As #1
 Line Input #1, lin3
 Open "data\attime3.txt" For Input As #2
 Line Input #2, tim3
 Open "data\atoc3.txt" For Input As #3
 Line Input #3, oc3
 Close #1
 Close #2
 Close #3
 If tim3 = CStr(Time) Then
  If oc3 = "1" Then
   Call Shell("cmd /c start " & lin3)
  Else
   Call Shell("cmd /c taskkill /f /im " & lin3)
  End If
 End If
 Open "data\atlink4.txt" For Input As #1
 Line Input #1, lin4
 Open "data\attime4.txt" For Input As #2
 Line Input #2, tim4
 Open "data\atoc4.txt" For Input As #3
 Line Input #3, oc4
 Close #1
 Close #2
 Close #3
 If tim4 = CStr(Time) Then
  If oc4 = "1" Then
   Call Shell("cmd /c start " & lin4)
  Else
   Call Shell("cmd /c taskkill /f /im " & lin4)
  End If
 End If
End If
Open "data\skatlink.txt" For Input As #1
Line Input #1, sklin
Close #1
Open "data\skatoc.txt" For Input As #1
Line Input #1, skoc
Close #1
If CDate(zxtime("zm1s")) = Time Or CDate(zxtime("zm2s")) = Time Or CDate(zxtime("zm3s")) = Time Or CDate(zxtime("zm4s")) = Time Or CDate(zxtime("zm5s")) = Time Or CDate(zxtime("zm6s")) = Time Or CDate(zxtime("zm7s")) = Time Or CDate(zxtime("zm8s")) = Time Then
 If aud = "1" Then
  WindowsMediaPlayer1.URL = "themes\" & thm & "\audio\sk.mp3"
 End If
 If Not sklin = "" Then
  If skoc = "1" Then
   Call Shell("cmd /c start " & sklin)
  Else
   Call Shell("cmd /c taskkill /f /im " & sklin)
  End If
 End If
End If
Open "data\xkatlink.txt" For Input As #1
Line Input #1, xklin
Close #1
Open "data\xkatoc.txt" For Input As #1
Line Input #1, xkoc
Close #1
If Not xklin = "" Then
 If fricls = 0 Then
  If CDate(zxtime("zm1x")) = Time Or CDate(zxtime("zm2x")) = Time Or CDate(zxtime("zm3x")) = Time Or CDate(zxtime("zm4x")) = Time Or CDate(zxtime("zm5x")) = Time Or CDate(zxtime("zm6x")) = Time Or CDate(zxtime("zm7x")) = Time Or CDate(zxtime("zm8x")) = Time Then
   If xkoc = "1" Then
    Call Shell("cmd /c start " & xklin)
   Else
    Call Shell("cmd /c taskkill /f /im " & xklin)
   End If
  End If
 Else
  If CDate(zxtime("zm1x")) = Time Or CDate(zxtime("zm2x")) = Time Or CDate(zxtime("zm3x")) = Time Or CDate(zxtime("zm4x")) = Time Or CDate(zxtime("zf5x")) = Time Or CDate(zxtime("zf6x")) = Time Then
   If xkoc = "1" Then
    Call Shell("cmd /c start " & xklin)
   Else
    Call Shell("cmd /c taskkill /f /im " & xklin)
   End If
  End If
 End If
End If
Open "data\xkchb.txt" For Input As #1
Line Input #1, xkchb
Close #1
Open "data\xkchbtime.txt" For Input As #1
Line Input #1, xkchbtm
Close #1
If fricls = 0 Then
 If CDate(zxtime("zm1x")) = Time Or CDate(zxtime("zm2x")) = Time Or CDate(zxtime("zm3x")) = Time Or CDate(zxtime("zm4x")) = Time Or CDate(zxtime("zm5x")) = Time Or CDate(zxtime("zm6x")) = Time Or CDate(zxtime("zm7x")) = Time Or CDate(zxtime("zm8x")) = Time Then
  If aud = "1" Then
   WindowsMediaPlayer1.URL = "themes\" & thm & "\audio\xk.mp3"
  End If
  If xkchb = "1" Then
   If lszr = "" Then
    ystext = Text1.Text & ",Çë²ÁºÚ°å"
   Else
    ystext = lszr & ",Çë²ÁºÚ°å"
    Close #1
   End If
   ystime = Val(xkchbtm)
   If ksmss = False Then
    Form28.Show
   End If
  End If
 End If
Else
 If CDate(zxtime("zm1x")) = Time Or CDate(zxtime("zm2x")) = Time Or CDate(zxtime("zm3x")) = Time Or CDate(zxtime("zm4x")) = Time Or CDate(zxtime("zf5x")) = Time Or CDate(zxtime("zf6x")) = Time Then
  If aud = "1" Then
   WindowsMediaPlayer1.URL = "themes\" & thm & "\audio\xk.mp3"
  End If
  If xkchb = "1" Then
   If lszr = "" Then
    ystext = Text1.Text & ",Çë²ÁºÚ°å"
   Else
    ystext = lszr & ",Çë²ÁºÚ°å"
    Close #1
   End If
   ystime = Val(xkchbtm)
   If ksmss = False Then
    Form28.Show
   End If
  End If
 End If
End If
Open "data\autostd.txt" For Input As #1
Line Input #1, autostd
Close #1
Open "data\zdgj.txt" For Input As #1
Line Input #1, zdgj
Close #1
If zdgj = "1" Then
 If CDate(autostd) = Time Then
  If aud = "1" Then
   WindowsMediaPlayer1.URL = "themes\" & thm & "\audio\gj.mp3"
  End If
  ystext = "¼ÆËã»ú½«ÔÚ5ÃëÄÚ¹Ø»ú"
  ystime = 3
  Form28.Show
  Timer5.Enabled = True
 End If
End If
Open "data\tctask1.txt" For Input As #1
Line Input #1, tlino
Open "data\tctask2.txt" For Input As #2
Line Input #2, tlint
Open "data\tctimejy.txt" For Input As #3
Line Input #3, ttimjy
Open "data\tctimejj.txt" For Input As #4
Line Input #4, ttimjj
Open "data\tcllq.txt" For Input As #5
Line Input #5, tllq
Open "data\tcmgr.txt" For Input As #6
Line Input #6, tmgr
Close #1
Close #2
Close #3
Close #4
Close #5
Close #6
If Not ttimjy = "" And Not ttimjj = "" Then
 If ttimjy <= Time And Time <= ttimjj Then
  sjkz = True
 Else
  sjkz = False
 End If
Else
 sjkz = False
End If
If sjkz = True Then
 If ttimjj <= Time Then
  sjkz = False
 Else
  If exitproc(tlino) Then
   If aud = "1" Then
    WindowsMediaPlayer1.URL = "themes\" & thm & "\audio\sj.mp3"
   End If
   Call Shell("cmd /c taskkill /f /im " & tlino)
  End If
  If exitproc(tlint) Then
   If aud = "1" Then
    WindowsMediaPlayer1.URL = "themes\" & thm & "\audio\sj.mp3"
   End If
   Call Shell("cmd /c taskkill /f /im " & tlint)
  End If
  If tllq = "1" Then
   If exitproc("iexplore.exe") Then
    If aud = "1" Then
     WindowsMediaPlayer1.URL = "themes\" & thm & "\audio\sj.mp3"
    End If
    Call Shell("cmd /c taskkill /f /im iexplore.exe")
   End If
   If exitproc("msedge.exe") Then
    If aud = "1" Then
     WindowsMediaPlayer1.URL = "themes\" & thm & "\audio\sj.mp3"
    End If
    Call Shell("cmd /c taskkill /f /im msedge.exe")
   End If
   If exitproc("chrome.exe") Then
    If aud = "1" Then
     WindowsMediaPlayer1.URL = "themes\" & thm & "\audio\sj.mp3"
    End If
    Call Shell("cmd /c taskkill /f /im chrome.exe")
   End If
   If exitproc("firefox.exe") Then
    If aud = "1" Then
     WindowsMediaPlayer1.URL = "themes\" & thm & "\audio\sj.mp3"
    End If
    Call Shell("cmd /c taskkill /f /im firefox.exe")
   End If
  End If
  If tmgr = "1" Then
   If exitproc("Taskmgr.exe") Then
    If aud = "1" Then
     WindowsMediaPlayer1.URL = "themes\" & thm & "\audio\sj.mp3"
    End If
    Call Shell("cmd /c taskkill /f /im Taskmgr.exe")
   End If
  End If
 End If
End If
Drive1.Refresh
Open "data\udatop.txt" For Input As #1
Line Input #1, udatop
Close #1
If udatop = "1" Then
 If Drive1.ListCount > disk Then
  If dskop = 0 Then
   Dim upf As String
   Open "data\upf.txt" For Input As #1
   Line Input #1, upf
   Close #1
   Call Shell("cmd /c start " & upf)
   Form24.Show
   dskop = 1
  End If
 Else
  dskop = 0
 End If
End If
End Sub

Private Sub Timer3_Timer()
If lskb = "" Then
 Open "data\fx.txt" For Input As #1
 Line Input #1, fx
 Close #1
 If fx = "5" Then
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
 ElseIf fx = "6" Then
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
  ElseIf Weekday(Date, 2) = 6 Then
   Open "data\z6.txt" For Input As #1
   Line Input #1, kb
   Close #1
   Text2.Text = kb
  Else
   Text2.Text = "½ñÌì²»ÉÏ¿Î"
  End If
 ElseIf fx = "7" Then
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
  ElseIf Weekday(Date, 2) = 6 Then
   Open "data\z6.txt" For Input As #1
   Line Input #1, kb
   Close #1
   Text2.Text = kb
  ElseIf Weekday(Date, 2) = 7 Then
   Open "data\z7.txt" For Input As #1
   Line Input #1, kb
   Close #1
   Text2.Text = kb
  Else
   Text2.Text = "½ñÌì²»ÉÏ¿Î"
  End If
 End If
Else
 Text2.Text = lskb
End If
If lszr = "" Then
 Open "data\namezr.txt" For Input As #1
 Line Input #1, nzr
 Close #1
 Text1.Text = namelist(Int(nzr))
Else
 Text1.Text = lszr
End If
Open "data\djstext.txt" For Input As #1
Line Input #1, djstext
Close #1
Open "data\djstime.txt" For Input As #1
Line Input #1, djstime
Close #1
Label11.Caption = "¾à" & djstext & "»¹ÓÐ"
Text5.Text = CDate(djstime) - Date
End Sub

Public Function namelist(n%) As String
Dim name As String
Open "data\name.txt" For Input As #1
Line Input #1, name
Close #1
If n < 10 Then
 namelist = Trim(Mid(name, InStr(name, CStr(n)) + 1, 4))
ElseIf n >= 10 Then
 namelist = Trim(Mid(name, InStr(name, CStr(n)) + 2, 4))
End If
End Function

Public Function zxtime(s$) As String
Dim timelist As String
Open "data\timelist.txt" For Input As #1
Line Input #1, timelist
Close #1
zxtime = Trim(Mid(timelist, InStr(timelist, CStr(s)) + 4, 9))
End Function

Private Sub Timer5_Timer()
Call Shell("cmd /c shutdown /s /t 0")
Timer5.Enabled = False
End Sub

Public Function collist(c$) As String
Dim allcolor As String
Open "themes\" & thm & "\color.txt" For Input As #1
Line Input #1, allcolor
Close #1
collist = Trim(Mid(allcolor, InStr(allcolor, CStr(c)) + 5, 3))
End Function

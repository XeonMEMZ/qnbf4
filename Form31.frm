VERSION 5.00
Begin VB.Form Form31 
   BackColor       =   &H00F9FFDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ñ¡Ôñ¿Î±í"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6015
   Icon            =   "Form31.frx":0000
   LinkTopic       =   "Form31"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6015
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command8 
      Caption         =   "ÖÜÁù"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "ÖÜÈÕ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "²»ÉÏ¿Î"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ÖÜÎå"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ÖÜËÄ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ÖÜÈý"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÖÜ¶þ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÖÜÒ»"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÇëÑ¡ÔñÉÏÖÜ¼¸µÄ¿Î"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2880
      TabIndex        =   1
      Top             =   1080
      Width           =   2880
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¼ì²âµ½½ñÈÕ²»ÉÏ¿Î"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2880
   End
End
Attribute VB_Name = "Form31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kb As String
Private Sub Command1_Click()
Open "data\z1.txt" For Input As #1
Line Input #1, kb
Close #1
lskb = kb
Dim zrnext As String
 Open "data\nextzr.txt" For Input As #1
 Line Input #1, zrnext
 Close #1
 If Not CStr(Day(Date)) = zrnext Then
  Dim zname As String
  Open "data\namezr.txt" For Input As #1
  Line Input #1, zname
  Close #1
  Dim namect As String
  Open "data\namec.txt" For Input As #1
  Line Input #1, namect
  Close #1
  If Not Fix(zname) >= Fix(namect) Then
   Dim zrname1 As String
   Open "data\namezr.txt" For Output As #1
   zrname1 = CStr(Fix(Val(zname)) + 1)
   Print #1, zrname1
   Close #1
  Else
   Dim zrname2 As String
   Open "data\namezr.txt" For Output As #1
   zrname2 = "1"
   Print #1, zrname2
   Close #1
  End If
  Dim znext As String
  Open "data\nextzr.txt" For Output As #1
  znext = CStr(Day(Date))
  Print #1, znext
  Close #1
 End If
Unload Me
End Sub

Private Sub Command2_Click()
Open "data\z2.txt" For Input As #1
Line Input #1, kb
Close #1
lskb = kb
Dim zrnext As String
 Open "data\nextzr.txt" For Input As #1
 Line Input #1, zrnext
 Close #1
 If Not CStr(Day(Date)) = zrnext Then
  Dim zname As String
  Open "data\namezr.txt" For Input As #1
  Line Input #1, zname
  Close #1
  Dim namect As String
  Open "data\namec.txt" For Input As #1
  Line Input #1, namect
  Close #1
  If Not Fix(zname) >= Fix(namect) Then
   Dim zrname1 As String
   Open "data\namezr.txt" For Output As #1
   zrname1 = CStr(Fix(Val(zname)) + 1)
   Print #1, zrname1
   Close #1
  Else
   Dim zrname2 As String
   Open "data\namezr.txt" For Output As #1
   zrname2 = "1"
   Print #1, zrname2
   Close #1
  End If
  Dim znext As String
  Open "data\nextzr.txt" For Output As #1
  znext = CStr(Day(Date))
  Print #1, znext
  Close #1
 End If
Unload Me
End Sub

Private Sub Command3_Click()
Open "data\z3.txt" For Input As #1
Line Input #1, kb
Close #1
lskb = kb
Dim zrnext As String
 Open "data\nextzr.txt" For Input As #1
 Line Input #1, zrnext
 Close #1
 If Not CStr(Day(Date)) = zrnext Then
  Dim zname As String
  Open "data\namezr.txt" For Input As #1
  Line Input #1, zname
  Close #1
  Dim namect As String
  Open "data\namec.txt" For Input As #1
  Line Input #1, namect
  Close #1
  If Not Fix(zname) >= Fix(namect) Then
   Dim zrname1 As String
   Open "data\namezr.txt" For Output As #1
   zrname1 = CStr(Fix(Val(zname)) + 1)
   Print #1, zrname1
   Close #1
  Else
   Dim zrname2 As String
   Open "data\namezr.txt" For Output As #1
   zrname2 = "1"
   Print #1, zrname2
   Close #1
  End If
  Dim znext As String
  Open "data\nextzr.txt" For Output As #1
  znext = CStr(Day(Date))
  Print #1, znext
  Close #1
 End If
Unload Me
End Sub

Private Sub Command4_Click()
Open "data\z4.txt" For Input As #1
Line Input #1, kb
Close #1
lskb = kb
Dim zrnext As String
 Open "data\nextzr.txt" For Input As #1
 Line Input #1, zrnext
 Close #1
 If Not CStr(Day(Date)) = zrnext Then
  Dim zname As String
  Open "data\namezr.txt" For Input As #1
  Line Input #1, zname
  Close #1
  Dim namect As String
  Open "data\namec.txt" For Input As #1
  Line Input #1, namect
  Close #1
  If Not Fix(zname) >= Fix(namect) Then
   Dim zrname1 As String
   Open "data\namezr.txt" For Output As #1
   zrname1 = CStr(Fix(Val(zname)) + 1)
   Print #1, zrname1
   Close #1
  Else
   Dim zrname2 As String
   Open "data\namezr.txt" For Output As #1
   zrname2 = "1"
   Print #1, zrname2
   Close #1
  End If
  Dim znext As String
  Open "data\nextzr.txt" For Output As #1
  znext = CStr(Day(Date))
  Print #1, znext
  Close #1
 End If
Unload Me
End Sub

Private Sub Command5_Click()
Open "data\z5.txt" For Input As #1
Line Input #1, kb
Close #1
lskb = kb
Dim fxs As String
Open "data\fx.txt" For Input As #1
Line Input #1, fxs
Close #1
If fxs = "5" Then
 fricls = 1
End If
Dim zrnext As String
 Open "data\nextzr.txt" For Input As #1
 Line Input #1, zrnext
 Close #1
 If Not CStr(Day(Date)) = zrnext Then
  Dim zname As String
  Open "data\namezr.txt" For Input As #1
  Line Input #1, zname
  Close #1
  Dim namect As String
  Open "data\namec.txt" For Input As #1
  Line Input #1, namect
  Close #1
  If Not Fix(zname) >= Fix(namect) Then
   Dim zrname1 As String
   Open "data\namezr.txt" For Output As #1
   zrname1 = CStr(Fix(Val(zname)) + 1)
   Print #1, zrname1
   Close #1
  Else
   Dim zrname2 As String
   Open "data\namezr.txt" For Output As #1
   zrname2 = "1"
   Print #1, zrname2
   Close #1
  End If
  Dim znext As String
  Open "data\nextzr.txt" For Output As #1
  znext = CStr(Day(Date))
  Print #1, znext
  Close #1
 End If
Unload Me
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
Open "data\z7.txt" For Input As #1
Line Input #1, kb
Close #1
lskb = kb
Dim fxs As String
Open "data\fx.txt" For Input As #1
Line Input #1, fxs
Close #1
If fxs = "7" Then
 fricls = 1
End If
Dim zrnext As String
 Open "data\nextzr.txt" For Input As #1
 Line Input #1, zrnext
 Close #1
 If Not CStr(Day(Date)) = zrnext Then
  Dim zname As String
  Open "data\namezr.txt" For Input As #1
  Line Input #1, zname
  Close #1
  Dim namect As String
  Open "data\namec.txt" For Input As #1
  Line Input #1, namect
  Close #1
  If Not Fix(zname) >= Fix(namect) Then
   Dim zrname1 As String
   Open "data\namezr.txt" For Output As #1
   zrname1 = CStr(Fix(Val(zname)) + 1)
   Print #1, zrname1
   Close #1
  Else
   Dim zrname2 As String
   Open "data\namezr.txt" For Output As #1
   zrname2 = "1"
   Print #1, zrname2
   Close #1
  End If
  Dim znext As String
  Open "data\nextzr.txt" For Output As #1
  znext = CStr(Day(Date))
  Print #1, znext
  Close #1
 End If
Unload Me
End Sub

Private Sub Command8_Click()
Open "data\z6.txt" For Input As #1
Line Input #1, kb
Close #1
lskb = kb
Dim fxs As String
Open "data\fx.txt" For Input As #1
Line Input #1, fxs
Close #1
If fxs = "6" Then
 fricls = 1
End If
Dim zrnext As String
 Open "data\nextzr.txt" For Input As #1
 Line Input #1, zrnext
 Close #1
 If Not CStr(Day(Date)) = zrnext Then
  Dim zname As String
  Open "data\namezr.txt" For Input As #1
  Line Input #1, zname
  Close #1
  Dim namect As String
  Open "data\namec.txt" For Input As #1
  Line Input #1, namect
  Close #1
  If Not Fix(zname) >= Fix(namect) Then
   Dim zrname1 As String
   Open "data\namezr.txt" For Output As #1
   zrname1 = CStr(Fix(Val(zname)) + 1)
   Print #1, zrname1
   Close #1
  Else
   Dim zrname2 As String
   Open "data\namezr.txt" For Output As #1
   zrname2 = "1"
   Print #1, zrname2
   Close #1
  End If
  Dim znext As String
  Open "data\nextzr.txt" For Output As #1
  znext = CStr(Day(Date))
  Print #1, znext
  Close #1
 End If
Unload Me
End Sub

Private Sub Form_Load()
Dim fxs As String
Open "data\fx.txt" For Input As #1
Line Input #1, fxs
Close #1
If fxs = "5" Then
 Command8.Enabled = False
 Command7.Enabled = False
ElseIf fxs = "6" Then
 Command7.Enabled = False
End If
End Sub

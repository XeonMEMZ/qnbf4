VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00F9FFDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "修改主题"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7455
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7455
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "深色"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   18
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "经典"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   18
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "清新"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   18
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   3480
      Width           =   975
   End
   Begin VB.Image Image9 
      Height          =   3135
      Left            =   6960
      Picture         =   "Form5.frx":6988A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image8 
      Height          =   2175
      Left            =   5160
      Picture         =   "Form5.frx":6E29C
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Image Image7 
      Height          =   975
      Left            =   5160
      Picture         =   "Form5.frx":78BC7
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image6 
      Height          =   3135
      Left            =   4440
      Picture         =   "Form5.frx":7CEA2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image5 
      Height          =   2175
      Left            =   2640
      Picture         =   "Form5.frx":82319
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   975
      Left            =   2640
      Picture         =   "Form5.frx":8A3A2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   3135
      Left            =   1920
      Picture         =   "Form5.frx":8E7B0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   120
      Picture         =   "Form5.frx":92219
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   2175
      Left            =   120
      Picture         =   "Form5.frx":95B62
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Open "themes\themes.txt" For Output As #1
Print #1, "1"
Close #1
Dim re As Integer
re = MsgBox("修改主题后重启软件生效!" & vbCrLf & "是否重启软件?", 52, "提示")
If re = 6 Then
 Call Shell("cmd /c start killme.bat")
 Unload Form4
 Unload Me
Else
 Unload Me
End If
End Sub

Private Sub Command2_Click()
Open "themes\themes.txt" For Output As #1
Print #1, "2"
Close #1
Dim re As Integer
re = MsgBox("修改主题后重启软件生效!" & vbCrLf & "是否重启软件?", 52, "提示")
If re = 6 Then
 Call Shell("cmd /c start killme.bat")
 Unload Form4
 Unload Me
Else
 Unload Me
End If
End Sub

Private Sub Command3_Click()
Open "themes\themes.txt" For Output As #1
Print #1, "3"
Close #1
Dim re As Integer
re = MsgBox("修改主题后重启软件生效!" & vbCrLf & "是否重启软件?", 52, "提示")
If re = 6 Then
 Call Shell("cmd /c start killme.bat")
 Unload Form4
 Unload Me
Else
 Unload Me
End If
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture("themes\1\bg1.jpg")
Image2.Picture = LoadPicture("themes\1\bg2.jpg")
Image3.Picture = LoadPicture("themes\1\bg3.jpg")
Image4.Picture = LoadPicture("themes\2\bg1.jpg")
Image5.Picture = LoadPicture("themes\2\bg2.jpg")
Image6.Picture = LoadPicture("themes\2\bg3.jpg")
Image7.Picture = LoadPicture("themes\3\bg1.jpg")
Image8.Picture = LoadPicture("themes\3\bg2.jpg")
Image9.Picture = LoadPicture("themes\3\bg3.jpg")
Dim l1c As String
Open "themes\1\name.txt" For Input As #1
Line Input #1, l1c
Close #1
Command1.Caption = l1c
Dim l2c As String
Open "themes\2\name.txt" For Input As #1
Line Input #1, l2c
Close #1
Command2.Caption = l2c
Dim l3c As String
Open "themes\3\name.txt" For Input As #1
Line Input #1, l3c
Close #1
Command3.Caption = l3c
End Sub

VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00F9FFDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "临时修改课表"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "临时修改课表"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Not Text1.Text = "" Then
 lskb = Text1.Text
 Unload Me
Else
 MsgBox "课表不能为空", vbCritical, "提示"
End If
End Sub

Private Sub Command2_Click()
lskb = ""
Unload Me
End Sub

Private Sub Form_Load()
Dim lkb As String
If Weekday(Date, 2) = 1 Then
 Open "data\z1.txt" For Input As #1
 Line Input #1, lkb
 Close #1
 Text1.Text = lkb
ElseIf Weekday(Date, 2) = 2 Then
 Open "data\z2.txt" For Input As #1
 Line Input #1, lkb
 Close #1
 Text1.Text = lkb
ElseIf Weekday(Date, 2) = 3 Then
 Open "data\z3.txt" For Input As #1
 Line Input #1, lkb
 Close #1
 Text1.Text = lkb
ElseIf Weekday(Date, 2) = 4 Then
 Open "data\z4.txt" For Input As #1
 Line Input #1, lkb
 Close #1
 Text1.Text = lkb
ElseIf Weekday(Date, 2) = 5 Then
 Open "data\z5.txt" For Input As #1
 Line Input #1, lkb
 Close #1
 Text1.Text = lkb
ElseIf Weekday(Date, 2) = 6 Then
 Open "data\z6.txt" For Input As #1
 Line Input #1, lkb
 Close #1
 Text1.Text = lkb
ElseIf Weekday(Date, 2) = 7 Then
 Open "data\z7.txt" For Input As #1
 Line Input #1, lkb
 Close #1
 Text1.Text = lkb
Else
 Text1.Text = "今天不上课"
End If
End Sub

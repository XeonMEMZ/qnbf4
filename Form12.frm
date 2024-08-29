VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00F9FFDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置时间控制"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4830
   Icon            =   "Form12.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4830
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text4 
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
      Left            =   1320
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00F9FFDD&
      Caption         =   "禁用任务管理器(可能无效)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "临时解禁"
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
      Left            =   1800
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text3 
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
      Left            =   1080
      TabIndex        =   7
      Top             =   720
      Width           =   3495
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00F9FFDD&
      Caption         =   "禁用浏览器"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
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
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   3495
   End
   Begin VB.TextBox Text2 
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
      Left            =   3600
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "创建时间控制"
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
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除时间控制"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "禁用时间:"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "进程2:"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "进程1:"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "解禁时间:"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   1020
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Not Text2.Text = "" And Not Text4.Text = "" And Len(Text2.Text) = 8 And Len(Text4.Text) = 8 Then
 If Not Text1.Text = "main.exe" And Not Text3.Text = "main.exe" And Not Text1.Text = "dual.exe" And Not Text3.Text = "dual.exe" And Text2.Text > Text4.Text Then
  Dim atlo As String
  Open "data\tctask1.txt" For Output As #1
  atlo = Text1.Text
  Print #1, atlo
  Close #1
  Dim atlt As String
  Open "data\tctask2.txt" For Output As #1
  atlt = Text3.Text
  Print #1, atlt
  Close #1
  Dim attjy As String
  Open "data\tctimejy.txt" For Output As #1
  attjy = Text4.Text
  Print #1, attjy
  Close #1
  Dim attjj As String
  Open "data\tctimejj.txt" For Output As #1
  attjj = Text2.Text
  Print #1, attjj
  Close #1
  Dim llq As String
  If Check1.Value = 1 Then
   Open "data\tcllq.txt" For Output As #1
   llq = "1"
   Print #1, llq
   Close #1
  Else
   Open "data\tcllq.txt" For Output As #1
   llq = "0"
   Print #1, llq
   Close #1
  End If
  Dim mgr As String
  If Check2.Value = 1 Then
   Open "data\tcmgr.txt" For Output As #1
   mgr = "1"
   Print #1, mgr
   Close #1
  Else
   Open "data\tcmgr.txt" For Output As #1
   mgr = "0"
   Print #1, mgr
   Close #1
  End If
  Unload Me
 Else
  MsgBox "不允许此操作", vbCritical, "警告"
 End If
Else
 MsgBox "请输入有效的参数", vbCritical, "提示"
End If
End Sub

Private Sub Command2_Click()
sjkz = False
Open "data\tctask1.txt" For Output As #1
Print #1, ""
Close #1
Open "data\tctask2.txt" For Output As #1
Print #1, ""
Close #1
Open "data\tctimejy.txt" For Output As #1
Print #1, ""
Close #1
Open "data\tctimejj.txt" For Output As #1
Print #1, ""
Close #1
Open "data\tcllq.txt" For Output As #1
Print #1, "0"
Close #1
Open "data\tcmgr.txt" For Output As #1
Print #1, "0"
Close #1
Unload Me
End Sub

Private Sub Command3_Click()
If sjkz = True Then
 sjkz = False
 Unload Me
Else
 MsgBox "时间控制未启用", vbCritical, "提示"
End If
End Sub

Private Sub Form_Load()
Dim lino As String
Open "data\tctask1.txt" For Input As #1
Line Input #1, lino
Text1.Text = lino
Close #1
Dim lint As String
Open "data\tctask2.txt" For Input As #1
Line Input #1, lint
Text3.Text = lint
Close #1
Dim timjy As String
Open "data\tctimejy.txt" For Input As #1
Line Input #1, timjy
Text4.Text = timjy
Close #1
Dim timjj As String
Open "data\tctimejj.txt" For Input As #1
Line Input #1, timjj
Text2.Text = timjj
Close #1
Dim tllq As String
Open "data\tcllq.txt" For Input As #1
Line Input #1, tllq
Close #1
Check1.Value = tllq
Dim tmgr As String
Open "data\tcmgr.txt" For Input As #1
Line Input #1, tmgr
Close #1
Check2.Value = tmgr
End Sub


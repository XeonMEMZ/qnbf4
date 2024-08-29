VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "软件设置"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8520
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleMode       =   0  'User
   ScaleWidth      =   8520
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "自动关机"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1695
      Left            =   6000
      TabIndex        =   21
      Top             =   2760
      Width           =   2175
      Begin VB.CommandButton Command26 
         BackColor       =   &H00FFC0C0&
         Caption         =   "禁用自动关机"
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
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command25 
         BackColor       =   &H00FFC0C0&
         Caption         =   "设置自动关机"
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
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00FFC0C0&
      Caption         =   "修改透明度"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H00FFC0C0&
      Caption         =   "设置自动任务"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00FFC0C0&
      Caption         =   "设置时间控制"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "上下课执行"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1695
      Left            =   360
      TabIndex        =   33
      Top             =   2760
      Width           =   2175
      Begin VB.CommandButton Command23 
         BackColor       =   &H00FFC0C0&
         Caption         =   "设置下课自动执行"
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
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00FFC0C0&
         Caption         =   "设置上课自动执行"
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
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00FFC0C0&
      Caption         =   "禁用自动打开U盘"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "修改主题"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command31 
      Height          =   1215
      Left            =   6960
      Picture         =   "Form4.frx":6988A
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command30 
      Height          =   1215
      Left            =   3240
      Picture         =   "Form4.frx":6D9B9
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command29 
      Height          =   1215
      Left            =   1680
      Picture         =   "Form4.frx":70D77
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command28 
      Height          =   1215
      Left            =   120
      Picture         =   "Form4.frx":7475C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "倒计时"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1095
      Left            =   6600
      TabIndex        =   24
      Top             =   5160
      Width           =   1815
      Begin VB.CommandButton Command27 
         BackColor       =   &H00FFC0C0&
         Caption         =   "修改倒计时"
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
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "U盘"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1695
      Left            =   4320
      TabIndex        =   16
      Top             =   1560
      Width           =   2175
      Begin VB.CommandButton Command16 
         BackColor       =   &H00FFC0C0&
         Caption         =   "修改U盘盘符"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "作息时间"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2895
      Left            =   4320
      TabIndex        =   14
      Top             =   3360
      Width           =   2175
      Begin VB.CommandButton Command21 
         BackColor       =   &H00FFC0C0&
         Caption         =   "改为特殊作息时间"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00FFC0C0&
         Caption         =   "修改作息时间"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "主程序功能"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3495
      Left            =   6600
      TabIndex        =   12
      Top             =   1560
      Width           =   1815
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FFC0C0&
         Caption         =   "禁用语音"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFC0C0&
         Caption         =   "禁用双进程保护"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "修改常用软件"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "值日和随机点名"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2895
      Left            =   2040
      TabIndex        =   9
      Top             =   3360
      Width           =   2175
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "临时修改值日"
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
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00FFC0C0&
         Caption         =   "修改值日"
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
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H00FFC0C0&
         Caption         =   "修改人数"
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
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "修改名单"
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
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "重启软件"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   1815
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "重启软件"
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
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "密码"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "修改控制密码"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "修改关闭密码"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "修改设置密码"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "课表"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1695
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "临时修改课表"
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
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "修改课表"
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
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6480
      Width           =   6855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      X1              =   0
      X2              =   8520
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      X1              =   0
      X2              =   8520
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form5.Show
End Sub

Private Sub Command10_Click()
Form18.Show
End Sub

Private Sub Command11_Click()
Form20.Show
End Sub

Private Sub Command12_Click()
If dt = "1" Then
 dt = "0"
 Command12.Caption = "启用双进程保护"
 Open "data\dt.txt" For Output As #1
 Print #1, "0"
 Close #1
 MsgBox "修改后关闭软件再次打开生效!", 48, "提示"
Else
 dt = "1"
 Command12.Caption = "禁用双进程保护"
 Open "data\dt.txt" For Output As #1
 Print #1, "1"
 Close #1
 MsgBox "修改后关闭软件再次打开生效!", 48, "提示"
End If
End Sub

Private Sub Command13_Click()
Form21.Show
End Sub

Private Sub Command14_Click()
If aud = "1" Then
 aud = "0"
 Command14.Caption = "启用语音"
 Open "data\aud.txt" For Output As #1
 Print #1, "0"
 Close #1
Else
 aud = "1"
 Command14.Caption = "禁用语音"
 Open "data\aud.txt" For Output As #1
 Print #1, "1"
 Close #1
End If
End Sub

Private Sub Command15_Click()
Form16.Show
End Sub

Private Sub Command16_Click()
Form17.Show
End Sub

Private Sub Command17_Click()
Call Shell("cmd /c start data\namec.txt")
End Sub

Private Sub Command18_Click()
Form19.Show
End Sub

Private Sub Command19_Click()
Dim tcp As String
Open "data\pswtc.txt" For Input As #1
Line Input #1, tcp
Close #1
If tcp = "" Then
 Form12.Show
Else
 Form11.Show
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command20_Click()
Dim udatopp As String
Open "data\udatop.txt" For Input As #1
Line Input #1, udatopp
Close #1
If udatopp = "1" Then
 Open "data\udatop.txt" For Output As #1
 Print #1, "0"
 Close #1
 Command20.Caption = "启用自动打开U盘"
ElseIf udatopp = "0" Then
 Open "data\udatop.txt" For Output As #1
 Print #1, "1"
 Close #1
 Command20.Caption = "禁用自动打开U盘"
End If
End Sub

Private Sub Command21_Click()
If fricls = 1 Then
 fricls = 0
 Command21.Caption = "改为特殊作息时间"
Else
 fricls = 1
 Command21.Caption = "改为正常作息时间"
End If
End Sub

Private Sub Command22_Click()
Form26.Show
End Sub

Private Sub Command23_Click()
Form27.Show
End Sub

Private Sub Command24_Click()
Dim tcp As String
Open "data\pswtc.txt" For Input As #1
Line Input #1, tcp
Close #1
If tcp = "" Then
 Form6.Show
Else
 Form14.Show
End If
End Sub

Private Sub Command25_Click()
Form29.Show
End Sub

Private Sub Command26_Click()
Dim zdgj As String
Open "data\zdgj.txt" For Input As #1
Line Input #1, zdgj
Close #1
If zdgj = "1" Then
 Open "data\zdgj.txt" For Output As #1
 Print #1, "0"
 Close #1
 Command26.Caption = "启用自动关机"
ElseIf zdgj = "0" Then
 Open "data\zdgj.txt" For Output As #1
 Print #1, "1"
 Close #1
 Command26.Caption = "禁用自动关机"
End If
End Sub

Private Sub Command27_Click()
Form32.Show
End Sub

Private Sub Command28_Click()
Frame1.Visible = True
Frame3.Visible = True
Frame4.Visible = True
Frame5.Visible = True
Frame6.Visible = True
Frame7.Visible = True
Frame8.Visible = True
Frame11.Visible = True

Command1.Visible = False
Command13.Visible = False

Command19.Visible = False
Command20.Visible = False
Command24.Visible = False
Frame2.Visible = False
Frame10.Visible = False
End Sub

Private Sub Command29_Click()
Frame1.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame8.Visible = False
Frame11.Visible = False

Command1.Visible = True
Command13.Visible = True

Command19.Visible = False
Command20.Visible = False
Command24.Visible = False
Frame2.Visible = False
Frame10.Visible = False
End Sub

Private Sub Command3_Click()
Form7.Show
End Sub

Private Sub Command30_Click()
Frame1.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Frame7.Visible = False
Frame8.Visible = False
Frame11.Visible = False

Command1.Visible = False
Command13.Visible = False

Command19.Visible = True
Command20.Visible = True
Command24.Visible = True
Frame2.Visible = True
Frame10.Visible = True
End Sub

Private Sub Command31_Click()
Form25.Show
End Sub

Private Sub Command4_Click()
Form8.Show
End Sub

Private Sub Command5_Click()
Form9.Show
End Sub

Private Sub Command6_Click()
Form10.Show
End Sub

Private Sub Command7_Click()
Call Shell("cmd /c start killme.bat")
Unload Me
End Sub

Private Sub Command8_Click()
Form13.Show
End Sub

Private Sub Command9_Click()
Call Shell("cmd /c start data\name.txt")
End Sub

Private Sub Form_Load()
Frame1.Visible = True
Frame3.Visible = True
Frame4.Visible = True
Frame5.Visible = True
Frame6.Visible = True
Frame7.Visible = True
Frame8.Visible = True
Frame11.Visible = True

Command1.Visible = False
Command13.Visible = False

Command19.Visible = False
Command20.Visible = False
Command24.Visible = False
Frame2.Visible = False
Frame10.Visible = False
Dim udatop As String
Open "data\udatop.txt" For Input As #1
Line Input #1, udatop
Close #1
If udatop = "1" Then
 Command20.Caption = "禁用自动打开U盘"
ElseIf udatop = "0" Then
 Command20.Caption = "启用自动打开U盘"
End If
If fricls = 0 Then
 Command21.Caption = "改为特殊作息时间"
Else
 Command21.Caption = "改为正常作息时间"
End If
Dim zdgj As String
Open "data\zdgj.txt" For Input As #1
Line Input #1, zdgj
Close #1
If zdgj = "1" Then
 Command26.Caption = "禁用自动关机"
ElseIf zdgj = "0" Then
 Command26.Caption = "启用自动关机"
End If
If dt = "1" Then
 Command12.Caption = "禁用双进程保护"
Else
 Command12.Caption = "启用双进程保护"
End If
If aud = "1" Then
 Command14.Caption = "禁用语音"
Else
 Command14.Caption = "启用语音"
End If
End Sub

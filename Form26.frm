VERSION 5.00
Begin VB.Form Form26 
   BackColor       =   &H00F9FFDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "上课自动执行"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "Form26.frx":0000
   LinkTopic       =   "Form26"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00F9FFDD&
      Caption         =   "上课自动执行"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4335
      Begin VB.OptionButton Option2 
         BackColor       =   &H00F9FFDD&
         Caption         =   "关闭软件"
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
         Left            =   2400
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00F9FFDD&
         Caption         =   "打开软件"
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
         Left            =   960
         TabIndex        =   4
         Top             =   840
         Value           =   -1  'True
         Width           =   1095
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
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "链接:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   540
      End
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
      TabIndex        =   1
      Top             =   2400
      Width           =   1815
   End
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
      TabIndex        =   0
      Top             =   2400
      Width           =   1815
   End
End
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Not Text1.Text = "" Then
 Open "data\skatlink.txt" For Output As #1
 Print #1, Text1.Text
 Close #1
 Open "data\skatoc.txt" For Output As #1
 If Option1.Value = True Then
  Print #1, "1"
 Else
  Print #1, "0"
 End If
 Close #1
 Unload Me
Else
 MsgBox "请输入有效的参数", vbCritical, "提示"
End If
End Sub

Private Sub Command2_Click()
Open "data\skatlink.txt" For Output As #1
Print #1, ""
Close #1
Open "data\skatoc.txt" For Output As #1
Print #1, ""
Close #1
Unload Me
End Sub

Private Sub Form_Load()
Dim lin As String
Open "data\skatlink.txt" For Input As #1
Line Input #1, lin
Text1.Text = lin
Close #1
Dim oc As String
Open "data\skatoc.txt" For Input As #1
Line Input #1, oc
If oc = "0" Then
 Option1.Value = False
 Option2.Value = True
Else
 Option1.Value = True
 Option2.Value = False
End If
Close #1
End Sub

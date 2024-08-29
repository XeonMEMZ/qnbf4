VERSION 5.00
Begin VB.Form Form20 
   BackColor       =   &H00F9FFDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÐÞ¸Ä³£ÓÃÈí¼þ"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4935
   Icon            =   "Form20.frx":0000
   LinkTopic       =   "Form20"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   4935
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame Frame1 
      BackColor       =   &H00F9FFDD&
      Caption         =   "³£ÓÃÈí¼þ1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   4695
      Begin VB.TextBox Text2 
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
         Left            =   720
         TabIndex        =   19
         Top             =   840
         Width           =   3855
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
         Left            =   720
         TabIndex        =   18
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Í¼±ê:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Á´½Ó:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F9FFDD&
      Caption         =   "³£ÓÃÈí¼þ2"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   4695
      Begin VB.TextBox Text3 
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
         Left            =   720
         TabIndex        =   14
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox Text4 
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
         Left            =   720
         TabIndex        =   13
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00F9FFDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Á´½Ó:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00F9FFDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Í¼±ê:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00F9FFDD&
      Caption         =   "³£ÓÃÈí¼þ3"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   4695
      Begin VB.TextBox Text5 
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
         Left            =   720
         TabIndex        =   9
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox Text6 
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
         Left            =   720
         TabIndex        =   8
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00F9FFDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Á´½Ó:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00F9FFDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Í¼±ê:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00F9FFDD&
      Caption         =   "³£ÓÃÈí¼þ4"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
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
      Top             =   4920
      Width           =   4695
      Begin VB.TextBox Text7 
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
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox Text8 
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
         Left            =   720
         TabIndex        =   3
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00F9FFDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Á´½Ó:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00F9FFDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Í¼±ê:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "È¡Ïû"
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
      Left            =   2760
      TabIndex        =   1
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "È·¶¨"
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
      TabIndex        =   0
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "×¢:Í¼±ê¸ñÊ½Îª .jpg »ò .ico"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   22
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Not Text1.Text = "" And Not Text3.Text = "" And Not Text5.Text = "" And Not Text7.Text = "" Then
 If Not Text2.Text = "" And Not Text4.Text = "" And Not Text6.Text = "" And Not Text8.Text = "" Then
  Open "data\cy1l.txt" For Output As #1
  Print #1, Text1.Text
  Close #1
  Open "data\cy1t.txt" For Output As #1
  Print #1, Text2.Text
  Close #1
  Open "data\cy2l.txt" For Output As #1
  Print #1, Text3.Text
  Close #1
  Open "data\cy2t.txt" For Output As #1
  Print #1, Text4.Text
  Close #1
  Open "data\cy3l.txt" For Output As #1
  Print #1, Text5.Text
  Close #1
  Open "data\cy3t.txt" For Output As #1
  Print #1, Text6.Text
  Close #1
  Open "data\cy4l.txt" For Output As #1
  Print #1, Text7.Text
  Close #1
  Open "data\cy4t.txt" For Output As #1
  Print #1, Text8.Text
  Close #1
  Unload Me
 Else
  MsgBox "ÎÞÐ§Í¼±ê", vbCritical, "ÌáÊ¾"
 End If
Else
 MsgBox "ÎÞÐ§ÎÄ¼þ", vbCritical, "ÌáÊ¾"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim cy1l As String
Open "data\cy1l.txt" For Input As #1
Line Input #1, cy1l
Close #1
Dim cy1t As String
Open "data\cy1t.txt" For Input As #1
Line Input #1, cy1t
Close #1
Dim cy2l As String
Open "data\cy2l.txt" For Input As #1
Line Input #1, cy2l
Close #1
Dim cy2t As String
Open "data\cy2t.txt" For Input As #1
Line Input #1, cy2t
Close #1
Dim cy3l As String
Open "data\cy3l.txt" For Input As #1
Line Input #1, cy3l
Close #1
Dim cy3t As String
Open "data\cy3t.txt" For Input As #1
Line Input #1, cy3t
Close #1
Dim cy4l As String
Open "data\cy4l.txt" For Input As #1
Line Input #1, cy4l
Close #1
Dim cy4t As String
Open "data\cy4t.txt" For Input As #1
Line Input #1, cy4t
Close #1
Text1.Text = cy1l
Text2.Text = cy1t
Text3.Text = cy2l
Text4.Text = cy2t
Text5.Text = cy3l
Text6.Text = cy3t
Text7.Text = cy4l
Text8.Text = cy4t
End Sub


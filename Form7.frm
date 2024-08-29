VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00F9FFDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÐÞ¸Ä¿Î±í"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4560
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
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
      TabIndex        =   14
      Top             =   3000
      Width           =   3615
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
      TabIndex        =   13
      Top             =   3480
      Width           =   3615
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
      Left            =   2520
      TabIndex        =   12
      Top             =   3960
      Width           =   1815
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
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   1815
   End
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
      TabIndex        =   10
      Top             =   2520
      Width           =   3615
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
      TabIndex        =   8
      Top             =   2040
      Width           =   3615
   End
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
      TabIndex        =   6
      Top             =   1560
      Width           =   3615
   End
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
      TabIndex        =   4
      Top             =   1080
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
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E7E5CD&
      BackStyle       =   0  'Transparent
      Caption         =   "ÖÜÁù"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E7E5CD&
      BackStyle       =   0  'Transparent
      Caption         =   "ÖÜÈÕ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E7E5CD&
      BackStyle       =   0  'Transparent
      Caption         =   "ÖÜÎå"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E7E5CD&
      BackStyle       =   0  'Transparent
      Caption         =   "ÖÜËÄ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E7E5CD&
      BackStyle       =   0  'Transparent
      Caption         =   "ÖÜÈý"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E7E5CD&
      BackStyle       =   0  'Transparent
      Caption         =   "ÖÜ¶þ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E7E5CD&
      BackStyle       =   0  'Transparent
      Caption         =   "ÖÜÒ»"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÐÞ¸Ä¿Î±í"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Not Text1.Text = "" And Not Text2.Text = "" And Not Text3.Text = "" And Not Text4.Text = "" And Not Text5.Text = "" Then
 Dim kbo As String
 Open "data\z1.txt" For Output As #1
 kbo = Text1.Text
 Print #1, kbo
 Close #1
 Open "data\z2.txt" For Output As #1
 kbo = Text2.Text
 Print #1, kbo
 Close #1
 Open "data\z3.txt" For Output As #1
 kbo = Text3.Text
 Print #1, kbo
 Close #1
 Open "data\z4.txt" For Output As #1
 kbo = Text4.Text
 Print #1, kbo
 Close #1
 Open "data\z5.txt" For Output As #1
 kbo = Text5.Text
 Print #1, kbo
 Close #1
 Open "data\z6.txt" For Output As #1
 kbo = Text7.Text
 Print #1, kbo
 Close #1
 Open "data\z7.txt" For Output As #1
 kbo = Text6.Text
 Print #1, kbo
 Close #1
 Unload Me
Else
 MsgBox "¿Î±í²»ÄÜÎª¿Õ", vbCritical, "ÌáÊ¾"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim kbi As String
Open "data\z1.txt" For Input As #1
Line Input #1, kbi
Text1.Text = kbi
Close #1
Open "data\z2.txt" For Input As #1
Line Input #1, kbi
Text2.Text = kbi
Close #1
Open "data\z3.txt" For Input As #1
Line Input #1, kbi
Text3.Text = kbi
Close #1
Open "data\z4.txt" For Input As #1
Line Input #1, kbi
Text4.Text = kbi
Close #1
Open "data\z5.txt" For Input As #1
Line Input #1, kbi
Text5.Text = kbi
Close #1
Open "data\z6.txt" For Input As #1
Line Input #1, kbi
Text7.Text = kbi
Close #1
Open "data\z7.txt" For Input As #1
Line Input #1, kbi
Text6.Text = kbi
Close #1
End Sub

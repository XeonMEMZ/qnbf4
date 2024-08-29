VERSION 5.00
Begin VB.Form Form27 
   BackColor       =   &H00F9FFDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÏÂ¿Î×Ô¶¯Ö´ÐÐ"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "Form27.frx":0000
   LinkTopic       =   "Form27"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
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
      Left            =   3120
      TabIndex        =   8
      Top             =   1800
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00F9FFDD&
      Caption         =   "ÌáÐÑ²ÁºÚ°å"
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
      Left            =   480
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F9FFDD&
      Caption         =   "ÏÂ¿Î×Ô¶¯Ö´ÐÐ"
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
      Top             =   240
      Width           =   4335
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
         TabIndex        =   5
         Top             =   360
         Width           =   3495
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00F9FFDD&
         Caption         =   "´ò¿ªÈí¼þ"
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
         Left            =   960
         TabIndex        =   4
         Top             =   840
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00F9FFDD&
         Caption         =   "¹Ø±ÕÈí¼þ"
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
         Left            =   2400
         TabIndex        =   3
         Top             =   840
         Width           =   1095
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
         TabIndex        =   6
         Top             =   360
         Width           =   540
      End
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
      TabIndex        =   1
      Top             =   2400
      Width           =   1815
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
      TabIndex        =   0
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ãë"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ÌáÐÑÊ±³¤:"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "Form27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
 Label2.Enabled = True
 Text2.Enabled = True
Else
 Label2.Enabled = False
 Text2.Enabled = False
End If
End Sub

Private Sub Command1_Click()
If Not Text1.Text = "" Then
 Open "data\xkatlink.txt" For Output As #1
 Print #1, Text1.Text
 Close #1
 Open "data\xkchbtime.txt" For Output As #1
 Print #1, Text2.Text
 Close #1
 Open "data\xkatoc.txt" For Output As #1
 If Option1.Value = True Then
  Print #1, "1"
 Else
  Print #1, "0"
 End If
 Close #1
 If Check1.Value = 1 Then
  Open "data\xkchb.txt" For Output As #1
  Print #1, "1"
  Close #1
 Else
  Open "data\xkchb.txt" For Output As #1
  Print #1, "0"
  Close #1
 End If
 Unload Me
Else
 Open "data\xkchbtime.txt" For Output As #1
 Print #1, Text2.Text
 Close #1
 If Check1.Value = 1 Then
  Open "data\xkchb.txt" For Output As #1
  Print #1, "1"
  Close #1
 Else
  Open "data\xkchb.txt" For Output As #1
  Print #1, "0"
  Close #1
 End If
 Unload Me
End If
End Sub

Private Sub Command2_Click()
Open "data\xkatlink.txt" For Output As #1
Print #1, ""
Close #1
Open "data\xkatoc.txt" For Output As #1
Print #1, ""
Close #1
Open "data\xkchb.txt" For Output As #1
Print #1, "1"
Close #1
Open "data\xkchbtime.txt" For Output As #1
Print #1, "120"
Close #1
Unload Me
End Sub

Private Sub Form_Load()
Dim lin As String
Open "data\xkatlink.txt" For Input As #1
Line Input #1, lin
Text1.Text = lin
Close #1
Dim oc As String
Open "data\xkatoc.txt" For Input As #1
Line Input #1, oc
If oc = "0" Then
 Option1.Value = False
 Option2.Value = True
Else
 Option1.Value = True
 Option2.Value = False
End If
Close #1
Dim chb As String
Open "data\xkchb.txt" For Input As #1
Line Input #1, chb
Close #1
Check1.Value = chb
Dim chbtm As String
Open "data\xkchbtime.txt" For Input As #1
Line Input #1, chbtm
Close #1
Text2.Text = chbtm
End Sub

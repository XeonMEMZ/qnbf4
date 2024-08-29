VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00F9FFDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÉèÖÃ×Ô¶¯ÈÎÎñ"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4935
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   4935
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame Frame4 
      BackColor       =   &H00F9FFDD&
      Caption         =   "×Ô¶¯ÈÎÎñ4"
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
      TabIndex        =   23
      Top             =   4440
      Width           =   4695
      Begin VB.OptionButton Option8 
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
         Left            =   3360
         TabIndex        =   27
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option7 
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
         Left            =   2160
         TabIndex        =   26
         Top             =   840
         Value           =   -1  'True
         Width           =   1095
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
         TabIndex        =   25
         Top             =   840
         Width           =   1335
      End
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
         TabIndex        =   24
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00F9FFDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Ê±¼ä:"
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
         TabIndex        =   29
         Top             =   840
         Width           =   540
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
         TabIndex        =   28
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00F9FFDD&
      Caption         =   "×Ô¶¯ÈÎÎñ3"
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
      TabIndex        =   16
      Top             =   3000
      Width           =   4695
      Begin VB.OptionButton Option6 
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
         Left            =   3360
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
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
         Left            =   2160
         TabIndex        =   19
         Top             =   840
         Value           =   -1  'True
         Width           =   1095
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
         TabIndex        =   18
         Top             =   840
         Width           =   1335
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
         TabIndex        =   17
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00F9FFDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Ê±¼ä:"
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
         TabIndex        =   22
         Top             =   840
         Width           =   540
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
         TabIndex        =   21
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F9FFDD&
      Caption         =   "×Ô¶¯ÈÎÎñ2"
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
      TabIndex        =   9
      Top             =   1560
      Width           =   4695
      Begin VB.OptionButton Option4 
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
         Left            =   3360
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
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
         Left            =   2160
         TabIndex        =   12
         Top             =   840
         Value           =   -1  'True
         Width           =   1095
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
         TabIndex        =   11
         Top             =   840
         Width           =   1335
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
         TabIndex        =   10
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00F9FFDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Ê±¼ä:"
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
         TabIndex        =   14
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F9FFDD&
      Caption         =   "×Ô¶¯ÈÎÎñ1"
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
      Top             =   120
      Width           =   4695
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
         TabIndex        =   6
         Top             =   360
         Width           =   3855
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
         TabIndex        =   5
         Top             =   840
         Width           =   1335
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
         Left            =   2160
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
         Left            =   3360
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
         TabIndex        =   8
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ê±¼ä:"
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
         TabIndex        =   7
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "É¾³ý×Ô¶¯ÈÎÎñ"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "´´½¨×Ô¶¯ÈÎÎñ"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   1695
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Not Text1.Text = "" And Not Text2.Text = "" Then
 zdrw = True
 Dim atl1 As String
 Open "data\atlink1.txt" For Output As #1
 atl1 = Text1.Text
 Print #1, atl1
 Close #1
 Dim att1 As String
 Open "data\attime1.txt" For Output As #1
 att1 = Text2.Text
 Print #1, att1
 Close #1
 Open "data\atoc1.txt" For Output As #1
 If Option1.Value = True Then
  Print #1, "1"
 Else
  Print #1, "0"
 End If
 Close #1
 Dim atl2 As String
 Open "data\atlink2.txt" For Output As #1
 atl2 = Text3.Text
 Print #1, atl2
 Close #1
 Dim att2 As String
 Open "data\attime2.txt" For Output As #1
 att2 = Text4.Text
 Print #1, att2
 Close #1
 Open "data\atoc2.txt" For Output As #1
 If Option3.Value = True Then
  Print #1, "1"
 Else
  Print #1, "0"
 End If
 Close #1
 Dim atl3 As String
 Open "data\atlink3.txt" For Output As #1
 atl3 = Text5.Text
 Print #1, atl3
 Close #1
 Dim att3 As String
 Open "data\attime3.txt" For Output As #1
 att3 = Text6.Text
 Print #1, att3
 Close #1
 Open "data\atoc3.txt" For Output As #1
 If Option5.Value = True Then
  Print #1, "1"
 Else
  Print #1, "0"
 End If
 Close #1
 Dim atl4 As String
 Open "data\atlink4.txt" For Output As #1
 atl4 = Text7.Text
 Print #1, atl4
 Close #1
 Dim att4 As String
 Open "data\attime4.txt" For Output As #1
 att4 = Text8.Text
 Print #1, att4
 Close #1
 Open "data\atoc4.txt" For Output As #1
 If Option7.Value = True Then
  Print #1, "1"
 Else
  Print #1, "0"
 End If
 Close #1
 Unload Me
Else
 MsgBox "ÇëÊäÈëÓÐÐ§µÄ²ÎÊý", vbCritical, "ÌáÊ¾"
End If
End Sub

Private Sub Command2_Click()
zdrw = False
Open "data\atlink1.txt" For Output As #1
Print #1, ""
Close #1
Open "data\attime1.txt" For Output As #1
Print #1, ""
Close #1
Open "data\atoc1.txt" For Output As #1
Print #1, ""
Close #1
Open "data\atlink2.txt" For Output As #1
Print #1, ""
Close #1
Open "data\attime2.txt" For Output As #1
Print #1, ""
Close #1
Open "data\atoc2.txt" For Output As #1
Print #1, ""
Close #1
Open "data\atlink3.txt" For Output As #1
Print #1, ""
Close #1
Open "data\attime3.txt" For Output As #1
Print #1, ""
Close #1
Open "data\atoc3.txt" For Output As #1
Print #1, ""
Close #1
Open "data\atlink4.txt" For Output As #1
Print #1, ""
Close #1
Open "data\attime4.txt" For Output As #1
Print #1, ""
Close #1
Open "data\atoc4.txt" For Output As #1
Print #1, ""
Close #1
Unload Me
End Sub

Private Sub Form_Load()
Dim lin1 As String
Open "data\atlink1.txt" For Input As #1
Line Input #1, lin1
Text1.Text = lin1
Close #1
Dim tim1 As String
Open "data\attime1.txt" For Input As #1
Line Input #1, tim1
Text2.Text = tim1
Close #1
Dim oc1 As String
Open "data\atoc1.txt" For Input As #1
Line Input #1, oc1
If oc1 = "0" Then
 Option1.Value = False
 Option2.Value = True
Else
 Option1.Value = True
 Option2.Value = False
End If
Close #1
Dim lin2 As String
Open "data\atlink2.txt" For Input As #1
Line Input #1, lin2
Text3.Text = lin2
Close #1
Dim tim2 As String
Open "data\attime2.txt" For Input As #1
Line Input #1, tim2
Text4.Text = tim2
Close #1
Dim oc2 As String
Open "data\atoc2.txt" For Input As #1
Line Input #1, oc2
If oc2 = "0" Then
 Option3.Value = False
 Option4.Value = True
Else
 Option3.Value = True
 Option4.Value = False
End If
Close #1
Dim lin3 As String
Open "data\atlink3.txt" For Input As #1
Line Input #1, lin3
Text5.Text = lin3
Close #1
Dim tim3 As String
Open "data\attime3.txt" For Input As #1
Line Input #1, tim3
Text6.Text = tim3
Close #1
Dim oc3 As String
Open "data\atoc3.txt" For Input As #1
Line Input #1, oc3
If oc3 = "0" Then
 Option5.Value = False
 Option6.Value = True
Else
 Option5.Value = True
 Option6.Value = False
End If
Close #1
Dim lin4 As String
Open "data\atlink4.txt" For Input As #1
Line Input #1, lin4
Text7.Text = lin4
Close #1
Dim tim4 As String
Open "data\attime4.txt" For Input As #1
Line Input #1, tim4
Text8.Text = tim4
Close #1
Dim oc4 As String
Open "data\atoc4.txt" For Input As #1
Line Input #1, oc4
If oc4 = "0" Then
 Option7.Value = False
 Option8.Value = True
Else
 Option7.Value = True
 Option8.Value = False
End If
Close #1
End Sub

Private Sub Option1_Click()
Option1.Value = True
Option2.Value = False
End Sub

Private Sub Option2_Click()
Option1.Value = False
Option2.Value = True
End Sub

Private Sub Option3_Click()
Option3.Value = True
Option4.Value = False
End Sub

Private Sub Option4_Click()
Option3.Value = False
Option4.Value = True
End Sub

Private Sub Option5_Click()
Option5.Value = True
Option6.Value = False
End Sub

Private Sub Option6_Click()
Option5.Value = False
Option6.Value = True
End Sub

Private Sub Option7_Click()
Option7.Value = True
Option8.Value = False
End Sub

Private Sub Option8_Click()
Option7.Value = False
Option8.Value = True
End Sub


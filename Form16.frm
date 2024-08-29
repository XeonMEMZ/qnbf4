VERSION 5.00
Begin VB.Form Form16 
   BackColor       =   &H00F9FFDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÐÞ¸Ä×÷Ï¢Ê±¼ä"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5160
   Icon            =   "Form16.frx":0000
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   5160
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame Frame3 
      BackColor       =   &H00F9FFDD&
      Caption         =   "ÌØÊâÊ±¼ä"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   20
      Top             =   5160
      Width           =   4935
      Begin VB.OptionButton Option3 
         BackColor       =   &H00F9FFDD&
         Caption         =   "ÖÜÈÕ·ÅÑ§"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3480
         TabIndex        =   47
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00F9FFDD&
         Caption         =   "ÖÜÁù·ÅÑ§"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1920
         TabIndex        =   46
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00F9FFDD&
         Caption         =   "ÖÜÎå·ÅÑ§"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   360
         TabIndex        =   45
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text15 
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3720
         TabIndex        =   34
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3720
         TabIndex        =   33
         Top             =   840
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
         Height          =   390
         Left            =   1200
         TabIndex        =   22
         Top             =   840
         Width           =   1095
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
         Height          =   390
         Left            =   1200
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚÎå½ÚÏÂ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   36
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚÁù½ÚÏÂ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   35
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚÁù½ÚÉÏ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚÎå½ÚÉÏ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   1095
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
      Left            =   2880
      TabIndex        =   19
      Top             =   7200
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
      Left            =   480
      TabIndex        =   18
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F9FFDD&
      Caption         =   "ÏÂÎç"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   4935
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3720
         TabIndex        =   40
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text19 
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3720
         TabIndex        =   39
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3720
         TabIndex        =   38
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3720
         TabIndex        =   37
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1200
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1200
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1200
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1200
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚ°Ë½ÚÏÂ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   44
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚÆß½ÚÏÂ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   43
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚÁù½ÚÏÂ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   42
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚÎå½ÚÏÂ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   41
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚÎå½ÚÉÏ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚÁù½ÚÉÏ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚÆß½ÚÉÏ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚ°Ë½ÚÉÏ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F9FFDD&
      Caption         =   "ÉÏÎç"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3720
         TabIndex        =   28
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3720
         TabIndex        =   27
         Top             =   840
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
         Height          =   390
         Left            =   3720
         TabIndex        =   26
         Top             =   1320
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
         Height          =   390
         Left            =   3720
         TabIndex        =   25
         Top             =   1800
         Width           =   1095
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
         Height          =   390
         Left            =   1200
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
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
         Height          =   390
         Left            =   1200
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
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
         Height          =   390
         Left            =   1200
         TabIndex        =   4
         Top             =   840
         Width           =   1095
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
         Height          =   390
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚÒ»½ÚÏÂ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   32
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚ¶þ½ÚÏÂ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚÈý½ÚÏÂ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚËÄ½ÚÏÂ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚËÄ½ÚÉÏ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚÈý½ÚÉÏ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚ¶þ½ÚÉÏ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "µÚÒ»½ÚÉÏ¿Î:"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Not Text1.Text = "" And Not Text2.Text = "" And Not Text3.Text = "" And Not Text4.Text = "" And Not Text5.Text = "" And Not Text6.Text = "" And Not Text7.Text = "" And Not Text8.Text = "" And Not Text9.Text = "" And Not Text10.Text = "" And Not Text11.Text = "" And Not Text12.Text = "" And Not Text13.Text = "" And Not Text14.Text = "" And Not Text15.Text = "" And Not Text16.Text = "" And Not Text17.Text = "" And Not Text18.Text = "" And Not Text19.Text = "" And Not Text20.Text = "" Then
 Dim alltime As String
 alltime = "zm1s" & " " & Text1.Text & " " & "zm1x" & " " & Text11.Text & " " & "zm2s" & " " & Text3.Text & " " & "zm2x" & " " & Text9.Text & " " & "zm3s" & " " & Text5.Text & " " & "zm3x" & " " & Text8.Text & " " & "zm4s" & " " & Text7.Text & " " & "zm4x" & " " & Text6.Text & " " & "zm5s" & " " & Text16.Text & " " & "zm5x" & " " & Text17.Text & " " & "zm6s" & " " & Text14.Text & " " & "zm6x" & " " & Text18.Text & " " & "zm7s" & " " & Text12.Text & " " & "zm7x" & " " & Text19.Text & " " & "zm8s" & " " & Text10.Text & " " & "zm8x" & " " & Text20.Text & " " & "zf5s" & " " & Text2.Text & " " & "zf5x" & " " & Text15.Text & " " & "zf6s" & " " & Text4.Text & " " & "zf6x" & " " & Text13.Text & " "
 Open "data\timelist.txt" For Output As #1
 Print #1, alltime
 Close #1
 If Option1.Value = True Then
  Open "data\fx.txt" For Output As #1
  Print #1, "5"
  Close #1
 ElseIf Option2.Value = True Then
  Open "data\fx.txt" For Output As #1
  Print #1, "6"
  Close #1
 ElseIf Option3.Value = True Then
  Open "data\fx.txt" For Output As #1
  Print #1, "7"
  Close #1
 End If
 Dim fxs As String
 Open "data\fx.txt" For Input As #1
 Line Input #1, fxs
 Close #1
 If fxs = "5" Then
  If Weekday(Date, 2) = 5 Then
   fricls = 1
  Else
   fricls = 0
  End If
 ElseIf fxs = "6" Then
  If Weekday(Date, 2) = 6 Then
   fricls = 1
  Else
   fricls = 0
  End If
 ElseIf fxs = "7" Then
  If Weekday(Date, 2) = 7 Then
   fricls = 1
  Else
   fricls = 0
  End If
 End If
 Unload Me
Else
 MsgBox "Ê±¼ä²»ÄÜÎª¿Õ", vbCritical, "ÌáÊ¾"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = zxtime("zm1s")
Text11.Text = zxtime("zm1x")
Text3.Text = zxtime("zm2s")
Text9.Text = zxtime("zm2x")
Text5.Text = zxtime("zm3s")
Text8.Text = zxtime("zm3x")
Text7.Text = zxtime("zm4s")
Text6.Text = zxtime("zm4x")
Text16.Text = zxtime("zm5s")
Text17.Text = zxtime("zm5x")
Text14.Text = zxtime("zm6s")
Text18.Text = zxtime("zm6x")
Text12.Text = zxtime("zm7s")
Text19.Text = zxtime("zm7x")
Text10.Text = zxtime("zm8s")
Text20.Text = zxtime("zm8x")
Text2.Text = zxtime("zf5s")
Text15.Text = zxtime("zf5x")
Text4.Text = zxtime("zf6s")
Text13.Text = zxtime("zf6x")
Dim fxs As String
Open "data\fx.txt" For Input As #1
Line Input #1, fxs
Close #1
If fxs = "5" Then
 Option1.Value = True
 Option2.Value = False
 Option3.Value = False
ElseIf fxs = "6" Then
 Option1.Value = False
 Option2.Value = True
 Option3.Value = False
ElseIf fxs = "7" Then
 Option1.Value = False
 Option2.Value = False
 Option3.Value = True
End If
End Sub

Public Function zxtime(s$) As String
Dim timelist As String
Open "data\timelist.txt" For Input As #1
Line Input #1, timelist
Close #1
zxtime = Trim(Mid(timelist, InStr(timelist, CStr(s)) + 4, 9))
End Function

Private Sub Option1_Click()
Option1.Value = True
Option2.Value = False
Option3.Value = False
End Sub

Private Sub Option2_Click()
Option1.Value = False
Option2.Value = True
Option3.Value = False
End Sub

Private Sub Option3_Click()
Option1.Value = False
Option2.Value = False
Option3.Value = True
End Sub

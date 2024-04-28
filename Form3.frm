VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H8000000B&
   Caption         =   "Form3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   FillColor       =   &H00FFFF80&
   LinkTopic       =   "Form3"
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ENTER TO QUIZ"
      BeginProperty Font 
         Name            =   "Mongolian Baiti"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11640
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   7560
      Width           =   4575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "--> A team cannot pass their question to other team."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3960
      TabIndex        =   9
      Top             =   5880
      Width           =   12975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "-->  If you know the answer please get up and answer the questions."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3960
      TabIndex        =   8
      Top             =   5040
      Width           =   12975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "--> One team should not support to another team,Here there is no partiality."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3960
      TabIndex        =   7
      Top             =   4320
      Width           =   13335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00400040&
      Caption         =   "     All The Best"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   4920
      TabIndex        =   6
      Top             =   7680
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "--> Select your team first and then answer your question."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   3960
      TabIndex        =   4
      Top             =   3480
      Width           =   12975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "--> For every question, particular team will have 60 Seconds of time."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3960
      TabIndex        =   5
      Top             =   2760
      Width           =   13095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "--> Each question carries one mark."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   2040
      Width           =   12975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "      RULES TO BE FOLLOW"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   8535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "--> There are 32 questions in the Quiz."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   1320
      Width           =   12975
   End
   Begin VB.Image Image1 
      Height          =   10215
      Left            =   -240
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form4.Show
End Sub


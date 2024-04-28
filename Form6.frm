VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form6"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   14760
      TabIndex        =   9
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0FF&
      Caption         =   "TEAM C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14760
      TabIndex        =   8
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF00&
      Caption         =   "TEAM B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11400
      TabIndex        =   7
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H000080FF&
      Caption         =   "TEAM A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   8040
      TabIndex        =   6
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "THANK YOU FOR ATTENDING GROUP 2 QUIZ"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   5
      Top             =   8400
      Width           =   12135
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "(Admin) = RAJEEV M M "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13320
      TabIndex        =   3
      Top             =   8520
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   4890
      Left            =   14400
      Picture         =   "Form6.frx":0000
      Top             =   3720
      Width           =   4530
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   11520
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "DEVELOPED  BY----------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   2
      Top             =   5280
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8160
      TabIndex        =   1
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "                     YOUR FINAL SCORE"
      BeginProperty Font 
         Name            =   "NSimSun"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   7455
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form6.Hide

Form7.Show

End Sub

Private Sub Form_Load()
Label2.Caption = Form4.Text1.Text
Label6.Caption = Form4.Text2.Text
Label10.Caption = Form4.Text4.Text
End Sub

Private Sub Image2_Click()

End Sub

VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000011&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4920
      TabIndex        =   5
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "clear"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8760
      TabIndex        =   6
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   8640
      TabIndex        =   4
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   8640
      TabIndex        =   3
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Image Image3 
      Height          =   3000
      Left            =   240
      Picture         =   "Form2.frx":0000
      Top             =   6480
      Width           =   3000
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000014&
      Caption         =   "   Student Login"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   720
      Picture         =   "Form2.frx":1D502
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   10335
      Left            =   0
      Picture         =   "Form2.frx":1F514
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Dim username As String
Dim password As String
username = "nagarjuna"
password = "rajeev@123"
If (username = Text1.Text And password = Text2.Text) Then
MsgBox "login successful......"
Form3.Show
Form1.Hide


Else
MsgBox "sorry..... login failed....try again...."
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Image3_Click()
Form1.Show

End Sub

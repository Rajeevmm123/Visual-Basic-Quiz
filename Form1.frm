VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "NAGARJUNA COLLEGE OF MANAGEMENT STUDIES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   960
      Width           =   9495
   End
   Begin VB.Image Image4 
      Height          =   2520
      Left            =   14520
      Picture         =   "Form1.frx":0B4B
      Top             =   0
      Width           =   4500
   End
   Begin VB.Image Image1 
      Height          =   2565
      Left            =   120
      Picture         =   "Form1.frx":3463
      Top             =   120
      Width           =   3180
   End
   Begin VB.Image Image5 
      Height          =   10215
      Left            =   -480
      Picture         =   "Form1.frx":429C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19440
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   8400
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   11415
      Left            =   0
      Picture         =   "Form1.frx":F164
      Stretch         =   -1  'True
      Top             =   0
      Width           =   21135
   End
   Begin WMPLibCtl.WindowsMediaPlayer mp 
      Height          =   4455
      Left            =   4800
      TabIndex        =   1
      Top             =   2280
      Width           =   8055
      URL             =   "C:\Users\Dell\Videos\Captures\InShot_20230804_202208020.mp4"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   -1  'True
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   -1  'True
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   14208
      _cy             =   7858
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form2.Show

End Sub


Private Sub Picture1_Click()
Form2.Show
End Sub

Private Sub Image1_Click()
Form2.Show


End Sub

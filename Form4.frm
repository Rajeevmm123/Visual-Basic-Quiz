VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form4"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Perpetua Titling MT"
      Size            =   48
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "<<< PREVIOUS"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   840
      TabIndex        =   23
      Top             =   8400
      Width           =   4215
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Option2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8640
      TabIndex        =   2
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   1305
      Left            =   2880
      TabIndex        =   22
      Text            =   "0"
      Top             =   8280
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox Text3 
      Height          =   1305
      Left            =   18300
      TabIndex        =   19
      Top             =   8760
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1935
      Left            =   480
      MousePointer    =   2  'Cross
      TabIndex        =   13
      Top             =   1200
      Width           =   14775
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Team C"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   11160
         TabIndex        =   21
         Top             =   720
         Width           =   3375
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFF00&
         Caption         =   "team B"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   6120
         TabIndex        =   15
         Top             =   720
         Width           =   3135
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H000080FF&
         Caption         =   "team a"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   2895
      End
   End
   Begin VB.TextBox Text1 
      Height          =   1305
      Left            =   240
      TabIndex        =   16
      Text            =   "0"
      Top             =   8280
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox Text2 
      Height          =   1305
      Left            =   1560
      TabIndex        =   17
      Text            =   "0"
      Top             =   8280
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Check Answer"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   6360
      TabIndex        =   12
      Top             =   8400
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT >>>"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12000
      TabIndex        =   5
      Top             =   8400
      Width           =   5055
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   17760
      Top             =   6360
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Option4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8640
      TabIndex        =   4
      Top             =   6600
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Option1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   1815
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Option3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   495
      Left            =   7920
      TabIndex        =   20
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "QUIZ"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1215
      Left            =   6960
      TabIndex        =   18
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   72
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2250
      Left            =   15720
      TabIndex        =   7
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      TabIndex        =   10
      Top             =   6600
      Width           =   5535
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10680
      TabIndex        =   11
      Top             =   6840
      Width           =   5295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10680
      TabIndex        =   9
      Top             =   4920
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1680
      TabIndex        =   0
      Top             =   3360
      Width           =   17175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      TabIndex        =   8
      Top             =   4920
      Width           =   5535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   10215
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   360
      Width           =   18975
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rec As DAO.Recordset
Dim score As Variant


Private Sub Command1_Click()
If rec.EOF = True Then
Form10.Show
Form4.Hide
Else
Option1.Visible = True
Option2.Visible = True
Option3.Visible = True
Option4.Visible = True
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False

 Label3.Caption = rec.Fields("slno")
Label2.Caption = rec.Fields("Question")
Label5.Caption = rec.Fields("option1")
Label6.Caption = rec.Fields("option2")
Label7.Caption = rec.Fields("option3")
Label8.Caption = rec.Fields("option4")
Text3.Text = rec.Fields("answer")
Timer1.Enabled = True
rec.MoveNext
Label4.Caption = 61

Label4.Caption = Label4.Caption - 1
If Label4.Caption = 0 Then
Timer1.Enabled = False
MsgBox "your time is over!", vbInformation, "message"
rec.MoveNext


End If
End If
End Sub



Private Sub Command3_Click()
If rec.BOF = True Then
Form10.Show
Form4.Hide
Else
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False

 Label3.Caption = rec.Fields("slno")
Label2.Caption = rec.Fields("Question")
Label5.Caption = rec.Fields("option1")
Label6.Caption = rec.Fields("option2")
Label7.Caption = rec.Fields("option3")
Label8.Caption = rec.Fields("option4")
Text3.Text = rec.Fields("answer")
Timer1.Enabled = False
rec.MovePrevious
Label4.Caption = 61

Label4.Caption = Label4.Caption - 1
If Label4.Caption = 0 Then
Timer1.Enabled = False
MsgBox "your time is over!", vbInformation, "message"
rec.MovePrevious


End If
End If
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
If Option5.Value = True Then


If Label3.Caption = 1 Or Label3.Caption = 7 Or Label3.Caption = 10 Or Label3.Caption = 13 Or Label3.Caption = 22 Or Label3.Caption = 24 Or Label3.Caption = 29 Then
If Option1.Value = True Then
Text1.Text = Text1.Text + 1
Form7.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
Else
Form5.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
End If
End If

If Label3.Caption = 2 Or Label3.Caption = 5 Or Label3.Caption = 9 Or Label3.Caption = 11 Or Label3.Caption = 14 Or Label3.Caption = 20 Or Label3.Caption = 21 Or Label3.Caption = 25 Or Label3.Caption = 28 Or Label3.Caption = 30 Then
If Option2.Value = True Then
Text1.Text = Text1.Text + 1
Form7.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
Else
Form5.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
End If
End If

If Label3.Caption = 3 Or Label3.Caption = 4 Or Label3.Caption = 6 Or Label3.Caption = 16 Or Label3.Caption = 17 Or Label3.Caption = 23 Then
If Option3.Value = True Then
Text1.Text = Text1.Text + 1
Form7.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
Else
Form5.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
End If
End If

If Label3.Caption = 8 Or Label3.Caption = 12 Or Label3.Caption = 15 Or Label3.Caption = 18 Or Label3.Caption = 19 Or Label3.Caption = 26 Or Label3.Caption = 27 Or Label3.Caption = 31 Or Label3.Caption = 32 Then
If Option4.Value = True Then
Text1.Text = Text1.Text + 1
Form7.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False

Else
Form5.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
End If
End If

End If



''' team b

If Option6.Value = True Then

If Label3.Caption = 1 Or Label3.Caption = 7 Or Label3.Caption = 10 Or Label3.Caption = 13 Or Label3.Caption = 22 Or Label3.Caption = 24 Or Label3.Caption = 29 Then
If Option1.Value = True Then
Text2.Text = Text2.Text + 1
Form7.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
Else
Form5.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
End If
End If

If Label3.Caption = 2 Or Label3.Caption = 5 Or Label3.Caption = 9 Or Label3.Caption = 11 Or Label3.Caption = 14 Or Label3.Caption = 20 Or Label3.Caption = 21 Or Label3.Caption = 25 Or Label3.Caption = 28 Or Label3.Caption = 30 Then
If Option2.Value = True Then
Text2.Text = Text2.Text + 1
Form7.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
Else
Form5.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
End If
End If

If Label3.Caption = 3 Or Label3.Caption = 4 Or Label3.Caption = 16 Or Label3.Caption = 17 Or Label3.Caption = 23 Then
If Option3.Value = True Then
Text2.Text = Text2.Text + 1
Form7.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
Else
Form5.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
End If
End If

If Label3.Caption = 8 Or Label3.Caption = 12 Or Label3.Caption = 15 Or Label3.Caption = 18 Or Label3.Caption = 19 Or Label3.Caption = 26 Or Label3.Caption = 27 Or Label3.Caption = 31 Or Label3.Caption = 32 Then
If Option4.Value = True Then
Text2.Text = Text2.Text + 1
Form7.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
Else
Form5.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
End If
End If
End If

'''team c

If Option7.Value = True Then


If Label3.Caption = 1 Or Label3.Caption = 7 Or Label3.Caption = 10 Or Label3.Caption = 13 Or Label3.Caption = 22 Or Label3.Caption = 24 Or Label3.Caption = 29 Then
If Option1.Value = True Then
Text4.Text = Text4.Text + 1
Form7.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
Else
Form5.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
End If
End If

If Label3.Caption = 2 Or Label3.Caption = 5 Or Label3.Caption = 9 Or Label3.Caption = 11 Or Label3.Caption = 14 Or Label3.Caption = 20 Or Label3.Caption = 21 Or Label3.Caption = 25 Or Label3.Caption = 28 Or Label3.Caption = 30 Then
If Option2.Value = True Then
Text4.Text = Text4.Text + 1
Form7.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
Else
Form5.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
End If
End If

If Label3.Caption = 3 Or Label3.Caption = 4 Or Label3.Caption = 6 Or Label3.Caption = 16 Or Label3.Caption = 17 Or Label3.Caption = 23 Then
If Option3.Value = True Then
Text4.Text = Text4.Text + 1
Form7.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
Else
Form5.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
End If
End If

If Label3.Caption = 8 Or Label3.Caption = 12 Or Label3.Caption = 15 Or Label3.Caption = 18 Or Label3.Caption = 19 Or Label3.Caption = 26 Or Label3.Caption = 27 Or Label3.Caption = 31 Or Label3.Caption = 32 Then
If Option4.Value = True Then
Text4.Text = Text4.Text + 1
Form7.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False

Else
Form5.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
End If
End If

End If

End Sub

Private Sub Form_Activate()
'''Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
Set db = OpenDatabase("C:\Users\Dell\Desktop\vb\rajeev vb\finalproject1.mdb")
Set rec = db.OpenRecordset("quiz1")
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False

If rec.EOF Then

Form4.Hide
Load Form4
Form4.Show

MsgBox "completed the survey"

End If
Label3.Caption = rec.Fields("slno")
Label2.Caption = rec.Fields("Question")
Label5.Caption = rec.Fields("option1")
Label6.Caption = rec.Fields("option2")
Label7.Caption = rec.Fields("option3")
Label8.Caption = rec.Fields("option4")
Text3.Text = rec.Fields("answer")
Timer1.Enabled = True
rec.MoveNext

End Sub


Private Sub Timer1_Timer()

Label4.Caption = Label4.Caption - 1
If Label4.Caption = 0 Then
Timer1.Enabled = False
Form11.Show
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
Label4.Caption = 60


End If

End Sub

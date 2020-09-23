VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   2280
      Top             =   2520
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1920
      Top             =   2520
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   2520
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   1200
      Top             =   2520
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   2520
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   2520
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "By"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "DEBRIS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Space "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
RememberScreenRes
ChangeScreenSettings 1024, 768, 16
Call sndPlaySound(App.Path & "\sounds\all_men_.wav", &H1)
End Sub

Private Sub Timer1_Timer()
Timer2.Enabled = False
If (Label1.FontSize < 30) Then
Label1.FontSize = Label1.FontSize + 1
Else
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
If (Label2.FontSize < 40) Then
Label2.FontSize = Label2.FontSize + 1
Else
Timer2.Enabled = False
Timer3.Enabled = True
Label3.Visible = True
Call sndPlaySound(App.Path & "\sounds\helicopt.wav", &H1)
End If
End Sub

Private Sub Timer3_Timer()
If (Label3.Left < 4000) Then
Label3.Left = Label3.Left + 150
Else
Label3.Visible = False
Timer3.Enabled = False
Form2.Show
Form3.Show
Timer4.Enabled = True
End If
End Sub

Private Sub Timer4_Timer()
Call sndPlaySound(App.Path & "\sounds\headcras.wav ", &H1)
Unload Form2
Timer4.Enabled = False
Timer5.Enabled = True
End Sub

Private Sub Timer5_Timer()
Call sndPlaySound(App.Path & "\sounds\headcras.wav ", &H1)
Unload Form3
Timer5.Enabled = False
Timer6.Enabled = True
End Sub

Private Sub Timer6_Timer()
Call sndPlaySound(App.Path & "\sounds\headcras.wav  ", &H1)
Timer6.Enabled = False
Timer7.Enabled = True
Form1.Visible = False 'To stop it from repeating again
End Sub

Private Sub Timer7_Timer()
Form5.Show
Timer7.Enabled = False
Unload Me
End Sub

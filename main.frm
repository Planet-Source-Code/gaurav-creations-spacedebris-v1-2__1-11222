VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Space Debris"
   ClientHeight    =   7830
   ClientLeft      =   1680
   ClientTop       =   1710
   ClientWidth     =   6780
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "main.frx":0442
   ScaleHeight     =   6830
   ScaleMode       =   0  'User
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   7320
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2880
      Top             =   7320
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   7320
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   7320
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   7320
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   7320
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   7815
      Left            =   3960
      ScaleHeight     =   7755
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton Command1 
         Enabled         =   0   'False
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   6600
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "www.question.chatbook.com"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   7440
         Width           =   2535
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Spc Bar"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   6840
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Right"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   6240
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Up"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   6240
         Width           =   495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Left"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   6240
         Width           =   495
      End
      Begin VB.Shape Shape2 
         Height          =   2415
         Left            =   120
         Top             =   4920
         Width           =   2535
      End
      Begin VB.Image Image11 
         Height          =   480
         Left            =   1800
         Picture         =   "main.frx":8267
         Top             =   5640
         Width           =   480
      End
      Begin VB.Image Image10 
         Height          =   480
         Left            =   1080
         Picture         =   "main.frx":86A9
         Top             =   5640
         Width           =   480
      End
      Begin VB.Image Image9 
         Height          =   480
         Left            =   360
         Picture         =   "main.frx":8AEB
         Top             =   5640
         Width           =   480
      End
      Begin VB.Shape Shape1 
         Height          =   2655
         Left            =   240
         Top             =   120
         Width           =   2175
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   3120
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3120
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rocket Controls"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   4920
         Width           =   2415
      End
      Begin VB.Image Image8 
         Height          =   480
         Index           =   4
         Left            =   2160
         Picture         =   "main.frx":8F2D
         Top             =   3960
         Width           =   480
      End
      Begin VB.Image Image8 
         Height          =   480
         Index           =   3
         Left            =   1680
         Picture         =   "main.frx":936F
         Top             =   3960
         Width           =   480
      End
      Begin VB.Image Image8 
         Height          =   480
         Index           =   2
         Left            =   1200
         Picture         =   "main.frx":97B1
         Top             =   3960
         Width           =   480
      End
      Begin VB.Image Image8 
         Height          =   480
         Index           =   1
         Left            =   720
         Picture         =   "main.frx":9BF3
         Top             =   3960
         Width           =   480
      End
      Begin VB.Image Image8 
         Height          =   480
         Index           =   0
         Left            =   240
         Picture         =   "main.frx":A035
         Top             =   3960
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Lives Left"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   4
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Level No."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -120
         TabIndex        =   2
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "SCORE"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   1080
      Top             =   7320
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   720
      Top             =   7320
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   7320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   7320
   End
   Begin VB.Image Image7 
      Height          =   1305
      Left            =   1680
      Picture         =   "main.frx":A477
      Stretch         =   -1  'True
      Top             =   -720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image6 
      Height          =   555
      Index           =   2
      Left            =   2400
      Picture         =   "main.frx":B010
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image6 
      Height          =   555
      Index           =   1
      Left            =   1320
      Picture         =   "main.frx":B31A
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   555
      Index           =   0
      Left            =   720
      Picture         =   "main.frx":B624
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   1170
      Left            =   480
      Picture         =   "main.frx":BA66
      Top             =   2880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   1560
      Picture         =   "main.frx":D1F0
      Stretch         =   -1  'True
      Top             =   6240
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image3 
      Height          =   345
      Index           =   2
      Left            =   1920
      Picture         =   "main.frx":D5AC
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   345
      Index           =   1
      Left            =   240
      Picture         =   "main.frx":E066
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image3 
      Height          =   345
      Index           =   0
      Left            =   3480
      Picture         =   "main.frx":E370
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image Image2 
      Height          =   1680
      Left            =   1680
      Picture         =   "main.frx":E7B2
      Top             =   6240
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   1680
      Left            =   1680
      Picture         =   "main.frx":128B4
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myvalue As Integer
Dim score As Integer
Dim counter As Integer
Dim sate As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'Keys
Select Case KeyCode

Case vbKeyUp
Call sndPlaySound(App.Path & "\sounds\thrust.wav", &H1)
Image2.Visible = False
Image1.Visible = True

If score < 500 Then
Randomize
myvalue = Int((3 * Rnd))
For x = 0 To myvalue Step 1
Image3(myvalue).Visible = True
Image3(myvalue).Top = Image3(myvalue).Top + 80
Next x
Timer1.Enabled = True
ElseIf score > 490 And score < 1000 Then
Timer1.Enabled = False
Timer5.Enabled = True
ElseIf score > 996 And score < 2490 Then
Timer1.Enabled = False
Timer5.Enabled = False
Timer7.Enabled = True
ElseIf score > 2490 Then
Timer7.Enabled = False
Timer9.Enabled = True
End If

Case vbKeyLeft
Image1.Left = Image1.Left - 60


Case vbKeyRight
Image1.Left = Image1.Left + 60


Case vbKeySpace
Call sndPlaySound(App.Path & "\sounds\missile.wav", &H1)
Image1.Visible = True
If score < 500 Then
Timer1.Enabled = True
Timer2.Enabled = True
ElseIf score > 490 And score < 1000 Then
Timer2.Enabled = False
Timer5.Enabled = True
Timer6.Enabled = True
ElseIf score > 996 And score < 2490 Then
Timer2.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = True
Timer8.Enabled = True
ElseIf score > 1995 Then
Timer8.Enabled = False
Timer9.Enabled = True
Timer10.Enabled = True
End If
Image4.Visible = True
End Select
' End of keys
End Sub

Private Sub Form_Load()
Form4.Picture = LoadPicture(App.Path & "\pics\sky.gif")
Image1.Picture = LoadPicture(App.Path & "\pics\rocket.gif")
Image2.Picture = LoadPicture(App.Path & "\pics\rocket1.gif")
Image3(0).Picture = LoadPicture(App.Path & "\pics\z.ico")
Image3(1).Picture = LoadPicture(App.Path & "\pics\disk06.ico")
Image3(2).Picture = LoadPicture(App.Path & "\pics\dan.ico")
Image6(0).Picture = LoadPicture(App.Path & "\pics\exe.ico")
Image6(1).Picture = LoadPicture(App.Path & "\pics\cdrom02.ico")
Image6(2).Picture = LoadPicture(App.Path & "\pics\unrzs.ico")
Image7.Picture = LoadPicture(App.Path & "\pics\sat.gif")
Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
Image8(0).Picture = LoadPicture(App.Path & "\pics\misc34.ico")
Image8(1).Picture = LoadPicture(App.Path & "\pics\misc34.ico")
Image8(2).Picture = LoadPicture(App.Path & "\pics\misc34.ico")
Image8(3).Picture = LoadPicture(App.Path & "\pics\misc34.ico")
Image8(4).Picture = LoadPicture(App.Path & "\pics\misc34.ico")
Image4.Picture = LoadPicture(App.Path & "\pics\weapon.gif")
Image9.Picture = LoadPicture(App.Path & "\pics\arw02lt.ico")
Image10.Picture = LoadPicture(App.Path & "\pics\arw02up.ico")
Image11.Picture = LoadPicture(App.Path & "\pics\arw02rt.ico")

Image1.Visible = False
Randomize
For x = 0 To myvalue Step 1
myvalue = Int((3 * Rnd))
h = Int((3600 * Rnd))
Image3(myvalue).Left = h
Next x
score = 0
counter = 0
sate = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)

   Dim Msg, Response   ' Declare variables.
   Msg = "Are you sure yo want to exit"
   Response = MsgBox(Msg, vbQuestion + vbYesNo, "Exit")
   Select Case Response
      Case vbYes
      Select Case iWidth 'put screen res back to normal
      Case 640
      ChangeScreenSettings 640, 480, 16
      Case 800
      ChangeScreenSettings 800, 600, 16
      Case 1024
      ChangeScreenSettings 1024, 768, 16
      End Select
      End
      Case vbNo
      Cancel = -1
    End Select
End Sub

Private Sub Label11_Click()
Shell "explorer http://www.geocities.com/gauravdhup"
End Sub

' Timer 1 for debris

Private Sub Timer1_Timer()
If counter = 1 Then
Image8(4).Visible = False
ElseIf counter = 2 Then
Image8(3).Visible = False
ElseIf counter = 3 Then
Image8(2).Visible = False
ElseIf counter = 4 Then
Image8(1).Visible = False
ElseIf counter = 5 Then
Image8(0).Visible = False
Form6.Show
End If

If score < 500 Then
myvalue = Int((3 * Rnd))

' Collision detection

Image3(myvalue).Visible = True
Image3(myvalue).Top = Image3(myvalue).Top + 70
    If Image3(myvalue).Left > Image1.Left And Image3(myvalue).Left < Image1.Left + Image1.Width Then
        If Image3(myvalue).Top > Image1.Top And Image3(myvalue).Top < Image1.Top + Image1.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer1.Enabled = False
        Timer2.Enabled = False
        Image5.Top = Image3(myvalue).Top
        Image5.Left = Image3(myvalue).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image3(myvalue).Visible = False
        Timer3.Enabled = True
        myvalue1 = Int((3600 * Rnd))
        Image3(myvalue).Top = 0
        Image3(myvalue).Left = myvalue1
        Image3(myvalue).Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        ElseIf Image3(myvalue).Top + Image3(myvalue).Height > Image1.Top And Image3(myvalue).Top + Image3(myvalue).Height < Image1.Top + Image1.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer1.Enabled = False
        Timer2.Enabled = False
        Image5.Top = Image3(myvalue).Top
        Image5.Left = Image3(myvalue).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image3(myvalue).Visible = False
        Timer3.Enabled = True
        myvalue1 = Int((3600 * Rnd))
        Image3(myvalue).Top = 0
        Image3(myvalue).Left = myvalue1
        Image3(myvalue).Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        End If
       
End If
    If Image3(myvalue).Left + Image3(myvalue).Width > Image1.Left And Image3(myvalue).Left + Image3(myvalue).Width < Image1.Left + Image1.Width Then
        If Image3(myvalue).Top + Image3(myvalue).Height > Image1.Top And Image3(myvalue).Top + Image3(myvalue).Height < Image1.Top + Image1.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer1.Enabled = False
        Timer2.Enabled = False
        Image5.Top = Image3(myvalue).Top
        Image5.Left = Image3(myvalue).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image3(myvalue).Visible = False
        Timer3.Enabled = True
        myvalue1 = Int((3600 * Rnd))
        Image3(myvalue).Top = 0
        Image3(myvalue).Left = myvalue1
        Image3(myvalue).Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        ElseIf Image3(myvalue).Top > Image1.Top And Image3(myvalue).Top < Image1.Top + Image1.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer1.Enabled = False
        Timer2.Enabled = False
        Image5.Top = Image3(myvalue).Top
        Image5.Left = Image3(myvalue).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image3(myvalue).Visible = False
        Timer3.Enabled = True
        myvalue1 = Int((3600 * Rnd))
        Image3(myvalue).Top = 0
        Image3(myvalue).Left = myvalue1
        Image3(myvalue).Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        End If
End If
' To check the end of screen for debris

If Image3(0).Top > 6600 Then
Image3(0).Top = 100
myvalue1 = Int((3600 * Rnd))
Image3(0).Left = myvalue1
End If

If Image3(1).Top > 6600 Then
Image3(1).Top = 100
myvalue1 = Int((3600 * Rnd))
Image3(1).Left = myvalue1
End If

If Image3(2).Top > 6600 Then
Image3(2).Top = 100
myvalue1 = Int((3600 * Rnd))
Image3(2).Left = myvalue1
End If

'End check the end of screen for debris
Else
Timer1.Enabled = False
Image3(0).Visible = False
Image3(1).Visible = False
Image3(2).Visible = False
Timer5.Enabled = True
Label4.Caption = 2
End If
 
End Sub

Private Sub Timer10_Timer()
If counter = 1 Then
Image8(4).Visible = False
ElseIf counter = 2 Then
Image8(3).Visible = False
ElseIf counter = 3 Then
Image8(2).Visible = False
ElseIf counter = 4 Then
Image8(1).Visible = False
ElseIf counter = 5 Then
Image8(0).Visible = False
Form6.Show
End If
If sate = 5 Then
Form7.Show
End If

Image4.Visible = True
Image4.Left = Image1.Left
Image4.Top = Image4.Top - 100
 
 If Image7.Left > Image4.Left And Image7.Left < Image4.Left + Image4.Width Then
        If Image7.Top > Image4.Top And Image7.Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image7.Top + Image7.Height > Image4.Top And Image7.Top + Image7.Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer10.Enabled = False
        Image5.Top = Image7.Top
        Image5.Left = Image7.Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image7.Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        'myvalue1 = Int((3600 * Rnd))
        'Image7.Top = 0
        'Image7.Left = myvalue1
        Image7.Visible = True
        score = score + 250
        Label2.Caption = score
        sate = sate + 1
        End If
   End If
    If Image7.Left + Image7.Width > Image4.Left And Image7.Left + Image7.Width < Image4.Left + Image4.Width Then
        If Image7.Top + Image7.Height > Image4.Top And Image7.Top + Image7.Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image7.Top > Image4.Top And Image7.Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer10.Enabled = False
        Image5.Top = Image7.Top
        Image5.Left = Image7.Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image7.Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        'myvalue1 = Int((3600 * Rnd))
        'Image7.Top = 0
        'Image7.Left = myvalue1
        Image7.Visible = True
        score = score + 250
        Label2.Caption = score
        sate = sate + 1
     End If
    End If
If Image4.Top < 0 Then
Image4.Visible = False
Image4.Top = Image1.Top
Timer10.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
If counter = 1 Then
Image8(4).Visible = False
ElseIf counter = 2 Then
Image8(3).Visible = False
ElseIf counter = 3 Then
Image8(2).Visible = False
ElseIf counter = 4 Then
Image8(1).Visible = False
ElseIf counter = 5 Then
Image8(0).Visible = False
Form6.Show
End If


Image4.Visible = True
Image4.Left = Image1.Left
Image4.Top = Image4.Top - 100
 
 If Image3(0).Left > Image4.Left And Image3(0).Left < Image4.Left + Image4.Width Then
        If Image3(0).Top > Image4.Top And Image3(0).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image3(0).Top + Image3(0).Height > Image4.Top And Image3(0).Top + Image3(0).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer2.Enabled = False
        Image5.Top = Image3(0).Top
        Image5.Left = Image3(0).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image3(0).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image3(0).Top = 0
        Image3(0).Left = myvalue1
        Image3(0).Visible = True
        score = score + 25
        Label2.Caption = score
        End If
   End If
    If Image3(0).Left + Image3(0).Width > Image4.Left And Image3(0).Left + Image3(0).Width < Image4.Left + Image4.Width Then
        If Image3(0).Top + Image3(0).Height > Image4.Top And Image3(0).Top + Image3(0).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image3(0).Top > Image4.Top And Image3(0).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer2.Enabled = False
        Image5.Top = Image3(0).Top
        Image5.Left = Image3(0).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image3(0).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image3(0).Top = 0
        Image3(0).Left = myvalue1
        Image3(0).Visible = True
        score = score + 25
        Label2.Caption = score
     End If
    End If
    
  If Image3(1).Left > Image4.Left And Image3(1).Left < Image4.Left + Image4.Width Then
       If Image3(1).Top > Image4.Top And Image3(1).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image3(1).Top + Image3(1).Height > Image4.Top And Image3(1).Top + Image3(1).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer2.Enabled = False
        Image5.Top = Image3(1).Top
        Image5.Left = Image3(1).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image3(1).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image3(1).Top = 0
        Image3(1).Left = myvalue1
        Image3(1).Visible = True
        score = score + 25
        Label2.Caption = score
        End If
      
End If
    If Image3(1).Left + Image3(1).Width > Image4.Left And Image3(1).Left + Image3(1).Width < Image4.Left + Image4.Width Then
        If Image3(1).Top + Image3(1).Height > Image4.Top And Image3(1).Top + Image3(1).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image3(1).Top > Image4.Top And Image3(1).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer2.Enabled = False
        Image5.Top = Image3(1).Top
        Image5.Left = Image3(1).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image3(1).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image3(1).Top = 0
        Image3(1).Left = myvalue1
        Image3(1).Visible = True
        score = score + 25
        Label2.Caption = score
        End If
   
    End If
   
   
If Image3(2).Left > Image4.Left And Image3(2).Left < Image4.Left + Image4.Width Then
       If Image3(2).Top > Image4.Top And Image3(2).Top < Image4.Top + Image4.Height Then
       Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
       ElseIf Image3(2).Top + Image3(2).Height > Image4.Top And Image3(2).Top + Image3(2).Height < Image4.Top + Image4.Width Then
       Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
       Timer2.Enabled = False
       Image5.Top = Image3(2).Top
       Image5.Left = Image3(2).Left
       Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
       Image5.Visible = True
       Image4.Visible = False
       Image3(2).Visible = False
       Timer3.Enabled = True
       Image4.Top = Image1.Top
       myvalue1 = Int((3600 * Rnd))
       Image3(2).Top = 0
       Image3(2).Left = myvalue1
       Image3(2).Visible = True
       score = score + 25
       Label2.Caption = score
       End If
     End If
    If Image3(2).Left + Image3(2).Width > Image4.Left And Image3(2).Left + Image3(2).Width < Image4.Left + Image4.Width Then
        If Image3(2).Top + Image3(2).Height > Image4.Top And Image3(2).Top + Image3(2).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image3(2).Top > Image4.Top And Image3(2).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer2.Enabled = False
        Image5.Top = Image3(2).Top
        Image5.Left = Image3(2).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image3(2).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image3(2).Top = 0
        Image3(2).Left = myvalue1
        Image3(2).Visible = True
        score = score + 25
        Label2.Caption = score
       End If
      End If

If Image4.Top < 0 Then
If counter = 1 Then
Image8(4).Visible = False
ElseIf counter = 2 Then
Image8(3).Visible = False
ElseIf counter = 3 Then
Image8(2).Visible = False
ElseIf counter = 4 Then
Image8(1).Visible = False
ElseIf counter = 5 Then
Image8(0).Visible = False
Form6.Show
'Unload Me
End If

Image4.Visible = False
Image4.Top = Image1.Top
Timer2.Enabled = False
End If
End Sub

Private Sub Timer3_Timer()
If counter = 1 Then
Image8(4).Visible = False
ElseIf counter = 2 Then
Image8(3).Visible = False
ElseIf counter = 3 Then
Image8(2).Visible = False
ElseIf counter = 4 Then
Image8(1).Visible = False
ElseIf counter = 5 Then
Image8(0).Visible = False
Form6.Show
End If

Image5.Visible = False
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
If counter = 1 Then
Image8(4).Visible = False
ElseIf counter = 2 Then
Image8(3).Visible = False
ElseIf counter = 3 Then
Image8(2).Visible = False
ElseIf counter = 4 Then
Image8(1).Visible = False
ElseIf counter = 5 Then
Image8(0).Visible = False
Form6.Show
End If

Call sndPlaySound(App.Path & "\sounds\beamin.wav", &H1)
For ope = 0 To 2845 Step 0.1
Form4.Width = 4035 + ope
Next ope
Timer4.Enabled = False
End Sub


' Timer 5 for debris 2
Private Sub Timer5_Timer()
If counter = 1 Then
Image8(4).Visible = False
ElseIf counter = 2 Then
Image8(3).Visible = False
ElseIf counter = 3 Then
Image8(2).Visible = False
ElseIf counter = 4 Then
Image8(1).Visible = False
ElseIf counter = 5 Then
Image8(0).Visible = False
Form6.Show
End If


If score > 490 And score < 995 Then
myvalue = Int((3 * Rnd))

' Collision detection

Image6(myvalue).Visible = True
Image6(myvalue).Top = Image6(myvalue).Top + 90
    If Image6(myvalue).Left > Image1.Left And Image6(myvalue).Left < Image1.Left + Image1.Width Then
        If Image6(myvalue).Top > Image1.Top And Image6(myvalue).Top < Image1.Top + Image1.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer5.Enabled = False
        Timer6.Enabled = False
        Image5.Top = Image6(myvalue).Top
        Image5.Left = Image6(myvalue).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image6(myvalue).Visible = False
        Timer3.Enabled = True
        myvalue1 = Int((3600 * Rnd))
        Image6(myvalue).Top = 0
        Image6(myvalue).Left = myvalue1
        Image6(myvalue).Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        ElseIf Image6(myvalue).Top + Image6(myvalue).Height > Image1.Top And Image6(myvalue).Top + Image6(myvalue).Height < Image1.Top + Image1.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer5.Enabled = False
        Timer6.Enabled = False
        Image5.Top = Image6(myvalue).Top
        Image5.Left = Image6(myvalue).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image6(myvalue).Visible = False
        Timer3.Enabled = True
        myvalue1 = Int((3600 * Rnd))
        Image6(myvalue).Top = 0
        Image6(myvalue).Left = myvalue1
        Image6(myvalue).Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        End If
       
End If
    If Image6(myvalue).Left + Image6(myvalue).Width > Image1.Left And Image6(myvalue).Left + Image6(myvalue).Width < Image1.Left + Image1.Width Then
        If Image6(myvalue).Top + Image6(myvalue).Height > Image1.Top And Image6(myvalue).Top + Image6(myvalue).Height < Image1.Top + Image1.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer5.Enabled = False
        Timer6.Enabled = False
        Image5.Top = Image6(myvalue).Top
        Image5.Left = Image6(myvalue).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image6(myvalue).Visible = False
        Timer3.Enabled = True
        myvalue1 = Int((3600 * Rnd))
        Image6(myvalue).Top = 0
        Image6(myvalue).Left = myvalue1
        Image6(myvalue).Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        ElseIf Image6(myvalue).Top > Image1.Top And Image6(myvalue).Top < Image1.Top + Image1.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer5.Enabled = False
        Timer6.Enabled = False
        Image5.Top = Image6(myvalue).Top
        Image5.Left = Image6(myvalue).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image6(myvalue).Visible = False
        Timer3.Enabled = True
        myvalue1 = Int((3600 * Rnd))
        Image6(myvalue).Top = 0
        Image6(myvalue).Left = myvalue1
        Image6(myvalue).Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        End If
End If
' To check the end of screen for debris

If Image6(0).Top > 6600 Then
Image6(0).Top = 100
myvalue1 = Int((3600 * Rnd))
Image6(0).Left = myvalue1
End If

If Image6(1).Top > 6600 Then
Image6(1).Top = 100
myvalue1 = Int((3600 * Rnd))
Image6(1).Left = myvalue1
End If

If Image6(2).Top > 6600 Then
Image6(2).Top = 100
myvalue1 = Int((3600 * Rnd))
Image6(2).Left = myvalue1
End If
Else
Timer5.Enabled = False
Timer7.Enabled = True
'End check the end of screen for debris
Label4.Caption = 3
End If
End Sub

'For Weapon 2

Private Sub Timer6_Timer()
If counter = 1 Then
Image8(4).Visible = False
ElseIf counter = 2 Then
Image8(3).Visible = False
ElseIf counter = 3 Then
Image8(2).Visible = False
ElseIf counter = 4 Then
Image8(1).Visible = False
ElseIf counter = 5 Then
Image8(0).Visible = False
Form6.Show
End If


Image4.Visible = True
Image4.Left = Image1.Left
Image4.Top = Image4.Top - 100
 
 If Image6(0).Left > Image4.Left And Image6(0).Left < Image4.Left + Image4.Width Then
        If Image6(0).Top > Image4.Top And Image6(0).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image6(0).Top + Image6(0).Height > Image4.Top And Image6(0).Top + Image6(0).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer6.Enabled = False
        Image5.Top = Image6(0).Top
        Image5.Left = Image6(0).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image6(0).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image6(0).Top = 0
        Image6(0).Left = myvalue1
        Image6(0).Visible = True
        score = score + 50
        Label2.Caption = score
        End If
   End If
    If Image6(0).Left + Image6(0).Width > Image4.Left And Image6(0).Left + Image6(0).Width < Image4.Left + Image4.Width Then
        If Image6(0).Top + Image6(0).Height > Image4.Top And Image6(0).Top + Image6(0).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image6(0).Top > Image4.Top And Image6(0).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer6.Enabled = False
        Image5.Top = Image6(0).Top
        Image5.Left = Image6(0).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image6(0).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image6(0).Top = 0
        Image6(0).Left = myvalue1
        Image6(0).Visible = True
        score = score + 50
        Label2.Caption = score
     End If
    End If
    
  If Image6(1).Left > Image4.Left And Image6(1).Left < Image4.Left + Image4.Width Then
       If Image6(1).Top > Image4.Top And Image6(1).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image6(1).Top + Image6(1).Height > Image4.Top And Image6(1).Top + Image6(1).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer6.Enabled = False
        Image5.Top = Image6(1).Top
        Image5.Left = Image6(1).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image6(1).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image6(1).Top = 0
        Image6(1).Left = myvalue1
        Image6(1).Visible = True
        score = score + 50
        Label2.Caption = score
        End If
      
End If
    If Image6(1).Left + Image6(1).Width > Image4.Left And Image6(1).Left + Image6(1).Width < Image4.Left + Image4.Width Then
        If Image6(1).Top + Image6(1).Height > Image4.Top And Image6(1).Top + Image6(1).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image6(1).Top > Image4.Top And Image6(1).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer6.Enabled = False
        Image5.Top = Image3(1).Top
        Image5.Left = Image3(1).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image6(1).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image6(1).Top = 0
        Image6(1).Left = myvalue1
        Image6(1).Visible = True
        score = score + 50
        Label2.Caption = score
        End If
   
    End If
   
   
If Image6(2).Left > Image4.Left And Image6(2).Left < Image4.Left + Image4.Width Then
       If Image3(2).Top > Image4.Top And Image3(2).Top < Image4.Top + Image4.Height Then
       Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
       ElseIf Image6(2).Top + Image6(2).Height > Image4.Top And Image6(2).Top + Image6(2).Height < Image4.Top + Image4.Width Then
       Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
       Timer6.Enabled = False
       Image5.Top = Image6(2).Top
       Image5.Left = Image6(2).Left
       Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
       Image5.Visible = True
       Image4.Visible = False
       Image6(2).Visible = False
       Timer3.Enabled = True
       Image4.Top = Image1.Top
       myvalue1 = Int((3600 * Rnd))
       Image6(2).Top = 0
       Image6(2).Left = myvalue1
       Image6(2).Visible = True
       score = score + 50
       Label2.Caption = score
       End If
     End If
    If Image6(2).Left + Image6(2).Width > Image4.Left And Image6(2).Left + Image6(2).Width < Image4.Left + Image4.Width Then
        If Image6(2).Top + Image6(2).Height > Image4.Top And Image6(2).Top + Image6(2).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image6(2).Top > Image4.Top And Image6(2).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer6.Enabled = False
        Image5.Top = Image6(2).Top
        Image5.Left = Image6(2).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image6(2).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image6(2).Top = 0
        Image6(2).Left = myvalue1
        Image6(2).Visible = True
        score = score + 50
        Label2.Caption = score
       End If
      End If

If Image4.Top < 0 Then
Image4.Visible = False
Image4.Top = Image1.Top
Timer6.Enabled = False
End If

End Sub


' Timer for level 3 debris
Private Sub Timer7_Timer()
If counter = 1 Then
Image8(4).Visible = False
ElseIf counter = 2 Then
Image8(3).Visible = False
ElseIf counter = 3 Then
Image8(2).Visible = False
ElseIf counter = 4 Then
Image8(1).Visible = False
ElseIf counter = 5 Then
Image8(0).Visible = False
Form6.Show
End If

If score > 996 And score < 1995 Then
myvalue = Int((3 * Rnd))

' Collision detection

Image6(myvalue).Visible = True
Image6(myvalue).Top = Image6(myvalue).Top + 90
    If Image6(myvalue).Left > Image1.Left And Image6(myvalue).Left < Image1.Left + Image1.Width Then
        If Image6(myvalue).Top > Image1.Top And Image6(myvalue).Top < Image1.Top + Image1.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer7.Enabled = False
        Timer8.Enabled = False
        Image5.Top = Image6(myvalue).Top
        Image5.Left = Image6(myvalue).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image6(myvalue).Visible = False
        Timer3.Enabled = True
        myvalue1 = Int((3600 * Rnd))
        Image6(myvalue).Top = 0
        Image6(myvalue).Left = myvalue1
        Image6(myvalue).Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        ElseIf Image6(myvalue).Top + Image6(myvalue).Height > Image1.Top And Image6(myvalue).Top + Image6(myvalue).Height < Image1.Top + Image1.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer7.Enabled = False
        Timer8.Enabled = False
        Image5.Top = Image6(myvalue).Top
        Image5.Left = Image6(myvalue).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image6(myvalue).Visible = False
        Timer3.Enabled = True
        myvalue1 = Int((3600 * Rnd))
        Image6(myvalue).Top = 0
        Image6(myvalue).Left = myvalue1
        Image6(myvalue).Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        End If
       
End If
    If Image6(myvalue).Left + Image6(myvalue).Width > Image1.Left And Image6(myvalue).Left + Image6(myvalue).Width < Image1.Left + Image1.Width Then
        If Image6(myvalue).Top + Image6(myvalue).Height > Image1.Top And Image6(myvalue).Top + Image6(myvalue).Height < Image1.Top + Image1.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer7.Enabled = False
        Timer8.Enabled = False
        Image5.Top = Image6(myvalue).Top
        Image5.Left = Image6(myvalue).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image6(myvalue).Visible = False
        Timer3.Enabled = True
        myvalue1 = Int((3600 * Rnd))
        Image6(myvalue).Top = 0
        Image6(myvalue).Left = myvalue1
        Image6(myvalue).Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        ElseIf Image6(myvalue).Top > Image1.Top And Image6(myvalue).Top < Image1.Top + Image1.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer7.Enabled = False
        Timer8.Enabled = False
        Image5.Top = Image6(myvalue).Top
        Image5.Left = Image6(myvalue).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image6(myvalue).Visible = False
        Timer3.Enabled = True
        myvalue1 = Int((3600 * Rnd))
        Image6(myvalue).Top = 0
        Image6(myvalue).Left = myvalue1
        Image6(myvalue).Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        End If
End If
' To check the end of screen for debris

If Image6(0).Top > 6600 Then
Image6(0).Top = 100
myvalue1 = Int((3600 * Rnd))
Image6(0).Left = myvalue1
End If

If Image6(1).Top > 6600 Then
Image6(1).Top = 100
myvalue1 = Int((3600 * Rnd))
Image6(1).Left = myvalue1
End If

If Image6(2).Top > 6600 Then
Image6(2).Top = 100
myvalue1 = Int((3600 * Rnd))
Image6(2).Left = myvalue1
End If

'End check the end of screen for debris
myvalue = Int((3 * Rnd))

' Collision detection

Image3(myvalue).Visible = True
Image3(myvalue).Top = Image3(myvalue).Top + 80
    If Image3(myvalue).Left > Image1.Left And Image3(myvalue).Left < Image1.Left + Image1.Width Then
        If Image3(myvalue).Top > Image1.Top And Image3(myvalue).Top < Image1.Top + Image1.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer7.Enabled = False
        Timer8.Enabled = False
        Image5.Top = Image3(myvalue).Top
        Image5.Left = Image3(myvalue).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image3(myvalue).Visible = False
        Timer3.Enabled = True
        myvalue1 = Int((3600 * Rnd))
        Image3(myvalue).Top = 0
        Image3(myvalue).Left = myvalue1
        Image3(myvalue).Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        ElseIf Image3(myvalue).Top + Image3(myvalue).Height > Image1.Top And Image3(myvalue).Top + Image3(myvalue).Height < Image1.Top + Image1.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer7.Enabled = False
        Timer8.Enabled = False
        Image5.Top = Image3(myvalue).Top
        Image5.Left = Image3(myvalue).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image3(myvalue).Visible = False
        Timer3.Enabled = True
        myvalue1 = Int((3600 * Rnd))
        Image3(myvalue).Top = 0
        Image3(myvalue).Left = myvalue1
        Image3(myvalue).Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        End If
       
End If
    If Image3(myvalue).Left + Image3(myvalue).Width > Image1.Left And Image3(myvalue).Left + Image3(myvalue).Width < Image1.Left + Image1.Width Then
        If Image3(myvalue).Top + Image3(myvalue).Height > Image1.Top And Image3(myvalue).Top + Image3(myvalue).Height < Image1.Top + Image1.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer7.Enabled = False
        Timer8.Enabled = False
        Image5.Top = Image3(myvalue).Top
        Image5.Left = Image3(myvalue).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image3(myvalue).Visible = False
        Timer3.Enabled = True
        myvalue1 = Int((3600 * Rnd))
        Image3(myvalue).Top = 0
        Image3(myvalue).Left = myvalue1
        Image3(myvalue).Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        ElseIf Image3(myvalue).Top > Image1.Top And Image3(myvalue).Top < Image1.Top + Image1.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer7.Enabled = False
        Timer8.Enabled = False
        Image5.Top = Image3(myvalue).Top
        Image5.Left = Image3(myvalue).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image3(myvalue).Visible = False
        Timer3.Enabled = True
        myvalue1 = Int((3600 * Rnd))
        Image3(myvalue).Top = 0
        Image3(myvalue).Left = myvalue1
        Image3(myvalue).Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        End If
End If
' To check the end of screen for debris

If Image3(0).Top > 6600 Then
Image3(0).Top = 100
myvalue1 = Int((3600 * Rnd))
Image3(0).Left = myvalue1
End If

If Image3(1).Top > 6600 Then
Image3(1).Top = 100
myvalue1 = Int((3600 * Rnd))
Image3(1).Left = myvalue1
End If

If Image3(2).Top > 6600 Then
Image3(2).Top = 100
myvalue1 = Int((3600 * Rnd))
Image3(2).Left = myvalue1
End If

'End check the end of screen for debris
Else
Image3(0).Visible = False
Image3(1).Visible = False
Image3(2).Visible = False
Image6(0).Visible = False
Image6(1).Visible = False
Image6(2).Visible = False
Image3(0).Enabled = False
Image3(1).Enabled = False
Image3(2).Enabled = False
Image6(0).Enabled = False
Image6(1).Enabled = False
Image6(2).Enabled = False
Timer7.Enabled = False
Label4.Caption = 4
Timer9.Enabled = True
End If
End Sub

Private Sub Timer8_Timer()
If counter = 1 Then
Image8(4).Visible = False
ElseIf counter = 2 Then
Image8(3).Visible = False
ElseIf counter = 3 Then
Image8(2).Visible = False
ElseIf counter = 4 Then
Image8(1).Visible = False
ElseIf counter = 5 Then
Image8(0).Visible = False
Form6.Show
End If


Image4.Visible = True
Image4.Left = Image1.Left
Image4.Top = Image4.Top - 100
 
 If Image3(0).Left > Image4.Left And Image3(0).Left < Image4.Left + Image4.Width Then
        If Image3(0).Top > Image4.Top And Image3(0).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image3(0).Top + Image3(0).Height > Image4.Top And Image3(0).Top + Image3(0).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer8.Enabled = False
        Image5.Top = Image3(0).Top
        Image5.Left = Image3(0).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image3(0).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image3(0).Top = 0
        Image3(0).Left = myvalue1
        Image3(0).Visible = True
        score = score + 25
        Label2.Caption = score
        End If
   End If
    If Image3(0).Left + Image3(0).Width > Image4.Left And Image3(0).Left + Image3(0).Width < Image4.Left + Image4.Width Then
        If Image3(0).Top + Image3(0).Height > Image4.Top And Image3(0).Top + Image3(0).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image3(0).Top > Image4.Top And Image3(0).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer8.Enabled = False
        Image5.Top = Image3(0).Top
        Image5.Left = Image3(0).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image3(0).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image3(0).Top = 0
        Image3(0).Left = myvalue1
        Image3(0).Visible = True
        score = score + 25
        Label2.Caption = score
     End If
    End If
    
  If Image3(1).Left > Image4.Left And Image3(1).Left < Image4.Left + Image4.Width Then
       If Image3(1).Top > Image4.Top And Image3(1).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image3(1).Top + Image3(1).Height > Image4.Top And Image3(1).Top + Image3(1).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer8.Enabled = False
        Image5.Top = Image3(1).Top
        Image5.Left = Image3(1).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image3(1).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image3(1).Top = 0
        Image3(1).Left = myvalue1
        Image3(1).Visible = True
        score = score + 25
        Label2.Caption = score
        End If
      
End If
    If Image3(1).Left + Image3(1).Width > Image4.Left And Image3(1).Left + Image3(1).Width < Image4.Left + Image4.Width Then
        If Image3(1).Top + Image3(1).Height > Image4.Top And Image3(1).Top + Image3(1).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image3(1).Top > Image4.Top And Image3(1).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer8.Enabled = False
        Image5.Top = Image3(1).Top
        Image5.Left = Image3(1).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image3(1).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image3(1).Top = 0
        Image3(1).Left = myvalue1
        Image3(1).Visible = True
        score = score + 25
        Label2.Caption = score
        End If
   
    End If
   
   
If Image3(2).Left > Image4.Left And Image3(2).Left < Image4.Left + Image4.Width Then
       If Image3(2).Top > Image4.Top And Image3(2).Top < Image4.Top + Image4.Height Then
       Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
       ElseIf Image3(2).Top + Image3(2).Height > Image4.Top And Image3(2).Top + Image3(2).Height < Image4.Top + Image4.Width Then
       Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
       Timer8.Enabled = False
       Image5.Top = Image3(2).Top
       Image5.Left = Image3(2).Left
       Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
       Image5.Visible = True
       Image4.Visible = False
       Image3(2).Visible = False
       Timer3.Enabled = True
       Image4.Top = Image1.Top
       myvalue1 = Int((3600 * Rnd))
       Image3(2).Top = 0
       Image3(2).Left = myvalue1
       Image3(2).Visible = True
       score = score + 25
       Label2.Caption = score
       End If
     End If
    If Image3(2).Left + Image3(2).Width > Image4.Left And Image3(2).Left + Image3(2).Width < Image4.Left + Image4.Width Then
        If Image3(2).Top + Image3(2).Height > Image4.Top And Image3(2).Top + Image3(2).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image3(2).Top > Image4.Top And Image3(2).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer8.Enabled = False
        Image5.Top = Image3(2).Top
        Image5.Left = Image3(2).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image3(2).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image3(2).Top = 0
        Image3(2).Left = myvalue1
        Image3(2).Visible = True
        score = score + 25
        Label2.Caption = score
       End If
      End If

 If Image6(0).Left > Image4.Left And Image6(0).Left < Image4.Left + Image4.Width Then
        If Image6(0).Top > Image4.Top And Image6(0).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image6(0).Top + Image6(0).Height > Image4.Top And Image6(0).Top + Image6(0).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer8.Enabled = False
        Image5.Top = Image6(0).Top
        Image5.Left = Image6(0).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image6(0).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image6(0).Top = 0
        Image6(0).Left = myvalue1
        Image6(0).Visible = True
        score = score + 50
        Label2.Caption = score
        End If
   End If
    If Image6(0).Left + Image6(0).Width > Image4.Left And Image6(0).Left + Image6(0).Width < Image4.Left + Image4.Width Then
        If Image6(0).Top + Image6(0).Height > Image4.Top And Image6(0).Top + Image6(0).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image6(0).Top > Image4.Top And Image6(0).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer8.Enabled = False
        Image5.Top = Image6(0).Top
        Image5.Left = Image6(0).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image6(0).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image6(0).Top = 0
        Image6(0).Left = myvalue1
        Image6(0).Visible = True
        score = score + 50
        Label2.Caption = score
     End If
    End If
    
  If Image6(1).Left > Image4.Left And Image6(1).Left < Image4.Left + Image4.Width Then
       If Image6(1).Top > Image4.Top And Image6(1).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image6(1).Top + Image6(1).Height > Image4.Top And Image6(1).Top + Image6(1).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer8.Enabled = False
        Image5.Top = Image6(1).Top
        Image5.Left = Image6(1).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image6(1).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image6(1).Top = 0
        Image6(1).Left = myvalue1
        Image6(1).Visible = True
        score = score + 50
        Label2.Caption = score
        End If
      
End If
    If Image6(1).Left + Image6(1).Width > Image4.Left And Image6(1).Left + Image6(1).Width < Image4.Left + Image4.Width Then
        If Image6(1).Top + Image6(1).Height > Image4.Top And Image6(1).Top + Image6(1).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image6(1).Top > Image4.Top And Image6(1).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer8.Enabled = False
        Image5.Top = Image3(1).Top
        Image5.Left = Image3(1).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image6(1).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image6(1).Top = 0
        Image6(1).Left = myvalue1
        Image6(1).Visible = True
        score = score + 50
        Label2.Caption = score
        End If
   
    End If
   
   
If Image6(2).Left > Image4.Left And Image6(2).Left < Image4.Left + Image4.Width Then
       If Image3(2).Top > Image4.Top And Image3(2).Top < Image4.Top + Image4.Height Then
       Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
       ElseIf Image6(2).Top + Image6(2).Height > Image4.Top And Image6(2).Top + Image6(2).Height < Image4.Top + Image4.Width Then
       Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
       Timer8.Enabled = False
       Image5.Top = Image6(2).Top
       Image5.Left = Image6(2).Left
       Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
       Image5.Visible = True
       Image4.Visible = False
       Image6(2).Visible = False
       Timer3.Enabled = True
       Image4.Top = Image1.Top
       myvalue1 = Int((3600 * Rnd))
       Image6(2).Top = 0
       Image6(2).Left = myvalue1
       Image6(2).Visible = True
       score = score + 50
       Label2.Caption = score
       End If
     End If
    If Image6(2).Left + Image6(2).Width > Image4.Left And Image6(2).Left + Image6(2).Width < Image4.Left + Image4.Width Then
        If Image6(2).Top + Image6(2).Height > Image4.Top And Image6(2).Top + Image6(2).Height < Image4.Top + Image4.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        ElseIf Image6(2).Top > Image4.Top And Image6(2).Top < Image4.Top + Image4.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer8.Enabled = False
        Image5.Top = Image6(2).Top
        Image5.Left = Image6(2).Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image4.Visible = False
        Image6(2).Visible = False
        Timer3.Enabled = True
        Image4.Top = Image1.Top
        myvalue1 = Int((3600 * Rnd))
        Image6(2).Top = 0
        Image6(2).Left = myvalue1
        Image6(2).Visible = True
        score = score + 50
        Label2.Caption = score
       End If
      End If

If Image4.Top < 0 Then
Image4.Visible = False
Image4.Top = Image1.Top
Image3(0).Visible = False
Image3(1).Visible = False
Image3(2).Visible = False
Image6(0).Visible = False
Image6(1).Visible = False
Image6(2).Visible = False
Image3(0).Enabled = False
Image3(1).Enabled = False
Image3(2).Enabled = False
Image6(0).Enabled = False
Image6(1).Enabled = False
Image6(2).Enabled = False
Timer8.Enabled = False
Timer10.Enabled = True
End If
End Sub

Private Sub Timer9_Timer()
If counter = 1 Then
Image8(4).Visible = False
ElseIf counter = 2 Then
Image8(3).Visible = False
ElseIf counter = 3 Then
Image8(2).Visible = False
ElseIf counter = 4 Then
Image8(1).Visible = False
ElseIf counter = 5 Then
Image8(0).Visible = False
Form6.Show
End If

' Collision detection
Image7.Visible = True
Image7.Top = Image7.Top + 30
    If Image7.Left > Image1.Left And Image7.Left < Image1.Left + Image1.Width Then
        If Image7.Top > Image1.Top And Image7.Top < Image1.Top + Image1.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer5.Enabled = False
        Timer6.Enabled = False
        Image5.Top = Image7.Top
        Image5.Left = Image7.Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image7.Visible = False
        Timer3.Enabled = True
        Image7.Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        ElseIf Image7.Top + Image7.Height > Image1.Top And Image7.Top + Image7.Height < Image1.Top + Image1.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer5.Enabled = False
        Timer6.Enabled = False
        Image5.Top = Image7.Top
        Image5.Left = Image7.Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image7.Visible = False
        Timer3.Enabled = True
        Image7.Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
     End If
    End If
    If Image7.Left + Image7.Width > Image1.Left And Image7.Left + Image7.Width < Image1.Left + Image1.Width Then
        If Image7.Top + Image7.Height > Image1.Top And Image7.Top + Image7.Height < Image1.Top + Image1.Width Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer5.Enabled = False
        Timer6.Enabled = False
        Image5.Top = Image7.Top
        Image5.Left = Image7.Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image6(myvalue).Visible = False
        Timer3.Enabled = True
        Image7.Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        ElseIf Image7.Top > Image1.Top And Image7.Top < Image1.Top + Image1.Height Then
        Call sndPlaySound(App.Path & "\sounds\headcras.wav", &H1)
        Timer5.Enabled = False
        Timer6.Enabled = False
        Image5.Top = Image7.Top
        Image5.Left = Image7.Left
        Image5.Picture = LoadPicture(App.Path & "\pics\boom.gif")
        Image5.Visible = True
        Image1.Visible = False
        Image7.Visible = False
        Timer3.Enabled = True
        Image7.Visible = True
        Image4.Visible = False
        Image4.Top = Image1.Top
        counter = counter + 1 ' lives
        End If
End If
' To check the end of screen for debris

If Image7.Top > 6600 Then
Image7.Top = 100
myvalue1 = Int((3600 * Rnd))
Image7.Left = myvalue1
End If

'End check the end of screen for debris

End Sub

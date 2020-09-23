VERSION 5.00
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "VTEXT.DLL"
Begin VB.Form Form8 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form8"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9840
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form8"
   ScaleHeight     =   4695
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin HTTSLibCtl.TextToSpeech TextToSpeech1 
      Height          =   2895
      Left            =   6480
      OleObjectBlob   =   "speak.frx":0000
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6120
      Top             =   120
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Height          =   1740
      Left            =   7200
      Picture         =   "speak.frx":0024
      ScaleHeight     =   1680
      ScaleWidth      =   2250
      TabIndex        =   1
      Top             =   2520
      Width           =   2310
   End
   Begin VB.PictureBox Picture1 
      Height          =   2055
      Left            =   6720
      Picture         =   "speak.frx":0CA4
      ScaleHeight     =   1995
      ScaleWidth      =   2955
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BEGIN"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2895
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "speak.frx":33AE
      Top             =   1080
      Width           =   6255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "www.question.chatbook.com"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   -480
      TabIndex        =   5
      Top             =   4080
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "DATE    : 24th August 2000 "
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "MISSION : To Clear Up The Space Debris"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Form4.Show
End Sub

Private Sub Form_Load()
TextToSpeech1.Select (4)
TextToSpeech1.Speak Label1
TextToSpeech1.Speak Label2
TextToSpeech1.Speak Text1
Picture1.Picture = LoadPicture(App.Path & "\pics\globe.gif")
Picture2.Picture = LoadPicture(App.Path & "\pics\object8.gif")
End Sub

Private Sub Timer1_Timer()
Picture1.Picture = LoadPicture(App.Path & "\pics\globe.gif")
Picture1.Visible = False
Picture1.Picture = LoadPicture(App.Path & "\pics\globe1.gif")
Picture1.Visible = True
End Sub

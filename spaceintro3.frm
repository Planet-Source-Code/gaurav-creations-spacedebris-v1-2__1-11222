VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   3135
   ClientLeft      =   4320
   ClientTop       =   4200
   ClientWidth     =   1185
   LinkTopic       =   "Form3"
   ScaleHeight     =   3150
   ScaleMode       =   0  'User
   ScaleWidth      =   1185
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "By"
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
      Left            =   480
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "G a u r a v"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "C R E A T I O N S"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form3.Top = Form1.Top + 50
Form3.Left = Form1.Left - Form3.Width
End Sub

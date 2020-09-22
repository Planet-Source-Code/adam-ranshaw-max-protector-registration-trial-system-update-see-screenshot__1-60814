VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invalid Registration Code"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "Form9.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4800
      Top             =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form9.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invalid Registration Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
End
End Sub



Private Sub ok_Click()
Form9.Visible = False
reg.Enabled = True
reg.Timer3.Enabled = True
End Sub

Private Sub Timer1_Timer()
Label2.Caption = "The Registration Code you provided was not found on my server or you have entered someone elses Registration Code.  Please ensure you are entering your Registration Code correctly.  You have entered a total of " + reg.labelpuk.Caption + " incorrect Registration Codes.  After 5 incorrect Registration Codes this software may become frozzen or disabled."
End Sub

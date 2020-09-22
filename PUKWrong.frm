VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Invalid PUK Code"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5145
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   1
      Top             =   2400
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "PUKWrong.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invalid PUK Code"
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
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"PUKWrong.frx":0442
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
      TabIndex        =   3
      Top             =   600
      Width           =   4335
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Label2.Caption = "The PUK Code you provided is not valid for your computer.  Please try again by clicking 'OK' or if you do not have a PUK Code for your computer then please send an email to aranshaw@aol.com with your Computer ID: " + Form10.Text1.Text + ".  By clicking 'Exit' this program will close."
End Sub

Private Sub ok_Click()
Form8.Visible = False
Form10.Enabled = True
End Sub

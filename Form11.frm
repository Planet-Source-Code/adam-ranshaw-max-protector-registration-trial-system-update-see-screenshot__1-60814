VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Live Connection Required"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   -120
      Picture         =   "Form11.frx":0000
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Re-Try"
      Default         =   -1  'True
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form11.frx":80A2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Live Internet Connection Required"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If CheckConnection = True Then
form3.Enabled = True
Unload Form1
Form11.Visible = False
form3.Label1.Visible = True
form3.unreg.Visible = False
form3.Visible = True
Else
MsgBox "No connection has been detected, if you have just connected wait 20 secounds and try again.", vbCritical
End If
End Sub

Private Sub Command2_Click()
form3.Visible = False
Form1.Visible = True
Form1.Enabled = True
form3.Enabled = True
Form11.Visible = False
End Sub

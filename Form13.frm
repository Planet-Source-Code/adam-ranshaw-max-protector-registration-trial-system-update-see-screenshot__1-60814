VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Live Connection Required"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Re-Try"
      Default         =   -1  'True
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form13.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Live Internet Connection Required"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
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

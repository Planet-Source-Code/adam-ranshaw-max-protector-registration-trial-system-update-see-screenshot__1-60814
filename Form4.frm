VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administration OK"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   7335
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "It appears you are using an Adranix computer system.  Please enter the admin password in the box below:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6975
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADRANIX"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
Unload Form4
End Sub



Private Sub ok_Click()
If Text1.Text = "@overcar33" Then
Text1.Text = ""
Form4.Visible = False
Form7.Visible = True
Else
MsgBox "The password is not correct.", vbCritical
End If
End Sub

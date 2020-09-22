VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H001463FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NOT OK!"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   6630
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton continue 
      Caption         =   "Continue"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form5.frx":0000
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
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6375
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
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub continue_Click()
Form5.Visible = False
Form6.Visible = True
End Sub



Private Sub ok_Click()
Form5.Visible = False
End Sub

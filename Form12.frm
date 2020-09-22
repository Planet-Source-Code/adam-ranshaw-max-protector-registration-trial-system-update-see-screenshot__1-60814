VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ERROR"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   $"Form12.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "An Unexpected Error has Occured!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "ERROR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form11.Visible = False
End Sub

Private Sub Command2_Click()
End
End Sub

VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00F48A2E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirm Administrator ID"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   6210
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton exit 
      Caption         =   "Back"
      Height          =   495
      Left            =   3480
      TabIndex        =   10
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton validate 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00F48A2E&
      Caption         =   "I understand that I am an ADRANIX administrator and that it is illigel to use this feature otherwise"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Value           =   1  'Checked
      Width           =   5775
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   7
      Text            =   "28/12/1989"
      Top             =   2400
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   5
      Text            =   "@overcar33"
      Top             =   1680
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "Adam Dale Kevin Ranshaw"
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "D.O.B:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Overide Key:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Admin ID:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Administration"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "ADRANIX"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub exit_Click()
Form6.Hide
End Sub



Private Sub validate_Click()
If Text1.Text = "Adam Dale Kevin Ranshaw" And Text2.Text = "@overcar33" And Text3.Text = "28/12/1989" And Check1.Value = 1 Then
Unload Form6
Form6.Visible = False
Form7.Visible = True
Else
MsgBox "All or some infomation is not valid and has been rejected!", vbCritical
End If
End Sub

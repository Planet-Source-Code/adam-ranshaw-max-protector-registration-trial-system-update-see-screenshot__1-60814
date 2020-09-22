VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administration"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   3510
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton chameleonButton4 
      Caption         =   "Close Options"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   3135
   End
   Begin VB.CommandButton chameleonButton3 
      Caption         =   "Register Software"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   3135
   End
   Begin VB.CommandButton chameleonButton2 
      Caption         =   "Set trial to X listed below"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   3135
   End
   Begin VB.CommandButton chameleonButton1 
      Caption         =   "Expire users trial"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   3135
   End
   Begin VB.CommandButton reset 
      Caption         =   "Put trial to 0 (Zero)"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      TabIndex        =   0
      Text            =   "0"
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Administrative Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
SaveSetting Me.Name, "Trial", "TimesOpen", 3720
Form1.Label1.Caption = 3720
MsgBox "Command was a succsess.", vbInformation
End Sub

Private Sub chameleonButton2_Click()
SaveSetting Me.Name, "Trial", "TimesOpen", Text1.Text
Form1.Label1.Caption = Text1.Text
MsgBox "Command was a succsess.", vbInformation
End Sub

Private Sub chameleonButton3_Click()
reg.name1.Text = "Administrator"
reg.code.Text = "283031293565"
reg.sec.Text = "111000"
reg.sec.SaveFile "c:\windows\system32\adranixsec000.rtf"
reg.name1.SaveFile "c:\windows\system32\adranixname000.rtf"
reg.code.SaveFile "c:\windows\system32\adranixcode000.rtf"
MsgBox "Command was a succsess.", vbInformation
End Sub

Private Sub chameleonButton4_Click()
Unload Form7
End Sub





Private Sub reset_Click()
SaveSetting Me.Name, "Trial", "TimesOpen", 60
Form1.Label1.Caption = 60
MsgBox "Command was a succsess.", vbInformation
End Sub

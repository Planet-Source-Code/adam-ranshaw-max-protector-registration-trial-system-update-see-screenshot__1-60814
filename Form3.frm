VERSION 5.00
Begin VB.Form form3 
   BackColor       =   &H00FF80FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2130
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   5700
   ControlBox      =   0   'False
   DrawWidth       =   80
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   2130
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Software Here"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
   Begin VB.Menu regmenu 
      Caption         =   "Registration"
      Begin VB.Menu regnow 
         Caption         =   "Register"
      End
      Begin VB.Menu unreg 
         Caption         =   "Unregister"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub cmdChoice1_Click(Index As Integer)
frmAcey.Show
End Sub

Private Sub cmdChoice2_Click(Index As Integer)
frmHam.Show
End Sub

Private Sub cmdChoice3_Click(Index As Integer)
frmEven.Show
End Sub

Private Sub cmdChoice4_Click(Index As Integer)
frmMemory.Show
End Sub

Private Sub cmdChoice5_Click(Index As Integer)
frmMug.Show
End Sub

Private Sub cmdChoice6_Click(Index As Integer)
frmJot.Show
End Sub

Private Sub cmdChoice7_Click(Index As Integer)
frmLunar.Show
End Sub

Private Sub cmdChoice8_Click(Index As Integer)
frmBandit.Show
End Sub

Private Sub exit_Click()
End
End Sub












Private Sub regnow_Click()
Unload reg
reg.Visible = False
reg.Visible = True
End Sub

Private Sub unreg_Click()
reg.name1.Text = ""
reg.code.Text = ""
reg.sec.Text = ""
Kill "c:\windows\system32\adranixsec000.rtf"
Kill "c:\windows\system32\adranixname000.rtf"
Kill "c:\windows\system32\adranixcode000.rtf"
Kill "c:\windows\system32\maxuser.reg"
SaveSetting Me.Name, "Trial", "TimesOpen", 0
Form1.Label1.Caption = 0
End
End Sub


Private Sub XpBs3_Click()
Unload reg
reg.Visible = False
reg.Visible = True
End Sub

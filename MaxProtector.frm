VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trial Expired"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   ControlBox      =   0   'False
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton back 
      Caption         =   "&Back"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4800
      TabIndex        =   13
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2520
      MaxLength       =   9
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Text            =   "64381"
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Text            =   "8546854"
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton continue 
      Caption         =   "Next"
      Default         =   -1  'True
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.OptionButton Option4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Enter PUK Code"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   2160
      Width           =   3255
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Close the software"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   2520
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Go to Main Start-UP Screen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   1800
      Width           =   3015
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Go to Registration page"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   1440
      Value           =   -1  'True
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      Picture         =   "MaxProtector.frx":0000
      ScaleHeight     =   3015
      ScaleWidth      =   2175
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton ok 
      Caption         =   "&Next"
      Height          =   495
      Left            =   6000
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "I have no PUK - Get PUK Code from Adranix"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Your 30 minute trial has Expired.  What whould you like to do now?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your trial is now Expired"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long



Private Sub back_Click()
Label2.Caption = "A PUK is needed, please exit this software or enter the PUK."
continue.Visible = True
ok.Visible = False
Option1.Visible = True
Option2.Visible = True
Option3.Visible = True
Option4.Visible = True
Text2.Visible = False
back.Enabled = False
End Sub

Private Sub continue_Click()
If Option1.Value = True Then
reg.labelpuk.Caption = GetSetting(Me.Name, "PUK", "TimesOpen")
Dim PCName As String
Dim P As Long
P = NameOfPC(PCName)
reg.Text1.Text = PCName
Form1.Visible = True
reg.Visible = True
End If
If Option2.Value = True Then
Form10.Visible = False
Form1.Visible = True
End If
If Option3.Value = True Then
Form10.Visible = False
End
End If
If Option4.Value = True Then
Label2.Caption = "Please enter your PUK code in the Box Below."
continue.Visible = False
ok.Visible = True
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
Text2.Visible = True
back.Enabled = True
End If
End Sub

Private Sub Label3_Click()
MsgBox "Please send an e-mail to aranshaw@aol.com along with your computer name: " + Text1.Text + " for your PUK code", vbExclamation
End Sub

Private Sub Timer1_Timer()
Dim PCName As String
Dim P As Long
P = NameOfPC(PCName)
Text1.Text = PCName
End Sub
Public Function NameOfPC(MachineName As String) As Long
    Dim NameSize As Long
    Dim x As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
End Function






Private Sub ok_Click()
'Registration  Code format
Dim i
Dim zip
Dim final
Dim code1 As Single
If Text1.Text = "" Or Text2.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
MsgBox "Please enter a PUK code before clicking Continueing.", vbExclamation
Text2.Text = ""
Exit Sub
End If



If Text5.Text = ("8546854") And Text6.Text = "64381" Then


Else
MsgBox "Invalid PUK Code was entered. You can try again as many times as you like.", vbCritical
Text2.Text = ""
Exit Sub
End If


For i = 1 To Len(Text1.Text) - 1
    code1 = Format(Asc(Right(Text1.Text, Len(Text1.Text) - i)) * 2 + (39 / i) + (i + 3 / 7), "#.#")
    zip = zip & code1
Next i
zip = Right(zip, 8)

For i = 1 To Len(zip) - 1
    code1 = Format(Asc(Right(zip, Len(zip) - i)) * 0.1 + (1 / i) + (i + 1 / 7), "#00")
    final = final & code1
Next i
final = Right(final, Len(final) - 4)
final = final & Asc(Text1)
'If reg code is correct
If Text2.Text = final Then
reg.Timer4.Enabled = True
MsgBox "You have Unblocked this software.  Please write down your PUK as you will need it again if you get locked out.", vbInformation
End
Else
MsgBox "Invalid PUK Code was entered. You can try again as many times as you like.", vbCritical
Text2.Text = ""
End If
End Sub

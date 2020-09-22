VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00F48A2E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Access Frozen"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   5805
   ControlBox      =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton getpuk 
      Caption         =   "Get PUK Code and Close"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton ok 
      Caption         =   "Validate PUK and Unlock"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Text            =   "8546854"
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Text            =   "64381"
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
      Left            =   1800
      MaxLength       =   9
      TabIndex        =   4
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PUK Code:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Form8.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Access is Frozen"
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
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long







Private Sub Form_Load()
Dim PCName As String
Dim P As Long
P = NameOfPC(PCName)
Text1.Text = PCName
End Sub



Private Sub getpuk_Click()
MsgBox "Please send an e-mail to aranshaw@aol.com along with your computer name: " + Text1.Text + " for your PUK code", vbExclamation
End
End Sub





Private Sub ok_Click()
'Registration  Code format
Dim i
Dim zip
Dim final
Dim code1 As Single
If Text1.Text = "" Or Text2.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
MsgBox "Please enter a PUK code beofre clicking the NExt button.", vbExclamation
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


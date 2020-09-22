VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alart Box "
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   ControlBox      =   0   'False
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   350
      Left            =   0
      Top             =   1080
   End
   Begin VB.OptionButton Option5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Buy this Software with PayPal"
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
      TabIndex        =   18
      Top             =   2160
      Width           =   3255
   End
   Begin MaxProtector.chameleonButton back 
      Height          =   495
      Left            =   4800
      TabIndex        =   17
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Back"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Warning.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MaxProtector.chameleonButton continue 
      Height          =   495
      Left            =   6000
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Next"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Warning.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   10
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
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Text            =   "64381"
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Text            =   "8546854"
      Top             =   360
      Visible         =   0   'False
      Width           =   855
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
      Top             =   2520
      Width           =   1935
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
      Top             =   2880
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
      Caption         =   "Go to Registration Screen"
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
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   120
      Picture         =   "Warning.frx":0038
      ScaleHeight     =   3015
      ScaleWidth      =   2175
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin MaxProtector.chameleonButton ok 
      Height          =   495
      Left            =   6000
      TabIndex        =   16
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Next"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Warning.frx":151FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label perlab 
      Caption         =   "0"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   1440
      Width           =   2535
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
      TabIndex        =   11
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Your 60 minute trial has Expired.  What whould you like to do now?"
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
      Height          =   495
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
Private Function TrialTime(TheForm As Form, TrialOverMSG As String, TrialOverMSGTitle As String, TrialOverMSGType As String, TrialCount As Integer, Work As Boolean)

    If Not Work Then SaveSetting TheForm.Name, "per", "TimesOpen", "."
'If Work = False then reset trial to 0 if Work = True then Count up the Trial

    SaveSetting TheForm.Name, "per", "TimesOpen", Val(GetSetting(TheForm.Name, "per", "TimesOpen")) + 1
'Write + 1 to the last to the last time opened

    If GetSetting(TheForm.Name, "per", "TimesOpen") > TrialCount Then SaveSetting TheForm.Name, "per", "TimesOpen", TrialCount: MsgBox TrialOverMSG, TrialOverMSGType, TrialOverMSGTitle: Timer1.Enabled = False
'If the amount of times open is > then the TrialCount..
'Reset it to the number in TrialCount specified
'Display a message and terminate the program


    If Not Work Then SaveSetting TheForm.Name, "try", "TimesOpen", "."
'If Work = False then reset trial to 0 if Work = True then Count up the Trial

    SaveSetting TheForm.Name, "try", "TimesOpen", Val(GetSetting(TheForm.Name, "try", "TimesOpen")) + 1
'Write + 1 to the last to the last time opened

    If GetSetting(TheForm.Name, "try", "TimesOpen") > TrialCount Then SaveSetting TheForm.Name, "try", "TimesOpen", TrialCount: MsgBox TrialOverMSG, TrialOverMSGType, TrialOverMSGTitle: Timer1.Enabled = False
'If the amount of times open is > then the TrialCount..
'Reset it to the number in TrialCount specified
'Display a message and terminate the program
End Function



Private Sub back_Click()
Label2.Caption = "A PUK is needed, please exit this software or enter the PUK."
continue.Visible = True
ok.Visible = False
Option1.Visible = True
Option2.Visible = True
Option3.Visible = True
Option4.Visible = True
Option5.Visible = True
Text2.Visible = False
back.Enabled = False
End Sub

Private Sub continue_Click()
If Option1.Value = True Then
reg.labelpuk.Caption = GetSetting(Me.Name, "PUK", "TimesOpen")
Disk = "C:\"
Text2.Text = VolumeSerialNumber(Disk)
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
Label2.Caption = "Please enter your PUK code in the Box Below the click 'Next'"
continue.Visible = False
ok.Visible = True
Option1.Visible = False
Option2.Visible = False
Option3.Visible = False
Option4.Visible = False
Option5.Visible = False
Text2.Visible = True
back.Enabled = True
End If
If Option5.Value = True Then
Disk = "C:\"
OpenWebsite ("https://www.paypal.com/xclick/business=aranshaw@aol.com&undefined_quantity=1&item_name=MaxProtector " & VolumeSerialNumber(Disk) & "&amount=$8.99")
End If
End Sub







Private Sub Form_Load()
If App.PrevInstance Then
ActivatePrevInstance
End If
Disk = "C:\"
reg.Text1.Text = VolumeSerialNumber(Disk)
perlab.Caption = GetSetting(Me.Name, "per", "TimesOpen")
End Sub

Private Sub Label3_Click()
MsgBox "Please send an e-mail to aranshaw@aol.com along with your ID: " + Text1.Text + " for your PUK code.", vbExclamation
End Sub




Private Sub Timer1_Timer()
Disk = "C:\"
Text1.Text = VolumeSerialNumber(Disk)
End Sub







Private Sub ok_Click()
'Registration  Code format
Dim i
Dim zip
Dim final
Dim code1 As Single
If Text1.Text = "" Or Text2.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
MsgBox "Please enter a PUK code before clicking the Next button.", vbExclamation
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
    code1 = Format(Asc(Right(zip, Len(zip) - i)) * 0.8 + (1 / i) + (i + 1 / 7), "#00")
    final = final & code1
Next i
final = Right(final, Len(final) - 0)
final = final & Asc(Text1)
'If PUK code is correct
If Text2.Text = final Then
reg.Timer4.Enabled = True
Form10.Visible = False
Form1.Visible = True
Form2.Visible = True
Form1.exit1.Visible = False
Form1.Command1.Visible = True
Unload Form9
SaveSetting Me.Name, "per", "TimesOpen", 1
perlab.Caption = 1
Else
Form10.Enabled = False
Form8.Visible = True
Text2.Text = ""
End If
End Sub

Private Sub Timer2_Timer()
continue.Enabled = True
End Sub

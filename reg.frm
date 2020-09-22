VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RichTx32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form reg 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registration of Cash Streak"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   10695
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet pidconnecter 
      Left            =   5400
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer7 
      Interval        =   1000
      Left            =   5520
      Top             =   0
   End
   Begin InetCtlsObjects.Inet regserver 
      Left            =   3240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer6 
      Interval        =   10
      Left            =   2760
      Top             =   0
   End
   Begin VB.Timer Timer5 
      Interval        =   10
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin RichTextLib.RichTextBox sec 
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"reg.frx":0000
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Text            =   "64381"
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Text            =   "8546854"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4095
      Begin MaxProtector.chameleonButton pay 
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   4800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Click here to buy with PayPal - $8.99"
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
         MICON           =   "reg.frx":0082
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"reg.frx":009E
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3975
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "How to Buy Cash Streak:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2895
      End
   End
   Begin RichTextLib.RichTextBox name1 
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"reg.frx":026E
   End
   Begin RichTextLib.RichTextBox code 
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"reg.frx":02F0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000000FF&
      ForeColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   4440
      TabIndex        =   4
      Top             =   1320
      Width           =   6135
      Begin MaxProtector.chameleonButton website 
         Height          =   615
         Left            =   120
         TabIndex        =   24
         Top             =   4800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Visit Website"
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
         MICON           =   "reg.frx":0372
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MaxProtector.chameleonButton webval 
         Default         =   -1  'True
         Height          =   615
         Left            =   3480
         TabIndex        =   20
         Top             =   4080
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Stage 1:  Validate over Internet"
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
         MICON           =   "reg.frx":038E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MaxProtector.chameleonButton back 
         Height          =   615
         Left            =   2040
         TabIndex        =   17
         Top             =   4800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Back"
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
         MICON           =   "reg.frx":03AA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MaxProtector.chameleonButton ok 
         Height          =   615
         Left            =   3480
         TabIndex        =   16
         Top             =   4800
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "Stage 2:  Confirm Registration"
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
         MICON           =   "reg.frx":03C6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Text            =   "Checking Internet Connection..."
         Top             =   3360
         Width           =   5775
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Text            =   "ADRANIX-12345678"
         Top             =   2280
         Width           =   5775
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"reg.frx":03E2
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Checking Internet Connection                                                Please Wait..."
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
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   4080
         Width           =   5895
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Registration Code:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Your Computer ID:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Already Paid?  Heres how to Register:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Label productid 
      Caption         =   "ERROR-"
      Height          =   255
      Left            =   4320
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   3960
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label labelpuk 
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   6000
      TabIndex        =   8
      Top             =   360
      Width           =   4575
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   1095
      Left            =   6000
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADRANIX"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Private Function TrialTime(TheForm As Form, TrialOverMSG As String, TrialOverMSGTitle As String, TrialOverMSGType As String, TrialCount As Integer, Work As Boolean)

    If Not Work Then SaveSetting TheForm.Name, "puk", "TimesOpen", "."
'If Work = False then reset trial to 0 if Work = True then Count up the Trial

    SaveSetting TheForm.Name, "puk", "TimesOpen", Val(GetSetting(TheForm.Name, "puk", "TimesOpen")) + 1
'Write + 1 to the last to the last time opened

    If GetSetting(TheForm.Name, "puk", "TimesOpen") > TrialCount Then SaveSetting TheForm.Name, "puk", "TimesOpen", TrialCount: MsgBox TrialOverMSG, TrialOverMSGType, TrialOverMSGTitle: Timer1.Enabled = False
'If the amount of times open is > then the TrialCount..
'Reset it to the number in TrialCount specified
'Display a message and terminate the program
End Function


Private Sub admin_Click()
If Text1.Text = "ADAMS-PC" Then
Form4.Visible = False
Form4.Visible = True
Else
Form5.Visible = False
Form5.Visible = True
End If
End Sub



Private Sub back_Click()
Unload reg
End Sub






Public Function NameOfPC(MachineName As String) As Long
    Dim NameSize As Long
    Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
End Function








Private Sub Form_Load()
productid.Caption = pidconnecter.OpenURL("www.adranix.co.uk/Trial/MaxProtector/MaxProtector.txt")
reg.labelpuk.Caption = GetSetting(Me.Name, "PUK", "TimesOpen")
Disk = "C:\"
Text1.Text = VolumeSerialNumber(Disk)
If Len(reg.Text1.Text) < 4 Then
Check1.Visible = True
 End If
 End Sub



Private Sub ok_Click()
'Registration  Code format
Dim i
Dim zip
Dim final
Dim code1 As Single
If Text1.Text = "" Or Text2.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
    reg.Enabled = False
    Form9.Visible = True
Text2.Text = ""
Exit Sub
End If


If Len(Text1.Text) < 4 Then
   MsgBox "Please change your computer name to somthing over 3 letter/numbers.", vbExclamation
Text2.Text = ""
    Exit Sub
End If

If Text5.Text = ("8546854") And Text6.Text = "64381" Then


Else
    reg.Enabled = False
    Form9.Visible = True
Text2.Text = ""
Exit Sub
End If


For i = 1 To Len(Text1.Text) - 1
    code1 = Format(Asc(Right(Text1.Text, Len(Text1.Text) - i)) * 2 + (39 / i) + (i + 3 / 7), "#.#")
    zip = zip & code1
Next i
zip = Right(zip, 8)

For i = 1 To Len(zip) - 1
    code1 = Format(Asc(Right(zip, Len(zip) - i)) * 0.5 + (1 / i) + (i + 1 / 7), "#00")
    final = final & code1
Next i
final = final & Asc(Text1)
'If reg code is correct
If Text2.Text = final Then
'Enable License file Frame
name1.Text = Text1.Text
code.Text = Text2.Text
sec.Text = "111000"
sec.SaveFile "c:\windows\system32\adranixsec000.rtf"
name1.SaveFile "c:\windows\system32\adranixname000.rtf"
code.SaveFile "c:\windows\system32\adranixcode000.rtf"
MsgBox "Registration of this software is compleate.  Please keep a note of your Registration Code in the event of it being needed in the future use.  Click OK to terminate."
End
Else
TrialTime Me, "A PUK Code is needed to continue.", "PUK Code Needed", vbCritical, 5, True
labelpuk.Caption = GetSetting(Me.Name, "PUK", "TimesOpen")
    reg.Enabled = False
    Form9.Visible = True
    Timer3.Enabled = False
Text2.Text = ""
End If

End Sub



Private Sub pay_Click()
Disk = "C:\"
OpenWebsite ("https://www.paypal.com/xclick/business=aranshaw@aol.com&undefined_quantity=1&item_name=MaxProtector " & VolumeSerialNumber(Disk) & "&amount=$8.99")
End Sub


Private Sub Timer1_Timer()
If Label5.Caption = "E X P I R E D Left" Then
Timer1.Enabled = False
Label5.Caption = "E X P I R E D"
Else
Label5.Caption = Form1.Label3.Caption + " Left"
End If
End Sub

Private Sub Timer2_Timer()
Text1.Text = PCName
End Sub

Private Sub Timer3_Timer()
If labelpuk.Caption = "5" And Form10.perlab.Caption = "" Then
Unload Form10
Form10.Visible = True
Form1.Command1.Visible = False
Form10.Option1.Enabled = False
Form10.Option2.Enabled = False
Form10.Option4.Value = True
Form10.Option4.Enabled = True
Form10.Option5.Enabled = False
Form10.Label1.Caption = "Software is Frozen"
Form10.Label2.Caption = "A PUK code is needed, please exit this software or enter the PUK."
Form1.Visible = False
reg.Visible = False
Form1.Timer5.Enabled = False
Timer3.Enabled = False
End If
If labelpuk.Caption = "5" And Form10.perlab.Caption = "1" Then
Unload Form10
Form10.Visible = True
Form1.Command1.Visible = False
Form10.Option1.Enabled = False
Form10.Option2.Enabled = False
Form10.Option4.Enabled = False
Form10.Option5.Enabled = False
Form10.Option3.Value = True
Form10.Label1.Caption = "Permanently Disabled"
Form10.Label2.Caption = "This software is Permanently Disabled due to 10 incorrect Registration Codes."
Form1.Visible = False
reg.Visible = False
Form1.Timer5.Enabled = False
Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
SaveSetting Me.Name, "puk", "TimesOpen", 0
reg.labelpuk.Caption = 0
End Sub

Private Sub Timer5_Timer()
If labelpuk.Caption = "5" Then
Form1.Visible = False
reg.Visible = False
End If
End Sub



Private Sub Timer7_Timer()
 If CheckConnection = True Then
 Label7.Caption = "Internet Connection Active                                              Please Proceed."
 webval.Enabled = True
 Text2.Text = ""
 Text2.Enabled = True
 Else
 Label7.Caption = "Internet Connection Required!                                        Please Connect."
 webval.Enabled = False
 ok.Enabled = False
 Text2.Text = "No Internet Connection."
 Text2.Enabled = False
 End If
 Timer7.Enabled = False
End Sub

Private Sub website_Click()
OpenWebsite ("www.adranix.co.uk")
End Sub

Private Sub webval_Click()
On Error Resume Next
If Text2.Text = "" Then
MsgBox "Please enter a Registration Code and Try Again.", vbCritical
Else
webval.Enabled = False
back.Enabled = False
Label6 = regserver.OpenURL("www.adranix.co.uk/Trial/MaxProtector/Codes/" + Text2.Text + ".txt")
If Label6.Caption = Text2.Text Then
webval.Visible = False
Text1.Enabled = False
Label7.Caption = "Your Registration Code appears to be valid.  Please click the 'Comfirm Registration' button."
ok.Enabled = True
Text2.Enabled = False
back.Enabled = False
webval.Enabled = True
Else
TrialTime Me, "A PUK Code is needed to continue.", "PUK Code Needed", vbCritical, 5, True
labelpuk.Caption = GetSetting(Me.Name, "PUK", "TimesOpen")
    reg.Enabled = False
    Form9.Visible = True
End If
End If
End Sub

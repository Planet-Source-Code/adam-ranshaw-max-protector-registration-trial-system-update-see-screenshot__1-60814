VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RichTx32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cash Strek v1- MaxProtector"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   10695
   ControlBox      =   0   'False
   Enabled         =   0   'False
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin MaxProtector.chameleonButton chameleonButton1 
      Height          =   615
      Left            =   6840
      TabIndex        =   40
      Top             =   5040
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Click Here to Buy Now - PayPal"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MICON           =   "Form1.frx":0442
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
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   480
      Top             =   480
   End
   Begin VB.Timer Timer5 
      Interval        =   10
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   960
      Top             =   480
   End
   Begin VB.Timer Timer7 
      Interval        =   1
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1440
      Top             =   480
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   0
   End
   Begin VB.Timer Timer10 
      Interval        =   1
      Left            =   1920
      Top             =   480
   End
   Begin VB.Timer Timer11 
      Interval        =   10
      Left            =   2400
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      Begin VB.CommandButton Command3 
         Caption         =   "About"
         Height          =   375
         Left            =   3360
         TabIndex        =   14
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Msn Alert"
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   1560
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Msn Type"
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Msn Email"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Original"
         Height          =   255
         Left            =   2280
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Msn Online"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Caption         =   "&Show alert"
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "Sorry the was an Error - Please ingnore this message!"
         Top             =   360
         Width           =   4215
      End
      Begin VB.Timer Timer14 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   960
         Top             =   0
      End
      Begin VB.Label Label24 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer13 
      Interval        =   10
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer15 
      Interval        =   10
      Left            =   2400
      Top             =   480
   End
   Begin VB.Timer Timer16 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2880
      Top             =   0
   End
   Begin VB.Timer Timer17 
      Interval        =   1000
      Left            =   2880
      Top             =   480
   End
   Begin VB.Timer Timer18 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   0
   End
   Begin VB.Timer Timer19 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   480
   End
   Begin VB.Timer Timer20 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   0
   End
   Begin VB.Timer Timer22 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   480
   End
   Begin MaxProtector.chameleonButton command2 
      Height          =   855
      Left            =   6600
      TabIndex        =   0
      Top             =   5880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Register"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
      MICON           =   "Form1.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MaxProtector.chameleonButton command1 
      Default         =   -1  'True
      Height          =   855
      Left            =   8520
      TabIndex        =   1
      Top             =   5880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Try Game"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
      MICON           =   "Form1.frx":047A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RichTextLib.RichTextBox tbox 
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0496
   End
   Begin RichTextLib.RichTextBox test 
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":050D
   End
   Begin RichTextLib.RichTextBox currentuser 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      TextRTF         =   $"Form1.frx":0584
   End
   Begin RichTextLib.RichTextBox username 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      TextRTF         =   $"Form1.frx":0600
   End
   Begin MaxProtector.chameleonButton exit1 
      Height          =   855
      Left            =   8520
      TabIndex        =   16
      Top             =   5880
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Exit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
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
      MICON           =   "Form1.frx":0678
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   4
      Height          =   1695
      Left            =   2280
      Top             =   3000
      Width           =   6375
   End
   Begin VB.Label Label11 
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
      Left            =   5640
      TabIndex        =   18
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2520
      TabIndex        =   19
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait while loading Remaining time..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   20
      Top             =   3960
      Width           =   6015
   End
   Begin VB.Label Label19 
      BackColor       =   &H00009EEA&
      Height          =   1695
      Left            =   2280
      TabIndex        =   21
      Top             =   3000
      Width           =   6375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Why Not Buy Now?  Only $8.99"
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
      Height          =   255
      Left            =   6840
      TabIndex        =   31
      Top             =   4680
      Width           =   3135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00F48A2E&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   3615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "* Play any where any time"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6720
      TabIndex        =   39
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "* Instant Game Activation"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6600
      TabIndex        =   38
      Top             =   3840
      Width           =   3495
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "* No CD requied to play"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   6720
      TabIndex        =   37
      Top             =   3360
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   6600
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X Minutes"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   6600
      TabIndex        =   33
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "of gameplay remaining"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6600
      TabIndex        =   32
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8520
      TabIndex        =   30
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Own the full version for unlimited gameplay."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   6600
      TabIndex        =   29
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   28
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   27
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   26
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0694
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
      Height          =   1455
      Left            =   840
      TabIndex        =   25
      Top             =   5040
      Width           =   4935
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Streak"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   44.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   0
      TabIndex        =   23
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click here for more infomation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7080
      TabIndex        =   22
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Have"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6600
      TabIndex        =   17
      Top             =   360
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0014EFFB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   6735
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0749
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
      Height          =   1455
      Left            =   840
      TabIndex        =   35
      Top             =   3360
      Width           =   4935
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":07FD
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
      Height          =   1815
      Left            =   840
      TabIndex        =   36
      Top             =   1320
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_MAXIMIZE = 3
Private Function TrialTime(TheForm As Form, TrialOverMSG As String, TrialOverMSGTitle As String, TrialOverMSGType As String, TrialCount As Integer, Work As Boolean)

    If Not Work Then SaveSetting TheForm.Name, "Trial", "TimesOpen", "."
'If Work = False then reset trial to 0 if Work = True then Count up the Trial

    SaveSetting TheForm.Name, "Trial", "TimesOpen", Val(GetSetting(TheForm.Name, "Trial", "TimesOpen")) + 1
'Write + 1 to the last to the last time opened


End Function




Function FileExists(FileNa As String) As Boolean
    Dim FRes As String
    On Error GoTo NotFound
    FRes = Dir$(FileNa)
    If FRes = "" Then FileExists = False Else FileExists = True
NotFound:
    If Err = 53 Then Resume Next
End Function



Private Sub about_Click()
MsgBox "MaxProtector was made by Adam Ranshaw, it was made on 03/08/05 for www.adranix.co.uk - email: aranshaw@aol.com"
End Sub

Private Sub buynow_Click()
MsgBox "Feature Disabled.", vbCritical
End Sub

Private Sub chameleonButton1_Click()
Disk = "C:\"
OpenWebsite ("https://www.paypal.com/xclick/business=aranshaw@aol.com&undefined_quantity=1&item_name=MaxProtector " & VolumeSerialNumber(Disk) & "&amount=$8.99")
End Sub

Private Sub Command1_Click()
Unload Form1
form3.Label1.Visible = True
form3.unreg.Visible = False
form3.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Unload reg
reg.Visible = True
End Sub




Private Sub exit_Click()
End
End Sub

Private Sub exit1_Click()
End
End Sub



Private Sub Form_Load()
On Error Resume Next
If Label1.Caption = "" Then
Form1.Visible = True
Form10.Visible = False
Form10.Label1.Caption = "First Run"
End If
Randomize
Label1.Caption = GetSetting(Me.Name, "Trial", "TimesOpen")
reg.sec.LoadFile "c:\windows\system32\adranixsec000.rtf"
reg.code.LoadFile "c:\windows\system32\adranixcode000.rtf"
reg.Text2.Text = reg.code.Text
If reg.sec.Text = "111000" Then
Command1.Visible = False
'Registration  Code format
Dim i
Dim zip
Dim final
Dim code1 As Single
If reg.Text1.Text = "" Or reg.Text2.Text = "" Or reg.Text5.Text = "" Or reg.Text6.Text = "" Then
Form10.Visible = True
Form1.Visible = False
Form10.Label1.Caption = "Registration Fixed"
Form10.Label2.Caption = "There was an error with one or more registration files, please re-register."
Kill "c:\windows\system32\adranixsec000.rtf"
Kill "c:\windows\system32\adranixname000.rtf"
Kill "c:\windows\system32\adranixcode000.rtf"
Exit Sub
End If


If Len(reg.Text1.Text) < 4 Then
Form10.Visible = True
Form1.Visible = False
Form10.Label1.Caption = "Registration Fixed"
Form10.Label2.Caption = "There was an error with one or more registration files, please re-register."
Kill "c:\windows\system32\adranixsec000.rtf"
Kill "c:\windows\system32\adranixname000.rtf"
Kill "c:\windows\system32\adranixcode000.rtf"
    Exit Sub
End If

If reg.Text5.Text = ("8546854") And reg.Text6.Text = "64381" Then


Else
Form10.Visible = True
Form1.Visible = False
Form10.Label1.Caption = "Registration Fixed"
Form10.Label2.Caption = "There was an error with one or more registration files, please re-register."
Kill "c:\windows\system32\adranixsec000.rtf"
Kill "c:\windows\system32\adranixname000.rtf"
Kill "c:\windows\system32\adranixcode000.rtf"
Exit Sub
End If


For i = 1 To Len(reg.Text1.Text) - 1
    code1 = Format(Asc(Right(reg.Text1.Text, Len(reg.Text1.Text) - i)) * 2 + (39 / i) + (i + 3 / 7), "#.#")
    zip = zip & code1
Next i
zip = Right(zip, 8)

For i = 1 To Len(zip) - 1
    code1 = Format(Asc(Right(zip, Len(zip) - i)) * 0.5 + (1 / i) + (i + 1 / 7), "#00")
    final = final & code1
Next i
final = final & Asc(reg.Text1)
'If reg code is correct
If reg.Text2.Text = final Then
'Enable License file Frame
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
reg.name1.Text = reg.Text1.Text
reg.code.Text = reg.Text2.Text
form3.regnow.Visible = False
Form1.Visible = False
form3.Visible = True
Else
Form10.Visible = True
Form1.Visible = False
Form10.Label1.Caption = "Registration Fixed"
Form10.Label2.Caption = "There was an error with one or more registration files, please re-register."
Kill "c:\windows\system32\adranixsec000.rtf"
Kill "c:\windows\system32\adranixname000.rtf"
Kill "c:\windows\system32\adranixcode000.rtf"
End If
End If
If App.PrevInstance Then
MsgBox "Software Already Active:  For security reasons you can not run Multiple Instances of this program.", vbInformation
ActivatePrevInstance
End If
On Local Error GoTo ErrorHandling
tbox.SaveFile "c:\windows\system32\lontel.adr"
Kill "c:\windows\system32\lontel.adr"
'Continue
Exit Sub
ErrorHandling:
MsgBox "It appears the you do not have Read/Write access to the C Drive, please logon as an administrator and then try to open this software again.  Click OK to exit.", vbCritical
End
Resume Next
End Sub



Private Sub Label1_Click()
    SaveSetting Me.Name, "Trial", "TimesOpen", 0
'Resets the trial
    Label1.Caption = 0
'Resets the Label
End Sub



Private Sub Label10_Click()
If Label23.Visible = False Then
SaveSetting Me.Name, "Trial", "TimesOpen", 1799
Label1.Caption = 1799
End
End If
End Sub



Private Sub Label15_Click()
End
End Sub



Private Sub Label22_Click()
If Label23.Visible = False Then
Kill "c:\windows\system32\maxuser.reg"
End
End If
End Sub

Private Sub Label23_Click()
Label23.Visible = False
End Sub




Private Sub Label8_Click()
If Label23.Visible = False Then
SaveSetting Me.Name, "Trial", "TimesOpen", 3660
Label1.Caption = 3660
End
End If
End Sub

Private Sub Label9_Click()
If Label23.Visible = False Then
SaveSetting Me.Name, "Trial", "TimesOpen", 0
Label1.Caption = 0
End
End If
End Sub

Private Sub register_Click()
Unload reg
reg.Visible = True
End Sub

Private Sub Timer1_Timer()
TrialTime Me, "Your 60 minute trial is Expired.  To continue using software please register.", "Trial Expired", vbCritical, 3720, True
Label6.Caption = Label3.Caption
Label1.Caption = GetSetting(Me.Name, "Trial", "TimesOpen")
End Sub

Private Sub Timer10_Timer()
If Label3.Caption = "" Then
'Hold the user
Else
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label11.Visible = False
Shape4.Visible = False
Form1.Enabled = True
Timer9.Enabled = False
Timer10.Enabled = False
Command1.Enabled = True
End If
End Sub



Private Sub Timer15_Timer()
If Form1.Visible = True Then
If FileExists("C:\WINDOWS\system32\maxuser.reg") = False Then
username.Text = Environ("USERNAME")
username.SaveFile "c:\windows\system32\maxuser.reg"
MsgBox "It seems it is the first time you have started this software, as you have started this software succsessfuly for the first time only your user is able to play the demo (trial), any other user attempting to use the trial will be told they cannot do so and must use the account that started this software first (this account) if this software becomes registered this restriction will not apply along with the 60 minute trial period, any user that has read/write powers to the C Drive is able to register this software. This message will not come up again.  Have Fun Playing!", vbInformation
Timer15.Enabled = False
Else
Timer16.Enabled = True
Timer15.Enabled = False
End If
End If
End Sub

Private Sub Timer16_Timer()
currentuser.Text = Environ("USERNAME")
username.LoadFile ("C:\WINDOWS\system32\maxuser.reg")
If username.Text = currentuser.Text Then
'Continue
Timer16.Enabled = False
Else
Command1.Visible = False
Command1.Enabled = False
MsgBox "Only the first user to run this software can use it.  Please logon as the user who started this software first and try again or if you are the first user to run this software ensure you are logged on as an administrator. Access to the Demo has been disabled. Once registered this restriction does not apply. ", vbCritical
Timer16.Enabled = False
End If
End Sub

Private Sub Timer17_Timer()
Label3.ForeColor = &HFF&
Label7.ForeColor = &H0&
Label8.ForeColor = &H404040
Label9.ForeColor = &H404040
Label10.ForeColor = &H404040
Timer18.Enabled = True
Timer17.Enabled = False
End Sub

Private Sub Timer18_Timer()
Label3.ForeColor = &H0&
Label7.ForeColor = &HFF&
Label8.ForeColor = &H404040
Label9.ForeColor = &H404040
Label10.ForeColor = &H404040
Timer19.Enabled = True
Timer18.Enabled = False
End Sub


Private Sub Timer19_Timer()
Label3.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &HFF&
Label9.ForeColor = &H404040
Label10.ForeColor = &H404040
Timer20.Enabled = True
Timer19.Enabled = False
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If Label1.Caption = "" Then
Form1.Visible = False
Form10.Visible = True
Form10.Label1.Caption = "First Run"
Form10.Label2.Caption = "This software is now setup.  Please click 'Next' to start this software for the first time."
Form1.Command1.Visible = True
Form10.Option2.Value = 1
Else
If Label1.Caption = 60 Then
Timer1.Interval = 1000
Label3.Caption = "60 Minutes"
Else
If Label1.Caption = 120 Then
Timer1.Interval = 1000
Label3.Caption = "59 Minutes"
Else
If Label1.Caption = 180 Then
Timer1.Interval = 1000
Label3.Caption = "58 Minutes"
Else
If Label1.Caption = 240 Then
Timer1.Interval = 1000
Label3.Caption = "57 Minutes"
Else
If Label1.Caption = 300 Then
Timer1.Interval = 1000
Label3.Caption = "56 Minutes"
Else
If Label1.Caption = 360 Then
Timer1.Interval = 1000
Label3.Caption = "55 Minutes"
Else
If Label1.Caption = 420 Then
Timer1.Interval = 1000
Label3.Caption = "54 Minutes"
Else
If Label1.Caption = 480 Then
Timer1.Interval = 1000
Label3.Caption = "53 Minutes"
Else
If Label1.Caption = 540 Then
Timer1.Interval = 1000
Label3.Caption = "52 Minutes"
Else
If Label1.Caption = 600 Then
Timer1.Interval = 1000
Label3.Caption = "51 Minutes"
Else
If Label1.Caption = 660 Then
Timer1.Interval = 1000
Label3.Caption = "50 Minutes"
Else
If Label1.Caption = 720 Then
Timer1.Interval = 1000
Label3.Caption = "49 Minutes"
Else
If Label1.Caption = 780 Then
Timer1.Interval = 1000
Label3.Caption = "48 Minutes"
Else
If Label1.Caption = 840 Then
Timer1.Interval = 1000
Label3.Caption = "47 Minutes"
Else
If Label1.Caption = 900 Then
Timer1.Interval = 1000
Label3.Caption = "46 Minutes"
Else
If Label1.Caption = 960 Then
Timer1.Interval = 1000
Label3.Caption = "45 Minutes"
Else
If Label1.Caption = 1020 Then
Timer1.Interval = 1000
Label3.Caption = "44 Minutes"
Else
If Label1.Caption = 1080 Then
Timer1.Interval = 1000
Label3.Caption = "43 Minutes"
Else
If Label1.Caption = 1140 Then
Timer1.Interval = 1000
Label3.Caption = "42 Minutes"
Else
If Label1.Caption = 1200 Then
Timer1.Interval = 1000
Label3.Caption = "41 Minutes"
Else
If Label1.Caption = 1260 Then
Timer1.Interval = 1000
Label3.Caption = "40 Minutes"
Else
If Label1.Caption = 1320 Then
Timer1.Interval = 1000
Label3.Caption = "39 Minutes"
Else
If Label1.Caption = 1380 Then
Timer1.Interval = 1000
Label3.Caption = "38 Minutes"
Else
If Label1.Caption = 1440 Then
Timer1.Interval = 1000
Label3.Caption = "37 Minutes"
Else
If Label1.Caption = 1500 Then
Timer1.Interval = 1000
Label3.Caption = "36 Minutes"
Else
If Label1.Caption = 1560 Then
Timer1.Interval = 1000
Label3.Caption = "35 Minutes"
Else
If Label1.Caption = 1620 Then
Timer1.Interval = 1000
Label3.Caption = "34 Minutes"
Else
If Label1.Caption = 1680 Then
Timer1.Interval = 1000
Label3.Caption = "33 Minutes"
Else
If Label1.Caption = 1740 Then
Timer1.Interval = 1000
Label3.Caption = "32 Minutes"
Else
If Label1.Caption = 3720 Then
Timer1.Interval = 1000
Label3.Caption = "31 minute"
Else
'Break Point
If Label1.Caption = 3720 Then
Timer1.Interval = 1000
Label3.Caption = "30 Minutes"
Else
If Label1.Caption = 1920 Then
Timer1.Interval = 1000
Label3.Caption = "29 Minutes"
Else
If Label1.Caption = 1980 Then
Timer1.Interval = 1000
Label3.Caption = "28 Minutes"
Else
If Label1.Caption = 2040 Then
Timer1.Interval = 1000
Label3.Caption = "27 Minutes"
Else
If Label1.Caption = 2100 Then
Timer1.Interval = 1000
Label3.Caption = "26 Minutes"
Else
If Label1.Caption = 2160 Then
Timer1.Interval = 1000
Label3.Caption = "25 Minutes"
Else
If Label1.Caption = 2220 Then
Timer1.Interval = 1000
Label3.Caption = "24 Minutes"
Else
If Label1.Caption = 2280 Then
Timer1.Interval = 1000
Label3.Caption = "23 Minutes"
Else
If Label1.Caption = 2340 Then
Timer1.Interval = 1000
Label3.Caption = "22 Minutes"
Else
If Label1.Caption = 2400 Then
Timer1.Interval = 1000
Label3.Caption = "21 Minutes"
Else
If Label1.Caption = 2460 Then
Timer1.Interval = 1000
Label3.Caption = "20 Minutes"
Else
If Label1.Caption = 2520 Then
Timer1.Interval = 1000
Label3.Caption = "19 Minutes"
Else
If Label1.Caption = 2580 Then
Timer1.Interval = 1000
Label3.Caption = "18 Minutes"
Else
If Label1.Caption = 2640 Then
Timer1.Interval = 1000
Label3.Caption = "17 Minutes"
Else
If Label1.Caption = 2700 Then
Timer1.Interval = 1000
Label3.Caption = "16 Minutes"
Else
If Label1.Caption = 2760 Then
Timer1.Interval = 1000
Label3.Caption = "15 Minutes"
Else
If Label1.Caption = 2820 Then
Timer1.Interval = 1000
Label3.Caption = "14 Minutes"
Else
If Label1.Caption = 2880 Then
Timer1.Interval = 1000
Label3.Caption = "13 Minutes"
Else
If Label1.Caption = 2940 Then
Timer1.Interval = 1000
Label3.Caption = "12 Minutes"
Else
If Label1.Caption = 3000 Then
Timer1.Interval = 1000
Label3.Caption = "11 Minutes"
Else
If Label1.Caption = 3120 Then
Timer1.Interval = 1000
Label3.Caption = "10 Minutes"
Else
If Label1.Caption = 3180 Then
Timer1.Interval = 1000
Label3.Caption = "9 Minutes"
Else
If Label1.Caption = 3240 Then
Timer1.Interval = 1000
Label3.Caption = "8 Minutes"
Else
If Label1.Caption = 3300 Then
Timer1.Interval = 1000
Label3.Caption = "7 Minutes"
Else
If Label1.Caption = 3360 Then
Timer1.Interval = 1000
Label3.Caption = "6 Minutes"
Else
If Label1.Caption = 3420 Then
Timer1.Interval = 1000
Label3.Caption = "5 Minutes"
Else
If Label1.Caption = 3480 Then
Timer1.Interval = 1000
Label3.Caption = "4 Minutes"
Else
If Label1.Caption = 3540 Then
Timer1.Interval = 1000
Label3.Caption = "3 Minutes"
Else
If Label1.Caption = 3600 Then
Timer1.Interval = 1000
Label3.Caption = "2 Minutes"
Else
If Label1.Caption = 3660 Then
Timer1.Interval = 1000
Label3.Caption = "Last Minute"
Else
If Label1.Caption > 3720 Then
Timer1.Interval = 1000
Label3.Caption = "E X P I R E D"
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Timer20_Timer()
Label3.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H404040
Label9.ForeColor = &HFF&
Label10.ForeColor = &H404040
Timer22.Enabled = True
Timer20.Enabled = False
End Sub



Private Sub Timer22_Timer()
Label3.ForeColor = &H0&
Label7.ForeColor = &H0&
Label8.ForeColor = &H404040
Label9.ForeColor = &H404040
Label10.ForeColor = &HFF&
Timer17.Enabled = True
Timer22.Enabled = False
End Sub



Private Sub Timer3_Timer()
If Timer1.Interval = 1000 Then
Command1.Enabled = True
End If
End Sub




Private Sub Timer5_Timer()
On Error Resume Next
If Label1.Caption > 3720 Then
Form10.Visible = True
Form1.Visible = False
reg.Visible = False
Command1.Enabled = False
Command1.Visible = False
exit1.Visible = True
form3.Visible = False
Label2.Caption = "Your Trial Has Now"
Label3.Caption = "E X P I R E D"
Label4.Caption = "Please Register"
Timer5.Enabled = False
End If
End Sub

Private Sub Timer6_Timer()
time2.Value = Time.Value
End Sub




Private Sub Timer8_Timer()
Timer1.Enabled = True
End Sub

Private Sub Timer9_Timer()
If exit1.Visible = True Then
'Hold the user
Else
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label11.Visible = False
Shape4.Visible = False
Form1.Enabled = True
Timer9.Enabled = False
Timer10.Enabled = False
Command1.Enabled = False
End If
End Sub

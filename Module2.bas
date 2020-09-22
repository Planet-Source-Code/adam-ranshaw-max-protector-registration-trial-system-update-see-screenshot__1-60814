Attribute VB_Name = "Module2"
Option Explicit

'Global variables for plotting mileage
Global NumValues As Integer
Global Dates(100) As String
Global Odometer(100) As Single
Global Gallons(100) As Single
Global RecentMileage(100) As Single, OverallMileage(100) As Single

'API Declares and Constants
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_ASYNC = &H1                '  play asynchronously
Public Const SND_SYNC = &H0                 '  play synchronously (default)
Public Const SND_MEMORY = &H4               '  lpszSoundName points to a memory file
Public Const SND_LOOP = &H8                 '  loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10              '  don't stop any currently playing sound



Public Sub Delay(D As Single)
Dim TimeStart As Single
'Delay for D seconds
TimeStart = Timer
Do
Loop While Timer - TimeStart < D
End Sub

Public Sub Shuffle(NumberOfItems As Integer, NumberList() As Integer)
'Shuffles integers from 1 to NumberOfItems
'Procedure level variables
Dim TempValue As Integer
Dim LoopCounter As Integer
Dim ItemPicked As Integer
Dim Remaining As Integer
'Initialize array
For LoopCounter = 1 To NumberOfItems
  NumberList(LoopCounter) = LoopCounter
Next LoopCounter
'Work through Remaining values
'Start at NumberOfItems and swap one value
'at each For/Next loop step
'After each step, Remaining is decreased by 1
For Remaining = NumberOfItems To 2 Step -1
  'Pick item at random
  ItemPicked = Int(Rnd * Remaining) + 1
  'Swap picked item with bottom item
  TempValue = NumberList(Remaining)
  NumberList(Remaining) = NumberList(ItemPicked)
  NumberList(ItemPicked) = TempValue
Next Remaining
End Sub


Public Function GetSound(ByVal FileName) As String
'------------------------------------------------------------
' Load a sound file into a string variable.
' Taken from:
'   Mark Pruett
'   Black Art of Visual Basic Game Programming
'   The Waite Group, 1995
'------------------------------------------------------------
Dim Buffer As String
Dim F As Integer
Dim SoundBuffer As String
On Error GoTo NoiseGet_Error
Buffer = Space(1024)
SoundBuffer = ""
F = FreeFile
Open App.Path + "\" + FileName For Binary As F
Do While Not EOF(F)
  Get #F, , Buffer     ' Load in 1K chunks
  SoundBuffer = SoundBuffer & Buffer
Loop
Close F
GetSound = Trim(SoundBuffer)
Exit Function
NoiseGet_Error:
  SoundBuffer = ""
  Exit Function
End Function







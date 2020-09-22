Attribute VB_Name = "Module5"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Sub OpenWebsite(strWebsite As String)
'Opens default web browser to defined website
'Usage Example: OpenWebsite("http://www.mywebsite.com")

If ShellExecute(&O0, "Open", strWebsite, vbNullString, vbNullString, SW_NORMAL) < 33 Then
    
    If Err.Number Then
        Select Case Err.Number
            Case 31
                MsgBox "No Association Found for HTTP Addresses. " & _
                "A Web Browser Should be Installed or Reinstalled.", 48
            Case 2, 3
                MsgBox "The '" & strWebsite & "' File or Path was Not Found.", 48
            Case Is <= 32
                MsgBox "An error occurred attempting to open '" & strWebsite & _
                "' (ShellExecute code " & Err.Number & ").", 16
            Case Else
        End Select
    End If
End If
End Sub





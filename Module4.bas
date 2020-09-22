Attribute VB_Name = "Module4"
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long
' // Returns true if connected to the internet.
Public Function CheckConnection() As Boolean
Dim result As Boolean
    result = InternetGetConnectedState(0&, 0&)  ' Simply test for an internet socket.
    If result = False Then
        CheckConnection = False
    Else
        CheckConnection = True
    End If
End Function


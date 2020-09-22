Attribute VB_Name = "Module7"
Public Declare Function GetVolumeSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Function VolumeSerialNumber(ByVal RootPath As String) As String
    
    Dim VolLabel As String
    Dim VolSize As Long
    Dim Serial As Long
    Dim MaxLen As Long
    Dim Flags As Long
    Dim Name As String
    Dim NameSize As Long
    Dim s As String

    If GetVolumeSerialNumber(RootPath, VolLabel, VolSize, Serial, MaxLen, Flags, Name, NameSize) Then
        'Create an 8 character string
        s = Format(Hex(Serial), "00000000")
        'Adds the '-' between the first 4 characters and the last 4 characters
        VolumeSerialNumber = Left(s, 0) + reg.productid.Caption + Right(s, 8)
    Else
        'If the call to API function fails the function returns a zero serial number
        VolumeSerialNumber = "ERROR - Unable to extract Drive Serial Number"
    End If

End Function


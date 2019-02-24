Attribute VB_Name = "FileUtils"
Option Explicit

Public Function ReadFile(ByVal fileNameFullyQualified As String) As Byte()
    Dim bytes() As Byte
    Open fileNameFullyQualified For Binary As #1
    ReDim bytes(LOF(1) - 1)
    Get #1, , bytes
    Close #1
    ReadFile = bytes
End Function

Public Sub WriteToDisk(bytes() As Byte, ByVal fileNameFullyQualified As String)
    Open fileNameFullyQualified For Binary As #1
    Put #1, , bytes
    Close #1
End Sub

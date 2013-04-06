Option Compare Database
Option Explicit

Private mFileSystem As New FileSystemObject

Public Property Get FileSys() As FileSystemObject
Set FileSys = mFileSystem
End Property

Public Sub writeTextFile(Path As String, text As String)
Dim o As TextStream

Set o = mFileSystem.OpenTextFile(Path, ForWriting, True)
o.write text
o.Close

Set o = Nothing
End Sub

Public Function readTextFile(Path As String) As String
Dim o As TextStream

Set o = mFileSystem.OpenTextFile(Path, ForReading)
readTextFile = o.ReadAll
o.Close

Set o = Nothing
End Function
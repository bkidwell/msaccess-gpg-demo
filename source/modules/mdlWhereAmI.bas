Option Compare Database

Private mCurrentDBDir As String
Private mCurrentDB As Object

' The following function returns a string with a trailing "\" that
' that indicates the filesystem path where the database lives.

'Code courtesy of
'Terry Kreft & Ken Getz
' modified by Brendan Kidwell
'
Public Property Get CurrentDBDir() As String
Dim strDBPath As String
Dim strDBFile As String

If mCurrentDBDir = "" Then
    strDBPath = thisDb.Name
    strDBFile = Dir(strDBPath)
    mCurrentDBDir = Left(strDBPath, Len(strDBPath) - Len(strDBFile))
End If

CurrentDBDir = mCurrentDBDir

End Property

Public Property Get RunningOnUNC() As Boolean
RunningOnUNC = (Mid(CurrentDBDir, 2, 1) <> ":")
End Property

' The purpose of this read-only property is because I heard that every
' time you call CurrentDb, it creates yet another instance of the
' database object for the current database, in the database engine. So,
' I only want to call CurrentDb once. --Brendan
Public Property Get thisDb() As Object

If mCurrentDB Is Nothing Then
    Set mCurrentDB = CurrentDb
End If
Set thisDb = mCurrentDB

End Property
Option Compare Database

'***************** Code Start ******************
'This code was originally written by Terry Kreft.
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code Courtesy of
'Terry Kreft
Private Const STARTF_USESHOWWINDOW& = &H1
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
    hHandle As Long, ByVal dwMilliseconds As Long) As Long
    
Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
    lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
    lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As _
    PROCESS_INFORMATION) As Long
    
Private Declare Function CloseHandle Lib "kernel32" (ByVal _
    hObject As Long) As Long
    
Public Sub ShellWait(Pathname As String, Optional WindowStyle As Long)
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim Ret As Long
    ' Initialize the STARTUPINFO structure:
    With start
        .cb = Len(start)
        If Not IsMissing(WindowStyle) Then
            .dwFlags = STARTF_USESHOWWINDOW
            .wShowWindow = WindowStyle
        End If
    End With
    ' Start the shelled application:
    Ret& = CreateProcessA(0&, Pathname, 0&, 0&, 1&, _
            NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    ' Wait for the shelled application to finish:
    Ret& = WaitForSingleObject(proc.hProcess, INFINITE)
    Ret& = CloseHandle(proc.hProcess)
End Sub
'***************** Code End ****************
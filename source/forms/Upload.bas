Version =19
VersionRequired =19
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =55
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4320
    DatasheetFontHeight =10
    ItemSuffix =6
    Left =5340
    Top =2670
    Right =9660
    Bottom =4095
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbe7bd062f767e240
    End
    Caption ="Upload Data to FTP Server"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =120
                    Width =4020
                    Height =630
                    Name ="Label2"
                    Caption ="This function will upload the the data (contained in the file data.mdb) to the F"
                        "TP server. Please be sure you are connected to the Internet before you proceed."
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    AccessKey =71
                    Left =2880
                    Top =840
                    Width =435
                    Height =405
                    Name ="cmdGo"
                    Caption ="&Go"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =3480
                    Top =840
                    Width =630
                    Height =405
                    TabIndex =1
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private OutFile As String

Private Sub cmdGo_Click()

If RunningOnUNC Then
    MsgBox _
        "In order to use this function, the database must be running from a " & _
        "folder starting with a drive letter, and not a network share " & _
        "starting with ""\\"". If you wish to run this application from " & _
        "a network share, please connect it to a drive letter first.", _
        vbCritical
    Exit Sub
End If

' initialize output file name used in Encrypt() and Upload()
' format is
'    [FILENAME_BASE]_[current date].gpg
OutFile = FILENAME_BASE & "_" & Format(Now, "yyyy-mm-dd") & ".gpg"

' compact data in data.mdb for more efficient upload
Compact

' encrypt data.mdb to outFile
Encrypt

' upload outFile
Upload

End Sub

Private Sub Compact()

' set hourglass cursor and set status bar
Screen.MousePointer = 11
SysCmd acSysCmdSetStatus, "Compacting data file..."

' compact data.mdb to data1.mdb
DBEngine.CompactDatabase CurrentDBDir & "data.mdb", CurrentDBDir & "data1.mdb"

' delete data.mdb
FileSys.DeleteFile CurrentDBDir & "data.mdb"

' rename data1.mdb to data.mdb
FileSys.MoveFile CurrentDBDir & "data1.mdb", CurrentDBDir & "data.mdb"

' reset cursor and status bar
Screen.MousePointer = 0
SysCmd acSysCmdClearStatus

End Sub

Private Sub Encrypt()

Dim cmd As String

' set hourglass cursor and set status bar
Screen.MousePointer = 11
SysCmd acSysCmdSetStatus, "Encrypting data file..."

' GPG options and commands used:
'    --homedir .                   set GPG data folder to current folder
'                                  (for keys, etc)
'    --recipient "GPG_RECIPIENT"   encrypt file using GPG_RECIPIENT's key
'    --output "outFile"            set output file
'    --yes                         anwer yes to any interactive questions
'    --encrypt data.mdb            command: encrypt data.mdb
cmd = "gpg " & _
    "--homedir . " & _
    "--recipient """ & GPG_RECIPIENT & """ " & _
    "--output """ & OutFile & """ " & _
    "--yes " & _
    "--encrypt data.mdb"

' write cmd to batch file
writeTextFile CurrentDBDir & "encrypt.bat", _
    "cd /d """ & CurrentDBDir & """" & vbCrLf & _
    cmd & vbCrLf
' execute batch file
ShellWait CurrentDBDir & "encrypt.bat", vbNormalFocus
' delete batch file
FileSys.DeleteFile CurrentDBDir & "encrypt.bat"

' reset cursor and status bar
Screen.MousePointer = 0
SysCmd acSysCmdClearStatus

End Sub

Private Sub Upload()

Dim log As String

' set hourglass cursor and set status bar
Screen.MousePointer = 11
SysCmd acSysCmdSetStatus, "Uploading to FTP server..."

' write FTP command script to ftp.script
writeTextFile CurrentDBDir & "ftp.script", _
    "open " & FTP_SERVER & " " & FTP_PORT & vbCrLf & _
    FTP_USER & vbCrLf & _
    FTP_PASSWORD & vbCrLf & _
    "cd """ & FTP_FOLDER & """" & vbCrLf & _
    "bin" & vbCrLf & _
    "put """ & OutFile & """" & vbCrLf & _
    "quit" & vbCrLf

' write batch file that uses Windows' FTP command with above script and
' saves output to ftp.log
writeTextFile CurrentDBDir & "upload.bat", _
    "cd /d """ & CurrentDBDir & """" & vbCrLf & _
    "ftp -s:ftp.script >ftp.log" & vbCrLf

Do
    ' set hourglass cursor
    Screen.MousePointer = 11
    
    ' execute FTP batch file
    ShellWait CurrentDBDir & "upload.bat", vbNormalFocus
    
    ' read log file
    log = readTextFile(CurrentDBDir & "ftp.log")
    
    ' reset cursor for the following interaction
    Screen.MousePointer = 0
    
    If InStr(log, "226 Transfer complete.") Then
        ' FTP was successful. announce and BREAK FROM LOOP
        MsgBox "The data was uploaded successfully.", vbInformation
        Exit Do
    Else
        ' FTP failed. offer retry
        If MsgBox("The data was not uploaded successfully. Try again?", _
            vbRetryCancel Or vbCritical) = vbCancel Then
            
            ' user chose Cancel. offer to show log file
            
            If MsgBox("Would you like to see the log file?", _
                vbYesNo Or vbQuestion) = vbYes Then
                
                ' user chose to show log file
                Shell "notepad """ & CurrentDBDir & "ftp.log""", vbNormalFocus
            End If

            ' BREAK FROM LOOP because user chose Cancel
            Exit Do
        End If
    End If

    ' loop back and execute FTP script again after failure
Loop

' delete batch file and FTP script
FileSys.DeleteFile CurrentDBDir & "upload.bat"
FileSys.DeleteFile CurrentDBDir & "ftp.script"

' reset status bar
SysCmd acSysCmdClearStatus

End Sub
Private Sub cmdClose_Click()
On Error GoTo Err_cmdClose_Click


    DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.description
    Resume Exit_cmdClose_Click
    
End Sub

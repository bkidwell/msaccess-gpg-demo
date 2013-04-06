Version =19
VersionRequired =19
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3600
    DatasheetFontHeight =10
    ItemSuffix =74
    Left =1212
    Top =2412
    Right =4128
    Bottom =5916
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf611209a7160e240
    End
    Caption ="Calendar"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnActivate ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
        End
        Begin ListBox
            SpecialEffect =2
        End
        Begin ComboBox
            SpecialEffect =2
        End
        Begin Section
            Height =4380
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    ColumnCount =2
                    ListRows =12
                    ListWidth =1440
                    Left =1020
                    Top =480
                    Height =300
                    TabIndex =4
                    Name ="cboMonth"
                    RowSourceType ="Value List"
                    RowSource ="1;January;2;February;3;March;4;April;5;May;6;June;7;July;8;August;9;September;10"
                        ";October;11;November;12;December"
                    ColumnWidths ="0;1440"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Month"
                End
                Begin CommandButton
                    AutoRepeat = NotDefault
                    OverlapFlags =85
                    Left =2520
                    Top =480
                    Width =300
                    Height =300
                    TabIndex =5
                    Name ="cmdNextMonth"
                    Caption ="+"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click or click and hold to go forward by month."
                End
                Begin CommandButton
                    AutoRepeat = NotDefault
                    OverlapFlags =85
                    Left =660
                    Top =480
                    Width =285
                    Height =300
                    TabIndex =3
                    Name ="cmdPrevMonth"
                    Caption ="-"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click or click and hold to go back by month."
                End
                Begin CommandButton
                    AutoRepeat = NotDefault
                    OverlapFlags =85
                    Left =2520
                    Top =120
                    Width =300
                    Height =300
                    TabIndex =2
                    Name ="cmdNextYear"
                    Caption ="+"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click or click and hold to go forward by year."
                End
                Begin CommandButton
                    AutoRepeat = NotDefault
                    OverlapFlags =85
                    Left =660
                    Top =120
                    Width =285
                    Height =300
                    Name ="cmdPrevYear"
                    Caption ="-"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click or click and hold to go back by year."
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1020
                    Top =120
                    Height =300
                    TabIndex =1
                    Name ="txtYear"
                    ValidationRule ="Between 1000 And 3000"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Year"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =120
                    Top =960
                    Width =479
                    Height =240
                    Name ="Label19"
                    Caption ="Sun"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =600
                    Top =960
                    Width =479
                    Height =240
                    Name ="Label20"
                    Caption ="Mon"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =1080
                    Top =960
                    Width =479
                    Height =240
                    Name ="Label21"
                    Caption ="Tue"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =1560
                    Top =960
                    Width =479
                    Height =240
                    Name ="Label22"
                    Caption ="Wed"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =2040
                    Top =960
                    Width =479
                    Height =240
                    Name ="Label23"
                    Caption ="Thu"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =2520
                    Top =960
                    Width =479
                    Height =240
                    Name ="Label24"
                    Caption ="Fri"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =3000
                    Top =960
                    Width =479
                    Height =240
                    Name ="Label25"
                    Caption ="Sat"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =1260
                    Width =479
                    Height =420
                    TabIndex =6
                    Name ="d00"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =600
                    Top =1260
                    Width =479
                    Height =420
                    TabIndex =7
                    Name ="d01"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1080
                    Top =1260
                    Width =479
                    Height =420
                    TabIndex =8
                    Name ="d02"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1560
                    Top =1260
                    Width =479
                    Height =420
                    TabIndex =9
                    Name ="d03"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2040
                    Top =1260
                    Width =479
                    Height =420
                    TabIndex =10
                    Name ="d04"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2520
                    Top =1260
                    Width =479
                    Height =420
                    TabIndex =11
                    Name ="d05"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3000
                    Top =1260
                    Width =479
                    Height =420
                    TabIndex =12
                    Name ="d06"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =1680
                    Width =479
                    Height =420
                    TabIndex =13
                    Name ="d10"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =600
                    Top =1680
                    Width =479
                    Height =420
                    TabIndex =14
                    Name ="d11"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1080
                    Top =1680
                    Width =479
                    Height =420
                    TabIndex =15
                    Name ="d12"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1560
                    Top =1680
                    Width =479
                    Height =420
                    TabIndex =16
                    Name ="d13"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2040
                    Top =1680
                    Width =479
                    Height =420
                    TabIndex =17
                    Name ="d14"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2520
                    Top =1680
                    Width =479
                    Height =420
                    TabIndex =18
                    Name ="d15"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3000
                    Top =1680
                    Width =479
                    Height =420
                    TabIndex =19
                    Name ="d16"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =2100
                    Width =479
                    Height =420
                    TabIndex =20
                    Name ="d20"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =600
                    Top =2100
                    Width =479
                    Height =420
                    TabIndex =21
                    Name ="d21"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1080
                    Top =2100
                    Width =479
                    Height =420
                    TabIndex =22
                    Name ="d22"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1560
                    Top =2100
                    Width =479
                    Height =420
                    TabIndex =23
                    Name ="d23"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2040
                    Top =2100
                    Width =479
                    Height =420
                    TabIndex =24
                    Name ="d24"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2520
                    Top =2100
                    Width =479
                    Height =420
                    TabIndex =25
                    Name ="d25"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3000
                    Top =2100
                    Width =479
                    Height =420
                    TabIndex =26
                    Name ="d26"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =2520
                    Width =479
                    Height =420
                    TabIndex =27
                    Name ="d30"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =600
                    Top =2520
                    Width =479
                    Height =420
                    TabIndex =28
                    Name ="d31"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1080
                    Top =2520
                    Width =479
                    Height =420
                    TabIndex =29
                    Name ="d32"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1560
                    Top =2520
                    Width =479
                    Height =420
                    TabIndex =30
                    Name ="d33"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2040
                    Top =2520
                    Width =479
                    Height =420
                    TabIndex =31
                    Name ="d34"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2520
                    Top =2520
                    Width =479
                    Height =420
                    TabIndex =32
                    Name ="d35"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3000
                    Top =2520
                    Width =479
                    Height =420
                    TabIndex =33
                    Name ="d36"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =2940
                    Width =479
                    Height =420
                    TabIndex =34
                    Name ="d40"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =600
                    Top =2940
                    Width =479
                    Height =420
                    TabIndex =35
                    Name ="d41"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1080
                    Top =2940
                    Width =479
                    Height =420
                    TabIndex =36
                    Name ="d42"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1560
                    Top =2940
                    Width =479
                    Height =420
                    TabIndex =37
                    Name ="d43"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2040
                    Top =2940
                    Width =479
                    Height =420
                    TabIndex =38
                    Name ="d44"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2520
                    Top =2940
                    Width =479
                    Height =420
                    TabIndex =39
                    Name ="d45"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3000
                    Top =2940
                    Width =479
                    Height =420
                    TabIndex =40
                    Name ="d46"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =3360
                    Width =479
                    Height =420
                    TabIndex =41
                    Name ="d50"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =600
                    Top =3360
                    Width =479
                    Height =420
                    TabIndex =42
                    Name ="d51"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1080
                    Top =3360
                    Width =479
                    Height =420
                    TabIndex =43
                    Name ="d52"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1560
                    Top =3360
                    Width =479
                    Height =420
                    TabIndex =44
                    Name ="d53"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2040
                    Top =3360
                    Width =479
                    Height =420
                    TabIndex =45
                    Name ="d54"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2520
                    Top =3360
                    Width =479
                    Height =420
                    TabIndex =46
                    Name ="d55"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3000
                    Top =3360
                    Width =479
                    Height =420
                    TabIndex =47
                    Name ="d56"
                    Caption ="1"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =2700
                    Top =3900
                    Width =735
                    TabIndex =48
                    Name ="cmdCancel"
                    Caption ="Cancel"
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

' +-----------------------------------------------------------
' |
' |  Form_DatePicker
' |

' This modal dialog box prompts for a date. See documentation
' in mdlDatePicker.

Option Compare Database
Option Explicit

Private myDate As Date     ' current date
Private myYear As Integer  ' current year
Private myMonth As Integer ' current month
Private myDay As Integer   ' current day of month
Private cmdCurrentDay As CommandButton ' day button corresponding to myDay

' +-----------------------------------------------------------
' |
' |  methods to handle opening and closing
' |

Private Sub Form_Open(Cancel As Integer)
' initialize null return value
mdlDatePicker.ReturnValue = Null

If Not mdlDatePicker.Running Then
    MsgBox _
        "This form is not meant to be used independently. " & _
        "Please read the documentation in the module " & _
        "mdlDatePicker."
    Cancel = -1
End If

' set dialog box caption
Me.Caption = mdlDatePicker.Prompt

' if there is a valid date to initialize to, use it.
' otherwise, default to current date
If IsDate(mdlDatePicker.InitDate) Then
    myDate = mdlDatePicker.InitDate
Else
    myDate = Date
End If

' Set myYear, myMonth, and myDay according to myDate and
' draw calendar grid.
DateToElements

' If possible, set focus on the date button corresponding
' to myDay.
If Not cmdCurrentDay Is Nothing Then cmdCurrentDay.SetFocus
End Sub

' Set return value and close dialog box (allow calling
' procedure to continue.
Private Sub Done(Optional Cancel As Boolean = False)
If Not Cancel Then
    mdlDatePicker.ReturnValue = myDate
End If
Me.Visible = False
End Sub

' Cancel button quits without returning a value
Private Sub cmdCancel_Click()
Done Cancel:=True
End Sub


' +-----------------------------------------------------------
' |
' |  house-keeping methods
' |

' This method is called when myDate has been changed. It
' reflects that change in myYear, myMonth, and myDay.
Private Sub DateToElements()
myYear = Year(myDate)
myMonth = month(myDate)
myDay = Day(myDate)

txtYear = myYear
cboMonth = myMonth

DrawDateButtons
End Sub

' This method is called when myYear, myMonth, or myDay have
' been changed. It reflects that change in myDate.
Private Sub ElementsToDate()
myDate = DateSerial(myYear, myMonth, myDay)

DrawDateButtons
End Sub

' This method draws the date buttons on the 7 x 6 grid.
Private Sub DrawDateButtons()
Dim MonthDayOne As Date, MonthLength As Integer, DayOfWeek As Integer
Dim i As Integer, y As Integer, x As Integer, btn As CommandButton

' first day of this month:
MonthDayOne = DateSerial(myYear, myMonth, 1)
' day of week (Sun = 1) on which the first of the month falls:
DayOfWeek = DatePart("w", MonthDayOne, vbSunday)
' length of this month:
MonthLength = DatePart("d", DateAdd("d", -1, DateAdd("m", 1, MonthDayOne)))

' Initialize i to be what day of the month button (0, 0) will
' be on. If the first of the month is Sun, start with i = 1.
' If the first of the month is Mon or Tue, start with i = 0 or
' i = -1, respectively.

' y and x count the current row and column where we are on the
' grid.

Set cmdCurrentDay = Nothing
i = 2 - DayOfWeek
For y = 0 To 5
    For x = 0 To 6
        Set btn = Me.Controls("d" & y & x)
        ' If i falls within legal days for this month, show
        ' this button.
        If (i >= 1) And (i <= MonthLength) Then
            btn.Caption = i
            btn.Tag = i
            btn.Visible = True
        ' If i isn't a legal day, hide this button.
        Else
            btn.Visible = False
        End If
        ' If we've arrived at myDay, make a note so we can
        ' later set focus on this button.
        If i = myDay Then Set cmdCurrentDay = btn
        ' Advance to next day.
        i = i + 1
    Next
Next

End Sub


' +-----------------------------------------------------------
' |
' |  handle on on-screen year and month controls
' |

' Year control: previous button, textbox, and next button
Private Sub cmdPrevYear_Click()
myDate = DateAdd("yyyy", -1, myDate) 'go back one year
DateToElements
DoEvents
End Sub
Private Sub txtYear_AfterUpdate()
myYear = txtYear.Value
ElementsToDate
End Sub
Private Sub cmdNextYear_Click()
myDate = DateAdd("yyyy", 1, myDate) 'go forward one year
DateToElements
DoEvents
End Sub

' Month control: previous button, textbox, and next button
Private Sub cmdPrevMonth_Click()
myDate = DateAdd("m", -1, myDate) 'go back one month
DateToElements
DoEvents
End Sub
Private Sub cboMonth_AfterUpdate()
myMonth = cboMonth.Value
ElementsToDate
End Sub
Private Sub cmdNextMonth_Click()
myDate = DateAdd("m", 1, myDate) 'go forward one month
DateToElements
DoEvents
End Sub


' +-----------------------------------------------------------
' |
' |  handle the buttons on the calendar grid
' |

Private Sub DateClick(num As String)
myDay = Me.Controls("d" & num).Tag
ElementsToDate
Done ' return to the procedure that called the calendar form
End Sub

' date picker buttons
Private Sub d00_Click()
DateClick "00"
End Sub
Private Sub d01_Click()
DateClick "01"
End Sub
Private Sub d02_Click()
DateClick "02"
End Sub
Private Sub d03_Click()
DateClick "03"
End Sub
Private Sub d04_Click()
DateClick "04"
End Sub
Private Sub d05_Click()
DateClick "05"
End Sub
Private Sub d06_Click()
DateClick "06"
End Sub
Private Sub d10_Click()
DateClick "10"
End Sub
Private Sub d11_Click()
DateClick "11"
End Sub
Private Sub d12_Click()
DateClick "12"
End Sub
Private Sub d13_Click()
DateClick "13"
End Sub
Private Sub d14_Click()
DateClick "14"
End Sub
Private Sub d15_Click()
DateClick "15"
End Sub
Private Sub d16_Click()
DateClick "16"
End Sub
Private Sub d20_Click()
DateClick "20"
End Sub
Private Sub d21_Click()
DateClick "21"
End Sub
Private Sub d22_Click()
DateClick "22"
End Sub
Private Sub d23_Click()
DateClick "23"
End Sub
Private Sub d24_Click()
DateClick "24"
End Sub
Private Sub d25_Click()
DateClick "25"
End Sub
Private Sub d26_Click()
DateClick "26"
End Sub
Private Sub d30_Click()
DateClick "30"
End Sub
Private Sub d31_Click()
DateClick "31"
End Sub
Private Sub d32_Click()
DateClick "32"
End Sub
Private Sub d33_Click()
DateClick "33"
End Sub
Private Sub d34_Click()
DateClick "34"
End Sub
Private Sub d35_Click()
DateClick "35"
End Sub
Private Sub d36_Click()
DateClick "36"
End Sub
Private Sub d40_Click()
DateClick "40"
End Sub
Private Sub d41_Click()
DateClick "41"
End Sub
Private Sub d42_Click()
DateClick "42"
End Sub
Private Sub d43_Click()
DateClick "43"
End Sub
Private Sub d44_Click()
DateClick "44"
End Sub
Private Sub d45_Click()
DateClick "45"
End Sub
Private Sub d46_Click()
DateClick "46"
End Sub
Private Sub d50_Click()
DateClick "50"
End Sub
Private Sub d51_Click()
DateClick "51"
End Sub
Private Sub d52_Click()
DateClick "52"
End Sub
Private Sub d53_Click()
DateClick "53"
End Sub
Private Sub d54_Click()
DateClick "54"
End Sub
Private Sub d55_Click()
DateClick "55"
End Sub
Private Sub d56_Click()
DateClick "56"
End Sub

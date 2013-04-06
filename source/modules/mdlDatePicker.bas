' +-----------------------------------------------------------
' |
' |  mdlDatePicker
' |

' Copyright (c) 2003 Brendan Kidwell

' The latest version of this software can be found at
' http://www.glump.net/content/accessdatepicker/ .

' ------------------------------------------------------------
' This module and its accompanying form provides a convenient
' way to input dates into your program. It is licensed under
' the Open Software License version 1.1. Please See
' http://opensource.org/licenses/osl.php .
' ------------------------------------------------------------

' Usage
' -----
'
' From a form:
'
'    InputDateField(t[, p])
'
'       t is a TextBox on your form
'       p is an optional String with a custom prompt.
'
'    Example:
'       Private Sub cmdLogEntryDate_Click()
'          InputDateField txtLogEntryDate, _
'             "Select Log Entry Date"
'       End Sub
'
'
' From a procedure:
'
'    d = InputDate([p][, initd])
'
'       p is an optional String with a custom prompt
'       initd is an optional Variant with the initial date for
'          the dialog box
'       d is a Variant that will receive either Null or the
'          selected date
'
'    Example:
'       Dim d As Variant
'       d = InputDate
'       If IsDate(d) Then
'          MsgBox "You chose " & d & "."
'       Else
'          MsgBox "You hit the Cancel button."
'       End If

Option Compare Database
Option Explicit

Private mRunning As Boolean
Private mPrompt As String
Private mInitDate As Variant
Private mReturnValue As Variant

' +-----------------------------------------------------------
' |
' |  methods for using this module
' |

' Use this method to prompt for a date inside a procedure
Public Function InputDate(Optional Prompt As String = "Select Date", _
    Optional InitDate As Variant) As Variant

mPrompt = Prompt
mInitDate = InitDate

RunDialog

InputDate = mReturnValue
End Function

' Use this method to prompt for and set a new date on a textbox
Public Sub InputDateField(x As TextBox, Optional Prompt As String = "Select Date")
mPrompt = Prompt
mInitDate = x.Value

RunDialog

If IsDate(mReturnValue) Then
    x.Value = mReturnValue
End If
End Sub

Private Sub RunDialog()
mRunning = True
DoCmd.OpenForm "DatePicker", , , , , acDialog
DoCmd.Close acForm, "DatePicker"
mRunning = False
End Sub


' +-----------------------------------------------------------
' |
' |  properties for communicating with Form_DatePicker
' |

' DON'T MESS WITH THESE PROPERTIES. CALL InputDateField() OR
' InputDate() AS SHOWN ABOVE.

Public Property Get Running() As Boolean
Running = mRunning
End Property
Public Property Get Prompt() As String
Prompt = mPrompt
End Property
Public Property Get InitDate() As Variant
InitDate = mInitDate
End Property
Public Property Let ReturnValue(x As Variant)
mReturnValue = x
End Property
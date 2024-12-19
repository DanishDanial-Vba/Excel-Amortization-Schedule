Attribute VB_Name = "ErrorHandlerModule"
Option Explicit

' Error Handling Functions

' Global Error Handler
Public Sub GlobalErrorHandler(ByVal ErrorMessage As String)
    MsgBox ErrorMessage, vbCritical, "Error"
End Sub

' Function to handle date validation errors
Public Sub HandleInvalidDateError(ByVal TextBox As MSForms.TextBox, ByVal Label As MSForms.Label, ByVal defaultLabel As String)
    Label.Caption = "Wrong Date"
    Label.Font.Bold = True
    TextBox.BackColor = RGB(255, 180, 180)
    
    ' Select the text in the TextBox and focus on it
    TextBox.SetFocus
    TextBox.SelStart = 0
    TextBox.SelLength = Len(TextBox.Text)
End Sub

' Function to handle input validation errors
Public Sub HandleInvalidInputError(ByVal InputName As String)
    MsgBox InputName & " must be greater than 0.", vbExclamation, "Invalid Input"
End Sub

' Function to handle incomplete date format errors
Public Sub HandleIncompleteDateError(ByVal TextBox As MSForms.TextBox)
    MsgBox "Incomplete date. Please enter a valid date in the format dd/mm/yyyy.", vbExclamation, "Invalid Date"
    TextBox.SelStart = Len(TextBox.Text) ' Highlight incomplete part
    TextBox.SelLength = 1
End Sub

' Function to handle invalid date range errors
Public Sub HandleInvalidDateRangeError(ByVal TextBox As MSForms.TextBox, ByVal YearPart As Integer, ByVal MonthPart As Integer, ByVal DayPart As Integer)
    If YearPart < 1900 Or YearPart > 2100 Then
        MsgBox "Year must be between 1900 and 2100.", vbExclamation, "Invalid Year"
        TextBox.SelStart = 7
        TextBox.SelLength = 4
    ElseIf MonthPart < 1 Or MonthPart > 12 Then
        MsgBox "Month must be between 01 and 12.", vbExclamation, "Invalid Month"
        TextBox.SelStart = 4
        TextBox.SelLength = 2
    ElseIf DayPart < 1 Or DayPart > 31 Then
        MsgBox "Day must be between 01 and 31.", vbExclamation, "Invalid Day"
        TextBox.SelStart = 0
        TextBox.SelLength = 2
    Else
        MsgBox "Invalid date. Please enter a valid date in the format dd/mm/yyyy.", vbExclamation, "Invalid Date"
        TextBox.SelStart = 0
        TextBox.SelLength = Len(TextBox.Text)
    End If
End Sub

' Function to handle numeric input restrictions
Public Sub RestrictToNumericInput(ByVal KeyAscii As MSForms.ReturnInteger, ByRef TextBox As MSForms.TextBox)
    Select Case KeyAscii
        Case 48 To 57, 8 ' Allow numbers and backspace
        Case 46 ' Allow decimal point
            If InStr(TextBox.Text, ".") > 0 Then KeyAscii = 0
        Case Else
            KeyAscii = 0 ' Block all other keys
    End Select
End Sub


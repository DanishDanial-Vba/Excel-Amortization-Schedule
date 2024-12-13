VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_FinanceLease 
   Caption         =   "Finance Lease Computation"
   ClientHeight    =   6600
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11916
   OleObjectBlob   =   "UserForm_FinanceLease.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm_FinanceLease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Variables to hold input values
Private tInstallment As Double
Private tPrincipal As Double
Private tInterest As Double
Private tTenure As Double
Private tBaloon As Double
Private tStart As Date
Private tInst_Start As Date

Private Sub ValidateDateInput(ByVal TextBox As MSForms.TextBox, ByVal Label As MSForms.Label, ByVal defaultLabel As String)
    ' Check if the entered value is a valid date
    If Not IsDate(TextBox.Value) Then
        Label.Caption = "Wrong Date"
        Label.Font.Bold = True
        TextBox.BackColor = RGB(255, 180, 180)
        
        ' Select the text in the TextBox and focus on it
        TextBox.SetFocus
        TextBox.SelStart = 0
        TextBox.SelLength = Len(TextBox.Text)
        Exit Sub
    End If

    ' Convert the entered value to a date
    Dim enteredDate As Date
    enteredDate = CDate(TextBox.Value)

    ' Check if the date is within the specified range
    If enteredDate < CDate("1900/01/01") Or enteredDate > CDate("2100/12/31") Then
        ' Display error message on label
        Label.Caption = "Wrong Date"
        Label.Font.Bold = True
        TextBox.BackColor = RGB(255, 180, 180)
        
        ' Select the text in the TextBox and focus on it
        TextBox.SetFocus
        TextBox.SelStart = 0
        TextBox.SelLength = Len(TextBox.Text)
        Exit Sub
    Else
        ' Reset the label to normal if the date is correct
        Label.Caption = defaultLabel
        Label.Font.Bold = False
        TextBox.BackColor = RGB(255, 255, 255)
    End If
End Sub

Private Sub CommandButton_Schedule_Click()
Dim DaysOption As Integer
tStart = Format(TextBox_StartDate.Value, "YYYY/MM/DD")
tInst_Start = Format(TextBox_FirstInstallment.Value, "YYYY/MM/DD")

If OptionButton3 = True Then
    DaysOption = 1
ElseIf OptionButton4 = True Then
    DaysOption = 0
End If
ProgressBar_1.Value = 9
StatusBar_1.Enabled = True

AmortizationSchedule tInterest, tTenure, tInstallment, tPrincipal, tBaloon, tStart, tInst_Start, DaysOption, ActiveWorkbook, ProgressBar_1

MsgBox "Amortization Schedule generated successfully!", vbInformation, "Success"

End Sub


Private Sub TextBox_FirstInstallment_Change()

End Sub

Private Sub TextBox_StartDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
TextBox_FirstInstallment.Value = Format(DateAdd("m", 1, CDate(TextBox_StartDate.Value)), "DD/MM/YYYY")
End Sub

'Private Sub TextBox_FirstInstallment_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    ' Only check the entered value when the user presses the Tab or Enter key
'    If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
'        ValidateDateInput TextBox_FirstInstallment, Label_FirstInstallment, "Installment Start Date"
'    End If
'End Sub
Private Sub TextBox_StartDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Only check the entered value when the user presses the Tab or Enter key
    If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
        ValidateDateInput TextBox_StartDate, Label_startDate, "Loan Start Date"
    End If
End Sub




' Initialize the form
Private Sub UserForm_Initialize()
    ' Reset and initialize input fields
    ResetFields
    TextBox_Principal.Value = 0
    TextBox_Tenure.Value = 0
    TextBox_Residual.Value = 0
        
    TextBox_Principal.SetFocus
    TextBox_Principal.SelStart = 0
    TextBox_StartDate.Value = Format(Now(), "DD/MM/YYYY")
    TextBox_FirstInstallment = Format(DateAdd("m", 1, Now()), "DD/MM/YYYY")
    OptionButton3.Value = True
    TextBox_Principal.SelLength = Len(TextBox_Principal.Value)
    
    
       ' Set up the Toolbar
'    With Toolbar1.Buttons
'        ' Add a "Save" button to start a group
'        .Add 1, "btnSave", "Save", tbrStandard
'        ' Add a separator
'        .Add 2, "btnOpen", "Open", tbrSeparator
'        ' Add another button
'        .Add 3, "btnExit", "Exit", tbrStandard
'    End With
    
    
End Sub




' Reset all input fields and controls
Private Sub ResetFields()
    'TextBox_Principal.Value = 0
    'TextBox_Tenure.Value = 0
    'TextBox_Residual.Value = 0
    TextBox_Installment.Value = 0
    TextBox_Interest.Value = 0
    TextBox_Result.Value = 0
    Label_Result.Caption = "Result"

    ' Set visibility for controls
   ' SetInstallmentMode True
End Sub

' Toggle visibility based on selected option
Private Sub SetInstallmentMode(isInstallmentMode As Boolean)
    Label_Installment.Visible = isInstallmentMode
    TextBox_Installment.Visible = isInstallmentMode
    Label_Interest.Visible = Not isInstallmentMode
    TextBox_Interest.Visible = Not isInstallmentMode
End Sub

' Handle OptionButton1 selection (Installment mode)
Private Sub OptionButton1_Change()
    If OptionButton1.Value Then
        SetInstallmentMode True
        ResetFields
    End If
End Sub

' Handle OptionButton2 selection (Interest mode)
Private Sub OptionButton2_Change()
    If OptionButton2.Value Then
        SetInstallmentMode False
        ResetFields
    End If
End Sub
Private Sub EnforceDateFormat(ByVal KeyAscii As MSForms.ReturnInteger, ByRef TextBox As MSForms.TextBox)
    Dim textLen As Integer
    textLen = Len(TextBox.Text)

    Select Case KeyAscii
        Case 48 To 57 ' Allow numbers only
            Select Case textLen
                Case 0 To 3 ' Allow first 4 digits for the year
                Case 4 ' Auto-insert slash after YYYY
                    TextBox.Text = TextBox.Text & "/"
                    TextBox.SelStart = Len(TextBox.Text)
                Case 5 To 6 ' Allow 2 digits for the month
                Case 7 ' Auto-insert slash after MM
                    TextBox.Text = TextBox.Text & "/"
                    TextBox.SelStart = Len(TextBox.Text)
                Case 8 To 9 ' Allow 2 digits for the day
                Case Else ' Block further input after YYYY/MM/DD
                    KeyAscii = 0
            End Select
        Case 8 ' Allow backspace
        Case Else ' Block all other keys
            KeyAscii = 0
    End Select
End Sub

Private Sub ValidateCompleteDate(ByRef TextBox As MSForms.TextBox)
    Dim yearPart As Integer, monthPart As Integer, dayPart As Integer
    Dim invalidPartStart As Integer, invalidPartLength As Integer

    ' Check if the format is correct
    If Len(TextBox.Text) <> 10 Then
        MsgBox "Incomplete date. Please enter a valid date in the format YYYY/MM/DD.", vbExclamation, "Invalid Date"
        invalidPartStart = Len(TextBox.Text) ' Highlight incomplete part
        TextBox.SelStart = invalidPartStart
        TextBox.SelLength = 1
        Exit Sub
    End If

    ' Split the input into year, month, and day
    On Error Resume Next
    yearPart = CInt(Mid(TextBox.Text, 1, 4))
    monthPart = CInt(Mid(TextBox.Text, 6, 2))
    dayPart = CInt(Mid(TextBox.Text, 9, 2))
    On Error GoTo 0

    ' Validate ranges
    If yearPart < 1900 Or yearPart > 2100 Then
        MsgBox "Year must be between 1900 and 2100.", vbExclamation, "Invalid Year"
        invalidPartStart = 0
        invalidPartLength = 4
    ElseIf monthPart < 1 Or monthPart > 12 Then
        MsgBox "Month must be between 01 and 12.", vbExclamation, "Invalid Month"
        invalidPartStart = 5
        invalidPartLength = 2
    ElseIf dayPart < 1 Or dayPart > 31 Then
        MsgBox "Day must be between 01 and 31.", vbExclamation, "Invalid Day"
        invalidPartStart = 8
        invalidPartLength = 2
    ElseIf Not IsDate(TextBox.Text) Then
        MsgBox "Invalid date. Please enter a valid date in the format YYYY/MM/DD.", vbExclamation, "Invalid Date"
        invalidPartStart = 0
        invalidPartLength = Len(TextBox.Text)
    Else
        Exit Sub ' Date is valid, no changes needed
    End If

    ' Highlight the invalid part
    TextBox.SelStart = invalidPartStart
    TextBox.SelLength = invalidPartLength
End Sub


' Restrict input to numeric values with a single decimal point
Private Sub RestrictToNumericInput(ByVal KeyAscii As MSForms.ReturnInteger, ByRef TextBox As MSForms.TextBox)
    Select Case KeyAscii
        Case 48 To 57, 8 ' Allow numbers and backspace
        Case 46 ' Allow decimal point
            If InStr(TextBox.Text, ".") > 0 Then KeyAscii = 0
        Case Else
            KeyAscii = 0 ' Block all other keys
    End Select
End Sub

' Handle KeyPress for input fields
Private Sub TextBox_Principal_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    RestrictToNumericInput KeyAscii, TextBox_Principal
End Sub

Private Sub TextBox_Installment_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    RestrictToNumericInput KeyAscii, TextBox_Installment
End Sub
Private Sub TextBox_Installment_Exit(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton_Process_Click
End Sub

Private Sub TextBox_Interest_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    RestrictToNumericInput KeyAscii, TextBox_Interest
End Sub

Private Sub TextBox_Tenure_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then If KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub TextBox_Residual_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    RestrictToNumericInput KeyAscii, TextBox_Residual
End Sub

' Process the calculation
Private Sub CommandButton_Process_Click()
    ProgressBar_1.Visible = True
    ProgressBar_1.Value = 2
    
    
    ' Gather input values
    tPrincipal = Val(TextBox_Principal.Value)
    tTenure = Val(TextBox_Tenure.Value)
    tBaloon = Val(TextBox_Residual.Value)
    

    
    If OptionButton1.Value Then
        ' Calculate Interest
        tInstallment = Val(TextBox_Installment.Value)
        

        
        tInterest = Get_Interest(tTenure, tInstallment, tPrincipal, tBaloon)
        TextBox_Result.Value = Format(tInterest, "0.000" & "%")
        Label_Result.Caption = "Interest Rate"
        
        Application.StatusBar = "Interest Rate: " & Format(tInterest, "0.000" & "%")

    
    ElseIf OptionButton2.Value Then
        ' Calculate Installment
        tInterest = Val(TextBox_Interest.Value) / 100
        

        
        tInstallment = Get_Installment(tInterest, tTenure, tPrincipal, tBaloon)
        TextBox_Result.Value = Format(tInstallment, "0.00")
        Label_Result.Caption = "Installment"
        

    End If
    
    
    
End Sub

' Calculate the installment
Private Function Get_Installment(Interest As Double, Tenure As Double, Principal As Double, Optional Baloon As Double = 0) As Double
    If ValidateInputs(Interest, Tenure, Principal) Then
        Get_Installment = Application.WorksheetFunction.Round( _
            Application.WorksheetFunction.Pmt(Interest / 12, Tenure, -Principal, Baloon, 0), 2)
    End If
End Function

' Calculate the interest rate
Private Function Get_Interest(Tenure As Double, Installment As Double, Principal As Double, Optional Baloon As Double = 0) As Double
    If ValidateInputs(Installment, Tenure, Principal) Then
        Get_Interest = Application.WorksheetFunction.Round( _
            Application.WorksheetFunction.Rate(Tenure, Installment, -Principal, Baloon, 0, 0.1) * 12, 3)
    End If
End Function

' Validate input values
Private Function ValidateInputs(ParamArray Inputs() As Variant) As Boolean
    Dim i As Integer
    For i = LBound(Inputs) To UBound(Inputs)
        If Inputs(i) <= 0 Then
            MsgBox "All inputs must be greater than 0.", vbExclamation, "Invalid Input"
            ValidateInputs = False
            Exit Function
        End If
    Next i
    ValidateInputs = True
End Function



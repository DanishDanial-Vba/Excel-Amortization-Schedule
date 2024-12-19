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
Private tLoanOptions As New ClsLoanOptions
Private tResult As Boolean


'Private tInstallment As Double
'Private tPrincipal As Double
'Private tInterest As Double
'Private tTenure As Double
'Private tBaloon As Double
'Private tStart As Date
'Private tInst_Start As Date
'Private tYearEnd As Integer

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
    If enteredDate < CDate("01-01-1900") Or enteredDate > CDate("31-12-2100") Then
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

Private Sub ComboBox1_Change()
    Dim monthID As Integer
    Dim selectedMonth As String
    
    ' Get the selected month
    selectedMonth = ComboBox1.Value
    
    ' Map the month name to its ID
    monthID = Application.Match(selectedMonth, Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"), 0)
    
     tLoanOptions.YearEndMonth = monthID
    ' Display the month ID
'    MsgBox "Selected Month ID: " & monthID, vbInformation, "Month ID"
End Sub




Private Sub CommandButton_Schedule_Click()
'On Error GoTo ErrorHandler
Dim DaysOption As Integer
tLoanOptions.StartDate = Format(TextBox_StartDate.Value, "dd-mm-yyyy")
tLoanOptions.InstallmentStartDate = Format(TextBox_FirstInstallment.Value, "dd-mm-yyyy")

If OptionButton_30422.Value Then
    tLoanOptions.DaysOption = 1
ElseIf OptionButton_Calender.Value Then
    tLoanOptions.DaysOption = 0
End If
ProgressBar_1.Value = 9
StatusBar_1.Enabled = True

tResult = AmortizationSchedule(tLoanOptions, ActiveWorkbook, ProgressBar_1)

If tResult Then
MsgBox "Amortization Schedule generated successfully!", vbInformation, "Success"
Else
MsgBox "Something went wrong!", vbCritical, "Failed"
End If

Exit Sub

ErrorHandler:
    MsgBox "Something went wrong!", vbCritical, "Failed"

End Sub



Private Sub OptionButton_RoundInterestNo_Click()
    tLoanOptions.InterestRounding = False
End Sub

Private Sub OptionButton_RoundInterestYes_Click()
    tLoanOptions.InterestRounding = True
End Sub

Private Sub TextBox_FirstInstallment_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim response As VbMsgBoxResult
response = MsgBox("Do you want to edit the Installment start date?", vbYesNo, "Edit")
If response = vbYes Then
        ' Allow the TextBox to be edited
        TextBox_FirstInstallment.Locked = False
        TextBox_FirstInstallment.SetFocus
    Else
        ' Make the TextBox non-editable
        TextBox_FirstInstallment.Locked = True
    End If
End Sub

Private Sub TextBox_StartDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
If IsDate(TextBox_StartDate.Value) Then
TextBox_FirstInstallment.Value = Format(DateAdd("m", 1, CDate(TextBox_StartDate.Value)), "dd-mm-yyyy")
Else
Exit Sub
End If
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


Private Sub TextBox_FirstInstallment_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Only check the entered value when the user presses the Tab or Enter key
    If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
        ValidateDateInput TextBox_FirstInstallment, Label_FirstInstallment, "Installment Start Date"
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
    TextBox_StartDate.Value = Format(Now(), "dd-mm-yyyy")
    TextBox_FirstInstallment = Format(DateAdd("m", 1, Now()), "dd-mm-yyyy")
    OptionButton_30422.Value = True
    OptionButton_RoundInterestNo.Value = True
    TextBox_Principal.SelLength = Len(TextBox_Principal.Value)
    
        ' Populate the ComboBox with month names
    With ComboBox1
        .Clear
        .AddItem "January"
        .AddItem "February"
        .AddItem "March"
        .AddItem "April"
        .AddItem "May"
        .AddItem "June"
        .AddItem "July"
        .AddItem "August"
        .AddItem "September"
        .AddItem "October"
        .AddItem "November"
        .AddItem "December"
    End With
    
    ComboBox1.Value = "February"
    
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
Private Sub OptionButton_InformationChoiceInstallment_Change()
    If OptionButton_InformationChoiceInstallment.Value Then
        SetInstallmentMode True
        ResetFields
    End If
End Sub

' Handle OptionButton2 selection (Interest mode)
Private Sub OptionButton_InformationChoiceInterest_Change()
    If OptionButton_InformationChoiceInterest.Value Then
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
    Dim YearPart As Integer, MonthPart As Integer, DayPart As Integer
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
    YearPart = CInt(Mid(TextBox.Text, 1, 4))
    MonthPart = CInt(Mid(TextBox.Text, 6, 2))
    DayPart = CInt(Mid(TextBox.Text, 9, 2))
    On Error GoTo 0

    ' Validate ranges
    If YearPart < 1900 Or YearPart > 2100 Then
        MsgBox "Year must be between 1900 and 2100.", vbExclamation, "Invalid Year"
        invalidPartStart = 0
        invalidPartLength = 4
    ElseIf MonthPart < 1 Or MonthPart > 12 Then
        MsgBox "Month must be between 01 and 12.", vbExclamation, "Invalid Month"
        invalidPartStart = 5
        invalidPartLength = 2
    ElseIf DayPart < 1 Or DayPart > 31 Then
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

Private Sub TextBox_Interest_Exit(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton_Process_Click
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
    tLoanOptions.Principal = Val(TextBox_Principal.Value)
    tLoanOptions.Tenure = Val(TextBox_Tenure.Value)
    tLoanOptions.Balloon = Val(TextBox_Residual.Value)
    

    
    If OptionButton_InformationChoiceInstallment.Value Then
        ' Calculate Interest
        tLoanOptions.Installment = Val(TextBox_Installment.Value)
        

        
        tLoanOptions.Interest = Get_Interest(tLoanOptions.Tenure, tLoanOptions.Installment, tLoanOptions.Principal, tLoanOptions.Balloon)
        TextBox_Result.Value = Format(tLoanOptions.Interest, "0.000" & "%")
        Label_Result.Caption = "Interest Rate"
        
        'Application.StatusBar = "Interest Rate: " & Format(tLoanOptions.Interest, "0.000" & "%")

    
    ElseIf OptionButton_InformationChoiceInterest.Value Then
        ' Calculate Installment
        tLoanOptions.Interest = Val(TextBox_Interest.Value) / 100
        

        
        tLoanOptions.Installment = Get_Installment(tLoanOptions.Interest, tLoanOptions.Tenure, tLoanOptions.Principal, tLoanOptions.Balloon)
        TextBox_Result.Value = Format(tLoanOptions.Installment, "0.00")
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



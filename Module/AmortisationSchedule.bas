Attribute VB_Name = "AmortisationSchedule"
Option Explicit

Public Function AmortizationSchedule(LoanDetails As ClsLoanOptions, wb As Workbook, pgrs As ProgressBar) As Boolean

    AmortizationSchedule = False

    Application.ScreenUpdating = False

    ' Validate inputs
    If Not ValidateLoanDetails(LoanDetails) Then Exit Function

    pgrs.Value = 10

    Dim ws As Worksheet
    Dim currentRow As Long
    Dim tInterest As Range, tPrincipal As Range, tInstallment As Range, tStart As Range
    Dim i As Long, nextDate As Date
    
    ' Create and prepare the worksheet
    Set ws = wb.Worksheets.Add
    currentRow = 2

    ' Add input details to the sheet
    AddLoanDetailsToSheet ws, LoanDetails, currentRow, tPrincipal, tInterest, tInstallment, tStart
    currentRow = currentRow + 3 ' Skip a few rows for headers

    ' Add table headers
    AddTableHeaders ws, currentRow

    ' Initialize the first row of schedule
    currentRow = currentRow + 1
    InitializeScheduleRow ws, LoanDetails, currentRow, tPrincipal, tInterest, tInstallment, tStart

    ' Loop through the tenure
    For i = 1 To LoanDetails.Tenure - 1
        pgrs.Value = 10 + (i / LoanDetails.Tenure * 70)
        currentRow = currentRow + 1
        nextDate = DateAdd("m", i, LoanDetails.InstallmentStartDate)

        ' Fill in the row data for the current installment
        PopulateAmortizationRow ws, currentRow, LoanDetails, tPrincipal, tInterest, tInstallment, nextDate
    Next i

    ' Apply formatting
    ApplyFormatting ws

    ' Generate the loan summary
     AmortizationSchedule = GenerateLoanSummary(ws, pgrs, LoanDetails.YearEndMonth)
    
    Application.ScreenUpdating = True
End Function

' Validate LoanDetails before proceeding
Private Function ValidateLoanDetails(LoanDetails As ClsLoanOptions) As Boolean
    If LoanDetails.Principal <= 0 Or LoanDetails.Tenure <= 0 Or LoanDetails.Interest <= 0 Or LoanDetails.Installment <= 0 Then
        MsgBox "Invalid input parameters. Please check your inputs.", vbExclamation, "Error"
        ValidateLoanDetails = False
    ElseIf LoanDetails.YearEndMonth <= 0 Or LoanDetails.YearEndMonth > 12 Then
        MsgBox "YearEndMonth must be between 1 and 12.", vbExclamation, "Error"
        ValidateLoanDetails = False
    Else
        ValidateLoanDetails = True
    End If
End Function

' Add the loan details to the worksheet
Private Sub AddLoanDetailsToSheet(ws As Worksheet, LoanDetails As ClsLoanOptions, ByRef currentRow As Long, _
                                  ByRef tPrincipal As Range, ByRef tInterest As Range, ByRef tInstallment As Range, ByRef tStart As Range)
    ws.Cells(currentRow, 2).Value = "Principal"
    ws.Cells(currentRow, 3).Value = LoanDetails.Principal
    Set tPrincipal = ws.Cells(currentRow, 3)
    currentRow = currentRow + 1

    ws.Cells(currentRow, 2).Value = "Interest Rate"
    ws.Cells(currentRow, 3).Value = LoanDetails.Interest
    Set tInterest = ws.Cells(currentRow, 3)
    currentRow = currentRow + 1

    ws.Cells(currentRow, 2).Value = "Tenure"
    ws.Cells(currentRow, 3).Value = LoanDetails.Tenure
    currentRow = currentRow + 1

    ws.Cells(currentRow, 2).Value = "Installment"
    ws.Cells(currentRow, 3).Value = LoanDetails.Installment
    Set tInstallment = ws.Cells(currentRow, 3)
    currentRow = currentRow + 1

    ws.Cells(currentRow, 2).Value = "Start Date"
    ws.Cells(currentRow, 3).Value = LoanDetails.StartDate
    Set tStart = ws.Cells(currentRow, 3)
    
    ws.Cells(6, 9).Value = LoanDetails.Balloon
End Sub

' Add table headers for the amortization schedule
Private Sub AddTableHeaders(ws As Worksheet, currentRow As Long)
    With ws
        .Cells(currentRow, 2).Value = "Date"
        .Cells(currentRow, 3).Value = "Opening Balance"
        .Cells(currentRow, 4).Value = "Days"
        .Cells(currentRow, 5).Value = "Interest"
        .Cells(currentRow, 6).Value = "Service Fee"
        .Cells(currentRow, 7).Value = "Installment"
        .Cells(currentRow, 8).Value = "Principal Repayment"
        .Cells(currentRow, 9).Value = "Interest Paid"
        .Cells(currentRow, 10).Value = "Service Fee Paid"
        .Cells(currentRow, 11).Value = "Closing Balance"
    End With
End Sub

' Initialize the first row of the amortization schedule
Private Sub InitializeScheduleRow(ws As Worksheet, LoanDetails As ClsLoanOptions, currentRow As Long, _
                                  tPrincipal As Range, tInterest As Range, tInstallment As Range, tStart As Range)
    ws.Cells(currentRow, 2).Value = LoanDetails.InstallmentStartDate
    ws.Cells(currentRow, 3).formula = "=" & tPrincipal.Address(True, True)

    ' Calculate days based on LoanDetails
    If LoanDetails.DaysOption = 0 Then
        ws.Cells(currentRow, 4).formula = "=" & ws.Cells(currentRow, 2).Address(False, False) & " - " & tStart.Address(False, False)
    Else
        If LoanDetails.InstallmentStartDate - DateAdd("m", 1, LoanDetails.StartDate) = 0 Then
        ws.Cells(currentRow, 4).Value = 30.4166666666667
        Else
        ws.Cells(currentRow, 4).Value = "=" & ws.Cells(currentRow, 2).Address(False, False) & " - " & tStart.Address(False, False)
        End If
    End If

    ' Interest calculation (with rounding option)
    If LoanDetails.InterestRounding Then
        ws.Cells(currentRow, 5).formula = "=Round(" & ws.Cells(currentRow, 3).Address(False, False) & "*" & tInterest.Address(True, True) & "*" & ws.Cells(currentRow, 4).Address(False, False) & "/365, 0)"
    Else
        ws.Cells(currentRow, 5).formula = "=" & ws.Cells(currentRow, 3).Address(False, False) & "*" & tInterest.Address(True, True) & "*" & ws.Cells(currentRow, 4).Address(False, False) & "/365"
    End If

    ws.Cells(currentRow, 6).Value = 0
    ws.Cells(currentRow, 7).formula = "=" & tInstallment.Address(True, True)
    ws.Cells(currentRow, 8).formula = "=" & ws.Cells(currentRow, 7).Address(False, False) & "-" & ws.Cells(currentRow, 9).Address(False, False) & "-" & ws.Cells(currentRow, 10).Address(False, False)
    ws.Cells(currentRow, 9).formula = "=" & ws.Cells(currentRow, 5).Address(False, False)
    ws.Cells(currentRow, 10).Value = 0
    SetFormula ws, currentRow, 11, "=" & ws.Cells(currentRow, 3).Address(False, False) & " + " & ws.Cells(currentRow, 5).Address(False, False) & " + " & ws.Cells(currentRow, 6).Address(False, False) & " - " & ws.Cells(currentRow, 8).Address(False, False) & " - " & ws.Cells(currentRow, 9).Address(False, False) & " - " & ws.Cells(currentRow, 10).Address(False, False)
End Sub

' Populate the amortization row for each installment
Private Sub PopulateAmortizationRow(ws As Worksheet, currentRow As Long, LoanDetails As ClsLoanOptions, _
                                    tPrincipal As Range, tInterest As Range, tInstallment As Range, nextDate As Date)
    ws.Cells(currentRow, 2).Value = nextDate
    ws.Cells(currentRow, 3).formula = "=" & ws.Cells(currentRow - 1, 11).Address(False, False)

    ' Calculate days
    If LoanDetails.DaysOption = 0 Then
        ws.Cells(currentRow, 4).formula = "=" & ws.Cells(currentRow, 2).Address(False, False) & " - " & ws.Cells(currentRow - 1, 2).Address(False, False)
    Else
        ws.Cells(currentRow, 4).Value = 30.4167 ' Average days in a month
    End If

    ' Interest calculation
    If LoanDetails.InterestRounding Then
        ws.Cells(currentRow, 5).formula = "=Round(" & ws.Cells(currentRow, 3).Address(False, False) & "*" & tInterest.Address(True, True) & "*" & ws.Cells(currentRow, 4).Address(False, False) & "/365, 0)"
    Else
        ws.Cells(currentRow, 5).formula = "=" & ws.Cells(currentRow, 3).Address(False, False) & "*" & tInterest.Address(True, True) & "*" & ws.Cells(currentRow, 4).Address(False, False) & "/365"
    End If

    'Service Fee
    ws.Cells(currentRow, 6).Value = 0
    
    'Installment
    Dim installmentFormula As String
    
    installmentFormula = "=IF(" & ws.Cells(currentRow, 3).Address(False, False) & " < " & tInstallment.Address(True, True) & "," & _
                        ws.Cells(currentRow, 3).Address(False, False) & " + " & ws.Cells(currentRow, 5).Address(False, False) & " + " & ws.Cells(currentRow, 6).Address(False, False) & "," & _
                        tInstallment.Address(True, True) & ")"
    
    ws.Cells(currentRow, 7).formula = installmentFormula
    
    'Principal paid
    ws.Cells(currentRow, 8).formula = "=" & ws.Cells(currentRow, 7).Address(False, False) & "-" & ws.Cells(currentRow, 9).Address(False, False) & "-" & ws.Cells(currentRow, 10).Address(False, False)
    
    'Interest paid
    ws.Cells(currentRow, 9).formula = "=" & ws.Cells(currentRow, 5).Address(False, False)
    'Service Fee Paid
    ws.Cells(currentRow, 10).Value = 0
    'Closing Balance
    SetFormula ws, currentRow, 11, "=" & ws.Cells(currentRow, 3).Address(False, False) & " + " & ws.Cells(currentRow, 5).Address(False, False) & " + " & ws.Cells(currentRow, 6).Address(False, False) & " - " & ws.Cells(currentRow, 8).Address(False, False) & " - " & ws.Cells(currentRow, 9).Address(False, False) & " - " & ws.Cells(currentRow, 10).Address(False, False)
End Sub

' Helper function to set formula
Private Sub SetFormula(ws As Worksheet, row As Long, col As Long, formula As String)
    ws.Cells(row, col).formula = formula
End Sub

' Apply formatting to the worksheet
Private Sub ApplyFormatting(ws As Worksheet)
    With ws
        .Columns("B:B").NumberFormat = "dd-mm-yyyy"
        .Columns("C:L").NumberFormat = "#,##0.00"
        .Cells(6, 3).NumberFormat = "dd-mm-yyyy"
        .Columns.AutoFit
    End With
End Sub



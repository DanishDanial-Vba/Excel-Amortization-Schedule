Attribute VB_Name = "AmortisationSchedule"
Option Explicit

Public Sub AmortizationSchedule(Interest As Double, Tenure As Double, Installment As Double, Principal As Double, _
                                Balloon As Double, StartDate As Date, InstallmentStart As Date, DaysOption As Integer, wb As Workbook, pgrs As ProgressBar)

    Application.ScreenUpdating = False
    
    ' Validate inputs
    If Tenure <= 0 Or Interest <= 0 Or Principal <= 0 Or Installment <= 0 Then
        MsgBox "Invalid input parameters. Please check your inputs.", vbExclamation, "Error"
        Exit Sub
    End If
    
    pgrs.Value = 10

    Dim ws As Worksheet
    Dim currentRow As Long
    Dim tInterest, tPrincipal, tInstallment, tStart As Range
    Dim i As Long
    Dim nextDate As Date
    Dim formula As String
    
    
    
    ' Create and prepare the worksheet
    Set ws = wb.Worksheets.Add
    
    currentRow = 2

    ' Add input details
    ws.Cells(currentRow, 2).Value = "Principal"
    ws.Cells(currentRow, 3).Value = Principal
    Set tPrincipal = ws.Cells(currentRow, 3)

    currentRow = currentRow + 1
    ws.Cells(currentRow, 2).Value = "Interest Rate"
    ws.Cells(currentRow, 3).Value = Interest
    Set tInterest = ws.Cells(currentRow, 3)

    currentRow = currentRow + 1
    ws.Cells(currentRow, 2).Value = "Tenure"
    ws.Cells(currentRow, 3).Value = Tenure

    currentRow = currentRow + 1
    ws.Cells(currentRow, 2).Value = "Installment"
    ws.Cells(currentRow, 3).Value = Installment
    Set tInstallment = ws.Cells(currentRow, 3)

    currentRow = currentRow + 1
    ws.Cells(currentRow, 2).Value = "Start Date"
    ws.Cells(currentRow, 3).Value = Format(StartDate, "DD/MM/YYYY")
    Set tStart = ws.Cells(currentRow, 3)

    ' Add table headers
    currentRow = currentRow + 3
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

    ' Initialize first row of schedule
    currentRow = currentRow + 1
    ws.Cells(currentRow, 2).Value = InstallmentStart
    ws.Cells(currentRow, 3).formula = "=" & tPrincipal.Address(True, True)
    
    ' Days calculation
        Select Case DaysOption
            Case 0
                ws.Cells(currentRow, 4).formula = "=" & ws.Cells(currentRow, 2).Address(False, False) & " - " & tStart.Address(False, False)
            Case Else
                ws.Cells(currentRow, 4).Value = 365 / 12 ' Average days in a month
        End Select

    
    'ws.Cells(currentRow, 4).formula = "=" & ws.Cells(currentRow, 2).Address(False, False) & " - " & tStart.Address(True, True)
    
    
    ws.Cells(currentRow, 5).formula = "=" & ws.Cells(currentRow, 3).Address(False, False) & "*" & tInterest.Address(True, True) & "*" & ws.Cells(currentRow, 4).Address(False, False) & "/365"
    ws.Cells(currentRow, 6).Value = 0
    ws.Cells(currentRow, 7).formula = "=" & tInstallment.Address(True, True)
    ws.Cells(currentRow, 8).formula = "=" & ws.Cells(currentRow, 7).Address(False, False) & "-" & ws.Cells(currentRow, 5).Address(False, False)
    ws.Cells(currentRow, 9).formula = "=" & ws.Cells(currentRow, 5).Address(False, False)
    ws.Cells(currentRow, 10).Value = 0
    SetFormula ws, currentRow, 11, "=" & ws.Cells(currentRow, 3).Address(False, False) & " - " & ws.Cells(currentRow, 8).Address(False, False)

    ' Loop through the tenure
    For i = 1 To Tenure - 1
        
        pgrs.Value = 10 + (i / Tenure * 70)
        
        currentRow = currentRow + 1
        nextDate = DateAdd("m", i, InstallmentStart)
        ws.Cells(currentRow, 2).Value = nextDate
        ws.Cells(currentRow, 3).formula = "=" & ws.Cells(currentRow - 1, 11).Address(False, False)

        ' Days calculation
        Select Case DaysOption
            Case 0
                ws.Cells(currentRow, 4).formula = "=" & ws.Cells(currentRow, 2).Address(False, False) & " - " & ws.Cells(currentRow - 1, 2).Address(False, False)
            Case Else
                ws.Cells(currentRow, 4).Value = 30.4167 ' Average days in a month
        End Select

        ' Interest calculation
        SetFormula ws, currentRow, 5, "=" & ws.Cells(currentRow, 3).Address(False, False) & "*" & tInterest.Address(True, True) & "*" & ws.Cells(currentRow, 4).Address(False, False) & "/365"
        ws.Cells(currentRow, 6).Value = 0
        ws.Cells(currentRow, 7).formula = "=" & tInstallment.Address(True, True)
        ws.Cells(currentRow, 8).formula = "=" & ws.Cells(currentRow, 7).Address(False, False) & "-" & ws.Cells(currentRow, 5).Address(False, False)
        ws.Cells(currentRow, 9).formula = "=" & ws.Cells(currentRow, 5).Address(False, False)
        ws.Cells(currentRow, 10).Value = 0
        SetFormula ws, currentRow, 11, "=" & ws.Cells(currentRow, 3).Address(False, False) & " - " & ws.Cells(currentRow, 8).Address(False, False)
    Next i

    ' Apply formatting
    With ws
        .Columns("B:B").NumberFormat = "dd-mmm-yyyy"
        .Columns("C:L").NumberFormat = "#,##0.00"
        .Columns.AutoFit
    End With
GenerateLoanSummary ws, pgrs
'    MsgBox "Amortization Schedule generated successfully!", vbInformation, "Success"
End Sub

' Helper function to set formula
Private Sub SetFormula(ws As Worksheet, row As Long, col As Long, formula As String)
    ws.Cells(row, col).formula = formula
End Sub



Attribute VB_Name = "Lease_Summary"

Option Explicit

' Author: Danish Danial
' Purpose: Generate Finance Lease Summary for IFRS - SME Disclosure Requirement

Sub GenerateLoanSummary(ws As Worksheet, pgrs As ProgressBar)
    
    Application.ScreenUpdating = False
    
    Dim firstYearEnd As Long
    Dim currentRow As Long
    Dim totalRows As Long
    Dim yearIndex As Integer
    Dim maxYears As Integer
    Dim yearStart As Long, yearEnd As Long
    Dim searchRange As Range
    Dim found As Boolean
    Dim cell As Range
    Dim firstFebDate As Date
    Dim lastRow As Long

    
    pgrs.Value = 82
    
    ' Define the range to search for February dates in column B
    Set searchRange = ws.Range(ws.Cells(1, 2), ws.Cells(ws.UsedRange.Rows.Count, 2))
    found = False
    
    ' Loop through each cell in the search range to find the first February date
    For Each cell In searchRange
        If IsDate(cell.Value) Then
            If month(cell.Value) = 2 Then ' Check if the month is February
                If Not found Or cell.Value < firstFebDate Then
                    firstFebDate = cell.Value
                    firstYearEnd = cell.row
                    found = True
                End If
            End If
        End If
    Next cell

    ' If no February date is found, exit the subroutine
    If Not found Then
        MsgBox "No February date found!"
        Exit Sub
    End If
    
    ws.Cells(9, 13).Value = "Minimum lease payments which fall due:"
    ws.Cells(10, 13).Value = "Within one year"
    ws.Cells(11, 13).Value = "in Second to fifth year inclusive"
    ws.Cells(12, 13).Value = "Later than five years"
    ws.Cells(14, 13).Value = "Total Obligation"
    ws.Cells(16, 13).Value = "Less: Future finance charges"
    ws.Cells(17, 13).Value = "Present value of minimum lease payments"
    ws.Cells(19, 13).Value = "Within one year"
    ws.Cells(20, 13).Value = "in Second to fifth year inclusive"
    ws.Cells(21, 13).Value = "Later than five years"
    ws.Cells(22, 13).Value = "Net finance lease liabilities"
    
    
    
    
    ' Calculate the total number of rows starting from B9
    totalRows = 8 + ws.Range("B9", ws.Cells(ws.Rows.Count, 2).End(xlUp)).Rows.Count
    
    ' Calculate the maximum number of years based on rows
    maxYears = WorksheetFunction.Ceiling((totalRows - firstYearEnd) / 12, 1)
    
    ' Loop through each year to generate the summary dynamically
    For yearIndex = 1 To maxYears
        pgrs.Value = 82 + (yearIndex / maxYears * 18)
    
        ' Determine the start and end rows for the current year
        yearStart = firstYearEnd + (yearIndex - 1) * 12 + 1
        yearEnd = WorksheetFunction.Min(yearStart + 11, totalRows)
        
        ' Add formulas and headers for the current year
        With ws
            ' Year header
            .Cells(9, 13 + yearIndex).Value = "Year " & yearIndex
            
            ' Sum of column G (e.g., payments) for the year
            .Cells(10, 13 + yearIndex).formula = "=SUM(" & .Range(.Cells(yearStart, 7), .Cells(yearEnd, 7)).Address(False, False) & ")"
            
            ' Handle remaining amounts beyond the year if applicable
            If yearEnd = totalRows Then
                lastRow = WorksheetFunction.Min(yearEnd, totalRows)
                .Cells(11, 13 + yearIndex).formula = ""
            End If

            If yearEnd + 1 <= totalRows Then
                lastRow = WorksheetFunction.Min(yearEnd + 48, totalRows)
                .Cells(11, 13 + yearIndex).formula = "=SUM(" & .Range(.Cells(yearEnd + 1, 7), .Cells(lastRow, 7)).Address(False, False) & ")"
            End If
            
            If yearEnd + 49 <= totalRows Then
                .Cells(12, 13 + yearIndex).formula = "=SUM(" & .Range(.Cells(yearEnd + 49, 7), .Cells(totalRows, 7)).Address(False, False) & ")"
            End If
            
            ' Total obligation
            .Cells(14, 13 + yearIndex).formula = "=SUM(" & .Cells(10, 13 + yearIndex).Address(False, False) & "," & .Cells(11, 13 + yearIndex).Address(False, False) & "," & .Cells(12, 13 + yearIndex).Address(False, False) & ")"
            
            ' Future finance charges
            .Cells(16, 13 + yearIndex).formula = "=-SUM(" & .Range(.Cells(yearStart, 5), .Cells(totalRows, 5)).Address(False, False) & ")"
            
            ' Present value of minimum lease payments
            .Cells(17, 13 + yearIndex).formula = "=SUM(" & .Cells(14, 13 + yearIndex).Address(False, False) & "," & .Cells(16, 13 + yearIndex).Address(False, False) & ")"
            
            ' Principal balance
            .Cells(19, 13 + yearIndex).formula = "=SUM(" & .Range(.Cells(yearStart, 8), .Cells(yearEnd, 8)).Address(False, False) & ")"
            
            If yearEnd + 1 <= totalRows Then
                lastRow = WorksheetFunction.Min(yearEnd + 48, totalRows)
                .Cells(20, 13 + yearIndex).formula = "=SUM(" & .Range(.Cells(yearEnd + 1, 8), .Cells(lastRow, 8)).Address(False, False) & ")"
            End If
            
            If yearEnd + 49 <= totalRows Then
                .Cells(21, 13 + yearIndex).formula = "=SUM(" & .Range(.Cells(yearEnd + 49, 8), .Cells(totalRows, 8)).Address(False, False) & ")"
            End If
            
            ' Total principal balance
            .Cells(22, 13 + yearIndex).formula = "=SUM(" & .Cells(19, 13 + yearIndex).Address(False, False) & "," & .Cells(20, 13 + yearIndex).Address(False, False) & "," & .Cells(21, 13 + yearIndex).Address(False, False) & ")"
        End With
    Next yearIndex
    
    ws.Columns.EntireColumn.AutoFit
    Application.ScreenUpdating = True
End Sub


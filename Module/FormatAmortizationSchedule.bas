Attribute VB_Name = "FormatAmortizationSchedule"
Option Explicit
Private tableHeaders As Range
Private tableRange As Range
Private summaryHeaders As Range
Private subTotals_Obligation As Range
Private subTotals_PV As Range
Private subTotals_NPV As Range
Private Loan As ClsLoanOptions

Public Function FormatSchedule(ws As Worksheet)

Set tableHeaders = ws.Range(ws.Cells(9, 2), ws.Cells(9, 11))
Set tableRange = ws.Range(ws.Cells(10, 2), ws.Cells(ws.Cells(ws.Rows.Count, "B").End(xlUp).row, 11))
Set summaryHeaders = ws.Range(ws.Cells(9, 13), ws.Cells(9, ws.Cells(9, ws.Columns.Count).End(xlToLeft).Column))
Set subTotals_Obligation = ws.Range(ws.Cells(14, 13), ws.Cells(14, ws.Cells(14, ws.Columns.Count).End(xlToLeft).Column))
Set subTotals_PV = ws.Range(ws.Cells(17, 13), ws.Cells(17, ws.Cells(17, ws.Columns.Count).End(xlToLeft).Column))
Set subTotals_NPV = ws.Range(ws.Cells(22, 13), ws.Cells(22, ws.Cells(22, ws.Columns.Count).End(xlToLeft).Column))


    With tableHeaders
        .Font.Bold = True
        .HorizontalAlignment = xlCenter ' Center align the text
        .VerticalAlignment = xlCenter ' Center align vertically
        '.WrapText = True ' Enable text wrapping
    End With
    
    With tableHeaders.Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0) ' Black border color
        .TintAndShade = 0 ' No shading
        .Weight = xlThin ' Set border thickness
    End With
    With tableHeaders.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0) ' Black border color
        .TintAndShade = 0 ' No shading
        .Weight = xlMedium ' Set border thickness
    End With
    
    
       
    ' Apply thin, continuous vertical borders
    With tableRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0) ' Black border color
        .TintAndShade = 0 ' No shading
        .Weight = xlThin ' Set border thickness
    End With
    
    With tableRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0) ' Black border color
        .TintAndShade = 0 ' No shading
        .Weight = xlThin ' Set border thickness
    End With
    
    With tableRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0) ' Black border color
        .TintAndShade = 0 ' No shading
        .Weight = xlThin ' Set border thickness
    End With

    ' Apply dotted/dashed inside borders (horizontal and vertical) within the table range
    With tableRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous ' Use a dotted line for vertical inside borders
        .Color = RGB(0, 0, 0) ' Black color
        .Weight = xlThin ' Set border thickness
    End With
    
    With tableRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous ' Use a dotted line for horizontal inside borders
        .Color = RGB(0, 0, 0) ' Black color
        .Weight = xlHairline ' Set border thickness
    End With
    
    With summaryHeaders
            .Font.Bold = True
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlMedium
            
    End With
    
    
    With subTotals_Obligation
            .Font.Bold = True
            .Borders(xlEdgeTop).Weight = xlThin
            '.Borders(xlEdgeBottom).Weight = xlMedium
            
    End With
    
    With subTotals_PV
        .Font.Bold = True
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    
    With subTotals_NPV
        .Font.Bold = True
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    'Additional Information Placeholders
    ws.Cells(2, 8).Value = "Entity"
    ws.Cells(3, 8).Value = "Asset Description"
    ws.Cells(4, 8).Value = "Financier"
    ws.Cells(6, 8).Value = "Baloon/Residual"
    
    
    
    tableRange.EntireColumn.AutoFit
End Function


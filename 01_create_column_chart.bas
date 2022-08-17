Option Explicit

Sub create_diagram_sheet()

Dim rngResult As Range
Dim wksSheet As Worksheet
Set wksSheet = Sheet1
    
Call create_data(wksSheet)

Set rngResult = wksSheet.UsedRange

ThisWorkbook.Charts.Add after:=wksSheet

With ActiveChart
  .ChartType = xlColumnClustered
  .SetSourceData Source:=rngResult
  .Name = "Column_Chart_1"
End With

End Sub

' ===============================================================
Function create_data(ByRef wksSheet As Worksheet)

Dim intCounter As Integer
Dim varRevenue As Variant
Dim rngResult As Range

varRevenue = Array(10, 12, 15, 13, 9, 10, 14)

For intCounter = LBound(varRevenue) To UBound(varRevenue)
  Sheet1.Cells(1, intCounter + 1).Value = varRevenue(intCounter)
Next intCounter

End Function

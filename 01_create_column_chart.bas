Option Explicit

' ===============================================================
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
Dim varDate, varRevenue As Variant
Dim rngResult As Range

varRevenue = Array(1000, 1200, 1500, 1300, 900, 1000, 1400)
varDate = Array(DateSerial(2022, 8, 7), DateSerial(2022, 8, 8), _
  DateSerial(2022, 8, 9), DateSerial(2022, 8, 10), DateSerial(2022, 8, 11), _
  DateSerial(2022, 8, 12), DateSerial(2022, 8, 13))
  
For intCounter = LBound(varRevenue) To UBound(varRevenue)
  Sheet1.Cells(intCounter + 1, 2).Value = CCur(varRevenue(intCounter))
  Sheet1.Cells(intCounter + 1, 1).Value = varDate(intCounter)
Next intCounter

End Function

Option Explicit

' Create and rename a WorkSheet to Sheet2(Sheet2)
' ==========================================================================
Sub create_embedded_diagram()

Dim chart_object As ChartObject
Dim chart As chart

Call create_data

Set chart_object = Sheet2.ChartObjects.Add(200, 10, 300, 150)

Set chart = chart_object.chart
chart.ChartType = xlLine
chart.SetSourceData Sheet2.Range("A1:C8")

Set chart = Nothing
Set chart_object = Nothing

End Sub

' ==========================================================================
Function create_data()

Dim intCounter, intCellFree As Integer
Dim varDate, varTempMin, varTempMax As Variant

varTempMin = Array(13, 16, 16, 14, 13, 12, 12)
varTempMax = Array(33, 28, 21, 22, 24, 23, 23)
varDate = Array(DateSerial(2022, 8, 7), DateSerial(2022, 8, 8), _
  DateSerial(2022, 8, 9), DateSerial(2022, 8, 10), DateSerial(2022, 8, 11), _
  DateSerial(2022, 8, 12), DateSerial(2022, 8, 13))
  
Sheet2.Cells(1, 1).Value = "Date"
Sheet2.Cells(1, 2).Value = "min_temp"
Sheet2.Cells(1, 3).Value = "max_temp"
intCellFree = Sheet2.Cells(Sheet2.Rows.Count, 1).End(xlUp).Row + 1

For intCounter = LBound(varDate) To UBound(varDate)
  Sheet2.Cells(intCellFree, 1).Value = varDate(intCounter)
  Sheet2.Cells(intCellFree, 2).Value = varTempMin(intCounter)
  Sheet2.Cells(intCellFree, 3).Value = varTempMax(intCounter)
  intCellFree = intCellFree + 1
Next intCounter

End Function

Attribute VB_Name = "PivotChart"
Option Compare Database

Public Sub BuildPivotChart()
  Dim objPivotChart As OWC10.ChChart
  Dim objChartSpace As OWC10.ChartSpace
  Dim frm As Access.Form
  Dim strExpression As String
  Dim rs As Recordset
  Dim values
  Dim axCategoryAxis
  Dim axValueAxis

  'Open the form in PivotChart view.
  DoCmd.OpenForm "frmPivotChart", acFormPivotChart
  Set frm = Forms("frmPivotChart")
  Set rs = frm.Recordset
  
  'Loop through Recordset to obtain data for the chart and put in strings.
  rs.MoveFirst
    Do While Not rs.EOF
        strExpression = strExpression & rs.Fields(0).Value & Chr(9)
        values = values & rs.Fields(1).Value & Chr(9)
        rs.MoveNext
    Loop
  rs.Close
  Set rs = Nothing
  
  'Trim any extra tabs from string.
  strExpression = Left(strExpression, Len(strExpression) - 1)
  values = Left(values, Len(values) - 1)
     
  'Clear existing Charts on Form if present and add a new chart to the form.
  'Set object variable equal to the new chart.
  Set objChartSpace = frm.ChartSpace
  objChartSpace.Clear
  objChartSpace.Charts.Add
  Set objPivotChart = objChartSpace.Charts.Item(0)
  
  'Set a variable to the Category (X) axis.
  Set axCategoryAxis = objChartSpace.Charts(0).Axes(0)
    
  ' Set a variable to the Value (Y) axis.
  Set axValueAxis = objChartSpace.Charts(0).Axes(1)

  ' The following two lines of code enable, and then
  ' set the title for the category axis.
  axCategoryAxis.HasTitle = True
  axCategoryAxis.Title.Caption = "Employees"
    
  ' The following two lines of code enable, and then
  ' set the title for the value axis.
  axValueAxis.HasTitle = True
  axValueAxis.Title.Caption = "Orders"
    
  'Add Series to Chart and set the caption.
  objPivotChart.SeriesCollection.Add
  objPivotChart.SeriesCollection(0).Caption = "Orders"
  
  'Add Data to the Series.
  objPivotChart.SeriesCollection(0).SetData chDimCategories, chDataLiteral, strExpression
  objPivotChart.SeriesCollection(0).SetData chDimValues, chDataLiteral, values
  
  'Set focus to the form and destroy the form object from memory.
  frm.SetFocus
  Set frm = Nothing
  
End Sub


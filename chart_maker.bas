Option Compare Database
Option Explicit
  
'adding globalization vars for charts
Public chCategoryAxisTitle As String
Public chValueAxisTitle As String
Public chSeriesTitle As String
Public chChartTitle As String

Function BuildPivotChart()
  Dim objPivotChart As OWC10.ChChart
  Dim objChartSpace As OWC10.ChartSpace
  Dim frm As Access.Form
  Dim strExpression As String
  Dim rs As Recordset
  Dim values
  Dim axCategoryAxis
  Dim axValueAxis
  Dim chTitle

  'Open the form in PivotChart view *better way is to preset the form to default on PivotChart view
  'DoCmd.OpenForm "sales_chart", acFormPivotChart
  'Call format_viewer("sales_chart", 0, 1)
  'Set frm = Forms("sales_chart")
  Set frm = Forms(objForm).Controls(objControl).Form.Controls(objSubformControl).Form
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
  objChartSpace.Charts(0).Type = chChartTypeSmoothLineStackedMarkers
  Set objPivotChart = objChartSpace.Charts.Item(0)
  
  'Set a variable to the Category (X) axis.
  Set axCategoryAxis = objChartSpace.Charts(0).Axes(0)
    
  'Set a variable to the Value (Y) axis.
  Set axValueAxis = objChartSpace.Charts(0).Axes(1)

  'Adding variable to the chart title
  'Set chTitle = objChartSpace.Charts(0).Title.Caption = chChartTitle
  Set chTitle = objChartSpace.Charts(0)
  
  'enabling chart title
  chTitle.HasTitle = True
  chTitle.Title.Caption = chChartTitle
    
  ' The following two lines of code enable, and then
  ' set the title for the category axis.
  axCategoryAxis.HasTitle = True
  axCategoryAxis.Title.Caption = chCategoryAxisTitle
    
  ' The following two lines of code enable, and then
  ' set the title for the value axis.
  axValueAxis.HasTitle = True
  axValueAxis.Title.Caption = chValueAxisTitle
    
  'Add Series to Chart and set the caption.
  objPivotChart.SeriesCollection.Add
  objPivotChart.SeriesCollection(0).Caption = chSeriesTitle
  
  'Add Data to the Series.
  objPivotChart.SeriesCollection(0).SetData chDimCategories, chDataLiteral, strExpression
  objPivotChart.SeriesCollection(0).SetData chDimValues, chDataLiteral, values
  
  'Set focus to the form and destroy the form object from memory.
  'frm.SetFocus
  Set frm = Nothing
  
End Function

Function set_chart_labels(myChartType As String)

  Select Case myChartType
    Case "sales"
      chCategoryAxisTitle = "Date"
      chValueAxisTitle = "Sales Value (in Php)"
      chSeriesTitle = "Sales"
    Case "hourly"
      chCategoryAxisTitle = "Hour of Day"
      chValueAxisTitle = "Sales Value (in Php)"
      chSeriesTitle = "Sales"
    Case "item_count"
      chCategoryAxisTitle = "Hour of Day"
      chValueAxisTitle = "No. of Items"
      chSeriesTitle = "Sales"
  End Select

End Function
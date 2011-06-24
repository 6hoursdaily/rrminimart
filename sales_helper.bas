Option Compare Database
Option Explicit

Function BuildSalesChart()
  
  objSubformControl = "chart"
  
  Call reset_sales_subforms
  Call format_viewer("sales_chart", 0, 1)
  Call BuildPivotChart
  
  'hide other subforms
  'Forms(objForm).Controls(objControl).Form.Controls("focustaker").SetFocus
  'Forms(objForm).Controls(objControl).Form.Controls("viewer").Visible = False
  'Forms(objForm).Controls(objControl).Form.Controls("info").Visible = False

End Function

Function reset_sales_subforms()

  Forms(objForm).Controls(objControl).Form.Controls("focustaker").SetFocus
  Forms(objForm).Controls(objControl).Form.Controls("viewer").Visible = False
  Forms(objForm).Controls(objControl).Form.Controls("info").Visible = False
  Forms(objForm).Controls(objControl).Form.Controls("chart").Visible = False
  Forms(objForm).Controls(objControl).Form.Controls("viewer").SourceObject = ""
  Forms(objForm).Controls(objControl).Form.Controls("info").SourceObject = ""
  Forms(objForm).Controls(objControl).Form.Controls("chart").SourceObject = ""
  
  Forms(objForm).Controls(objControl).Form.Controls("focustaker").SetFocus
  
End Function
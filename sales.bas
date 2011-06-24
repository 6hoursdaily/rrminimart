VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Dim strNoOfDays As Integer

Private Sub Form_Load()
    
  'hide links
  Call show_chart_controls(False)
    
  'reset subform objects
  Call reset_sales_subforms
  
  strSQLWhere = "WHERE Status = 'REG'"

End Sub

Private Sub img_sales_monitor_Click()
    
  Call reset_sales_subforms
  
  objSubformControl = "viewer"
  Call format_viewer("sales_viewer", 1, 1)
  
  objSubformControl = "info"
  Call format_viewer("sales_summary", 1, 1)
    
  Me.focustaker.SetFocus
  
End Sub

Private Sub img_sales_trends_Click()
  
  Call EndTimer
  Call set_chart_labels("sales")
  'Call Subforms_reset
  
  chChartTitle = "14-Day Sales Trending"
  
  strSQLRank = "TOP 14"
  strSQLChart = "SELECT " & strSQLRank & " DtlsDate, SUM(ExtPriceEff) AS totalSales FROM SalesDtls " & strSQLWhere & " GROUP BY DtlsDate ORDER BY DtlsDate DESC"
  
  Call BuildSalesChart
  
  Call show_chart_controls(True)
  
End Sub
Private Sub btn_this_month_Click()
  
  Call EndTimer
  Call set_chart_labels("sales")
  
  chChartTitle = "30-Day Sales Trending"
  
  strSQLRank = "TOP 30"
  strSQLChart = "SELECT " & strSQLRank & " DtlsDate, SUM(ExtPriceEff) AS totalSales FROM SalesDtls " & strSQLWhere & " GROUP BY DtlsDate ORDER BY DtlsDate DESC"
  
  Call BuildSalesChart
  
End Sub

Private Sub btn_this_quarter_Click()
  
  Call EndTimer
  Call set_chart_labels("sales")
  
  chChartTitle = "90-Day Sales Trending"
  
  strSQLRank = "TOP 90"
  strSQLChart = "SELECT " & strSQLRank & " DtlsDate, SUM(ExtPriceEff) AS totalSales FROM SalesDtls " & strSQLWhere & " GROUP BY DtlsDate ORDER BY DtlsDate DESC"
  
  Call BuildSalesChart

End Sub

Private Sub btn_this_week_Click()
  
  Call EndTimer
  Call set_chart_labels("sales")
  
  chChartTitle = "14-Day Sales Trending"
  
  strSQLRank = "TOP 14"
  strSQLChart = "SELECT " & strSQLRank & " DtlsDate, SUM(ExtPriceEff) AS totalSales FROM SalesDtls " & strSQLWhere & " GROUP BY DtlsDate ORDER BY DtlsDate DESC"
  
  Call BuildSalesChart

End Sub

Private Sub btn_hourly_14_Click()
  
  Call EndTimer
  Call set_chart_labels("hourly")
  
  strNoOfDays = 14
  chChartTitle = "Ave. Sales per Hour Over Past " & strNoOfDays & " Days"
  
  strSQLChart = "SELECT FORMAT(hour) AS hour_of_day, FORMAT(Avg(hourly_sales),'Standard') AS avg_hourly_sales FROM hourly_data_query WHERE DtlsDate BETWEEN NOW() AND NOW()-" & strNoOfDays & " GROUP BY hour" 'DateValue('6/1/2011') AND DateValue('6/30/2011') GROUP BY hour"

  Call BuildSalesChart

End Sub

Private Sub btn_hourly_30_Click()
  
  Call EndTimer
  Call set_chart_labels("hourly")
  
  strNoOfDays = 30
  chChartTitle = "Ave. Sales per Hour Over Past " & strNoOfDays & " Days"
    
  strSQLChart = "SELECT FORMAT(hour) AS hour_of_day, FORMAT(Avg(hourly_sales),'Standard') AS avg_hourly_sales FROM hourly_data_query WHERE DtlsDate BETWEEN NOW() AND NOW()-" & strNoOfDays & " GROUP BY hour" 'DateValue('6/1/2011') AND DateValue('6/30/2011') GROUP BY hour"

  Call BuildSalesChart

End Sub

Private Sub btn_hourly_sum_14_Click()
  
  Call EndTimer
  Call set_chart_labels("hourly")
  
  strNoOfDays = 14
  chChartTitle = "Aggregate Sales per Hour Over Past " & strNoOfDays & " Days"
    
  strSQLChart = "SELECT hour, sum(hourly_sales) as sum_hourly_sales FROM hourly_data_query WHERE DtlsDate BETWEEN NOW() AND NOW()-" & strNoOfDays & " GROUP BY hour"
  
  Call BuildSalesChart

End Sub

Private Sub btn_hourly_sum_30_Click()
  
  Call EndTimer
  Call set_chart_labels("hourly")
  
  strNoOfDays = 30
  chChartTitle = "Aggregate Sales per Hour Over Past " & strNoOfDays & " Days"
    
  strSQLChart = "SELECT hour, sum(hourly_sales) as sum_hourly_sales FROM hourly_data_query WHERE DtlsDate BETWEEN NOW() AND NOW()-" & strNoOfDays & " GROUP BY hour"
  
  Call BuildSalesChart

End Sub

Private Sub btn_hourly_transactions_14_Click()
  
  Call EndTimer
  Call set_chart_labels("item_count")
  
  strNoOfDays = 14
  chChartTitle = "Aggregate Items Sold per Hour Over Past " & strNoOfDays & " Days"
    
  strSQLChart = "SELECT TIMESERIAL(FORMAT(SalesDtls.EndTime,'HH'),0,0) AS [hour], COUNT(ExtPriceEff) AS trans_count FROM SalesDtls WHERE DtlsDate BETWEEN NOW() AND NOW()-" & strNoOfDays & " AND Status = 'REG' GROUP BY TIMESERIAL(FORMAT(SalesDtls.EndTime,'HH'),0,0)"
  
  Call BuildSalesChart

End Sub

Private Sub btn_hourly_transactions_30_Click()
  
  Call EndTimer
  Call set_chart_labels("hourly")
  
  strNoOfDays = 30
  chChartTitle = "Aggregate Items Sold per Hour Over Past " & strNoOfDays & " Days"
    
  strSQLChart = "SELECT TIMESERIAL(FORMAT(SalesDtls.EndTime,'HH'),0,0) AS [hour], COUNT(ExtPriceEff) AS trans_count FROM SalesDtls WHERE DtlsDate BETWEEN NOW() AND NOW()-" & strNoOfDays & " AND Status = 'REG' GROUP BY TIMESERIAL(FORMAT(SalesDtls.EndTime,'HH'),0,0)"
  
  Call BuildSalesChart

End Sub

Function show_chart_controls(thisVisibility As Boolean)

  'sales trending controls
  Me.a_this_week.Visible = thisVisibility
  Me.a_this_month.Visible = thisVisibility
  Me.a_this_quarter.Visible = thisVisibility
  Me.btn_this_month.Visible = thisVisibility
  Me.btn_this_quarter.Visible = thisVisibility
  Me.btn_this_week.Visible = thisVisibility
  
  'hourly trending controls
  Me.a_hourly_14.Visible = thisVisibility
  Me.a_hourly_30.Visible = thisVisibility
  Me.a_hourly_sum_14.Visible = thisVisibility
  Me.a_hourly_sum_30.Visible = thisVisibility
  Me.a_hourly_transactions_14.Visible = thisVisibility
  Me.a_hourly_transactions_30.Visible = thisVisibility
  
  Me.btn_hourly_14.Visible = thisVisibility
  Me.btn_hourly_30.Visible = thisVisibility
  Me.btn_hourly_sum_14.Visible = thisVisibility
  Me.btn_hourly_sum_30.Visible = thisVisibility
  Me.btn_hourly_transactions_14.Visible = thisVisibility
  Me.btn_hourly_transactions_30.Visible = thisVisibility
  
  'labels
  Me.label_hourly.Visible = thisVisibility
  Me.label_avg.Visible = thisVisibility
  Me.label_sum.Visible = thisVisibility
  Me.label_transactions.Visible = thisVisibility
  
End Function



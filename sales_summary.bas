VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sales_summary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
  
  strSQLSearch = "SELECT SUM(Qty) AS sumQty, SUM(ExtPriceEff) AS sumSales FROM SalesDtls WHERE Status = 'REG'"
  strGroupBy = " GROUP BY DtlsDate"
  Me.RecordSource = strSQLSearch & strGroupBy
  
End Sub

Private Sub a_daily_Click()
  
  strGroupBy = " GROUP BY DtlsDate"
  Call myNavcolors("daily")
  
End Sub

Private Sub a_weekly_Click()
  
  strGroupBy = " GROUP BY DatePart('ww',DtlsDate)"
  Call myNavcolors("weekly")

End Sub

Private Sub a_monthly_Click()
  
  strGroupBy = " GROUP BY DatePart('m',DtlsDate)"
  Call myNavcolors("monthly")

End Sub

Function myNavcolors(navitem As String)

  Me.a_daily.ForeColor = 1279872587
  Me.a_weekly.ForeColor = 1279872587
  Me.a_monthly.ForeColor = 1279872587
  
  Me.label_average_sales.Caption = "Ave. " & navitem & " sales:"
  navitem = "a_" & navitem
  Me.Controls(navitem).ForeColor = 255
  
  Me.RecordSource = strSQLSearch & strGroupBy

End Function

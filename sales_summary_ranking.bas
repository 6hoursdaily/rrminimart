VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sales_summary_ranking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim thisGroupBy As String
Dim thisOrderBy As String

Private Sub Form_Load()
  
  strSQL = "SELECT TOP 10 ProdName, Sum(Qty) AS sumQty, Sum([ExtPriceEff]-[UPrice]) AS Net FROM SalesDtls WHERE Status = 'REG'"
  thisGroupBy = " GROUP BY ProdName"
  thisOrderBy = " ORDER BY Sum(Qty) DESC , Sum([ExtPriceEff]-[UPrice]) DESC"
  Me.RecordSource = strSQL & thisGroupBy & thisOrderBy
  
End Sub

Private Sub a_all_Click()
  
  strSQLWhere = ""
  Call myNavcolors("a_all")


End Sub

Private Sub a_week_Click()
  
  strSQLWhere = "  AND DtlsDate > NOW() - 7"
  Call myNavcolors("a_week")

End Sub

Private Sub a_month_Click()
  
  strSQLWhere = " AND DtlsDate Like  MONTH(Now()) & '/*'"
  Call myNavcolors("a_month")

End Sub

Function myNavcolors(navitem As String)
  
  strSQL = "SELECT TOP 10 ProdName, Sum(Qty) AS sumQty, Sum([ExtPriceEff]-[UPrice]) AS Net FROM SalesDtls WHERE Status = 'REG'"

  Me.a_all.ForeColor = 1279872587
  Me.a_week.ForeColor = 1279872587
  Me.a_month.ForeColor = 1279872587
  
  Me.Controls(navitem).ForeColor = 255
  
  Me.RecordSource = strSQL & strSQLWhere & thisGroupBy & thisOrderBy

End Function

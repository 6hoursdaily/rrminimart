VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sales_viewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()

  timerForm = "viewer"
  
  strSQLSearch = "SELECT DISTINCT DtlsDate, SUM(Qty) AS totalQty, SUM(ExtPriceEff) As totalSales FROM SalesDtls WHERE Status = 'REG' GROUP BY DtlsDate ORDER BY DtlsDate DESC"
  Me.RecordSource = strSQLSearch
  
  'uncomment upon production
  Call StartTimer 'init timer to refresh the sales data
  
  Me.a_view.Hyperlink.Address = "#"
  
End Sub
Private Sub Form_Close()
    
  Call EndTimer
    
End Sub

Private Sub btn_view_Click()
  
  objSubformControl = "viewer"
  strSQL = Me.DtlsDate
  subformwidth = 7920 + 250 '250 is for the scrollbar
  
  Call EndTimer
  Call format_viewer("sales_reports", 1, 1)
  
  Forms(objForm).Controls(objControl).Form.Controls("focustaker").SetFocus
  Forms(objForm).Controls(objControl).Form.Controls("info").Visible = False

End Sub

Attribute VB_Name = "Global"
Global LastPrice As String
Public Sub StockUpdate(Title, Caption, Optional Good = 0)
frmStocks.Show
frmStocks.Image1.Visible = True
frmStocks.Label1 = "Stock Quote Update : " & Title
frmStocks.Label2 = Caption
End Sub

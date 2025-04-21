Sub UpdateAllTablesAndProfit()
    Call UpdateExpenseTable
    Call UpdateSalesTable
    Call UpdateProfitStatement
    MsgBox "すべてのシートが更新されました！", vbInformation
End Sub


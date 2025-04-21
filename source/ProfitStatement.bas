Sub UpdateProfitStatement()
    Dim ws As Worksheet
    Dim tblExp As ListObject
    Dim tblSales As ListObject
    Dim salesExTax As Double, salesTax As Double
    Dim expExTax As Double, expTax As Double
    Dim profit As Double, taxDue As Double, netProfit As Double
    Dim i As Long
    Dim labelList As Variant
    Dim valueDict As Object
    Set valueDict = CreateObject("Scripting.Dictionary")

    ' ▼ テーブル取得
    Set tblExp = ThisWorkbook.Sheets("経費管理").ListObjects("ExpenseTable")
    Set tblSales = ThisWorkbook.Sheets("売上表").ListObjects("SalesTable")

    ' ▼ 合計取得
    salesExTax = tblSales.ListColumns("税抜金額").Total
    salesTax = tblSales.ListColumns("消費税額").Total
    expExTax = tblExp.ListColumns("税抜金額").Total
    expTax = tblExp.ListColumns("消費税額").Total

    ' ▼ 計算
    profit = salesExTax - expExTax
    taxDue = salesTax - expTax
    If taxDue < 0 Then taxDue = 0
    netProfit = profit - taxDue

    ' ▼ 値格納
    valueDict("売上高（税抜）") = salesExTax
    valueDict("売上消費税") = salesTax
    valueDict("経費合計（税抜）") = expExTax
    valueDict("経費消費税") = expTax
    valueDict("営業利益（税抜）") = profit
    valueDict("納税予定額（消費税）") = taxDue
    valueDict("純利益（税引後）") = netProfit

    ' ▼ 項目順
    labelList = Array("売上高（税抜）", "売上消費税", "経費合計（税抜）", "経費消費税", "営業利益（税抜）", "納税予定額（消費税）", "純利益（税引後）")

    ' ▼ シート指定
    Set ws = ThisWorkbook.Sheets("損益計算書")

    ' ▼ 項目列が空なら作成
    If ws.Cells(1, 1).Value = "" Then
        For i = 0 To UBound(labelList)
            ws.Cells(i + 1, 1).Value = labelList(i)
        Next i
    End If

    ' ▼ 値の反映（B列）
    For i = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim label As String
        label = ws.Cells(i, 1).Value

        If valueDict.exists(label) Then
            ws.Cells(i, 2).Value = valueDict(label)
            ws.Cells(i, 2).NumberFormat = """\""#,##0"
        End If
    Next i

    ' ▼ テーブル見た目調整（枠線、太字、背景色）
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    With ws.Range("A1:B" & lastRow)
        .Borders.LineStyle = xlContinuous
        .Font.Name = "メイリオ"
        .Font.Size = 10
        .Columns.AutoFit
    End With

    ' ▼ ヘッダー行の装飾
    With ws.Range("A1:B1")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 240)
    End With

    ' ▼ 純利益行を強調
    With ws.Range("A" & lastRow & ":B" & lastRow)
        .Font.Bold = True
        .Interior.Color = RGB(220, 255, 220)
    End With
    
    
    ws.Columns.AutoFit
    MsgBox "損益計算書が更新されました！", vbInformation
End Sub


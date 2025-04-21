Sub UpdateExpenseTable()
    Dim folderPath As String
    Dim fileSystem As Object
    Dim file As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblRow As ListRow
    Dim fileNameParts As Variant
    Dim fileDate As Date
    Dim itemName As String
    Dim amount As Double
    Dim remarks As String
    Dim hyperlinkAddress As String
    Dim taxRateStr As String
    Dim taxRate As Double
    Dim taxAmount As Double
    Dim amountExTax As Double
    Dim taxRateDisplay As String
    Dim existingFileNames As Collection
    Dim i As Long

    ' ▼動的な年を取得
    Dim currentYear As String
    currentYear = Year(Date)

    ' ▼経費フォルダのパス
    folderPath = "あなたのファイルパスをここに貼り付けてください（currentYearを用いて現在年月自動取得可能）"
    ' 例 "G:\マイドライブ\領収書管理\" & currentYear & "年\経費\"

    ' ▼ワークシート取得
    Set ws = ThisWorkbook.Sheets("経費管理")

    ' ▼ヘッダーが無ければ自動挿入（1行目に）
    If ws.Cells(1, 1).Value = "" Then
        Dim headers As Variant
        headers = Array("日付", "勘定科目", "金額", "摘要", "リンク", "消費税率", "税抜金額", "消費税額")
        For i = 0 To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
        Next i
    End If

    ' ▼テーブルが未作成なら作成
    If ws.ListObjects.Count = 0 Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
        tbl.Name = "ExpenseTable"
    Else
        Set tbl = ws.ListObjects(1)
    End If

    ' ▼既存ファイル名を収集（重複防止）
    Set existingFileNames = New Collection
    On Error Resume Next
    For i = 1 To tbl.ListRows.Count
        Dim cellValue As String
        cellValue = tbl.ListRows(i).Range(1, 5).Text
        If Len(cellValue) > 0 Then
            existingFileNames.Add cellValue, CStr(cellValue)
        End If
    Next i
    On Error GoTo 0

    ' ▼ファイル取得開始
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(folderPath) Then
        MsgBox "フォルダが見つかりません: " & folderPath, vbExclamation
        Exit Sub
    End If

    For Each file In fileSystem.GetFolder(folderPath).Files
        If InStr(file.Name, ".pdf") > 0 Then
            Dim fileName As String
            fileName = file.Name
            hyperlinkAddress = folderPath & fileName

            Dim fileNameExists As Boolean
            fileNameExists = False
            On Error Resume Next
            fileNameExists = Not IsEmpty(existingFileNames(fileName))
            On Error GoTo 0

            If Not fileNameExists Then
                fileNameParts = Split(file.Name, "_")
                If UBound(fileNameParts) >= 6 Then
                    fileDate = CDate(fileNameParts(0) & "/" & fileNameParts(1) & "/" & fileNameParts(2))
                    itemName = fileNameParts(3)
                    amount = Val(fileNameParts(4))
                    remarks = fileNameParts(5)

                    ' ▼消費税率抽出（例："10%.pdf" → "10"）
                    taxRateStr = Replace(Split(fileNameParts(6), ".")(0), "%", "")
                    If IsNumeric(taxRateStr) Then
                        taxRate = CDbl(taxRateStr) / 100
                    Else
                        taxRate = 0
                    End If
                    taxRateDisplay = Format(taxRate * 100, "0") & "%"

                    ' ▼税抜・税額計算
                    amountExTax = Round(amount / (1 + taxRate), 0)
                    taxAmount = amount - amountExTax

                    ' ▼テーブルに行追加
                    Set tblRow = tbl.ListRows.Add
                    With tblRow
                        .Range(1, 1).Value = fileDate
                        .Range(1, 2).Value = itemName
                        .Range(1, 3).Value = amount
                        .Range(1, 4).Value = remarks
                        ws.Hyperlinks.Add Anchor:=.Range(1, 5), Address:=hyperlinkAddress, TextToDisplay:=fileName
                        .Range(1, 6).Value = taxRateDisplay
                        .Range(1, 7).Value = amountExTax
                        .Range(1, 8).Value = taxAmount
                    End With
                End If
            End If
        End If
    Next file

    ' ▼合計行の更新
    On Error Resume Next
    tbl.ShowTotals = False
    tbl.ShowTotals = True
    On Error GoTo 0

    ' ▼合計列の指定
    With tbl
        .ListColumns("金額").TotalsCalculation = xlTotalsCalculationSum
        .ListColumns("税抜金額").TotalsCalculation = xlTotalsCalculationSum
        .ListColumns("消費税額").TotalsCalculation = xlTotalsCalculationSum
    End With

    ' ▼合計行のフォーマット
    With tbl.TotalsRowRange
        .Columns(3).NumberFormat = """\""#,##0"
        .Columns(7).NumberFormat = """\""#,##0"
        .Columns(8).NumberFormat = """\""#,##0"
    End With

    ' ▼日付で昇順ソート
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=tbl.ListColumns(1).Range, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

    ' ▼列フォーマット
    ws.Columns(3).NumberFormat = """\""#,##0"
    ws.Columns(7).NumberFormat = """\""#,##0"
    ws.Columns(8).NumberFormat = """\""#,##0"
    
    ws.Columns.AutoFit

    MsgBox "経費シートが更新されました！", vbInformation
End Sub

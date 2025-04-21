Sub UpdateSalesTable()
    Dim folderPath As String
    Dim fileSystem As Object
    Dim file As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fileNameParts As Variant
    Dim fileDate As Date
    Dim amount As Double
    Dim clientName As String
    Dim hyperlinkAddress As String
    Dim amountExTax As Double
    Dim taxAmount As Double
    Dim existingFileNames As Collection
    Dim i As Long

    ' ▼動的な年を取得
    Dim currentYear As String
    currentYear = Year(Date)

    ' ▼請求書フォルダのパス
    folderPath = "G:\マイドライブ\領収書管理\" & currentYear & "年\請求書\"

    ' ▼売上シート取得
    Set ws = ThisWorkbook.Sheets("売上表")

    ' ▼ヘッダーが無ければ自動挿入
    If ws.Cells(1, 1).Value = "" Then
        Dim headers As Variant
        headers = Array("日付", "顧客名", "金額（税込）", "税抜金額", "消費税額", "リンク")
        For i = 0 To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
        Next i
    End If

    ' ▼テーブルが未作成なら作成
    If ws.ListObjects.Count = 0 Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
        tbl.Name = "SalesTable"
    Else
        Set tbl = ws.ListObjects(1)
    End If

    ' ▼既存ファイル名の取得（重複防止）
    Set existingFileNames = New Collection
    On Error Resume Next
    For i = 1 To tbl.ListRows.Count
        Dim linkText As String
        linkText = tbl.ListRows(i).Range(1, 6).Text
        If Len(linkText) > 0 Then
            existingFileNames.Add linkText, CStr(linkText)
        End If
    Next i
    On Error GoTo 0

    ' ▼フォルダ存在確認
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(folderPath) Then
        MsgBox "フォルダが見つかりません: " & folderPath, vbExclamation
        Exit Sub
    End If

    ' ▼ファイル走査
    For Each file In fileSystem.GetFolder(folderPath).Files
        If InStr(file.Name, ".pdf") > 0 Then
            Dim fileName As String
            fileName = file.Name
            hyperlinkAddress = folderPath & fileName

            ' ▼重複チェック（ファイル名ベース）
            Dim fileNameExists As Boolean
            fileNameExists = False
            On Error Resume Next
            fileNameExists = Not IsEmpty(existingFileNames(fileName))
            On Error GoTo 0

            If Not fileNameExists Then
                fileNameParts = Split(file.Name, "_")
                If UBound(fileNameParts) >= 3 Then
                    fileDate = CDate(fileNameParts(0) & "/" & fileNameParts(1) & "/" & fileNameParts(2))
                    amount = Val(fileNameParts(3))
                    clientName = Replace(Split(file.Name, "_", 5)(4), ".pdf", "")

                    ' ▼税率一律10%
                    amountExTax = Round(amount / 1.1, 0)
                    taxAmount = amount - amountExTax

                    ' ▼テーブルに追加
                    Dim tblRow As ListRow
                    Set tblRow = tbl.ListRows.Add
                    With tblRow
                        .Range(1, 1).Value = fileDate
                        .Range(1, 2).Value = clientName & "様"
                        .Range(1, 3).Value = amount
                        .Range(1, 4).Value = amountExTax
                        .Range(1, 5).Value = taxAmount
                        ws.Hyperlinks.Add Anchor:=.Range(1, 6), Address:=hyperlinkAddress, TextToDisplay:=fileName
                    End With
                End If
            End If
        End If
    Next file

    ' ▼合計行の設定
    On Error Resume Next
    tbl.ShowTotals = False
    tbl.ShowTotals = True
    On Error GoTo 0

    With tbl
        .ListColumns("金額（税込）").TotalsCalculation = xlTotalsCalculationSum
        .ListColumns("税抜金額").TotalsCalculation = xlTotalsCalculationSum
        .ListColumns("消費税額").TotalsCalculation = xlTotalsCalculationSum
        .ListColumns("リンク").TotalsCalculation = xlTotalsCalculationNone
    End With

    ' ▼合計行フォーマット
    With tbl.TotalsRowRange
        .Columns(3).NumberFormat = """\""#,##0"
        .Columns(4).NumberFormat = """\""#,##0"
        .Columns(5).NumberFormat = """\""#,##0"
    End With

    ' ▼列フォーマット
    ws.Columns(3).NumberFormat = """\""#,##0"
    ws.Columns(4).NumberFormat = """\""#,##0"
    ws.Columns(5).NumberFormat = """\""#,##0"

    ' ▼日付で昇順
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=tbl.ListColumns(1).Range, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

    ws.Columns.AutoFit
    MsgBox "? 売上シートが更新されました！（ファイル名ベース重複チェック）", vbInformation
End Sub


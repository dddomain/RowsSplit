Option Explicit

Sub RowsSplit()

Dim headerRange As Range
Set headerRange = Application.InputBox _
    ( _
    prompt:="ヘッダー範囲を選択してください", _
    Title:="ヘッダー範囲選択", _
    Type:=8 _
    )

'エラー処理
If Err.Number <> 0 Then
    MsgBox "キャンセルされました。"
    Exit Sub
End If

'ヘッダーの最初と最後の行及びテーブルの最後の列を取得する
Dim firstHeaderRow As Long
Dim lastHeaderRow As Long
Dim lastCol As Long
firstheadeerrow = headerRange.row
lastHeaderRow = headerRange.Rows(headerRange.Rows.Count).row
lastCol = headerRange.CurrentRegion.Columns.Count
'テーブルの最後の行を取得する
Dim lastDataRow As Long
lastDataRow = Cells(Rows.Count, 1).End(xlUp).row

'ヘッダー範囲から最初のデータ行を取得する
Dim firstDataRow As Long
firstDataRow = lastHeaderRow + 1
firstDataRow = Application.InputBox _
    ( _
    prompt:="データ範囲は" & firstDataRow & "行目からでよろしいですか？", _
    Title:="データの開始行確認", _
    Default:=firstDataRow, _
    Type:=2 _
    )

'処理の実行
Dim execRow As Long: execRow = firstDataRow
Dim i As Long
For i = firstDataRow To lastDataRow
    Dim createdWorkbook As Workbook
    Set createdWorkbook = Workbooks.Add
    'ヘッダーをコピー
    createdWorkbook.Range (headerRange)
    'データをコピー
    
    Dim filename As String
    filename = td(1, 1).Value
    wb.SaveAs ThisWorkbook.Path & "\" & filename & ".xlsx"
    wb.Close
    
Next i


End Sub

Function Split(th As Range, td As Range)
    Dim createdWorkbook As Workbook
    Set createdWorkbook = Workbooks.Add
    
    'ヘッダーをコピー
    createdWorkbook.Range (headerRange)
    'データをコピー

    
    Dim filename As String
    filename = td(1, 1).Value
    wb.SaveAs ThisWorkbook.Path & "\" & filename & ".xlsx"
    wb.Close
End Function

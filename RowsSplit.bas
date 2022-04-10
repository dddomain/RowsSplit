Option Explicit

Sub RowsSplit()

Dim headerRange As Range
Set headerRange = Application.InputBox _
    ( _
    prompt:="ヘッダー範囲を選択してください", _
    Title:="ヘッダー範囲選択", _
    Type:=8 _
    )

If Err.Number <> 0 Then
    MsgBox "キャンセルされました。"
    Exit Sub
End If

'ヘッダー範囲からテーブルの最後の行を取得する
Dim lastCol As Long
lastCol = headerRange.CurrentRegion.Columns.Count

'ヘッダー範囲から最初のデータ行を取得する
Dim firstDataRow As Long
firstDataRow = headerRange.Rows(headerRange.Rows.Count).row + 1
firstDataRow = Application.InputBox _
    ( _
    prompt:="データ範囲は" & firstDataRow & "行目からでよろしいですか？", _
    Title:="データの開始行確認", _
    Default:=firstDataRow, _
    Type:=2 _
    )

End Sub

Function Split(th As Range, td As Range)
    Dim wb As Workbook
    Set wb = Workbooks.Add
    
    Dim row As Long
    Dim col As Long
    row = td.Rows.Count
    col = td.Columns.Count
    
    With wb.Worksheets(1)
        'header
        Range(.Cells(1, 1), .Cells(1, col)).Value = th.Value
        'data
        Range(.Cells(2, 1), .Cells(row, col)).Value = td.Value
    End With
    
    Dim fn As String
    fn = td(1, 1).Value
    
    wb.SaveAs ThisWorkbook.Path & "\" & fn & ".xlsx"
    wb.Close
End Function

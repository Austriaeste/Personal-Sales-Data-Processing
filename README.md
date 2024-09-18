# Personal Sales Data Processing

このリポジトリでは、特定の名前に基づいて件数、合計、平均を計算し、新しいExcelワークブックに保存するVBAマクロを提供します。

## 使用方法
Excelを開き、VBAエディタを起動します（Alt + F11）。
新しいモジュールを挿入し、上記のマクロコードを貼り付けます。
マクロを実行すると、指定された名前ごとに新しいExcelファイルが作成され、件数、合計、平均が記載されます。

## 注意事項
保存パスは C:\Work\ に設定されています。必要に応じて変更してください。
名前のリストは配列で定義されています。必要に応じて追加や変更を行ってください。

## 対象のデータ

以下のデータを使用して、各名前ごとに件数、合計、平均を計算します。

| 名前   | 売上 |
| ------ | ---- |
| 菊池   | 88   |
| 佐々木 | 70   |
| 桜井   | 87   |
| 佐々木 | 73   |
| 田中   | 77   |
| 佐々木 | 52   |

## マクロ

以下のVBAマクロを使用して、各名前ごとに新しいワークブックを作成し、件数、合計、平均を記載します。

```vba
Sub PersonalSalesDate()
    Dim names As Variant
    Dim i As Integer
    Dim countName As Long
    Dim sumName As Double
    Dim avgName As Double
    Dim savePath As String
    Dim name As String
    
    ' 名前のリスト
    names = Array("菊池", "佐々木", "桜井", "田中")
    
    ' 各名前ごとに処理
    For i = LBound(names) To UBound(names)
        name = names(i)
        
        ' 件数、合計、平均を計算
        countName = WorksheetFunction.CountIf(Range("A1:A6"), name)
        sumName = WorksheetFunction.SumIf(Range("A1:A6"), name, Range("B1:B6"))
        If countName <> 0 Then
            avgName = sumName / countName
        Else
            avgName = 0
        End If
        
        ' 新しいワークブックを作成
        Workbooks.Add
        
        ' 件数、合計、平均を新しいワークブックに書き込み
        ActiveSheet.Cells(1, 1).Value = name & "の件数"
        ActiveSheet.Cells(1, 2).Value = countName
        ActiveSheet.Cells(2, 1).Value = name & "の合計"
        ActiveSheet.Cells(2, 2).Value = sumName
        ActiveSheet.Cells(3, 1).Value = name & "の平均"
        ActiveSheet.Cells(3, 2).Value = avgName
        
        ' 保存パスとファイル名を設定
        savePath = "C:\Work\" & Year(Now)
        
        ' フォルダが存在するか確認し、存在しない場合は作成
        If Dir(savePath, vbDirectory) = "" Then
            MkDir savePath
        End If
        
        ' ワークブックを保存
        ActiveWorkbook.SaveAs Filename:=savePath & "\売上_" & name & ".xlsx"
        
        ' 新しいワークブックを閉じる
        ActiveWorkbook.Close
    Next i
End Sub

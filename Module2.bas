Attribute VB_Name = "Module2"
Option Explicit

Sub csvPortingExcelFile()
    Dim fso As New Scripting.FileSystemObject
    Dim csvFile As Object
    Dim csvData As String
    Dim splitcsvData As Variant
    Dim i As Integer
    Dim j As Integer
    Dim overWorkCheck As Boolean
        
    '読み込み
    Set csvFile = fso.OpenTextFile("C:\Users\t.kawano\Desktop\残業代作成ツール\daily_2017-09-01_2017-10-01.csv", 1)
    '---------------↑読み込み対象ファイルべた書きなのでテキストボックス？から場所を読むように直す_1011_1100-------------
    
    '読んだcsvファイルの最後の行まで読み込む
    i = 1
    Do While csvFile.AtEndOfStream = False
        csvData = csvFile.ReadLine
        splitcsvData = Split(csvData, ",")
        j = UBound(splitcsvData) + 1
        
        'ファイルを作ろう
        Sheet2.Range(Sheet2.Cells(i, 1), Sheet2.Cells(i, j)).Value = splitcsvData
        i = i + 1
    Loop

    '残業時間色付けメソッドの呼び出し
    Call overTimeColoring
    
    'エクセルファイルをかきだす

    csvFile.Close
    Set csvFile = Nothing
    Set fso = Nothing
    
    '最後一番左上に選択セルを戻す
    Range("A1").Select

End Sub

'残業時間の規模に応じて背景色をつける
Private Sub overTimeColoring()
    Dim MaxCol As Integer '最大行
    Dim MaxRow As Integer '最大列
    Dim overTimeRow As Integer '残業時間の入った列
    Dim overTime As Date '残業時間
    'Dim paintTargetCell As Object '色をつけるセル
    
    Dim k As Integer
    Dim l As Integer

    '残業時間の入った行を探す
    MaxCol = Cells(1, Columns.Count).End(xlToLeft).Column
    For k = 1 To MaxCol
        If Cells(1, k).Value = "残業時間" Then
            overTimeRow = k
        End If
    Next k

    MaxRow = Cells(Rows.Count, 1).End(xlUp).Row
    '1行目は見出し,残業時間データは2行目から最後尾まで
    For l = 2 To MaxRow
        overTime = CDate(Cells(l, overTimeRow).Value)
        'paintTargetCell = Cells(l, overTimeRow)
        
       '残業時間の条件に沿って色をつける(色と時間を定数化する？)
        If overTime >= "3:00:00" Then
            Cells(l, overTimeRow).Interior.Color = RGB(226, 43, 48)
        ElseIf overTime >= "2:00:00" Then
            Cells(l, overTimeRow).Interior.Color = RGB(182, 59, 64)
        ElseIf overTime >= "1:00:00" Then
            Cells(l, overTimeRow).Interior.Color = RGB(233, 115, 155)
        Else
            '1時間未満なら色付けなし
        End If
    Next l
End Sub


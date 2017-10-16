Attribute VB_Name = "Module2"
Option Explicit

'参照先をexcファイルとしてシート出力するボタン
Sub outputExcFile_btn()
    Dim fso As New Scripting.FileSystemObject
    Dim csvFile As Object
    Dim csvData As String
    Dim splitcsvData As Variant
    Dim i As Integer
    Dim j As Integer
    Dim overWorkCheck As Boolean
        
    '参照ファイル読み込み
    Set csvFile = fso.OpenTextFile(Sheets("入力フォーム").Range("A2").Value, 1)
    i = 1
    Do While csvFile.AtEndOfStream = False
        'csvファイルを整形して読み込む
        csvData = Replace(csvFile.ReadLine, """", "")
        splitcsvData = Split(csvData, ",")
        j = UBound(splitcsvData) + 1
        '出力シートに書き出し
        Sheets("出力").Range(Sheets("出力").Cells(i, 1), Sheets("出力").Cells(i, j)).Value = splitcsvData
        i = i + 1
    Loop
    
    '残業時間色付けメソッドの呼び出し
    Call overTimeColoring
   
    'クローズ対応
    csvFile.Close
    Set csvFile = Nothing
    Set fso = Nothing
    '出力シートのA1指定状態にして動作終了する
    Sheets("出力").Activate
    Sheets("出力").Range("A1").Select

End Sub

'残業時間の規模に応じて各セルに背景色をつける
Private Sub overTimeColoring()
    Dim MaxCol As Integer '最大行
    Dim MaxRow As Integer '最大列
    Dim overTimeRow As Integer '残業時間カラム
    Dim overTime As Date '残業時間
    Dim k As Integer
    Dim l As Integer
    
    '残業時間カラムの位置を探す
    MaxCol = Sheets("出力").Cells(1, Columns.Count).End(xlToLeft).Column
    For k = 1 To MaxCol
        If Sheets("出力").Cells(1, k).Value = "残業時間" Then
            overTimeRow = k
        End If
    Next k
    
    '残業時間を2行目から最後尾で取得する
    MaxRow = Sheets("出力").Cells(Rows.Count, 1).End(xlUp).Row
    For l = 2 To MaxRow
        overTime = CDate(Sheets("出力").Cells(l, overTimeRow).Value)
       '残業時間の条件に沿って色をつける
       With Sheets("出力").Cells(l, overTimeRow)
            If overTime >= "3:00:00" Then
                .Interior.Color = RGB(226, 43, 48) '真赤
            ElseIf overTime >= "2:00:00" Then
                .Interior.Color = RGB(182, 59, 64) '薄め赤
            ElseIf overTime >= "1:00:00" Then
                .Interior.Color = RGB(233, 115, 155) 'もう少し薄め赤
            Else
                '1時間未満の残業なら色付けなし
            End If
        End With
    Next l
End Sub

'---------終了ボタン--------
Public Sub exitApp_btn()
    Application.Quit
 End Sub
 
 '---------参照ボタン--------
Private Sub choiceFile_btn()
    Dim objFS           As New FileSystemObject
    Dim strPath         As String
    Dim strFile         As String
    Dim strFolder       As String
    Dim ofdFileDlg    As Office.FileDialog

    strPath = Sheets("入力フォーム").Range("A2").Value

    ' 初期パスの設定
    If Len(strPath) > 0 Then
        ' 末尾の"\"削除
        If Right(strPath, 1) = "\" Then
            strPath = Left(strPath, Len(strPath) - 1)
        End If

        ' ファイルが存在
        If objFS.FileExists(strPath) Then
            ' ファイル名のみ取得
            strFile = objFS.GetFileName(strPath)
            ' フォルダパスのみ取得
            strFolder = objFS.GetParentFolderName(strPath)
        ' ファイルが存在しない
        Else
            ' フォルダが存在
            If objFS.FolderExists(strPath) Then
                strFile = ""
                strFolder = strPath
            ' フォルダが存在しない
            Else
                ' ファイル名のみ取得
                strFile = objFS.GetFileName(strPath)
                ' 親フォルダを取得
                strFolder = objFS.GetParentFolderName(strPath)
                ' 親フォルダが存在しない
                If Not objFS.FolderExists(strFolder) Then
                    strFolder = ThisWorkbook.Path
                End If
            End If
        End If
        Set objFS = Nothing
    Else
        strFolder = ThisWorkbook.Path
        strFile = ""
    End If

    ' ファイル選択ダイアログ設定
    Set ofdFileDlg = Application.FileDialog(msoFileDialogFilePicker)
    With ofdFileDlg
        .ButtonName = "選択"
        '「ファイルの種類」をクリア
        .Filters.Clear
        '「ファイルの種類」を登録
        .Filters.Add "CSVファイル", "*.?sv", 1
        .Filters.Add "全ファイル", "*.*", 2

        ' 初期フォルダ
        .InitialFileName = strFolder & "\" & strFile
        ' 複数選択不可
        .AllowMultiSelect = False
        '表示するアイコンの大きさを指定
        .InitialView = msoFileDialogViewDetails
    End With

    ' フォルダ選択ダイアログ表示
    If ofdFileDlg.Show() = -1 Then
        ' フォルダパス設定
        Sheets("入力フォーム").Range("A2").Value = ofdFileDlg.SelectedItems(1)
    End If

    Set ofdFileDlg = Nothing
    Exit Sub
End Sub

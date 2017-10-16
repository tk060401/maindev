Attribute VB_Name = "Module2"
Option Explicit

'参照先をexcファイルとしてシート出力するボタン
Sub outputExcFile_btn()
    '参照ファイル読み込みとマスター書き出し
    Call readCsvExportMasterSheet
    '書き出されたシートに色をつける
    Call touchCollerSheet
    
    '出力シートのA1指定状態にして動作終了する
    Sheets("出力").Activate
    Sheets("出力").Range("A1").Select
End Sub
'参照ファイル読み込み
Private Sub readCsvExportMasterSheet()
    Dim fso As New Scripting.FileSystemObject
    Dim csvFile As Object
    Dim csvData As String
    Dim splitcsvData As Variant
    Dim i As Integer
    Dim j As Integer
    Set csvFile = fso.OpenTextFile(Sheets("入力フォーム").Range("A2").Value, 1)
    i = 1
    Do While csvFile.AtEndOfStream = False
        'csvファイルを整形して読み込む
        csvData = Replace(csvFile.ReadLine, """", "")
        splitcsvData = Split(csvData, ",")
        j = UBound(splitcsvData) + 1
        '出力シート(マスター)に書き出し
        Sheets("出力").Range(Sheets("出力").Cells(i, 1), Sheets("出力").Cells(i, j)).Value = splitcsvData
        i = i + 1
    Loop
    
    'クローズ対応
    csvFile.Close
    Set csvFile = Nothing
    Set fso = Nothing

End Sub
'残業時間の規模に応じて各セルに背景色をつける
Private Sub touchCollerSheet()
    Dim MaxCol As Integer '最大行
    Dim MaxRow As Integer '最大列
    Dim overTimeRow As Integer '残業時間カラム番号
    Dim employerCordRow As Integer '社員コードカラム番号
    Dim overTime As Date '残業時間
    Dim employerCord As Integer '社員コード
    Dim k As Integer
    Dim l As Integer
    Dim s As Integer
    Dim y As Integer
    s = 2
    y = 2
        
    '出力に必要な最後尾カラムを取得
    MaxCol = Sheets("出力").Cells(1, Columns.Count).End(xlToLeft).Column
    MaxRow = Sheets("出力").Cells(Rows.Count, 1).End(xlUp).Row
    
    '残業時間と社員コードの入った列を取得
    For k = 1 To MaxCol
        With Sheets("出力").Cells(1, k)
            If .Value = "残業時間" Then
                overTimeRow = k
            ElseIf .Value = "社員コード" Then
                employerCordRow = k
            End If
        End With
    Next k
    
    '管理者別シートの作成(べた書き・・管理者変数 as Array みたいにしたいけどvbaよくわからない)
    With Worksheets.Add(after:=Worksheets(Worksheets.Count))
        .Name = "sakai"
    End With
    With Worksheets.Add(after:=Worksheets(Worksheets.Count))
       .Name = "yoshiike"
    End With
    
    '1行目だけ先にまとめて出力(カラムの見出し部分,管理者が増えるたびに追加しなければならないのがつらい)
    Sheets("sakai").Range(Sheets("sakai").Cells(1, 1), Sheets("sakai").Cells(1, MaxCol)).Value = Sheets("出力").Range(Sheets("出力").Cells(1, 1), Sheets("出力").Cells(1, MaxCol)).Value
    Sheets("yoshiike").Range(Sheets("yoshiike").Cells(1, 1), Sheets("yoshiike").Cells(1, MaxCol)).Value = Sheets("出力").Range(Sheets("出力").Cells(1, 1), Sheets("出力").Cells(1, MaxCol)).Value

    For l = 2 To MaxRow
        overTime = CDate(Sheets("出力").Cells(l, overTimeRow).Value)
        With Sheets("出力").Cells(l, overTimeRow)
            '残業時間条件に沿ってセル着色
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
        
       '勤怠管理者ごとにシートを作成してコピー
       employerCord = Sheets("出力").Cells(l, employerCordRow).Value
       'とりあえず条件ベタ書き(in_arrayがなくて困っている)
       Select Case employerCord
            Case 44, 48, 52, 58, 66, 137, 149, 151, 167, 203, 227, 270, 297
                '酒井さんグループ(仮)
                Sheets("sakai").Range(Sheets("sakai").Cells(s, 1), Sheets("sakai").Cells(s, MaxCol)).Value = Sheets("出力").Range(Sheets("出力").Cells(l, 1), Sheets("出力").Cells(l, MaxCol)).Value
                Sheets("sakai").Cells(s, overTimeRow).Interior.Color = Sheets("出力").Cells(l, overTimeRow).Interior.Color
                s = s + 1
            Case 8, 314, 343, 355, 357, 365, 368, 373, 382, 384, 396, 401, 408
                '吉池さんグループ(仮)
                Sheets("yoshiike").Range(Sheets("yoshiike").Cells(y, 1), Sheets("yoshiike").Cells(y, MaxCol)).Value = Sheets("出力").Range(Sheets("出力").Cells(l, 1), Sheets("出力").Cells(l, MaxCol)).Value
                Sheets("yoshiike").Cells(y, overTimeRow).Interior.Color = Sheets("出力").Cells(l, overTimeRow).Interior.Color
                y = y + 1
            Case Else
       End Select
    Next l
End Sub
'---------終了ボタン--------
Private Sub exitApp_btn()
    Application.Quit
 End Sub
 '--------管理者シートの削除
 Private Sub deleteSheet_btn()
 
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
                    strFolder = ThisWorkbook.path
                End If
            End If
        End If
        Set objFS = Nothing
    Else
        strFolder = ThisWorkbook.path
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

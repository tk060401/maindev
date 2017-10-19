Attribute VB_Name = "Module2"
Option Explicit
Sub outputExcFile_btn()
    Dim g_masterSheet As String
    g_masterSheet = "出力"
    '参照ファイル読み込み、出力シートへマスター吐き出し
    Call loadCsvOutputMasterSheet(g_masterSheet)
    'シート加工、管理者シート作成
    Call editSheet(g_masterSheet)
    
    '出力シートのA1指定状態にして動作終了する
    Sheets(g_masterSheet).Activate
    Sheets(g_masterSheet).Range("A1").Select
End Sub
'参照ファイル読み込みと出力シートに無編集で書き出し
Private Sub loadCsvOutputMasterSheet(ByVal g_masterSheet As String)
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
        '出力シートに書き出し
        Sheets(g_masterSheet).Range(Sheets(g_masterSheet).Cells(i, 1), Sheets(g_masterSheet).Cells(i, j)).Value = splitcsvData
        i = i + 1
    Loop
    
    'クローズ対応
    csvFile.Close
    Set csvFile = Nothing
    Set fso = Nothing

End Sub
Private Sub editSheet(ByVal g_masterSheet As String)
    Dim MaxCol As Integer '最大行
    Dim MaxRow As Integer '最大列
    Dim overTimeRow As Integer '残業時間カラム番号
    Dim employerCordRow As Integer '社員コードカラム番号
    Dim ymRow As Integer '月度カラム番号
    Dim overTime As Date '残業時間
    Dim employerCord As Long '社員コード
    Dim ymCord As Long '年月度

    '出力に必要な最後尾カラムを取得
    MaxCol = Sheets(g_masterSheet).Cells(1, Columns.count).End(xlToLeft).Column
    MaxRow = Sheets(g_masterSheet).Cells(Rows.count, 1).End(xlUp).Row
    '「残業時間」と「社員コード」と「月度」の入った列を取得
    Dim k As Integer
    For k = 1 To MaxCol
        With Sheets(g_masterSheet).Cells(1, k)
            If .Value = "残業時間" Then
                overTimeRow = k
            ElseIf .Value = "社員コード" Then
                employerCordRow = k
            ElseIf .Value = "月度" Then
                ymRow = k
            End If
        End With
    Next k

    '(年月,社員ID)検索条件読み込み
    Dim inputYm As Variant '入力された年月
    Dim inputEmployerCord As Variant '入力された社員コード
    inputYm = Sheets("入力フォーム").Range("H3").Value
    inputEmployerCord = Sheets("入力フォーム").Range("H5").Value
    
     'ユーザーが検索したかどうか
    Dim isYm As Boolean '検索年月の有無
    Dim isEmployerCord As Boolean '検索社員コードの有無
    isYm = IIf(inputYm <> 0, True, False)
    isEmployerCord = IIf(inputEmployerCord <> 0, True, False)
    
    Dim masterRecord As Range
    Dim masterRecordOverTimeColler As Integer
    Dim managerRecord As Range
    Dim managerRecordOverTimeColler As Integer
    
    '管理者名一覧設定
    Dim managerNameList() As Variant
    managerNameList = Array("sakai", "yoshiike", "hogehoge")
    Dim employeeCordList(0 To 2) As Variant '管理者ごとの社員コード設定
    employeeCordList(0) = Array(44, 48, 52, 58, 66, 137, 149, 151, 167, 203, 227, 270, 297) '酒井さんグループ(仮)
    employeeCordList(1) = Array(8, 314, 343, 355, 357, 365, 368, 373, 382, 384, 396, 401, 408) '吉池さんグループ(仮)
    employeeCordList(2) = Array(314, 343, 355, 357, 365, 368, 373, 382, 384, 396, 401, 408) 'hogeさんグループ(仮)
    
    Dim managerSum As Integer
    managerSum = UBound(managerNameList) '管理者人数
    
    '管理者シートの作成(1行目カラム一覧まで)
    Call createManagerSheet(MaxCol, managerNameList, g_masterSheet)
    '出力シートの残業時間による着色
    Call touchCollerOverTimeCell(overTimeRow, MaxRow, g_masterSheet)

    Dim managerCount As Integer
    Dim recordCount As Integer
    Dim l As Integer
    For managerCount = 0 To managerSum
            recordCount = 2
            For l = 2 To MaxRow
                '社員ID
                employerCord = Sheets(g_masterSheet).Cells(l, employerCordRow).Value
                '年月度
                ymCord = Sheets(g_masterSheet).Cells(l, ymRow).Value
                '出力シートのレコード位置
                Set masterRecord = Sheets(g_masterSheet).Range(Sheets(g_masterSheet).Cells(l, 1), Sheets(g_masterSheet).Cells(l, MaxCol))
                
                If inArray(employerCord, employeeCordList(managerCount)) = False Then
                    GoTo Continue ' 検索外の社員コードならコピー(以下処理)しない
                End If
                '管理者シートのレコード
                Set managerRecord = Sheets(managerNameList(managerCount)).Range(Sheets(managerNameList(managerCount)).Cells(recordCount, 1), Sheets(managerNameList(managerCount)).Cells(recordCount, MaxCol))
                
                Dim IsInportRecord As Boolean
                If IsInportRecord = isInportRecordSheet(isYm, inputEmployerCord, inputYm, inputEmployerCord, ymCord, employerCord) Then
                    managerRecord.Value = masterRecord.Value
                    recordCount = recordCount + 1
                End If
Continue:
            Next l
            'マスターシートと同じように色付け
            Call touchCollerOverTimeCell(overTimeRow, Sheets(managerNameList(managerCount)).Cells(Rows.count, 1).End(xlUp).Row, managerNameList(managerCount))
    Next managerCount
End Sub
Private Function isInportRecordSheet(ByVal isYm As Boolean, ByVal isEmployerCord As Boolean, ByVal inputYm As Long, ByVal inputEmployerCord As Long, ByVal ymCord As Long, ByVal employerCord As Long)
    Dim IsCopyRecord As Boolean
    IsCopyRecord = False
    If isYm And isEmployerCord And inputYm = ymCord And inputEmployerCord = employerCord Then
        IsCopyRecord = True
    ElseIf isYm And inputYm = ymCord Then
        IsCopyRecord = True
    ElseIf isEmployerCord And inputEmployerCord = employerCord Then
        IsCopyRecord = True
    End If
End Function
'配列内検索
Public Function inArray(ByVal needle As Variant, ByVal haystack As Variant) As Boolean
    Dim theValue As Variant
    For Each theValue In haystack
        If needle = theValue Then
            inArray = True
            Exit Function
        End If
    Next theValue
    inArray = False
End Function
'残業時間セルへの色付け
Public Sub touchCollerOverTimeCell(ByVal overTimeRow As Integer, ByVal MaxRow As Integer, ByVal g_masterSheet As String)
    Dim l As Integer
    Dim overTime As Date
    For l = 2 To MaxRow
        overTime = CDate(Sheets(g_masterSheet).Cells(l, overTimeRow).Value)
        With Sheets(g_masterSheet).Cells(l, overTimeRow)
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
    Next l
End Sub
'管理者シートの作成(1行目カラム一覧まで)
Private Sub createManagerSheet(ByVal MaxCol As Integer, ByVal managerNameList As Variant, ByVal g_masterSheet As String)
    Dim mName As Variant
    Dim sCount As Integer
    sCount = Sheets.count
    Dim i As Integer
    For Each mName In managerNameList
        '同名シートの存在チェック
        For i = 1 To sCount
            If mName = Worksheets(i).Name Then
                MsgBox "古い管理者シートを捨てるか名前を変えてください"
                End
            End If
        Next i
        With Worksheets.Add(after:=Worksheets(Worksheets.count))
            .Name = mName
        End With
        Sheets(mName).Range(Sheets(mName).Cells(1, 1), Sheets(mName).Cells(1, MaxCol)).Value = _
            Sheets(g_masterSheet).Range(Sheets(g_masterSheet).Cells(1, 1), Sheets(g_masterSheet).Cells(1, MaxCol)).Value
    Next
End Sub
'---------終了ボタン--------
Private Sub exitApp_btn()
    Application.Quit
End Sub
'--------管理者シート削除ボタン
Private Sub deleteSheet_btn()
     '管理者名一覧の読み込み
    Dim defaultSheetList As Variant
    defaultSheetList = Array("入力フォーム", "出力")
    Dim i As Long
    Dim sCount As Long
    
    'シート枚数
    sCount = Sheets.count
    For i = 1 To sCount
        '「入力フォーム」「出力」シートは削除処理不要
        If sCount = 2 Then
            Exit Sub
        End If
        If inArray(Worksheets(i).Name, defaultSheetList) = False Then
            Application.DisplayAlerts = False
            Worksheets(Worksheets(i).Name).Delete
            Application.DisplayAlerts = True
            
            'シート削除すると枚数が1枚少なくなるので処理を入れる(範囲エラーが起きる)
            sCount = Sheets.count
            i = i - 1
        End If
    Next i
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


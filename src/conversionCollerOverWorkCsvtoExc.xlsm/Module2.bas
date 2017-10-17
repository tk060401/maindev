Attribute VB_Name = "Module2"
Option Explicit

Dim g_masterSheet As String
Sub outputExcFile_btn()
    g_masterSheet = "出力"
    '参照ファイル読み込みとマスター書き出し
    Call loadCsvExportMasterSheet
    '残業時間に色をつける
    Call touchCollerOverTimeCell
    
    '出力シートのA1指定状態にして動作終了する
    Sheets(g_masterSheet).Activate
    Sheets(g_masterSheet).Range("A1").Select
End Sub
'参照ファイル読み込み
Private Sub loadCsvExportMasterSheet()
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
        Sheets(g_masterSheet).Range(Sheets(g_masterSheet).Cells(i, 1), Sheets(g_masterSheet).Cells(i, j)).Value = splitcsvData
        i = i + 1
    Loop
    
    'クローズ対応
    csvFile.Close
    Set csvFile = Nothing
    Set fso = Nothing

End Sub
Private Sub touchCollerOverTimeCell()
    Dim MaxCol As Integer '最大行
    Dim MaxRow As Integer '最大列
    Dim overTimeRow As Integer '残業時間カラム番号
    Dim employerCordRow As Integer '社員コードカラム番号
    Dim ymRow As Integer '月度カラム番号
    Dim overTime As Date '残業時間
    Dim employerCord As Integer '社員コード
        
    '出力に必要な最後尾カラムを取得
    MaxCol = Sheets(g_masterSheet).Cells(1, Columns.Count).End(xlToLeft).Column
    MaxRow = Sheets(g_masterSheet).Cells(Rows.Count, 1).End(xlUp).Row
    '残業時間と社員コードの入った列を取得
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
    
    '管理者名一覧の読み込み
    Dim managerNameList() As Variant
    managerNameList = Array("sakai", "yoshiike")
    Dim mName As Variant
    '管理者ごとにシートを作ってテンプレ作成
    For Each mName In managerNameList
        With Worksheets.Add(after:=Worksheets(Worksheets.Count))
            .Name = mName
        End With
        Sheets(mName).Range(Sheets(mName).Cells(1, 1), Sheets(mName).Cells(1, MaxCol)).Value = _
            Sheets(g_masterSheet).Range(Sheets(g_masterSheet).Cells(1, 1), Sheets(g_masterSheet).Cells(1, MaxCol)).Value
    Next
    
    '検索条件読み込み
    Dim ym As String
    Dim empCord As Integer
    Dim isDateCondition As Boolean
    Dim l As Integer
    Dim s As Integer
    Dim y As Integer
    s = 2
    y = 2

    ym = Sheets("入力フォーム").Range("H3").Value
    empCord = Sheets("入力フォーム").Range("H5").Value
    If IsNull(ym) Then
        isDateCondition = False
    Else
        isDateCondition = True
    End If
    
    If IsNull(empCord) Then
        isEmpCord = False
    Else
        isEmpCord = True
    End If
    
    '社員コード一覧の読み込み
    Dim employeeCordList(0 To 1) As Variant
    employeeCordList(0) = Array(44, 48, 52, 58, 66, 137, 149, 151, 167, 203, 227, 270, 297)
    employeeCordList(1) = Array(8, 314, 343, 355, 357, 365, 368, 373, 382, 384, 396, 401, 408)

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
        
       '勤怠管理者ごとにシートを作成してコピー
       employerCord = Sheets(g_masterSheet).Cells(l, employerCordRow).Value
       'とりあえず条件ベタ書き(in_arrayがなくて困っている)
       Select Case employerCord
            Case 44, 48, 52, 58, 66, 137, 149, 151, 167, 203, 227, 270, 297
                '酒井さんグループ(仮)
                If isDateCondition = True Then
                    If ym = Sheets(g_masterSheet).Cells(l, ymRow).Value Then
                        Sheets(managerNameList(0)).Range(Sheets(managerNameList(0)).Cells(s, 1), Sheets(managerNameList(0)).Cells(s, MaxCol)).Value = _
                            Sheets(g_masterSheet).Range(Sheets(g_masterSheet).Cells(l, 1), Sheets(g_masterSheet).Cells(l, MaxCol)).Value
                        Sheets(managerNameList(0)).Cells(s, overTimeRow).Interior.Color = Sheets(g_masterSheet).Cells(l, overTimeRow).Interior.Color
                        s = s + 1
                    End If
                Else
                    Sheets(managerNameList(0)).Range(Sheets(managerNameList(0)).Cells(s, 1), Sheets(managerNameList(0)).Cells(s, MaxCol)).Value = _
                        Sheets(g_masterSheet).Range(Sheets(g_masterSheet).Cells(l, 1), Sheets(g_masterSheet).Cells(l, MaxCol)).Value
                    Sheets(managerNameList(0)).Cells(s, overTimeRow).Interior.Color = Sheets(g_masterSheet).Cells(l, overTimeRow).Interior.Color
                    s = s + 1
                End If
            Case 8, 314, 343, 355, 357, 365, 368, 373, 382, 384, 396, 401, 408
                '吉池さんグループ(仮)
                If isDateCondition = True Then
                    If ym = Sheets(g_masterSheet).Cells(l, ymRow).Value Then
                        Sheets(managerNameList(1)).Range(Sheets(managerNameList(1)).Cells(y, 1), Sheets(managerNameList(1)).Cells(y, MaxCol)).Value = _
                            Sheets(g_masterSheet).Range(Sheets(g_masterSheet).Cells(l, 1), Sheets(g_masterSheet).Cells(l, MaxCol)).Value
                        Sheets(managerNameList(1)).Cells(y, overTimeRow).Interior.Color = Sheets(g_masterSheet).Cells(l, overTimeRow).Interior.Color
                        y = y + 1
                    End If
                Else
                    Sheets(managerNameList(1)).Range(Sheets(managerNameList(1)).Cells(y, 1), Sheets(managerNameList(1)).Cells(y, MaxCol)).Value = _
                        Sheets(g_masterSheet).Range(Sheets(g_masterSheet).Cells(l, 1), Sheets(g_masterSheet).Cells(l, MaxCol)).Value
                    Sheets(managerNameList(1)).Cells(y, overTimeRow).Interior.Color = Sheets(g_masterSheet).Cells(l, overTimeRow).Interior.Color
                    y = y + 1
                End If
            Case Else
       End Select
    Next l
End Sub
'検索条件の読み込み
Private Function loadSearchConditions() As Variant
        
    'loadSearchConditions(0) = Sheets("入力フォーム").Range("H3").Value
    'loadSearchConditions(1) = Sheets("入力フォーム").Range("H4").Value
    
End Function
'---------終了ボタン--------
Private Sub exitApp_btn()
    Application.Quit
End Sub
 '--------管理者シート削除ボタン
Private Sub deleteSheet_btn()
     '管理者名一覧の読み込み
    'Dim managerNameList As Variant
    'managerNameList = readManagerNameList()
    managerNameList = Array("sakai", "yoshiike")
End Sub
'--------管理者一覧
Private Function readManagerNameList() As Variant
    '管理者リスト
    Dim managerNameList() As Variant
    managerNameList = Array("sakai", "yoshiike")
End Function
'--------社員コード一覧
Private Function readEmployeeCordList()
    'Dim employeeCordList(0 To 1) As Variant
    'employeeCordList(0) = Array(44, 48, 52, 58, 66, 137, 149, 151, 167, 203, 227, 270, 297)
    'employeeCordList(1) = Array(8, 314, 343, 355, 357, 365, 368, 373, 382, 384, 396, 401, 408)
End Function
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

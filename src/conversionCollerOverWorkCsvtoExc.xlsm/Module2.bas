Attribute VB_Name = "Module2"
Option Explicit
Sub outputExcFile_btn()
    Dim g_masterSheet As String
    g_masterSheet = "�o��"
    '�Q�ƃt�@�C���ǂݍ��݁A�o�̓V�[�g�փ}�X�^�[�f���o��
    Call loadCsvOutputMasterSheet(g_masterSheet)
    '�V�[�g���H�A�Ǘ��҃V�[�g�쐬
    Call editSheet(g_masterSheet)
    
    '�o�̓V�[�g��A1�w���Ԃɂ��ē���I������
    Sheets(g_masterSheet).Activate
    Sheets(g_masterSheet).Range("A1").Select
End Sub
'�Q�ƃt�@�C���ǂݍ��݂Əo�̓V�[�g�ɖ��ҏW�ŏ����o��
Private Sub loadCsvOutputMasterSheet(ByVal g_masterSheet As String)
    Dim fso As New Scripting.FileSystemObject
    Dim csvFile As Object
    Dim csvData As String
    Dim splitcsvData As Variant
    Dim i As Integer
    Dim j As Integer
    Set csvFile = fso.OpenTextFile(Sheets("���̓t�H�[��").Range("A2").Value, 1)
    i = 1
    Do While csvFile.AtEndOfStream = False
        'csv�t�@�C���𐮌`���ēǂݍ���
        csvData = Replace(csvFile.ReadLine, """", "")
        splitcsvData = Split(csvData, ",")
        j = UBound(splitcsvData) + 1
        '�o�̓V�[�g�ɏ����o��
        Sheets(g_masterSheet).Range(Sheets(g_masterSheet).Cells(i, 1), Sheets(g_masterSheet).Cells(i, j)).Value = splitcsvData
        i = i + 1
    Loop
    
    '�N���[�Y�Ή�
    csvFile.Close
    Set csvFile = Nothing
    Set fso = Nothing

End Sub
Private Sub editSheet(ByVal g_masterSheet As String)
    Dim MaxCol As Integer '�ő�s
    Dim MaxRow As Integer '�ő��
    Dim overTimeRow As Integer '�c�Ǝ��ԃJ�����ԍ�
    Dim employerCordRow As Integer '�Ј��R�[�h�J�����ԍ�
    Dim ymRow As Integer '���x�J�����ԍ�
    Dim overTime As Date '�c�Ǝ���
    Dim employerCord As Long '�Ј��R�[�h
    Dim ymCord As Long '�N���x

    '�o�͂ɕK�v�ȍŌ���J�������擾
    MaxCol = Sheets(g_masterSheet).Cells(1, Columns.count).End(xlToLeft).Column
    MaxRow = Sheets(g_masterSheet).Cells(Rows.count, 1).End(xlUp).Row
    '�u�c�Ǝ��ԁv�Ɓu�Ј��R�[�h�v�Ɓu���x�v�̓���������擾
    Dim k As Integer
    For k = 1 To MaxCol
        With Sheets(g_masterSheet).Cells(1, k)
            If .Value = "�c�Ǝ���" Then
                overTimeRow = k
            ElseIf .Value = "�Ј��R�[�h" Then
                employerCordRow = k
            ElseIf .Value = "���x" Then
                ymRow = k
            End If
        End With
    Next k

    '(�N��,�Ј�ID)���������ǂݍ���
    Dim inputYm As Variant '���͂��ꂽ�N��
    Dim inputEmployerCord As Variant '���͂��ꂽ�Ј��R�[�h
    inputYm = Sheets("���̓t�H�[��").Range("H3").Value
    inputEmployerCord = Sheets("���̓t�H�[��").Range("H5").Value
    
     '���[�U�[�������������ǂ���
    Dim isYm As Boolean '�����N���̗L��
    Dim isEmployerCord As Boolean '�����Ј��R�[�h�̗L��
    isYm = IIf(inputYm <> 0, True, False)
    isEmployerCord = IIf(inputEmployerCord <> 0, True, False)
    
    Dim masterRecord As Range
    Dim masterRecordOverTimeColler As Integer
    Dim managerRecord As Range
    Dim managerRecordOverTimeColler As Integer
    
    '�Ǘ��Җ��ꗗ�ݒ�
    Dim managerNameList() As Variant
    managerNameList = Array("sakai", "yoshiike", "hogehoge")
    Dim employeeCordList(0 To 2) As Variant '�Ǘ��҂��Ƃ̎Ј��R�[�h�ݒ�
    employeeCordList(0) = Array(44, 48, 52, 58, 66, 137, 149, 151, 167, 203, 227, 270, 297) '���䂳��O���[�v(��)
    employeeCordList(1) = Array(8, 314, 343, 355, 357, 365, 368, 373, 382, 384, 396, 401, 408) '�g�r����O���[�v(��)
    employeeCordList(2) = Array(314, 343, 355, 357, 365, 368, 373, 382, 384, 396, 401, 408) 'hoge����O���[�v(��)
    
    Dim managerSum As Integer
    managerSum = UBound(managerNameList) '�Ǘ��Ґl��
    
    '�Ǘ��҃V�[�g�̍쐬(1�s�ڃJ�����ꗗ�܂�)
    Call createManagerSheet(MaxCol, managerNameList, g_masterSheet)
    '�o�̓V�[�g�̎c�Ǝ��Ԃɂ�钅�F
    Call touchCollerOverTimeCell(overTimeRow, MaxRow, g_masterSheet)

    Dim managerCount As Integer
    Dim recordCount As Integer
    Dim l As Integer
    For managerCount = 0 To managerSum
            recordCount = 2
            For l = 2 To MaxRow
                '�Ј�ID
                employerCord = Sheets(g_masterSheet).Cells(l, employerCordRow).Value
                '�N���x
                ymCord = Sheets(g_masterSheet).Cells(l, ymRow).Value
                '�o�̓V�[�g�̃��R�[�h�ʒu
                Set masterRecord = Sheets(g_masterSheet).Range(Sheets(g_masterSheet).Cells(l, 1), Sheets(g_masterSheet).Cells(l, MaxCol))
                
                If inArray(employerCord, employeeCordList(managerCount)) = False Then
                    GoTo Continue ' �����O�̎Ј��R�[�h�Ȃ�R�s�[(�ȉ�����)���Ȃ�
                End If
                '�Ǘ��҃V�[�g�̃��R�[�h
                Set managerRecord = Sheets(managerNameList(managerCount)).Range(Sheets(managerNameList(managerCount)).Cells(recordCount, 1), Sheets(managerNameList(managerCount)).Cells(recordCount, MaxCol))
                
                Dim IsInportRecord As Boolean
                If IsInportRecord = isInportRecordSheet(isYm, inputEmployerCord, inputYm, inputEmployerCord, ymCord, employerCord) Then
                    managerRecord.Value = masterRecord.Value
                    recordCount = recordCount + 1
                End If
Continue:
            Next l
            '�}�X�^�[�V�[�g�Ɠ����悤�ɐF�t��
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
'�z�������
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
'�c�Ǝ��ԃZ���ւ̐F�t��
Public Sub touchCollerOverTimeCell(ByVal overTimeRow As Integer, ByVal MaxRow As Integer, ByVal g_masterSheet As String)
    Dim l As Integer
    Dim overTime As Date
    For l = 2 To MaxRow
        overTime = CDate(Sheets(g_masterSheet).Cells(l, overTimeRow).Value)
        With Sheets(g_masterSheet).Cells(l, overTimeRow)
            '�c�Ǝ��ԏ����ɉ����ăZ�����F
            If overTime >= "3:00:00" Then
                .Interior.Color = RGB(226, 43, 48) '�^��
            ElseIf overTime >= "2:00:00" Then
                .Interior.Color = RGB(182, 59, 64) '���ߐ�
            ElseIf overTime >= "1:00:00" Then
                .Interior.Color = RGB(233, 115, 155) '�����������ߐ�
            Else
                '1���Ԗ����̎c�ƂȂ�F�t���Ȃ�
            End If
        End With
    Next l
End Sub
'�Ǘ��҃V�[�g�̍쐬(1�s�ڃJ�����ꗗ�܂�)
Private Sub createManagerSheet(ByVal MaxCol As Integer, ByVal managerNameList As Variant, ByVal g_masterSheet As String)
    Dim mName As Variant
    Dim sCount As Integer
    sCount = Sheets.count
    Dim i As Integer
    For Each mName In managerNameList
        '�����V�[�g�̑��݃`�F�b�N
        For i = 1 To sCount
            If mName = Worksheets(i).Name Then
                MsgBox "�Â��Ǘ��҃V�[�g���̂Ă邩���O��ς��Ă�������"
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
'---------�I���{�^��--------
Private Sub exitApp_btn()
    Application.Quit
End Sub
'--------�Ǘ��҃V�[�g�폜�{�^��
Private Sub deleteSheet_btn()
     '�Ǘ��Җ��ꗗ�̓ǂݍ���
    Dim defaultSheetList As Variant
    defaultSheetList = Array("���̓t�H�[��", "�o��")
    Dim i As Long
    Dim sCount As Long
    
    '�V�[�g����
    sCount = Sheets.count
    For i = 1 To sCount
        '�u���̓t�H�[���v�u�o�́v�V�[�g�͍폜�����s�v
        If sCount = 2 Then
            Exit Sub
        End If
        If inArray(Worksheets(i).Name, defaultSheetList) = False Then
            Application.DisplayAlerts = False
            Worksheets(Worksheets(i).Name).Delete
            Application.DisplayAlerts = True
            
            '�V�[�g�폜����Ɩ�����1�����Ȃ��Ȃ�̂ŏ���������(�͈̓G���[���N����)
            sCount = Sheets.count
            i = i - 1
        End If
    Next i
End Sub
 '---------�Q�ƃ{�^��--------
Private Sub choiceFile_btn()
    Dim objFS           As New FileSystemObject
    Dim strPath         As String
    Dim strFile         As String
    Dim strFolder       As String
    Dim ofdFileDlg    As Office.FileDialog

    strPath = Sheets("���̓t�H�[��").Range("A2").Value

    ' �����p�X�̐ݒ�
    If Len(strPath) > 0 Then
        ' ������"\"�폜
        If Right(strPath, 1) = "\" Then
            strPath = Left(strPath, Len(strPath) - 1)
        End If

        ' �t�@�C��������
        If objFS.FileExists(strPath) Then
            ' �t�@�C�����̂ݎ擾
            strFile = objFS.GetFileName(strPath)
            ' �t�H���_�p�X�̂ݎ擾
            strFolder = objFS.GetParentFolderName(strPath)
        ' �t�@�C�������݂��Ȃ�
        Else
            ' �t�H���_������
            If objFS.FolderExists(strPath) Then
                strFile = ""
                strFolder = strPath
            ' �t�H���_�����݂��Ȃ�
            Else
                ' �t�@�C�����̂ݎ擾
                strFile = objFS.GetFileName(strPath)
                ' �e�t�H���_���擾
                strFolder = objFS.GetParentFolderName(strPath)
                ' �e�t�H���_�����݂��Ȃ�
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

    ' �t�@�C���I���_�C�A���O�ݒ�
    Set ofdFileDlg = Application.FileDialog(msoFileDialogFilePicker)
    With ofdFileDlg
        .ButtonName = "�I��"
        '�u�t�@�C���̎�ށv���N���A
        .Filters.Clear
        '�u�t�@�C���̎�ށv��o�^
        .Filters.Add "CSV�t�@�C��", "*.?sv", 1
        .Filters.Add "�S�t�@�C��", "*.*", 2

        ' �����t�H���_
        .InitialFileName = strFolder & "\" & strFile
        ' �����I��s��
        .AllowMultiSelect = False
        '�\������A�C�R���̑傫�����w��
        .InitialView = msoFileDialogViewDetails
    End With

    ' �t�H���_�I���_�C�A���O�\��
    If ofdFileDlg.Show() = -1 Then
        ' �t�H���_�p�X�ݒ�
        Sheets("���̓t�H�[��").Range("A2").Value = ofdFileDlg.SelectedItems(1)
    End If

    Set ofdFileDlg = Nothing
    Exit Sub
End Sub


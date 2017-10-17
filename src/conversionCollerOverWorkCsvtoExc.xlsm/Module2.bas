Attribute VB_Name = "Module2"
Option Explicit

Dim g_masterSheet As String
Sub outputExcFile_btn()
    g_masterSheet = "�o��"
    '�Q�ƃt�@�C���ǂݍ��݂ƃ}�X�^�[�����o��
    Call loadCsvExportMasterSheet
    '�c�Ǝ��ԂɐF������
    Call touchCollerOverTimeCell
    
    '�o�̓V�[�g��A1�w���Ԃɂ��ē���I������
    Sheets(g_masterSheet).Activate
    Sheets(g_masterSheet).Range("A1").Select
End Sub
'�Q�ƃt�@�C���ǂݍ���
Private Sub loadCsvExportMasterSheet()
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
        '�o�̓V�[�g(�}�X�^�[)�ɏ����o��
        Sheets(g_masterSheet).Range(Sheets(g_masterSheet).Cells(i, 1), Sheets(g_masterSheet).Cells(i, j)).Value = splitcsvData
        i = i + 1
    Loop
    
    '�N���[�Y�Ή�
    csvFile.Close
    Set csvFile = Nothing
    Set fso = Nothing

End Sub
Private Sub touchCollerOverTimeCell()
    Dim MaxCol As Integer '�ő�s
    Dim MaxRow As Integer '�ő��
    Dim overTimeRow As Integer '�c�Ǝ��ԃJ�����ԍ�
    Dim employerCordRow As Integer '�Ј��R�[�h�J�����ԍ�
    Dim ymRow As Integer '���x�J�����ԍ�
    Dim overTime As Date '�c�Ǝ���
    Dim employerCord As Integer '�Ј��R�[�h
        
    '�o�͂ɕK�v�ȍŌ���J�������擾
    MaxCol = Sheets(g_masterSheet).Cells(1, Columns.Count).End(xlToLeft).Column
    MaxRow = Sheets(g_masterSheet).Cells(Rows.Count, 1).End(xlUp).Row
    '�c�Ǝ��ԂƎЈ��R�[�h�̓���������擾
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
    
    '�Ǘ��Җ��ꗗ�̓ǂݍ���
    Dim managerNameList() As Variant
    managerNameList = Array("sakai", "yoshiike")
    Dim mName As Variant
    '�Ǘ��҂��ƂɃV�[�g������ăe���v���쐬
    For Each mName In managerNameList
        With Worksheets.Add(after:=Worksheets(Worksheets.Count))
            .Name = mName
        End With
        Sheets(mName).Range(Sheets(mName).Cells(1, 1), Sheets(mName).Cells(1, MaxCol)).Value = _
            Sheets(g_masterSheet).Range(Sheets(g_masterSheet).Cells(1, 1), Sheets(g_masterSheet).Cells(1, MaxCol)).Value
    Next
    
    '���������ǂݍ���
    Dim ym As String
    Dim empCord As Integer
    Dim isDateCondition As Boolean
    Dim l As Integer
    Dim s As Integer
    Dim y As Integer
    s = 2
    y = 2

    ym = Sheets("���̓t�H�[��").Range("H3").Value
    empCord = Sheets("���̓t�H�[��").Range("H5").Value
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
    
    '�Ј��R�[�h�ꗗ�̓ǂݍ���
    Dim employeeCordList(0 To 1) As Variant
    employeeCordList(0) = Array(44, 48, 52, 58, 66, 137, 149, 151, 167, 203, 227, 270, 297)
    employeeCordList(1) = Array(8, 314, 343, 355, 357, 365, 368, 373, 382, 384, 396, 401, 408)

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
        
       '�ΑӊǗ��҂��ƂɃV�[�g���쐬���ăR�s�[
       employerCord = Sheets(g_masterSheet).Cells(l, employerCordRow).Value
       '�Ƃ肠���������x�^����(in_array���Ȃ��č����Ă���)
       Select Case employerCord
            Case 44, 48, 52, 58, 66, 137, 149, 151, 167, 203, 227, 270, 297
                '���䂳��O���[�v(��)
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
                '�g�r����O���[�v(��)
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
'���������̓ǂݍ���
Private Function loadSearchConditions() As Variant
        
    'loadSearchConditions(0) = Sheets("���̓t�H�[��").Range("H3").Value
    'loadSearchConditions(1) = Sheets("���̓t�H�[��").Range("H4").Value
    
End Function
'---------�I���{�^��--------
Private Sub exitApp_btn()
    Application.Quit
End Sub
 '--------�Ǘ��҃V�[�g�폜�{�^��
Private Sub deleteSheet_btn()
     '�Ǘ��Җ��ꗗ�̓ǂݍ���
    'Dim managerNameList As Variant
    'managerNameList = readManagerNameList()
    managerNameList = Array("sakai", "yoshiike")
End Sub
'--------�Ǘ��҈ꗗ
Private Function readManagerNameList() As Variant
    '�Ǘ��҃��X�g
    Dim managerNameList() As Variant
    managerNameList = Array("sakai", "yoshiike")
End Function
'--------�Ј��R�[�h�ꗗ
Private Function readEmployeeCordList()
    'Dim employeeCordList(0 To 1) As Variant
    'employeeCordList(0) = Array(44, 48, 52, 58, 66, 137, 149, 151, 167, 203, 227, 270, 297)
    'employeeCordList(1) = Array(8, 314, 343, 355, 357, 365, 368, 373, 382, 384, 396, 401, 408)
End Function
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

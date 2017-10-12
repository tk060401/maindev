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
        
    '�ǂݍ���
    Set csvFile = fso.OpenTextFile(Sheet1.Range("A2").Value, 1)
    '---------------���ǂݍ��ݑΏۃt�@�C���ׂ������Ȃ̂Ńe�L�X�g�{�b�N�X�H����ꏊ��ǂނ悤�ɒ���_1011_1100-------------
    
    '�ǂ�csv�t�@�C���̍Ō�̍s�܂œǂݍ���
    i = 1
    Do While csvFile.AtEndOfStream = False
        csvData = csvFile.ReadLine
        
        csvData = Replace(csvData, """", "")
        splitcsvData = Split(csvData, ",")
        j = UBound(splitcsvData) + 1
        
        '�t�@�C������낤
        Sheet2.Range(Sheet2.Cells(i, 1), Sheet2.Cells(i, j)).Value = splitcsvData
        i = i + 1
    Loop

    '�c�Ǝ��ԐF�t�����\�b�h�̌Ăяo��
    Call overTimeColoring
    
    '�G�N�Z���t�@�C������������

    csvFile.Close
    Set csvFile = Nothing
    Set fso = Nothing
    
    '�Ō��ԍ���ɑI���Z����߂�
    Range("A1").Select

End Sub

'�c�Ǝ��Ԃ̋K�͂ɉ����Ĕw�i�F������
Private Sub overTimeColoring()
    Dim MaxCol As Integer '�ő�s
    Dim MaxRow As Integer '�ő��
    Dim overTimeRow As Integer '�c�Ǝ��Ԃ̓�������
    Dim overTime As Date '�c�Ǝ���
    'Dim paintTargetCell As Object '�F������Z��
    
    Dim k As Integer
    Dim l As Integer

    '�c�Ǝ��Ԃ̓������s��T��
    MaxCol = Sheet2.Cells(1, Columns.Count).End(xlToLeft).Column
    For k = 1 To MaxCol
        If Sheet2.Cells(1, k).Value = "�c�Ǝ���" Then
            overTimeRow = k
        End If
    Next k

    MaxRow = Sheet2.Cells(Rows.Count, 1).End(xlUp).Row
    '1�s�ڂ͌��o��,�c�Ǝ��ԃf�[�^��2�s�ڂ���Ō���܂�
    For l = 2 To MaxRow
        overTime = CDate(Sheet2.Cells(l, overTimeRow).Value)
        'paintTargetCell = Cells(l, overTimeRow)
        
       '�c�Ǝ��Ԃ̏����ɉ����ĐF������(�F�Ǝ��Ԃ�萔������H)
        If overTime >= "3:00:00" Then
            Sheet2.Cells(l, overTimeRow).Interior.Color = RGB(226, 43, 48)
        ElseIf overTime >= "2:00:00" Then
            Sheet2.Cells(l, overTimeRow).Interior.Color = RGB(182, 59, 64)
        ElseIf overTime >= "1:00:00" Then
            Sheet2.Cells(l, overTimeRow).Interior.Color = RGB(233, 115, 155)
        Else
            '1���Ԗ����Ȃ�F�t���Ȃ�
        End If
    Next l
End Sub

'---------�I���{�^��--------
Public Sub exitAppButton()
    Application.Quit
 End Sub
 
 '---------�Q�ƃ{�^��--------
Private Sub choiceFileBtn()
    Dim objFS           As New FileSystemObject
    Dim strPath         As String
    Dim strFile         As String
    Dim strFolder       As String
    Dim ofdFileDlg    As Office.FileDialog

    strPath = Sheet1.Range("A2").Value

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
                    strFolder = ThisWorkbook.Path
                End If
            End If
        End If
        Set objFS = Nothing
    Else
        strFolder = ThisWorkbook.Path
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
        Sheet1.Range("A2").Value = ofdFileDlg.SelectedItems(1)
    End If

    Set ofdFileDlg = Nothing
    Exit Sub
End Sub

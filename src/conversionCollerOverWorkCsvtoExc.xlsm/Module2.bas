Attribute VB_Name = "Module2"
Option Explicit

'�Q�Ɛ��exc�t�@�C���Ƃ��ăV�[�g�o�͂���{�^��
Sub outputExcFile_btn()
    Dim fso As New Scripting.FileSystemObject
    Dim csvFile As Object
    Dim csvData As String
    Dim splitcsvData As Variant
    Dim i As Integer
    Dim j As Integer
    Dim overWorkCheck As Boolean
        
    '�Q�ƃt�@�C���ǂݍ���
    Set csvFile = fso.OpenTextFile(Sheets("���̓t�H�[��").Range("A2").Value, 1)
    i = 1
    Do While csvFile.AtEndOfStream = False
        'csv�t�@�C���𐮌`���ēǂݍ���
        csvData = Replace(csvFile.ReadLine, """", "")
        splitcsvData = Split(csvData, ",")
        j = UBound(splitcsvData) + 1
        '�o�̓V�[�g�ɏ����o��
        Sheets("�o��").Range(Sheets("�o��").Cells(i, 1), Sheets("�o��").Cells(i, j)).Value = splitcsvData
        i = i + 1
    Loop
    
    '�c�Ǝ��ԐF�t�����\�b�h�̌Ăяo��
    Call overTimeColoring
   
    '�N���[�Y�Ή�
    csvFile.Close
    Set csvFile = Nothing
    Set fso = Nothing
    '�o�̓V�[�g��A1�w���Ԃɂ��ē���I������
    Sheets("�o��").Activate
    Sheets("�o��").Range("A1").Select

End Sub

'�c�Ǝ��Ԃ̋K�͂ɉ����Ċe�Z���ɔw�i�F������
Private Sub overTimeColoring()
    Dim MaxCol As Integer '�ő�s
    Dim MaxRow As Integer '�ő��
    Dim overTimeRow As Integer '�c�Ǝ��ԃJ����
    Dim overTime As Date '�c�Ǝ���
    Dim k As Integer
    Dim l As Integer
    
    '�c�Ǝ��ԃJ�����̈ʒu��T��
    MaxCol = Sheets("�o��").Cells(1, Columns.Count).End(xlToLeft).Column
    For k = 1 To MaxCol
        If Sheets("�o��").Cells(1, k).Value = "�c�Ǝ���" Then
            overTimeRow = k
        End If
    Next k
    
    '�c�Ǝ��Ԃ�2�s�ڂ���Ō���Ŏ擾����
    MaxRow = Sheets("�o��").Cells(Rows.Count, 1).End(xlUp).Row
    For l = 2 To MaxRow
        overTime = CDate(Sheets("�o��").Cells(l, overTimeRow).Value)
       '�c�Ǝ��Ԃ̏����ɉ����ĐF������
       With Sheets("�o��").Cells(l, overTimeRow)
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

'---------�I���{�^��--------
Public Sub exitApp_btn()
    Application.Quit
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
        Sheets("���̓t�H�[��").Range("A2").Value = ofdFileDlg.SelectedItems(1)
    End If

    Set ofdFileDlg = Nothing
    Exit Sub
End Sub

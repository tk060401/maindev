Attribute VB_Name = "Module2"
Option Explicit

'�Q�Ɛ��exc�t�@�C���Ƃ��ăV�[�g�o�͂���{�^��
Sub outputExcFile_btn()
    '�Q�ƃt�@�C���ǂݍ��݂ƃ}�X�^�[�����o��
    Call readCsvExportMasterSheet
    '�����o���ꂽ�V�[�g�ɐF������
    Call touchCollerSheet
    
    '�o�̓V�[�g��A1�w���Ԃɂ��ē���I������
    Sheets("�o��").Activate
    Sheets("�o��").Range("A1").Select
End Sub
'�Q�ƃt�@�C���ǂݍ���
Private Sub readCsvExportMasterSheet()
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
        Sheets("�o��").Range(Sheets("�o��").Cells(i, 1), Sheets("�o��").Cells(i, j)).Value = splitcsvData
        i = i + 1
    Loop
    
    '�N���[�Y�Ή�
    csvFile.Close
    Set csvFile = Nothing
    Set fso = Nothing

End Sub
'�c�Ǝ��Ԃ̋K�͂ɉ����Ċe�Z���ɔw�i�F������
Private Sub touchCollerSheet()
    Dim MaxCol As Integer '�ő�s
    Dim MaxRow As Integer '�ő��
    Dim overTimeRow As Integer '�c�Ǝ��ԃJ�����ԍ�
    Dim employerCordRow As Integer '�Ј��R�[�h�J�����ԍ�
    Dim overTime As Date '�c�Ǝ���
    Dim employerCord As Integer '�Ј��R�[�h
    Dim k As Integer
    Dim l As Integer
    Dim s As Integer
    Dim y As Integer
    s = 2
    y = 2
        
    '�o�͂ɕK�v�ȍŌ���J�������擾
    MaxCol = Sheets("�o��").Cells(1, Columns.Count).End(xlToLeft).Column
    MaxRow = Sheets("�o��").Cells(Rows.Count, 1).End(xlUp).Row
    
    '�c�Ǝ��ԂƎЈ��R�[�h�̓���������擾
    For k = 1 To MaxCol
        With Sheets("�o��").Cells(1, k)
            If .Value = "�c�Ǝ���" Then
                overTimeRow = k
            ElseIf .Value = "�Ј��R�[�h" Then
                employerCordRow = k
            End If
        End With
    Next k
    
    '�Ǘ��ҕʃV�[�g�̍쐬(�ׂ������E�E�Ǘ��ҕϐ� as Array �݂����ɂ���������vba�悭�킩��Ȃ�)
    With Worksheets.Add(after:=Worksheets(Worksheets.Count))
        .Name = "sakai"
    End With
    With Worksheets.Add(after:=Worksheets(Worksheets.Count))
       .Name = "yoshiike"
    End With
    
    '1�s�ڂ�����ɂ܂Ƃ߂ďo��(�J�����̌��o������,�Ǘ��҂������邽�тɒǉ����Ȃ���΂Ȃ�Ȃ��̂��炢)
    Sheets("sakai").Range(Sheets("sakai").Cells(1, 1), Sheets("sakai").Cells(1, MaxCol)).Value = Sheets("�o��").Range(Sheets("�o��").Cells(1, 1), Sheets("�o��").Cells(1, MaxCol)).Value
    Sheets("yoshiike").Range(Sheets("yoshiike").Cells(1, 1), Sheets("yoshiike").Cells(1, MaxCol)).Value = Sheets("�o��").Range(Sheets("�o��").Cells(1, 1), Sheets("�o��").Cells(1, MaxCol)).Value

    For l = 2 To MaxRow
        overTime = CDate(Sheets("�o��").Cells(l, overTimeRow).Value)
        With Sheets("�o��").Cells(l, overTimeRow)
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
       employerCord = Sheets("�o��").Cells(l, employerCordRow).Value
       '�Ƃ肠���������x�^����(in_array���Ȃ��č����Ă���)
       Select Case employerCord
            Case 44, 48, 52, 58, 66, 137, 149, 151, 167, 203, 227, 270, 297
                '���䂳��O���[�v(��)
                Sheets("sakai").Range(Sheets("sakai").Cells(s, 1), Sheets("sakai").Cells(s, MaxCol)).Value = Sheets("�o��").Range(Sheets("�o��").Cells(l, 1), Sheets("�o��").Cells(l, MaxCol)).Value
                Sheets("sakai").Cells(s, overTimeRow).Interior.Color = Sheets("�o��").Cells(l, overTimeRow).Interior.Color
                s = s + 1
            Case 8, 314, 343, 355, 357, 365, 368, 373, 382, 384, 396, 401, 408
                '�g�r����O���[�v(��)
                Sheets("yoshiike").Range(Sheets("yoshiike").Cells(y, 1), Sheets("yoshiike").Cells(y, MaxCol)).Value = Sheets("�o��").Range(Sheets("�o��").Cells(l, 1), Sheets("�o��").Cells(l, MaxCol)).Value
                Sheets("yoshiike").Cells(y, overTimeRow).Interior.Color = Sheets("�o��").Cells(l, overTimeRow).Interior.Color
                y = y + 1
            Case Else
       End Select
    Next l
End Sub
'---------�I���{�^��--------
Private Sub exitApp_btn()
    Application.Quit
 End Sub
 '--------�Ǘ��҃V�[�g�̍폜
 Private Sub deleteSheet_btn()
 
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

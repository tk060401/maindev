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
    Set csvFile = fso.OpenTextFile("C:\Users\t.kawano\Desktop\�c�Ƒ�쐬�c�[��\daily_2017-09-01_2017-10-01.csv", 1)
    '---------------���ǂݍ��ݑΏۃt�@�C���ׂ������Ȃ̂Ńe�L�X�g�{�b�N�X�H����ꏊ��ǂނ悤�ɒ���_1011_1100-------------
    
    '�ǂ�csv�t�@�C���̍Ō�̍s�܂œǂݍ���
    i = 1
    Do While csvFile.AtEndOfStream = False
        csvData = csvFile.ReadLine
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
    MaxCol = Cells(1, Columns.Count).End(xlToLeft).Column
    For k = 1 To MaxCol
        If Cells(1, k).Value = "�c�Ǝ���" Then
            overTimeRow = k
        End If
    Next k

    MaxRow = Cells(Rows.Count, 1).End(xlUp).Row
    '1�s�ڂ͌��o��,�c�Ǝ��ԃf�[�^��2�s�ڂ���Ō���܂�
    For l = 2 To MaxRow
        overTime = CDate(Cells(l, overTimeRow).Value)
        'paintTargetCell = Cells(l, overTimeRow)
        
       '�c�Ǝ��Ԃ̏����ɉ����ĐF������(�F�Ǝ��Ԃ�萔������H)
        If overTime >= "3:00:00" Then
            Cells(l, overTimeRow).Interior.Color = RGB(226, 43, 48)
        ElseIf overTime >= "2:00:00" Then
            Cells(l, overTimeRow).Interior.Color = RGB(182, 59, 64)
        ElseIf overTime >= "1:00:00" Then
            Cells(l, overTimeRow).Interior.Color = RGB(233, 115, 155)
        Else
            '1���Ԗ����Ȃ�F�t���Ȃ�
        End If
    Next l
End Sub


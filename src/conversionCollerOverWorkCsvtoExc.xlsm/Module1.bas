Attribute VB_Name = "Module1"
Option Explicit


Sub overWorkColorRank()

'csv�t�@�C����ǂݍ���
Dim buf As String 'buf���ĕϐ����͂������̂Ō�Œu������
'�t�@�C���ꏊ�͂��ƂŃ{�^������ǂݍ���(�Ƃ肠�����x�^����)
Open "C:\Users\t.kawano\Desktop\�c�Ƒ�쐬�c�[��\daily_2017-09-01_2017-10-01.csv" For Input As #1

Do Until EOF(1)
    Line Input #1, buf
    '�ǂݍ��񂾃f�[�^���Z���ɑ������
    Loop
    Close #1

'�c�Ǝ��ԃJ����(AC)�ŁA���Ԃ����݂��镔���̔w�i�F��h��
'�c�Ǝ��ԋK�͂ɂ���Ĕw�i�F��������
'�J�������Ȃ��Ȃ����炨���
'�f���o����

End Sub


'�G�N�Z���ɃJ�������Ԃ����ނ��(����Ȃ��˂���)
Sub insertExcColumn()
    Dim tmp As Variant
    '���J���������x�^�����ǂ��Ԃ񒼂�����
    tmp = Split("�����R�[�h,������,�Ζ��̌n,�Ј��R�[�h,�Ј���,���x,���t,�j��,���,�x�ɐ\��,�U�֐\��,�Ζ����ԕύX�\��,�x���o�ΐ\��,�c�Ɛ\��,�����Ζ��\��,�x���\��,���ސ\��,�����m��,�����m��,�V�t�g,�n��,�I��,�o��,�ގ�,�o��(�ۂ߂Ȃ�),�ގ�(�ۂ߂Ȃ�),���J������,���J������,�c�Ǝ���,�@��x���J������,�[��J������,���Ύ���,�x�e����,������,�H�����v����,���l" _
, ",")
    Range("A1:AJ1").Value = tmp
End Sub

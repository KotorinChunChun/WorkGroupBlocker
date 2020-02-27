Attribute VB_Name = "AppMain"
Rem
Rem @appname WorkGroupBlocker - ��ƃO���[�v�֎~�A�h�C��
Rem
Rem @module AppMain
Rem
Rem @author @KotorinChunChun
Rem
Rem @update
Rem    2020/02/18 : �����
Rem    2020/02/27 : Git���J
Rem
Option Explicit
Option Private Module

Public Const APP_NAME = "��ƃO���[�v�֎~�A�h�C��"
Public Const APP_CREATER = "@KotorinChunChun"
Public Const APP_VERSION = "0.11"
Public Const APP_UPDATE = "2020/02/27"
Public Const APP_URL = "https://www.excel-chunchun.com/entry/work_group_blocker"

Public instBlockMultiSelectSheet As BlockMultiSelectSheet

'--------------------------------------------------
'�A�h�C�����s��
Sub AddinStart()
    MsgBox "���Ȃ��̐g���y��ƃO���[�v�z���犮�S�Ɍ��܂��I�I�I" & vbLf & _
             vbLf & _
            "�����̃V�[�g��I�т��ςȂ��ɂ���" & vbLf & _
            "�y��ƃO���[�v�z�͐�΂ɋ����܂���I�I�I", _
                vbInformation + vbOKOnly, APP_NAME
    Call MonitorStart
End Sub

'�A�h�C���ꎞ��~��
Sub AddinStop()
    Dim item
    For Each item In Array( _
        "���`��ƃO���[�v�������Ⴄ�́`�H", _
        "��ƃO���[�v���ĕ����V�[�g�I���̂��Ƃ���", _
        "�����V�[�g�I�������܂܍�Ƃ���ƁA�f�[�^�󂵂��Ⴄ������H", _
        "�����V�[�g�I�������܂ܕۑ�����ƁA���Ɏg���l���f�[�^�󂵂��Ⴄ������H", _
        "����ł���ƃO���[�v�g�������́H")
        If MsgBox(item, vbExclamation + vbYesNo, APP_NAME) = vbNo Then
            MsgBox "����ˁ`��ƃO���[�v�Ȃ�Ă���Ȃ���ˁ`", vbOKOnly, APP_NAME
            Exit Sub
        End If
    Next
    MsgBox "�����ǂ��Ȃ��Ă��m��Ȃ��񂾂�����I�I�I", vbOKOnly, APP_NAME
    Call MonitorStop
End Sub

'�A�h�C���ݒ�\��
Sub AddinConfig(): Call SettingForm.Show: End Sub

'�A�h�C�����\��
Sub AddinInfo()
    Select Case MsgBox(ThisWorkbook.Name & vbLf & vbLf & _
            "�o�[�W���� : " & APP_VERSION & vbLf & _
            "�X�V���@�@ : " & APP_UPDATE & vbLf & _
            "�J���ҁ@�@ : " & APP_CREATER & vbLf & _
            "���s�p�X�@ : " & ThisWorkbook.Path & vbLf & _
            "���J�y�[�W : " & APP_URL & vbLf & _
            vbLf & _
            "�g������ŐV�ł�T���Ɍ��J�y�[�W���J���܂����H" & _
            "", vbInformation + vbYesNo, "�o�[�W�������")
        Case vbNo
            '
        Case vbYes
            CreateObject("Wscript.Shell").Run APP_URL, 3
    End Select
End Sub

'�A�h�C�����S�I��
Sub AddinEnd(): ThisWorkbook.Close False: End Sub
'--------------------------------------------------

'�Ď��J�n
'Workbook_Open����Ă΂��
Sub MonitorStart(): Set instBlockMultiSelectSheet = New BlockMultiSelectSheet: End Sub

'�Ď���~
Sub MonitorStop(): Set instBlockMultiSelectSheet = Nothing: End Sub

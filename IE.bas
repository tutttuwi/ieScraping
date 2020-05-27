Option Explicit

'��������pSleepAPI
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'�����I�ɍőO�ʂɂ�����
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'�ŏ�������Ă��邩���ׂ�
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
'���̑傫���ɖ߂�API
Private Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'---------------------------------------------------------------
'�֐����FieView
'���e�F�w�肳�ꂽURL��IE�ŕ\������
'����1�FIE�I�u�W�F�N�g
'����2�FURL��
'����3�FIE�I�u�W�F�N�g��\�����邩�ǂ����B����lTrue
'�߂�l�F�^�C���A�E�g�̗L��
'---------------------------------------------------------------
Sub ieView(objIE As InternetExplorer, _
           urlName As String, _
           Optional viewFlg As Boolean = True)

  'IE(InternetExplorer)�̃I�u�W�F�N�g���쐬����
  Set objIE = CreateObject("InternetExplorer.Application")

  'IE(InternetExplorer)��\���E��\��
  objIE.Visible = viewFlg

  '�w�肵��URL�̃y�[�W��\������
  objIE.Navigate urlName
 
 'IE�����S�\�������܂őҋ@
 Call ieCheck(objIE)

End Sub

'---------------------------------------------------------------
'�֐����FieCheck
'���e�FIE��Busy��Ԃ����������܂őҋ@
'����1�FIE�I�u�W�F�N�g
'����2�F�^�C���A�E�g�̎��ԁi�I�v�V�����j
'����3�FSleep�֐��ŋx�ގ��ԁi�I�v�V�����j
'�߂�l�F�^�C���A�E�g�̗L��
'---------------------------------------------------------------
Function ieCheck(ByVal objIE As Object, Optional ByVal timeout As String = "0:00:00", Optional ByVal breaktime As Long = 100) As Long
    Dim flg As Boolean
    Dim setTime As Date
    flg = False
    If CDate(timeout) > CDate("0:00:00") Then
        flg = True
        setTime = Now + CDate(timeout)
    End If
    Do While objIE.Busy Or objIE.ReadyState <> 4
        If flg Then
            If Now >= setTime Then
                ieCheck = 1
                Exit Function
            End If
        End If
        Sleep breaktime
        DoEvents
    Loop
    Do While objIE.Document.ReadyState <> "complete"
        If flg Then
            If Now >= setTime Then
                ieCheck = 1
                Exit Function
            End If
        End If
        Sleep breaktime
        DoEvents
    Loop
    ieCheck = 0
End Function

'---------------------------------------------------------------
'�֐����FgetObjIE
'���e�F�w�肳�ꂽ��������܂ރ^�C�g����URL��IE�I�u�W�F�N�g���擾
'����1�F����������
'�߂�l�FIE�I�u�W�F�N�g
'---------------------------------------------------------------
Function getObjIE(Key)
    Dim KeyWord, ie, Reg
    Set ie = Nothing
    Set Reg = CreateObject("VBScript.RegExp")
    Reg.Pattern = ".*" & Key & ".*"
    On Error Resume Next
    For Each obj In CreateObject("Shell.Application").Windows
       If TypeName(obj.Document) = "HTMLDocument" Then
            If Reg.test(obj.LocationName) Or Reg.test(obj.LocationURL) Then
                Set ie = obj
            End If
        End If
    Next
    On Error GoTo 0
    Set Reg = Nothing
    If ie Is Nothing Then
        MsgBox "�w���ie��������܂���ł����B"
    Else
        Set getObjIE = ie
    End If
End Function

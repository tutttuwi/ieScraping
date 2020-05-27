'==========================================================================================================================================
' File   : IE�X�N���C�s���O �e���v���[�g
' Author : T.tsutsui
' Date   : 2020/05/27
' Purpose: IE�X�N���C�s���O�c�[�����J������ۂ̃e���v���[�g���`
'==========================================================================================================================================
Option Explicit

'��������pSleepAPI
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'�����I�ɍőO�ʂɂ�����
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'�ŏ�������Ă��邩���ׂ�
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
'���̑傫���ɖ߂�API
Private Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'----------------------
'      �萔��`
'----------------------
Const TOP_LEFT_RANGE_NAME As String = "C3" ' �������ɁA�E��A�����̃����W���擾���邽�߂Ɏg�p

' �e�[�u���w�b�_�[�ꗗ
Const NO_RANGE_NAME As String = "C3"
Const DATE_RANGE_NAME As String = "D3"
Const PERSON_RANGE_NAME As String = "E3"
Const TYPE_RANGE_NAME As String = "F3"
Const TITLE_RANGE_NAME As String = "G3"

' �V�[�g��
Const TARGET_SHEET_NAME As String = "�ꗗ"
Const SETTINGS_SHEET_NAME As String = "�ݒ�"

'----------------------
'   �������ϐ���`
'----------------------
Dim TOP_RIGHT_RANGE As range
Dim BOTTOM_RIGHT_RANGE As range

Dim NO_RANGE As range
Dim DATE_RANGE As range
Dim PERSON_RANGE As range
Dim TYPE_RANGE As range
Dim TITLE_RANGE As range

Dim NAVIGATE_URL As String ' ����A�N�Z�X������URL�ݒ�p



'==========================================================================================================================================
' Method : �又��
' Author : T.tsutsui
' Date   : YYYY/MM/DD
' Purpose: ���C������
'==========================================================================================================================================
Sub main()
    DoEvents ' ���s���G�N�Z��������\�ɂ���
    Call init
    On Error GoTo checkError
    'Call requredCheck
    Call service
    GoTo EndProc
checkError:
    '���������G���[��No.�Ɠ��e�����b�Z�[�W�{�b�N�X�ŕ\��
    MsgBox "�G���[No.�F" & Err.Number & vbCrLf _
    & "�G���[���e�F" & Err.Description, vbCritical, _
    "[error message]"
    Exit Sub
EndProc:
    Call finally
End Sub

'==========================================================================================================================================
' Method : ��������
' Author : T.tsutsui
' Date   : YYYY/MM/DD
' Purpose: �������������{����
'==========================================================================================================================================
Sub init()
       
    ' �ݒ�l�擾
'    With Worksheets("�ݒ�")
'        targetDirectryStr = .range(TARGET_DIRECTRY_RNG).value
'        targetExtensionStr = .range(TARGET_EXTENSION_RNG).value
'        targetIndexFlg = .range(TARGET_INDEX_RNG).value
'        targetPasswardFlg = .range(TARGET_PASSWARD_RNG).value
'    End With
    
    ' ���L�ϐ�������
    Set NO_RANGE = Worksheets(TARGET_SHEET_NAME).range(NO_RANGE_NAME)
    Dim topLeftRangeBottom As range
    Dim topLeftRangeRight As range
    Set topLeftRangeBottom = Worksheets(TARGET_SHEET_NAME).range(TOP_LEFT_RANGE_NAME).Offset(1, 0)
    Set topLeftRangeRight = Worksheets(TARGET_SHEET_NAME).range(TOP_LEFT_RANGE_NAME).Offset(0, 1)
    Debug.Print topLeftRangeBottom
    If topLeftRangeBottom <> "" Then
        Set BOTTOM_RIGHT_RANGE = Worksheets(TARGET_SHEET_NAME).range(TOP_LEFT_RANGE_NAME).End(xlDown)
    Else
        Set BOTTOM_RIGHT_RANGE = Worksheets(TARGET_SHEET_NAME).range(TOP_LEFT_RANGE_NAME)
    End If
    If topLeftRangeRight <> "" Then
        Set TOP_RIGHT_RANGE = Worksheets(TARGET_SHEET_NAME).range(TOP_LEFT_RANGE_NAME).End(xlToRight)
    Else
        Set TOP_RIGHT_RANGE = Worksheets(TARGET_SHEET_NAME).range(TOP_LEFT_RANGE_NAME)
    End If
    Stop
End Sub

'==========================================================================================================================================
' Method : �K�{�`�F�b�N
' Author : T.tsutsui
' Date   : 2018/08/28
' Purpose: �ݒ�l�̕K�{���ڂ��`�F�b�N����
'==========================================================================================================================================
Sub requredCheck()
    Const ERROR_TITLE As String = "�K�{���ڃG���["
    
    ' �ݒ荀�ڂ�����ꍇ�K�{�`�F�b�N
    If targetDirectryStr = Null Or targetDirectryStr = "" Then
        'MsgBox "�����Ώۂ̃f�B���N�g�����w�肳��Ă��܂���B", vbCritical, ERROR_TITLE
        '�w�蕶���񂪖����ꍇ�A�G���[�𔭐�������
        '���[�U�[��`�G���[�ԍ���513�`65535���Ŏw��
        Err.Raise Number:=9999, Description:=ERROR_TITLE
        Exit Sub
    End If
    
    ' ���̓{�b�N�X����擾����ꍇ
    NAVIGATE_URL = InputBox("���擾��URL���w�肵�Ă���������", "URL�w��", Worksheets("�\��").range("O5"))
    
    If NAVIGATE_URL = "" Or Not (Left(NAVIGATE_URL, 8) = "https://" Or Left(NAVIGATE_URL, 7) = "http://") Then
        MsgBox "URL������Ɍ�肪����܂��B", vbExclamation
        Err.Raise Number:=9999, Description:=ERROR_TITLE
        Exit Sub
    End If
    
End Sub

'==========================================================================================================================================
' Method : �T�[�r�X����
' Author : T.tsutsui
' Date   : 2018/08/28
' Purpose: �T�[�r�X���W�b�N���L�ڂ���
'==========================================================================================================================================
Sub service()
    
    '==============================================
    'IE�I�u�W�F�N�g�̐ݒ�A�w��y�[�W���J��
    '==============================================
    Dim objIE As InternetExplorer
    Set objIE = CreateObject("InternetExplorer.application")
    NAVIGATE_URL = "https://google.com"
    Call ie.ieView(objIE, NAVIGATE_URL)
    objIE.Quit
    Set objIE = Nothing
    U.DBG "�f�o�b�O�֐����s"
End Sub

'==========================================================================================================================================
' Method : ���ߏ���
' Author : T.tsutsui
' Date   : YYYY/MM/DD
' Purpose: ���C������
'==========================================================================================================================================
Sub finally()

End Sub


Sub snipet()
    'Dim dic
    'Set dic = CreateObject("Scripting.Dictionary")  ' �z����g�p�������ꍇ
    
    'Dim appWord, objDoc
    'Set appWord = CreateObject("Word.Application")  'Word�A�v���P�[�V�����̋N��
    'Set objDoc = appWord.Documents.Add              '�V�K�����I�u�W�F�N�g�̍쐬
    
    ' FOR��
    '
    Dim value As Variant
    Dim vlaues As Variant
    vlaues = Array(1, 2, 3, 4, 5)
    For Each value In vlaues
        Debug.Print value
    Next value
    
    ' �Q�s�ɕ�����Ƃy�̕����ɓǂݍ��܂��
    For Each value In range("C3:I4")
        Debug.Print value
    Next value
    
    Dim csvStr As String: csvStr = "1,2,3,4,5,6,7,8,9,10"
    For Each value In Split(csvStr, ",")
        Debug.Print value
    Next value
    
    Dim i As Long
    For i = 0 To 10
    Next i
    
    Dim ubStr As Variant: ubStr = Split(csvStr, ",")
    For i = 0 To UBound(Split(csvStr, ","))
        Debug.Print "ubStr(i) : " & ubStr(i)
    Next i
    
    '�e�[�u���̒[���擾����
    Dim MaxRow, MaxCol As Long
    MaxRow = Worksheets("�ꗗ").range("B8").End(xlDown).Row
    MaxCol = Worksheets("�ꗗ").range("B5").End(xlToRight).Column
    Debug.Print "MaxRow : " & MaxRow & " MaxCol :" & MaxCol
    
    ' Contains�̕\��
    Dim conStr As String: conStr = "��������������������"
    If InStr(conStr, "����") > 0 Then
        Debug.Print "Hit Contain Text!!"
    End If
    
    ' ���K�\����v ( *,?,#)
    If "abc123" Like "*c1*#" Then
        Debug.Print "Hit Regex Text"
    End If
    
    Dim cl As range
    Set cl = Cells(1, 1)
    Debug.Print Cells.Cells
    Debug.Print Cells(1, 1)
    Debug.Print Cells(4, 4)
    
    ' �����W�N���A
    Worksheets("�ꗗ").range("B5:C10000").ClearContents

    ' ��ʍX�V�}�~���ď������x�グ��
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True

End Sub

Sub snippetDoWhile()
    Do
    Loop While 1 = 1 And 2 = 2
End Sub

Sub snippetIE()
    ' IE����
    Dim objIE As InternetExplorer
    Set objIE = CreateObject("InternetExplorer.application")
    NAVIGATE_URL = "https://google.com"
    Call ie.ieView(objIE, NAVIGATE_URL)
    objIE.Document.Script.setTimeout "javascript:alert('���M�{�^����������܂���')", 1000
    objIE.Quit
    Set objIE = Nothing

End Sub


Sub �X�e�[�^�X�o�[()
    Dim i As Integer
    For i = 0 To 10
        Application.StatusBar = "�v���V�[�W���[������" & String(i, "��") & _
            String(10 - i, "��")
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    Application.StatusBar = False
End Sub

'TODO:
' FSO�֘A �t�@�C��/�t�H���_����(�쐬/�폜/�ǋL/�I��)
' �Z������ �t�H���g �w�i �r��

' MISC:
' �t�@�C���W��
' �ڎ��쐬







'==========================================================================================================================================
' File   : UTIL �v���V�[�W��
' Author : T.tsutsui
' Date   : 2020/05/27
' Purpose: �ėp�֐����`
'==========================================================================================================================================
Option Explicit

'---------------------------------------------------------------------------------------
' Method : �f�o�b�O�֐�
' Author : T.tsutsui
' Date   : 2018/08/28
' Purpose: �f�o�b�O�p�̏������L�q
'---------------------------------------------------------------------------------------
Sub DBG(value As String)
    Debug.Print Date; Time(); " : " + value
End Sub


'---------------------------------------------------------------------------------------
' Method : �I�u�W�F�N�g�폜
' Author : T.tsutsui
' Date   :
' Purpose:
'---------------------------------------------------------------------------------------
Sub �I�u�W�F�N�g�폜()
    Dim delOb As Object
    For Each delOb In Worksheet("MAIN").Object
    
End Sub

'---------------------------------------------------------------------------------------
' Method :
' Author :
' Date   :
' Purpose:
'---------------------------------------------------------------------------------------
Sub �}�`���񂳂��A�N�e�B�u�V�[�g��̐}�`���ꊇ�폜����()
  ActiveSheet.Shapes.SelectAll
  Selection.Delete

End Sub

'---------------------------------------------------------------------------------------
' Method :
' Author :
' Date   :
' Purpose:
'---------------------------------------------------------------------------------------
Sub �A�N�e�B�u�u�b�N�̑S���[�N�V�[�g�̐}�`���ꊇ�폜����()
  Dim ws As Worksheet
  Dim shp As Shape

  For Each ws In ActiveWorkbook.Worksheets

    For Each shp In ws.Shapes
      shp.Delete
    Next shp

  Next ws

End Sub


'---------------------------------------------------------------------------------------
' Method :
' Author :
' Date   :
' Purpose:
'---------------------------------------------------------------------------------------
Sub �A�N�e�B�u�V�[�g��̐}�`���ꊇ�폜����()
  Dim shp As Shape

  For Each shp In ActiveSheet.Shapes
    shp.Delete
  Next shp

End Sub





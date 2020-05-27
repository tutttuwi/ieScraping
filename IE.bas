Option Explicit

'処理安定用SleepAPI
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'強制的に最前面にさせる
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'最小化されているか調べる
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
'元の大きさに戻すAPI
Private Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'---------------------------------------------------------------
'関数名：ieView
'内容：指定されたURLをIEで表示する
'引数1：IEオブジェクト
'引数2：URL名
'引数3：IEオブジェクトを表示するかどうか。既定値True
'戻り値：タイムアウトの有無
'---------------------------------------------------------------
Sub ieView(objIE As InternetExplorer, _
           urlName As String, _
           Optional viewFlg As Boolean = True)

  'IE(InternetExplorer)のオブジェクトを作成する
  Set objIE = CreateObject("InternetExplorer.Application")

  'IE(InternetExplorer)を表示・非表示
  objIE.Visible = viewFlg

  '指定したURLのページを表示する
  objIE.Navigate urlName
 
 'IEが完全表示されるまで待機
 Call ieCheck(objIE)

End Sub

'---------------------------------------------------------------
'関数名：ieCheck
'内容：IEのBusy状態が解除されるまで待機
'引数1：IEオブジェクト
'引数2：タイムアウトの時間（オプション）
'引数3：Sleep関数で休む時間（オプション）
'戻り値：タイムアウトの有無
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
'関数名：getObjIE
'内容：指定された文字列を含むタイトルかURLのIEオブジェクトを取得
'引数1：検索文字列
'戻り値：IEオブジェクト
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
        MsgBox "指定のieが見つかりませんでした。"
    Else
        Set getObjIE = ie
    End If
End Function

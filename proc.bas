'==========================================================================================================================================
' File   : IEスクレイピング テンプレート
' Author : T.tsutsui
' Date   : 2020/05/27
' Purpose: IEスクレイピングツールを開発する際のテンプレートを定義
'==========================================================================================================================================
Option Explicit

'処理安定用SleepAPI
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'強制的に最前面にさせる
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'最小化されているか調べる
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
'元の大きさに戻すAPI
Private Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'----------------------
'      定数定義
'----------------------
Const TOP_LEFT_RANGE_NAME As String = "C3" ' 左上を基準に、右上、左下のレンジを取得するために使用

' テーブルヘッダー一覧
Const NO_RANGE_NAME As String = "C3"
Const DATE_RANGE_NAME As String = "D3"
Const PERSON_RANGE_NAME As String = "E3"
Const TYPE_RANGE_NAME As String = "F3"
Const TITLE_RANGE_NAME As String = "G3"

' シート名
Const TARGET_SHEET_NAME As String = "一覧"
Const SETTINGS_SHEET_NAME As String = "設定"

'----------------------
'   初期化変数定義
'----------------------
Dim TOP_RIGHT_RANGE As range
Dim BOTTOM_RIGHT_RANGE As range

Dim NO_RANGE As range
Dim DATE_RANGE As range
Dim PERSON_RANGE As range
Dim TYPE_RANGE As range
Dim TITLE_RANGE As range

Dim NAVIGATE_URL As String ' 初回アクセスしたいURL設定用



'==========================================================================================================================================
' Method : 主処理
' Author : T.tsutsui
' Date   : YYYY/MM/DD
' Purpose: メイン処理
'==========================================================================================================================================
Sub main()
    DoEvents ' 実行中エクセル操作を可能にする
    Call init
    On Error GoTo checkError
    'Call requredCheck
    Call service
    GoTo EndProc
checkError:
    '発生したエラーのNo.と内容をメッセージボックスで表示
    MsgBox "エラーNo.：" & Err.Number & vbCrLf _
    & "エラー内容：" & Err.Description, vbCritical, _
    "[error message]"
    Exit Sub
EndProc:
    Call finally
End Sub

'==========================================================================================================================================
' Method : 初期処理
' Author : T.tsutsui
' Date   : YYYY/MM/DD
' Purpose: 初期処理を実施する
'==========================================================================================================================================
Sub init()
       
    ' 設定値取得
'    With Worksheets("設定")
'        targetDirectryStr = .range(TARGET_DIRECTRY_RNG).value
'        targetExtensionStr = .range(TARGET_EXTENSION_RNG).value
'        targetIndexFlg = .range(TARGET_INDEX_RNG).value
'        targetPasswardFlg = .range(TARGET_PASSWARD_RNG).value
'    End With
    
    ' 共有変数初期化
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
' Method : 必須チェック
' Author : T.tsutsui
' Date   : 2018/08/28
' Purpose: 設定値の必須項目をチェックする
'==========================================================================================================================================
Sub requredCheck()
    Const ERROR_TITLE As String = "必須項目エラー"
    
    ' 設定項目がある場合必須チェック
    If targetDirectryStr = Null Or targetDirectryStr = "" Then
        'MsgBox "検索対象のディレクトリが指定されていません。", vbCritical, ERROR_TITLE
        '指定文字列が無い場合、エラーを発生させる
        'ユーザー定義エラー番号は513～65535内で指定
        Err.Raise Number:=9999, Description:=ERROR_TITLE
        Exit Sub
    End If
    
    ' 入力ボックスから取得する場合
    NAVIGATE_URL = InputBox("★取得先URLを指定してください★", "URL指定", Worksheets("表紙").range("O5"))
    
    If NAVIGATE_URL = "" Or Not (Left(NAVIGATE_URL, 8) = "https://" Or Left(NAVIGATE_URL, 7) = "http://") Then
        MsgBox "URL文字列に誤りがあります。", vbExclamation
        Err.Raise Number:=9999, Description:=ERROR_TITLE
        Exit Sub
    End If
    
End Sub

'==========================================================================================================================================
' Method : サービス処理
' Author : T.tsutsui
' Date   : 2018/08/28
' Purpose: サービスロジックを記載する
'==========================================================================================================================================
Sub service()
    
    '==============================================
    'IEオブジェクトの設定、指定ページを開く
    '==============================================
    Dim objIE As InternetExplorer
    Set objIE = CreateObject("InternetExplorer.application")
    NAVIGATE_URL = "https://google.com"
    Call ie.ieView(objIE, NAVIGATE_URL)
    objIE.Quit
    Set objIE = Nothing
    U.DBG "デバッグ関数実行"
End Sub

'==========================================================================================================================================
' Method : 締め処理
' Author : T.tsutsui
' Date   : YYYY/MM/DD
' Purpose: メイン処理
'==========================================================================================================================================
Sub finally()

End Sub


Sub snipet()
    'Dim dic
    'Set dic = CreateObject("Scripting.Dictionary")  ' 配列を使用したい場合
    
    'Dim appWord, objDoc
    'Set appWord = CreateObject("Word.Application")  'Wordアプリケーションの起動
    'Set objDoc = appWord.Documents.Add              '新規文書オブジェクトの作成
    
    ' FOR文
    '
    Dim value As Variant
    Dim vlaues As Variant
    vlaues = Array(1, 2, 3, 4, 5)
    For Each value In vlaues
        Debug.Print value
    Next value
    
    ' ２行に分けるとＺの方向に読み込まれる
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
    
    'テーブルの端を取得する
    Dim MaxRow, MaxCol As Long
    MaxRow = Worksheets("一覧").range("B8").End(xlDown).Row
    MaxCol = Worksheets("一覧").range("B5").End(xlToRight).Column
    Debug.Print "MaxRow : " & MaxRow & " MaxCol :" & MaxCol
    
    ' Containsの表現
    Dim conStr As String: conStr = "あいうえおかきくけこ"
    If InStr(conStr, "おか") > 0 Then
        Debug.Print "Hit Contain Text!!"
    End If
    
    ' 正規表現一致 ( *,?,#)
    If "abc123" Like "*c1*#" Then
        Debug.Print "Hit Regex Text"
    End If
    
    Dim cl As range
    Set cl = Cells(1, 1)
    Debug.Print Cells.Cells
    Debug.Print Cells(1, 1)
    Debug.Print Cells(4, 4)
    
    ' レンジクリア
    Worksheets("一覧").range("B5:C10000").ClearContents

    ' 画面更新抑止して処理速度上げる
    Application.ScreenUpdating = False
    Application.ScreenUpdating = True

End Sub

Sub snippetDoWhile()
    Do
    Loop While 1 = 1 And 2 = 2
End Sub

Sub snippetIE()
    ' IE操作
    Dim objIE As InternetExplorer
    Set objIE = CreateObject("InternetExplorer.application")
    NAVIGATE_URL = "https://google.com"
    Call ie.ieView(objIE, NAVIGATE_URL)
    objIE.Document.Script.setTimeout "javascript:alert('送信ボタンが押されました')", 1000
    objIE.Quit
    Set objIE = Nothing

End Sub


Sub ステータスバー()
    Dim i As Integer
    For i = 0 To 10
        Application.StatusBar = "プロシージャー処理中" & String(i, "■") & _
            String(10 - i, "□")
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    Application.StatusBar = False
End Sub

'TODO:
' FSO関連 ファイル/フォルダ操作(作成/削除/追記/選択)
' セル操作 フォント 背景 罫線

' MISC:
' ファイル集約
' 目次作成







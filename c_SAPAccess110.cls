VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_SAPAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ###########################################################
' # ---------------------------------------------------------
' # === SAP ACCESS Common Class ===
' #
' # Version : 20181031 Upd tranSQ01Bat スクリプト変更, エラーメッセージ変更
' # Version : 20181030 Add setParaVal パラメータ列内に\マークがあったら、パラメータとはしない処理を追加。
' # Version : 20181030 Add tranSQ01Bat SQ01からトランザクションを選択するまでのまとめ処理を追加
' # Version : 20181029 Upd LogonSAP 多重ログオン時に画面戻り処理の追加。２回以上多重ログオンするとその他エラーになる為
' # Version : 20181029 Upd ckParaVal ファイルパスをチェックする機能を追加
' # Version : 20181029 Add dlMhtmlBat Mhtmlでダウンロードするまとめ処理を追加
' # Version : 20181024 Upd ckParaVal Time Outパターンの検索条件変更。
' # Version : 20181022 Add selVariantBat バリアント選択のまとめ処理を追加。
' # Version : 20181022 Upd ckParaVal　横展開の未入力チェックを追加。
' # Version : 20181019 Upd LogonSAP メッセージコードを修正。L21=>L101に変更。
' # Version : 20181019 Add ckParaVal パラメータのチェックを行う。
' # Version : 20181018 Add ckFileExported2 ファイルが出力されたかどうかまで待つ。
' # Version : 20181017 Upd setParaVal, getParaVal, putParaVal - 横パラメータのブランクに対応。縦パラメータは頭に*を入れる。
' # Version : 20181017 Upd dlXXLBat - saveSAPExcelFile呼び出し後のエラーハンドリングを追加。
' # Version : 20181017 Upd saveSAPExcelFile - 処理の前にDoEvents, WaitSecを入れる。
' # Version : 20181016 Upd closeExcelFile Windows7,10対応
' # Version : 20181015 Upd setParaVal, getParaVal, putParaVal 縦展開の複数パラメータに対応。
' # Version : 20181015 Upd setParaVal, getParaVal 縦展開のパラメータも対応できる
' # Version : 20181015 Add putParaVal パラメータ名で値を設定出来る様にした。
' # Version : 20181012 Add setParaVal, getParaVal パラメータ名で値を取得出来る様にした。
' # Version : 20181011 Add dlXXLBat XXLダウンロード処理のまとめ
' # Version : 20181010 Add closeExcelFile SAPから自動起動されたExcel関連ファイルを閉じる。
' # Version : 20181010 Upd ckFileExported 待ち時間を指定出来る様にした。
' # Version : 20181010 Del ckFileExported2 ckFileExportedに統合
' # Version : 20181005 Upd ckScreenTransition, ckStatusTransition - 1秒待ちをループの最後に持って行った。処理が1秒早くなる。
' # Version : 20181004 Upd cnvExcelFile - DoEventsを追加
' # Version : 20181001 Upd SAP Logon
' # Version : 20180926 Add isLogonSAP, True:SAP Logon, False:SAP Started
' # Version : 20180906 Add insCol
' # Version : 20180904 Upd saveSAPExcelFile - optionalでOriginalのファイル名を指定出来るようにした。
' # Version : 20180903 Add getGridLineNoClick - Clickパターン
' # Version : 20180830 Upd saveSAPExcelFile - On Error処理追加
' # Version : 20180829 Upd LogonSAP - Logon処理変更
' # Version : 20180824 Upd LogonSAP - Logonしている時にも実行可能
' # Version : 20180807 Add saveSAPExcelFile
' # Version : 20180806 Upd cnvExcelFile - Add Delete File Error Code
' # Version : 20180803 Add ckDataSelected
' # Version : 20180731 Upd cnvExcelFile - Error Process
' # Version : 20180730 Add ckFileExported2
' # Version : 20180725 Upd ckScreenTransition - Delete Logoff Process
' # Version : 20180724 Add cnvExcelFile2
' # Version : 20180724 Upd ckStatusTransition
' # Version : 20180724 Add ckWaitingStatus
' # Version : 20180719 Update logonSAP - Log Out
' # Version : 20180719 Update logonSAP - Error Process
' # Version : 20180717 Add Function-getGridLineNo
' # Version : 20180709 Created
' # ---------------------------------------------------------
' ###########################################################
Public SAPSesi As Object
Public SAPGrid As Object
Dim SapGuiApp As Object
Dim oConnection As Object
Dim SAPGUIAuto As Object ' 既にセッションがある時に使用

Dim xConnection As String
Dim xClient As String
Dim xUser As String
Dim xPassword As String
Dim xLanguage As String
Dim xTranCd As String
Dim xIsLogonSAP As Boolean
Dim xSaveFilePath As String

Dim xParaArray() As Variant
Dim xValRowNo() As Variant
Dim xValColNo() As Variant
Dim xParaValShtNm As String

Property Let Connection(inConnection As String)
    xConnection = inConnection
End Property

Property Let Client(inClient As String)
    xClient = inClient
End Property

Property Let User(inUser As String)
    xUser = inUser
End Property

Property Let Password(inPassword As String)
    xPassword = inPassword
End Property

Property Let Language(inLanguage As String)
    xLanguage = inLanguage
End Property

Property Let TranCd(inTranCd As String)
    xTranCd = inTranCd
End Property

Property Get isLogonSAP() As Boolean
    isLogonSAP = xIsLogonSAP
End Property

Property Get SaveFilePath() As String
    SaveFilePath = xSaveFilePath
End Property

' **************************************************
' SAP Logon
' SAPにログオンする。
' Return(0:Normal, 1:SAP Create Object Error, 2:Server Connection Error, 3:Logon Parameter Error)
' **************************************************
Public Function LogonSAP() As Integer
    Dim w_CatchNo As Integer
    Dim w_ErrorText As String

    xIsLogonSAP = False
' 既にSAPにログオンされている時
    On Error Resume Next
    Set SAPGUIAuto = GetObject("SAPGUI")
    If Err.Number <> 0 Then
        On Error GoTo Exception
        GoTo LogonProc
    End If
    Set SapGuiApp = SAPGUIAuto.GetScriptingEngine

    On Error GoTo Exception
' 既にあるログオンされているセッションを取得
    w_CatchNo = 10
    If SapGuiApp.Children.Count = 0 Then
        GoTo LogonProc2
    End If
    Set oConnection = SapGuiApp.Children(0)
    
    w_CatchNo = 11
    Set SAPSesi = oConnection.Children(0)

' ログイン画面が開かれていた時
    w_CatchNo = 12
    If SAPSesi.findById("wnd[0]").text = "SAP" Then
        GoTo LogonSAP
    End If

' SAP Easy Access画面に遷移する。
    w_CatchNo = 13
    SAPSesi.findById("wnd[0]/tbar[0]/okcd").text = "/ns000"
    w_CatchNo = 14
    SAPSesi.findById("wnd[0]").sendVKey 0
    Select Case ckScreenTransition("wnd[0]", "*SAP Easy Access*", get_Timeout(1))
        Case 0
            ' OK : Continue
        Case 1
            w_CatchNo = 15
            GoTo Exception
        Case 2
            w_CatchNo = 16
            GoTo Exception
        Case Else
            w_CatchNo = 17
            GoTo Exception
    End Select
    
    LogonSAP = 0
    Exit Function

LogonProc:
    'SAPのログオンを行う。プログラム内でSAP GUIスクリプティングAPIを使用する｡
    w_CatchNo = 1
    Set SapGuiApp = CreateObject("Sapgui.ScriptingCtrl.1")

LogonProc2:
    ' SAP指定サーバーに接続
    w_CatchNo = 2
    Set oConnection = SapGuiApp.OpenConnection(xConnection, True)
    
    ' ログオン画面表示［起動時第一画面］
    Set SAPSesi = oConnection.Children(0)

    On Error GoTo 0

LogonSAP:
    With SAPSesi
    ' ログオン画面パラメータ設定
    .findById("wnd[0]").maximize                              ' 画面最大
    .findById("wnd[0]/usr/txtRSYST-MANDT").text = xClient     'クライアント設定
    .findById("wnd[0]/usr/txtRSYST-BNAME").text = xUser       'ユーザ設定
    .findById("wnd[0]/usr/pwdRSYST-BCODE").text = xPassword   'パスワード設定
    .findById("wnd[0]/usr/txtRSYST-LANGU").text = xLanguage   '言語設定
    .findById("wnd[0]").sendVKey 0                            'Enter Key
            
    waitSec 2 '1秒待つ
            
    ' ログオン後のチェック
    If .findById("wnd[0]").text Like "*SAP Easy Access*" Then
        outMsg "L101", "Logon User = " & xUser & ", Computer Name = " & getComputerName
        LogonSAP = 0
        xIsLogonSAP = True
        Exit Function
    End If
    
    If .findById("wnd[0]/sbar/pane[0]").text Like "*ユーザ名または*" Then
        w_CatchNo = 3
        GoTo Exception
    End If
            
    If .findById("wnd[0]/sbar/pane[0]").text Like "*必要な入力項目すべて*" Then
        w_CatchNo = 6
        GoTo Exception
    End If

    If .Children.Count > 1 Then
        If .findById("wnd[1]").text Like "*多重ログオン*" Then
            ' 前の画面に戻る。
            .findById("wnd[1]").sendVKey 0
            w_CatchNo = 4
            GoTo Exception
        Else
            w_CatchNo = 7
            ' 前の画面に戻る。
            .findById("wnd[1]").sendVKey 0
            w_CatchNo = 4
            GoTo Exception
        End If
    Else
        If .findById("wnd[0]/sbar/pane[0]").text <> "" Then
             w_ErrorText = .findById("wnd[0]/sbar/pane[0]").text
            w_CatchNo = 5
            GoTo Exception
        Else
            w_CatchNo = 7
            GoTo Exception
        End If
    End If
        
    End With
    
    Exit Function

Exception:
    logoffSAP
        
    Select Case w_CatchNo
        Case 1
            outMsg "E101", "Create Objectに失敗しました。" & vbCrLf & "Sapgui.ScriptingCtrl.1"
        Case 2
            outMsg "E102", "Open Connectionを確立出来ないか、キャンセルされました。"
        Case 3
            outMsg "E103", "ユーザ名またはパスワードが正しくありません。" & vbCrLf & _
            "再度ログインしてください。" & vbCrLf & _
            "▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼" & vbCrLf & _
            "      Client:" & get_Client & vbCrLf & _
            "      User:" & get_User & vbCrLf & _
            "      Language:" & get_Language
        Case 4
             outMsg "E104", "他のPCでログオンされています。" & vbCrLf & _
            "▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼" & vbCrLf & _
            "      Client:" & get_Client & vbCrLf & _
            "      User:" & get_User & vbCrLf
        Case 5
             outMsg "E105", "その他のログオンエラー。" & vbCrLf & w_ErrorText
        Case 6
             outMsg "E106", "ユーザー名/パスワードを入力して下さい。"
        Case 7
             outMsg "E107", "その他のログオンエラー(2)。" & vbCrLf & _
            "▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼" & vbCrLf & _
            "      Client:" & get_Client & vbCrLf & _
            "      User:" & get_User & vbCrLf
        Case 10, 11
            outMsg "E108", "開かれているSAPセッションに接続出来ませんでした。"
        Case 12, 13, 14, 15, 16, 17
            outMsg "E109", "開かれているSAPセッション画面からSAP Easy Access画面に行けませんでした。"
    End Select

    LogonSAP = w_CatchNo
End Function

' **************************************************
' Logoff SAP
' SAP関連Objectをクリア
' Return = 0:Normal
' **************************************************
Public Function logoffSAP() As Integer
    Set SAPSesi = Nothing
    Set oConnection = Nothing
    Set SapGuiApp = Nothing
    Set SAPGUIAuto = Nothing
    
    logoffSAP = 0
End Function

' **************************************************
' Set SAP Transaction Code
' SAPトランザクションを設定する。
' **************************************************
Public Sub setTranCd()
    SAPSesi.findById("wnd[0]/tbar[0]/okcd").text = xTranCd  ' コードセット
    SAPSesi.findById("wnd[0]").sendVKey 0                  ' Enter
    
End Sub

' **************************************************
' Check Screen Transition
' 画面が遷移されたかどうかのチェック
' Return = 0 : Normal, 1 : Timeオーバー
' **************************************************
Public Function ckScreenTransition(inTierarchy As String, inTitle As String, inTimeOutSec As Integer) As Integer
    Dim w_TimCnt As Integer
    Dim w_CheckText As String

    For w_TimCnt = 0 To inTimeOutSec
        ckScreenTransition = 2
        ' 明細結果チェック
        w_CheckText = SAPSesi.findById(inTierarchy).text
        If w_CheckText <> "" Then
            If w_CheckText Like inTitle Then
                ckScreenTransition = 0
            Else
                ckScreenTransition = 1
            End If
            Exit For
        End If
        
        waitSec '1秒待つ
    Next w_TimCnt
End Function

' **************************************************
' Check Status Transition
' 画面のステータスが表示された時のチェック
' Return = 0:Status Matching 1:Status Unmatching 2:Time Over
' **************************************************
Public Function ckStatusTransition(inStatus As String, inTimeOutSec As Integer) As Integer
    Dim w_TimCnt As Integer
    Dim w_CheckText As String

    For w_TimCnt = 0 To inTimeOutSec
        ' 明細結果チェック
        w_CheckText = SAPSesi.findById("wnd[0]/sbar/pane[0]").text
        If "" <> w_CheckText Then
            If w_CheckText Like inStatus Then
                ckStatusTransition = 0
            Else
                ckStatusTransition = 1
            End If
            Exit Function
        End If
        
        waitSec '1秒待つ
    Next w_TimCnt
    
    ckStatusTransition = 2
End Function

' **************************************************
' Check Data Selected
' データが抽出されたかどうかのチェック
' Return = True:正常 False:抽出エラー
' **************************************************
Public Function ckDataSelected(inDataSelectedMsg As String) As Boolean
    Dim w_MessageType As String
    Dim w_MessageStr As String

    w_MessageType = SAPSesi.findById("wnd[0]/sbar").messagetype
    w_MessageStr = SAPSesi.findById("wnd[0]/sbar").text
    
    Select Case w_MessageType
        Case ""
            ckDataSelected = True
        Case "E"
            ckDataSelected = False
        Case "S"
            If inDataSelectedMsg = "" Then
                ckDataSelected = False
            Else
                If w_MessageStr Like inDataSelectedMsg Then
                    ckDataSelected = True
                Else
                    ckDataSelected = False
                End If
            End If
        Case Else
            ckDataSelected = False
    End Select
End Function

' **************************************************
' Check SAP File Exported
' SAPから出力され、自動起動たファイルが使用中になるまで待つ。
' Return = 0:Normal 1:Time Over
' **************************************************
Public Function ckFileExported(inPathNm As String, Optional inWaitSec As Variant) As Integer
    Dim i As Integer
    Dim wkWaitSec As Integer
    
    ckFileExported = 1
    
    If IsMissing(inWaitSec) Then
        wkWaitSec = 600 ' 10分
    Else
        wkWaitSec = CInt(inWaitSec)
    End If
    
    For i = 0 To wkWaitSec
        If Dir(inPathNm) <> "" Then
            ' ファイルが使用中になるまで待つ。SAPで自動でExcelを開く為。
            Dim n As Integer
            n = FreeFile()
            On Error Resume Next
            Open inPathNm For Binary Access Read Lock Read Write As #n
            Close #n
            If Err Then
                ckFileExported = 0
                Exit For
            End If
            On Error GoTo 0
        End If
        waitSec '1秒待つ
        DoEvents
    Next i
End Function

' **************************************************
' Check SAP File Exported
' SAPから出力され、ファイルが出来たかどうかまで待つ。
' SAPから自動起動されて開かれないパターン。
' Return = 0:Normal 1:Time Over
' **************************************************
Public Function ckFileExported2(inPathNm As String, Optional inWaitSec As Variant) As Integer
    Dim i As Integer
    Dim wkWaitSec As Integer
    
    If IsMissing(inWaitSec) Then
        wkWaitSec = 10 ' デフォルト10秒
    Else
        wkWaitSec = CInt(inWaitSec)
    End If
    
    ckFileExported2 = 1
    For i = 0 To wkWaitSec
        If Dir(inPathNm) <> "" Then
            ckFileExported2 = 0
            Exit For
        End If
        waitSec '1秒待つ
        DoEvents
    Next i
End Function

' **************************************************
' Convert SAP File to Excel Standart Type
' SAPから出力されたファイル
' Return = 0:Normal 1:Excel Open Error 2:Excel Copy Error
' ■■ SAPを起動して作成されたスプレッドシートは、別のExcelを起動されたものなので
' ■■ マクロから実行されたExcel範囲外の管理となる。
' ■■ フルパス指定で実行されているアプリケーションのオブジェクトを取得し、
' ■■ そのオブジェクト内で制御を行う。
' ■■ Excel認識待ちが必要です。
' **************************************************
Public Function cnvExcelFile(inPathNm As String, inCnvPathNm As String) As Integer
    Dim xlapp As Application
    Dim WB As Workbook
    Dim w_TimCnt As Integer
    Dim w_RtnNo As Integer
    Dim w_BookCnt As Integer

    DoEvents ' 重要：当Excelの画面がアクティブになっていると、SAPから出力されたファイルをGetObject出来ない。一度、DoEventsでOSに制御を戻す。

    w_RtnNo = 1
    On Error GoTo Exception
    Set xlapp = GetObject(inPathNm).Application
    On Error GoTo 0
    
    For w_TimCnt = 0 To 600
        waitSec '1秒待つ
        If xlapp Is Nothing Then
            'Continue
        Else
            w_RtnNo = 0
            Exit For
        End If
    Next
    
    If w_RtnNo > 0 Then
        cnvExcelFile = w_RtnNo
        Exit Function
    End If
    
    w_RtnNo = 2
    w_BookCnt = 0
    For Each WB In xlapp.Workbooks
        w_BookCnt = w_BookCnt + 1
        If WB.Name = Dir(inPathNm) Then 'コンバート元のファイル名をパスからDir関数で取得する。
            
            ' xlsx形式で保存する
            On Error GoTo Exception
            WB.SaveAs FileName:=inCnvPathNm, FileFormat:=xlWorkbookDefault
            On Error GoTo 0

            For w_TimCnt = 0 To 600
                waitSec '1秒待つ
                If Dir(inCnvPathNm) <> "" Then
                ' ##### ファイル検知 !!!
                    w_RtnNo = 0
                    Exit For
                End If
            Next w_TimCnt

            WB.Close
        End If
    Next
    
    If w_RtnNo > 0 Then ' 1つのExcelのインスタンスで開かれた場合。
        cnvExcelFile = w_RtnNo
        Exit Function
    End If
    
    If w_BookCnt > 1 Then
        ' Continue
    Else
        xlapp.Quit        ' 該当アプリケーションを終了する。
    End If
    
    Set WB = Nothing
    Set xlapp = Nothing
    
    
    ' MHTMLファイル削除
    w_RtnNo = 3
    If Dir(inPathNm) <> "" Then
        On Error GoTo Exception
        Kill inPathNm
        w_RtnNo = 0
    End If

Exception:

    cnvExcelFile = w_RtnNo
End Function

' **************************************************
' エクセルファイル変換２, 当インスタンスで実行
' Return = 0:Normal, 1:Open File Error, 2:XLSX保存エラー, 3:Delete Fileエラー
'
' Excelフォーマットだが、拡張子が無かったり、CSVファイルなどで、
' GetObject().Applicationでは開けないファイルを、
' 当インスタンスから開いて、XLSX形式で保存する。
' **************************************************
Public Function cnvExcelFile2(inPathNm As String, inCnvPathNm As String)
    Dim WB As Workbook
    Dim w_RtnNo As Integer

    On Error GoTo Exception
    w_RtnNo = 1
    Workbooks.Open inPathNm
    
    w_RtnNo = 2
    For Each WB In Workbooks
        If WB.Name = Dir(inPathNm) Then
            WB.SaveAs FileName:=inCnvPathNm, FileFormat:=xlWorkbookDefault
            WB.Close
            Exit For
        End If
    Next
    
    Set WB = Nothing
    
    w_RtnNo = 3
    If Dir(inPathNm) <> "" Then
        Kill inPathNm
    End If
    
    On Error GoTo 0
    
    cnvExcelFile2 = 0
    
    Exit Function
        
Exception:
    cnvExcelFile2 = w_RtnNo
End Function

' **************************************************
' SAPから自動出力されたExcelファイル(デフォルト:XXL)を名前を付けて保存する。
' Return = 0:Normal, 1:ファイルが見つからない。, 2:SAVE AS Error
' ### Windows10専用 ###
' **************************************************
Public Function saveSAPExcelFile(inSavePathNm As String, Optional inOrgFileNm As Variant) As Integer
    Dim WB As Workbook
    Dim w_RtnNo As Integer
    Dim w_OrgFileNm As String

    DoEvents
    waitSec '1秒待つ

    If IsMissing(inOrgFileNm) Then
        w_OrgFileNm = "*XXL*"
    Else
        w_OrgFileNm = inOrgFileNm
    End If

    w_RtnNo = 1 ' Original File Nothing
    For Each WB In Workbooks
        If WB.Name Like w_OrgFileNm Then
            On Error GoTo Exception
            WB.SaveAs FileName:=inSavePathNm, FileFormat:=xlWorkbookDefault
            WB.Close
            On Error GoTo 0
            w_RtnNo = 0
            Exit For
        End If
    Next
    
    Set WB = Nothing
    
    saveSAPExcelFile = w_RtnNo
    
    Exit Function
    
Exception:
    Set WB = Nothing

    saveSAPExcelFile = 2 ' File Save Error
End Function

' **************************************************
' Get Glid Line No, 行選択も同時に行う。
' 一覧情報から該当の内容のラインNoを取得する。
' Return = 0 : 該当無し, 1以上 : ラインNo
' **************************************************
Public Function getGridLineNo(inGridObjStr As String, inFieldNm As String, inTitle As String) As Long
    Dim i As Integer
    Dim hasDataFlg As Boolean
    
    hasDataFlg = False
    
    Set SAPGrid = SAPSesi.findById(inGridObjStr)
    
    For i = 1 To SAPGrid.RowCount
        If SAPGrid.getcellvalue(i - 1, inFieldNm) Like inTitle Then
            SAPSesi.findById(inGridObjStr).selectedRows = CStr(i - 1)
            hasDataFlg = True
            Exit For
        End If
    Next i

    If hasDataFlg Then
        getGridLineNo = i
    Else
        getGridLineNo = 0
    End If
End Function

' **************************************************
' Get Glid Line No, 行クリックも同時に行う。
' 一覧情報から該当の内容のラインNoを取得して、クリックする。
' Return = 0 : 該当無し, 1以上 : ラインNo
' **************************************************
Public Function getGridLineNoClick(inGridObjStr As String, inFieldNm As String, inTitle As String) As Long
    Dim i As Integer
    Dim myLineNo As Integer
    
    myLineNo = getGridLineNo(inGridObjStr, inFieldNm, inTitle)
    
    If myLineNo > 0 Then
        SAPSesi.findById(inGridObjStr).setCurrentCell myLineNo - 1, inFieldNm
        SAPSesi.findById(inGridObjStr).clickCurrentCell
    End If

    getGridLineNoClick = myLineNo
End Function

' **************************************************
' Get Layout Line No, , 行選択も同時に行う。
' 一覧情報から該当の内容のラインNoを取得する。
' Return = 0 : 該当無し, 1以上 : ラインNo
' **************************************************
Public Function getLayoutLineNo(inLayoutObjStr As String, inTitle As String) As Long
    Dim SAPUsr As Object
    Dim ChildCount As Integer
    Dim LastChildId As String
    Dim VisRow As Integer
    Dim TotScrol As Integer
    Dim SelLineNo As Integer
    Dim i As Integer
    
    Set SAPUsr = SAPSesi.findById(inLayoutObjStr)

    SAPUsr.verticalScrollbar.Position = 0

    ChildCount = SAPUsr.Children.Count

    LastChildId = SAPUsr.Children(ChildCount - 1).ID

    VisRow = CLng(Right(Left(LastChildId, Len(LastChildId) - 1), Len(LastChildId) - InStr(1, LastChildId, ",") - 1))

    SelLineNo = 0
    For i = 3 To VisRow
        If SAPSesi.findById(inLayoutObjStr & "/lbl[1," & i & "]").text Like inTitle Then
            SAPSesi.findById(inLayoutObjStr & "/lbl[1," & i & "]").SetFocus
            SelLineNo = i
            Exit For
        End If
    Next i
    
    If 0 = SelLineNo Then
        TotScrol = SAPUsr.verticalScrollbar.Maximum
        If TotScrol > 0 Then
            For i = 1 To TotScrol - 1
                SAPUsr.verticalScrollbar.Position = i
                If SAPSesi.findById(inLayoutObjStr & "/lbl[1," & VisRow & "]").text Like inTitle Then
                    SAPSesi.findById(inLayoutObjStr & "/lbl[1," & VisRow & "]").SetFocus
                    SelLineNo = VisRow
                    Exit For
                End If
            Next i
        End If
    End If

    getLayoutLineNo = SelLineNo
    Debug.Print "Title = " & inTitle & ", Index = " & i
End Function

' **************************************************
' SAPから自動起動されたExcel関連ファイルを閉じる。
' Windows7, 10対応, Windows7では別にExcelが立ち上がる。Winsows10では開いているExcelの中で立ち上がる。
' @In : File Path Name
' @Out 0:Normal, 1:ファイルが見つからない。, 2:Close Error
' **************************************************
Public Function closeExcelFile(inPathNm As Variant) As Integer
    Dim xlapp As Application
    Dim WB As Workbook
    Dim i As Integer
    Dim w_RtnNo As Integer
    Dim w_BookCnt As Integer

    DoEvents ' 重要：当Excelの画面がアクティブになっていると、SAPから出力されたファイルをGetObject出来ない。一度、DoEventsでOSに制御を戻す。

    w_RtnNo = 1
    On Error GoTo Exception
    Set xlapp = GetObject(inPathNm).Application
    On Error GoTo 0
    
    For i = 0 To 600
        If xlapp Is Nothing Then
            'Continue
        Else
            w_RtnNo = 0
            Exit For
        End If
        waitSec '1秒待つ
    Next i
    
    If w_RtnNo > 0 Then
        closeExcelFile = w_RtnNo
        Exit Function
    End If
    
    w_RtnNo = 2
    w_BookCnt = 0
    For Each WB In xlapp.Workbooks
        w_BookCnt = w_BookCnt + 1
        If WB.Name = Dir(inPathNm) Then 'コンバート元のファイル名をパスからDir関数で取得する。
            w_RtnNo = 0
            WB.Close
        End If
    Next
    
    If w_RtnNo > 0 Then ' 1つのExcelのインスタンスで開かれた場合。
        closeExcelFile = w_RtnNo
        Exit Function
    End If
    
    If w_BookCnt > 1 Then
        ' Continue
    Else
        xlapp.Quit ' 該当アプリケーションを終了する。
    End If
    
    Set WB = Nothing
    Set xlapp = Nothing
    
    closeExcelFile = w_RtnNo
    
    Exit Function

Exception:
    Set WB = Nothing
    Set xlapp = Nothing
        
    closeExcelFile = w_RtnNo
End Function

' **************************************************
' 画面のパラメータの内容を配列に保存する。
' @In : シート名
' @In : パラメータの行No
' @In : 値の行No
' **************************************************
Public Sub setParaVal(inShtNm As String, inParaColNo As Integer, inValColNo As Integer)
    Dim WS As Worksheet
    
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim hasDataFlg As Boolean
    
    Set WS = ThisWorkbook.Worksheets(inShtNm)
    
    hasDataFlg = True
    i = 1 ' セル行のカウント, タイトルは除く（1行目)
    j = -1 ' データ格納のカウント
    k = 0 'データ無しのカウント
    Do While hasDataFlg
        i = i + 1
        If WS.Cells(i, inParaColNo) = "" Or InStr(WS.Cells(i, inParaColNo), "\") > 0 Then
            k = k + 1
        Else
            j = j + 1
            ReDim Preserve xParaArray(j)
            ReDim Preserve xValRowNo(j)
            ReDim Preserve xValColNo(j)
            xParaArray(j) = WS.Cells(i, inParaColNo)
            
            If Left(WS.Cells(i, inParaColNo), 1) = "*" Then
                xValRowNo(j) = i + 1
                xValColNo(j) = inParaColNo
            Else
                xValRowNo(j) = i
                xValColNo(j) = inValColNo
            End If
            
            k = 0
'            Debug.Print j & ", " & xParaArray(j) & ", " & xValRowNo(j) & ", " & xValColNo(j)
        End If
                
        If k > 20 Then
            hasDataFlg = False
        End If
    Loop
    
    xParaValShtNm = inShtNm
    
    Set WS = Nothing
End Sub

' **************************************************
' パラメータ名より値を取得する。
' @In : パラメータ名
' @In : パラメータNo
' @Out : 値
' **************************************************
Public Function getParaVal(inParaNm As String, Optional inRowSeqNo As Variant) As String
    Dim WS As Worksheet
    
    Dim i As Integer
    Dim j As Integer
    
    Set WS = ThisWorkbook.Worksheets(xParaValShtNm)
        
    If IsMissing(inRowSeqNo) Then
        j = 0
    Else
        j = CInt(inRowSeqNo) - 1
    End If
    
    For i = 0 To UBound(xParaArray)
        If inParaNm = xParaArray(i) Then
            If Left(xParaArray(i), 1) = "*" Then
                getParaVal = WS.Cells(xValRowNo(i) + j, xValColNo(i))
            Else
                getParaVal = WS.Cells(xValRowNo(i), xValColNo(i) + j)
            End If
            
            Set WS = Nothing
            Exit Function
        End If
    Next i
    
    Set WS = Nothing
        
    getParaVal = ""
    
    Debug.Print "C_SAPAccess.GetParaVal Error = " & inParaNm
End Function

' **************************************************
' パラメータに値を設定する。
' @In : パラメータ名
' @In : パラメータNo
' @Out : 値
' **************************************************
Public Sub putParaVal(inParaNm As String, inVal As String, Optional inRowSeqNo As Variant)
    Dim WS As Worksheet
    
    Dim i As Integer
    Dim j As Integer
        
    Set WS = ThisWorkbook.Worksheets(xParaValShtNm)
    
    If IsMissing(inRowSeqNo) Then
        j = 0
    Else
        j = CInt(inRowSeqNo) - 1
    End If
    
    For i = 0 To UBound(xParaArray)
        If inParaNm = xParaArray(i) Then
            If Left(xParaArray(i), 1) = "*" Then
                WS.Cells(xValRowNo(i) + j, xValColNo(i)) = inVal
            Else
                WS.Cells(xValRowNo(i), xValColNo(i) + j) = inVal
            End If
            Set WS = Nothing
            Exit Sub
        End If
    Next i
    
    Set WS = Nothing
    
    Debug.Print "C_SAPAccess.SetParaVal Error = " & inParaNm
End Sub

' **************************************************
' パラメータのチェックを行う。
' @In : 追加チェック項目（未入力チェック) @マークの後ろに数字を入れると、横展開のチェックを行う。例:バリアント@2
' @In : 追加チェック項目２（未入力、保管場所チェック、最後\チェック)
' @In : 追加チェック項目３（未入力、ファイルパスチェック)
' @Out : True - エラー無し, False - エラーあり
' **************************************************
Public Function ckParaVal(Optional inCkParaNm As Variant, Optional inCkParaNm2 As Variant, Optional inCkParaNm3 As Variant) As Boolean
    Dim WS As Worksheet
    
    Dim ParaNmStr As String
    Dim ParaNmStr2 As String
    Dim ParaNmStr3 As String
    Dim ParaNmArray As Variant
    Dim ParaNmArray2 As Variant
    Dim ParaNmArray3 As Variant
    
    Dim i As Integer
    Dim j As Integer
    
    Dim wkPara As String
    Dim wkVal As String
    
    Dim wkPosNo As Integer
    Dim wkSeqNo As Integer
    Dim wkParaNm As String
    
    Set WS = ThisWorkbook.Worksheets(xParaValShtNm)
    
    ' Log情報の設定
    For i = 4 To 6
        j = UBound(xParaArray) + 1
        ReDim Preserve xParaArray(j)
        ReDim Preserve xValRowNo(j)
        ReDim Preserve xValColNo(j)
        xParaArray(j) = WS.Cells(i, 3).Value
        xValRowNo(j) = i
        xValColNo(j) = 5
    Next i
    
    ' Logon, Time Out情報の設定
    For i = 4 To 20
        wkPara = WS.Cells(i, 8).Value
        If wkPara <> "" Then
            j = UBound(xParaArray) + 1
            ReDim Preserve xParaArray(j)
            ReDim Preserve xValRowNo(j)
            ReDim Preserve xValColNo(j)
            
            xParaArray(j) = wkPara
            xValRowNo(j) = i
            xValColNo(j) = 9
        End If
    Next i

    ParaNmStr = "トランザクション*,バリアント名*,*ファイル名*,*フォーマット*,Operating Type,Log File Name,Connection,User,Client,Password,Language,*パターン*" '未入力チェック項目
    ParaNmStr2 = "*保管場所*,Log Folder" '未入力、保管場所チェック、最後\チェック項目
    ParaNmStr3 = "*ファイルパス*" '未入力、ファイルパスチェック

    If IsMissing(inCkParaNm) = False Then
        ParaNmStr = ParaNmStr & "," & inCkParaNm
    End If
    
    If IsMissing(inCkParaNm2) = False Then
        ParaNmStr2 = ParaNmStr2 & "," & inCkParaNm2
    End If
    
    If IsMissing(inCkParaNm3) = False Then
        ParaNmStr3 = ParaNmStr3 & "," & inCkParaNm3
    End If
    
    ParaNmArray = Split(ParaNmStr, ",")
    ParaNmArray2 = Split(ParaNmStr2, ",")
    ParaNmArray3 = Split(ParaNmStr3, ",")
    
    ' チェック処理
    ckParaVal = True
    
    For i = 0 To UBound(xParaArray)
        For j = 0 To UBound(ParaNmArray)
            wkPosNo = InStr(1, ParaNmArray(j), "@")
            If wkPosNo > 0 Then
                wkSeqNo = CInt(Mid(ParaNmArray(j), wkPosNo + 1)) - 1
                wkParaNm = Left(ParaNmArray(j), wkPosNo - 1)
            Else
                wkSeqNo = 0
                wkParaNm = ParaNmArray(j)
            End If
            
            If xParaArray(i) Like wkParaNm Then
                wkVal = WS.Cells(xValRowNo(i), xValColNo(i) + wkSeqNo).Value
                If wkVal = "" Then
                    outMsg "E130", xParaArray(i) & IIf(wkSeqNo > 0, "@" & wkSeqNo + 1, "") & "を入力して下さい。"
                    WS.Cells(xValRowNo(i), xValColNo(i) + wkSeqNo).Select
                    ckParaVal = False
                    Exit Function
                End If
            End If
        Next j
        
        For j = 0 To UBound(ParaNmArray2)
            If xParaArray(i) Like ParaNmArray2(j) Then
                wkVal = WS.Cells(xValRowNo(i), xValColNo(i))
                If wkVal = "" Then
                    outMsg "E131", xParaArray(i) & "を入力して下さい。"
                    WS.Cells(xValRowNo(i), xValColNo(i)).Select
                    ckParaVal = False
                    Exit Function
                End If
                
                On Error Resume Next ' Dir関数でエラーになる時がある。
                If Dir(wkVal, vbDirectory) = "" Then
                    outMsg "E132", xParaArray(i) & "が存在しません。"
                    WS.Cells(xValRowNo(i), xValColNo(i)).Select
                    ckParaVal = False
                    Exit Function
                End If
                On Error GoTo 0
                
                If Right(wkVal, 1) <> "\" Then
                    outMsg "E133", xParaArray(i) & "の最後に'\'マークを付けて下さい。"
                    WS.Cells(xValRowNo(i), xValColNo(i)).Select
                    ckParaVal = False
                    Exit Function
                End If
            End If
        Next j
        
        For j = 0 To UBound(ParaNmArray3)
            If xParaArray(i) Like ParaNmArray3(j) Then
                wkVal = WS.Cells(xValRowNo(i), xValColNo(i))
                If wkVal = "" Then
                    outMsg "E134", xParaArray(i) & "を入力して下さい。"
                    WS.Cells(xValRowNo(i), xValColNo(i)).Select
                    ckParaVal = False
                    Exit Function
                End If
                
                On Error Resume Next ' Dir関数でエラーになる時がある。
                If Dir(wkVal) = "" Then
                    outMsg "E135", xParaArray(i) & "が存在しません。"
                    WS.Cells(xValRowNo(i), xValColNo(i)).Select
                    ckParaVal = False
                    Exit Function
                End If
                On Error GoTo 0
            End If
        Next j
    Next i
        
    Set WS = Nothing
End Function

' **************************************************
' XXLダウンロード処理のまとめ
' @In : Save Folder Name
' @In : Save File Name
' @Out : True - OK, False - NG
' **************************************************
Public Function dlXXLBat(inSaveFolderNm As String, inSaveFileNm As String) As Boolean
    ' スプレッドシート選択画面遷移確認
    Select Case ckScreenTransition("wnd[1]", "*スプレッド*", get_Timeout(2))
        Case 0
            ' OK : Continue
        Case 1
            outMsg "E121", "ファイル保存設定 画面出力失敗！"
            dlXXLBat = False
            Exit Function
        Case 2
            outMsg "E122", "ファイル保存設定画面 タイムオーバー"
            dlXXLBat = False
            Exit Function
        Case Else
            outMsg "S101", "System Error!"
            dlXXLBat = False
            Exit Function
    End Select
    
    With SAPSesi
     ' スプレッドシート選択操作　ファイルフォーマット選択（export.XXL）
    .findById("wnd[1]/usr/radRB_OTHERS").Select
    .findById("wnd[1]/usr/cmbG_LISTBOX").Key = "08"
    .findById("wnd[1]").sendVKey 0
    
    waitSec  '1秒待ち
    
    ' 情報画面(フィルター基準～考慮されません）の操作
    .findById("wnd[0]").sendVKey 0
    
    waitSec  '1秒待ち
    
    '一覧をXXLにエクスポート画面の操作
    .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select 'Tableを選択
    .findById("wnd[1]").sendVKey 0

    waitSec  '1秒待ち

    '一覧をXXLにエクスポート画面(2)の操作
    '選択は、Microsoft Excelのみ
    .findById("wnd[1]").sendVKey 0
        
    ' 出力ファイルの編集
    xSaveFilePath = inSaveFolderNm & inSaveFileNm & "_" & Format(Now, "YYYYMMDDhhmmss") & ".xlsx"
    
    ' SAPよりExcelが自動で立ち上がる。
    Select Case saveSAPExcelFile(xSaveFilePath)
        Case 0
            'OK : Continue
        Case 1
            outMsg "E124", "SAPより開かれたExcelファイルが見つかりません。"
            dlXXLBat = False
            Exit Function
        Case 2
            outMsg "E125", "Excelファイル保存エラー。"
            dlXXLBat = False
            Exit Function
        Case Else
            outMsg "S103", "System Error!"
            dlXXLBat = False
            Exit Function
    End Select
    
    ' 情報画面(スプレッドシートにデータを保存して下さい。）の操作
    .findById("wnd[1]").sendVKey 0
    
    End With
    
    dlXXLBat = True
End Function

' **************************************************
' バリアント検索処理のまとめ。最初にバリアント検索画面が表示されるパターンで使用する。
' @In : Variant名
' @Out : True - OK, False - NG
' **************************************************
Public Function selVariantBat(inVariantNm As String) As Boolean
    With SAPSesi

    ' バリアント検索画面へ
    .findById("wnd[0]").sendVKey 17
    
    waitSec '1秒待ち
    
    'バリアント検索画面の操作
    .findById("wnd[1]/usr/txtENAME-LOW").text = ""
    .findById("wnd[1]").sendVKey 8
    
    waitSec '1秒待ち

    'バリアント画面の操作
    If getGridLineNo("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell", "VARIANT", inVariantNm & "*") = 0 Then
        outMsg "E141", "バリアント名が存在しません。" & vbCrLf & "バリアント名 = " & inVariantNm
        selVariantBat = False
        Exit Function
    End If
    .findById("wnd[1]").sendVKey 2
    
    End With
    
    waitSec '1秒待ち
        
    selVariantBat = True
End Function

' **************************************************
' Mhtmlダウンロード処理のまとめ
' @In : Save Folder Name
' @In : Save File Name　*日付時間を入れたい時は、呼び出し元で設定する。
' @Out : True - OK, False - NG
' **************************************************
Public Function dlMhtmlBat(inSaveFolderNm As String, inSaveFileNm As String) As Boolean
    Dim wkMhtmlFilePath As String
    
    ' スプレッドシート選択画面遷移確認
    Select Case ckScreenTransition("wnd[1]", "*スプレッド*", get_Timeout(2))
        Case 0
            ' OK : Continue
        Case 1
            outMsg "E151", "ファイル保存設定 画面出力失敗！"
            dlMhtmlBat = False
            Exit Function
        Case 2
            outMsg "E152", "ファイル保存設定画面 タイムオーバー"
            dlMhtmlBat = False
            Exit Function
        Case Else
            outMsg "S151", "System Error!"
            dlMhtmlBat = False
            Exit Function
    End Select
    
    With SAPSesi
    
    ' スプレッドシート選択操作　ファイルフォーマット選択（export.MHTML）
    .findById("wnd[1]/usr/radRB_1").Select
    .findById("wnd[1]").sendVKey 0  'Enter

    waitSec
    
    ' ファイル名入力画面,保存しない
    .findById("wnd[1]/usr/ctxtDY_PATH").text = inSaveFolderNm  '出力フォルダを設定
    .findById("wnd[1]/usr/ctxtDY_FILENAME").text = inSaveFileNm & ".MHTML" '出力ファイル名を設定
    waitSec '1秒待ち
    .findById("wnd[1]").sendVKey 0  'Enter
    
    ' ファイル転送済みステータス確認
    Select Case ckStatusTransition("*転送*", get_Timeout(2))
        Case 0
            ' OK : Continue
        Case 1
            outMsg "E153", "Mhtmlファイル保存 失敗！" & vbCrLf & "フォルダ = " & inSaveFolderNm & vbCrLf & "ファイル名 = " & inSaveFileNm
            dlMhtmlBat = False
            Exit Function
        Case 2
            outMsg "E154", "Mhtmlファイル保存 タイムオーバー" & vbCrLf & "フォルダ = " & inSaveFolderNm & vbCrLf & "ファイル名 = " & inSaveFileNm
            dlMhtmlBat = False
            Exit Function
        Case Else
            outMsg "S152", "System Error!"
            dlMhtmlBat = False
            Exit Function
    End Select
         
    waitSec '1秒待ち
    
    wkMhtmlFilePath = inSaveFolderNm & inSaveFileNm & ".MHTML"
    
    ' 出力されたファイルがSAPで開かれるかチェックする。
    Select Case ckFileExported(wkMhtmlFilePath, 10) '10秒待ち
        Case 0
            ' OK : Continue
        Case 1
            outMsg "E155", "SAPからファイルが開かれません。タイムオーバー。" & vbCrLf & "ファイルパス = " & wkMhtmlFilePath
            dlMhtmlBat = False
            Exit Function
        Case Else
            outMsg "E156", "SAPからファイルが開かれないその他のエラー。" & vbCrLf & "ファイルパス = " & wkMhtmlFilePath
            dlMhtmlBat = False
            Exit Function
    End Select
    
    '抽出データオーバー可能性ありWindows7
    waitSec get_Timeout(3)
    
    xSaveFilePath = inSaveFolderNm & inSaveFileNm & ".xlsx"
    
    Select Case cnvExcelFile(wkMhtmlFilePath, xSaveFilePath)
        Case 0
            'OK : Continue
        Case 1
            outMsg "E157", "MHTMLファイル読み込みエラー" & vbCrLf & wkMhtmlFilePath
            dlMhtmlBat = False
            Exit Function
        Case 2
            outMsg "E158", "MHTML => Excelファイル変換エラー" & vbCrLf & wkMhtmlFilePath & vbCrLf & xSaveFilePath
            dlMhtmlBat = False
            Exit Function
        Case Else
            outMsg "S153", "System Error!"
            dlMhtmlBat = False
            Exit Function
    End Select
    
    End With
    
    dlMhtmlBat = True
End Function

' **************************************************
' SQ01画面遷移のまとめ
' @In : 部門コード
' @In : SQ01画面内のトランザクションコード
' @In : 遷移画面名
' @Out : True - OK, False - NG
' **************************************************
Public Function tranSQ01Bat(inDeptCd As String, inTranCd As String, inFormNm As String) As Boolean
    ' 指定トランザクションを設定→実行する｡
    xTranCd = "SQ01"
    setTranCd
    
    ' ユーザーグループ画面の遷移確認
    Select Case ckScreenTransition("wnd[0]", "*ユーザグループ*", get_Timeout(1))
        Case 0
            ' OK : Continue
        Case 1
            outMsg "E161", "「ユーザグループ」画面の出力が失敗しました。" & vbCrLf & "トランザクションコード = SQ01"
            tranSQ01Bat = False
            Exit Function
        Case 2
            outMsg "E162", "ユーザグループ画面 タイムオーバー"
            tranSQ01Bat = False
            Exit Function
        Case Else
            outMsg "S161", "System Error!"
            tranSQ01Bat = False
            Exit Function
    End Select
    
    With SAPSesi
    
    ' ユーザーグループ画面の操作
    .findById("wnd[0]").sendVKey 19
    
    waitSec '1秒待ち
    
    '部門選択画面の操作
    If getGridLineNo("wnd[1]/usr/cntlGRID1/shellcont/shell", "DBGBNUM", inDeptCd) = 0 Then
        outMsg "E163", "部門名が取得出来ません。Code = inDeptCd"
        tranSQ01Bat = False
        Exit Function
    End If
    .findById("wnd[0]").sendVKey 0
    
    waitSec '1秒待ち
    
    ' ユーザーグループ画面の操作
    If getGridLineNo("wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell", "QNUM", inTranCd) = 0 Then
        outMsg "E164", "SQ01にトランザクションがありません。" & vbCrLf & "トランザクション = " & inTranCd
        tranSQ01Bat = False
        Exit Function
    End If
    .findById("wnd[0]").sendVKey 8

    '　指定の画面遷移確認
    Select Case ckScreenTransition("wnd[0]", "*" & inFormNm & "*", get_Timeout(1))
        Case 0
            ' OK : Continue
        Case 1
            outMsg "E165", "「" & inFormNm & "」画面の出力が失敗しました。" & vbCrLf & "トランザクション = " & inTranCd
            tranSQ01Bat = False
            Exit Function
        Case 2
            outMsg "E166", inFormNm & "画面 タイムオーバー"
            tranSQ01Bat = False
            Exit Function
        Case Else
            outMsg "S162", "System Error!"
            tranSQ01Bat = False
            Exit Function
    End Select
    
    End With
    
    tranSQ01Bat = True
End Function


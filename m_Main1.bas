Attribute VB_Name = "m_Main"
Option Explicit

Public Sub MainProc()
    Dim SAPTran As New c_SAPAccess
    
    Dim w_MhtmlFilePath As String
    Dim w_XlsxFilePath As String
    Dim w_path As String
    Dim w_fname As String
    Dim w_Now As String
    Dim w_bname As String
    Dim i As Integer
    
    ' Script Start Log
    outMsg "L01", "*** Script Cost110 Started. ***"
    outMsg "I01", "処理を開始します。"
    
    ' Excelメッセージ非表示に設定する｡
    Application.DisplayAlerts = False
    
    ' 保管場所が正しいかチェック
    If ckSaveFolder = False Then
        Exit Sub
    End If
    
    ' Save File Pathをクリア
    For i = 1 To 9
        set_SpreadX_SaveFullPath i, ""
    Next i
    DoEvents

    ' ログオンSAP
    SAPTran.Connection = get_Connection ' Connection設定
    SAPTran.Client = get_Client 'クライアント設定
    SAPTran.User = get_User       'ユーザ設定
    SAPTran.Password = get_Password 'パスワード設定
    SAPTran.Language = get_Language    '言語設定
    If SAPTran.logonSAP() > 0 Then 'エラーなら処理を終了
        Exit Sub
    End If
    
    ' 指定トランザクションを設定→実行する｡
    SAPTran.TranCd = "SQ01"
    SAPTran.setTranCd
    
    ' ユーザグループ画面の遷移確認
    Select Case SAPTran.ckScreenTransition("wnd[0]", "*ユーザグループ*", get_Timeout(1))
        Case 0
            ' OK : Continue
        Case 1
            outMsg "E11", "ユーザグループ　画面出力失敗！"
            Exit Sub
        Case 2
            outMsg "E12", "ユーザグループ　タイムオーバー"
            Exit Sub
        Case Else
            outMsg "S11", "System Error!"
            Exit Sub
    End Select
    
    With SAPTran.SAPSesi
    
    ' ユーザグループ画面の操作
    .findById("wnd[0]").sendVKey 19 'Shift+F7, ユーザグループ部門選択画面の表示
    
    ' ユーザグループ部門選択画面の遷移確認
    Select Case SAPTran.ckScreenTransition("wnd[1]", "*ユーザグループ*", get_Timeout(1))
        Case 0
            ' OK : Continue
        Case 1
            outMsg "E13", "ユーザグループ部門選択　画面出力失敗！"
            Exit Sub
        Case 2
            outMsg "E14", "ユーザグループ部門選択　タイムオーバー"
            Exit Sub
        Case Else
            outMsg "S12", "System Error!"
            Exit Sub
    End Select
    
    ' ユーザグループ部門選択画面の操作
    .findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectColumn "DBGBNUM"
    .findById("wnd[1]/usr/cntlGRID1/shellcont/shell").contextMenu
    .findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&FILTER"
    
    waitSec '1秒待ち
    
    'Filtter画面
    .findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "PRISM03"
    .findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").Text = "PRISM03"
    .findById("wnd[2]").sendVKey 0
    
    waitSec '1秒待ち
    
    ' ユーザグループ部門選択画面の操作
    .findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "0"
    .findById("wnd[1]").sendVKey 0
    
    waitSec '1秒待ち
    
    ' ユーザグループ画面の操作
    .findById("wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell").selectColumn "QNUM"
    .findById("wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell").pressToolbarButton "&MB_FILTER"
    
    waitSec '1秒待ち
    
    'Filtter画面
    .findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "ZISMSD0017"
    .findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").Text = "ZISMSD0017"
    .findById("wnd[1]").sendVKey 0

    waitSec '1秒待ち
    
    ' ユーザグループ画面の操作
    .findById("wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell").selectedRows = "0"
    .findById("wnd[0]").sendVKey 8

    w_Now = Format(Now, "yyyymmddhhnnss")  '最初にファイルの作成時間を取得する。
    For i = 1 To 9
'        Select Case i
'            Case 2, 5, 6
'                ' continue
'            Case Else
'                GoTo loopnext
'        End Select
        
        If 5 = i Then
            .findById("wnd[0]").sendVKey 3
            .findById("wnd[0]").sendVKey 3
            .findById("wnd[0]/tbar[0]/okcd").Text = "ZISM_SD_R0042"
            .findById("wnd[0]").sendVKey 0
        End If

        ' 売上/売上原価明細画面の遷移確認
        Select Case SAPTran.ckScreenTransition("wnd[0]", "*売上*", get_Timeout(1))
            Case 0
                ' OK : Continue
            Case 1
                outMsg "E15", "売上/売上原価明細画面出力失敗！"
                Exit Sub
            Case 2
                outMsg "E16", "売上/売上原価明細画面タイムオーバー"
                Exit Sub
            Case Else
                outMsg "S13", "System Error!"
                Exit Sub
        End Select
        
        ' 売上/売上原価明細画面の操作
        .findById("wnd[0]").sendVKey 17  'shift+F5
        
        If i >= 5 Then
            .findById("wnd[1]/usr/txtENAME-LOW").Text = ""
            .findById("wnd[1]/tbar[0]/btn[8]").press
        End If

        'バリアント画面の操作
        If SAPTran.getGridLineNo("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell", "VARIANT", getVariantNm(i)) = 0 Then
            outMsg "E25", "バリアントの内容が正しくありません。" & vbCrLf & "Variant=" & getVariantNm(i)
            Exit Sub
        End If
        
        waitSec '1秒待ち
        .findById("wnd[1]").sendVKey 2 'PF2
    
        waitSec '1秒待ち
    
        ' 売上画面の操作
        .findById("wnd[0]").sendVKey 8
        
        ' 売上/売上原価明細画面の遷移確認
        Select Case SAPTran.ckScreenTransition("wnd[0]", "*売上*", get_Timeout(2))
            Case 0
                ' OK : Continue
            Case 1
                outMsg "E17", "売上/売上原価明細画面(2)出力失敗！"
                Exit Sub
            Case 2
                outMsg "E18", "売上/売上原価明細画面(2)タイムオーバー"
                Exit Sub
            Case Else
                outMsg "S13", "System Error!"
                Exit Sub
        End Select
        
        ' 売上画面の操作 その２
        If i <= 4 Then
            .findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
            .findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
        Else
            .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow = -1
            .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
            .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
        End If

        ' スプレッドシート選択画面遷移確認
        Select Case SAPTran.ckScreenTransition("wnd[1]", "*スプレッド*", get_Timeout(3))
            Case 0
                ' OK : Continue
            Case 1
                outMsg "E19", "ファイル保存設定　画面出力失敗！"
                Exit Sub
            Case 2
                outMsg "E20", "ファイル保存設定画面　タイムオーバー"
                Exit Sub
            Case Else
                outMsg "S14", "System Error!"
                Exit Sub
        End Select
    
        ' スプレッドシート選択操作　ファイルフォーマット選択（export.MHTML）
        .findById("wnd[1]/usr/radRB_1").SetFocus
        .findById("wnd[1]/usr/radRB_1").Select
        waitSec 3 '3秒待ち
        .findById("wnd[1]").sendVKey 0  'Enter
            
        ' 売上画面の遷移確認　四角い画面　30分ガード
        Select Case SAPTran.ckScreenTransition("wnd[1]", "*売上*", get_Timeout(3))
            Case 0
                ' OK : Continue
            Case 1
                outMsg "E21", "売上/売上原価明細画面(3)出力失敗！"
                Exit Sub
            Case 2
                outMsg "E22", "売上/売上原価明細画面(3)タイムオーバー"
                Exit Sub
            Case Else
                MsgBox "S15", "System Error!"
                Exit Sub
        End Select
            
        ' ファイル名入力画面
        w_path = get_SpreadT_SavePath
        .findById("wnd[1]/usr/ctxtDY_PATH").Text = w_path  '出力フォルダーを設定
        w_fname = getFileNm(i) & "_" & w_Now & ".MHTML"
        .findById("wnd[1]/usr/ctxtDY_FILENAME").Text = w_fname '出力ファイル名を設定
        waitSec 2 '2秒待ち
        .findById("wnd[1]").sendVKey 0  'Enter
             
        ' ファイル転送済みステータス確認
        Select Case SAPTran.ckStatusTransition("*転送*", get_Timeout(3))
            Case 0
                ' OK : Continue
            Case 1
                outMsg "E23", "ファイル保存失敗！"
                Exit Sub
            Case 2
                outMsg "E24", "ファイル保存　タイムオーバー"
                Exit Sub
            Case Else
                outMsg "S16", "System Error!"
                Exit Sub
        End Select
             
        ' SAPから出力されたファイル(MHTML)が出力完了かどうかの確認, 30分待ち
        w_MhtmlFilePath = w_path & w_fname
        Select Case SAPTran.ckFileExported(w_MhtmlFilePath)
            Case 0
                ' OK : Continue
            Case 1
                outMsg "E51", "ファイル保存　タイムオーバー"
                Exit Sub
            Case Else
                outMsg "S17", "System Error!"
                Exit Sub
        End Select
        
        ' Excel Save File
        If i <= 4 Then
            w_path = getSaveXlsxFolder(1)
        Else
            w_path = getSaveXlsxFolder(2)
        End If
        w_bname = getFileNm(i) & "_" & w_Now
        w_XlsxFilePath = w_path & w_bname & ".xlsx"
            
        Select Case SAPTran.cnvExcelFile(w_MhtmlFilePath, w_XlsxFilePath)
            Case 0
                'OK : Continue
            Case 1
                outMsg "E52", "THXMLファイル読み込みエラー"
                Exit Sub
            Case 2
                outMsg "E53", "THXML => Excelファイル変換エラー"
                Exit Sub
            Case Else
                outMsg "S18", "System Error!"
                Exit Sub
        End Select
        
        set_SpreadX_SaveFullPath i, w_XlsxFilePath
        outMsg "L02", "File Created. Path = " & w_XlsxFilePath
        
        ' 前の画面に戻る。
        .findById("wnd[0]").sendVKey 3
        
        DoEvents '制御を戻す。
loopnext:
    Next i
    
    .findById("wnd[0]").sendVKey 3

    End With
    
    ' Logoff SAP
    Select Case SAPTran.logoffSAP
        Case 0 'OK
            ' Continue
        Case Else
            outMsg "S19", "System Error!"
            Exit Sub
    End Select
    
    Application.DisplayAlerts = True
                
TestEnd:
    
    outMsg "I03", "処理が正常に終了しました。"
    
    outMsg "L03", "*** Script Cost110 Ended ***"
End Sub


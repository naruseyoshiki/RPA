Attribute VB_Name = "m_Main"
Option Explicit

Public Sub MainProc()
Attribute MainProc.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim SAPTran As New c_SAPAccess
    Dim i As Integer
    
    ' Script Start Log
    outMsg "L01", "*** Script Cost110 DL_Started. ***"
    
    'パラメータの設定
    SAPTran.setParaVal "コントロール", 1, 2
    
    'Excel入力欄チェック
    If SAPTran.ckParaVal("勘定区分*@2,勘定区分*@3,勘定区分*@4,勘定区分_2@5,バリアント名*@2,バリアント名*@3,バリアント名*@4,バリアント名_2@5,バリアント名*@2,バリアント名*@3,バリアント名*@4,バリアント名_2@5") = False Then
        Exit Sub
    End If
        
    If outMsg("Q01", "処理を開始しますか?") = vbNo Then
        outMsg "L02", "処理がキャンセルされました。"
        Exit Sub
    End If
    
    ' Excelメッセージ非表示に設定する｡
    Application.DisplayAlerts = False
    
    ' Save File Pathをクリア
    For i = 1 To 9
        SAPTran.putParaVal "*売上原価明細保存場所結果（フルパス）", "", i
    Next i
    
    For i = 1 To 3
        SAPTran.putParaVal "*Excel編集ファイル保存場所結果（フルパス）", "", i
    Next i

    ' ログオンSAP
    SAPTran.Connection = get_Connection ' Connection設定
    SAPTran.Client = get_Client 'クライアント設定
    SAPTran.User = get_User       'ユーザ設定
    SAPTran.Password = get_Password 'パスワード設定
    SAPTran.Language = get_Language    '言語設定
    If SAPTran.LogonSAP() > 0 Then 'エラーなら処理を終了
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
            outMsg "E01", "「ユーザグループ」画面の遷移が失敗しました。" & vbCrLf & "トランザクションコード = SQ01"
            Exit Sub
        Case 2
            outMsg "E02", "ユーザグループ　タイムオーバー"
            Exit Sub
        Case Else
            outMsg "S01", "System Error!"
            Exit Sub
    End Select
    
    With SAPTran.SAPSesi
    
    ' ユーザグループ画面の操作
    .findById("wnd[0]").sendVKey 19 'Shift+F7, ユーザグループ部門選択画面の表示
    
    waitSec '1秒待ち
   
    'バリアント画面の操作
    If SAPTran.getGridLineNo("wnd[1]/usr/cntlGRID1/shellcont/shell", "DBGBNUM", "PRISM03") = 0 Then
        outMsg "E03", "クエリが存在しません。 " & "PRISM03"
        Exit Sub
    End If
    .findById("wnd[1]").sendVKey 2
    
    ' ユーザグループ画面の操作
    If SAPTran.getGridLineNo("wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell", "QNUM", SAPTran.getParaVal("トランザクション(SQ01)")) = 0 Then
        outMsg "E04", "トランザクションが存在しません。" & vbCrLf & "トランザクション = " & SAPTran.getParaVal("トランザクション(SQ01)")
        Exit Sub
    End If
    .findById("wnd[0]").sendVKey 8
    
    'ここからループ
    For i = 1 To 9
        
        If 5 = i Then
            .findById("wnd[0]").sendVKey 3
            .findById("wnd[0]").sendVKey 3
            .findById("wnd[0]/tbar[0]/okcd").text = SAPTran.getParaVal("トランザクションコード（統計)")
            .findById("wnd[0]").sendVKey 0
        End If

        ' 売上/売上原価明細画面の遷移確認
        Select Case SAPTran.ckScreenTransition("wnd[0]", "*売上*", get_Timeout(1))
            Case 0
                ' OK : Continue
            Case 1
                If i < 5 Then
                    outMsg "E05", "売上/売上原価明細画面出力失敗！" & vbCrLf & "ファイル名 = " & SAPTran.getParaVal("ファイル名", i)
                    Exit Sub
                Else
                    outMsg "E06", "売上/売上原価明細画面出力失敗！" & vbCrLf & "ファイル名 = " & SAPTran.getParaVal("ファイル名_2", i - 4)
                    Exit Sub
                End If
            Case 2
                outMsg "E07", "売上/売上原価明細画面タイムオーバー"
                Exit Sub
            Case Else
                outMsg "S02", "System Error!"
                Exit Sub
        End Select
        
        'バリアント画面の操作
        If i < 5 Then
            ' 売上/売上原価明細画面の操作
            .findById("wnd[0]").sendVKey 17  'shift+F5
        
            If SAPTran.getGridLineNo("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell", "VARIANT", SAPTran.getParaVal("バリアント名", i)) = 0 Then
                outMsg "E08", "バリアントの内容が正しくありません。" & vbCrLf & "Variant=" & SAPTran.getParaVal("バリアント名", i)
                Exit Sub
            End If
            
            waitSec '1秒待ち
            
            .findById("wnd[1]").sendVKey 2 'PF2
    
        Else
            If SAPTran.selVariantBat(SAPTran.getParaVal("バリアント名_2", i - 4)) = False Then
               Exit Sub
            End If
        End If
        
        ' 売上画面の操作
        .findById("wnd[0]").sendVKey 8
        
        If i < 5 Then
            If SAPTran.ckDataSelected("") = False Then
                outMsg "L03", "データは選択されませんでした。No=" & i & vbCrLf & "Variant=" & SAPTran.getParaVal("バリアント名", i)
                GoTo LoopNext
            End If
        Else
            If SAPTran.ckDataSelected("") = False Then
                outMsg "L04", "データは選択されませんでした。No=" & i & vbCrLf & "Variant=" & SAPTran.getParaVal("バリアント名_2", i - 4)
                GoTo LoopNext
            End If
        End If
        ' 売上画面の操作 その２,XXL形式データ保存
        If i < 5 Then
            .findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
            .findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
            
            If SAPTran.dlXXLBat(SAPTran.getParaVal("SQ01の保管場所"), SAPTran.getParaVal("ファイル名", i)) = False Then
                Exit Sub
            End If
            
        Else
            .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow = -1
            .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
            .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
            
            If SAPTran.dlXXLBat(SAPTran.getParaVal("統計の保管場所"), SAPTran.getParaVal("ファイル名_2", i - 4)) = False Then
                Exit Sub
            End If
        End If
        
        SAPTran.putParaVal "*売上原価明細保存場所結果（フルパス）", SAPTran.SaveFilePath, i
        outMsg "L05", "File Created. Path = " & SAPTran.getParaVal("*売上原価明細保存場所結果（フルパス）", i)

        
        ' 前の画面に戻る。
        .findById("wnd[0]").sendVKey 3
        
        DoEvents '制御を戻す。
LoopNext:
    Next i
    
    .findById("wnd[0]").sendVKey 3

    End With
    
    ' Logoff SAP
    Select Case SAPTran.logoffSAP
        Case 0 'OK
            ' Continue
        Case Else
            outMsg "S03", "System Error!"
            Exit Sub
    End Select
    
    Set SAPTran = Nothing
    
    Application.DisplayAlerts = True
                
TestEnd:
    
    outMsg "I03", "処理が正常に終了しました。"
    
    outMsg "L06", "*** Script Cost110 DL_Ended. ***"
End Sub


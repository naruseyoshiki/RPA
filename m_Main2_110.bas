Attribute VB_Name = "m_main2"
Option Explicit

Public Sub MainProc2()
    Dim MergeSht As New c_ExcelUser    ' 結合ファイルのワークシートを設定
    Dim SAPTran As New c_SAPAccess     'コントロールシート
    Dim DLSht As New c_ExcelUser       'DLファイル
    Dim Cost_DL As New c_ExcelUser     '売上原価明細フォーマット
    Dim Cost_Pivo As New c_ExcelUser   '売上原価明細フォーマットぴぼ
    Dim Cost_Paste As New c_ExcelUser  '売上原価明細フォーマット貼付用
    Dim New_Paste As New c_ExcelUser   '新規ファイル貼付用
    Dim i As Integer
    Dim bond As String                  '結合ファイルパス
    Dim item As String                  '明細ファイルパス
    Dim paste As String                 '貼付用ファイルパス
    
    '画面描画抑止
    Application.ScreenUpdating = False
    
    ' Script Start Log
    outMsg "L07", "*** Script Cost110 EditExcel_Started. ***"
    
    'パラメータの設定
    SAPTran.setParaVal "コントロール", 1, 2
    
    'Excel入力欄チェック
    If SAPTran.ckParaVal("勘定区分*@2,勘定区分*@3,勘定区分*@4,勘定区分_2@5,バリアント名*@2,バリアント名*@3,バリアント名*@4,バリアント名_2@5,バリアント名*@2,バリアント名*@3,バリアント名*@4,バリアント名_2@5") = False Then
        Exit Sub
    End If
    
    If outMsg("Q02", "処理を開始しますか?") = vbNo Then
        outMsg "L08", "処理がキャンセルされました。"
        Exit Sub
    End If
    
    ' 結合フォーマットファイルを開く
    If MergeSht.openWorkBook(SAPTran.getParaVal("売上原価明細結合フォーマットファイルパス")) = False Then
        outMsg "E09", "売上原価明細結合フォーマットファイルを開けません。" & vbCrLf & "ファイルパス = " & SAPTran.getParaVal("売上原価明細結合フォーマットファイルパス")
        Exit Sub
    End If
    
    '結合ファイルパスの連結
    bond = SAPTran.getParaVal("結合ファイルの保管場所") & SAPTran.getParaVal("結合ファイルのファイル名") & "_" & Format(Now, "yyyymmddhhnnss")
    
    ' 結合ファイルとして名前を付けて保存
    If MergeSht.SaveAsWorkBook(bond) = False Then
        outMsg "E10", "結合ファイルを名前を付けて保存できませんでした。" & vbCrLf & "結合ファイル = " & bond
        Exit Sub
    End If

    ' 結合ファイルのワークシートを設定
    If MergeSht.setWorkSht("", 1, 1, 16) = False Then
        outMsg "E11", "結合ファイルのシートを設定出来ませんでした。" & vbCrLf & "結合ファイル = " & bond
        Exit Sub
    End If
            
    For i = 1 To 5
        If i > 1 Then
            ' 結合ファイルのMax Noのリセット
            MergeSht.resetMaxNo
        End If
        
        'DLファイルパス有無チェック
        If SAPTran.getParaVal("*売上原価明細保存場所結果（フルパス）", i + 4) = "" Then
            outMsg "L09", SAPTran.getParaVal("ファイル名_2", i) & "のDLファイルが見つかりませんでした。" & vbCrLf & "DLファイル = " & SAPTran.getParaVal("*売上原価明細保存場所結果（フルパス）", i + 4)
            GoTo notDL
        End If
        
        ' DLファイルを開く
        If DLSht.openWorkBook(SAPTran.getParaVal("*売上原価明細保存場所結果（フルパス）", i + 4)) = False Then
            outMsg "E12", "DLファイルを開けませんでした。" & vbCrLf & "DLファイル = " & SAPTran.getParaVal("*売上原価明細保存場所結果（フルパス）", i + 4)
            Exit Sub
        End If
        Debug.Print SAPTran.getParaVal("*売上原価明細保存場所結果（フルパス）", i + 4)
        
        '結合ファイル保存場所結果を設定
        SAPTran.putParaVal "*Excel編集ファイル保存場所結果（フルパス）", bond & ".xlsx", 1
        
        ' DLファイルのワークシートを設定
        If DLSht.setWorkSht("", 1, 1, 16) = False Then
            outMsg "E13", "DLファイルのシートを設定出来ませんでした。" & vbCrLf & "DLファイル = " & SAPTran.getParaVal("*売上原価明細保存場所結果（フルパス）", i + 4)
            Exit Sub
        End If

        If DLSht.MaxRowNo = DLSht.StartRowNo Then
            GoTo Down
        End If
        
        ' 結合ファイルのフォーマットのコピー
        MergeSht.copyFormatRow MergeSht.StartRowNo + 1, MergeSht.MaxRowNo + 1, DLSht.MaxRowNo - DLSht.StartRowNo
        
        ' DLファイルから結合ファイルへのコピー
        DLSht.copyData DLSht.StartRowNo + 1, DLSht.getColNo("集計キー"), DLSht.MaxRowNo, DLSht.getColNo("計画原価")
        MergeSht.pasteValueData MergeSht.MaxRowNo + 1, MergeSht.getColNo("集計キー")
        
        '勘定区分の設定
        MergeSht.WS.Cells(MergeSht.MaxRowNo + 1, MergeSht.getColNo("勘定区分")) = SAPTran.getParaVal("勘定区分_2", i)
        MergeSht.WS.Cells(MergeSht.MaxRowNo + 1, MergeSht.getColNo("会計期間")) = Month(DateAdd("m", -4, Now))
        MergeSht.copyData MergeSht.MaxRowNo + 1, MergeSht.getColNo("勘定区分"), MergeSht.MaxRowNo + 1, MergeSht.getColNo("会計期間")
        MergeSht.pasteValueData MergeSht.MaxRowNo + 1, MergeSht.getColNo("勘定区分"), MergeSht.MaxRowNo + DLSht.MaxRowNo - DLSht.StartRowNo, MergeSht.getColNo("会計期間")
        
Down:
        If DLSht.closeWorkBook() = False Then
            outMsg "E14", "DLファイルをクローズ出来ませんでした。" & vbCrLf & "DLファイル = " & SAPTran.getParaVal("*売上原価明細保存場所結果（フルパス）", i + 4)
            Exit Sub
        End If
        
        outMsg "L10", "--Copy Completed-- " & SAPTran.getParaVal("*Excel編集ファイル保存場所結果（フルパス）", 1) & " DLファイル追加件数 = " & DLSht.MaxRowNo - DLSht.StartRowNo
notDL:
        ' 結合ファイルのMax Noのリセット
        MergeSht.resetMaxNo
        
    Next i
        
    '結合ファイルのデータ有無チェック
    If MergeSht.MaxRowNo = MergeSht.StartRowNo Then
        '結合ファイル保存場所結果を設定
        SAPTran.putParaVal "*Excel編集ファイル保存場所結果（フルパス）", bond & ".xlsx", 1
        outMsg "E15", "追加するデータがありませんでした。" & vbCrLf & "結合ファイル = " & SAPTran.getParaVal("*Excel編集ファイル保存場所結果（フルパス）", 1)
        
        If MergeSht.closeWorkBook() = False Then
            outMsg "E16", "結合ファイルをクローズ出来ませんでした。" & vbCrLf & "結合ファイル = " & SAPTran.getParaVal("*Excel編集ファイル保存場所結果（フルパス）", 1)
            Set MergeSht = Nothing
            Exit Sub
        End If
    
        Exit Sub
    End If
        
    outMsg "L11", "結合ファイル全件数 = " & DLSht.RangeRowCnt + MergeSht.MaxRowNo - MergeSht.StartRowNo

    Set DLSht = Nothing

    ' 売上原価明細フォーマットを開く
    If Cost_DL.openWorkBook(SAPTran.getParaVal("売上原価明細フォーマットファイルパス")) = False Then
        outMsg "E17", "売上原価明細フォーマットを開けませんでした。" & vbCrLf & "フォーマット = " & SAPTran.getParaVal("売上原価明細フォーマットファイルパス")
        Exit Sub
    End If
    
    '売上原価明細ファイル名の連結
    item = SAPTran.getParaVal("明細ファイルの保管場所") & SAPTran.getParaVal("明細ファイルのファイル名") & "_" & Format(Now, "yyyymmddhhnnss")
                    
    '売上原価明細ファイルとして名前を付けて保存
    If Cost_DL.SaveAsWorkBook(item) = False Then
        outMsg "E18", "売上原価明細ファイルを保存できませんでした。" & vbCrLf & "売上原価明細ファイル = " & item
        Exit Sub
    End If
    
    '売上原価明細ファイル保存場所結果を設定
    SAPTran.putParaVal "*Excel編集ファイル保存場所結果（フルパス）", item & ".xlsx", 2
    
    '明細のワークシートを設定
    If Cost_DL.setWorkSht("ダウンロードデータ", 1, 1, 16) = False Then
        outMsg "E19", "「ダウンロードデータ」シートを設定できませんでした。" & vbCrLf & "明細ファイル = " & SAPTran.getParaVal("*Excel編集ファイル保存場所結果（フルパス）", 2)
        Exit Sub
    End If
    
    ' シートダウンロードデータへのフォーマットのコピー
    Cost_DL.copyFormatRow Cost_DL.StartRowNo + 1, Cost_DL.StartRowNo + 1, MergeSht.MaxRowNo - MergeSht.StartRowNo
    
    ' フォーミュラーのコピー
    Cost_DL.copyFormulaRow2 Cost_DL.getColNo("地区"), Cost_DL.getColNo("件名読替"), Cost_DL.StartRowNo + 1, Cost_DL.StartRowNo + 2, MergeSht.MaxRowNo - MergeSht.StartRowNo - 1
    
    ' DLファイルから結合ファイルへのコピー
    MergeSht.copyData MergeSht.StartRowNo + 1, MergeSht.getColNo("集計キー"), MergeSht.MaxRowNo, MergeSht.getColNo("勘定区分")
    Cost_DL.pasteValueData Cost_DL.StartRowNo + 1, Cost_DL.getColNo("集計キー")

    '結合ファイルを保存、閉じる
    If MergeSht.saveWorkBook() = False Then
        outMsg "E20", "結合ファイルを保存できませんでした。" & vbCrLf & "ファイル = " & SAPTran.getParaVal("*Excel編集ファイル保存場所結果（フルパス）", 1)
        Exit Sub
    End If
        
    If MergeSht.closeWorkBook() = False Then
        outMsg "E21", "結合ファイルをクローズ出来ませんでした。" & vbCrLf & "ファイル = " & SAPTran.getParaVal("*Excel編集ファイル保存場所結果（フルパス）", 1)
        Exit Sub
    End If
        
    Set MergeSht = Nothing

    ' 売上原価明細ファイル-ダウンロードのMax Noを再取得
    Cost_DL.resetMaxNo
    
    Cost_DL.mkTable "TBL_ダウンロードデータ", Cost_DL.StartRowNo, Cost_DL.getColNo("集計キー"), Cost_DL.MaxRowNo, Cost_DL.getColNo("件名読替")
    
    'ピボファイルの設定
    Cost_Pivo.setWorkBook Cost_DL.WB
    
    If Cost_Pivo.setWorkSht("ぴぼ", 4, 1) = False Then
        outMsg "E22", "ぴぼのワークシートを設定出来ませんでした。" & vbCrLf & "ファイル = " & SAPTran.getParaVal("*Excel編集ファイル保存場所結果（フルパス）", 2)
        Exit Sub
    End If
    
    Cost_Pivo.updPivotDS "TBL_ぴぼ", "TBL_ダウンロードデータ"
    
    'ぴぼのMax Noを再取得
    Cost_Pivo.resetMaxNo
    
    '貼付用の設定
    Cost_Paste.setWorkBook Cost_Pivo.WB
    
    If Cost_Paste.setWorkSht("貼付用", 1, 1) = False Then
        outMsg "E23", "貼付用ワークシートを設定出来ませんでした。" & vbCrLf & "ファイル = " & SAPTran.getParaVal("*Excel編集ファイル保存場所結果（フルパス）", 2)
        Exit Sub
    End If
    
    ' ぴぼシートから貼付用シートへのコピー
    Cost_Paste.copyData Cost_Paste.StartRowNo + 1, Cost_Paste.getColNo("会計期間"), Cost_Paste.StartRowNo + 1, Cost_Paste.getColNo("受注先区分")
    Cost_Paste.pasteData Cost_Paste.StartRowNo + 2, Cost_Paste.getColNo("会計期間"), Cost_Pivo.MaxRowNo - Cost_Pivo.StartRowNo + Cost_Paste.StartRowNo, Cost_Paste.getColNo("会計期間")
    
    
    '貼付用ブック新規作成
    If New_Paste.addWorkBook() = False Then
        outMsg "E24", "ワークブックを新規作成できませんでした。"
        Exit Sub
    End If
    
    '新規貼付用シートを設定
    If New_Paste.setWorkSht("", 1, 1) = False Then
        outMsg "E25", "新規シートに設定出来ませんでした。"
        Exit Sub
    End If
    
    '貼付用シートをコピー、新規ブックに貼り付け
    Cost_Paste.copyAllShtData
    New_Paste.pasteAllShtData
    
    '新規ブックより任意のシート名変更
    If New_Paste.chgShtNm(SAPTran.getParaVal("貼付用ファイルのシート名")) = False Then
        outMsg "E26", "新規ブックのシート名を変更出来ませんでした。"
        Exit Sub
    End If
    
    '新規貼付用ファイル名の連結
    paste = SAPTran.getParaVal("貼付用ファイルの保管場所") & SAPTran.getParaVal("貼付用ファイルのファイル名") & "_" & Format(Now, "yyyymmddhhnnss")
    
    '新規ブックを指定パスで名前を付けて保存
    If New_Paste.SaveAsWorkBook(paste) = False Then
        outMsg "E27", "新規ブックを名前を付けて保存出来ませんでした。" & vbCrLf & "新規ファイルパス = " & paste
        Exit Sub
    End If
    
    '新規ブック保存場所結果を設定
    SAPTran.putParaVal "*Excel編集ファイル保存場所結果（フルパス）", paste & ".xlsx", 3
    
    '明細を保存,閉じる
    If Cost_DL.saveWorkBook() = False Then
        outMsg "E28", "明細を保存できませんでした。" & vbCrLf & "明細ファイル = " & SAPTran.getParaVal("*Excel編集ファイル保存場所結果（フルパス）", 2)
        Exit Sub
    End If
        
    If Cost_DL.closeWorkBook() = False Then
        outMsg "E29", "明細をクローズ出来ませんでした。" & vbCrLf & "明細ファイル = " & SAPTran.getParaVal("*Excel編集ファイル保存場所結果（フルパス）", 2)
        Exit Sub
    End If
    
    If New_Paste.closeWorkBook() = False Then
        outMsg "E30", "新規ブックをクローズ出来ませんでした。" & vbCrLf & "新規ファイル = " & SAPTran.getParaVal("*Excel編集ファイル保存場所結果（フルパス）", 3)
        Exit Sub
    End If
    
    outMsg "L12", "【ぴぼ】全件数 = " & Cost_Pivo.MaxRowNo - Cost_Pivo.StartRowNo
    
    outMsg "L13", "【" & SAPTran.getParaVal("結合ファイルのファイル名") & "】" & "File Created. Path = " & SAPTran.getParaVal("*Excel編集ファイル保存場所結果（フルパス）", 1)
    
    outMsg "L14", "【" & SAPTran.getParaVal("明細ファイルのファイル名") & "】" & "File Created And Copy Completed. Path = " & SAPTran.getParaVal("*Excel編集ファイル保存場所結果（フルパス）", 2)
    
    outMsg "L15", "【" & SAPTran.getParaVal("貼付用ファイルのファイル名") & "】" & "File Created. Path = " & SAPTran.getParaVal("*Excel編集ファイル保存場所結果（フルパス）", 3)
    
    ' 各ファイルの初期化
    Set Cost_Pivo = Nothing
    Set Cost_DL = Nothing
    Set MergeSht = Nothing
    Set DLSht = Nothing
    Set New_Paste = Nothing
    
    outMsg "I04", "処理が正常に終了しました。"
    
    outMsg "L16", "*** Script Cost110 EditExcel_Ended. ***"

    Application.ScreenUpdating = True
    
End Sub

Attribute VB_Name = "m_Common"
Option Explicit
' ###########################################################
' # ---------------------------------------------------------
' # Common Module
' #
' # Version : 20181024 Upd outMsg,Right(getLogFolder, 1)<>"\" ログフォルダ\確認
' # Version : 20181017 Add setCurrentDir,GetFNmFromFPath, copyClipboard, addStrValue
' # Version : 20181017 Add calAccYear, calAccMonth 会計年度、会計期間(月)を取得
' #
' # ---------------------------------------------------------
' ###########################################################

' ###########################################################
' # ---------------------------------------------------------
' # get_DeskTopPath
' # [概要]
' # 　デスクトップパスを取得する。
' # [In]
' #   ─
' # [Out]
' #   デスクトップパス設定値
' # ---------------------------------------------------------
' ###########################################################
Public Function get_DeskTopPath() As String
    
    Dim WSH As Variant
    
    Set WSH = CreateObject("Wscript.Shell") 'WSHオブジェクト
    get_DeskTopPath = WSH.SpecialFolders("Desktop") & "\"
    Set WSH = Nothing 'オブジェクト解放
    
End Function

' ###########################################################
' # ---------------------------------------------------------
' # outMsg
' # [概要]
' # 　メッセージを取得
' # [In]
' #   ─
' # [Out]
' #   エラーコード＋メッセージ内容
' # ---------------------------------------------------------
' ###########################################################
Public Function outMsg(inCode As String, inMsg As String) As Integer
    If getOpeType() = "Manual" Then
        '各メッセージをコードの頭文字表示
        Select Case Left(inCode, 1)
            Case "E"
                MsgBox inCode & " : " & inMsg, vbCritical, "エラー"
            Case "S"
                MsgBox inCode & " : " & inMsg, vbExclamation, "システムエラー"
            Case "I"
                MsgBox inCode & " : " & inMsg, vbInformation, "インフォメーション"
            Case "Q"
                outMsg = MsgBox(inCode & ":" & inMsg, vbQuestion + vbYesNo, "確認")
            Case "L"
                ' Continue ログ専用
            Case Else
                MsgBox inCode & "" & inMsg, vbCritical, "不明なエラー"
                Exit Function
        End Select
    End If
    
    Select Case Left(inCode, 1)
        Case "E", "S", "I", "L"
            ' Continue
        Case "Q"
            If getOpeType() = "Auto" Then
                outMsg = vbYes
            End If
        Case Else
            Exit Function
    End Select
    
    ' ログフォルダの有無チェック、無ならファイル生成無し。
    If Right(getLogFolder, 1) <> "\" Then
        GoTo Exception1
    End If
        
    ' ログの出力
    On Error GoTo Exception1
    Open getLogFolder & getLogFileNm For Append As #1
    Print #1, Format(Now, "YYYY/MM/DD hh:nn:ss") & vbTab & inCode & " : " & Replace(Replace(Replace(inMsg, vbCrLf, " "), "▼", ""), "     ", "")
    Close #1
    On Error GoTo 0
    
    Exit Function
Exception1:
    Debug.Print "Function=outMsg, Log Output Error!, 【Log Folder】 = " & getLogFolder & " 【Log File Name】= " & getLogFileNm
    
End Function

' ###########################################################
' # ---------------------------------------------------------
' # waitSec
' # [概要]
' # 　画面遷移時の待機時間を指定
' # [In]
' #   ─
' # [Out]
' # 　thisSec設定値
' # ---------------------------------------------------------
' ###########################################################

Public Sub waitSec(Optional inTime As Variant)
    
    Dim thisSec As Variant
    
    '引数省略の真偽確認。tureならデフォルト1秒にする。
    If IsMissing(inTime) Then
        thisSec = 1
    Else
        thisSec = inTime
    End If
    
    
    Application.Wait Now + TimeValue("00:00:" & Format(thisSec, "00"))

End Sub

' ###########################################################
' # ---------------------------------------------------------
' # getExcelVer
' # [概要]
' # 　使用中のExcelのバージョンを取得(ver2010まで)
' # [In]
' #   ─
' # [Out]
' #   ExcelVerの取得値
' # ---------------------------------------------------------
' ###########################################################

Public Function getExcelVer() As String

    'Excelバージョンの取得
    Select Case Application.Version
        Case "16.0"
            getExcelVer = "Excel2016"
        Case "15.0"
            getExcelVer = "Excel2013"
        Case "14.0"
            getExcelVer = "Excel2010"
        Case Else
            getExcelVer = Application.Version
    End Select

End Function

' ###########################################################
' # ---------------------------------------------------------
' # getComputerName
' # [概要]
' # 　使用PCのコンピュータ名を取得
' # [In]
' #   ─
' # [Out]
' #   ComputerNameの取得値
' # ---------------------------------------------------------
' ###########################################################
Public Function getComputerName()

    'コンピュータ名を取得。文字列型
    getComputerName = Environ("COMPUTERNAME")

End Function

' ###########################################################
' # ---------------------------------------------------------
' # Get First Day
' # [概要]
' # 　月の初日を取得
' # [In]
' #   ─
' # [Out]
' #   YYYYMMDD
' # ---------------------------------------------------------
' ###########################################################
Public Function getBeginDayStr(inDate As Date)
    getBeginDayStr = Format(DateSerial(Year(inDate), Month(inDate), 1), "YYYYMMDD")
End Function

' ###########################################################
' # ---------------------------------------------------------
' # Get End Day
' # [概要]
' # 　月の最終日を取得
' # [In]
' #   ─
' # [Out]
' #   YYYYMMDD
' # ---------------------------------------------------------
' ###########################################################
Public Function getEndDayStr(inDate As Date)
    getEndDayStr = Format(DateSerial(Year(inDate), Month(inDate) + 1, 1) - 1, "YYYYMMDD")
End Function

' ###########################################################
' # ---------------------------------------------------------
' # Convert YYYYMMDD to Date Type
' # [概要]
' # YYYYMMDDを日付タイプに変換する。
' # [In]
' #   ─
' # [Out]
' #   Date Type
' # ---------------------------------------------------------
' ###########################################################

Public Function cnvDateFromYMD(inYYYYMMDD As String) As Date
    cnvDateFromYMD = CDate(Mid(inYYYYMMDD, 1, 4) & "/" & Mid(inYYYYMMDD, 5, 2) & "/" & Mid(inYYYYMMDD, 7, 2))
End Function

' ###########################################################
' # ---------------------------------------------------------
' # Convert YYYYMMDDhhnnss to Date Type
' # [概要]
' # YYYYMMDDhhnnssを日付時間タイプに変換する。
' # [In]
' #   ─
' # [Out]
' #   Date Type
' # ---------------------------------------------------------
' ###########################################################

Public Function cnvDateYMDhns(inYMDhns As String) As Date
    cnvDateYMDhns = CDate(Mid(inYMDhns, 1, 4) & "/" & Mid(inYMDhns, 5, 2) & "/" & Mid(inYMDhns, 7, 2) & " " & Mid(inYMDhns, 9, 2) & ":" & Mid(inYMDhns, 11, 2) & ":" & Mid(inYMDhns, 13, 2))
End Function

' ###########################################################
' # ---------------------------------------------------------
' # Array to Boolean Type
' # [概要]
' # 変数を配列かどうか判断する。
' # [In]
' #   ─
' # [Out]
' #   Boolean Type
' # ---------------------------------------------------------
' ###########################################################

Public Function isArrayEmpty(inArray As Variant) As Boolean

    If IsArray(inArray) Then
    
        isArrayEmpty = True
    Else
        isArrayEmpty = False
    End If
    
End Function

' ###########################################################
' # ---------------------------------------------------------
' # calAccYear
' # [概要]
' # 　会計年度を取得
' # [In]
' #   ─
' # [Out]
' #   会計年度の取得値
' # ---------------------------------------------------------
' ###########################################################
Public Function calAccYear(inDate As Date) As Integer

    '会計年度を取得。数値型
    calAccYear = Year(DateAdd("m", -3, inDate))
    
End Function

' ###########################################################
' # ---------------------------------------------------------
' # calAccMonth
' # [概要]
' # 　会計月を取得
' # [In]
' #   ─
' # [Out]
' #   会計月の取得値
' # ---------------------------------------------------------
' ###########################################################
Public Function calAccMonth(inDate As Date) As Integer

    '会計月を取得。数値型
    calAccMonth = Month(DateAdd("m", -3, inDate))
    
End Function

' ###########################################################
' # ---------------------------------------------------------
' # Set Current Directory
' # [概要]
' # 　カレントディレクトリーに設定する。
' # [In]
' #   ─
' # [Out]
' #   String
' # ---------------------------------------------------------
' ###########################################################
Public Sub setCurrentDir(inDirectoryNm As String)
    With CreateObject("WScript.Shell")
        .CurrentDirectory = inDirectoryNm
    End With
End Sub

' ###########################################################
' # ---------------------------------------------------------
' # Get File Name From File Path
' # [概要]
' # ファイルパスから拡張子を除いたファイル名を編集
' # [In]
' #   ─
' # [Out]
' #   String
' # ---------------------------------------------------------
' ###########################################################
Public Function GetFNmFromFPath(inPath As String) As String
    Dim FSO As Object
    Dim myFileNmExt As String
    Dim myFileDotPoint As Integer
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    myFileNmExt = FSO.GetFileName(inPath)
    myFileDotPoint = InStrRev(myFileNmExt, ".")
    
    GetFNmFromFPath = Left(myFileNmExt, myFileDotPoint - 1)
    
    Set FSO = Nothing
 End Function
 
' ###########################################################
' # --------------------------------------------------------
' # Copy Clipboard
' # [概要]
' # クリップボードにコピーする。
' # [In]
' #  Text
' # [Out]
' #   -
' # ---------------------------------------------------------
' ###########################################################
Public Sub copyClipboard(inText As Variant)
    Dim CB As New DataObject

    CB.SetText inText
    CB.PutInClipboard
    
    Set CB = Nothing
End Sub

' ###########################################################
' # --------------------------------------------------------
' # Add String Value
' # [概要]
' # 文字列の連結
' # [In]
' #  Text
' # [Out]
' #   -
' # ---------------------------------------------------------
' ###########################################################
Public Function addStrValue(inStrValue As String, inAddValue As String) As String
    If inStrValue = "" Then
        addStrValue = inAddValue
    Else
        addStrValue = inStrValue & ", " & inAddValue
    End If
End Function


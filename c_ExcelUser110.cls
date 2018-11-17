VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ExcelUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ###########################################################
' # ---------------------------------------------------------
' # Excel Worksheet Object Control Class
' # Ver.20181102    Upd updPivotDS 各OfficeVersionに対応
' # Ver.20181015    Upd mkDic, mkDicArray IndexをIntからLongに変更
' # Ver.20181004    Upd mkDicArray 複数キーを指定出来る様に変更
' # Ver.20181004    Upd mkDic 複数キーを指定出来る様に変更
' # Ver.20181001    Upd mkDic 行Noをデータに持つ機能を追加
' # Ver.20180927    Add duplicateAdd, duplicateDelete, openUserNo
' # Ver.20180927    Upd AddSortField, RunSort
' # Ver.20180920    Add delRow
' # Ver.20180912    Add mkDicArray, getDicDataArray, removeDic
' # Ver.20180905    Add addWorkBook
' # Ver.20180828-3  Add pasteData
' # Ver.20180828-2  Add updPivotDS
' # Ver.20180828    Add mkTable
' # Ver.20180827    Upd copyData, pasteValueData
' # Ver.20180823-2  Upd resetMaxNo - Column名の保存時に、改行コードをブランクに置き換え。
' # Ver.20180823    Upd calRangeRowNoの初期チェックを追加。
' # Ver.20180822-2  Upd openWorkBookのOn Errorの位置を修正。
' #                 Add copyFormulaRow
' #                 Add chgShtNm
' # Ver.20180822    Upd calRangeRowNo - データが存在しない時は、xRangeRowCntを0にする。
' # Ver.20180820    Upd CloseWorkBook
' # Ver.20180820    Created
' # ---------------------------------------------------------
' ###########################################################

Public WB As Workbook
Public WS As Worksheet
Dim Dic As Object

Dim xMaxRowNo As Long
Dim xMaxColNo As Integer
Dim xStartColNo As Integer
Dim xStartRowNo As Integer
Dim xColNm() As String
Dim xMaxRowCkColNo As Integer

Dim xRangeRowNoFrom As Long
Dim xRangeRowNoTo As Long
Dim xRangeRowCnt As Long

Dim xColNoArray() As Variant

Public Enum eColorIndex
    Black = 1
    White = 2
    Red = 3
    BrightGreen = 4
    Blue00 = 5
    xYellow = 6
    Pink00 = 7
    Turquoise00 = 8
    DarkRed00 = 9
    Green = 10
    DarkBlue00 = 11
    DarkYellow = 12
    Violet00 = 13
    Teal00 = 14
    Gray25 = 15
    Gray50 = 16
    Periwinkle = 17
    Plum = 18
    Ivory = 19
    LightTurquoise = 20
    DarkPurple = 21
    Coral = 22
    OceanBlue = 23
    IceBlue = 24
    DarkBlue = 25
    Pink = 26
    Yellow = 27
    Turquoise = 28
    Violet = 29
    DarkRed = 30
    Teal = 31
    Blue = 32
    SkyBlue = 33
    LightGreen = 35
    LightYellow = 36
    PaleBlue = 37
    Rose = 38
    Lavender = 39
    Tan = 40
    LightBlue = 41
    Aqua = 42
    Lime = 43
    gold = 44
    LightOrange = 45
    Orange = 46
    BlueGray = 47
    Gray40 = 48
    DarkTeal = 49
    SeaGreen = 50
    DarkGreen = 51
    Brown = 53
    Indigo = 55
    Gray80 = 56
End Enum

' Max Row Noの取得
Public Property Get MaxRowNo() As Long
    MaxRowNo = xMaxRowNo
End Property

' Max Column Noの取得
Public Property Get MaxColNo() As Long
    MaxColNo = xMaxColNo
End Property

' 開始Column Noを取得
Public Property Get StartColNo() As Integer
    StartColNo = xStartColNo
End Property

' 開始Row Noを取得
Public Property Get StartRowNo() As Long
    StartRowNo = xStartRowNo
End Property

' 条件レンジ開始Row No Fromを取得
Public Property Get RangeRowNoFrom() As Long
    RangeRowNoFrom = xRangeRowNoFrom
End Property

' 条件レンジ開始Row No Toを取得
Public Property Get RangeRowNoTo() As Long
    RangeRowNoTo = xRangeRowNoTo
End Property

' 条件レンジ開始Row Countを取得
Public Property Get RangeRowCnt() As Long
    RangeRowCnt = xRangeRowCnt
End Property

' 条件レンジの計算
Public Sub calRangeRowNo(inColNo As Integer, inValueFrom As Variant, inValueTo As Variant)
    Dim i As Long

    xRangeRowNoFrom = 0
    xRangeRowNoTo = 0

    If WS.Cells(xStartRowNo + 1, inColNo).Value > inValueTo Then
        xRangeRowCnt = 0
        Exit Sub
    End If
    
    If WS.Cells(xMaxRowNo, inColNo).Value < inValueFrom Then
        xRangeRowCnt = 0
        Exit Sub
    End If

    For i = xStartRowNo + 1 To xMaxRowNo
        If WS.Cells(i, inColNo).Value >= inValueFrom Then
            xRangeRowNoFrom = i
            Exit For
        End If
    Next i
            
    If xRangeRowNoFrom = 0 Then
        xRangeRowCnt = 0
        Exit Sub
    End If
            
    For i = xMaxRowNo To xStartRowNo + 1 Step -1
        If WS.Cells(i, inColNo).Value <= inValueTo Then
            xRangeRowNoTo = i
            Exit For
        End If
    Next i
    
    If xRangeRowNoTo = 0 Then
        xRangeRowCnt = 0
        xRangeRowNoFrom = 0
        Exit Sub
    End If
    
    xRangeRowCnt = xRangeRowNoTo - xRangeRowNoFrom + 1
End Sub

' Work Bookを開く
' Parameter - inWorkBookPath : ブランクならば開いているExcelをインスタンスにする。
Public Function openWorkBook(inWorkBookPath As String) As Boolean
    On Error GoTo Exception
    If inWorkBookPath = "" Then
        Set WB = ThisWorkbook
    Else
        Workbooks.Open inWorkBookPath
        Set WB = Workbooks(Dir(inWorkBookPath))
    End If
    On Error GoTo 0
    
    openWorkBook = True
    
    Exit Function
    
Exception:
    outMsg "L81", "Excel Open Work Book Error. Description = " & Err.Description
    
    openWorkBook = False
End Function

' 新しいWorkbookを作成する。
Public Function addWorkBook() As Boolean
    On Error GoTo Exception
        Set WB = Workbooks.Add
    On Error GoTo 0
    
    addWorkBook = True
    
    Exit Function
    
Exception:
    outMsg "L87", "Excel Add Work Book Error. Description = " & Err.Description
    
    addWorkBook = False
End Function

' すでに開いているWork Bookからクラスをインスタンス化する。クラスはシート単位の為。
Public Function setWorkBook(inWorkBookObj As Workbook) As Boolean
    On Error GoTo Exception
    
    Set WB = inWorkBookObj
    
    On Error GoTo 0
    
    setWorkBook = True
    
    Exit Function
    
Exception:
    setWorkBook = False
End Function

' Work Seetをセットする。Max値、Col名-Noの対応表を作成
Public Function setWorkSht(inWorkShtNm As String, inStartRowNo As Integer, inStartColNo As Integer, Optional inMaxRowCkColNo As Variant) As Boolean
    On Error GoTo Exception
    
    If inWorkShtNm = "" Then
        Set WS = WB.Worksheets(1)
    Else
        Set WS = WB.Worksheets(inWorkShtNm)
    End If
    
    On Error GoTo 0
    
    xStartColNo = inStartColNo
    xStartRowNo = inStartRowNo
    If IsMissing(inMaxRowCkColNo) Then
        xMaxRowCkColNo = xStartColNo
    Else
        xMaxRowCkColNo = inMaxRowCkColNo
    End If
    
    resetMaxNo
    
    setWorkSht = True
    
    Exit Function
    
Exception:
    outMsg "L82", "Excel Open Work Sheet Description = " & Err.Description
    
    setWorkSht = False
End Function

' Max値のリセット、Col名-Noの対応表の再作成
Public Sub resetMaxNo()
    Dim i As Integer

    If WS.Cells(xStartRowNo, xStartColNo) = "" Then
        xMaxColNo = 0
        xMaxRowNo = 0
    Else
        If WS.Cells(xStartRowNo, xStartColNo + 1) = "" Then
            xMaxColNo = xStartColNo
        Else
            xMaxColNo = WS.Cells(xStartRowNo, xStartColNo).End(xlToRight).Column
        End If
        
        If WS.Cells(xStartRowNo + 1, xMaxRowCkColNo) = "" Then
            xMaxRowNo = xStartRowNo
        Else
            xMaxRowNo = WS.Cells(xStartRowNo, xMaxRowCkColNo).End(xlDown).Row
        End If
    End If
    
    ReDim xColNm(xMaxColNo)
    For i = xStartColNo To xMaxColNo
        xColNm(i) = Replace(WS.Cells(xStartRowNo, i), vbLf, "")
    Next i
End Sub

' 開いているExcel Bookに対して名前を付けて保存する。
Public Function SaveAsWorkBook(inFilePath As String, Optional inMacroEnable As Variant) As Boolean
    On Error GoTo Exception
    If IsMissing(inMacroEnable) Then
        WB.SaveAs FileName:=inFilePath, FileFormat:=xlWorkbookDefault
    Else
        If inMacroEnable Then
            WB.SaveAs FileName:=inFilePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Else
            WB.SaveAs FileName:=inFilePath, FileFormat:=xlWorkbookDefault
        End If
    End If
    On Error GoTo 0
    
    SaveAsWorkBook = True
    
    Exit Function
    
Exception:
    outMsg "L83", "Excel Save Aa File Error. Description = " & Err.Description
    
    SaveAsWorkBook = False
End Function

' Work Sheetの追加
Public Function addWorkSht(inWorkShtNm As String, inStartRowNo As Integer, inStartColNo As Integer, Optional inMaxRowCkColNo As Variant) As Boolean
    On Error GoTo Exception
    Set WS = WB.Worksheets.Add
    On Error GoTo 0
    
    WS.Name = inWorkShtNm
    
    xStartColNo = inStartColNo
    xStartRowNo = inStartRowNo
    If IsMissing(inMaxRowCkColNo) Then
        xMaxRowCkColNo = xStartColNo
    Else
        xMaxRowCkColNo = inMaxRowCkColNo
    End If
    
    resetMaxNo

    addWorkSht = True
    
    Exit Function

Exception:
    addWorkSht = False
End Function

' Column名称より、Column Noを取得する。
Public Function getColNo(inColNm As String) As Integer
    Dim i As Long

    For i = xStartColNo To xMaxColNo
        If inColNm = xColNm(i) Then
            getColNo = i
            Exit Function
        End If
    Next i
    
    outMsg "L84", "Excel Get Column No Error. Name = " & inColNm
    Debug.Print "Excel Get Column No Error. Name = " & inColNm
    
    getColNo = 1
End Function

' シートの全クリア
Public Sub clrAllData()
    WS.Cells.Clear
End Sub

' Work Bookを保存する。
Public Function saveWorkBook() As Boolean
    On Error GoTo Exception
    WB.Save
    On Error GoTo 0
    
    saveWorkBook = True

    Exit Function
    
Exception:
    outMsg "L85", "Excel Save Error. Description = " & Err.Description
    
    saveWorkBook = False
End Function

' Work BookをCloseする。Saveはしない。
Public Function closeWorkBook() As Boolean
    On Error GoTo Exception
    WB.Close False
    On Error GoTo 0
    
    closeWorkBook = True
    
    Exit Function

Exception:
    outMsg "L86", "Excel Close Error. Description = " & Err.Description
    
    closeWorkBook = False
End Function

' 空白の行挿入を行う。
Public Sub insRow(inStartRowNo As Long, inInsCnt As Long)
    WS.Rows(inStartRowNo & ":" & inStartRowNo + inInsCnt - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub

' シート一括でフォーマットのコピー、ペーストを行う。
Public Sub copyFormatRow(inCopyFromRowNo As Long, inCopyToStartRowNo As Long, inCopyCnt As Long)
    WS.Range(WS.Cells(inCopyFromRowNo, xStartColNo), WS.Cells(inCopyFromRowNo, xMaxColNo)).Copy
    WS.Range(WS.Cells(inCopyToStartRowNo, xStartColNo), WS.Cells(inCopyToStartRowNo + inCopyCnt - 1, xMaxColNo)).PasteSpecial paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub

' カラム単位のフォーミュラーのコピー、ペーストを行う。
Public Sub copyFormulaRow(inColNo As Integer, inCopyFromRowNo As Long, inCopyToStartRowNo As Long, inCopyCnt As Long)
    WS.Range(WS.Cells(inCopyFromRowNo, inColNo), WS.Cells(inCopyFromRowNo, inColNo)).Copy
    WS.Range(WS.Cells(inCopyToStartRowNo, inColNo), WS.Cells(inCopyToStartRowNo + inCopyCnt - 1, inColNo)).PasteSpecial paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub

' カラム単位の入力規則のコピー、ペーストを行う。
Public Sub copyValidationRow(inColNo As Integer, inCopyFromRowNo As Long, inCopyToStartRowNo As Long, inCopyCnt As Long)
    WS.Range(WS.Cells(inCopyFromRowNo, inColNo), WS.Cells(inCopyFromRowNo, inColNo)).Copy
    WS.Range(WS.Cells(inCopyToStartRowNo, inColNo), WS.Cells(inCopyToStartRowNo + inCopyCnt - 1, inColNo)).PasteSpecial paste:=xlPasteValidation, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub

' カラムFrom-Toのフォーミュラーのコピー、ペーストを行う。
' ※コピーした同じセルからペースト出来ない。エラー「配列の一部を変更できません」となる。
Public Sub copyFormulaRow2(inStartColNo As Integer, inEndColNo As Integer, inCopyFromRowNo As Long, inCopyToStartRowNo As Long, inCopyCnt As Long)
    WS.Range(WS.Cells(inCopyFromRowNo, inStartColNo), WS.Cells(inCopyFromRowNo, inEndColNo)).Copy
    WS.Range(WS.Cells(inCopyToStartRowNo, inStartColNo), WS.Cells(inCopyToStartRowNo + inCopyCnt - 1, inStartColNo)).PasteSpecial paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub

' コピーを行う。
Public Sub copyData(inCopyStartRowNo As Long, inCopyStartColNo As Integer, Optional inCopyEndRowNo As Variant, Optional inCopyEndColNo As Variant)
    If IsMissing(inCopyEndRowNo) Then
'        WS.Range(WS.Cells(inCopyStartRowNo, inCopyStartColNo), WS.Cells(inCopyStartRowNo, inCopyStartColNo)).Copy
        WS.Cells(inCopyStartRowNo, inCopyStartColNo).Copy
    Else
        WS.Range(WS.Cells(inCopyStartRowNo, inCopyStartColNo), WS.Cells(inCopyEndRowNo, inCopyEndColNo)).Copy
    End If
End Sub

' 値のペーストを行う。
Public Sub pasteValueData(inPasteStartRowNo As Long, inPasteStartColNo As Integer, Optional inPasteEndRowNo As Variant, Optional inPasteEndColNo As Variant)
    If IsMissing(inPasteEndRowNo) Then
        WS.Cells(inPasteStartRowNo, inPasteStartColNo).PasteSpecial paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Else
        WS.Range(WS.Cells(inPasteStartRowNo, inPasteStartColNo), WS.Cells(inPasteEndRowNo, inPasteEndColNo)).PasteSpecial paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End If
End Sub

' すべてのペーストを行う。
Public Sub pasteData(inPasteStartRowNo As Long, inPasteStartColNo As Integer, Optional inPasteEndRowNo As Variant, Optional inPasteEndColNo As Variant)
    If IsMissing(inPasteEndRowNo) Then
        WS.Cells(inPasteStartRowNo, inPasteStartColNo).paste
    Else
        WS.Range(WS.Cells(inPasteStartRowNo, inPasteStartColNo), WS.Cells(inPasteEndRowNo, inPasteEndColNo)).PasteSpecial paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End If
End Sub

' 指定されたColumnの最大値を求める。
Public Function getMaxValue(inRowNoFrom As Long, inRowNoTo As Long, inColNo As Integer) As Variant
    getMaxValue = WorksheetFunction.Max(WS.Range(WS.Cells(inRowNoFrom, inColNo), WS.Cells(inRowNoTo, inColNo)))
End Function

' Dictionaryの生成、キー・データの入れ込みを行う。inDataColNoに0を設定すると、Dataに行Noを設定する。
Public Sub mkDic(inKeyColNo As Variant, inDataColNo As Integer)
    Dim i As Long
    Dim j As Integer
    Dim KeyBuf As String
    Dim DataBuf As String
    Dim KeyColNo As Variant
    Dim KeyCnt As Integer
    
    KeyColNo = Split(inKeyColNo, ",")
    KeyCnt = UBound(KeyColNo)
    
    Set Dic = CreateObject("Scripting.Dictionary")
    
    For i = xMaxRowNo To xStartRowNo + 1 Step -1
        KeyBuf = ""
        For j = 0 To KeyCnt
            KeyBuf = KeyBuf & WS.Cells(i, CInt(KeyColNo(j))).Value
        Next j
        
        If inDataColNo = 0 Then
            DataBuf = i
        Else
            DataBuf = WS.Cells(i, inDataColNo).Value
        End If
        
        If Dic.Exists(KeyBuf) Then
           'continue
        Else
            Dic.Add KeyBuf, DataBuf
        End If
    Next i
End Sub

' Dictionaryからパラメータのキー情報を元にデータを取得する。
Public Function getDicData(inKeyValue As String) As String
    getDicData = Dic.item(inKeyValue)
End Function

' Dictionaryに存在するかどうかチェック
Public Function ckDicExist(inKeyValue As String) As Boolean
    ckDicExist = Dic.Exists(inKeyValue)
End Function

Public Sub removeDic()
    Dic.RemoveAll
End Sub

' Dictionaryの生成、キー・データの入れ込みを行う。
Public Sub mkDicArray(inKeyColNo As Variant, inDataColNoStr As String)
    Dim i As Long
    Dim j As Integer
    Dim KeyBuf As String
    Dim DataColNo As Variant
    Dim DataBuf() As String
    Dim KeyColNo As Variant
    Dim KeyCnt As Integer
    
    KeyColNo = Split(inKeyColNo, ",")
    KeyCnt = UBound(KeyColNo)

    Set Dic = CreateObject("Scripting.Dictionary")
    
    DataColNo = Split(inDataColNoStr, ",")
    ReDim DataBuf(UBound(DataColNo))
        
    For i = xMaxRowNo To xStartRowNo + 1 Step -1
        KeyBuf = ""
        For j = 0 To KeyCnt
            KeyBuf = KeyBuf & WS.Cells(i, CInt(KeyColNo(j))).Value
        Next j
                
        If Dic.Exists(KeyBuf) Then
           'continue
        Else
            For j = 0 To UBound(DataBuf)
                DataBuf(j) = WS.Cells(i, CInt(DataColNo(j))).Value
            Next j
        
            Dic.Add KeyBuf, DataBuf
        End If
    Next i
End Sub

Public Function getDicDataArray(inKeyValue As String, inDataNo As Integer) As String
    Dim ItemData As Variant

    ItemData = Dic.item(inKeyValue)
    
    getDicDataArray = ItemData(inDataNo - 1)
End Function

' Clear Sort Area
Public Sub clearSortArea()
    WS.Sort.SortFields.Clear
End Sub

' Sort Fieldの追加
Public Sub addSortField(inColNo As Integer, Optional inStartRowNo As Variant, Optional inEndRowNo As Variant)
    Dim StartRowNo As Long
    Dim EndRowNo As Long
    
    If IsMissing(inStartRowNo) Then
        StartRowNo = xStartRowNo + 1
        EndRowNo = xMaxRowNo
    Else
        StartRowNo = CLng(inStartRowNo)
        EndRowNo = CLng(inEndRowNo)
    End If

    WS.Sort.SortFields.Add Key:=Range(Cells(StartRowNo, inColNo), Cells(EndRowNo, inColNo)) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
End Sub

' Sortの実行
Public Sub runSort(Optional inStartRowNo As Variant, Optional inEndRowNo As Variant)
    Dim StartRowNo As Long
    Dim EndRowNo As Long
    Dim hasHeader As Variant
    
    If IsMissing(inStartRowNo) Then
        StartRowNo = xStartRowNo
        EndRowNo = xMaxRowNo
        hasHeader = xlYes
    Else
        StartRowNo = CLng(inStartRowNo)
        EndRowNo = CLng(inEndRowNo)
        hasHeader = xlNo
    End If
    
    With WS.Sort
        .SetRange Range(Cells(StartRowNo, xStartColNo), Cells(EndRowNo, xMaxColNo))
        .Header = hasHeader
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

' シート名の変更
Public Function chgShtNm(inShtNm As String) As Boolean
    On Error GoTo Exception
    Worksheets(1).Name = inShtNm
    On Error GoTo 0
    
    chgShtNm = True
    
    Exit Function
Exception:
    chgShtNm = False
End Function

' Excelのテーブルを作成する。
Public Sub mkTable(inTableNm As String, inTableStartRowNo As Long, inTableStartColNo As Integer, inTableEndRowNo As Long, inTableEndColNo As Integer)
    WS.ListObjects.Add(SourceType:=xlSrcRange, Source:=WS.Range(WS.Cells(inTableStartRowNo, inTableStartColNo), WS.Cells(inTableEndRowNo, inTableEndColNo))).Name = inTableNm
    WS.ListObjects(inTableNm).TableStyle = "TableStyleLight2"
End Sub

' Pivotのデータセット名を変更し、Pivotの更新を行う。
Public Sub updPivotDS(inPivotTblNm As String, inPivotDSNm As String)
    Dim ExcelVer As Integer
    
    Select Case getExcelVer()
        Case "Excel2016"
            ExcelVer = 6
        Case "Excel2013"
            ExcelVer = 5
        Case "Excel2010"
            ExcelVer = 4
        Case Else
            ExcelVer = 3
    End Select
    
    WS.PivotTables(inPivotTblNm).ChangePivotCache WB.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=inPivotDSNm, Version:=ExcelVer)
    WS.PivotTables(inPivotTblNm).PivotCache.Refresh
End Sub

' Copy All Sheet Data
Public Sub copyAllShtData()
    WS.Cells.Copy
End Sub

' Paste All Sheet Data
Public Sub pasteAllShtData()
    WS.paste
End Sub

' Fit All Column
Public Sub fitAllCol()
    WS.Cells.EntireColumn.AutoFit
End Sub

Public Sub delCol(inColNo As Integer)
    WS.Columns(inColNo).Delete
End Sub

Public Sub insCol(inColNo As Integer)
    WS.Columns(inColNo).Insert
End Sub

Public Sub delRow(inRowNo As Long)
    WS.Rows(inRowNo).Delete
End Sub

'重複削除項目取得
Public Sub duplicateAdd(inColNo As Integer)
    Dim i As Integer
    
    If (Not xColNoArray) = -1 Then
        ReDim xColNoArray(0) As Variant
        xColNoArray(0) = inColNo
        Debug.Print xColNoArray(0)
'        ReDim Preserve xColNoArray(UBound(xColNoArray) + 1) As Variant
    Else
        i = UBound(xColNoArray) + 1
        ReDim Preserve xColNoArray(i) As Variant
        xColNoArray(i) = inColNo
        Debug.Print xColNoArray(UBound(xColNoArray))
    End If
    
    
End Sub

'重複削除duplicateDelete
Public Sub duplicateDelete()

    WS.Range(WS.Cells(xStartRowNo, xStartColNo), WS.Cells(xMaxRowNo, xMaxColNo)).RemoveDuplicates Columns:=(xColNoArray), Header:=xlYes

End Sub
'Excel利用ユーザー数
Public Function openUserNo()
    Dim users
    users = WB.UserStatus
   
    openUserNo = UBound(users)

End Function


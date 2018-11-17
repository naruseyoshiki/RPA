Attribute VB_Name = "m_Parameter"
Option Explicit
' Connectionの設定値を取得する。
Public Function get_Connection() As String

    Dim WS                  As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("コントロール")
    get_Connection = WS.Cells(4, 9).Value
    Set WS = Nothing

End Function

' Clientの設定値を取得する。
Public Function get_Client() As String

    Dim WS                  As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("コントロール")
    get_Client = WS.Cells(5, 9).Value
    Set WS = Nothing

End Function

' Userの設定値を取得する。
Public Function get_User() As String

    Dim WS                  As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("コントロール")
    get_User = WS.Cells(6, 9).Value
    Set WS = Nothing

End Function

' Passwordの設定値を取得する。
Public Function get_Password() As String

    Dim WS                  As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("コントロール")
    get_Password = WS.Cells(7, 9).Value
    Set WS = Nothing

End Function

' Passwordの設定値を取得する。
Public Sub clear_Password()
Attribute clear_Password.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim WS                  As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("コントロール")
    WS.Cells(7, 9).Value = ""
    Set WS = Nothing

End Sub

' Userの設定値を取得する。
Public Function get_Language() As String

    Dim WS                  As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("コントロール")
    get_Language = WS.Cells(8, 9).Value
    Set WS = Nothing

End Function

' タイムアウト設定値を取得する。
Public Function get_Timeout(w_ii As Integer) As Integer

    Dim WS                  As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("コントロール")
    get_Timeout = WS.Cells(11 + w_ii, 9).Value
    Set WS = Nothing

End Function

' Get Operation Type
Public Function getOpeType() As String
    Dim WS As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("コントロール")
    getOpeType = WS.Cells(4, 5).Value
    
    Set WS = Nothing
End Function

' Get Log Folder
Public Function getLogFolder() As String
    Dim WS As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("コントロール")
    getLogFolder = WS.Cells(5, 5).Value
    
    Set WS = Nothing
End Function

' Get Log File Name
Public Function getLogFileNm() As String
    Dim WS As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("コントロール")
    getLogFileNm = WS.Cells(6, 5).Value
    
    Set WS = Nothing
End Function


Attribute VB_Name = "m_Parameter"
Option Explicit
' Connection�̐ݒ�l���擾����B
Public Function get_Connection() As String

    Dim WS                  As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("�R���g���[��")
    get_Connection = WS.Cells(4, 9).Value
    Set WS = Nothing

End Function

' Client�̐ݒ�l���擾����B
Public Function get_Client() As String

    Dim WS                  As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("�R���g���[��")
    get_Client = WS.Cells(5, 9).Value
    Set WS = Nothing

End Function

' User�̐ݒ�l���擾����B
Public Function get_User() As String

    Dim WS                  As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("�R���g���[��")
    get_User = WS.Cells(6, 9).Value
    Set WS = Nothing

End Function

' Password�̐ݒ�l���擾����B
Public Function get_Password() As String

    Dim WS                  As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("�R���g���[��")
    get_Password = WS.Cells(7, 9).Value
    Set WS = Nothing

End Function

' Password�̐ݒ�l���擾����B
Public Sub clear_Password()
Attribute clear_Password.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim WS                  As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("�R���g���[��")
    WS.Cells(7, 9).Value = ""
    Set WS = Nothing

End Sub

' User�̐ݒ�l���擾����B
Public Function get_Language() As String

    Dim WS                  As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("�R���g���[��")
    get_Language = WS.Cells(8, 9).Value
    Set WS = Nothing

End Function

' �^�C���A�E�g�ݒ�l���擾����B
Public Function get_Timeout(w_ii As Integer) As Integer

    Dim WS                  As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("�R���g���[��")
    get_Timeout = WS.Cells(11 + w_ii, 9).Value
    Set WS = Nothing

End Function

' Get Operation Type
Public Function getOpeType() As String
    Dim WS As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("�R���g���[��")
    getOpeType = WS.Cells(4, 5).Value
    
    Set WS = Nothing
End Function

' Get Log Folder
Public Function getLogFolder() As String
    Dim WS As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("�R���g���[��")
    getLogFolder = WS.Cells(5, 5).Value
    
    Set WS = Nothing
End Function

' Get Log File Name
Public Function getLogFileNm() As String
    Dim WS As Worksheet
    
    Set WS = ThisWorkbook.Worksheets("�R���g���[��")
    getLogFileNm = WS.Cells(6, 5).Value
    
    Set WS = Nothing
End Function


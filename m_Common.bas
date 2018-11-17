Attribute VB_Name = "m_Common"
Option Explicit
' ###########################################################
' # ---------------------------------------------------------
' # Common Module
' #
' # Version : 20181024 Upd outMsg,Right(getLogFolder, 1)<>"\" ���O�t�H���_\�m�F
' # Version : 20181017 Add setCurrentDir,GetFNmFromFPath, copyClipboard, addStrValue
' # Version : 20181017 Add calAccYear, calAccMonth ��v�N�x�A��v����(��)���擾
' #
' # ---------------------------------------------------------
' ###########################################################

' ###########################################################
' # ---------------------------------------------------------
' # get_DeskTopPath
' # [�T�v]
' # �@�f�X�N�g�b�v�p�X���擾����B
' # [In]
' #   ��
' # [Out]
' #   �f�X�N�g�b�v�p�X�ݒ�l
' # ---------------------------------------------------------
' ###########################################################
Public Function get_DeskTopPath() As String
    
    Dim WSH As Variant
    
    Set WSH = CreateObject("Wscript.Shell") 'WSH�I�u�W�F�N�g
    get_DeskTopPath = WSH.SpecialFolders("Desktop") & "\"
    Set WSH = Nothing '�I�u�W�F�N�g���
    
End Function

' ###########################################################
' # ---------------------------------------------------------
' # outMsg
' # [�T�v]
' # �@���b�Z�[�W���擾
' # [In]
' #   ��
' # [Out]
' #   �G���[�R�[�h�{���b�Z�[�W���e
' # ---------------------------------------------------------
' ###########################################################
Public Function outMsg(inCode As String, inMsg As String) As Integer
    If getOpeType() = "Manual" Then
        '�e���b�Z�[�W���R�[�h�̓������\��
        Select Case Left(inCode, 1)
            Case "E"
                MsgBox inCode & " : " & inMsg, vbCritical, "�G���["
            Case "S"
                MsgBox inCode & " : " & inMsg, vbExclamation, "�V�X�e���G���["
            Case "I"
                MsgBox inCode & " : " & inMsg, vbInformation, "�C���t�H���[�V����"
            Case "Q"
                outMsg = MsgBox(inCode & ":" & inMsg, vbQuestion + vbYesNo, "�m�F")
            Case "L"
                ' Continue ���O��p
            Case Else
                MsgBox inCode & "" & inMsg, vbCritical, "�s���ȃG���["
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
    
    ' ���O�t�H���_�̗L���`�F�b�N�A���Ȃ�t�@�C�����������B
    If Right(getLogFolder, 1) <> "\" Then
        GoTo Exception1
    End If
        
    ' ���O�̏o��
    On Error GoTo Exception1
    Open getLogFolder & getLogFileNm For Append As #1
    Print #1, Format(Now, "YYYY/MM/DD hh:nn:ss") & vbTab & inCode & " : " & Replace(Replace(Replace(inMsg, vbCrLf, " "), "��", ""), "     ", "")
    Close #1
    On Error GoTo 0
    
    Exit Function
Exception1:
    Debug.Print "Function=outMsg, Log Output Error!, �yLog Folder�z = " & getLogFolder & " �yLog File Name�z= " & getLogFileNm
    
End Function

' ###########################################################
' # ---------------------------------------------------------
' # waitSec
' # [�T�v]
' # �@��ʑJ�ڎ��̑ҋ@���Ԃ��w��
' # [In]
' #   ��
' # [Out]
' # �@thisSec�ݒ�l
' # ---------------------------------------------------------
' ###########################################################

Public Sub waitSec(Optional inTime As Variant)
    
    Dim thisSec As Variant
    
    '�����ȗ��̐^�U�m�F�Bture�Ȃ�f�t�H���g1�b�ɂ���B
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
' # [�T�v]
' # �@�g�p����Excel�̃o�[�W�������擾(ver2010�܂�)
' # [In]
' #   ��
' # [Out]
' #   ExcelVer�̎擾�l
' # ---------------------------------------------------------
' ###########################################################

Public Function getExcelVer() As String

    'Excel�o�[�W�����̎擾
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
' # [�T�v]
' # �@�g�pPC�̃R���s���[�^�����擾
' # [In]
' #   ��
' # [Out]
' #   ComputerName�̎擾�l
' # ---------------------------------------------------------
' ###########################################################
Public Function getComputerName()

    '�R���s���[�^�����擾�B������^
    getComputerName = Environ("COMPUTERNAME")

End Function

' ###########################################################
' # ---------------------------------------------------------
' # Get First Day
' # [�T�v]
' # �@���̏������擾
' # [In]
' #   ��
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
' # [�T�v]
' # �@���̍ŏI�����擾
' # [In]
' #   ��
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
' # [�T�v]
' # YYYYMMDD����t�^�C�v�ɕϊ�����B
' # [In]
' #   ��
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
' # [�T�v]
' # YYYYMMDDhhnnss����t���ԃ^�C�v�ɕϊ�����B
' # [In]
' #   ��
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
' # [�T�v]
' # �ϐ���z�񂩂ǂ������f����B
' # [In]
' #   ��
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
' # [�T�v]
' # �@��v�N�x���擾
' # [In]
' #   ��
' # [Out]
' #   ��v�N�x�̎擾�l
' # ---------------------------------------------------------
' ###########################################################
Public Function calAccYear(inDate As Date) As Integer

    '��v�N�x���擾�B���l�^
    calAccYear = Year(DateAdd("m", -3, inDate))
    
End Function

' ###########################################################
' # ---------------------------------------------------------
' # calAccMonth
' # [�T�v]
' # �@��v�����擾
' # [In]
' #   ��
' # [Out]
' #   ��v���̎擾�l
' # ---------------------------------------------------------
' ###########################################################
Public Function calAccMonth(inDate As Date) As Integer

    '��v�����擾�B���l�^
    calAccMonth = Month(DateAdd("m", -3, inDate))
    
End Function

' ###########################################################
' # ---------------------------------------------------------
' # Set Current Directory
' # [�T�v]
' # �@�J�����g�f�B���N�g���[�ɐݒ肷��B
' # [In]
' #   ��
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
' # [�T�v]
' # �t�@�C���p�X����g���q���������t�@�C������ҏW
' # [In]
' #   ��
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
' # [�T�v]
' # �N���b�v�{�[�h�ɃR�s�[����B
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
' # [�T�v]
' # ������̘A��
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


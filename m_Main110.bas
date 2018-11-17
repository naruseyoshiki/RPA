Attribute VB_Name = "m_Main"
Option Explicit

Public Sub MainProc()
Attribute MainProc.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim SAPTran As New c_SAPAccess
    Dim i As Integer
    
    ' Script Start Log
    outMsg "L01", "*** Script Cost110 DL_Started. ***"
    
    '�p�����[�^�̐ݒ�
    SAPTran.setParaVal "�R���g���[��", 1, 2
    
    'Excel���͗��`�F�b�N
    If SAPTran.ckParaVal("����敪*@2,����敪*@3,����敪*@4,����敪_2@5,�o���A���g��*@2,�o���A���g��*@3,�o���A���g��*@4,�o���A���g��_2@5,�o���A���g��*@2,�o���A���g��*@3,�o���A���g��*@4,�o���A���g��_2@5") = False Then
        Exit Sub
    End If
        
    If outMsg("Q01", "�������J�n���܂���?") = vbNo Then
        outMsg "L02", "�������L�����Z������܂����B"
        Exit Sub
    End If
    
    ' Excel���b�Z�[�W��\���ɐݒ肷��
    Application.DisplayAlerts = False
    
    ' Save File Path���N���A
    For i = 1 To 9
        SAPTran.putParaVal "*���㌴�����וۑ��ꏊ���ʁi�t���p�X�j", "", i
    Next i
    
    For i = 1 To 3
        SAPTran.putParaVal "*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", "", i
    Next i

    ' ���O�I��SAP
    SAPTran.Connection = get_Connection ' Connection�ݒ�
    SAPTran.Client = get_Client '�N���C�A���g�ݒ�
    SAPTran.User = get_User       '���[�U�ݒ�
    SAPTran.Password = get_Password '�p�X���[�h�ݒ�
    SAPTran.Language = get_Language    '����ݒ�
    If SAPTran.LogonSAP() > 0 Then '�G���[�Ȃ珈�����I��
        Exit Sub
    End If
    
    ' �w��g�����U�N�V������ݒ聨���s����
    SAPTran.TranCd = "SQ01"
    SAPTran.setTranCd
    
    ' ���[�U�O���[�v��ʂ̑J�ڊm�F
    Select Case SAPTran.ckScreenTransition("wnd[0]", "*���[�U�O���[�v*", get_Timeout(1))
        Case 0
            ' OK : Continue
        Case 1
            outMsg "E01", "�u���[�U�O���[�v�v��ʂ̑J�ڂ����s���܂����B" & vbCrLf & "�g�����U�N�V�����R�[�h = SQ01"
            Exit Sub
        Case 2
            outMsg "E02", "���[�U�O���[�v�@�^�C���I�[�o�["
            Exit Sub
        Case Else
            outMsg "S01", "System Error!"
            Exit Sub
    End Select
    
    With SAPTran.SAPSesi
    
    ' ���[�U�O���[�v��ʂ̑���
    .findById("wnd[0]").sendVKey 19 'Shift+F7, ���[�U�O���[�v����I����ʂ̕\��
    
    waitSec '1�b�҂�
   
    '�o���A���g��ʂ̑���
    If SAPTran.getGridLineNo("wnd[1]/usr/cntlGRID1/shellcont/shell", "DBGBNUM", "PRISM03") = 0 Then
        outMsg "E03", "�N�G�������݂��܂���B " & "PRISM03"
        Exit Sub
    End If
    .findById("wnd[1]").sendVKey 2
    
    ' ���[�U�O���[�v��ʂ̑���
    If SAPTran.getGridLineNo("wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell", "QNUM", SAPTran.getParaVal("�g�����U�N�V����(SQ01)")) = 0 Then
        outMsg "E04", "�g�����U�N�V���������݂��܂���B" & vbCrLf & "�g�����U�N�V���� = " & SAPTran.getParaVal("�g�����U�N�V����(SQ01)")
        Exit Sub
    End If
    .findById("wnd[0]").sendVKey 8
    
    '�������烋�[�v
    For i = 1 To 9
        
        If 5 = i Then
            .findById("wnd[0]").sendVKey 3
            .findById("wnd[0]").sendVKey 3
            .findById("wnd[0]/tbar[0]/okcd").text = SAPTran.getParaVal("�g�����U�N�V�����R�[�h�i���v)")
            .findById("wnd[0]").sendVKey 0
        End If

        ' ����/���㌴�����׉�ʂ̑J�ڊm�F
        Select Case SAPTran.ckScreenTransition("wnd[0]", "*����*", get_Timeout(1))
            Case 0
                ' OK : Continue
            Case 1
                If i < 5 Then
                    outMsg "E05", "����/���㌴�����׉�ʏo�͎��s�I" & vbCrLf & "�t�@�C���� = " & SAPTran.getParaVal("�t�@�C����", i)
                    Exit Sub
                Else
                    outMsg "E06", "����/���㌴�����׉�ʏo�͎��s�I" & vbCrLf & "�t�@�C���� = " & SAPTran.getParaVal("�t�@�C����_2", i - 4)
                    Exit Sub
                End If
            Case 2
                outMsg "E07", "����/���㌴�����׉�ʃ^�C���I�[�o�["
                Exit Sub
            Case Else
                outMsg "S02", "System Error!"
                Exit Sub
        End Select
        
        '�o���A���g��ʂ̑���
        If i < 5 Then
            ' ����/���㌴�����׉�ʂ̑���
            .findById("wnd[0]").sendVKey 17  'shift+F5
        
            If SAPTran.getGridLineNo("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell", "VARIANT", SAPTran.getParaVal("�o���A���g��", i)) = 0 Then
                outMsg "E08", "�o���A���g�̓��e������������܂���B" & vbCrLf & "Variant=" & SAPTran.getParaVal("�o���A���g��", i)
                Exit Sub
            End If
            
            waitSec '1�b�҂�
            
            .findById("wnd[1]").sendVKey 2 'PF2
    
        Else
            If SAPTran.selVariantBat(SAPTran.getParaVal("�o���A���g��_2", i - 4)) = False Then
               Exit Sub
            End If
        End If
        
        ' �����ʂ̑���
        .findById("wnd[0]").sendVKey 8
        
        If i < 5 Then
            If SAPTran.ckDataSelected("") = False Then
                outMsg "L03", "�f�[�^�͑I������܂���ł����BNo=" & i & vbCrLf & "Variant=" & SAPTran.getParaVal("�o���A���g��", i)
                GoTo LoopNext
            End If
        Else
            If SAPTran.ckDataSelected("") = False Then
                outMsg "L04", "�f�[�^�͑I������܂���ł����BNo=" & i & vbCrLf & "Variant=" & SAPTran.getParaVal("�o���A���g��_2", i - 4)
                GoTo LoopNext
            End If
        End If
        ' �����ʂ̑��� ���̂Q,XXL�`���f�[�^�ۑ�
        If i < 5 Then
            .findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
            .findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
            
            If SAPTran.dlXXLBat(SAPTran.getParaVal("SQ01�̕ۊǏꏊ"), SAPTran.getParaVal("�t�@�C����", i)) = False Then
                Exit Sub
            End If
            
        Else
            .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").currentCellRow = -1
            .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
            .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
            
            If SAPTran.dlXXLBat(SAPTran.getParaVal("���v�̕ۊǏꏊ"), SAPTran.getParaVal("�t�@�C����_2", i - 4)) = False Then
                Exit Sub
            End If
        End If
        
        SAPTran.putParaVal "*���㌴�����וۑ��ꏊ���ʁi�t���p�X�j", SAPTran.SaveFilePath, i
        outMsg "L05", "File Created. Path = " & SAPTran.getParaVal("*���㌴�����וۑ��ꏊ���ʁi�t���p�X�j", i)

        
        ' �O�̉�ʂɖ߂�B
        .findById("wnd[0]").sendVKey 3
        
        DoEvents '�����߂��B
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
    
    outMsg "I03", "����������ɏI�����܂����B"
    
    outMsg "L06", "*** Script Cost110 DL_Ended. ***"
End Sub


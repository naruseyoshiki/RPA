Attribute VB_Name = "m_main2"
Option Explicit

Public Sub MainProc2()
    Dim MergeSht As New c_ExcelUser    ' �����t�@�C���̃��[�N�V�[�g��ݒ�
    Dim SAPTran As New c_SAPAccess     '�R���g���[���V�[�g
    Dim DLSht As New c_ExcelUser       'DL�t�@�C��
    Dim Cost_DL As New c_ExcelUser     '���㌴�����׃t�H�[�}�b�g
    Dim Cost_Pivo As New c_ExcelUser   '���㌴�����׃t�H�[�}�b�g�҂�
    Dim Cost_Paste As New c_ExcelUser  '���㌴�����׃t�H�[�}�b�g�\�t�p
    Dim New_Paste As New c_ExcelUser   '�V�K�t�@�C���\�t�p
    Dim i As Integer
    Dim bond As String                  '�����t�@�C���p�X
    Dim item As String                  '���׃t�@�C���p�X
    Dim paste As String                 '�\�t�p�t�@�C���p�X
    
    '��ʕ`��}�~
    Application.ScreenUpdating = False
    
    ' Script Start Log
    outMsg "L07", "*** Script Cost110 EditExcel_Started. ***"
    
    '�p�����[�^�̐ݒ�
    SAPTran.setParaVal "�R���g���[��", 1, 2
    
    'Excel���͗��`�F�b�N
    If SAPTran.ckParaVal("����敪*@2,����敪*@3,����敪*@4,����敪_2@5,�o���A���g��*@2,�o���A���g��*@3,�o���A���g��*@4,�o���A���g��_2@5,�o���A���g��*@2,�o���A���g��*@3,�o���A���g��*@4,�o���A���g��_2@5") = False Then
        Exit Sub
    End If
    
    If outMsg("Q02", "�������J�n���܂���?") = vbNo Then
        outMsg "L08", "�������L�����Z������܂����B"
        Exit Sub
    End If
    
    ' �����t�H�[�}�b�g�t�@�C�����J��
    If MergeSht.openWorkBook(SAPTran.getParaVal("���㌴�����׌����t�H�[�}�b�g�t�@�C���p�X")) = False Then
        outMsg "E09", "���㌴�����׌����t�H�[�}�b�g�t�@�C�����J���܂���B" & vbCrLf & "�t�@�C���p�X = " & SAPTran.getParaVal("���㌴�����׌����t�H�[�}�b�g�t�@�C���p�X")
        Exit Sub
    End If
    
    '�����t�@�C���p�X�̘A��
    bond = SAPTran.getParaVal("�����t�@�C���̕ۊǏꏊ") & SAPTran.getParaVal("�����t�@�C���̃t�@�C����") & "_" & Format(Now, "yyyymmddhhnnss")
    
    ' �����t�@�C���Ƃ��Ė��O��t���ĕۑ�
    If MergeSht.SaveAsWorkBook(bond) = False Then
        outMsg "E10", "�����t�@�C���𖼑O��t���ĕۑ��ł��܂���ł����B" & vbCrLf & "�����t�@�C�� = " & bond
        Exit Sub
    End If

    ' �����t�@�C���̃��[�N�V�[�g��ݒ�
    If MergeSht.setWorkSht("", 1, 1, 16) = False Then
        outMsg "E11", "�����t�@�C���̃V�[�g��ݒ�o���܂���ł����B" & vbCrLf & "�����t�@�C�� = " & bond
        Exit Sub
    End If
            
    For i = 1 To 5
        If i > 1 Then
            ' �����t�@�C����Max No�̃��Z�b�g
            MergeSht.resetMaxNo
        End If
        
        'DL�t�@�C���p�X�L���`�F�b�N
        If SAPTran.getParaVal("*���㌴�����וۑ��ꏊ���ʁi�t���p�X�j", i + 4) = "" Then
            outMsg "L09", SAPTran.getParaVal("�t�@�C����_2", i) & "��DL�t�@�C����������܂���ł����B" & vbCrLf & "DL�t�@�C�� = " & SAPTran.getParaVal("*���㌴�����וۑ��ꏊ���ʁi�t���p�X�j", i + 4)
            GoTo notDL
        End If
        
        ' DL�t�@�C�����J��
        If DLSht.openWorkBook(SAPTran.getParaVal("*���㌴�����וۑ��ꏊ���ʁi�t���p�X�j", i + 4)) = False Then
            outMsg "E12", "DL�t�@�C�����J���܂���ł����B" & vbCrLf & "DL�t�@�C�� = " & SAPTran.getParaVal("*���㌴�����וۑ��ꏊ���ʁi�t���p�X�j", i + 4)
            Exit Sub
        End If
        Debug.Print SAPTran.getParaVal("*���㌴�����וۑ��ꏊ���ʁi�t���p�X�j", i + 4)
        
        '�����t�@�C���ۑ��ꏊ���ʂ�ݒ�
        SAPTran.putParaVal "*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", bond & ".xlsx", 1
        
        ' DL�t�@�C���̃��[�N�V�[�g��ݒ�
        If DLSht.setWorkSht("", 1, 1, 16) = False Then
            outMsg "E13", "DL�t�@�C���̃V�[�g��ݒ�o���܂���ł����B" & vbCrLf & "DL�t�@�C�� = " & SAPTran.getParaVal("*���㌴�����וۑ��ꏊ���ʁi�t���p�X�j", i + 4)
            Exit Sub
        End If

        If DLSht.MaxRowNo = DLSht.StartRowNo Then
            GoTo Down
        End If
        
        ' �����t�@�C���̃t�H�[�}�b�g�̃R�s�[
        MergeSht.copyFormatRow MergeSht.StartRowNo + 1, MergeSht.MaxRowNo + 1, DLSht.MaxRowNo - DLSht.StartRowNo
        
        ' DL�t�@�C�����猋���t�@�C���ւ̃R�s�[
        DLSht.copyData DLSht.StartRowNo + 1, DLSht.getColNo("�W�v�L�["), DLSht.MaxRowNo, DLSht.getColNo("�v�挴��")
        MergeSht.pasteValueData MergeSht.MaxRowNo + 1, MergeSht.getColNo("�W�v�L�[")
        
        '����敪�̐ݒ�
        MergeSht.WS.Cells(MergeSht.MaxRowNo + 1, MergeSht.getColNo("����敪")) = SAPTran.getParaVal("����敪_2", i)
        MergeSht.WS.Cells(MergeSht.MaxRowNo + 1, MergeSht.getColNo("��v����")) = Month(DateAdd("m", -4, Now))
        MergeSht.copyData MergeSht.MaxRowNo + 1, MergeSht.getColNo("����敪"), MergeSht.MaxRowNo + 1, MergeSht.getColNo("��v����")
        MergeSht.pasteValueData MergeSht.MaxRowNo + 1, MergeSht.getColNo("����敪"), MergeSht.MaxRowNo + DLSht.MaxRowNo - DLSht.StartRowNo, MergeSht.getColNo("��v����")
        
Down:
        If DLSht.closeWorkBook() = False Then
            outMsg "E14", "DL�t�@�C�����N���[�Y�o���܂���ł����B" & vbCrLf & "DL�t�@�C�� = " & SAPTran.getParaVal("*���㌴�����וۑ��ꏊ���ʁi�t���p�X�j", i + 4)
            Exit Sub
        End If
        
        outMsg "L10", "--Copy Completed-- " & SAPTran.getParaVal("*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", 1) & " DL�t�@�C���ǉ����� = " & DLSht.MaxRowNo - DLSht.StartRowNo
notDL:
        ' �����t�@�C����Max No�̃��Z�b�g
        MergeSht.resetMaxNo
        
    Next i
        
    '�����t�@�C���̃f�[�^�L���`�F�b�N
    If MergeSht.MaxRowNo = MergeSht.StartRowNo Then
        '�����t�@�C���ۑ��ꏊ���ʂ�ݒ�
        SAPTran.putParaVal "*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", bond & ".xlsx", 1
        outMsg "E15", "�ǉ�����f�[�^������܂���ł����B" & vbCrLf & "�����t�@�C�� = " & SAPTran.getParaVal("*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", 1)
        
        If MergeSht.closeWorkBook() = False Then
            outMsg "E16", "�����t�@�C�����N���[�Y�o���܂���ł����B" & vbCrLf & "�����t�@�C�� = " & SAPTran.getParaVal("*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", 1)
            Set MergeSht = Nothing
            Exit Sub
        End If
    
        Exit Sub
    End If
        
    outMsg "L11", "�����t�@�C���S���� = " & DLSht.RangeRowCnt + MergeSht.MaxRowNo - MergeSht.StartRowNo

    Set DLSht = Nothing

    ' ���㌴�����׃t�H�[�}�b�g���J��
    If Cost_DL.openWorkBook(SAPTran.getParaVal("���㌴�����׃t�H�[�}�b�g�t�@�C���p�X")) = False Then
        outMsg "E17", "���㌴�����׃t�H�[�}�b�g���J���܂���ł����B" & vbCrLf & "�t�H�[�}�b�g = " & SAPTran.getParaVal("���㌴�����׃t�H�[�}�b�g�t�@�C���p�X")
        Exit Sub
    End If
    
    '���㌴�����׃t�@�C�����̘A��
    item = SAPTran.getParaVal("���׃t�@�C���̕ۊǏꏊ") & SAPTran.getParaVal("���׃t�@�C���̃t�@�C����") & "_" & Format(Now, "yyyymmddhhnnss")
                    
    '���㌴�����׃t�@�C���Ƃ��Ė��O��t���ĕۑ�
    If Cost_DL.SaveAsWorkBook(item) = False Then
        outMsg "E18", "���㌴�����׃t�@�C����ۑ��ł��܂���ł����B" & vbCrLf & "���㌴�����׃t�@�C�� = " & item
        Exit Sub
    End If
    
    '���㌴�����׃t�@�C���ۑ��ꏊ���ʂ�ݒ�
    SAPTran.putParaVal "*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", item & ".xlsx", 2
    
    '���ׂ̃��[�N�V�[�g��ݒ�
    If Cost_DL.setWorkSht("�_�E�����[�h�f�[�^", 1, 1, 16) = False Then
        outMsg "E19", "�u�_�E�����[�h�f�[�^�v�V�[�g��ݒ�ł��܂���ł����B" & vbCrLf & "���׃t�@�C�� = " & SAPTran.getParaVal("*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", 2)
        Exit Sub
    End If
    
    ' �V�[�g�_�E�����[�h�f�[�^�ւ̃t�H�[�}�b�g�̃R�s�[
    Cost_DL.copyFormatRow Cost_DL.StartRowNo + 1, Cost_DL.StartRowNo + 1, MergeSht.MaxRowNo - MergeSht.StartRowNo
    
    ' �t�H�[�~�����[�̃R�s�[
    Cost_DL.copyFormulaRow2 Cost_DL.getColNo("�n��"), Cost_DL.getColNo("�����Ǒ�"), Cost_DL.StartRowNo + 1, Cost_DL.StartRowNo + 2, MergeSht.MaxRowNo - MergeSht.StartRowNo - 1
    
    ' DL�t�@�C�����猋���t�@�C���ւ̃R�s�[
    MergeSht.copyData MergeSht.StartRowNo + 1, MergeSht.getColNo("�W�v�L�["), MergeSht.MaxRowNo, MergeSht.getColNo("����敪")
    Cost_DL.pasteValueData Cost_DL.StartRowNo + 1, Cost_DL.getColNo("�W�v�L�[")

    '�����t�@�C����ۑ��A����
    If MergeSht.saveWorkBook() = False Then
        outMsg "E20", "�����t�@�C����ۑ��ł��܂���ł����B" & vbCrLf & "�t�@�C�� = " & SAPTran.getParaVal("*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", 1)
        Exit Sub
    End If
        
    If MergeSht.closeWorkBook() = False Then
        outMsg "E21", "�����t�@�C�����N���[�Y�o���܂���ł����B" & vbCrLf & "�t�@�C�� = " & SAPTran.getParaVal("*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", 1)
        Exit Sub
    End If
        
    Set MergeSht = Nothing

    ' ���㌴�����׃t�@�C��-�_�E�����[�h��Max No���Ď擾
    Cost_DL.resetMaxNo
    
    Cost_DL.mkTable "TBL_�_�E�����[�h�f�[�^", Cost_DL.StartRowNo, Cost_DL.getColNo("�W�v�L�["), Cost_DL.MaxRowNo, Cost_DL.getColNo("�����Ǒ�")
    
    '�s�{�t�@�C���̐ݒ�
    Cost_Pivo.setWorkBook Cost_DL.WB
    
    If Cost_Pivo.setWorkSht("�҂�", 4, 1) = False Then
        outMsg "E22", "�҂ڂ̃��[�N�V�[�g��ݒ�o���܂���ł����B" & vbCrLf & "�t�@�C�� = " & SAPTran.getParaVal("*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", 2)
        Exit Sub
    End If
    
    Cost_Pivo.updPivotDS "TBL_�҂�", "TBL_�_�E�����[�h�f�[�^"
    
    '�҂ڂ�Max No���Ď擾
    Cost_Pivo.resetMaxNo
    
    '�\�t�p�̐ݒ�
    Cost_Paste.setWorkBook Cost_Pivo.WB
    
    If Cost_Paste.setWorkSht("�\�t�p", 1, 1) = False Then
        outMsg "E23", "�\�t�p���[�N�V�[�g��ݒ�o���܂���ł����B" & vbCrLf & "�t�@�C�� = " & SAPTran.getParaVal("*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", 2)
        Exit Sub
    End If
    
    ' �҂ڃV�[�g����\�t�p�V�[�g�ւ̃R�s�[
    Cost_Paste.copyData Cost_Paste.StartRowNo + 1, Cost_Paste.getColNo("��v����"), Cost_Paste.StartRowNo + 1, Cost_Paste.getColNo("�󒍐�敪")
    Cost_Paste.pasteData Cost_Paste.StartRowNo + 2, Cost_Paste.getColNo("��v����"), Cost_Pivo.MaxRowNo - Cost_Pivo.StartRowNo + Cost_Paste.StartRowNo, Cost_Paste.getColNo("��v����")
    
    
    '�\�t�p�u�b�N�V�K�쐬
    If New_Paste.addWorkBook() = False Then
        outMsg "E24", "���[�N�u�b�N��V�K�쐬�ł��܂���ł����B"
        Exit Sub
    End If
    
    '�V�K�\�t�p�V�[�g��ݒ�
    If New_Paste.setWorkSht("", 1, 1) = False Then
        outMsg "E25", "�V�K�V�[�g�ɐݒ�o���܂���ł����B"
        Exit Sub
    End If
    
    '�\�t�p�V�[�g���R�s�[�A�V�K�u�b�N�ɓ\��t��
    Cost_Paste.copyAllShtData
    New_Paste.pasteAllShtData
    
    '�V�K�u�b�N���C�ӂ̃V�[�g���ύX
    If New_Paste.chgShtNm(SAPTran.getParaVal("�\�t�p�t�@�C���̃V�[�g��")) = False Then
        outMsg "E26", "�V�K�u�b�N�̃V�[�g����ύX�o���܂���ł����B"
        Exit Sub
    End If
    
    '�V�K�\�t�p�t�@�C�����̘A��
    paste = SAPTran.getParaVal("�\�t�p�t�@�C���̕ۊǏꏊ") & SAPTran.getParaVal("�\�t�p�t�@�C���̃t�@�C����") & "_" & Format(Now, "yyyymmddhhnnss")
    
    '�V�K�u�b�N���w��p�X�Ŗ��O��t���ĕۑ�
    If New_Paste.SaveAsWorkBook(paste) = False Then
        outMsg "E27", "�V�K�u�b�N�𖼑O��t���ĕۑ��o���܂���ł����B" & vbCrLf & "�V�K�t�@�C���p�X = " & paste
        Exit Sub
    End If
    
    '�V�K�u�b�N�ۑ��ꏊ���ʂ�ݒ�
    SAPTran.putParaVal "*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", paste & ".xlsx", 3
    
    '���ׂ�ۑ�,����
    If Cost_DL.saveWorkBook() = False Then
        outMsg "E28", "���ׂ�ۑ��ł��܂���ł����B" & vbCrLf & "���׃t�@�C�� = " & SAPTran.getParaVal("*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", 2)
        Exit Sub
    End If
        
    If Cost_DL.closeWorkBook() = False Then
        outMsg "E29", "���ׂ��N���[�Y�o���܂���ł����B" & vbCrLf & "���׃t�@�C�� = " & SAPTran.getParaVal("*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", 2)
        Exit Sub
    End If
    
    If New_Paste.closeWorkBook() = False Then
        outMsg "E30", "�V�K�u�b�N���N���[�Y�o���܂���ł����B" & vbCrLf & "�V�K�t�@�C�� = " & SAPTran.getParaVal("*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", 3)
        Exit Sub
    End If
    
    outMsg "L12", "�y�҂ځz�S���� = " & Cost_Pivo.MaxRowNo - Cost_Pivo.StartRowNo
    
    outMsg "L13", "�y" & SAPTran.getParaVal("�����t�@�C���̃t�@�C����") & "�z" & "File Created. Path = " & SAPTran.getParaVal("*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", 1)
    
    outMsg "L14", "�y" & SAPTran.getParaVal("���׃t�@�C���̃t�@�C����") & "�z" & "File Created And Copy Completed. Path = " & SAPTran.getParaVal("*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", 2)
    
    outMsg "L15", "�y" & SAPTran.getParaVal("�\�t�p�t�@�C���̃t�@�C����") & "�z" & "File Created. Path = " & SAPTran.getParaVal("*Excel�ҏW�t�@�C���ۑ��ꏊ���ʁi�t���p�X�j", 3)
    
    ' �e�t�@�C���̏�����
    Set Cost_Pivo = Nothing
    Set Cost_DL = Nothing
    Set MergeSht = Nothing
    Set DLSht = Nothing
    Set New_Paste = Nothing
    
    outMsg "I04", "����������ɏI�����܂����B"
    
    outMsg "L16", "*** Script Cost110 EditExcel_Ended. ***"

    Application.ScreenUpdating = True
    
End Sub

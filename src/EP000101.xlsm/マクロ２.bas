Attribute VB_Name = "�}�N���Q"
Dim henkan, henkan_2, henkan_3, henkan_4, kioku_3, lank_9      '�ǉ� 99/ 1/21 K.Y
'******************************************
'
' �Ǘ��K�i�ݒ� �}�N��
'
' 1998/07/22 �C�� ���`���؋K�i���羯Ă��鍀�ڐ��̕ύX�i�o�O�Ή�)
'******************************************
Sub �Ǘ��K�i()
Attribute �Ǘ��K�i.VB_ProcData.VB_Invoke_Func = " \n14"

    Message = Array("���i�}�ɁA���L ���L���L������Ă��܂����H", "� �V�F�[�r���O�a � ���I�X�h�b�N�Ɋ|�����Ă��܂����H")
    
    w_��� = ""
    
    '���`�E���ؐݒ�     ���`�A���؋K�i�� �� JIS4���̎�   aaaa
    j = 0
    If Sheets("�\�t�����").Range("M3") = "��" And Mid(Sheets("���ͼ��").Range("D10"), 4, 1) = "4" Then
        
        '''�I�X�h�b�Nchk
        w_ans = MsgBox(Message(1) & Chr(13), vbYesNoCancel, "���`�E���؋K�i�ݒ�")
        Select Case w_ans
        Case Is = vbYes
              With Sheets("���ͼ��").DrawingObjects("���L")
                .Visible = True
              End With
        
              w_ans = MsgBox(Message(0) & Chr(13), vbYesNoCancel, "���`�E���؋K�i�ݒ�")
        
              With Sheets("���ͼ��").DrawingObjects("���L")
                .Visible = False
              End With
              Select Case w_ans
                     Case Is = vbYes
                        Sheets("���ͼ��").Range("E10") = "���}�n�U��"
                        For i = 0 To 9
                        If Workbooks(XLS_KTRTN).Sheets("���`���؋K�i").Cells(i + 3, j + 4) = "" Then
                            Sheets("���ͼ��").Cells(i + 7, j + 11) = "�|"
                        Else
                            Sheets("���ͼ��").Cells(i + 7, j + 11) = Workbooks(XLS_KTRTN).Sheets("���`���؋K�i").Cells(i + 3, j + 4)
                        End If
                        Next i
                    Case Is = vbNo
                        Sheets("���ͼ��").Range("E10") = "���}�n�S��"
                        For i = 0 To 9
                        If Workbooks(XLS_KTRTN).Sheets("���`���؋K�i").Cells(i + 3, j + 2) = "" Then
                            Sheets("���ͼ��").Cells(i + 7, j + 11) = "�|"
                        Else
                            Sheets("���ͼ��").Cells(i + 7, j + 11) = Workbooks(XLS_KTRTN).Sheets("���`���؋K�i").Cells(i + 3, j + 2)
                        End If
                        Next i
                        a = MsgBox("� �u�`��� � ���o���ĉ�����", vbExclamation)
              End Select
        
        
        Case Is = vbNo
            Sheets("���ͼ��").Range("E10") = "���}�n�S��"
'            For i = 0 To 8  1998/07/22 �C��
            For i = 0 To 9
                If Workbooks(XLS_KTRTN).Sheets("���`���؋K�i").Cells(i + 3, j + 2) = "" Then
                    Sheets("���ͼ��").Cells(i + 7, j + 11) = "�|"
                Else
                    Sheets("���ͼ��").Cells(i + 7, j + 11) = Workbooks(XLS_KTRTN).Sheets("���`���؋K�i").Cells(i + 3, j + 2)
                End If
            Next i
            Exit Sub
        End Select
    End If
End Sub
'******************************************
'
' �����ݒ� �}�N��
'
'******************************************
Sub �����\��()
Attribute �����\��.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim w_kousa1, w_kousa2
    
    w_kousa1 = Worksheets("���ͼ��").Range("D8")
    w_kousa2 = Worksheets("���ͼ��").Range("D9")
    
    If (Not IsNumeric(w_kousa1)) Or (Not IsNumeric(w_kousa2)) Then
    
        If w_kousa1 = "*****" And w_kousa2 = "*****" Then
        
            Sheets("�H��}").TextBoxes("kousa0").Visible = False
            Sheets("�H��}").TextBoxes("kousa1").Visible = False
            Sheets("�H��}").TextBoxes("kousa2").Visible = False
        Else
            a = MsgBox("�����ɂ́A���l��ݒ肵�Ă��������B", vbExclamation)
        End If
        
        Exit Sub
        
    End If
    
    If Abs(w_kousa1) = Abs(w_kousa2) Then                       '�����̐�Βl��������
        With Worksheets("�H��}").TextBoxes("kousa0")     '"�}" �\��
            .Caption = "�}" + CStr(Abs(w_kousa1))
            .Font.Size = 16
            .Visible = True
        End With
        
        With Worksheets("�H��}").TextBoxes("kousa1")     '"���" ��\��
            .Visible = False
        End With
        
        With Worksheets("�H��}").TextBoxes("kousa2")    '"����" ��\��
            .Visible = False
        End With
    Else
        With Worksheets("�H��}").TextBoxes("kousa0")    '"�}" ��\��
            .Visible = False
        End With
        '**********************************
        RET = Application.Run(XLS_KTRTN + "!����_�ҏW", w_kousa1, w_kousa2)
        If RET = "err" Then
           Exit Sub
        Else
           w_kousa1 = Mid(RET, 1, Application.Search("#", RET, 1) - 1)
           w_kousa2 = Mid(RET, Application.Search("#", RET, 1) + 1, 6)
        End If
        '**********************************
                                        
        '"���" �\��
        With Worksheets("�H��}").TextBoxes("kousa1")
            .Font.Name = "�l�r �S�V�b�N"
            .Caption = w_kousa1
            .Font.Size = 11
            .Visible = True
        End With
        
        '"����"�\��
        With Worksheets("�H��}").TextBoxes("kousa2")
            .Font.Name = "�l�r �S�V�b�N"
            .Caption = w_kousa2
            .Font.Size = 11
            .Visible = True
        End With
    End If
    
End Sub
'******************************************
'
' �T�C�Y�i�w�� �}�N��
'
'******************************************
Sub �T�C�Y�w��()
Attribute �T�C�Y�w��.VB_ProcData.VB_Invoke_Func = " \n14"

    Sheets("���ͼ��").Unprotect
    Sheets("���ͼ��").Range("C23:C30") = "�|"
    w_ans = MsgBox("���̃��[�N �́A� �T�C�Y�i � �ł����H" & Chr(13), vbYesNoCancel, "���ޕi �ݒ�")
    
    Select Case w_ans
    Case Is = vbYes
        
        Sheets("�H��}").TextBoxes("size2").Visible = True
'        Sheets("�H��}").TextBoxes("size4").Visible = True
        
        Sheets("���ͼ��").TextBoxes("size5").Visible = True
        
        If Worksheets("�\�t�����").Range("H3") = "M" Then
           �T�C�Y�I��
        Else
          If Worksheets("�\�t�����").Range("H3") = "O" Then
msg_4:      Message_4 = "�T�C�Y�\�̃p�^�[���͂ǂ����܂����H" & Chr(13) & "���̔ԍ�����I�����ĉ������B" & Chr(13) & Chr(13) & _
                        "�P�D����      �Q�D�~��"
            Title = "�T�C�Y�\�p�^�[���I��"
            w_Default = "1"
            MyValue_4 = InputBox(Message_4, Title, w_Default)
            henkan_4 = CStr(MyValue_4)
            Select Case henkan_4
              Case "1"
                �\������_1
              Case "2"
                �\������_1
              Case Else
                If henkan_4 <> "" Then
                  c_4 = MsgBox("�͈͊O�̐��l�ł��B������x���͂��ĉ������B", vbQuestion, "����")
                  GoTo msg_4
                Else
                  Exit Sub
                End If
            End Select
          End If
        End If
        Sheets("�H��}").Range("AW16").Value = "1/50"    '2000/1/25 ADD
        Sheets("�H��}").Range("AW18").Value = "1/50"    '2000/1/25 ADD
    Case vbNo
        
        �T�C�Y�i�łȂ��Ƃ�_2              '�T�C�Y�i���̐ݒ����������
        Sheets("�H��}").Range("AW16").Value = "1/2h"    '2000/1/25 ADD
        Sheets("�H��}").Range("AW18").Value = "1/2h"    '2000/1/25 ADD
        
    End Select
    'test
End Sub
        
'******************************************
'
'  �}�^�M�����I�����̃}�N��
'
'******************************************
Sub �T�C�Y�I��()
Attribute �T�C�Y�I��.VB_ProcData.VB_Invoke_Func = " \n14"

     Sheets("���ͼ��").DrawingObjects("size6").Visible = True
        Sheets("���ͼ��").DrawingObjects("�B���U").Visible = True
'�I�v�V�����{�^���̔�\��          Excel2007�Ή�
    Sheets("���ͼ��").Shapes("�{�^�� 312").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��1").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��2").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��3").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��4").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��5").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��6").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��7").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��8").Visible = False
       
msg_1:  Message = "�ǂ̂悤�ȕ\�����@�ɂ��܂����H" & Chr(13) & "���̔ԍ�����I�����ĉ������B" & Chr(13) & Chr(13) & _
                  "�P�D�}�^�M�����̂�   �Q�D����"
        Title = "�T�C�Y�i�I��"
        w_Default = "1"
        MyValue = InputBox(Message, Title, w_Default)
        henkan = CStr(MyValue)
        Select Case henkan
          Case "1"
msg_2:      Message_2 = "�T�C�Y�\�̃p�^�[���͂ǂ����܂����H" & Chr(13) & "���̔ԍ�����I�����ĉ������B" & Chr(13) & Chr(13) & _
                        "�P�D����      �Q�D�~��"
            Title = "�T�C�Y�\�p�^�[���I��"
            w_Default = "1"
            MyValue_2 = InputBox(Message_2, Title, w_Default)
            henkan_2 = CStr(MyValue_2)
            Select Case henkan_2
              Case "1"
                Sheets("���ͼ��").Range("C7").Value = "�}�^�M����"
                Sheets("���ͼ��").Range("C7").Font.Size = 11
                �\������_1
              Case "2"
                Sheets("���ͼ��").Range("C7").Value = "�}�^�M����"
                Sheets("���ͼ��").Range("C7").Font.Size = 11
                �\������_1
              Case Else
                If henkan_2 <> "" Then
                  c_2 = MsgBox("�͈͊O�̐��l�ł��B������x���͂��ĉ������B", vbQuestion, "����")
                  GoTo msg_2
                Else
                  Exit Sub
                End If
            End Select
          Case "2"
msg_3:      Message_3 = "�T�C�Y�\�̃p�^�[���͂ǂ����܂����H" & Chr(13) & "���̔ԍ�����I�����ĉ������B" & Chr(13) & Chr(13) & _
                        "�P�D����      �Q�D�~��"
            Title = "�T�C�Y�\�p�^�[���I��"
            w_Default = "1"
            MyValue_3 = InputBox(Message_3, Title, w_Default)
            henkan_3 = CStr(MyValue_3)
            Select Case henkan_3
              Case "1"
                kioku_3 = henkan_3
                Sheets("�����v�Z").Select
                Range("G3").Select
                a_res = MsgBox("���F�̃Z���ɐ��l����͂��ĉ������B���͌�A[���̓V�[�g�֖߂�]�{�^���������ĉ������B", vbInformation, "�m�F")
              Case "2"
                kioku_3 = henkan_3
                Sheets("�����v�Z�Q").Select
                Range("G3").Select
                a_res = MsgBox("���F�̃Z���ɐ��l����͂��ĉ������B���͌�A[���̓V�[�g�֖߂�]�{�^���������ĉ������B", vbInformation, "�m�F")
              Case Else
                If henkan_3 <> "" Then
                  c_3 = MsgBox("�͈͊O�̐��l�ł��B������x���͂��ĉ������B", vbQuestion, "����")
                  GoTo msg_3
                Else
                  Exit Sub
                End If
            End Select
          Case Else
            If henkan <> "" Then
               c = MsgBox("�͈͊O�̐��l�ł��B������x���͂��ĉ������B", vbQuestion, "����")
               GoTo msg_1
            Else
               Exit Sub
            End If
        
        End Select
        
End Sub

'******************************************
'
'  �T�C�Y�i�łȂ����Ɠ��͒l�N���A�|���̃}�N��
'
'******************************************
Sub �T�C�Y�i�łȂ��Ƃ�_1()
Attribute �T�C�Y�i�łȂ��Ƃ�_1.VB_ProcData.VB_Invoke_Func = " \n14"
    
    With Sheets("�H��}").TextBoxes("text2")
        .Visible = True
    End With
    With Sheets("�H��}").TextBoxes("text3")
        .Visible = True
    End With
    With Sheets("�H��}").TextBoxes("size1_�O")
        .Visible = False
    End With
    
    Sheets("�H��}").DrawingObjects("�B���k").Visible = True
    Sheets("�H��}").DrawingObjects("�B���q").Visible = True
    
    Sheets("�H��}").TextBoxes("size2").Visible = False
    Sheets("�H��}").TextBoxes("size4").Visible = False
    Sheets("�H��}").TextBoxes("size4�̍�").Visible = False
        
    Sheets("���ͼ��").TextBoxes("size5").Visible = False
    Sheets("���ͼ��").DrawingObjects("�B���U").Visible = True
    With Sheets("���ͼ��").DrawingObjects("size6")
        .Visible = True
        .BringToFront
    End With
'�I�v�V�����{�^���̔�\��          Excel2007�Ή�
    Sheets("���ͼ��").Shapes("�{�^�� 312").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��1").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��2").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��3").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��4").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��5").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��6").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��7").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��8").Visible = False
    
    If Sheets("�\�t�����").Range("H3") = "O" Then               '�I�[�o�[�s���̏ꍇ         '�ǉ� 99/ 1/18
       Sheets("���ͼ��").Range("C7").Value = "�I�[�o�[�s���a"
       Sheets("���ͼ��").Range("C7").Font.Size = 10
       Sheets("���ͼ��").Range("C15").Value = "�s���a"
       Sheets("���ͼ��").Range("C15").Font.Size = 11
       Sheets("���ͼ��").Range("G16").Value = "���ް��݌a�����·"
       Sheets("���ͼ��").Range("G16").Font.Size = 10
       Sheets("���ͼ��").Range("H16").Value = 0.05
       Sheets("���ͼ��").Range("H16").Font.Size = 11
'       Sheets("�H��}").Range("AG44").Value = "�����}�C�N��"                                '�ǉ� 99/ 1/18
'       Sheets("�H��}").Range("AG45").Value = "���ް���ϲ��"                               '�ǉ� 99/ 1/18
    ElseIf Sheets("�\�t�����").Range("H3") = "M" Then           '�}�^�M�̏ꍇ               '�ǉ� 99/ 1/18
       Sheets("���ͼ��").Range("C7").Value = "�}�^�M����"
       Sheets("���ͼ��").Range("C7").Font.Size = 11
       Sheets("���ͼ��").Range("C15").Value = "�}�^�M����"
       Sheets("���ͼ��").Range("C15").Font.Size = 11
       Sheets("���ͼ��").Range("G16").Value = "�s�b�`�덷"
       Sheets("���ͼ��").Range("G16").Font.Size = 11
       Sheets("���ͼ��").Range("H16").Value = 0.02
       Sheets("���ͼ��").Range("H16").Font.Size = 11
'       Sheets("�H��}").Range("AG44").Value = "�����}�C�N��"                                '�ǉ� 99/ 1/18
'       Sheets("�H��}").Range("AG45").Value = "�����}�C�N��"                                '�ǉ� 99/ 1/18
    End If                                                                                 '�ǉ� 99/ 1/18

    With Sheets("�H��}").TextBoxes("text1")
        .Visible = True
        .Caption = Sheets("���ͼ��").Range("C7")
    End With
    With Sheets("�H��}").TextBoxes("size2")           '�ǉ� 99/ 1/18
        .Visible = False                               '�ǉ� 99/ 1/18
        .Caption = "�������ޕ\�ɂ��"                   '�ǉ� 99/ 1/18
        .Font.Size = 6                                 '�ǉ� 99/ 1/18
    End With                                           '�ǉ� 99/ 1/18
    With Sheets("�H��}").TextBoxes("size4�̍�")        '�ǉ� 99/ 1/18
        .Visible = False                               '�ǉ� 99/ 1/18
        .Caption = "�������ޕ\�ɂ��"                   '�ǉ� 99/ 1/18
        .Font.Size = 12                                '�ǉ� 99/ 1/18
    End With                                           '�ǉ� 99/ 1/18
    With Sheets("�H��}").TextBoxes("size4")           '�ǉ� 99/ 1/18
        .Visible = False                               '�ǉ� 99/ 1/18
        .Caption = "��"                                '�ǉ� 99/ 1/18
        .Font.Size = 16                                '�ǉ� 99/ 1/18
    End With                                           '�ǉ� 99/ 1/18
        
End Sub

'******************************************
'
'  �T�C�Y�i�łȂ����Ɠ��͒l�N���A�|���̃}�N��2
'
'******************************************
Sub �T�C�Y�i�łȂ��Ƃ�_2()
Attribute �T�C�Y�i�łȂ��Ƃ�_2.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Application.Run (XLS_KTRTN + "!�H��}�A�����b�N")
    With Sheets("�H��}").TextBoxes("text2")
        .Visible = True
    End With
    With Sheets("�H��}").TextBoxes("size1_�O")
        .Visible = False
    End With
    
    Sheets("�H��}").DrawingObjects("�B���k").Visible = True
    Sheets("�H��}").DrawingObjects("�B���q").Visible = True
    
    Sheets("�H��}").TextBoxes("size2").Visible = False
    Sheets("�H��}").TextBoxes("size4").Visible = False
    Sheets("�H��}").TextBoxes("size4�̍�").Visible = False
        
    Sheets("���ͼ��").TextBoxes("size5").Visible = False
    Sheets("���ͼ��").DrawingObjects("�B���U").Visible = True
    With Sheets("���ͼ��").DrawingObjects("size6")
        .Visible = True
        .BringToFront
    End With
'�I�v�V�����{�^���̔�\��          Excel2007�Ή�
    Sheets("���ͼ��").Shapes("�{�^�� 312").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��1").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��2").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��3").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��4").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��5").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��6").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��7").Visible = False
    Sheets("���ͼ��").Shapes("�{�^��8").Visible = False
    
    If Sheets("�\�t�����").Range("H3") = "O" Then               '�I�[�o�[�s���̏ꍇ         '�ǉ� 99/ 1/18
       Sheets("���ͼ��").Range("C7").Value = "�I�[�o�[�s���a"
       Sheets("���ͼ��").Range("C7").Font.Size = 10
       Sheets("���ͼ��").Range("C15").Value = "�s���a"
       Sheets("���ͼ��").Range("C15").Font.Size = 11
       Sheets("���ͼ��").Range("G16").Value = "���ް��݌a�����·"
       Sheets("���ͼ��").Range("G16").Font.Size = 10
       Sheets("���ͼ��").Range("H16").Value = 0.05
       Sheets("���ͼ��").Range("H16").Font.Size = 11
'       Sheets("�H��}").Range("AG44").Value = "�����}�C�N��"                                '�ǉ� 99/ 1/18
'       Sheets("�H��}").Range("AG45").Value = "���ް���ϲ��"                               '�ǉ� 99/ 1/18
       With Sheets("�H��}").TextBoxes("textp1")                                           '�ǉ� 99/ 1/14
           .Visible = True                                                                 '�ǉ� 99/ 1/14
           .Font.Size = 20                                                                 '�ǉ� 99/ 1/14
           .Caption = "���ް��݌a�����·  " & Sheets("���ͼ��").Range("H16") & "  �ȉ�"     '�ǉ� 99/ 1/14
       End With                                                                            '�ǉ� 99/ 1/14
       With Sheets("�H��}").TextBoxes("text3")                                           '�ǉ� 99/ 1/14
           .Visible = True                                                                '�ǉ� 99/ 1/14
           .Font.Size = 18                                                                '�ǉ� 99/ 1/14
           .Caption = "(�s���a  " & Sheets("���ͼ��").Range("D15") & ")"                   '�ǉ� 99/ 1/14
       End With                                                                           '�ǉ� 99/ 1/14

    ElseIf Sheets("�\�t�����").Range("H3") = "M" Then           '�}�^�M�̏ꍇ               '�ǉ� 99/ 1/18
       Sheets("���ͼ��").Range("C7").Value = "�}�^�M����"
       Sheets("���ͼ��").Range("C7").Font.Size = 11
       Sheets("���ͼ��").Range("C15").Value = "�}�^�M����"
       Sheets("���ͼ��").Range("C15").Font.Size = 11
       Sheets("���ͼ��").Range("G16").Value = "�s�b�`�덷"
       Sheets("���ͼ��").Range("G16").Font.Size = 11
       Sheets("���ͼ��").Range("H16").Value = 0.02
       Sheets("���ͼ��").Range("H16").Font.Size = 11
'       Sheets("�H��}").Range("AG44").Value = "�����}�C�N��"                                '�ǉ� 99/ 1/18
'       Sheets("�H��}").Range("AG45").Value = "�����}�C�N��"                                '�ǉ� 99/ 1/18
       With Sheets("�H��}").TextBoxes("textp1")                                           '�ǉ� 99/ 1/14
           .Visible = True                                                                 '�ǉ� 99/ 1/14
           .Font.Size = 20                                                                 '�ǉ� 99/ 1/14
           .Caption = "�s�b�`�덷  " & Sheets("���ͼ��").Range("H16") & "  �ȉ�"             '�ǉ� 99/ 1/14
       End With                                                                             '�ǉ� 99/ 1/14
       With Sheets("�H��}").TextBoxes("text3")                                             '�ǉ� 99/ 1/14
           .Visible = True                                                                  '�ǉ� 99/ 1/14
           .Font.Size = 18                                                                  '�ǉ� 99/ 1/14
           .Caption = "(�}�^�M����  " & Sheets("���ͼ��").Range("D15") & "��)"               '�ǉ� 99/ 1/14
       End With                                                                             '�ǉ� 99/ 1/14
       With Sheets("�H��}").TextBoxes("micro_text1")                                       '�ǉ� 99/ 1/18
           .Caption = "����ϲ��"                                                            '�ǉ� 99/ 1/18
           .Font.Size = 9                                                                   '�ǉ� 99/ 1/18
           .Visible = True                                                                  '�ǉ� 99/ 1/18
       End With                                                                             '�ǉ� 99/ 1/18
    End If                                                                                  '�ǉ� 99/ 1/18

    With Sheets("�H��}").TextBoxes("text1")
        .Visible = True
        .Caption = Sheets("���ͼ��").Range("C7")
    End With
    With Sheets("�H��}").TextBoxes("size2")           '�ǉ� 99/ 1/18
        .Visible = False                               '�ǉ� 99/ 1/18
        .Caption = "�������ޕ\�ɂ��"                   '�ǉ� 99/ 1/18
        .Font.Size = 6                                 '�ǉ� 99/ 1/18
    End With                                           '�ǉ� 99/ 1/18
    With Sheets("�H��}").TextBoxes("size4�̍�")        '�ǉ� 99/ 1/18
        .Visible = False                               '�ǉ� 99/ 1/18
        .Caption = "�������ޕ\�ɂ��"                   '�ǉ� 99/ 1/18
        .Font.Size = 12                                '�ǉ� 99/ 1/18
    End With                                           '�ǉ� 99/ 1/18
    With Sheets("�H��}").TextBoxes("size4")           '�ǉ� 99/ 1/18
        .Visible = False                               '�ǉ� 99/ 1/18
        .Caption = "��"                                '�ǉ� 99/ 1/18
        .Font.Size = 16                                '�ǉ� 99/ 1/18
    End With                                           '�ǉ� 99/ 1/18
        
    Application.Run (XLS_KTRTN + "!�H��}���b�N")
End Sub
'****************************************
'
'  �����v�Z����̓V�|�g�֖߂鎞�̃}�N��
'
'****************************************
Sub ���̓V�[�g�֖߂�()
Attribute ���̓V�[�g�֖߂�.VB_ProcData.VB_Invoke_Func = " \n14"
    �\������_2
End Sub

'****************************************
'
'  ���l��\�ɏ����}�N�� ���̂P
'
'****************************************
Sub �\������_1()
Attribute �\������_1.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("���ͼ��").Unprotect
    Application.Run (XLS_KTRTN + "!�H��}�A�����b�N")
    With Sheets("�H��}").TextBoxes("size1_�O")
        .Visible = False
    End With
    If Worksheets("���ͼ��").Range("C7") = "�}�^�M����" Then
      If henkan_2 = "1" Then                                    '�}�^�M����  �����̎�
        For i = 1 To 8
          Sheets("���ͼ��").Range("D" & Trim(i + 22)) = Sheets("�����v�Z").Cells(i + 3, 33)
        Next i
        For i = 1 To 8
          Sheets("�H��}").TextBoxes("����" & Trim(i)).Caption = Sheets("���ͼ��").Range("D" & Trim(i + 22))
        Next i
      Else
        If henkan_2 = "2" Then                                  '�}�^�M����  �~���̎�
          For i = 4 To 11
            Sheets("���ͼ��").Range("D" & Trim(i + 19)) = Sheets("�����v�Z�Q").Cells(i, 33)
          Next i
          For i = 1 To 8
            Sheets("�H��}").TextBoxes("����" & Trim(i)).Caption = Sheets("���ͼ��").Range("D" & Trim(i + 22))
          Next i
        End If
      End If
       
      Sheets("�H��}").TextBoxes("SUUTI_1").Caption = "( " & Sheets("���ͼ��").Range("D15") & "��" & " )"
       
'++++++++ �R�}���h�{�^���̕\���A��\����؂�ւ��� 2012/01/01 +++++++++
      For i = 1 To 8
        With Sheets("���ͼ��").OptionButtons("�{�^��" & Trim(Str(i)))
            .Value = xlOff
            .Visible = True
        End With
      Next i
       
      Sheets("���ͼ��").Range("C15").Value = "�}�^�M����"
      Sheets("���ͼ��").Range("C15").Font.Size = 11
      Sheets("���ͼ��").Range("G16").Value = "�s�b�`�덷"
      Sheets("���ͼ��").Range("G16").Font.Size = 11
      Sheets("���ͼ��").Range("H16").Value = 0.02
      Sheets("���ͼ��").Range("H16").Font.Size = 11

      With Sheets("�H��}").TextBoxes("textp1")                                    '�ǉ� 99/ 1/14
          .Visible = True                                                          '�ǉ� 99/ 1/14
          .Font.Size = 20                                                          '�ǉ� 99/ 1/14
          .Caption = "�s�b�`�덷  " & Sheets("���ͼ��").Range("H16") & "  �ȉ�"     '�ǉ� 99/ 1/14
      End With                                                                     '�ǉ� 99/ 1/14
      With Sheets("�H��}").TextBoxes("text1")                                      '�ǉ� 99/ 1/14
          .Visible = True                                                           '�ǉ� 99/ 1/14
          .Caption = "�}�^�M����"                                                    '�ǉ� 99/ 1/14
          .Font.Size = 20                                                           '�ǉ� 99/ 1/14
      End With                                                                      '�ǉ� 99/ 1/14
      With Sheets("�H��}").TextBoxes("text2")                                      '�ǉ� 99/ 1/14
          .Visible = True                                                           '�ǉ� 99/ 1/14
      End With                                                                      '�ǉ� 99/ 1/14
      With Sheets("�H��}").TextBoxes("text3")                                      '�ǉ� 99/ 1/14
          .Visible = True                                                           '�ǉ� 99/ 1/14
          .Font.Size = 18                                                           '�ǉ� 99/ 1/14
          .Caption = "(�}�^�M����  " & Sheets("���ͼ��").Range("D15") & "��)"        '�ǉ� 99/ 1/14
      End With                                                                      '�ǉ� 99/ 1/14

'      Sheets("�H��}").Range("AG44").Value = "�����}�C�N��"                          '�ǉ� 99/ 1/18
'      Sheets("�H��}").Range("AG45").Value = "�����}�C�N��"                          '�ǉ� 99/ 1/18
      Sheets("�H��}").Range("AG45").Font.Size = 11
      With Sheets("�H��}").TextBoxes("micro_text11")
          .Visible = False
      End With
      With Sheets("�H��}").TextBoxes("micro_text12")
          .Visible = False
      End With
      With Sheets("�H��}").TextBoxes("micro_text1")                                '�ǉ� 99/ 1/18
          .Caption = "����ϲ��"                                                     '�ǉ� 99/ 1/18
          .Font.Size = 9                                                            '�ǉ� 99/ 1/18
          .Visible = True                                                           '�ǉ� 99/ 1/18
      End With                                                                      '�ǉ� 99/ 1/18
      Sheets("���ͼ��").DrawingObjects("size6").Visible = False
    '�I�v�V�����{�^���̕\��          Excel2007�Ή�
      Sheets("���ͼ��").Shapes("�{�^�� 312").Visible = True
      Sheets("���ͼ��").Shapes("�{�^��1").Visible = True
      Sheets("���ͼ��").Shapes("�{�^��2").Visible = True
      Sheets("���ͼ��").Shapes("�{�^��3").Visible = True
      Sheets("���ͼ��").Shapes("�{�^��4").Visible = True
      Sheets("���ͼ��").Shapes("�{�^��5").Visible = True
      Sheets("���ͼ��").Shapes("�{�^��6").Visible = True
      Sheets("���ͼ��").Shapes("�{�^��7").Visible = True
      Sheets("���ͼ��").Shapes("�{�^��8").Visible = True
      Sheets("���ͼ��").DrawingObjects("�B���U").Visible = True
      Worksheets("���ͼ��").Range("D22") = "�}�^�M����"
      Sheets("�H��}").DrawingObjects("�B���k").Visible = False
      Sheets("�H��}").DrawingObjects("�B���q").Visible = True
      Sheets("���ͼ��").Select
      Range("C23").Select
      a_res = MsgBox("�T�C�Y����͂��ĉ������B���͌�[�m�F]�{�^���������ĉ������B", vbExclamation, "�m�F")
      
      With Sheets("�H��}").DrawingObjects("�\�B��")
          .Visible = True
          .Top = 386.5
          .Height = 177
          .Left = 6
          .Width = 287.25
      End With
      
    Else
      If Worksheets("���ͼ��").Range("C7") = "�I�[�o�[�s���a" Then
msg_9:   Message_9 = "1�����N�̐��@�́H" & Chr(13) & "���ɓ��͂��ĉ������B"
         Title = "1�����N�̐��@����"
         w_Default = "0.02"
         MyValue_9 = InputBox(Message_9, Title, w_Default)
         lank_9 = CStr(MyValue_9)
         If lank_9 <> "" Then
         Else
           Exit Sub
         End If
        If henkan_4 = "1" Then                                             '�I�[�o�[�s���a  �����̎�
          w_kanrikikaku_1 = Sheets("���ͼ��").Range("D7") - Sheets("���ͼ��").Range("D13") + Sheets("���ͼ��").Range("D9")
          w_kanrikikaku_2 = w_kanrikikaku_1 + lank_9
          For i = 1 To 8
            Sheets("���ͼ��").Range("X" & Trim(i + 6)) = w_kanrikikaku_1
            Sheets("���ͼ��").Range("Y" & Trim(i + 6)) = w_kanrikikaku_2
            w_kanrikikaku_1 = w_kanrikikaku_1 + lank_9
            w_kanrikikaku_2 = w_kanrikikaku_2 + lank_9
          Next i
         
          For i = 1 To 8
            Sheets("���ͼ��").Range("D" & Trim(i + 22)) = Sheets("���ͼ��").Range("Z" & Trim(i + 6))
            Sheets("�H��}").TextBoxes("����" & Trim(i)).Caption = Sheets("���ͼ��").Range("D" & Trim(i + 22))
          Next i
        Else
          If henkan_4 = "2" Then                                           '�I�[�o�[�s���a  �~���̎�
            w_kanrikikaku_1 = Sheets("���ͼ��").Range("D7") - Sheets("���ͼ��").Range("D13") + Sheets("���ͼ��").Range("D9")
            w_kanrikikaku_2 = w_kanrikikaku_1 + lank_9
            For i = 1 To 8
              Sheets("���ͼ��").Range("X" & Trim(i + 6)) = w_kanrikikaku_1
              Sheets("���ͼ��").Range("Y" & Trim(i + 6)) = w_kanrikikaku_2
              w_kanrikikaku_1 = w_kanrikikaku_1 + lank_9
              w_kanrikikaku_2 = w_kanrikikaku_2 + lank_9
            Next i
            j = 0
            For i = 7 To 14
              Sheets("���ͼ��").Range("D" & Trim(i + 23 - j)) = Sheets("���ͼ��").Range("Z" & Trim(i))
              j = j + 2
            Next i
            For i = 1 To 8
              Sheets("�H��}").TextBoxes("����" & Trim(i)).Caption = Sheets("���ͼ��").Range("D" & Trim(i + 22))
            Next i
          End If
        End If
         
        Sheets("�H��}").TextBoxes("SUUTI_1").Caption = "(" & Sheets("���ͼ��").Range("D15") & ")"
         
'++++++++ �R�}���h�{�^���̕\���A��\����؂�ւ��� 2012/01/01 +++++++++
        For i = 1 To 8
          With Sheets("���ͼ��").OptionButtons("�{�^��" & Trim(Str(i)))
              .Value = xlOff
              .Visible = True
          End With
        Next i
         
        With Sheets("�H��}").TextBoxes("text3")                                           '�ǉ� 99/ 1/14
            .Visible = True                                                                '�ǉ� 99/ 1/14
            .Font.Size = 18                                                                '�ǉ� 99/ 1/14
            .Caption = "(�s���a  " & Sheets("���ͼ��").Range("D15") & ")"                   '�ǉ� 99/ 1/14
        End With                                                                           '�ǉ� 99/ 1/14
        With Sheets("�H��}").TextBoxes("text1")                                           '�ǉ� 99/ 1/14
            .Visible = True                                                                '�ǉ� 99/ 1/14
            .Caption = "�I�[�o�[�s���a"                                                     '�ǉ� 99/ 1/14
            .Font.Size = 20                                                                '�ǉ� 99/ 1/14
        End With                                                                           '�ǉ� 99/ 1/14
        With Sheets("�H��}").TextBoxes("text2")                                            '�ǉ� 99/ 1/14
            .Visible = True                                                                 '�ǉ� 99/ 1/14
        End With                                                                            '�ǉ� 99/ 1/14
        With Sheets("�H��}").TextBoxes("textp1")                                           '�ǉ� 99/ 1/14
            .Visible = True                                                                 '�ǉ� 99/ 1/14
            .Font.Size = 20                                                                 '�ǉ� 99/ 1/14
            .Caption = "���ް��݌a�����·  " & Sheets("���ͼ��").Range("H16") & "  �ȉ�"     '�ǉ� 99/ 1/14
        End With                                                                            '�ǉ� 99/ 1/14
'        Sheets("�H��}").Range("AG44").Value = "�����}�C�N��"                                '�ǉ� 99/ 1/18
'        Sheets("�H��}").Range("AG45").Value = "���ް���ϲ��"                               '�ǉ� 99/ 1/18
        Sheets("���ͼ��").DrawingObjects("size6").Visible = False
    '�I�v�V�����{�^���̕\��          Excel2007�Ή�
        Sheets("���ͼ��").Shapes("�{�^�� 312").Visible = True
        Sheets("���ͼ��").Shapes("�{�^��1").Visible = True
        Sheets("���ͼ��").Shapes("�{�^��2").Visible = True
        Sheets("���ͼ��").Shapes("�{�^��3").Visible = True
        Sheets("���ͼ��").Shapes("�{�^��4").Visible = True
        Sheets("���ͼ��").Shapes("�{�^��5").Visible = True
        Sheets("���ͼ��").Shapes("�{�^��6").Visible = True
        Sheets("���ͼ��").Shapes("�{�^��7").Visible = True
        Sheets("���ͼ��").Shapes("�{�^��8").Visible = True
        
        Sheets("���ͼ��").DrawingObjects("�B���U").Visible = True
        Worksheets("���ͼ��").Range("D22") = "�I�[�o�[�s���a"
        Sheets("�H��}").DrawingObjects("�B���k").Visible = False
        Sheets("�H��}").DrawingObjects("�B���q").Visible = True
        Sheets("���ͼ��").Select
        Range("C23").Select
        a_res = MsgBox("�T�C�Y����͂��ĉ������B���͌�[�m�F]�{�^���������ĉ������B", vbExclamation, "�m�F")
        
        With Sheets("�H��}").DrawingObjects("�\�B��")
            .Visible = True
            .Top = 386.5
            .Height = 177
            .Left = 6
            .Width = 287.25
        End With
         
      End If
    End If
End Sub

'****************************************
'
'  ���l��\�ɏ����}�N�� ���̂Q
'
'****************************************
Sub �\������_2()
Attribute �\������_2.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("���ͼ��").Unprotect
    Application.Run (XLS_KTRTN + "!�H��}�A�����b�N")
    With Sheets("�H��}").TextBoxes("size1_�O")
        .Visible = False
    End With
    If kioku_3 = "1" Then                           '�I�[�o�[�s���a  �����̎�
      For i = 4 To 11
        Sheets("���ͼ��").Range("D" & Trim(i + 19)) = Sheets("�����v�Z").Cells(i, 8)
        Sheets("���ͼ��").Range("G" & Trim(i + 19)) = Sheets("�����v�Z").Cells(i, 33)
      Next i
    
      For i = 1 To 8
        Sheets("�H��}").TextBoxes("����" & Trim(i)).Caption = Sheets("���ͼ��").Range("D" & Trim(i + 22))
        Sheets("�H��}").TextBoxes("�E��" & Trim(i)).Caption = Sheets("���ͼ��").Range("G" & Trim(i + 22))
      Next i
    Else
      If kioku_3 = "2" Then                         '�I�[�o�[�s���a  �~���̎�
        For i = 4 To 11
          Sheets("���ͼ��").Range("D" & Trim(i + 19)) = Sheets("�����v�Z�Q").Cells(i, 8)
          Sheets("���ͼ��").Range("G" & Trim(i + 19)) = Sheets("�����v�Z�Q").Cells(i, 33)
        Next i
    
        For i = 1 To 8
          Sheets("�H��}").TextBoxes("����" & Trim(i)).Caption = Sheets("���ͼ��").Range("D" & Trim(i + 22))
          Sheets("�H��}").TextBoxes("�E��" & Trim(i)).Caption = Sheets("���ͼ��").Range("G" & Trim(i + 22))
        Next i
      End If
    End If
        
    Sheets("�H��}").TextBoxes("SUUTI_1").Caption = "( " & "��" & Sheets("�����v�Z").Range("G3") & " )"
    Sheets("�H��}").TextBoxes("SUUTI_2").Caption = "( " & Sheets("�����v�Z").Range("E4") & "��" & " )"
    
'++++++++ �R�}���h�{�^���̕\���A��\����؂�ւ��� 2012/01/01 +++++++++
    For i = 1 To 8
      With Sheets("���ͼ��").OptionButtons("�{�^��" & Trim(Str(i)))
          .Value = xlOff
          .Visible = True
      End With
    Next i
    
    Sheets("���ͼ��").Range("C7").Value = "�I�[�o�[�s���a"
    Sheets("���ͼ��").Range("C7").Font.Size = 10
    Sheets("���ͼ��").Range("G16").Value = "���ް��݌a�����·"
    Sheets("���ͼ��").Range("G16").Font.Size = 10
    Sheets("���ͼ��").Range("H16").Value = 0.05
    Sheets("���ͼ��").Range("H16").Font.Size = 11
'    Sheets("�H��}").Range("AG45").Value = "���ް���ϲ��"
'    Sheets("�H��}").Range("AG45").Font.Size = 11

    With Sheets("�H��}").TextBoxes("textp1")                                            '�ǉ� 99/ 1/14
        .Visible = True                                                                 '�ǉ� 99/ 1/14
        .Font.Size = 20                                                                 '�ǉ� 99/ 1/14
        .Caption = "���ް��݌a�����·  " & Sheets("���ͼ��").Range("H16") & "  �ȉ�"     '�ǉ� 99/ 1/14
    End With                                                                            '�ǉ� 99/ 1/14
    With Sheets("�H��}").TextBoxes("text3")                                             '�ǉ� 99/ 1/14
         .Visible = True                                                                 '�ǉ� 99/ 1/14
         .Font.Size = 18                                                                 '�ǉ� 99/ 1/14
         .Caption = "(�s���a  ��" & Sheets("�����v�Z").Range("G3") & ")"                  '�ǉ� 99/ 1/14
     End With                                                                            '�ǉ� 99/ 1/14
     With Sheets("�H��}").TextBoxes("text1")                                            '�ǉ� 99/ 1/14
         .Visible = True                                                                '�ǉ� 99/ 1/14
         .Caption = "�I�[�o�[�s���a"                                                     '�ǉ� 99/ 1/14
         .Font.Size = 20                                                                '�ǉ� 99/ 1/14
     End With                                                                           '�ǉ� 99/ 1/14
    With Sheets("�H��}").TextBoxes("text2")                                            '�ǉ� 99/ 1/14
        .Visible = True                                                                 '�ǉ� 99/ 1/14
    End With                                                                            '�ǉ� 99/ 1/14


    Sheets("���ͼ��").DrawingObjects("size6").Visible = False
    Sheets("���ͼ��").DrawingObjects("�B���U").Visible = False
'�I�v�V�����{�^���̕\��          Excel2007�Ή�
    Sheets("���ͼ��").Shapes("�{�^�� 312").Visible = True
    Sheets("���ͼ��").Shapes("�{�^��1").Visible = True
    Sheets("���ͼ��").Shapes("�{�^��2").Visible = True
    Sheets("���ͼ��").Shapes("�{�^��3").Visible = True
    Sheets("���ͼ��").Shapes("�{�^��4").Visible = True
    Sheets("���ͼ��").Shapes("�{�^��5").Visible = True
    Sheets("���ͼ��").Shapes("�{�^��6").Visible = True
    Sheets("���ͼ��").Shapes("�{�^��7").Visible = True
    Sheets("���ͼ��").Shapes("�{�^��8").Visible = True
    Worksheets("���ͼ��").Range("D22") = "�I�[�o�[�s���a"
    Worksheets("���ͼ��").Range("G22") = "�}�^�M����(�Q�l)"
    Sheets("�H��}").DrawingObjects("�B���k").Visible = False
    Sheets("�H��}").DrawingObjects("�B���q").Visible = False
    Sheets("���ͼ��").Select
    Range("C23").Select
    a_res = MsgBox("�T�C�Y����͂��ĉ������B���͌�[�m�F]�{�^���������ĉ������B", vbExclamation, "�m�F")
    
    With Sheets("�H��}").DrawingObjects("�\�B��")
        .Visible = True
        .Top = 386.5
        .Height = 177
        .Left = 6
        .Width = 287.25
    End With
    
End Sub

'******************************************
'
'  �T�C�Y�w���H��}�ɕ\�����邽�߂̃}�N��
'
'******************************************
Sub �m�F()

       wk_top = 417    ' 399.75
       wk_height = 150     '153.75
       wk_left = 6
       wk_width = 287.25
         
       For i = 1 To 8
           If Worksheets("���ͼ��").Range("C" & Trim(i + 22)) <> "�|" Then
              wk_top = wk_top + 17.75
              wk_height = wk_height - 17.75
           Else
              GoTo ����
           End If
       Next i
����:  With Sheets("�H��}").DrawingObjects("�\�B��")
           .Visible = True
           .Top = wk_top
           .Height = wk_height
           .Left = wk_left
           .Width = wk_width
       End With
End Sub

'******************************************
'
' �T�C�Y�w��\��
'
'******************************************

Sub �T�C�Y1()
Attribute �T�C�Y1.VB_ProcData.VB_Invoke_Func = " \n14"
    w_�{�^�� = 1
    �T�C�Y
End Sub

Sub �T�C�Y2()
Attribute �T�C�Y2.VB_ProcData.VB_Invoke_Func = " \n14"
    w_�{�^�� = 2
    �T�C�Y
End Sub

Sub �T�C�Y3()
Attribute �T�C�Y3.VB_ProcData.VB_Invoke_Func = " \n14"
    w_�{�^�� = 3
    �T�C�Y
End Sub

Sub �T�C�Y4()
Attribute �T�C�Y4.VB_ProcData.VB_Invoke_Func = " \n14"
    w_�{�^�� = 4
    �T�C�Y
End Sub

Sub �T�C�Y5()
Attribute �T�C�Y5.VB_ProcData.VB_Invoke_Func = " \n14"
    w_�{�^�� = 5
    �T�C�Y
End Sub

Sub �T�C�Y6()
Attribute �T�C�Y6.VB_ProcData.VB_Invoke_Func = " \n14"
    w_�{�^�� = 6
    �T�C�Y
End Sub

Sub �T�C�Y7()
Attribute �T�C�Y7.VB_ProcData.VB_Invoke_Func = " \n14"
    w_�{�^�� = 7
    �T�C�Y
End Sub

Sub �T�C�Y8()
Attribute �T�C�Y8.VB_ProcData.VB_Invoke_Func = " \n14"
    w_�{�^�� = 8
    �T�C�Y
End Sub

Sub �T�C�Y()
Attribute �T�C�Y.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("�����v�Z").Unprotect
    Sheets("�����v�Z�Q").Unprotect
    With Sheets("�H��}").TextBoxes("text2")
        .Visible = False
    End With
    If Sheets("���ͼ��").Range("D22") = "�}�^�M����" And Sheets("���ͼ��").DrawingObjects("�B���U").Visible = True Then
'       Sheets("�H��}").Range("AG44").Value = "�����}�C�N��"                          '�ǉ� 99/ 1/18
'       Sheets("�H��}").Range("AG45").Value = "�����}�C�N��"                          '�ǉ� 99/ 1/18
       With Sheets("�H��}").TextBoxes("micro_text1")                                '�ǉ� 99/ 1/18
           .Caption = "����ϲ��"                                                     '�ǉ� 99/ 1/18
           .Font.Size = 9                                                            '�ǉ� 99/ 1/18
           .Visible = True                                                           '�ǉ� 99/ 1/18
       End With                                                                      '�ǉ� 99/ 1/18

       With Sheets("�H��}").TextBoxes("text1")
           .Visible = True
       End With
       With Sheets("�H��}").TextBoxes("text3")
           .Visible = True
       End With
       If henkan_2 = "1" Then                            '�}�^�M����  �����̎�
          With Sheets("�H��}").TextBoxes("size1_�O")
              .Visible = True
              .Caption = Format(Application.RoundUp(Sheets("�����v�Z").Range("F" & Trim(w_�{�^�� + 3)), 2), "##0.00") & "  �`  " & _
                         Format(Application.RoundUp(Sheets("�����v�Z").Range("F" & Trim(w_�{�^�� + 4)), 2), "##0.00") & "  (�ʏ�)"
              .Font.Size = 26
          End With
       Else
         If henkan_2 = "2" Then                           '�}�^�M����  �~���̎�
            For i = 4 To 11
              Sheets("�����v�Z�Q").Range("AH" & Trim(i)) = Sheets("�����v�Z�Q").Range("F" & Trim(i + 1))
              Sheets("�����v�Z�Q").Range("AI" & Trim(i)) = Sheets("�����v�Z�Q").Range("F" & Trim(i))
            Next i
            With Sheets("�H��}").TextBoxes("size1_�O")
                .Visible = True
                .Caption = Format(Application.RoundUp(Sheets("�����v�Z�Q").Range("AH" & Trim(w_�{�^�� + 3)), 2), "##0.00") & "  �`  " & _
                           Format(Application.RoundUp(Sheets("�����v�Z�Q").Range("AI" & Trim(w_�{�^�� + 3)), 2), "##0.00") & "  (�ʏ�)"
                .Font.Size = 26
            End With
         End If
       End If
    Else
       If Sheets("���ͼ��").Range("D22") = "�I�[�o�[�s���a" And Sheets("���ͼ��").DrawingObjects("�B���U").Visible = True Then
'          Sheets("�H��}").Range("AG44").Value = "�n�a�c�����"                                '�ǉ� 99/ 1/18
          Sheets("�H��}").Range("AG46").Font.Size = 11                                       '�ǉ� 99/ 1/18
'          Sheets("�H��}").Range("AG45").Value = "�n�a�c�����"                                '�ǉ� 99/ 1/18
          Sheets("�H��}").Range("AG47").Font.Size = 11                                       '�ǉ� 99/ 1/18
          With Sheets("�H��}").TextBoxes("micro_text1")                                      '�ǉ� 99/ 1/18
              .Caption = "OBD�����"                                                          '�ǉ� 99/ 1/18
              .Font.Size = 8                                                                  '�ǉ� 99/ 1/18
              .Visible = True                                                                 '�ǉ� 99/ 1/18
          End With                                                                            '�ǉ� 99/ 1/18
          With Sheets("�H��}").TextBoxes("micro_text11")                                      '�ǉ� 99/ 1/18
              .Caption = "���ް���"                                                           '�ǉ� 99/ 1/18
              .Font.Size = 8                                                                  '�ǉ� 99/ 1/18
              .Visible = False                                                                '�ǉ� 99/ 1/18
          End With                                                                            '�ǉ� 99/ 1/18
          With Sheets("�H��}").TextBoxes("micro_text12")                                      '�ǉ� 99/ 1/18
              .Caption = "ϲ��"                                                               '�ǉ� 99/ 1/18
              .Font.Size = 8                                                                  '�ǉ� 99/ 1/18
              .Visible = False                                                                '�ǉ� 99/ 1/18
          End With                                                                            '�ǉ� 99/ 1/18
          With Sheets("�H��}").TextBoxes("text1")
              .Visible = True
          End With
          With Sheets("�H��}").TextBoxes("text3")
              .Visible = True
          End With
          If henkan_4 = "1" Then                            '�I�[�o�[�s���a  �����̎�
             With Sheets("�H��}").TextBoxes("size1_�O")
                 .Visible = True
                 .Caption = Format(Sheets("���ͼ��").Range("X" & Trim(w_�{�^�� + 6)), "##0.000") & "  �`  " & _
                            Format(Sheets("���ͼ��").Range("Y" & Trim(w_�{�^�� + 6)), "##0.000") & "  (�ʏ�)"
                 .Font.Size = 26
             End With
          Else
            If henkan_4 = "2" Then                          '�I�[�o�[�s���a  �~���̎�
               j = 0
               For i = 7 To 14
                 Sheets("���ͼ��").Range("AB" & Trim(i + 7 - j)) = Sheets("���ͼ��").Range("X" & Trim(i))
                 Sheets("���ͼ��").Range("AC" & Trim(i + 7 - j)) = Sheets("���ͼ��").Range("Y" & Trim(i))
                 j = j + 2
               Next i
               With Sheets("�H��}").TextBoxes("size1_�O")
                   .Visible = True
                   .Caption = Format(Sheets("���ͼ��").Range("AB" & Trim(w_�{�^�� + 6)), "##0.000") & "  �`  " & _
                              Format(Sheets("���ͼ��").Range("AC" & Trim(w_�{�^�� + 6)), "##0.000") & "  (�ʏ�)"
                   .Font.Size = 26
               End With
            End If
          End If
       Else
          If Sheets("���ͼ��").Range("D22") = "�I�[�o�[�s���a" And Sheets("���ͼ��").Range("G22") = "�}�^�M����(�Q�l)" Then
'             Sheets("�H��}").Range("AG44").Value = "�n�a�c�����"                                '�ǉ� 99/ 1/18
             Sheets("�H��}").Range("AG46").Font.Size = 11                                       '�ǉ� 99/ 1/18
'             Sheets("�H��}").Range("AG45").Value = "�n�a�c�����"                                '�ǉ� 99/ 1/18
             Sheets("�H��}").Range("AG47").Font.Size = 11                                       '�ǉ� 99/ 1/18
             With Sheets("�H��}").TextBoxes("micro_text1")                                      '�ǉ� 99/ 1/18
                 .Caption = "OBD�����"                                                          '�ǉ� 99/ 1/18
                 .Font.Size = 8                                                                  '�ǉ� 99/ 1/18
                 .Visible = True                                                                 '�ǉ� 99/ 1/18
             End With                                                                            '�ǉ� 99/ 1/18

             With Sheets("�H��}").TextBoxes("text1")
                 .Visible = True
             End With
             With Sheets("�H��}").TextBoxes("text3")
                 .Visible = True
             End With
             If kioku_3 = "1" Then                           '�I�[�o�[�s���a  �����̎�
                With Sheets("�H��}").TextBoxes("size1_�O")
                    .Visible = True
                    .Caption = Format(Sheets("�����v�Z").Range("G" & Trim(w_�{�^�� + 3)), "###0.000") & "  �`  " & _
                               Format(Sheets("�����v�Z").Range("G" & Trim(w_�{�^�� + 4)), "###0.000") & "  (�ʏ�)"
                    .Font.Size = 26
                End With
             Else
               If kioku_3 = "2" Then                         '�I�[�o�[�s���a  �~���̎�
                  For i = 4 To 11
                    Sheets("�����v�Z�Q").Range("AL" & Trim(i)) = Sheets("�����v�Z�Q").Range("G" & Trim(i + 1))
                    Sheets("�����v�Z�Q").Range("AM" & Trim(i)) = Sheets("�����v�Z�Q").Range("G" & Trim(i))
                  Next i
                  With Sheets("�H��}").TextBoxes("size1_�O")
                      .Visible = True
                      .Caption = Format(Sheets("�����v�Z�Q").Range("AL" & Trim(w_�{�^�� + 3)), "###0.000") & "  �`  " & _
                                 Format(Sheets("�����v�Z�Q").Range("AM" & Trim(w_�{�^�� + 3)), "###0.000") & "  (�ʏ�)"
                      .Font.Size = 26
                  End With
               End If
             End If
          End If
       End If
    End If
             
    With Sheets("�H��}").TextBoxes("size2")
        .Visible = True
        .Caption = "�������ޕ\�ɂ��"
        .Font.Size = 6
    End With
'    With Sheets("�H��}").TextBoxes("size4�̍�")
'        .Visible = True
'        .Caption = "�������ޕ\�ɂ��"
'        .Font.Size = 12
'    End With
'    With Sheets("�H��}").TextBoxes("size4")
'        .Visible = True
'        .Caption = "��"
'        .Font.Size = 16
'    End With
 
End Sub

'******************************************
'
' �����}�[�N�}�� ���ʃ}�N��
'
'******************************************

Sub ����()
Attribute ����.VB_ProcData.VB_Invoke_Func = " \n14"
    L_act_sheet = ActiveSheet.Name
    Application.Run (XLS_KTRTN + "!�����ԍ�"), XLS_FILE, L_act_sheet
End Sub

'******************************************
'
' ���������L�� ���ʃ}�N��       ���L �F ���O�p�̖��O�� "�O�p1","�O�p2",�E�E�E�E"�O�px"�ƕύX���Ă����Ă��������B
'
'       ����1.�H��}�̃t�@�C���� �iEX:"EO000101.XLS"�j
'       ����2.�A�N�e�B�u�V�[�g��
'       ����3.���������̍s��     �iMAX�l�j
'       ����4.�������𗓂̐擪�s�̂w���̈ʒl
'       ����5.�������𗓂̐擪�s�̂x���̔N�����̈ʒl
'       ����6.�������𗓂̐擪�s�̂x���̉����m���̈ʒl
'       ����7.�������𗓂̐擪�s�̂x���̉����ӏ��̈ʒl
'       ����8.�������𗓂̐擪�s�̂x���̉������R�̈ʒl
'       ����9.�������𗓂̐擪�s�̂x���̌���̈ʒl

'****************************************************


Sub ��������()
Attribute ��������.VB_ProcData.VB_Invoke_Func = " \n14"
    w_max = 5
    w_y = 64        '�w�� -- (�ŏ��̈ʒu)�ɐݒ�
                    '�x�� -- ���ꂼ��L������ʒu��ݒ肷��
    w_x1 = 4        '�N����
    w_x2 = 7        '��������
    w_x3 = 12       '�����ӏ�
    w_x4 = 18       '�������R
    w_x5 = 0        '����   �i���󗓂Ȃ��Ƃ� �O�j
    L_act_sheet = ActiveSheet.Name
    Application.Run (XLS_KTRTN + "!�����L��"), XLS_FILE, L_act_sheet, w_max, w_y, w_x1, w_x2, w_x3, w_x4, w_x5
End Sub
'****************************************************
'*TEST TEST TEST TEST TEST TEST TEST TEST           *
'   ������������������L���������ƂɎ������������A   *
'   �����������N���A�[���������͍��O�p�}�[�N���\��     *
'   ����Ă��Ȃ��Ɠo�^�ł��Ȃ��̂ŁA�O�p�\��() �𗬂��B*
'****************************************************
Sub �O�p�\��()                                  '�e�X�g�p --- �S���O�p�\��
Attribute �O�p�\��.VB_ProcData.VB_Invoke_Func = " \n14"
    For i = 1 To 5
        Sheets("�H��}").DrawingObjects("�O�p" & i).Visible = True
    Next i
End Sub
'****************************************************

'   �O�쌟��        listbox("�O��")�Ɗ֘A�t����

'****************************************************
Sub �O�쌟��()
Attribute �O�쌟��.VB_ProcData.VB_Invoke_Func = " \n14"
        '*****************************************
        '���ͼ�đO�H���A��H���̃Z���ʒu�������Ƃ���
        '*****************************************
    y1 = 4      '�O�H��-�c�ʒu
    x1 = 3      '�O�H��-���ʒu
    y2 = 4      '��H��-�c�ʒu
    x2 = 4      '��H��-���ʒu
    
    Application.Run (XLS_KTRTN + "!�O��O��"), y1, x1, y2, x2
    
End Sub
'******************************************
'
' �H���m�n
'
'******************************************
Sub �H��NO_click()
Attribute �H��NO_click.VB_ProcData.VB_Invoke_Func = " \n14"
    Static w_a
    
    If w_a <> 1 Then
        With ActiveSheet.ListBoxes("�H��_���X�g")
            .RemoveAllItems
            .AddItem Text:="�e�P", Index:=1
            .AddItem Text:="�e�Q", Index:=2
            .AddItem Text:="�e�R", Index:=3
            .AddItem Text:="���ݸ", Index:=4
        End With
        Sheets("���ͼ��").ListBoxes("�H��_���X�g").Visible = True
        Sheets("���ͼ��").ListBoxes("�H��_���X�g").Enabled = True
        w_a = 1
    Else
        Sheets("���ͼ��").ListBoxes("�H��_���X�g").Visible = False
        Sheets("���ͼ��").ListBoxes("�H��_���X�g").Enabled = False
        w_a = 0
    End If
End Sub
Sub �H��NO_select()
Attribute �H��NO_select.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim �H��_select As String
    
    With ActiveSheet.ListBoxes("�H��_���X�g")
        �H��_select = .List(.ListIndex)
        If �H��_select = "���ݸ" Then
            �H��_select = ""
        End If
        ActiveSheet.Range("B4") = �H��_select       '�C��
    End With
    
    �H��NO_click
        
End Sub

'******************************************************************************************************************
Sub test1()
Attribute test1.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("���ͼ��").DrawingObjects("size6").Visible = True
    Sheets("���ͼ��").DrawingObjects("�B���U").Visible = True
'    With Sheets("�H��}").DrawingObjects("�\�B��")
'        .Visible = True
'    End With
    Sheets("�H��}").DrawingObjects("�B���k").Visible = True
    Sheets("�H��}").DrawingObjects("�B���q").Visible = True
End Sub


Sub ccc()
Attribute ccc.VB_ProcData.VB_Invoke_Func = " \n14"
    With Sheets("�H��}").TextBoxes("size2")           '�ǉ� 99/ 1/18
        .Visible = False                               '�ǉ� 99/ 1/18
    End With                                           '�ǉ� 99/ 1/18
    With Sheets("�H��}").TextBoxes("size4")           '�ǉ� 99/ 1/18
        .Visible = False                               '�ǉ� 99/ 1/18
    End With                                           '�ǉ� 99/ 1/18
End Sub


Attribute VB_Name = "�}�N��"
'*****************************************************

'   �H������  �F �V�F�[�r���O
'
'   �ݔ�GR    �F �t���g�y�r�s�|�V

'*****************************************************
'   �萔��`�̈�
'*****************************************************
'XLS�t�@�C��
'Public Const XLS_FILE = "EP000101.XLSM"
Public XLS_FILE As String

'CSV�t�@�C��
'Public Const DATA_FILE = "SV.csv"                          'FOR TEST
'Public Const DATA_FILE_PATH = "\�͍�\csv\SV.csv"           'FOR TEST

Public Const DATA_FILE = "EP0001.CSV"
Public Const DATA_FILE_PATH = "C:\CS50\EP0001.CSV"

'*****************************************************

'       ���ʃt�@�C��

'*****************************************************

'���ʃ}�N��
Public Const XLS_KTRTN = "EPC007.XLSM"
Public Const KYOTU_FILE_PATH = "S:\CS50\DOCUMENT\EPC007.XLSM"
'Public Const KYOTU_FILE_PATH = "D:\CS50_Document\EPC007.XLSM"

'*****************************************************

'       ���F��

'*****************************************************

Public Const HANKO_FILE = "EPC006.XLSM"
Public Const HANKO_FILE_PATH = "S:\CS50\DOCUMENT\EPC006.XLSM"
'Public Const HANKO_FILE_PATH = "D:\CS50_Document\EPC006.XLSM"

Public Const EPC006_CSV = "EPC006.CSV"
Public Const EPC006_CSV_PATH = "C:\CS50\EPC006.CSV"
'*****************************************************

'       �ϐ��錾

'*****************************************************

Public w_name, w_���@, w_���, w_����, w_����, w_���W���[��, w_���݂�������, w_�{�^��

'*****************************************************
'
' auto_open �}�N��
'
'*****************************************************'
Sub AUTO_OPEN()
Attribute AUTO_OPEN.VB_ProcData.VB_Invoke_Func = " \n14"

    XLS_FILE = ActiveWorkbook.Name
    
    oldStatusBar = Application.DisplayStatusBar  '�K��l�̕ۑ�
    Application.DisplayStatusBar = True '�ð���ް�̕\��

    kyotu_macro_open                                '���ʃ}�N���t�@�C���n�o�d�m

    Windows(XLS_FILE).Activate
'    ActiveWindow.WindowState = xlMinimized
    Windows(XLS_FILE).Activate


    If Trim(Sheets("�H��}").TextBoxes("text_hinban").Caption) = "" Then

       Application.StatusBar = "���΂炭���҂�������..�D�D"
       Application.Cursor = xlWait  '�����v�^���ٕ\��

       Sheets("�\�t�����").Visible = True

       open_file                                    '�b�r�u�t�@�C�����J��

       csv_harituke                                 'csv�t�@�C���̓��e��\�t����Ăɓ\��t����

       close_file                                   '�b�r�u�t�@�C�������

       '�V�X�e�����t���[���T�v���[�X�����`�Ńt�H�[�}�b�g�������^�ɕϊ����H��}�ɓ\��t����
       Dim m_wk, m_ans
       m_wk = Sheets("���ͼ��").Range("K3")
       m_ans = "'" + CStr(Format(m_wk, "�fyy�Dm�Dd"))
       Sheets("�H��}").Range("BB7") = m_ans

       ���͒l_clear

       Sheets("�\�t�����").Visible = False

'       �������� �| �V�}���s
'           ���� ----
'            "XLŞ�ٖ�","��è�޼�Ė�","�c�ʒu","���ʒu-�N����","���ʒu-����No","���ʒu-�����ӏ�","���ʒu-�������R","���ʒu-����(�Ȃ��Ƃ��O)"
        Sheets("�H��}").Activate
        Application.Run (XLS_KTRTN + "!�V�}"), XLS_FILE, ActiveSheet.Name, 64, 4, 7, 12, 18, 0

    End If

    text_harituke   '���͗p�e�L�X�g�̃t�H���g�ݒ�i�쐬�ρA�� �ɂ�����炸�j

    HANKO_open_file                                     '���F�󋤒ʃ}�N���t�@�C�����J��


    Windows(XLS_FILE).Activate

    Application.Run (XLS_KTRTN + "!���͎Q�Ɛݒ�")

    Sheets("�H��}").Select

    Application.Run (XLS_KTRTN + "!�Q�Ɛݒ�")            '���ʃ}�N���̒��ڎQ�Ɛݒ�

    Application.Run (XLS_KTRTN + "!�H��}���b�N")

    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    Application.Cursor = xlNormal
    ActiveWindow.WindowState = xlMaximized

    Range("A1").Select

    '�������ݏI����
   For i = 1 To 10
      Beep  '�r�[�v�����Ȃ炵�܂�
   Next i

End Sub
'*****************************************************
'
'  ���͒l��clear ����
'
'*****************************************************'
Sub ���͒l_clear()
Attribute ���͒l_clear.VB_ProcData.VB_Invoke_Func = " \n14"

    Sheets("���ͼ��").Range("H13") = ""
    Sheets("���ͼ��").Range("K18") = ""
    Sheets("���ͼ��").Range("B4") = ""
   
    '�T�C�Y�\
    Sheets("���ͼ��").Range("C23:C30") = "�|"
    Sheets("�����v�Z").Range("C9") = 0.01               '�ǉ� 99/ 1/18
    Sheets("�����v�Z�Q").Range("C9") = -0.01            '�ǉ� 99/ 1/26
    
    �T�C�Y�i�łȂ��Ƃ�_1              'call "�T�C�Y�i���̐ݒ����������"
    
    For i = 1 To 8
        With Sheets("���ͼ��").OptionButtons("�{�^��" & Trim(Str(i)))
            .Value = xlOff
        End With
    Next i
   
         
End Sub
'*****************************************************
'
' auto_close �}�N��
'
'*****************************************************'
Sub auto_close()
Attribute auto_close.VB_ProcData.VB_Invoke_Func = " \n14"


    Application.Run (XLS_KTRTN + "!�H��}�A�����b�N")
    Application.Run (XLS_KTRTN + "!���͎Q�Ɖ���"), XLS_FILE
    Application.Run (XLS_KTRTN + "!�Q�Ɖ���"), XLS_FILE            '���ʃ}�N���̒��ڎQ�Ɖ���
    
    
    ''EP0001R.CSV�̕ۑ��ꏊ��荞
    AREA_G = Workbooks(EPC006_CSV).Sheets("Epc006").Range("C1") & "\EP0001R.CSV"
    
    Application.Run (HANKO_FILE + "!Epc006_close")            'EXCEL��Ԃ̐ݒ�
       
    '���F�󋤒ʃ}�N���t�@�C���b�k�n�r�d
    Workbooks(HANKO_FILE).Close savechanges:=False
    
    '��Əꏊ�b�r�u�t�@�C���b�k�n�r�d
    Workbooks(EPC006_CSV).Close savechanges:=False
    
    '���ʃ}�N���t�@�C���b�k�n�r�d
    Windows(XLS_KTRTN).Close savechanges:=False

   '�����[�X��Ԃ�EP0001R.CSV�ɏ���
    Open AREA_G For Output As #1
    'If Mid(Sheets("�H��}").Range("AM1"), 1, 2) = "�l�k" Then
    '    w_ok = Sheets("�H��}").Range("AM1")
    If Mid(Sheets("�H��}").Range("AS1"), 1, 2) = "�l�k" Then
        w_ok = Sheets("�H��}").Range("AS1")
        Write #1, w_ok:
    Else
        Write #1, "NO":
    End If
    Close #1

    ''''�ۑ�����EXCEL�I��
    On Error Resume Next
    Workbooks(XLS_FILE).Save
'    Application.Quit
    '�ۑ��m�F������邽�߁A�ۑ��ς݂ɂ���
'    ThisWorkbook.Saved = True
    If Workbooks.Count <= 1 Then Application.CommandBars.FindControl(ID:=752).Execute
    ThisWorkbook.Close False
    
End Sub

'*****************************************************
'
' open_file �}�N��
'
'*****************************************************'
Sub open_file()
Attribute open_file.VB_ProcData.VB_Invoke_Func = " \n14"

    '�b�r�u�t�@�C���n�o�d�m
    Workbooks.Open filename:=DATA_FILE_PATH, ReadOnly _
         :=True
         
'   ActiveWindow.WindowState = xlMinimized

End Sub
'*****************************************************
'
'  ���ʃ}�N���t�@�C���n�o�d�m �}�N��
'
'*****************************************************'
Sub kyotu_macro_open()
Attribute kyotu_macro_open.VB_ProcData.VB_Invoke_Func = " \n14"
    '���ʃ}�N���t�@�C���n�o�d�m

    Workbooks.Open filename:=KYOTU_FILE_PATH, ReadOnly _
         :=True
'    ActiveWindow.WindowState = xlMinimized

End Sub

'*****************************************************
'
' ���F�� �}�N��
'
'*****************************************************'
Sub HANKO_open_file()
Attribute HANKO_open_file.VB_ProcData.VB_Invoke_Func = " \n14"

    '��Ԃb�r�u�t�@�C���n�o�d�m
    Workbooks.Open filename:=EPC006_CSV_PATH, ReadOnly _
         :=True

'    ActiveWindow.WindowState = xlMinimized

    '���F�󋤒ʃ}�N���t�@�C���n�o�d�m
    Workbooks.Open filename:=HANKO_FILE_PATH, ReadOnly _
         :=True
  
'    ActiveWindow.WindowState = xlMinimized

End Sub

'*****************************************************
'
' close_file �}�N��
'
'*****************************************************
Sub close_file()
Attribute close_file.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.DisplayAlerts = False
    '�b�r�u�t�@�C���b�k�n�r�d
    Windows(DATA_FILE).Close

    Application.DisplayAlerts = True
End Sub


'*****************************************************
' csv_harituke �}�N��
'
' �E csv�t�@�C���̓��e��\�t����Ăɓ\��t����
' �E �e���ڂ̒l�����߂�}�N�����Ă�
'
'*****************************************************
Sub csv_harituke()
Attribute csv_harituke.VB_ProcData.VB_Invoke_Func = " \n14"

    '''�b�r�u�t�@�C���̓��e���R�s�[���y�[�X�g

    Windows(DATA_FILE).Activate
    Range("A1:AZ10").Select
    Selection.Copy
    Windows(XLS_FILE).Activate
    Sheets("�\�t�����").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select

    data_bunbetu        '�\�t����� �� ���ͼ��

    hafure_get          '���U��v�Z

    �K�i�ݒ�            '���`�E���؋K�i�ݒ�

'    �����\��            '�����\���̐ݒ�   2003/4/17DEL

End Sub

'*****************************************************
'
' DATA_BUNBETU ϸ�(�\�t����Ă��ް�����ͼ�Ăɺ�߰
'
'*****************************************************
Sub data_bunbetu()

    '�^��
    With Sheets("�H��}").TextBoxes("text0")
        .Caption = Mid(Sheets("�\�t�����").Range("A1"), 1, 3)
        .Font.Size = 48
    End With
    
    '�O�H��
    Sheets("���ͼ��").Range("C4") = Sheets("�\�t�����").Range("E1")
    
    '��H��
    Sheets("���ͼ��").Range("D4") = Sheets("�\�t�����").Range("F1")
    
    ' ���i�ǔԂ̐F�t�� 2010/12/09 ADD
    If Right(Worksheets("�\�t�����").Range("A1"), 2) = "00" Then
        Worksheets("�H��}").Shapes("���i�ǔ�").Visible = False
    Else
        Worksheets("�H��}").Shapes("���i�ǔ�").Visible = True
    End If

    ''��A���Ή�05/02/08  (From)
    If Sheets("�\�t�����").Range("B1") = "�V�F�[�r���O�P" Then
       l_disp = True
       l_DATA = Mid(Sheets("�\�t�����").Range("A1"), 8, 1) & "�o��"
    ElseIf Sheets("�\�t�����").Range("B1") = "�V�F�[�r���O�Q" Then
       l_disp = True
       l_DATA = Trim(Val(Mid(Sheets("�\�t�����").Range("A1"), 8, 1)) + 1) & "�o��"
    Else
       l_disp = False
       l_DATA = ""
    End If
    With Sheets("�H��}").DrawingObjects("type1_txt")
        .Visible = l_disp
        .Caption = l_DATA
        .Font.Size = 40
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ''��A���Ή�05/02/08  (To)
    
    w_���� = Sheets("�\�t�����").Range("I3")    '���ގ��� or �I�[�o�[�s���a
    w_��� = Sheets("�\�t�����").Range("J3")    '��������l
    w_���� = Sheets("�\�t�����").Range("K3")    '���������l
    
    Sheets("���ͼ��").Range("D7") = w_����
    '��������l
    Sheets("���ͼ��").Range("D8") = w_���
    '���������l
    Sheets("���ͼ��").Range("D9") = w_����
    '���`�E���؋K�i�O
    Sheets("���ͼ��").Range("D10") = Sheets("�\�t�����").Range("E3")
    '���W���[��
    w_���W���[�� = Sheets("�\�t�����").Range("A3")
    Sheets("���ͼ��").Range("H7") = w_���W���[��
    '���͊p
    Sheets("���ͼ��").Range("H8") = Sheets("�\�t�����").Range("B3")
    '����
    Sheets("���ͼ��").Range("H9") = Sheets("�\�t�����").Range("C3")
    '�l�W���p�ƕ���
    If Sheets("�\�t�����").Range("F3") <> "" And Sheets("�\�t�����").Range("G3") <> "" Then
        Sheets("���ͼ��").Range("H11") = CStr(Sheets("�\�t�����").Range("F3")) + Sheets("�\�t�����").Range("G3")
    Else
        Sheets("���ͼ��").Range("H11") = ""
    End If
    '���݂������蒷��
    w_���݂������� = Sheets("�\�t�����").Range("D3")
    Sheets("���ͼ��").Range("H15") = Application.RoundUp((w_���݂������� + (0.375 * w_���W���[��)) * 4, 0)
    '�O�a
    Sheets("���ͼ��").Range("N7") = Sheets("�\�t�����").Range("R3")
'    '�����ݸމ��H���������i���݂������蒷���Ɠ��j
'�폜 (K.Y)    Sheets("���ͼ��").Range("H12") = Sheets("���ͼ��").Range("H14")
    
    '�����ݸ޶���
    Sheets("���ͼ��").Range("D18") = Sheets("�\�t�����").Range("C4")
    '�����ݸޱ��ް
    Sheets("���ͼ��").Range("D19") = Sheets("�\�t�����").Range("C5")
    
    '****************************************************
    '�}�^�M�̏ꍇ�ƃI�[�o�[�s���̏ꍇ�œ\�蕪����
    '****************************************************
    
    If Sheets("�\�t�����").Range("H3") = "O" Then               '�I�[�o�[�s���̏ꍇ
       
       Sheets("���ͼ��").Range("C7").Value = "�I�[�o�[�s���a"
       Sheets("���ͼ��").Range("C7").Font.Size = 10
       Sheets("���ͼ��").Range("C15").Value = "�s���a"
       Sheets("���ͼ��").Range("C15").Font.Size = 11
       Sheets("���ͼ��").Range("G16").Value = "���ް��݌a�����·"
       Sheets("���ͼ��").Range("G16").Font.Size = 10
       Sheets("���ͼ��").Range("H16").Value = 0.05
       Sheets("���ͼ��").Range("H16").Font.Size = 11
'       Sheets("�H��}").Range("AG45").Value = "���ް���ϲ��"
'       Sheets("�H��}").Range("AG45").Font.Size = 11
       With Sheets("�H��}").TextBoxes("textp1")                                            '�ǉ� 99/ 1/14
           .Visible = True                                                                 '�ǉ� 99/ 1/14
           .Font.Size = 20                                                                 '�ǉ� 99/ 1/14
           .Caption = "���ް��݌a�����·  " & Sheets("���ͼ��").Range("H16") & "  �ȉ�"     '�ǉ� 99/ 1/14
       End With                                                                            '�ǉ� 99/ 1/14

       
       With Sheets("�H��}").TextBoxes("text1")
                            .Caption = "�I�[�o�[�s���a"
                            .Font.Size = 20       '�ύX
                            .Visible = True
       End With
       With Sheets("�H��}").TextBoxes("micro_text11")
                            .Caption = "���ް���"
                            .Font.Size = 8
                            .Visible = True
       End With
       With Sheets("�H��}").TextBoxes("micro_text12")
                            .Caption = "ϲ��"
                            .Font.Size = 8
                            .Visible = True
       End With
       With Sheets("�H��}").TextBoxes("micro_text1")
                            .Visible = False
       End With
       
        '�s���a
        Sheets("���ͼ��").Range("D15") = "��" & Sheets("�\�t�����").Range("L3")
        With Sheets("�H��}").TextBoxes("text3")                                           '�ǉ� 99/ 1/14
            .Visible = True                                                                '�ǉ� 99/ 1/14
            .Font.Size = 18                                                                '�ǉ� 99/ 1/14
            .Caption = "(�s���a  " & Sheets("���ͼ��").Range("D15") & ")"                   '�ǉ� 99/ 1/14
       End With
                                                                                  '�ǉ� 99/ 1/14
       '���ް��݌a �Ǘ��K�i
       If w_���� = "" Or w_���� = "" Then
           Sheets("���ͼ��").Range("E7") = ""
       Else
           Sheets("���ͼ��").Range("E7") = CStr(Application.Round(w_���� + w_���� + 0.02, 2)) & "�}0.02"
       End If
       
    ElseIf Sheets("�\�t�����").Range("H3") = "M" Then                '�}�^�M�����̏ꍇ
    
       Sheets("���ͼ��").Range("D15") = Sheets("�\�t�����").Range("L3")
       '�}�^�M�̏ꍇ�̐��@�v�Z
       If w_���� = "" Or w_���� = "" Or w_��� = "" Then
           Sheets("���ͼ��").Range("E7") = ""
       Else
       Sheets("���ͼ��").Range("E7") = CStr(Application.Round((w_���� + w_���) / 2 + w_���� - 0.015, 2)) & "�}0.01"
       End If
    Else
    
       Sheets("���ͼ��").Range("C7").Value = ""
       Sheets("���ͼ��").Range("E7").Value = ""
       Sheets("���ͼ��").Range("D15").Value = ""
       Sheets("���ͼ��").Range("C15").Value = ""
       Sheets("���ͼ��").Range("G16").Value = ""
       Sheets("���ͼ��").Range("H16").Value = ""
       Sheets("���ͼ��").Range("H8").Value = ""
    End If
    
    
End Sub
'*****************************************************
'
' ���`�E���؋K�i �ݒ�
'
'*****************************************************

Sub �K�i�ݒ�()
Attribute �K�i�ݒ�.VB_ProcData.VB_Invoke_Func = " \n14"
    w_���x = Sheets("�\�t�����").Range("E3")
    w_�K�i = Sheets("�\�t�����").Range("M3")
    
    Sheets("���ͼ��").Unprotect                'yamauchi
     
    '�O�񕪂̃N���A�[
    For i = 7 To 16
        Sheets("���ͼ��").Cells(i, 11) = ""
    Next i
    
    Sheets("���ͼ��").Range("K7:K16").Interior.ColorIndex = 40              '�O���[
    Sheets("���ͼ��").Range("E10").Interior.ColorIndex = 40
    
    If Mid(w_���x, 4, 1) = "4" Then                                         '"JIS�S��"�̏ꍇ
        If w_�K�i = "�L" Then
            With Sheets("���ͼ��").Buttons("�{�^�� 114")
                                    .Visible = False
            End With
            'Sheets("���ͼ��").Range("E10") = "���}�n�S��"
            Sheets("���ͼ��").Range("E10").Formula = "=IF(D10<>"""",""���}�n""&MID(D10,4,1)&""��"")"         '�ǉ� 99/ 1/14
            Sheets("���ͼ��").Range("K7:K16").Interior.ColorIndex = 19       '�N���[���F
            
            '���͎w��
            Sheets("���ͼ��").Range("K7") = "XX �` XX�ʂ�"
            Sheets("���ͼ��").Range("K8") = "�|"
            Sheets("���ͼ��").Range("K9") = "X �ʂ��ȓ�"
            Sheets("���ͼ��").Range("K10") = "XX �` XX�ʂ�"
            Sheets("���ͼ��").Range("K11") = "XX mm�ȉ�"
            Sheets("���ͼ��").Range("K12") = "�|"
            Sheets("���ͼ��").Range("K13") = "�|"
            Sheets("���ͼ��").Range("K14") = " 0 �}  5�ʂ�"
            Sheets("���ͼ��").Range("K15") = "XX �ʂ��ȓ�"
            Sheets("���ͼ��").Range("K16") = "35 �ʂ��ȓ�"
            
        ElseIf w_�K�i = "��" Then
            With Sheets("���ͼ��").Buttons("�{�^�� 114")                    '���`�E���؋K�i �|�| ���`�E���؋K�i�� "��" ���A�{�^����\��
                                  .Visible = True
            End With
            Sheets("���ͼ��").Range("E10") = ""
        End If
    Else
        With Sheets("���ͼ��").Buttons("�{�^�� 114")
                                .Visible = False
        End With
'        Sheets("���ͼ��").Range("E10") = ""                                 '"JIS�S��" �ȊO�͎����
        Sheets("���ͼ��").Range("E10").Formula = "=IF(D10<>"""",""���}�n""&MID(D10,4,1)&""��"")"         '�ǉ� 99/ 1/14
        Sheets("���ͼ��").Range("E10").Interior.ColorIndex = 19             '�N���[���F
        Sheets("���ͼ��").Range("K7:K16").Interior.ColorIndex = 19          '�N���[���F
        
        '���͎w��
        Sheets("���ͼ��").Range("K7") = "XX �` XX�ʂ�"
        Sheets("���ͼ��").Range("K8") = "�|"
        Sheets("���ͼ��").Range("K9") = "X �ʂ��ȓ�"
        Sheets("���ͼ��").Range("K10") = "XX �` XX�ʂ�"
        Sheets("���ͼ��").Range("K11") = "XX mm�ȉ�"
        Sheets("���ͼ��").Range("K12") = "�|"
        Sheets("���ͼ��").Range("K13") = "�|"
        Sheets("���ͼ��").Range("K14") = " 0 �}  5�ʂ�"
        Sheets("���ͼ��").Range("K15") = "XX �ʂ��ȓ�"
        Sheets("���ͼ��").Range("K16") = "35 �ʂ��ȓ�"
    End If
    
    
End Sub

'*****************************************************
'
' �t�H���g�T�C�Y�̐ݒ� �}�N��
'
'*****************************************************
Sub text_harituke()
Attribute text_harituke.VB_ProcData.VB_Invoke_Func = " \n14"

    '�^��
    With Sheets("�H��}").TextBoxes("text0")
        .Font.Size = 48
    End With
    '�i��
    With Sheets("�H��}").TextBoxes("text_hinmei")
        .Font.Size = 14
    End With
    '�i��
    With Sheets("�H��}").TextBoxes("text_hinban")
        .Font.Size = 14
    End With
    '�}�^�M�����D�I�[�o�[�s���a
    With Sheets("�H��}").TextBoxes("text2")
        .Font.Size = 26
    End With
    
    '�}�^�M�����D�s���a
    With Sheets("�H��}").TextBoxes("text_kakoumen")
        .Font.Size = 7
    End With
    '�H���m��
    With Sheets("�H��}").TextBoxes("�H��_text")
        .Font.Size = 20
    End With
    
End Sub
'*****************************************************
'
'  Macro Name : hafure_get
' �i���t�������ʃ}�N���t�@�C����(EPC007.XLS)��"hafure_keisan"�}�N����苁�߂�j
'
'*****************************************************'
Sub hafure_get()
Attribute hafure_get.VB_ProcData.VB_Invoke_Func = " \n14"
  
    Dim RET, l_��, l_mod, l_P�a
        If Sheets("�\�t�����").Range("E3") = "" Or Sheets("�\�t�����").Range("A3") = "" Or Sheets("�\�t�����").Range("C3") = "" Then
            
            Sheets("���ͼ��").Range("H17") = ""
            Exit Sub
        End If
        
       l_�� = CInt(Mid(Sheets("�\�t�����").Range("E3"), 4, 1))  '���i�}�K�i���x
       l_mod = Sheets("�\�t�����").Range("A3")                  '���W���[��
       l_P�a = l_mod * Sheets("�\�t�����").Range("C3")          '�s�b�`�~�a (Ӽޭ�� * ����)
    
    RET = Application.Run(XLS_KTRTN + "!hafure_keisan", l_��, l_mod, l_P�a)
    
    Windows(XLS_FILE).Activate
    
    If RET <> 99999 Then
        
        Sheets("���ͼ��").Range("H17") = Application.RoundDown(RET, 2)
        
    Else
       ''�G���[��
       Sheets("���ͼ��").Range("F10") = Null
       MsgBox ("���U��v�Z �ɃG���[������܂����B")
    End If
End Sub
'*****************************************************
'
'  ���a���E�l
'       ���a�̖��Ռ��E�l ��ݒ肷��
'
'*****************************************************'
Sub ���a���E�l()
Attribute ���a���E�l.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim w_���, w_�M�O���a
    
    w_��� = Sheets("�\�t�����").Range("N3")
    w_���E�l = ""
    w_�M�O���a = ""
    
    w_ans = MsgBox("� ���a��� � �A� �M�O���a � ��� � �����ݸޱ��ް�̎��a ��� ���Ռ��E�l ��ݒ肵�܂��B" & _
                    Chr(13), vbOKCancel + vbExclamation, "���a���Ռ��E�l�ݒ�")
    Select Case w_ans
    Case Is = vbOK
    Case Else
        w_���E�l = Sheets("���ͼ��").Range("K18")
        GoTo last
    End Select
    
    Select Case w_���
    Case Is = "�pSP"
        w_ans = MsgBox("�A�[�o�[���̎�� �� � �pSP � �ł����H" & Chr(13) & Chr(13) & _
                       "     �� �F � ������ � ��I������Ǝ��̎�ނ͢ �� � �ƂȂ�܂��B " & Chr(13), _
                        vbYesNoCancel + vbQuestion, "���ް��� - �pSP ")
        Select Case w_ans
        Case Is = vbYes
            w_�M�O���a = Sheets("�\�t�����").Range("P3")
            
        Case Is = vbNo
            w_�M�O���a = Sheets("�\�t�����").Range("Q3")
            
        Case Else
            Exit Sub
        End Select
    
    Case Is = "INV-SP"
        w_ans = MsgBox("���̎�� �� � ����ح�Ľ��ײ� � �ł����H" & Chr(13) & Chr(13) & _
                       "     �� �F � ������ � ��I������Ǝ��̎�ނ͢ �� � �ƂȂ�܂��B " & Chr(13), _
                        vbYesNoCancel + vbQuestion, "���ް��� - INV-SP ")
        Select Case w_ans
        Case Is = vbYes             '����ح�Ľ��ײ� -- �M�O��a�Ă�
            w_�M�O���a = Sheets("�\�t�����").Range("P3")
            
        Case Is = vbNo              '�ۓ��a --  �M�OSP���a�Ă�
            w_�M�O���a = Sheets("�\�t�����").Range("Q3")
            
        Case Else
            Exit Sub
        End Select
    
    Case Is = "��", "�ԕ�", "�ޯ��", "�װ", "��", "�ض�"
        w_�M�O���a = Sheets("�\�t�����").Range("O3")        '�M�O�ۓ��a�Ă�
        
    Case Else
        a = MsgBox("� ���a��� � ���o�^����Ă��܂���B �m�F���Ă��������B", vbExclamation)
        GoTo last
    End Select
    
    If w_�M�O���a <> "" Then
        If IsNumeric(w_�M�O���a) Then
            w_���E�l = Application.Round(w_�M�O���a - 0.03, 2)
        Else
            a = MsgBox("� �M�O���a�l � �����l�łȂ��B", vbExclamation)
            GoTo last
        End If
    Else
        a = MsgBox("���a���@���o�^����Ă��܂���B", vbExclamation)
    End If
    
last:
    Sheets("���ͼ��").Range("K18") = w_���E�l
    
End Sub

'*******************************************************************************************************************************
'             �e  �X  �g  �p
'*******************************************************************************************************************************
Sub xxxx()
Attribute xxxx.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.Cursor = xlNormal
End Sub

Sub ggg()
Attribute ggg.VB_ProcData.VB_Invoke_Func = " \n14"
    With Sheets("�H��}").TextBoxes("textp1")
        .Visible = True
        .Font.Size = 20
        .Caption = "�s�b�`�덷  " & Sheets("���ͼ��").Range("H16") & "  �ȉ�"
    End With
'    With Sheets("�H��}").TextBoxes("textp1")
'        .Visible = True
'        .Font.Size = 20
'        .Caption = "���ް��݌a�����·  " & Sheets("���ͼ��").Range("H16") & "  �ȉ�"
'    End With
    With Sheets("�H��}").TextBoxes("text3")
        .Visible = True
        .Font.Size = 18
        .Caption = "(�}�^�M����  " & Sheets("���ͼ��").Range("D13") & "��)"
    End With
'     With Sheets("�H��}").TextBoxes("text3")
'         .Visible = True
'         .Font.Size = 18
'         .Caption = "(�s���a  ��" & Sheets("���ͼ��").Range("D13") & ")"
'     End With
'     With Sheets("�H��}").TextBoxes("text1")
'         .Visible = True
'         .Caption = "�I�[�o�[�s���a"
'         .Font.Size = 20
'     End With
     With Sheets("�H��}").TextBoxes("text1")
         .Visible = True
         .Caption = "�}�^�M����"
         .Font.Size = 20
     End With
    With Sheets("�H��}").TextBoxes("text2")
        .Visible = True
    End With
'    With Sheets("�H��}").TextBoxes("size1_�O")
'        .Visible = False
'        .Caption = Sheets("�����v�Z").Range("F4") & "  �`  " & Sheets("�����v�Z").Range("F5") & "  (�ʏ�)"
'        .Font.Size = 26
'    End With
'    With Sheets("�H��}").TextBoxes("size1_�O")
'        .Visible = False
'        .Caption = Sheets("�����v�Z").Range("G4") & "  �`  " & Sheets("�����v�Z").Range("G5") & "  (�ʏ�)"
'        .Font.Size = 26
'    End With
End Sub

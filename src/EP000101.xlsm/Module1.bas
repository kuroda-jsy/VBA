Attribute VB_Name = "Module1"
Function TH(�C���{�����[�g, �S�T)
Attribute TH.VB_ProcData.VB_Invoke_Func = " \n14"
    TH = 0
    P = 3.141592654 / 180
    For N = -1 To 20
        N1 = 10 ^ (-N)
        For N2 = 1 To 9
            TH = TH + N1
            INVTH = Tan(TH * P) - TH * P
            SABUN = �C���{�����[�g - INVTH
            If Abs(SABUN) < �S�T Then Exit For
            If SABUN < 0 Then TH = TH - N1
        Next N2
    Next N
End Function

Sub fff()
     Sheets("���ͼ��").DrawingObjects("size6").Visible = True
        Sheets("���ͼ��").DrawingObjects("�B���U").Visible = True
'�I�v�V�����{�^���̔�\��          Excel2007�Ή�
    ActiveSheet.Shapes("�{�^�� 312").Visible = False
    ActiveSheet.Shapes("�{�^��1").Visible = False
    ActiveSheet.Shapes("�{�^��2").Visible = False
    ActiveSheet.Shapes("�{�^��3").Visible = False
    ActiveSheet.Shapes("�{�^��4").Visible = False
    ActiveSheet.Shapes("�{�^��5").Visible = False
    ActiveSheet.Shapes("�{�^��6").Visible = False
    ActiveSheet.Shapes("�{�^��7").Visible = False
    ActiveSheet.Shapes("�{�^��8").Visible = False
    
     Sheets("���ͼ��").DrawingObjects("size6").Visible = False
        Sheets("���ͼ��").DrawingObjects("�B���U").Visible = False
'�I�v�V�����{�^���̕\��          Excel2007�Ή�
    ActiveSheet.Shapes("�{�^�� 312").Visible = True
    ActiveSheet.Shapes("�{�^��1").Visible = True
    ActiveSheet.Shapes("�{�^��2").Visible = True
    ActiveSheet.Shapes("�{�^��3").Visible = True
    ActiveSheet.Shapes("�{�^��4").Visible = True
    ActiveSheet.Shapes("�{�^��5").Visible = True
    ActiveSheet.Shapes("�{�^��6").Visible = True
    ActiveSheet.Shapes("�{�^��7").Visible = True
    ActiveSheet.Shapes("�{�^��8").Visible = True
End Sub


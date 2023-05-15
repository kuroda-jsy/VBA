Attribute VB_Name = "Module1"
Function TH(インボリュート, ゴサ)
Attribute TH.VB_ProcData.VB_Invoke_Func = " \n14"
    TH = 0
    P = 3.141592654 / 180
    For N = -1 To 20
        N1 = 10 ^ (-N)
        For N2 = 1 To 9
            TH = TH + N1
            INVTH = Tan(TH * P) - TH * P
            SABUN = インボリュート - INVTH
            If Abs(SABUN) < ゴサ Then Exit For
            If SABUN < 0 Then TH = TH - N1
        Next N2
    Next N
End Function

Sub fff()
     Sheets("入力ｼｰﾄ").DrawingObjects("size6").Visible = True
        Sheets("入力ｼｰﾄ").DrawingObjects("隠しⅡ").Visible = True
'オプションボタンの非表示          Excel2007対応
    ActiveSheet.Shapes("ボタン 312").Visible = False
    ActiveSheet.Shapes("ボタン1").Visible = False
    ActiveSheet.Shapes("ボタン2").Visible = False
    ActiveSheet.Shapes("ボタン3").Visible = False
    ActiveSheet.Shapes("ボタン4").Visible = False
    ActiveSheet.Shapes("ボタン5").Visible = False
    ActiveSheet.Shapes("ボタン6").Visible = False
    ActiveSheet.Shapes("ボタン7").Visible = False
    ActiveSheet.Shapes("ボタン8").Visible = False
    
     Sheets("入力ｼｰﾄ").DrawingObjects("size6").Visible = False
        Sheets("入力ｼｰﾄ").DrawingObjects("隠しⅡ").Visible = False
'オプションボタンの表示          Excel2007対応
    ActiveSheet.Shapes("ボタン 312").Visible = True
    ActiveSheet.Shapes("ボタン1").Visible = True
    ActiveSheet.Shapes("ボタン2").Visible = True
    ActiveSheet.Shapes("ボタン3").Visible = True
    ActiveSheet.Shapes("ボタン4").Visible = True
    ActiveSheet.Shapes("ボタン5").Visible = True
    ActiveSheet.Shapes("ボタン6").Visible = True
    ActiveSheet.Shapes("ボタン7").Visible = True
    ActiveSheet.Shapes("ボタン8").Visible = True
End Sub


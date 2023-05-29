Attribute VB_Name = "マクロ２"
Dim henkan, henkan_2, henkan_3, henkan_4, kioku_3, lank_9      '追加 99/ 1/21 K.Y
'******************************************
'
' 管理規格設定 マクロ
'
' 1998/07/22 修正 歯形歯筋規格からｾｯﾄする項目数の変更（バグ対応)
'******************************************
Sub 管理規格()
Attribute 管理規格.VB_ProcData.VB_Invoke_Func = " \n14"

    Message = Array("製品図に、下記 注記が記入されていますか？", "｢ シェービング径 ｣ がオスドックに掛かっていますか？")
    
    w_取代 = ""
    
    '歯形・歯筋設定     歯形、歯筋規格無 ＆ JIS4級の時   aaaa
    j = 0
    If Sheets("貼付けｼｰﾄ").Range("M3") = "無" And Mid(Sheets("入力ｼｰﾄ").Range("D10"), 4, 1) = "4" Then
        
        '''オスドックchk
        w_ans = MsgBox(Message(1) & Chr(13), vbYesNoCancel, "歯形・歯筋規格設定")
        Select Case w_ans
        Case Is = vbYes
              With Sheets("入力ｼｰﾄ").DrawingObjects("注記")
                .Visible = True
              End With
        
              w_ans = MsgBox(Message(0) & Chr(13), vbYesNoCancel, "歯形・歯筋規格設定")
        
              With Sheets("入力ｼｰﾄ").DrawingObjects("注記")
                .Visible = False
              End With
              Select Case w_ans
                     Case Is = vbYes
                        Sheets("入力ｼｰﾄ").Range("E10") = "ヤマハ６級"
                        For i = 0 To 9
                        If Workbooks(XLS_KTRTN).Sheets("歯形歯筋規格").Cells(i + 3, j + 4) = "" Then
                            Sheets("入力ｼｰﾄ").Cells(i + 7, j + 11) = "－"
                        Else
                            Sheets("入力ｼｰﾄ").Cells(i + 7, j + 11) = Workbooks(XLS_KTRTN).Sheets("歯形歯筋規格").Cells(i + 3, j + 4)
                        End If
                        Next i
                    Case Is = vbNo
                        Sheets("入力ｼｰﾄ").Range("E10") = "ヤマハ４級"
                        For i = 0 To 9
                        If Workbooks(XLS_KTRTN).Sheets("歯形歯筋規格").Cells(i + 3, j + 2) = "" Then
                            Sheets("入力ｼｰﾄ").Cells(i + 7, j + 11) = "－"
                        Else
                            Sheets("入力ｼｰﾄ").Cells(i + 7, j + 11) = Workbooks(XLS_KTRTN).Sheets("歯形歯筋規格").Cells(i + 3, j + 2)
                        End If
                        Next i
                        a = MsgBox("｢ ＶＡ提案 ｣ を提出して下さい", vbExclamation)
              End Select
        
        
        Case Is = vbNo
            Sheets("入力ｼｰﾄ").Range("E10") = "ヤマハ４級"
'            For i = 0 To 8  1998/07/22 修正
            For i = 0 To 9
                If Workbooks(XLS_KTRTN).Sheets("歯形歯筋規格").Cells(i + 3, j + 2) = "" Then
                    Sheets("入力ｼｰﾄ").Cells(i + 7, j + 11) = "－"
                Else
                    Sheets("入力ｼｰﾄ").Cells(i + 7, j + 11) = Workbooks(XLS_KTRTN).Sheets("歯形歯筋規格").Cells(i + 3, j + 2)
                End If
            Next i
            Exit Sub
        End Select
    End If
End Sub
'******************************************
'
' 公差設定 マクロ
'
'******************************************
Sub 公差表示()
Attribute 公差表示.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim w_kousa1, w_kousa2
    
    w_kousa1 = Worksheets("入力ｼｰﾄ").Range("D8")
    w_kousa2 = Worksheets("入力ｼｰﾄ").Range("D9")
    
    If (Not IsNumeric(w_kousa1)) Or (Not IsNumeric(w_kousa2)) Then
    
        If w_kousa1 = "*****" And w_kousa2 = "*****" Then
        
            Sheets("工作図").TextBoxes("kousa0").Visible = False
            Sheets("工作図").TextBoxes("kousa1").Visible = False
            Sheets("工作図").TextBoxes("kousa2").Visible = False
        Else
            a = MsgBox("公差には、数値を設定してください。", vbExclamation)
        End If
        
        Exit Sub
        
    End If
    
    If Abs(w_kousa1) = Abs(w_kousa2) Then                       '公差の絶対値が同じ時
        With Worksheets("工作図").TextBoxes("kousa0")     '"±" 表示
            .Caption = "±" + CStr(Abs(w_kousa1))
            .Font.Size = 16
            .Visible = True
        End With
        
        With Worksheets("工作図").TextBoxes("kousa1")     '"上限" 非表示
            .Visible = False
        End With
        
        With Worksheets("工作図").TextBoxes("kousa2")    '"下限" 非表示
            .Visible = False
        End With
    Else
        With Worksheets("工作図").TextBoxes("kousa0")    '"±" 非表示
            .Visible = False
        End With
        '**********************************
        RET = Application.Run(XLS_KTRTN + "!公差_編集", w_kousa1, w_kousa2)
        If RET = "err" Then
           Exit Sub
        Else
           w_kousa1 = Mid(RET, 1, Application.Search("#", RET, 1) - 1)
           w_kousa2 = Mid(RET, Application.Search("#", RET, 1) + 1, 6)
        End If
        '**********************************
                                        
        '"上限" 表示
        With Worksheets("工作図").TextBoxes("kousa1")
            .Font.Name = "ＭＳ ゴシック"
            .Caption = w_kousa1
            .Font.Size = 11
            .Visible = True
        End With
        
        '"下限"表示
        With Worksheets("工作図").TextBoxes("kousa2")
            .Font.Name = "ＭＳ ゴシック"
            .Caption = w_kousa2
            .Font.Size = 11
            .Visible = True
        End With
    End If
    
End Sub
'******************************************
'
' サイズ品指定 マクロ
'
'******************************************
Sub サイズ指定()
Attribute サイズ指定.VB_ProcData.VB_Invoke_Func = " \n14"

    Sheets("入力ｼｰﾄ").Unprotect
    Sheets("入力ｼｰﾄ").Range("C23:C30") = "－"
    w_ans = MsgBox("このワーク は、｢ サイズ品 ｣ ですか？" & Chr(13), vbYesNoCancel, "ｻｲｽﾞ品 設定")
    
    Select Case w_ans
    Case Is = vbYes
        
        Sheets("工作図").TextBoxes("size2").Visible = True
'        Sheets("工作図").TextBoxes("size4").Visible = True
        
        Sheets("入力ｼｰﾄ").TextBoxes("size5").Visible = True
        
        If Worksheets("貼付けｼｰﾄ").Range("H3") = "M" Then
           サイズ選択
        Else
          If Worksheets("貼付けｼｰﾄ").Range("H3") = "O" Then
msg_4:      Message_4 = "サイズ表のパターンはどうしますか？" & Chr(13) & "下の番号から選択して下さい。" & Chr(13) & Chr(13) & _
                        "１．昇順      ２．降順"
            Title = "サイズ表パターン選択"
            w_Default = "1"
            MyValue_4 = InputBox(Message_4, Title, w_Default)
            henkan_4 = CStr(MyValue_4)
            Select Case henkan_4
              Case "1"
                表を書く_1
              Case "2"
                表を書く_1
              Case Else
                If henkan_4 <> "" Then
                  c_4 = MsgBox("範囲外の数値です。もう一度入力して下さい。", vbQuestion, "注意")
                  GoTo msg_4
                Else
                  Exit Sub
                End If
            End Select
          End If
        End If
        Sheets("工作図").Range("AW16").Value = "1/50"    '2000/1/25 ADD
        Sheets("工作図").Range("AW18").Value = "1/50"    '2000/1/25 ADD
    Case vbNo
        
        サイズ品でないとき_2              'サイズ品時の設定を解除する
        Sheets("工作図").Range("AW16").Value = "1/2h"    '2000/1/25 ADD
        Sheets("工作図").Range("AW18").Value = "1/2h"    '2000/1/25 ADD
        
    End Select
    'test
End Sub
        
'******************************************
'
'  マタギ歯厚選択時のマクロ
'
'******************************************
Sub サイズ選択()
Attribute サイズ選択.VB_ProcData.VB_Invoke_Func = " \n14"

     Sheets("入力ｼｰﾄ").DrawingObjects("size6").Visible = True
        Sheets("入力ｼｰﾄ").DrawingObjects("隠しⅡ").Visible = True
'オプションボタンの非表示          Excel2007対応
    Sheets("入力ｼｰﾄ").Shapes("ボタン 312").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン1").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン2").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン3").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン4").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン5").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン6").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン7").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン8").Visible = False
       
msg_1:  Message = "どのような表示方法にしますか？" & Chr(13) & "下の番号から選択して下さい。" & Chr(13) & Chr(13) & _
                  "１．マタギ歯厚のみ   ２．両方"
        Title = "サイズ品選択"
        w_Default = "1"
        MyValue = InputBox(Message, Title, w_Default)
        henkan = CStr(MyValue)
        Select Case henkan
          Case "1"
msg_2:      Message_2 = "サイズ表のパターンはどうしますか？" & Chr(13) & "下の番号から選択して下さい。" & Chr(13) & Chr(13) & _
                        "１．昇順      ２．降順"
            Title = "サイズ表パターン選択"
            w_Default = "1"
            MyValue_2 = InputBox(Message_2, Title, w_Default)
            henkan_2 = CStr(MyValue_2)
            Select Case henkan_2
              Case "1"
                Sheets("入力ｼｰﾄ").Range("C7").Value = "マタギ歯厚"
                Sheets("入力ｼｰﾄ").Range("C7").Font.Size = 11
                表を書く_1
              Case "2"
                Sheets("入力ｼｰﾄ").Range("C7").Value = "マタギ歯厚"
                Sheets("入力ｼｰﾄ").Range("C7").Font.Size = 11
                表を書く_1
              Case Else
                If henkan_2 <> "" Then
                  c_2 = MsgBox("範囲外の数値です。もう一度入力して下さい。", vbQuestion, "注意")
                  GoTo msg_2
                Else
                  Exit Sub
                End If
            End Select
          Case "2"
msg_3:      Message_3 = "サイズ表のパターンはどうしますか？" & Chr(13) & "下の番号から選択して下さい。" & Chr(13) & Chr(13) & _
                        "１．昇順      ２．降順"
            Title = "サイズ表パターン選択"
            w_Default = "1"
            MyValue_3 = InputBox(Message_3, Title, w_Default)
            henkan_3 = CStr(MyValue_3)
            Select Case henkan_3
              Case "1"
                kioku_3 = henkan_3
                Sheets("歯厚計算").Select
                Range("G3").Select
                a_res = MsgBox("黄色のセルに数値を入力して下さい。入力後、[入力シートへ戻る]ボタンを押して下さい。", vbInformation, "確認")
              Case "2"
                kioku_3 = henkan_3
                Sheets("歯厚計算２").Select
                Range("G3").Select
                a_res = MsgBox("黄色のセルに数値を入力して下さい。入力後、[入力シートへ戻る]ボタンを押して下さい。", vbInformation, "確認")
              Case Else
                If henkan_3 <> "" Then
                  c_3 = MsgBox("範囲外の数値です。もう一度入力して下さい。", vbQuestion, "注意")
                  GoTo msg_3
                Else
                  Exit Sub
                End If
            End Select
          Case Else
            If henkan <> "" Then
               c = MsgBox("範囲外の数値です。もう一度入力して下さい。", vbQuestion, "注意")
               GoTo msg_1
            Else
               Exit Sub
            End If
        
        End Select
        
End Sub

'******************************************
'
'  サイズ品でない時と入力値クリア－時のマクロ
'
'******************************************
Sub サイズ品でないとき_1()
Attribute サイズ品でないとき_1.VB_ProcData.VB_Invoke_Func = " \n14"
    
    With Sheets("工作図").TextBoxes("text2")
        .Visible = True
    End With
    With Sheets("工作図").TextBoxes("text3")
        .Visible = True
    End With
    With Sheets("工作図").TextBoxes("size1_前")
        .Visible = False
    End With
    
    Sheets("工作図").DrawingObjects("隠しＬ").Visible = True
    Sheets("工作図").DrawingObjects("隠しＲ").Visible = True
    
    Sheets("工作図").TextBoxes("size2").Visible = False
    Sheets("工作図").TextBoxes("size4").Visible = False
    Sheets("工作図").TextBoxes("size4の左").Visible = False
        
    Sheets("入力ｼｰﾄ").TextBoxes("size5").Visible = False
    Sheets("入力ｼｰﾄ").DrawingObjects("隠しⅡ").Visible = True
    With Sheets("入力ｼｰﾄ").DrawingObjects("size6")
        .Visible = True
        .BringToFront
    End With
'オプションボタンの非表示          Excel2007対応
    Sheets("入力ｼｰﾄ").Shapes("ボタン 312").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン1").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン2").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン3").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン4").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン5").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン6").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン7").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン8").Visible = False
    
    If Sheets("貼付けｼｰﾄ").Range("H3") = "O" Then               'オーバーピンの場合         '追加 99/ 1/18
       Sheets("入力ｼｰﾄ").Range("C7").Value = "オーバーピン径"
       Sheets("入力ｼｰﾄ").Range("C7").Font.Size = 10
       Sheets("入力ｼｰﾄ").Range("C15").Value = "ピン径"
       Sheets("入力ｼｰﾄ").Range("C15").Font.Size = 11
       Sheets("入力ｼｰﾄ").Range("G16").Value = "ｵｰﾊﾞｰﾋﾟﾝ径のﾊﾞﾗﾂｷ"
       Sheets("入力ｼｰﾄ").Range("G16").Font.Size = 10
       Sheets("入力ｼｰﾄ").Range("H16").Value = 0.05
       Sheets("入力ｼｰﾄ").Range("H16").Font.Size = 11
'       Sheets("工作図").Range("AG44").Value = "歯厚マイクロ"                                '追加 99/ 1/18
'       Sheets("工作図").Range("AG45").Value = "ｵｰﾊﾞｰﾋﾟﾝﾏｲｸﾛ"                               '追加 99/ 1/18
    ElseIf Sheets("貼付けｼｰﾄ").Range("H3") = "M" Then           'マタギの場合               '追加 99/ 1/18
       Sheets("入力ｼｰﾄ").Range("C7").Value = "マタギ歯厚"
       Sheets("入力ｼｰﾄ").Range("C7").Font.Size = 11
       Sheets("入力ｼｰﾄ").Range("C15").Value = "マタギ枚数"
       Sheets("入力ｼｰﾄ").Range("C15").Font.Size = 11
       Sheets("入力ｼｰﾄ").Range("G16").Value = "ピッチ誤差"
       Sheets("入力ｼｰﾄ").Range("G16").Font.Size = 11
       Sheets("入力ｼｰﾄ").Range("H16").Value = 0.02
       Sheets("入力ｼｰﾄ").Range("H16").Font.Size = 11
'       Sheets("工作図").Range("AG44").Value = "歯厚マイクロ"                                '追加 99/ 1/18
'       Sheets("工作図").Range("AG45").Value = "歯厚マイクロ"                                '追加 99/ 1/18
    End If                                                                                 '追加 99/ 1/18

    With Sheets("工作図").TextBoxes("text1")
        .Visible = True
        .Caption = Sheets("入力ｼｰﾄ").Range("C7")
    End With
    With Sheets("工作図").TextBoxes("size2")           '追加 99/ 1/18
        .Visible = False                               '追加 99/ 1/18
        .Caption = "歯厚分類表による"                   '追加 99/ 1/18
        .Font.Size = 6                                 '追加 99/ 1/18
    End With                                           '追加 99/ 1/18
    With Sheets("工作図").TextBoxes("size4の左")        '追加 99/ 1/18
        .Visible = False                               '追加 99/ 1/18
        .Caption = "歯厚分類表による"                   '追加 99/ 1/18
        .Font.Size = 12                                '追加 99/ 1/18
    End With                                           '追加 99/ 1/18
    With Sheets("工作図").TextBoxes("size4")           '追加 99/ 1/18
        .Visible = False                               '追加 99/ 1/18
        .Caption = "←"                                '追加 99/ 1/18
        .Font.Size = 16                                '追加 99/ 1/18
    End With                                           '追加 99/ 1/18
        
End Sub

'******************************************
'
'  サイズ品でない時と入力値クリア－時のマクロ2
'
'******************************************
Sub サイズ品でないとき_2()
Attribute サイズ品でないとき_2.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Application.Run (XLS_KTRTN + "!工作図アンロック")
    With Sheets("工作図").TextBoxes("text2")
        .Visible = True
    End With
    With Sheets("工作図").TextBoxes("size1_前")
        .Visible = False
    End With
    
    Sheets("工作図").DrawingObjects("隠しＬ").Visible = True
    Sheets("工作図").DrawingObjects("隠しＲ").Visible = True
    
    Sheets("工作図").TextBoxes("size2").Visible = False
    Sheets("工作図").TextBoxes("size4").Visible = False
    Sheets("工作図").TextBoxes("size4の左").Visible = False
        
    Sheets("入力ｼｰﾄ").TextBoxes("size5").Visible = False
    Sheets("入力ｼｰﾄ").DrawingObjects("隠しⅡ").Visible = True
    With Sheets("入力ｼｰﾄ").DrawingObjects("size6")
        .Visible = True
        .BringToFront
    End With
'オプションボタンの非表示          Excel2007対応
    Sheets("入力ｼｰﾄ").Shapes("ボタン 312").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン1").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン2").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン3").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン4").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン5").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン6").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン7").Visible = False
    Sheets("入力ｼｰﾄ").Shapes("ボタン8").Visible = False
    
    If Sheets("貼付けｼｰﾄ").Range("H3") = "O" Then               'オーバーピンの場合         '追加 99/ 1/18
       Sheets("入力ｼｰﾄ").Range("C7").Value = "オーバーピン径"
       Sheets("入力ｼｰﾄ").Range("C7").Font.Size = 10
       Sheets("入力ｼｰﾄ").Range("C15").Value = "ピン径"
       Sheets("入力ｼｰﾄ").Range("C15").Font.Size = 11
       Sheets("入力ｼｰﾄ").Range("G16").Value = "ｵｰﾊﾞｰﾋﾟﾝ径のﾊﾞﾗﾂｷ"
       Sheets("入力ｼｰﾄ").Range("G16").Font.Size = 10
       Sheets("入力ｼｰﾄ").Range("H16").Value = 0.05
       Sheets("入力ｼｰﾄ").Range("H16").Font.Size = 11
'       Sheets("工作図").Range("AG44").Value = "歯厚マイクロ"                                '追加 99/ 1/18
'       Sheets("工作図").Range("AG45").Value = "ｵｰﾊﾞｰﾋﾟﾝﾏｲｸﾛ"                               '追加 99/ 1/18
       With Sheets("工作図").TextBoxes("textp1")                                           '追加 99/ 1/14
           .Visible = True                                                                 '追加 99/ 1/14
           .Font.Size = 20                                                                 '追加 99/ 1/14
           .Caption = "ｵｰﾊﾞｰﾋﾟﾝ径のﾊﾞﾗﾂｷ  " & Sheets("入力ｼｰﾄ").Range("H16") & "  以下"     '追加 99/ 1/14
       End With                                                                            '追加 99/ 1/14
       With Sheets("工作図").TextBoxes("text3")                                           '追加 99/ 1/14
           .Visible = True                                                                '追加 99/ 1/14
           .Font.Size = 18                                                                '追加 99/ 1/14
           .Caption = "(ピン径  " & Sheets("入力ｼｰﾄ").Range("D15") & ")"                   '追加 99/ 1/14
       End With                                                                           '追加 99/ 1/14

    ElseIf Sheets("貼付けｼｰﾄ").Range("H3") = "M" Then           'マタギの場合               '追加 99/ 1/18
       Sheets("入力ｼｰﾄ").Range("C7").Value = "マタギ歯厚"
       Sheets("入力ｼｰﾄ").Range("C7").Font.Size = 11
       Sheets("入力ｼｰﾄ").Range("C15").Value = "マタギ枚数"
       Sheets("入力ｼｰﾄ").Range("C15").Font.Size = 11
       Sheets("入力ｼｰﾄ").Range("G16").Value = "ピッチ誤差"
       Sheets("入力ｼｰﾄ").Range("G16").Font.Size = 11
       Sheets("入力ｼｰﾄ").Range("H16").Value = 0.02
       Sheets("入力ｼｰﾄ").Range("H16").Font.Size = 11
'       Sheets("工作図").Range("AG44").Value = "歯厚マイクロ"                                '追加 99/ 1/18
'       Sheets("工作図").Range("AG45").Value = "歯厚マイクロ"                                '追加 99/ 1/18
       With Sheets("工作図").TextBoxes("textp1")                                           '追加 99/ 1/14
           .Visible = True                                                                 '追加 99/ 1/14
           .Font.Size = 20                                                                 '追加 99/ 1/14
           .Caption = "ピッチ誤差  " & Sheets("入力ｼｰﾄ").Range("H16") & "  以下"             '追加 99/ 1/14
       End With                                                                             '追加 99/ 1/14
       With Sheets("工作図").TextBoxes("text3")                                             '追加 99/ 1/14
           .Visible = True                                                                  '追加 99/ 1/14
           .Font.Size = 18                                                                  '追加 99/ 1/14
           .Caption = "(マタギ枚数  " & Sheets("入力ｼｰﾄ").Range("D15") & "枚)"               '追加 99/ 1/14
       End With                                                                             '追加 99/ 1/14
       With Sheets("工作図").TextBoxes("micro_text1")                                       '追加 99/ 1/18
           .Caption = "歯厚ﾏｲｸﾛ"                                                            '追加 99/ 1/18
           .Font.Size = 9                                                                   '追加 99/ 1/18
           .Visible = True                                                                  '追加 99/ 1/18
       End With                                                                             '追加 99/ 1/18
    End If                                                                                  '追加 99/ 1/18

    With Sheets("工作図").TextBoxes("text1")
        .Visible = True
        .Caption = Sheets("入力ｼｰﾄ").Range("C7")
    End With
    With Sheets("工作図").TextBoxes("size2")           '追加 99/ 1/18
        .Visible = False                               '追加 99/ 1/18
        .Caption = "歯厚分類表による"                   '追加 99/ 1/18
        .Font.Size = 6                                 '追加 99/ 1/18
    End With                                           '追加 99/ 1/18
    With Sheets("工作図").TextBoxes("size4の左")        '追加 99/ 1/18
        .Visible = False                               '追加 99/ 1/18
        .Caption = "歯厚分類表による"                   '追加 99/ 1/18
        .Font.Size = 12                                '追加 99/ 1/18
    End With                                           '追加 99/ 1/18
    With Sheets("工作図").TextBoxes("size4")           '追加 99/ 1/18
        .Visible = False                               '追加 99/ 1/18
        .Caption = "←"                                '追加 99/ 1/18
        .Font.Size = 16                                '追加 99/ 1/18
    End With                                           '追加 99/ 1/18
        
    Application.Run (XLS_KTRTN + "!工作図ロック")
End Sub
'****************************************
'
'  歯厚計算後入力シ－トへ戻る時のマクロ
'
'****************************************
Sub 入力シートへ戻る()
Attribute 入力シートへ戻る.VB_ProcData.VB_Invoke_Func = " \n14"
    表を書く_2
End Sub

'****************************************
'
'  数値を表に書くマクロ その１
'
'****************************************
Sub 表を書く_1()
Attribute 表を書く_1.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("入力ｼｰﾄ").Unprotect
    Application.Run (XLS_KTRTN + "!工作図アンロック")
    With Sheets("工作図").TextBoxes("size1_前")
        .Visible = False
    End With
    If Worksheets("入力ｼｰﾄ").Range("C7") = "マタギ歯厚" Then
      If henkan_2 = "1" Then                                    'マタギ歯厚  昇順の時
        For i = 1 To 8
          Sheets("入力ｼｰﾄ").Range("D" & Trim(i + 22)) = Sheets("歯厚計算").Cells(i + 3, 33)
        Next i
        For i = 1 To 8
          Sheets("工作図").TextBoxes("左左" & Trim(i)).Caption = Sheets("入力ｼｰﾄ").Range("D" & Trim(i + 22))
        Next i
      Else
        If henkan_2 = "2" Then                                  'マタギ歯厚  降順の時
          For i = 4 To 11
            Sheets("入力ｼｰﾄ").Range("D" & Trim(i + 19)) = Sheets("歯厚計算２").Cells(i, 33)
          Next i
          For i = 1 To 8
            Sheets("工作図").TextBoxes("左左" & Trim(i)).Caption = Sheets("入力ｼｰﾄ").Range("D" & Trim(i + 22))
          Next i
        End If
      End If
       
      Sheets("工作図").TextBoxes("SUUTI_1").Caption = "( " & Sheets("入力ｼｰﾄ").Range("D15") & "枚" & " )"
       
'++++++++ コマンドボタンの表示、非表示を切り替える 2012/01/01 +++++++++
      For i = 1 To 8
        With Sheets("入力ｼｰﾄ").OptionButtons("ボタン" & Trim(Str(i)))
            .Value = xlOff
            .Visible = True
        End With
      Next i
       
      Sheets("入力ｼｰﾄ").Range("C15").Value = "マタギ枚数"
      Sheets("入力ｼｰﾄ").Range("C15").Font.Size = 11
      Sheets("入力ｼｰﾄ").Range("G16").Value = "ピッチ誤差"
      Sheets("入力ｼｰﾄ").Range("G16").Font.Size = 11
      Sheets("入力ｼｰﾄ").Range("H16").Value = 0.02
      Sheets("入力ｼｰﾄ").Range("H16").Font.Size = 11

      With Sheets("工作図").TextBoxes("textp1")                                    '追加 99/ 1/14
          .Visible = True                                                          '追加 99/ 1/14
          .Font.Size = 20                                                          '追加 99/ 1/14
          .Caption = "ピッチ誤差  " & Sheets("入力ｼｰﾄ").Range("H16") & "  以下"     '追加 99/ 1/14
      End With                                                                     '追加 99/ 1/14
      With Sheets("工作図").TextBoxes("text1")                                      '追加 99/ 1/14
          .Visible = True                                                           '追加 99/ 1/14
          .Caption = "マタギ歯厚"                                                    '追加 99/ 1/14
          .Font.Size = 20                                                           '追加 99/ 1/14
      End With                                                                      '追加 99/ 1/14
      With Sheets("工作図").TextBoxes("text2")                                      '追加 99/ 1/14
          .Visible = True                                                           '追加 99/ 1/14
      End With                                                                      '追加 99/ 1/14
      With Sheets("工作図").TextBoxes("text3")                                      '追加 99/ 1/14
          .Visible = True                                                           '追加 99/ 1/14
          .Font.Size = 18                                                           '追加 99/ 1/14
          .Caption = "(マタギ枚数  " & Sheets("入力ｼｰﾄ").Range("D15") & "枚)"        '追加 99/ 1/14
      End With                                                                      '追加 99/ 1/14

'      Sheets("工作図").Range("AG44").Value = "歯厚マイクロ"                          '追加 99/ 1/18
'      Sheets("工作図").Range("AG45").Value = "歯厚マイクロ"                          '追加 99/ 1/18
      Sheets("工作図").Range("AG45").Font.Size = 11
      With Sheets("工作図").TextBoxes("micro_text11")
          .Visible = False
      End With
      With Sheets("工作図").TextBoxes("micro_text12")
          .Visible = False
      End With
      With Sheets("工作図").TextBoxes("micro_text1")                                '追加 99/ 1/18
          .Caption = "歯厚ﾏｲｸﾛ"                                                     '追加 99/ 1/18
          .Font.Size = 9                                                            '追加 99/ 1/18
          .Visible = True                                                           '追加 99/ 1/18
      End With                                                                      '追加 99/ 1/18
      Sheets("入力ｼｰﾄ").DrawingObjects("size6").Visible = False
    'オプションボタンの表示          Excel2007対応
      Sheets("入力ｼｰﾄ").Shapes("ボタン 312").Visible = True
      Sheets("入力ｼｰﾄ").Shapes("ボタン1").Visible = True
      Sheets("入力ｼｰﾄ").Shapes("ボタン2").Visible = True
      Sheets("入力ｼｰﾄ").Shapes("ボタン3").Visible = True
      Sheets("入力ｼｰﾄ").Shapes("ボタン4").Visible = True
      Sheets("入力ｼｰﾄ").Shapes("ボタン5").Visible = True
      Sheets("入力ｼｰﾄ").Shapes("ボタン6").Visible = True
      Sheets("入力ｼｰﾄ").Shapes("ボタン7").Visible = True
      Sheets("入力ｼｰﾄ").Shapes("ボタン8").Visible = True
      Sheets("入力ｼｰﾄ").DrawingObjects("隠しⅡ").Visible = True
      Worksheets("入力ｼｰﾄ").Range("D22") = "マタギ歯厚"
      Sheets("工作図").DrawingObjects("隠しＬ").Visible = False
      Sheets("工作図").DrawingObjects("隠しＲ").Visible = True
      Sheets("入力ｼｰﾄ").Select
      Range("C23").Select
      a_res = MsgBox("サイズを入力して下さい。入力後[確認]ボタンを押して下さい。", vbExclamation, "確認")
      
      With Sheets("工作図").DrawingObjects("表隠し")
          .Visible = True
          .Top = 386.5
          .Height = 177
          .Left = 6
          .Width = 287.25
      End With
      
    Else
      If Worksheets("入力ｼｰﾄ").Range("C7") = "オーバーピン径" Then
msg_9:   Message_9 = "1ランクの寸法は？" & Chr(13) & "下に入力して下さい。"
         Title = "1ランクの寸法入力"
         w_Default = "0.02"
         MyValue_9 = InputBox(Message_9, Title, w_Default)
         lank_9 = CStr(MyValue_9)
         If lank_9 <> "" Then
         Else
           Exit Sub
         End If
        If henkan_4 = "1" Then                                             'オーバーピン径  昇順の時
          w_kanrikikaku_1 = Sheets("入力ｼｰﾄ").Range("D7") - Sheets("入力ｼｰﾄ").Range("D13") + Sheets("入力ｼｰﾄ").Range("D9")
          w_kanrikikaku_2 = w_kanrikikaku_1 + lank_9
          For i = 1 To 8
            Sheets("入力ｼｰﾄ").Range("X" & Trim(i + 6)) = w_kanrikikaku_1
            Sheets("入力ｼｰﾄ").Range("Y" & Trim(i + 6)) = w_kanrikikaku_2
            w_kanrikikaku_1 = w_kanrikikaku_1 + lank_9
            w_kanrikikaku_2 = w_kanrikikaku_2 + lank_9
          Next i
         
          For i = 1 To 8
            Sheets("入力ｼｰﾄ").Range("D" & Trim(i + 22)) = Sheets("入力ｼｰﾄ").Range("Z" & Trim(i + 6))
            Sheets("工作図").TextBoxes("左左" & Trim(i)).Caption = Sheets("入力ｼｰﾄ").Range("D" & Trim(i + 22))
          Next i
        Else
          If henkan_4 = "2" Then                                           'オーバーピン径  降順の時
            w_kanrikikaku_1 = Sheets("入力ｼｰﾄ").Range("D7") - Sheets("入力ｼｰﾄ").Range("D13") + Sheets("入力ｼｰﾄ").Range("D9")
            w_kanrikikaku_2 = w_kanrikikaku_1 + lank_9
            For i = 1 To 8
              Sheets("入力ｼｰﾄ").Range("X" & Trim(i + 6)) = w_kanrikikaku_1
              Sheets("入力ｼｰﾄ").Range("Y" & Trim(i + 6)) = w_kanrikikaku_2
              w_kanrikikaku_1 = w_kanrikikaku_1 + lank_9
              w_kanrikikaku_2 = w_kanrikikaku_2 + lank_9
            Next i
            j = 0
            For i = 7 To 14
              Sheets("入力ｼｰﾄ").Range("D" & Trim(i + 23 - j)) = Sheets("入力ｼｰﾄ").Range("Z" & Trim(i))
              j = j + 2
            Next i
            For i = 1 To 8
              Sheets("工作図").TextBoxes("左左" & Trim(i)).Caption = Sheets("入力ｼｰﾄ").Range("D" & Trim(i + 22))
            Next i
          End If
        End If
         
        Sheets("工作図").TextBoxes("SUUTI_1").Caption = "(" & Sheets("入力ｼｰﾄ").Range("D15") & ")"
         
'++++++++ コマンドボタンの表示、非表示を切り替える 2012/01/01 +++++++++
        For i = 1 To 8
          With Sheets("入力ｼｰﾄ").OptionButtons("ボタン" & Trim(Str(i)))
              .Value = xlOff
              .Visible = True
          End With
        Next i
         
        With Sheets("工作図").TextBoxes("text3")                                           '追加 99/ 1/14
            .Visible = True                                                                '追加 99/ 1/14
            .Font.Size = 18                                                                '追加 99/ 1/14
            .Caption = "(ピン径  " & Sheets("入力ｼｰﾄ").Range("D15") & ")"                   '追加 99/ 1/14
        End With                                                                           '追加 99/ 1/14
        With Sheets("工作図").TextBoxes("text1")                                           '追加 99/ 1/14
            .Visible = True                                                                '追加 99/ 1/14
            .Caption = "オーバーピン径"                                                     '追加 99/ 1/14
            .Font.Size = 20                                                                '追加 99/ 1/14
        End With                                                                           '追加 99/ 1/14
        With Sheets("工作図").TextBoxes("text2")                                            '追加 99/ 1/14
            .Visible = True                                                                 '追加 99/ 1/14
        End With                                                                            '追加 99/ 1/14
        With Sheets("工作図").TextBoxes("textp1")                                           '追加 99/ 1/14
            .Visible = True                                                                 '追加 99/ 1/14
            .Font.Size = 20                                                                 '追加 99/ 1/14
            .Caption = "ｵｰﾊﾞｰﾋﾟﾝ径のﾊﾞﾗﾂｷ  " & Sheets("入力ｼｰﾄ").Range("H16") & "  以下"     '追加 99/ 1/14
        End With                                                                            '追加 99/ 1/14
'        Sheets("工作図").Range("AG44").Value = "歯厚マイクロ"                                '追加 99/ 1/18
'        Sheets("工作図").Range("AG45").Value = "ｵｰﾊﾞｰﾋﾟﾝﾏｲｸﾛ"                               '追加 99/ 1/18
        Sheets("入力ｼｰﾄ").DrawingObjects("size6").Visible = False
    'オプションボタンの表示          Excel2007対応
        Sheets("入力ｼｰﾄ").Shapes("ボタン 312").Visible = True
        Sheets("入力ｼｰﾄ").Shapes("ボタン1").Visible = True
        Sheets("入力ｼｰﾄ").Shapes("ボタン2").Visible = True
        Sheets("入力ｼｰﾄ").Shapes("ボタン3").Visible = True
        Sheets("入力ｼｰﾄ").Shapes("ボタン4").Visible = True
        Sheets("入力ｼｰﾄ").Shapes("ボタン5").Visible = True
        Sheets("入力ｼｰﾄ").Shapes("ボタン6").Visible = True
        Sheets("入力ｼｰﾄ").Shapes("ボタン7").Visible = True
        Sheets("入力ｼｰﾄ").Shapes("ボタン8").Visible = True
        
        Sheets("入力ｼｰﾄ").DrawingObjects("隠しⅡ").Visible = True
        Worksheets("入力ｼｰﾄ").Range("D22") = "オーバーピン径"
        Sheets("工作図").DrawingObjects("隠しＬ").Visible = False
        Sheets("工作図").DrawingObjects("隠しＲ").Visible = True
        Sheets("入力ｼｰﾄ").Select
        Range("C23").Select
        a_res = MsgBox("サイズを入力して下さい。入力後[確認]ボタンを押して下さい。", vbExclamation, "確認")
        
        With Sheets("工作図").DrawingObjects("表隠し")
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
'  数値を表に書くマクロ その２
'
'****************************************
Sub 表を書く_2()
Attribute 表を書く_2.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("入力ｼｰﾄ").Unprotect
    Application.Run (XLS_KTRTN + "!工作図アンロック")
    With Sheets("工作図").TextBoxes("size1_前")
        .Visible = False
    End With
    If kioku_3 = "1" Then                           'オーバーピン径  昇順の時
      For i = 4 To 11
        Sheets("入力ｼｰﾄ").Range("D" & Trim(i + 19)) = Sheets("歯厚計算").Cells(i, 8)
        Sheets("入力ｼｰﾄ").Range("G" & Trim(i + 19)) = Sheets("歯厚計算").Cells(i, 33)
      Next i
    
      For i = 1 To 8
        Sheets("工作図").TextBoxes("左左" & Trim(i)).Caption = Sheets("入力ｼｰﾄ").Range("D" & Trim(i + 22))
        Sheets("工作図").TextBoxes("右左" & Trim(i)).Caption = Sheets("入力ｼｰﾄ").Range("G" & Trim(i + 22))
      Next i
    Else
      If kioku_3 = "2" Then                         'オーバーピン径  降順の時
        For i = 4 To 11
          Sheets("入力ｼｰﾄ").Range("D" & Trim(i + 19)) = Sheets("歯厚計算２").Cells(i, 8)
          Sheets("入力ｼｰﾄ").Range("G" & Trim(i + 19)) = Sheets("歯厚計算２").Cells(i, 33)
        Next i
    
        For i = 1 To 8
          Sheets("工作図").TextBoxes("左左" & Trim(i)).Caption = Sheets("入力ｼｰﾄ").Range("D" & Trim(i + 22))
          Sheets("工作図").TextBoxes("右左" & Trim(i)).Caption = Sheets("入力ｼｰﾄ").Range("G" & Trim(i + 22))
        Next i
      End If
    End If
        
    Sheets("工作図").TextBoxes("SUUTI_1").Caption = "( " & "φ" & Sheets("歯厚計算").Range("G3") & " )"
    Sheets("工作図").TextBoxes("SUUTI_2").Caption = "( " & Sheets("歯厚計算").Range("E4") & "枚" & " )"
    
'++++++++ コマンドボタンの表示、非表示を切り替える 2012/01/01 +++++++++
    For i = 1 To 8
      With Sheets("入力ｼｰﾄ").OptionButtons("ボタン" & Trim(Str(i)))
          .Value = xlOff
          .Visible = True
      End With
    Next i
    
    Sheets("入力ｼｰﾄ").Range("C7").Value = "オーバーピン径"
    Sheets("入力ｼｰﾄ").Range("C7").Font.Size = 10
    Sheets("入力ｼｰﾄ").Range("G16").Value = "ｵｰﾊﾞｰﾋﾟﾝ径のﾊﾞﾗﾂｷ"
    Sheets("入力ｼｰﾄ").Range("G16").Font.Size = 10
    Sheets("入力ｼｰﾄ").Range("H16").Value = 0.05
    Sheets("入力ｼｰﾄ").Range("H16").Font.Size = 11
'    Sheets("工作図").Range("AG45").Value = "ｵｰﾊﾞｰﾋﾟﾝﾏｲｸﾛ"
'    Sheets("工作図").Range("AG45").Font.Size = 11

    With Sheets("工作図").TextBoxes("textp1")                                            '追加 99/ 1/14
        .Visible = True                                                                 '追加 99/ 1/14
        .Font.Size = 20                                                                 '追加 99/ 1/14
        .Caption = "ｵｰﾊﾞｰﾋﾟﾝ径のﾊﾞﾗﾂｷ  " & Sheets("入力ｼｰﾄ").Range("H16") & "  以下"     '追加 99/ 1/14
    End With                                                                            '追加 99/ 1/14
    With Sheets("工作図").TextBoxes("text3")                                             '追加 99/ 1/14
         .Visible = True                                                                 '追加 99/ 1/14
         .Font.Size = 18                                                                 '追加 99/ 1/14
         .Caption = "(ピン径  φ" & Sheets("歯厚計算").Range("G3") & ")"                  '追加 99/ 1/14
     End With                                                                            '追加 99/ 1/14
     With Sheets("工作図").TextBoxes("text1")                                            '追加 99/ 1/14
         .Visible = True                                                                '追加 99/ 1/14
         .Caption = "オーバーピン径"                                                     '追加 99/ 1/14
         .Font.Size = 20                                                                '追加 99/ 1/14
     End With                                                                           '追加 99/ 1/14
    With Sheets("工作図").TextBoxes("text2")                                            '追加 99/ 1/14
        .Visible = True                                                                 '追加 99/ 1/14
    End With                                                                            '追加 99/ 1/14


    Sheets("入力ｼｰﾄ").DrawingObjects("size6").Visible = False
    Sheets("入力ｼｰﾄ").DrawingObjects("隠しⅡ").Visible = False
'オプションボタンの表示          Excel2007対応
    Sheets("入力ｼｰﾄ").Shapes("ボタン 312").Visible = True
    Sheets("入力ｼｰﾄ").Shapes("ボタン1").Visible = True
    Sheets("入力ｼｰﾄ").Shapes("ボタン2").Visible = True
    Sheets("入力ｼｰﾄ").Shapes("ボタン3").Visible = True
    Sheets("入力ｼｰﾄ").Shapes("ボタン4").Visible = True
    Sheets("入力ｼｰﾄ").Shapes("ボタン5").Visible = True
    Sheets("入力ｼｰﾄ").Shapes("ボタン6").Visible = True
    Sheets("入力ｼｰﾄ").Shapes("ボタン7").Visible = True
    Sheets("入力ｼｰﾄ").Shapes("ボタン8").Visible = True
    Worksheets("入力ｼｰﾄ").Range("D22") = "オーバーピン径"
    Worksheets("入力ｼｰﾄ").Range("G22") = "マタギ歯厚(参考)"
    Sheets("工作図").DrawingObjects("隠しＬ").Visible = False
    Sheets("工作図").DrawingObjects("隠しＲ").Visible = False
    Sheets("入力ｼｰﾄ").Select
    Range("C23").Select
    a_res = MsgBox("サイズを入力して下さい。入力後[確認]ボタンを押して下さい。", vbExclamation, "確認")
    
    With Sheets("工作図").DrawingObjects("表隠し")
        .Visible = True
        .Top = 386.5
        .Height = 177
        .Left = 6
        .Width = 287.25
    End With
    
End Sub

'******************************************
'
'  サイズ指定後工作図に表示するためのマクロ
'
'******************************************
Sub 確認()

       wk_top = 417    ' 399.75
       wk_height = 150     '153.75
       wk_left = 6
       wk_width = 287.25
         
       For i = 1 To 8
           If Worksheets("入力ｼｰﾄ").Range("C" & Trim(i + 22)) <> "－" Then
              wk_top = wk_top + 17.75
              wk_height = wk_height - 17.75
           Else
              GoTo 次へ
           End If
       Next i
次へ:  With Sheets("工作図").DrawingObjects("表隠し")
           .Visible = True
           .Top = wk_top
           .Height = wk_height
           .Left = wk_left
           .Width = wk_width
       End With
End Sub

'******************************************
'
' サイズ指定表示
'
'******************************************

Sub サイズ1()
Attribute サイズ1.VB_ProcData.VB_Invoke_Func = " \n14"
    w_ボタン = 1
    サイズ
End Sub

Sub サイズ2()
Attribute サイズ2.VB_ProcData.VB_Invoke_Func = " \n14"
    w_ボタン = 2
    サイズ
End Sub

Sub サイズ3()
Attribute サイズ3.VB_ProcData.VB_Invoke_Func = " \n14"
    w_ボタン = 3
    サイズ
End Sub

Sub サイズ4()
Attribute サイズ4.VB_ProcData.VB_Invoke_Func = " \n14"
    w_ボタン = 4
    サイズ
End Sub

Sub サイズ5()
Attribute サイズ5.VB_ProcData.VB_Invoke_Func = " \n14"
    w_ボタン = 5
    サイズ
End Sub

Sub サイズ6()
Attribute サイズ6.VB_ProcData.VB_Invoke_Func = " \n14"
    w_ボタン = 6
    サイズ
End Sub

Sub サイズ7()
Attribute サイズ7.VB_ProcData.VB_Invoke_Func = " \n14"
    w_ボタン = 7
    サイズ
End Sub

Sub サイズ8()
Attribute サイズ8.VB_ProcData.VB_Invoke_Func = " \n14"
    w_ボタン = 8
    サイズ
End Sub

Sub サイズ()
Attribute サイズ.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("歯厚計算").Unprotect
    Sheets("歯厚計算２").Unprotect
    With Sheets("工作図").TextBoxes("text2")
        .Visible = False
    End With
    If Sheets("入力ｼｰﾄ").Range("D22") = "マタギ歯厚" And Sheets("入力ｼｰﾄ").DrawingObjects("隠しⅡ").Visible = True Then
'       Sheets("工作図").Range("AG44").Value = "歯厚マイクロ"                          '追加 99/ 1/18
'       Sheets("工作図").Range("AG45").Value = "歯厚マイクロ"                          '追加 99/ 1/18
       With Sheets("工作図").TextBoxes("micro_text1")                                '追加 99/ 1/18
           .Caption = "歯厚ﾏｲｸﾛ"                                                     '追加 99/ 1/18
           .Font.Size = 9                                                            '追加 99/ 1/18
           .Visible = True                                                           '追加 99/ 1/18
       End With                                                                      '追加 99/ 1/18

       With Sheets("工作図").TextBoxes("text1")
           .Visible = True
       End With
       With Sheets("工作図").TextBoxes("text3")
           .Visible = True
       End With
       If henkan_2 = "1" Then                            'マタギ歯厚  昇順の時
          With Sheets("工作図").TextBoxes("size1_前")
              .Visible = True
              .Caption = Format(Application.RoundUp(Sheets("歯厚計算").Range("F" & Trim(w_ボタン + 3)), 2), "##0.00") & "  ～  " & _
                         Format(Application.RoundUp(Sheets("歯厚計算").Range("F" & Trim(w_ボタン + 4)), 2), "##0.00") & "  (通常)"
              .Font.Size = 26
          End With
       Else
         If henkan_2 = "2" Then                           'マタギ歯厚  降順の時
            For i = 4 To 11
              Sheets("歯厚計算２").Range("AH" & Trim(i)) = Sheets("歯厚計算２").Range("F" & Trim(i + 1))
              Sheets("歯厚計算２").Range("AI" & Trim(i)) = Sheets("歯厚計算２").Range("F" & Trim(i))
            Next i
            With Sheets("工作図").TextBoxes("size1_前")
                .Visible = True
                .Caption = Format(Application.RoundUp(Sheets("歯厚計算２").Range("AH" & Trim(w_ボタン + 3)), 2), "##0.00") & "  ～  " & _
                           Format(Application.RoundUp(Sheets("歯厚計算２").Range("AI" & Trim(w_ボタン + 3)), 2), "##0.00") & "  (通常)"
                .Font.Size = 26
            End With
         End If
       End If
    Else
       If Sheets("入力ｼｰﾄ").Range("D22") = "オーバーピン径" And Sheets("入力ｼｰﾄ").DrawingObjects("隠しⅡ").Visible = True Then
'          Sheets("工作図").Range("AG44").Value = "ＯＢＤ測定器"                                '追加 99/ 1/18
          Sheets("工作図").Range("AG46").Font.Size = 11                                       '追加 99/ 1/18
'          Sheets("工作図").Range("AG45").Value = "ＯＢＤ測定器"                                '追加 99/ 1/18
          Sheets("工作図").Range("AG47").Font.Size = 11                                       '追加 99/ 1/18
          With Sheets("工作図").TextBoxes("micro_text1")                                      '追加 99/ 1/18
              .Caption = "OBD測定器"                                                          '追加 99/ 1/18
              .Font.Size = 8                                                                  '追加 99/ 1/18
              .Visible = True                                                                 '追加 99/ 1/18
          End With                                                                            '追加 99/ 1/18
          With Sheets("工作図").TextBoxes("micro_text11")                                      '追加 99/ 1/18
              .Caption = "ｵｰﾊﾞｰﾋﾟﾝ"                                                           '追加 99/ 1/18
              .Font.Size = 8                                                                  '追加 99/ 1/18
              .Visible = False                                                                '追加 99/ 1/18
          End With                                                                            '追加 99/ 1/18
          With Sheets("工作図").TextBoxes("micro_text12")                                      '追加 99/ 1/18
              .Caption = "ﾏｲｸﾛ"                                                               '追加 99/ 1/18
              .Font.Size = 8                                                                  '追加 99/ 1/18
              .Visible = False                                                                '追加 99/ 1/18
          End With                                                                            '追加 99/ 1/18
          With Sheets("工作図").TextBoxes("text1")
              .Visible = True
          End With
          With Sheets("工作図").TextBoxes("text3")
              .Visible = True
          End With
          If henkan_4 = "1" Then                            'オーバーピン径  昇順の時
             With Sheets("工作図").TextBoxes("size1_前")
                 .Visible = True
                 .Caption = Format(Sheets("入力ｼｰﾄ").Range("X" & Trim(w_ボタン + 6)), "##0.000") & "  ～  " & _
                            Format(Sheets("入力ｼｰﾄ").Range("Y" & Trim(w_ボタン + 6)), "##0.000") & "  (通常)"
                 .Font.Size = 26
             End With
          Else
            If henkan_4 = "2" Then                          'オーバーピン径  降順の時
               j = 0
               For i = 7 To 14
                 Sheets("入力ｼｰﾄ").Range("AB" & Trim(i + 7 - j)) = Sheets("入力ｼｰﾄ").Range("X" & Trim(i))
                 Sheets("入力ｼｰﾄ").Range("AC" & Trim(i + 7 - j)) = Sheets("入力ｼｰﾄ").Range("Y" & Trim(i))
                 j = j + 2
               Next i
               With Sheets("工作図").TextBoxes("size1_前")
                   .Visible = True
                   .Caption = Format(Sheets("入力ｼｰﾄ").Range("AB" & Trim(w_ボタン + 6)), "##0.000") & "  ～  " & _
                              Format(Sheets("入力ｼｰﾄ").Range("AC" & Trim(w_ボタン + 6)), "##0.000") & "  (通常)"
                   .Font.Size = 26
               End With
            End If
          End If
       Else
          If Sheets("入力ｼｰﾄ").Range("D22") = "オーバーピン径" And Sheets("入力ｼｰﾄ").Range("G22") = "マタギ歯厚(参考)" Then
'             Sheets("工作図").Range("AG44").Value = "ＯＢＤ測定器"                                '追加 99/ 1/18
             Sheets("工作図").Range("AG46").Font.Size = 11                                       '追加 99/ 1/18
'             Sheets("工作図").Range("AG45").Value = "ＯＢＤ測定器"                                '追加 99/ 1/18
             Sheets("工作図").Range("AG47").Font.Size = 11                                       '追加 99/ 1/18
             With Sheets("工作図").TextBoxes("micro_text1")                                      '追加 99/ 1/18
                 .Caption = "OBD測定器"                                                          '追加 99/ 1/18
                 .Font.Size = 8                                                                  '追加 99/ 1/18
                 .Visible = True                                                                 '追加 99/ 1/18
             End With                                                                            '追加 99/ 1/18

             With Sheets("工作図").TextBoxes("text1")
                 .Visible = True
             End With
             With Sheets("工作図").TextBoxes("text3")
                 .Visible = True
             End With
             If kioku_3 = "1" Then                           'オーバーピン径  昇順の時
                With Sheets("工作図").TextBoxes("size1_前")
                    .Visible = True
                    .Caption = Format(Sheets("歯厚計算").Range("G" & Trim(w_ボタン + 3)), "###0.000") & "  ～  " & _
                               Format(Sheets("歯厚計算").Range("G" & Trim(w_ボタン + 4)), "###0.000") & "  (通常)"
                    .Font.Size = 26
                End With
             Else
               If kioku_3 = "2" Then                         'オーバーピン径  降順の時
                  For i = 4 To 11
                    Sheets("歯厚計算２").Range("AL" & Trim(i)) = Sheets("歯厚計算２").Range("G" & Trim(i + 1))
                    Sheets("歯厚計算２").Range("AM" & Trim(i)) = Sheets("歯厚計算２").Range("G" & Trim(i))
                  Next i
                  With Sheets("工作図").TextBoxes("size1_前")
                      .Visible = True
                      .Caption = Format(Sheets("歯厚計算２").Range("AL" & Trim(w_ボタン + 3)), "###0.000") & "  ～  " & _
                                 Format(Sheets("歯厚計算２").Range("AM" & Trim(w_ボタン + 3)), "###0.000") & "  (通常)"
                      .Font.Size = 26
                  End With
               End If
             End If
          End If
       End If
    End If
             
    With Sheets("工作図").TextBoxes("size2")
        .Visible = True
        .Caption = "歯厚分類表による"
        .Font.Size = 6
    End With
'    With Sheets("工作図").TextBoxes("size4の左")
'        .Visible = True
'        .Caption = "歯厚分類表による"
'        .Font.Size = 12
'    End With
'    With Sheets("工作図").TextBoxes("size4")
'        .Visible = True
'        .Caption = "←"
'        .Font.Size = 16
'    End With
 
End Sub

'******************************************
'
' 改訂マーク挿入 共通マクロ
'
'******************************************

Sub 改訂()
Attribute 改訂.VB_ProcData.VB_Invoke_Func = " \n14"
    L_act_sheet = ActiveSheet.Name
    Application.Run (XLS_KTRTN + "!改訂番号"), XLS_FILE, L_act_sheet
End Sub

'******************************************
'
' 改訂履歴記入 共通マクロ       注記 ： 黒三角の名前を "三角1","三角2",・・・・"三角x"と変更しておいてください。
'
'       引数1.工作図のファイル名 （EX:"EO000101.XLS"）
'       引数2.アクティブシート名
'       引数3.改訂履歴の行数     （MAX値）
'       引数4.改訂履歴欄の先頭行のＸ軸の位値
'       引数5.改訂履歴欄の先頭行のＹ軸の年月日の位値
'       引数6.改訂履歴欄の先頭行のＹ軸の改訂Ｎｏの位値
'       引数7.改訂履歴欄の先頭行のＹ軸の改訂箇所の位値
'       引数8.改訂履歴欄の先頭行のＹ軸の改訂理由の位値
'       引数9.改訂履歴欄の先頭行のＹ軸の検印の位値

'****************************************************


Sub 改訂履歴()
Attribute 改訂履歴.VB_ProcData.VB_Invoke_Func = " \n14"
    w_max = 5
    w_y = 64        'Ｘ軸 -- (最初の位置)に設定
                    'Ｙ軸 -- それぞれ記入する位置を設定する
    w_x1 = 4        '年月日
    w_x2 = 7        '改訂ｎｏ
    w_x3 = 12       '改訂箇所
    w_x4 = 18       '改訂理由
    w_x5 = 0        '検印   （検印欄ないとき ０）
    L_act_sheet = ActiveSheet.Name
    Application.Run (XLS_KTRTN + "!改訂記入"), XLS_FILE, L_act_sheet, w_max, w_y, w_x1, w_x2, w_x3, w_x4, w_x5
End Sub
'****************************************************
'*TEST TEST TEST TEST TEST TEST TEST TEST           *
'   いったん改訂履歴を記入したあとに取り消したい時、   *
'   改訂履歴をクリアーしたい時は黒三角マークが表示     *
'   されていないと登録できないので、三角表示() を流す。*
'****************************************************
Sub 三角表示()                                  'テスト用 --- 全黒三角表示
Attribute 三角表示.VB_ProcData.VB_Invoke_Func = " \n14"
    For i = 1 To 5
        Sheets("工作図").DrawingObjects("三角" & i).Visible = True
    Next i
End Sub
'****************************************************

'   外作検索        listbox("前後")と関連付ける

'****************************************************
Sub 外作検索()
Attribute 外作検索.VB_ProcData.VB_Invoke_Func = " \n14"
        '*****************************************
        '入力ｼｰﾄ前工程、後工程のセル位置を引数とする
        '*****************************************
    y1 = 4      '前工程-縦位置
    x1 = 3      '前工程-横位置
    y2 = 4      '後工程-縦位置
    x2 = 4      '後工程-横位置
    
    Application.Run (XLS_KTRTN + "!外作前後"), y1, x1, y2, x2
    
End Sub
'******************************************
'
' 工程ＮＯ
'
'******************************************
Sub 工程NO_click()
Attribute 工程NO_click.VB_ProcData.VB_Invoke_Func = " \n14"
    Static w_a
    
    If w_a <> 1 Then
        With ActiveSheet.ListBoxes("工程_リスト")
            .RemoveAllItems
            .AddItem Text:="Ｆ１", Index:=1
            .AddItem Text:="Ｆ２", Index:=2
            .AddItem Text:="Ｆ３", Index:=3
            .AddItem Text:="ﾌﾞﾗﾝｸ", Index:=4
        End With
        Sheets("入力ｼｰﾄ").ListBoxes("工程_リスト").Visible = True
        Sheets("入力ｼｰﾄ").ListBoxes("工程_リスト").Enabled = True
        w_a = 1
    Else
        Sheets("入力ｼｰﾄ").ListBoxes("工程_リスト").Visible = False
        Sheets("入力ｼｰﾄ").ListBoxes("工程_リスト").Enabled = False
        w_a = 0
    End If
End Sub
Sub 工程NO_select()
Attribute 工程NO_select.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim 工程_select As String
    
    With ActiveSheet.ListBoxes("工程_リスト")
        工程_select = .List(.ListIndex)
        If 工程_select = "ﾌﾞﾗﾝｸ" Then
            工程_select = ""
        End If
        ActiveSheet.Range("B4") = 工程_select       '任意
    End With
    
    工程NO_click
        
End Sub

'******************************************************************************************************************
Sub test1()
Attribute test1.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("入力ｼｰﾄ").DrawingObjects("size6").Visible = True
    Sheets("入力ｼｰﾄ").DrawingObjects("隠しⅡ").Visible = True
'    With Sheets("工作図").DrawingObjects("表隠し")
'        .Visible = True
'    End With
    Sheets("工作図").DrawingObjects("隠しＬ").Visible = True
    Sheets("工作図").DrawingObjects("隠しＲ").Visible = True
End Sub


Sub ccc()
Attribute ccc.VB_ProcData.VB_Invoke_Func = " \n14"
    With Sheets("工作図").TextBoxes("size2")           '追加 99/ 1/18
        .Visible = False                               '追加 99/ 1/18
    End With                                           '追加 99/ 1/18
    With Sheets("工作図").TextBoxes("size4")           '追加 99/ 1/18
        .Visible = False                               '追加 99/ 1/18
    End With                                           '追加 99/ 1/18
End Sub


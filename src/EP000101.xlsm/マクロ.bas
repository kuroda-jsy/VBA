Attribute VB_Name = "マクロ"
'*****************************************************

'   工程名称  ： シェービング
'
'   設備GR    ： フルトＺＳＴ－７

'*****************************************************
'   定数定義領域
'*****************************************************
'XLSファイル
'Public Const XLS_FILE = "EP000101.XLSM"
Public XLS_FILE As String

'CSVファイル
'Public Const DATA_FILE = "SV.csv"                          'FOR TEST
'Public Const DATA_FILE_PATH = "\河合\csv\SV.csv"           'FOR TEST

Public Const DATA_FILE = "EP0001.CSV"
Public Const DATA_FILE_PATH = "C:\CS50\EP0001.CSV"

'*****************************************************

'       共通ファイル

'*****************************************************

'共通マクロ
Public Const XLS_KTRTN = "EPC007.XLSM"
Public Const KYOTU_FILE_PATH = "S:\CS50\DOCUMENT\EPC007.XLSM"
'Public Const KYOTU_FILE_PATH = "D:\CS50_Document\EPC007.XLSM"

'*****************************************************

'       承認印

'*****************************************************

Public Const HANKO_FILE = "EPC006.XLSM"
Public Const HANKO_FILE_PATH = "S:\CS50\DOCUMENT\EPC006.XLSM"
'Public Const HANKO_FILE_PATH = "D:\CS50_Document\EPC006.XLSM"

Public Const EPC006_CSV = "EPC006.CSV"
Public Const EPC006_CSV_PATH = "C:\CS50\EPC006.CSV"
'*****************************************************

'       変数宣言

'*****************************************************

Public w_name, w_方法, w_上限, w_下限, w_歯厚, w_モジュール, w_かみあい長さ, w_ボタン

'*****************************************************
'
' auto_open マクロ
'
'*****************************************************'
Sub AUTO_OPEN()
Attribute AUTO_OPEN.VB_ProcData.VB_Invoke_Func = " \n14"

    XLS_FILE = ActiveWorkbook.Name
    
    oldStatusBar = Application.DisplayStatusBar  '規定値の保存
    Application.DisplayStatusBar = True 'ｽﾃｰﾀｽﾊﾞｰの表示

    kyotu_macro_open                                '共通マクロファイルＯＰＥＮ

    Windows(XLS_FILE).Activate
'    ActiveWindow.WindowState = xlMinimized
    Windows(XLS_FILE).Activate


    If Trim(Sheets("工作図").TextBoxes("text_hinban").Caption) = "" Then

       Application.StatusBar = "しばらくお待ち下さい..．．"
       Application.Cursor = xlWait  '砂時計型ｶｰｿﾙ表示

       Sheets("貼付けｼｰﾄ").Visible = True

       open_file                                    'ＣＳＶファイルを開く

       csv_harituke                                 'csvファイルの内容を貼付けｼｰﾄに貼り付ける

       close_file                                   'ＣＳＶファイルを閉じる

       'システム日付をゼロサプレースした形でフォーマットし文字型に変換し工作図に貼り付ける
       Dim m_wk, m_ans
       m_wk = Sheets("入力ｼｰﾄ").Range("K3")
       m_ans = "'" + CStr(Format(m_wk, "’yy．m．d"))
       Sheets("工作図").Range("BB7") = m_ans

       入力値_clear

       Sheets("貼付けｼｰﾄ").Visible = False

'       改訂履歴 － 新図発行
'           引数 ----
'            "XLSﾌｧｲﾙ名","ｱｸﾃｨﾌﾞｼｰﾄ名","縦位置","横位置-年月日","横位置-改訂No","横位置-改訂箇所","横位置-改訂理由","横位置-検印(ないとき０)"
        Sheets("工作図").Activate
        Application.Run (XLS_KTRTN + "!新図"), XLS_FILE, ActiveSheet.Name, 64, 4, 7, 12, 18, 0

    End If

    text_harituke   '入力用テキストのフォント設定（作成済、否 にかかわらず）

    HANKO_open_file                                     '承認印共通マクロファイルを開く


    Windows(XLS_FILE).Activate

    Application.Run (XLS_KTRTN + "!入力参照設定")

    Sheets("工作図").Select

    Application.Run (XLS_KTRTN + "!参照設定")            '共通マクロの直接参照設定

    Application.Run (XLS_KTRTN + "!工作図ロック")

    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    Application.Cursor = xlNormal
    ActiveWindow.WindowState = xlMaximized

    Range("A1").Select

    '流し込み終了音
   For i = 1 To 10
      Beep  'ビープ音をならします
   Next i

End Sub
'*****************************************************
'
'  入力値をclear する
'
'*****************************************************'
Sub 入力値_clear()
Attribute 入力値_clear.VB_ProcData.VB_Invoke_Func = " \n14"

    Sheets("入力ｼｰﾄ").Range("H13") = ""
    Sheets("入力ｼｰﾄ").Range("K18") = ""
    Sheets("入力ｼｰﾄ").Range("B4") = ""
   
    'サイズ表
    Sheets("入力ｼｰﾄ").Range("C23:C30") = "－"
    Sheets("歯厚計算").Range("C9") = 0.01               '追加 99/ 1/18
    Sheets("歯厚計算２").Range("C9") = -0.01            '追加 99/ 1/26
    
    サイズ品でないとき_1              'call "サイズ品時の設定を解除する"
    
    For i = 1 To 8
        With Sheets("入力ｼｰﾄ").OptionButtons("ボタン" & Trim(Str(i)))
            .Value = xlOff
        End With
    Next i
   
         
End Sub
'*****************************************************
'
' auto_close マクロ
'
'*****************************************************'
Sub auto_close()
Attribute auto_close.VB_ProcData.VB_Invoke_Func = " \n14"


    Application.Run (XLS_KTRTN + "!工作図アンロック")
    Application.Run (XLS_KTRTN + "!入力参照解除"), XLS_FILE
    Application.Run (XLS_KTRTN + "!参照解除"), XLS_FILE            '共通マクロの直接参照解除
    
    
    ''EP0001R.CSVの保存場所取り込
    AREA_G = Workbooks(EPC006_CSV).Sheets("Epc006").Range("C1") & "\EP0001R.CSV"
    
    Application.Run (HANKO_FILE + "!Epc006_close")            'EXCEL状態の設定
       
    '承認印共通マクロファイルＣＬＯＳＥ
    Workbooks(HANKO_FILE).Close savechanges:=False
    
    '作業場所ＣＳＶファイルＣＬＯＳＥ
    Workbooks(EPC006_CSV).Close savechanges:=False
    
    '共通マクロファイルＣＬＯＳＥ
    Windows(XLS_KTRTN).Close savechanges:=False

   'リリース状態をEP0001R.CSVに書く
    Open AREA_G For Output As #1
    'If Mid(Sheets("工作図").Range("AM1"), 1, 2) = "浜北" Then
    '    w_ok = Sheets("工作図").Range("AM1")
    If Mid(Sheets("工作図").Range("AS1"), 1, 2) = "浜北" Then
        w_ok = Sheets("工作図").Range("AS1")
        Write #1, w_ok:
    Else
        Write #1, "NO":
    End If
    Close #1

    ''''保存してEXCEL終了
    On Error Resume Next
    Workbooks(XLS_FILE).Save
'    Application.Quit
    '保存確認を避けるため、保存済みにする
'    ThisWorkbook.Saved = True
    If Workbooks.Count <= 1 Then Application.CommandBars.FindControl(ID:=752).Execute
    ThisWorkbook.Close False
    
End Sub

'*****************************************************
'
' open_file マクロ
'
'*****************************************************'
Sub open_file()
Attribute open_file.VB_ProcData.VB_Invoke_Func = " \n14"

    'ＣＳＶファイルＯＰＥＮ
    Workbooks.Open filename:=DATA_FILE_PATH, ReadOnly _
         :=True
         
'   ActiveWindow.WindowState = xlMinimized

End Sub
'*****************************************************
'
'  共通マクロファイルＯＰＥＮ マクロ
'
'*****************************************************'
Sub kyotu_macro_open()
Attribute kyotu_macro_open.VB_ProcData.VB_Invoke_Func = " \n14"
    '共通マクロファイルＯＰＥＮ

    Workbooks.Open filename:=KYOTU_FILE_PATH, ReadOnly _
         :=True
'    ActiveWindow.WindowState = xlMinimized

End Sub

'*****************************************************
'
' 承認印 マクロ
'
'*****************************************************'
Sub HANKO_open_file()
Attribute HANKO_open_file.VB_ProcData.VB_Invoke_Func = " \n14"

    '状態ＣＳＶファイルＯＰＥＮ
    Workbooks.Open filename:=EPC006_CSV_PATH, ReadOnly _
         :=True

'    ActiveWindow.WindowState = xlMinimized

    '承認印共通マクロファイルＯＰＥＮ
    Workbooks.Open filename:=HANKO_FILE_PATH, ReadOnly _
         :=True
  
'    ActiveWindow.WindowState = xlMinimized

End Sub

'*****************************************************
'
' close_file マクロ
'
'*****************************************************
Sub close_file()
Attribute close_file.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.DisplayAlerts = False
    'ＣＳＶファイルＣＬＯＳＥ
    Windows(DATA_FILE).Close

    Application.DisplayAlerts = True
End Sub


'*****************************************************
' csv_harituke マクロ
'
' ・ csvファイルの内容を貼付けｼｰﾄに貼り付ける
' ・ 各項目の値を求めるマクロを呼ぶ
'
'*****************************************************
Sub csv_harituke()
Attribute csv_harituke.VB_ProcData.VB_Invoke_Func = " \n14"

    '''ＣＳＶファイルの内容をコピー＆ペースト

    Windows(DATA_FILE).Activate
    Range("A1:AZ10").Select
    Selection.Copy
    Windows(XLS_FILE).Activate
    Sheets("貼付けｼｰﾄ").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select

    data_bunbetu        '貼付けｼｰﾄ → 入力ｼｰﾄ

    hafure_get          '歯振れ計算

    規格設定            '歯形・歯筋規格設定

'    公差表示            '公差表示の設定   2003/4/17DEL

End Sub

'*****************************************************
'
' DATA_BUNBETU ﾏｸﾛ(貼付けｼｰﾄのﾃﾞｰﾀを入力ｼｰﾄにｺﾋﾟｰ
'
'*****************************************************
Sub data_bunbetu()

    '型式
    With Sheets("工作図").TextBoxes("text0")
        .Caption = Mid(Sheets("貼付けｼｰﾄ").Range("A1"), 1, 3)
        .Font.Size = 48
    End With
    
    '前工程
    Sheets("入力ｼｰﾄ").Range("C4") = Sheets("貼付けｼｰﾄ").Range("E1")
    
    '後工程
    Sheets("入力ｼｰﾄ").Range("D4") = Sheets("貼付けｼｰﾄ").Range("F1")
    
    ' 部品追番の色付け 2010/12/09 ADD
    If Right(Worksheets("貼付けｼｰﾄ").Range("A1"), 2) = "00" Then
        Worksheets("工作図").Shapes("部品追番").Visible = False
    Else
        Worksheets("工作図").Shapes("部品追番").Visible = True
    End If

    ''報連相対応05/02/08  (From)
    If Sheets("貼付けｼｰﾄ").Range("B1") = "シェービング１" Then
       l_disp = True
       l_DATA = Mid(Sheets("貼付けｼｰﾄ").Range("A1"), 8, 1) & "Ｐ側"
    ElseIf Sheets("貼付けｼｰﾄ").Range("B1") = "シェービング２" Then
       l_disp = True
       l_DATA = Trim(Val(Mid(Sheets("貼付けｼｰﾄ").Range("A1"), 8, 1)) + 1) & "Ｐ側"
    Else
       l_disp = False
       l_DATA = ""
    End If
    With Sheets("工作図").DrawingObjects("type1_txt")
        .Visible = l_disp
        .Caption = l_DATA
        .Font.Size = 40
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ''報連相対応05/02/08  (To)
    
    w_歯厚 = Sheets("貼付けｼｰﾄ").Range("I3")    'ﾏﾀｷﾞ歯厚 or オーバーピン径
    w_上限 = Sheets("貼付けｼｰﾄ").Range("J3")    '公差上限値
    w_下限 = Sheets("貼付けｼｰﾄ").Range("K3")    '公差下限値
    
    Sheets("入力ｼｰﾄ").Range("D7") = w_歯厚
    '公差上限値
    Sheets("入力ｼｰﾄ").Range("D8") = w_上限
    '公差下限値
    Sheets("入力ｼｰﾄ").Range("D9") = w_下限
    '歯形・歯筋規格前
    Sheets("入力ｼｰﾄ").Range("D10") = Sheets("貼付けｼｰﾄ").Range("E3")
    'モジュール
    w_モジュール = Sheets("貼付けｼｰﾄ").Range("A3")
    Sheets("入力ｼｰﾄ").Range("H7") = w_モジュール
    '圧力角
    Sheets("入力ｼｰﾄ").Range("H8") = Sheets("貼付けｼｰﾄ").Range("B3")
    '歯数
    Sheets("入力ｼｰﾄ").Range("H9") = Sheets("貼付けｼｰﾄ").Range("C3")
    'ネジレ角と方向
    If Sheets("貼付けｼｰﾄ").Range("F3") <> "" And Sheets("貼付けｼｰﾄ").Range("G3") <> "" Then
        Sheets("入力ｼｰﾄ").Range("H11") = CStr(Sheets("貼付けｼｰﾄ").Range("F3")) + Sheets("貼付けｼｰﾄ").Range("G3")
    Else
        Sheets("入力ｼｰﾄ").Range("H11") = ""
    End If
    'かみあい判定長さ
    w_かみあい長さ = Sheets("貼付けｼｰﾄ").Range("D3")
    Sheets("入力ｼｰﾄ").Range("H15") = Application.RoundUp((w_かみあい長さ + (0.375 * w_モジュール)) * 4, 0)
    '外径
    Sheets("入力ｼｰﾄ").Range("N7") = Sheets("貼付けｼｰﾄ").Range("R3")
'    'ｼｪｰﾋﾞﾝｸﾞ加工長さ下限（かみあい判定長さと同）
'削除 (K.Y)    Sheets("入力ｼｰﾄ").Range("H12") = Sheets("入力ｼｰﾄ").Range("H14")
    
    'ｼｪｰﾋﾞﾝｸﾞｶｯﾀｰ
    Sheets("入力ｼｰﾄ").Range("D18") = Sheets("貼付けｼｰﾄ").Range("C4")
    'ｼｪｰﾋﾞﾝｸﾞｱｰﾊﾞｰ
    Sheets("入力ｼｰﾄ").Range("D19") = Sheets("貼付けｼｰﾄ").Range("C5")
    
    '****************************************************
    'マタギの場合とオーバーピンの場合で貼り分ける
    '****************************************************
    
    If Sheets("貼付けｼｰﾄ").Range("H3") = "O" Then               'オーバーピンの場合
       
       Sheets("入力ｼｰﾄ").Range("C7").Value = "オーバーピン径"
       Sheets("入力ｼｰﾄ").Range("C7").Font.Size = 10
       Sheets("入力ｼｰﾄ").Range("C15").Value = "ピン径"
       Sheets("入力ｼｰﾄ").Range("C15").Font.Size = 11
       Sheets("入力ｼｰﾄ").Range("G16").Value = "ｵｰﾊﾞｰﾋﾟﾝ径のﾊﾞﾗﾂｷ"
       Sheets("入力ｼｰﾄ").Range("G16").Font.Size = 10
       Sheets("入力ｼｰﾄ").Range("H16").Value = 0.05
       Sheets("入力ｼｰﾄ").Range("H16").Font.Size = 11
'       Sheets("工作図").Range("AG45").Value = "ｵｰﾊﾞｰﾋﾟﾝﾏｲｸﾛ"
'       Sheets("工作図").Range("AG45").Font.Size = 11
       With Sheets("工作図").TextBoxes("textp1")                                            '追加 99/ 1/14
           .Visible = True                                                                 '追加 99/ 1/14
           .Font.Size = 20                                                                 '追加 99/ 1/14
           .Caption = "ｵｰﾊﾞｰﾋﾟﾝ径のﾊﾞﾗﾂｷ  " & Sheets("入力ｼｰﾄ").Range("H16") & "  以下"     '追加 99/ 1/14
       End With                                                                            '追加 99/ 1/14

       
       With Sheets("工作図").TextBoxes("text1")
                            .Caption = "オーバーピン径"
                            .Font.Size = 20       '変更
                            .Visible = True
       End With
       With Sheets("工作図").TextBoxes("micro_text11")
                            .Caption = "ｵｰﾊﾞｰﾋﾟﾝ"
                            .Font.Size = 8
                            .Visible = True
       End With
       With Sheets("工作図").TextBoxes("micro_text12")
                            .Caption = "ﾏｲｸﾛ"
                            .Font.Size = 8
                            .Visible = True
       End With
       With Sheets("工作図").TextBoxes("micro_text1")
                            .Visible = False
       End With
       
        'ピン径
        Sheets("入力ｼｰﾄ").Range("D15") = "φ" & Sheets("貼付けｼｰﾄ").Range("L3")
        With Sheets("工作図").TextBoxes("text3")                                           '追加 99/ 1/14
            .Visible = True                                                                '追加 99/ 1/14
            .Font.Size = 18                                                                '追加 99/ 1/14
            .Caption = "(ピン径  " & Sheets("入力ｼｰﾄ").Range("D15") & ")"                   '追加 99/ 1/14
       End With
                                                                                  '追加 99/ 1/14
       'ｵｰﾊﾞｰﾋﾟﾝ径 管理規格
       If w_歯厚 = "" Or w_下限 = "" Then
           Sheets("入力ｼｰﾄ").Range("E7") = ""
       Else
           Sheets("入力ｼｰﾄ").Range("E7") = CStr(Application.Round(w_歯厚 + w_下限 + 0.02, 2)) & "±0.02"
       End If
       
    ElseIf Sheets("貼付けｼｰﾄ").Range("H3") = "M" Then                'マタギ歯厚の場合
    
       Sheets("入力ｼｰﾄ").Range("D15") = Sheets("貼付けｼｰﾄ").Range("L3")
       'マタギの場合の寸法計算
       If w_歯厚 = "" Or w_下限 = "" Or w_上限 = "" Then
           Sheets("入力ｼｰﾄ").Range("E7") = ""
       Else
       Sheets("入力ｼｰﾄ").Range("E7") = CStr(Application.Round((w_下限 + w_上限) / 2 + w_歯厚 - 0.015, 2)) & "±0.01"
       End If
    Else
    
       Sheets("入力ｼｰﾄ").Range("C7").Value = ""
       Sheets("入力ｼｰﾄ").Range("E7").Value = ""
       Sheets("入力ｼｰﾄ").Range("D15").Value = ""
       Sheets("入力ｼｰﾄ").Range("C15").Value = ""
       Sheets("入力ｼｰﾄ").Range("G16").Value = ""
       Sheets("入力ｼｰﾄ").Range("H16").Value = ""
       Sheets("入力ｼｰﾄ").Range("H8").Value = ""
    End If
    
    
End Sub
'*****************************************************
'
' 歯形・歯筋規格 設定
'
'*****************************************************

Sub 規格設定()
Attribute 規格設定.VB_ProcData.VB_Invoke_Func = " \n14"
    w_精度 = Sheets("貼付けｼｰﾄ").Range("E3")
    w_規格 = Sheets("貼付けｼｰﾄ").Range("M3")
    
    Sheets("入力ｼｰﾄ").Unprotect                'yamauchi
     
    '前回分のクリアー
    For i = 7 To 16
        Sheets("入力ｼｰﾄ").Cells(i, 11) = ""
    Next i
    
    Sheets("入力ｼｰﾄ").Range("K7:K16").Interior.ColorIndex = 40              'グレー
    Sheets("入力ｼｰﾄ").Range("E10").Interior.ColorIndex = 40
    
    If Mid(w_精度, 4, 1) = "4" Then                                         '"JIS４級"の場合
        If w_規格 = "有" Then
            With Sheets("入力ｼｰﾄ").Buttons("ボタン 114")
                                    .Visible = False
            End With
            'Sheets("入力ｼｰﾄ").Range("E10") = "ヤマハ４級"
            Sheets("入力ｼｰﾄ").Range("E10").Formula = "=IF(D10<>"""",""ヤマハ""&MID(D10,4,1)&""級"")"         '追加 99/ 1/14
            Sheets("入力ｼｰﾄ").Range("K7:K16").Interior.ColorIndex = 19       'クリーム色
            
            '入力指示
            Sheets("入力ｼｰﾄ").Range("K7") = "XX ～ XXμｍ"
            Sheets("入力ｼｰﾄ").Range("K8") = "－"
            Sheets("入力ｼｰﾄ").Range("K9") = "X μｍ以内"
            Sheets("入力ｼｰﾄ").Range("K10") = "XX ～ XXμｍ"
            Sheets("入力ｼｰﾄ").Range("K11") = "XX mm以下"
            Sheets("入力ｼｰﾄ").Range("K12") = "－"
            Sheets("入力ｼｰﾄ").Range("K13") = "－"
            Sheets("入力ｼｰﾄ").Range("K14") = " 0 ±  5μｍ"
            Sheets("入力ｼｰﾄ").Range("K15") = "XX μｍ以内"
            Sheets("入力ｼｰﾄ").Range("K16") = "35 μｍ以内"
            
        ElseIf w_規格 = "無" Then
            With Sheets("入力ｼｰﾄ").Buttons("ボタン 114")                    '歯形・歯筋規格 －－ 歯形・歯筋規格が "無" 時、ボタンを表示
                                  .Visible = True
            End With
            Sheets("入力ｼｰﾄ").Range("E10") = ""
        End If
    Else
        With Sheets("入力ｼｰﾄ").Buttons("ボタン 114")
                                .Visible = False
        End With
'        Sheets("入力ｼｰﾄ").Range("E10") = ""                                 '"JIS４級" 以外は手入力
        Sheets("入力ｼｰﾄ").Range("E10").Formula = "=IF(D10<>"""",""ヤマハ""&MID(D10,4,1)&""級"")"         '追加 99/ 1/14
        Sheets("入力ｼｰﾄ").Range("E10").Interior.ColorIndex = 19             'クリーム色
        Sheets("入力ｼｰﾄ").Range("K7:K16").Interior.ColorIndex = 19          'クリーム色
        
        '入力指示
        Sheets("入力ｼｰﾄ").Range("K7") = "XX ～ XXμｍ"
        Sheets("入力ｼｰﾄ").Range("K8") = "－"
        Sheets("入力ｼｰﾄ").Range("K9") = "X μｍ以内"
        Sheets("入力ｼｰﾄ").Range("K10") = "XX ～ XXμｍ"
        Sheets("入力ｼｰﾄ").Range("K11") = "XX mm以下"
        Sheets("入力ｼｰﾄ").Range("K12") = "－"
        Sheets("入力ｼｰﾄ").Range("K13") = "－"
        Sheets("入力ｼｰﾄ").Range("K14") = " 0 ±  5μｍ"
        Sheets("入力ｼｰﾄ").Range("K15") = "XX μｍ以内"
        Sheets("入力ｼｰﾄ").Range("K16") = "35 μｍ以内"
    End If
    
    
End Sub

'*****************************************************
'
' フォントサイズの設定 マクロ
'
'*****************************************************
Sub text_harituke()
Attribute text_harituke.VB_ProcData.VB_Invoke_Func = " \n14"

    '型式
    With Sheets("工作図").TextBoxes("text0")
        .Font.Size = 48
    End With
    '品名
    With Sheets("工作図").TextBoxes("text_hinmei")
        .Font.Size = 14
    End With
    '品番
    With Sheets("工作図").TextBoxes("text_hinban")
        .Font.Size = 14
    End With
    'マタギ歯厚．オーバーピン径
    With Sheets("工作図").TextBoxes("text2")
        .Font.Size = 26
    End With
    
    'マタギ枚数．ピン径
    With Sheets("工作図").TextBoxes("text_kakoumen")
        .Font.Size = 7
    End With
    '工程Ｎｏ
    With Sheets("工作図").TextBoxes("工程_text")
        .Font.Size = 20
    End With
    
End Sub
'*****************************************************
'
'  Macro Name : hafure_get
' （歯フレを共通マクロファイル内(EPC007.XLS)の"hafure_keisan"マクロより求める）
'
'*****************************************************'
Sub hafure_get()
Attribute hafure_get.VB_ProcData.VB_Invoke_Func = " \n14"
  
    Dim RET, l_級, l_mod, l_P径
        If Sheets("貼付けｼｰﾄ").Range("E3") = "" Or Sheets("貼付けｼｰﾄ").Range("A3") = "" Or Sheets("貼付けｼｰﾄ").Range("C3") = "" Then
            
            Sheets("入力ｼｰﾄ").Range("H17") = ""
            Exit Sub
        End If
        
       l_級 = CInt(Mid(Sheets("貼付けｼｰﾄ").Range("E3"), 4, 1))  '製品図規格精度
       l_mod = Sheets("貼付けｼｰﾄ").Range("A3")                  'モジュール
       l_P径 = l_mod * Sheets("貼付けｼｰﾄ").Range("C3")          'ピッチ円径 (ﾓｼﾞｭｰﾙ * 歯数)
    
    RET = Application.Run(XLS_KTRTN + "!hafure_keisan", l_級, l_mod, l_P径)
    
    Windows(XLS_FILE).Activate
    
    If RET <> 99999 Then
        
        Sheets("入力ｼｰﾄ").Range("H17") = Application.RoundDown(RET, 2)
        
    Else
       ''エラー時
       Sheets("入力ｼｰﾄ").Range("F10") = Null
       MsgBox ("歯振れ計算 にエラーがありました。")
    End If
End Sub
'*****************************************************
'
'  軸径限界値
'       軸径の摩耗限界値 を設定する
'
'*****************************************************'
Sub 軸径限界値()
Attribute 軸径限界値.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim w_種類, w_熱前内径
    
    w_種類 = Sheets("貼付けｼｰﾄ").Range("N3")
    w_限界値 = ""
    w_熱前内径 = ""
    
    w_ans = MsgBox("｢ 内径種類 ｣ 、｢ 熱前内径 ｣ より ｢ ｼｪｰﾋﾞﾝｸﾞｱｰﾊﾞｰの軸径 ｣の 摩耗限界値 を設定します。" & _
                    Chr(13), vbOKCancel + vbExclamation, "軸径摩耗限界値設定")
    Select Case w_ans
    Case Is = vbOK
    Case Else
        w_限界値 = Sheets("入力ｼｰﾄ").Range("K18")
        GoTo last
    End Select
    
    Select Case w_種類
    Case Is = "角SP"
        w_ans = MsgBox("アーバー軸の種類 は ｢ 角SP ｣ ですか？" & Chr(13) & Chr(13) & _
                       "     注 ： ｢ いいえ ｣ を選択すると軸の種類は｢ 丸 ｣ となります。 " & Chr(13), _
                        vbYesNoCancel + vbQuestion, "ｱｰﾊﾞｰ種類 - 角SP ")
        Select Case w_ans
        Case Is = vbYes
            w_熱前内径 = Sheets("貼付けｼｰﾄ").Range("P3")
            
        Case Is = vbNo
            w_熱前内径 = Sheets("貼付けｼｰﾄ").Range("Q3")
            
        Case Else
            Exit Sub
        End Select
    
    Case Is = "INV-SP"
        w_ans = MsgBox("軸の種類 は ｢ ｲﾝﾎﾞﾘｭｰﾄｽﾌﾟﾗｲﾝ ｣ ですか？" & Chr(13) & Chr(13) & _
                       "     注 ： ｢ いいえ ｣ を選択すると軸の種類は｢ 丸 ｣ となります。 " & Chr(13), _
                        vbYesNoCancel + vbQuestion, "ｱｰﾊﾞｰ種類 - INV-SP ")
        Select Case w_ans
        Case Is = vbYes             'ｲﾝﾎﾞﾘｭｰﾄｽﾌﾟﾗｲﾝ -- 熱前大径呼び
            w_熱前内径 = Sheets("貼付けｼｰﾄ").Range("P3")
            
        Case Is = vbNo              '丸内径 --  熱前SP小径呼び
            w_熱前内径 = Sheets("貼付けｼｰﾄ").Range("Q3")
            
        Case Else
            Exit Sub
        End Select
    
    Case Is = "丸", "花柄", "ﾌﾞｯｼｭ", "ｶﾗｰ", "ｷｰ", "ﾍﾘｶﾙ"
        w_熱前内径 = Sheets("貼付けｼｰﾄ").Range("O3")        '熱前丸内径呼び
        
    Case Else
        a = MsgBox("｢ 内径種類 ｣ が登録されていません。 確認してください。", vbExclamation)
        GoTo last
    End Select
    
    If w_熱前内径 <> "" Then
        If IsNumeric(w_熱前内径) Then
            w_限界値 = Application.Round(w_熱前内径 - 0.03, 2)
        Else
            a = MsgBox("｢ 熱前内径値 ｣ が数値でない。", vbExclamation)
            GoTo last
        End If
    Else
        a = MsgBox("内径寸法が登録されていません。", vbExclamation)
    End If
    
last:
    Sheets("入力ｼｰﾄ").Range("K18") = w_限界値
    
End Sub

'*******************************************************************************************************************************
'             テ  ス  ト  用
'*******************************************************************************************************************************
Sub xxxx()
Attribute xxxx.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.Cursor = xlNormal
End Sub

Sub ggg()
Attribute ggg.VB_ProcData.VB_Invoke_Func = " \n14"
    With Sheets("工作図").TextBoxes("textp1")
        .Visible = True
        .Font.Size = 20
        .Caption = "ピッチ誤差  " & Sheets("入力ｼｰﾄ").Range("H16") & "  以下"
    End With
'    With Sheets("工作図").TextBoxes("textp1")
'        .Visible = True
'        .Font.Size = 20
'        .Caption = "ｵｰﾊﾞｰﾋﾟﾝ径のﾊﾞﾗﾂｷ  " & Sheets("入力ｼｰﾄ").Range("H16") & "  以下"
'    End With
    With Sheets("工作図").TextBoxes("text3")
        .Visible = True
        .Font.Size = 18
        .Caption = "(マタギ枚数  " & Sheets("入力ｼｰﾄ").Range("D13") & "枚)"
    End With
'     With Sheets("工作図").TextBoxes("text3")
'         .Visible = True
'         .Font.Size = 18
'         .Caption = "(ピン径  φ" & Sheets("入力ｼｰﾄ").Range("D13") & ")"
'     End With
'     With Sheets("工作図").TextBoxes("text1")
'         .Visible = True
'         .Caption = "オーバーピン径"
'         .Font.Size = 20
'     End With
     With Sheets("工作図").TextBoxes("text1")
         .Visible = True
         .Caption = "マタギ歯厚"
         .Font.Size = 20
     End With
    With Sheets("工作図").TextBoxes("text2")
        .Visible = True
    End With
'    With Sheets("工作図").TextBoxes("size1_前")
'        .Visible = False
'        .Caption = Sheets("歯厚計算").Range("F4") & "  ～  " & Sheets("歯厚計算").Range("F5") & "  (通常)"
'        .Font.Size = 26
'    End With
'    With Sheets("工作図").TextBoxes("size1_前")
'        .Visible = False
'        .Caption = Sheets("歯厚計算").Range("G4") & "  ～  " & Sheets("歯厚計算").Range("G5") & "  (通常)"
'        .Font.Size = 26
'    End With
End Sub

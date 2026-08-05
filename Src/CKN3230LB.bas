Attribute VB_Name = "CKN3230LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   社内かんばん発行(中間工程)
'*
'*   モジュールID  : CKN3230LB.BAS
'*
'*                   CKN3230LPrint(Integer)
'*
'*   初期設定値 : CKN3230_PrtKB(かんばん種類) :  1 固定 /社内
'*                CKN3230_PrtSy(発行種類)     :  1 固定 /中間工程
'*
'*   引数 : PrtSy    =  0  :新規発行
'*                      1  :再発行
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:
'    If Not CKN3230LPrint(0) Then Exit Sub
'
'*   作成日  : 2009/10/20   By CIS
'*
'*---------------------------------------------------------------------------------
Global CKN3230_PrtKB            As Byte
Global CKN3230_PrtSy            As Byte
Global CKN3230_ImagePath        As String
Global CKN3230_CoReport         As New CoReports85        ' クラス定義

Global CKN3230_ColorCode(9)     As String
Global CKN3230_Color(9)         As Long
Global CKN3230_ColorSetFlg      As Boolean

Global kBas_Proc        As String
    
Dim CorItem                     As CrObject
Dim User_BarCode                As String
Dim SetColor                    As Long
Dim EndNo                       As Long

Dim bCnt                        As Long
Dim bStr                        As String

Dim BarHeght                    As Long
Dim OutPrinter                  As String           ' 2006/02/14 Insert

Global CKN3230_TestPrt          As Boolean
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/
'_/             印　　刷　　処　　理
'_/
'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Function CKN3230LPrint(Optional PrtSy As Integer = 0)
    CKN3230LPrint = False
    
    Qr_Jisya = ""
    Qr_HMoto = ""
    ' 自社コード
    Qr_Jisya = RTrim(gCompany)
    '発行元
    If Not CKN3230_TestPrt Then
       Call ItemsClearIRN
       IRN.bk = "Option": IRN.rno = 4
       Call IRNRead("", 1, 1)
       Qr_HMoto = RTrim(Mid(IRN.oyian, 1, 7))
       If Trim(Qr_HMoto) = "" Then Qr_HMoto = "SIMPR-A"
    Else
       Qr_HMoto = "TEST"
    End If
    
    App.Title = "工程かんばん"
' 印刷進行画面表示ﾌﾗｸﾞｾｯﾄ
    pref_MsgDisp = True

    gSL_Select = ""
    gSL_Select = gSL_Select & " Select KH.*,HT.最終発行連番"
    gSL_Select = gSL_Select & "   From かんばん発行テーブル KH"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番取引先マスタ HT"
    gSL_Select = gSL_Select & "             ON HT.品番 = KH.品番30"
    gSL_Select = gSL_Select & "            AND HT.取引先 = KH.取引先"
    gSL_Select = gSL_Select & "            AND HT.受入 = KH.受入"
    gSL_Select = gSL_Select & "            AND HT.品目 = 1"
'    gSL_Select = gSL_Select & "            AND HT.手配区分 = '2'"
    gSL_Select = gSL_Select & "  Where KH.かんばん種類 = 1"
    gSL_Select = gSL_Select & "    And KH.発行種類 = 1"
    gSL_Select = gSL_Select & "    And KH.発行済 = 0"
    If CKN3230_TestPrt Then
       gSL_Select = gSL_Select & "    AND KH.発行区分 = 3"
    End If
    gSL_Select = gSL_Select & "  Order By KH.発行種類,KH.取引先,KH.受入,KH.背番号,KH.品番,KH.発行NO"

    If Not KHTRead(gSL_Select) Then
        Exit Function
    End If

    pref_TCnt = CisDB.RecordCount

'+++++++++++++++++++++++++++++++'
'+  印刷処理
'+++++++++++++++++++++++++++++++'
    Set pref_Form = CKN3230L
    pref_Msg = "【 工程かんばん 印 刷 中 】"

    pref_Form.Show vbModal
    Unload pref_Form
    Set pref_Form = Nothing

    CKN3230LPrint = True
End Function
'+----------------------------------------------------------+
'+                    印  刷  処  理                        ＋
'+----------------------------------------------------------+
Public Function CoPrintN3230() As Boolean
    Dim PageCnt     As Long
    
    CoPrintN3230 = False
    On Error GoTo CoPrintN3230_Er
    
' イメージ格納先
    With CisFun
        .INI_FileID = gIniExe & gwsIniName
        .INI_Section = "Option"
        .INI_Key = "Image":       .INI_Default = gIniExe
        .GetIni
        If Right(RTrim(.INI_String), 1) <> "\" Then
            CKN3230_ImagePath = RTrim(.INI_String) & "\"
        Else
            CKN3230_ImagePath = RTrim(.INI_String)
        End If
    ' かんばん印刷先
        OutPrinter = ""
        .INI_Section = "かんばん印刷先"
        .INI_Default = ""
        .INI_Key = CisFun.WSID
        .GetIni
        If Trim(.INI_String) <> "" Then
           OutPrinter = RTrim(.INI_String)
        End If
        .INI_FileID = gIniExe & gIniName
    End With
    
    pref_NCnt = 0                ' 現在発行ページ初期化
    
    
    With CKN3230_CoReport
        .SetDraw = pref_Form.CrDraw1

        .SetPrtFormCnt = 1
        .SetPrtForm = "LKN3230"        ' ﾌｫｰﾑIDをｾｯﾄ
        
        .SetPrtPath = gIniExe & "Prt\" ' ﾌｫｰﾑ格納先
        .SetPrtOut = PrinterX          ' ﾌﾟﾘﾝﾀｰ
        .SetPrtName = "工程かんばん"                      ' 帳票名
        .SetPrinter = OutPrinter
        .SetPaperSize = A4
        .SetPaperOrientation = 縦
        PageCnt = 0
        '------ 印刷処理開始
        If Not .CorOutStart Then GoTo CoPrintN3230_Er
        
        ' ﾌｫｰﾑ項目非表示
        Call .CorItemHeide
                
        Do Until Not KHT_RDSTS
            PageCnt = PageCnt + 1
        '-----------------------------------------------    最終発行NOを求める(新規発行時)
            EndNo = 0
            If KHT.発行区分 = 0 And KHT.発行済 = 0 Then
               EndNo = KHT.最終発行連番
               gSL_Select = "Select Max(発行NO) 発行NO From かんばん流動社内テーブル "
               gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(KHT.取引先) & "'"
               gSL_Select = gSL_Select & " And   受入 = '" & RTrim(KHT.受入) & "'"
               gSL_Select = gSL_Select & " And   品番 = '" & RTrim(KHT.品番) & "'"
               If KRSRead(gSL_Select, 1, 1) Then
                  If KRS.発行NO > KHT.最終発行連番 Then
                     EndNo = KRS.発行NO
                  End If
               End If
               KHT.発行NO = EndNo + 1
            End If
            
            '------ かんばん流動管理テーブル 版数獲得
             gSL_Select = "Select 版数 From かんばん流動社内テーブル"
             gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(KHT.取引先) & "'"
             gSL_Select = gSL_Select & " And   受入 = '" & RTrim(KHT.受入) & "'"
             gSL_Select = gSL_Select & " And   品番 = '" & RTrim(KHT.品番) & "'"
             gSL_Select = gSL_Select & " And   発行NO = " & KHT.発行NO & ""
             If Not KRTRead(gSL_Select, 1, 1) Then
                 KHT.版数 = 1
             Else
                 KHT.版数 = KRT.版数 + 1
            End If
        
            ' QRエリア転送
            Call QRAreaClear
            With QRArea
                .品番 = KHT.品番
                .品目 = KHT.品目
                .取引先 = KHT.取引先
                .受入 = KHT.受入
                .手配区分 = KHT.手配区分
                .版数 = KRT.版数
                .発行日 = Format(Now(), "YYYYMMDD")
                .背番号 = KHT.背番号
                .連番 = KHT.発行NO
                .収容数 = KHT.収容数
                .サイクル = KHT.サイクル
                .納入先 = KHT.納入先
                .納入先受入 = KHT.納入先受入
                .工程 = KHT.工程
                .次工程 = KHT.次工程
                .編集品番 = KHT.表示品番
                .手配区分名 = KHT.手配区分名
                .アドレス = KHT.アドレス
    '            If kn_PrtKB = 2 Then
    '               .品目 = 2
    '               .品番 = KHT.材質
    '               .板厚 = KHT.板厚
    '               .幅 = KHT.幅
    '               .長さ = KHT.長さ
    '               .寸法区分 = KHT.寸法区分
    '               .材料管理番号 = KHT.材料管理番号
    '               .材料収容数 = KHT.収容数
    '            End If
            End With
        
            '------ 項目転送
            If Not ItemMoveFore(PageCnt) Then GoTo CoPrintN3230_Er
            '------ 印刷(１ﾍﾟｰｼﾞ分)
            If PageCnt >= 3 Then
               If Not .CoPrintOut Then GoTo CoPrintN3230_Er
               Call .CorItemHeide
                PageCnt = 0
            End If
            
            If Not CKN3230_TestPrt Then
                 '------ かんばん流動社内テーブル更新
                  gSL_Select = "Select * From かんばん流動社内テーブル"
                  gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(KHT.取引先) & "'"
                  gSL_Select = gSL_Select & " And   受入 = '" & RTrim(KHT.受入) & "'"
                  gSL_Select = gSL_Select & " And   品番 = '" & RTrim(KHT.品番) & "'"
                  gSL_Select = gSL_Select & " And   発行NO = " & KHT.発行NO & ""
                  If Not KRSRead(gSL_Select, 1, 1) Then
                     Call ItemsClearKRS
                     With KRS
                          .取引先 = KHT.取引先
                          .受入 = KHT.受入
                          .品番 = KHT.品番
                          .発行NO = KHT.発行NO
                          .発行区分 = 0
                          .収容区分 = KHT.収容区分
                          .収容数 = KHT.収容数
                          .箱種 = KHT.収容器
                          .版数 = 1
                      End With
                      Call KRSInsert(1)
                  Else
                      With KRS
                          .版数 = .版数 + 1
                          .発行区分 = 0
                          .収容区分 = KHT.収容区分
                          .収容数 = KHT.収容数
                          .箱種 = KHT.収容器
                      End With
                      Call KRSUpdate(1)
                 End If
                '--------------------- かんばん発行ワーク更新
                 With CisDB
                     .ConnectNo = 1
                     .SQL = "Update かんばん発行テーブル Set 発行済 = 1"
                     If KHT.発行区分 = 0 And KHT.発行済 = 0 Then
                        .SQL = .SQL & " , 発行NO = " & KHT.発行NO
                     End If
                     .SQL = .SQL & " Where かんばん種類 = 1"
                     .SQL = .SQL & " And   品番 = '" & RTrim(KHT.品番) & "'"
                     .SQL = .SQL & " And   取引先 = '" & RTrim(KHT.取引先) & "'"
                     .SQL = .SQL & " And   受入 = '" & RTrim(KHT.受入) & "'"
                     .SQL = .SQL & " And   発行NO = " & KHT.発行NO
                     If Not .DBExec Then GoTo CoPrintN3230_Er
                 '--------------------- 品番取引先マスタ　最終発行連番 更新
                     If KHT.発行区分 = 0 And KHT.発行済 = 0 Then
                        .SQL = "Update 品番取引先マスタ Set  "
                        .SQL = .SQL & "最終発行連番 = Case When 最終発行連番 > " & KHT.発行NO & " THEN 最終発行連番 Else " & KHT.発行NO & " End,"
                        .SQL = .SQL & "更新日 = GETDATE(),"
                        .SQL = .SQL & "更新者 = '" & gTanto & "',"
                        .SQL = .SQL & "更新端末 = SubString(Host_Name(),1,20)"
                        .SQL = .SQL & " Where 品番 = '" & RTrim(KHT.品番) & "'"
                        .SQL = .SQL & "   And 取引先 = '" & RTrim(KHT.取引先) & "'"
                        .SQL = .SQL & "   And 受入 = '" & RTrim(KHT.受入) & "'"
                        If Not .DBExec Then GoTo CoPrintN3230_Er
                     End If
                 End With
            Else
                With CisDB
                    .SQL = ""
                    .SQL = .SQL & " DELETE FROM かんばん発行テーブル"
                    .SQL = .SQL & "  WHERE 発行区分 = 3"
                    Call .DBExec
                End With
            End If
            
            ' 進捗ゲージ更新
            pref_NCnt = pref_NCnt + 1
            pref_Cnt = Format(pref_NCnt, "###") & " / " & Format(pref_TCnt, "###")
            DoEvents
            '中断処理
            If pref_Cansel_Flg Then
                If MsgBox("処理を中断します。よろしいですか？", 33, "印刷") = 1 Then
                   GoTo CoPrintN3230_ED
                Else
                   pref_Cansel_Flg = False
                End If
            End If
            
            Call KHTReadNext
        Loop
        
        If PageCnt <> 0 Then
           If Not .CoPrintOut Then GoTo CoPrintN3230_Er
        End If
        
CoPrintN3230_ED:
        Call KHTClose
        
        '------ 印刷終了
        If Not .CoOutEnd Then GoTo CoPrintN3230_Er
    End With
    
    
    CoPrintN3230 = True
    Exit Function
CoPrintN3230_Er:
    Call CisFun.ErrorBox("印刷エラー")
End Function
'*------------------------------------------------------------------*
'*                                                                  *
'*      かんばん　発行内容　項目転送                                *
'*                                                                  *
'*------------------------------------------------------------------*
Private Function ItemMoveFore(PageNo As Long) As Boolean
    Dim wCnt As Integer

    ItemMoveFore = False

    With CKN3230_CoReport
        '---------- 部品区分
        Select Case KHT.部品区分
               Case 0: gStr = "完  成  品"
               Case 1: gStr = "購  入  品"
               Case 2: gStr = "仕  掛  品"
        End Select
        If PageNo = 1 Then
           If Not .CorData("部品区分-" & Format(PageNo, "0"), gStr, "S", 10) Then Exit Function
        Else
           If Not .CorData("部品区分-" & Format(PageNo, "0"), gStr, "", 10) Then Exit Function
        End If
        '---------- 品番
        If Not .CorData("品番-" & Format(PageNo, "0"), KHT.表示品番, "", 18) Then Exit Function
        '---------- 得意先
        If Not .CorData("得意先-" & Format(PageNo, "0"), KHT.得意先名 & " 向け", "", 40) Then Exit Function
        '---------- 背番号
        If Not .CorData("背番号-" & Format(PageNo, "0"), KHT.背番号, "", 8) Then Exit Function
        '---------- 所番地
        If Not .CorData("所番地-" & Format(PageNo, "0"), KHT.置場, "", 10) Then Exit Function
        '---------- 収容器
        If Not .CorData("収容器-" & Format(PageNo, "0"), KHT.収容器, "", 10) Then Exit Function
        '---------- 収容数
        If Not .CorData("収容数-" & Format(PageNo, "0"), Format(KHT.収容数, "#"), "", 7) Then Exit Function
        '---------- ロット
        If Not .CorData("ロット-" & Format(PageNo, "0"), Format(KHT.生産ロット, "#"), "", 7) Then Exit Function
        '---------- 生産区分
        If Not .CorData("生産区分-" & Format(PageNo, "0"), KHT.生産区分, "", 1) Then Exit Function
        '---------- 部署名
        If Not .CorData("部署名-" & Format(PageNo, "0"), KHT.部署名, "", 12) Then Exit Function
        '---------- 返却場所名
        If Not .CorData("返却場所名-" & Format(PageNo, "0"), KHT.返却場所名, "", 12) Then Exit Function
        '---------- 後工程名１
        If Not .CorData("後工程-" & Format(PageNo, "0"), KHT.後工程名1, "", 12) Then Exit Function
        '---------- 後工程名２
        If Not .CorData("後工程名-" & Format(PageNo, "0"), KHT.後工程名2, "", 12) Then Exit Function
        '---------- 発行連番
        If Not .CorData("発行連番-" & Format(PageNo, "0"), Format(KHT.発行NO, "#"), "", 3) Then Exit Function
        '---------- 版数
        If Not .CorData("版数-" & Format(PageNo, "0"), Format(KHT.版数, "#"), "", 3) Then Exit Function
        '---------- 発行日付
        If Not .CorData("発行日付-" & Format(PageNo, "0"), Format(Now(), "yyyy.MM.dd"), "", 10) Then Exit Function
        '---------- QR
        gStr = QR_OutCreate
        If Not .CorQRData("QR-" & Format(PageNo, "0"), gStr, "", 150) Then Exit Function
        Debug.Print (gStr)
    
        If RTrim(KHT.IMAGE) <> "" Then
            With CisFun
                .FE_FullPath = KHT.IMAGE
                .GetPathName
                gStr = Replace(.FE_FullPath, .FE_File, "")
            End With
            Call .CorImageSet("Image1-" & Format(PageNo, "0"), gStr, CisFun.FE_File)
        End If
        If RTrim(KHT.IMAGE) <> "" Then
            With CisFun
                .FE_FullPath = KHT.ポイントIMAGE
                .GetPathName
                gStr = Replace(.FE_FullPath, .FE_File, "")
            End With
            Call .CorImageSet("Image2-" & Format(PageNo, "0"), gStr, CisFun.FE_File)
        End If
        With CisFun
            .FE_FullPath = CKN3230_ImagePath & "Logo.jpg"
            .GetPathName
            gStr = Replace(.FE_FullPath, .FE_File, "")
        End With
        Call .CorImageSet("Image3-" & Format(PageNo, "0"), gStr, CisFun.FE_File)
    
        '---------- 固定項目表示
        For wCnt = 1 To 12
            .CorForm.CrObjects("Box" & Format(wCnt, "0") & "-" & Format(PageNo, "0")).Visible = True
        Next wCnt
        For wCnt = 1 To 13
            .CorForm.CrObjects("Label" & Format(wCnt, "0") & "-" & Format(PageNo, "0")).Visible = True
        Next wCnt
        For wCnt = 1 To 9
            .CorForm.CrObjects("Line" & Format(wCnt, "0") & "-" & Format(PageNo, "0")).Visible = True
        Next wCnt
        Call .CorFixedColorSet("Box1-" & Format(PageNo, "0"), True, CLng(KHT.かんばん枠色))
        Call .CorFixedColorSet("Box2-" & Format(PageNo, "0"), True, CLng(KHT.かんばん枠色))
                    
            
        
    End With
    
    ItemMoveFore = True

End Function


Attribute VB_Name = "CKNZ1230LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   発注かんばん発行
'*
'*   モジュールID  : CKNZ1230LB.BAS
'*
'*                   CKNZ1230LPrint(Integer)
'*
'*
'*   引数 : PrtSy    =  0  :新規発行
'*                      1  :再発行
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:
'    If Not CKNZ1230LPrint(0) Then Exit Sub
'
'*   作成日  : 2008/06/18  By CIS
'*
'*---------------------------------------------------------------------------------
Global CKNZ1230_PrtKB            As Byte
Global CKNZ1230_PrtSy            As Byte
Global CKNZ1230_ImagePath        As String
Global CKNZ1230_CoReport         As New CoReports85        ' クラス定義
Global CKNZ1230_Jisya                 As String
Global CKNZ1230_JisyaNM               As String
Global CKNZ1230_JisyaKJNM             As String

Global CKNZ1230_ColorCode(9)     As String
Global CKNZ1230_Color(9)         As Long
Global CKNZ1230_ColorSetFlg      As Boolean
Global CKNZ1230_BaseColor        As Long

Dim CorItem                     As CrObject
Dim User_BarCode                As String
Dim SetColor                    As Long
Dim EndNo                       As Long
Dim SaveHNo                     As Long

Dim bCnt                        As Long
Dim bStr                        As String

Dim BarHeght                    As Long
Dim OutPrinter                  As String

'+--------------------------------------------+
'+      かんばん色初期内容セット
'+--------------------------------------------+
Public Sub CKNZ1230ColorSet()
' かんばん色取得
    CKNZ1230_BaseColor = 0
    CKNZ1230_ColorSetFlg = True
End Sub
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/
'_/             印　　刷　　処　　理
'_/
'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Function CKNZ1230LPrint(Optional PrtSy As Integer = 0)
    CKNZ1230LPrint = False
    App.Title = "発注かんばん"
' 印刷進行画面表示ﾌﾗｸﾞｾｯﾄ
    pref_MsgDisp = True
'
    gSL_Select = "SELECT  KH.*,ISNULL(SY1.数字1, 14737632) かんばん色,"
    gSL_Select = gSL_Select & "  ZT.最終発行連番,ISNULL(SY2.数字1,0) かんばん枠色"
    gSL_Select = gSL_Select & " , ISNULL(SY3.値名称, '') 自社受入名"
    gSL_Select = gSL_Select & "  FROM かんばん発行テーブル KH"
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 材料マスタ ZM"
    gSL_Select = gSL_Select & "      ON ZM.材料管理番号 = KH.材料管理番号"
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 材料取引先マスタ ZT"
    gSL_Select = gSL_Select & "       ON ZT.材料管理番号 = KH.材料管理番号"
    gSL_Select = gSL_Select & "      AND ZT.取引先 = KH.取引先"
    gSL_Select = gSL_Select & "      AND ZT.受入 = KH.受入"
    gSL_Select = gSL_Select & "      AND ZT.品目 = 1"
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 色設定テーブル IST"
    gSL_Select = gSL_Select & "       ON IST.受入 = KH.受入"
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 名称マスタ SY1"
    gSL_Select = gSL_Select & "       ON SY1.区分名称 = '色設定'"
    gSL_Select = gSL_Select & "      AND SY1.区分タイプ = ''"
    gSL_Select = gSL_Select & "      AND SY1.値 = ISNULL(IST.背景色, '00')"
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 名称マスタ SY2"
    gSL_Select = gSL_Select & "       ON SY2.区分名称 = '色設定'"
    gSL_Select = gSL_Select & "      AND SY2.区分タイプ = ''"
    gSL_Select = gSL_Select & "      AND SY2.値 = ISNULL(IST.枠色, '16')"
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 名称マスタ SY3"
    gSL_Select = gSL_Select & "       ON SY3.区分名称 = '自社工区'"
    gSL_Select = gSL_Select & "      AND SY3.英数字1 = KH.受入"
    gSL_Select = gSL_Select & "      AND SY3.区分タイプ = ''"
    gSL_Select = gSL_Select & " WHERE かんばん種類 = " & CKNZ1230_PrtKB
    gSL_Select = gSL_Select & "   AND 発行種類 = " & CKNZ1230_PrtSy
    If PrtSy = 0 Then
        gSL_Select = gSL_Select & "   AND 発行済 = 0"
    Else
        gSL_Select = gSL_Select & "   AND 発行FLG <> ''"
    End If
    gSL_Select = gSL_Select & " ORDER BY KH.発行種類,KH.取引先,KH.受入,KH.背番号,KH.品番,KH.発行NO"
    
    If Not KHTRead(gSL_Select, 0, 0) Then
        Exit Function
    End If

    pref_TCnt = CisDB.RecordCount

'+++++++++++++++++++++++++++++++'
'+  印刷処理
'+++++++++++++++++++++++++++++++'
    Set pref_Form = CKNZ1211
    pref_Msg = "【 発注かんばん(材料)　印刷中 】"

    pref_Form.Show vbModal
    Unload pref_Form
    Set pref_Form = Nothing

    CKNZ1230LPrint = True
End Function
'+----------------------------------------------------------+
'+                    印  刷  処  理                        ＋
'+----------------------------------------------------------+
Public Function CoPrintN1230() As Boolean
    Dim PageCnt     As Long
    
    CoPrintN1230 = False
    On Error GoTo CoPrintN1230_Er
    
    'かんばん色（基準）
    Call CKNZ1230ColorSet
    '発行元
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 4
    Call IRNRead("", 1, 1)
    Qr_HMoto = RTrim(Mid(IRN.oyian, 1, 7))
    If Trim(Qr_HMoto) = "" Then Qr_HMoto = "SIMPR-A"
    ' 自社コード
    Qr_Jisya = RTrim(gCompany)

    With CisFun
    ' かんばん印刷先
        OutPrinter = ""
        .INI_FileID = gIniExe & gwsIniName
        .INI_Section = "かんばん印刷先"
        .INI_Default = ""
        .INI_Key = CisFun.WSID
        .GetIni
        If Trim(.INI_String) <> "" Then
           OutPrinter = RTrim(.INI_String)
        End If
    End With

    pref_NCnt = 0                ' 現在発行ページ初期化
    
    With CKNZ1230_CoReport
        .SetDraw = pref_Form.CrDraw1

        .SetPrtFormCnt = 1
        .SetPrtForm = "LKN1231"        ' ﾌｫｰﾑIDをｾｯﾄ
        .SetPrtPath = gIniExe & "Prt\" ' ﾌｫｰﾑ格納先
        .SetPrtOut = PrinterX          ' ﾌﾟﾘﾝﾀｰ
        .SetPrtName = "発注かんばん(材料)"                      ' 帳票名
        .SetPrinter = OutPrinter
        .SetPaperSize = A4
        .SetPaperOrientation = 縦
'
        PageCnt = 0
        
        '------ 印刷処理開始
        If Not .CorOutStart Then GoTo CoPrintN1230_Er
        
        ' ﾌｫｰﾑ項目非表示
        Call .CorItemHeide
                
        Do Until Not KHT_RDSTS
            PageCnt = PageCnt + 1
        '-----------------------------------------------    最終発行NOを求める(新規発行時)
            EndNo = 0: SaveHNo = KHT.発行NO
            If KHT.発行区分 = 0 And KHT.発行済 = 0 Then
               EndNo = KHT.発行NO
               gSL_Select = "Select Max(発行NO) 発行NO From かんばん材料管理テーブル "
               gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(KHT.取引先) & "'"
               gSL_Select = gSL_Select & " And   受入 = '" & RTrim(KHT.受入) & "'"
               gSL_Select = gSL_Select & " And   材料管理番号 = '" & RTrim(KHT.材料管理番号) & "'"
               If KZKRead(gSL_Select, 1, 1) Then
                    If KZK.発行NO > KHT.最終発行連番 Then
                       EndNo = KZK.発行NO + 1
                    End If
               End If
                KHT.発行NO = EndNo
            End If
            '------ かんばん材料管理テーブル 版数獲得
             gSL_Select = "Select 版数 From かんばん材料管理テーブル"
             gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(KHT.取引先) & "'"
             gSL_Select = gSL_Select & " And   受入 = '" & RTrim(KHT.受入) & "'"
             gSL_Select = gSL_Select & " And   材料管理番号 = '" & RTrim(KHT.材料管理番号) & "'"
             gSL_Select = gSL_Select & " And   発行NO = " & KHT.発行NO & ""
             If Not KZKRead(gSL_Select, 1, 1) Then
                 KHT.版数 = 1
             Else
                 KHT.版数 = KZK.版数 + 1
            End If
            '------ 項目転送
            If Not ItemMoveFore(PageCnt) Then GoTo CoPrintN1230_Er
            '------ 印刷(１ﾍﾟｰｼﾞ分)
            If PageCnt >= 3 Then
               If Not .CoPrintOut Then GoTo CoPrintN1230_Er
               Call .CorItemHeide
                PageCnt = 0
            End If

            '------ かんばん材料管理テーブル更新
             gSL_Select = "Select * From かんばん材料管理テーブル"
             gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(KHT.取引先) & "'"
             gSL_Select = gSL_Select & " And   受入 = '" & RTrim(KHT.受入) & "'"
             gSL_Select = gSL_Select & " And   材料管理番号 = '" & RTrim(KHT.材料管理番号) & "'"
             gSL_Select = gSL_Select & " And   発行NO = " & KHT.発行NO & ""
             If Not KZKRead(gSL_Select, 1, 1) Then
                Call ItemsClearKZK
                With KZK
                     .取引先 = KHT.取引先
                     .受入 = KHT.受入
                     .材料管理番号 = KHT.材料管理番号
                     .版数 = 1
                     .発行NO = KHT.発行NO
                     .発行区分 = 0
                     .収容区分 = KHT.収容区分
                 End With
                 Call KZKInsert(1)
             Else
                 With KZK
                     .版数 = .版数 + 1
                     .発行区分 = 0
                     .収容区分 = KHT.収容区分
                 End With
                 Call KZKUpdate(1)
            End If
           '--------------------- かんばん発行ワーク更新
            With CisDB
                .ConnectNo = 1
                .SQL = "Update かんばん発行テーブル Set 発行済 = 1"
                If KHT.発行区分 = 0 And KHT.発行済 = 0 Then
                   .SQL = .SQL & " , 発行NO = " & KHT.発行NO
                End If
                .SQL = .SQL & " Where かんばん種類 = " & CKNZ1230_PrtKB
                .SQL = .SQL & " And   品番 = '" & RTrim(KHT.品番) & "'"
                .SQL = .SQL & " And   取引先 = '" & RTrim(KHT.取引先) & "'"
                .SQL = .SQL & " And   受入 = '" & RTrim(KHT.受入) & "'"
                .SQL = .SQL & " And   発行NO = " & SaveHNo
                If Not .DBExec Then GoTo CoPrintN1230_Er
            '--------------------- 材料取引先マスタ　最終発行連番 更新
                If KHT.発行区分 = 0 And KHT.発行済 = 0 Then
                   .SQL = "Update 材料取引先マスタ Set  "
                   .SQL = .SQL & "最終発行連番 = Case When 最終発行連番 > " & KHT.発行NO & " THEN 最終発行連番 Else " & KHT.発行NO & " End,"
                   .SQL = .SQL & "更新日 = GETDATE(),"
                   .SQL = .SQL & "更新者 = '" & gTanto & "',"
                   .SQL = .SQL & "更新端末 = SubString(Host_Name(),1,20)"
                   .SQL = .SQL & " Where 材料管理番号 = '" & RTrim(KHT.材料管理番号) & "'"
                   .SQL = .SQL & "   And 取引先 = '" & RTrim(KHT.取引先) & "'"
                   .SQL = .SQL & "   And 受入 = '" & RTrim(KHT.受入) & "'"
                   If Not .DBExec Then GoTo CoPrintN1230_Er
                End If
            End With

            ' 進捗ゲージ更新
            pref_NCnt = pref_NCnt + 1
            pref_Cnt = Format(pref_NCnt, "###") & " / " & Format(pref_TCnt, "###")
            DoEvents
            '中断処理
            If pref_Cansel_Flg Then
                If MsgBox("処理を中断します。よろしいですか？", 33, "印刷") = 1 Then
                   GoTo CoPrintN1230_ED
                Else
                   pref_Cansel_Flg = False
                End If
            End If
            Call KHTReadNext(0)
        Loop
        
        If PageCnt <> 0 Then
           If Not .CoPrintOut Then GoTo CoPrintN1230_Er
        End If
        
CoPrintN1230_ED:
        Call KHTClose(0)
        
        '------ 印刷終了
        If Not .CoOutEnd Then GoTo CoPrintN1230_Er
    End With
    
    
    CoPrintN1230 = True
    Exit Function
CoPrintN1230_Er:
    Call CisFun.ErrorBox("印刷エラー")
End Function
'*------------------------------------------------------------------*
'*                                                                  *
'*      かんばん　発行内容　項目転送                                *
'*                                                                  *
'*------------------------------------------------------------------*
Private Function ItemMoveFore(PageNo As Long) As Boolean
    ItemMoveFore = False

    With CKNZ1230_CoReport
    '--- 仕入先名
        gStr = CisFun.Mid2(KHT.取引先名, 1, 20)
        If Not .CorData("仕入先名" & Format(PageNo, "0"), gStr, "", 20) Then Exit Function
    '--- 受入
        If Not .CorData("受入" & Format(PageNo, "0"), KHT.受入, "", 2) Then Exit Function
    '--- 所番地
        If Not .CorData("所番地" & Format(PageNo, "0"), KHT.置場, "", 10) Then Exit Function
    '--- 仕入先CD
        If Not .CorData("仕入先CD" & Format(PageNo, "0"), KHT.取引先CD, "", 7) Then Exit Function
    '--- 背番号
        If Not .CorData("背番号" & Format(PageNo, "0"), KHT.背番号, "", 8) Then Exit Function
    '--- 材質
        If KHT.材質 = "" Then
           gStr = Mid(KHT.材質18, 1, 18)
        Else
           gStr = Mid(KHT.材質, 1, 16)
        End If
        If Not .CorData("材質" & Format(PageNo, "0"), gStr, "", 18) Then Exit Function
    '--- 名称
        If Not .CorData("名称" & Format(PageNo, "0"), KHT.材料名, "", 40) Then Exit Function
    '--- 寸法
        If Not .CorData("寸法" & Format(PageNo, "0"), KHT.表示品番, "", 30) Then Exit Function
    '--- 収容数
        gStr = Format(KHT.収容数, "#,##0.000;;#")
        If Not .CorData("収容数" & Format(PageNo, "0"), gStr, "", 11) Then Exit Function
    '--- 発行日
        gStr = Format(Now(), "YYYY/MM/DD")
        If Not .CorData("発行日付" & Format(PageNo, "0"), gStr, "", 10) Then Exit Function
    '--- 発行連番
        gStr = Format(KHT.発行NO, "000")
        If Not .CorData("発行連番" & Format(PageNo, "0"), gStr, "", 3) Then Exit Function
    '--- 最終発行連番
        gStr = Format(KHT.最終発行連番, "000")
        If Not .CorData("最終発行連番" & Format(PageNo, "0"), gStr, "", 3) Then Exit Function
    '--- 版数
        gStr = Format(KHT.版数, "00")
        If Not .CorData("版数" & Format(PageNo, "0"), gStr, "", 2) Then Exit Function
    '--- 荷姿
        If Not .CorData("荷姿" & Format(PageNo, "0"), KHT.収容器, "", 10) Then Exit Function
    
    '--- 社名
        gStr = Mid(KHT.自社名1, 1, 16)
        If Not .CorData("自社名" & Format(PageNo, "0"), gStr, "", 18) Then Exit Function
    '--- 自社受入名
        gStr = Mid(KHT.自社受入名, 1, 16)
        If Not .CorData("工区" & Format(PageNo, "0"), gStr, "", 18) Then Exit Function
    '--- イメージ　ロゴ
        Call .CorImageSet("Image2" & Format(PageNo, "0"), gIniExe & "Bmp", "Logo.jpg")
    '--- イメージ１
'        If RTrim(KHT.IMage) <> "" Then
'            With CisFun
'                .FE_FullPath = KHT.IMage
'                .GetPathName
'                gStr = Replace(.FE_FullPath, .FE_File, "")
'            End With
'            Call .CorImageSet("Image1" & Format(PageNo, "0"), gStr, CisFun.FE_File)
'        End If
        If KHT.かんばん枠色 = 0 Then KHT.かんばん枠色 = CKNZ1230_BaseColor
        If KHT.かんばん色 = 0 Then KHT.かんばん色 = CKNZ1230_BaseColor
        ' QRエリア転送
        Call QRAreaClear
        With QRArea
            .品番 = KHT.品番
            .品目 = KHT.品目
            .取引先 = KHT.取引先
            .受入 = KHT.受入
            .手配区分 = KHT.手配区分
            .版数 = KZK.版数
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
            .品目 = 2
            .品番 = KHT.材質
            .板厚 = KHT.板厚
            .幅 = KHT.幅
            .長さ = KHT.長さ
            .寸法区分 = KHT.寸法区分
            .材料管理番号 = KHT.材料管理番号
            .材料収容数 = KHT.収容数
        End With
        gStr = QR_OutCreate
        
        If Not .CorQRData("QR" & Format(PageNo, "0"), gStr, "", 171) Then Exit Function

       '--- 固定文字・区切り線
        With CKNZ1230_CoReport
            .CorForm.CrObjects("lb_仕入先" & Format(PageNo, "0")).Visible = True
            .CorForm.CrObjects("lb_受入" & Format(PageNo, "0")).Visible = True
            .CorForm.CrObjects("lb_所番地" & Format(PageNo, "0")).Visible = True
            .CorForm.CrObjects("lb_仕入先CD" & Format(PageNo, "0")).Visible = True
            .CorForm.CrObjects("lb_背番号" & Format(PageNo, "0")).Visible = True
            
            .CorForm.CrObjects("lb_材質" & Format(PageNo, "0")).Visible = True
            .CorForm.CrObjects("lb_名称" & Format(PageNo, "0")).Visible = True
            .CorForm.CrObjects("lb_寸法" & Format(PageNo, "0")).Visible = True
            .CorForm.CrObjects("lb_収容数" & Format(PageNo, "0")).Visible = True
            .CorForm.CrObjects("lb_荷姿" & Format(PageNo, "0")).Visible = True
            .CorForm.CrObjects("lb_発行日付" & Format(PageNo, "0")).Visible = True
            .CorForm.CrObjects("lb_発行連番" & Format(PageNo, "0")).Visible = True
            .CorForm.CrObjects("lb_版数" & Format(PageNo, "0")).Visible = True
            .CorForm.CrObjects("lb_Slash" & Format(PageNo, "0")).Visible = True
            
            Call .CorFixedColorSet("BoxA1" & Format(PageNo, "0"), True, KHT.かんばん枠色)
            Call .CorFixedColorSet("BoxA2" & Format(PageNo, "0"), True, KHT.かんばん枠色)
            Call .CorFixedColorSet("BoxA3" & Format(PageNo, "0"), True, KHT.かんばん枠色)
            Call .CorFixedColorSet("BoxA4" & Format(PageNo, "0"), True, KHT.かんばん枠色)
            Call .CorFixedColorSet("BoxA5" & Format(PageNo, "0"), True, KHT.かんばん枠色)
            
            Call .CorFixedColorSet("Line1" & Format(PageNo, "0"), True, KHT.かんばん枠色)
            Call .CorFixedColorSet("Line2" & Format(PageNo, "0"), True, KHT.かんばん枠色)
            Call .CorFixedColorSet("Line4" & Format(PageNo, "0"), True, KHT.かんばん枠色)
            Call .CorFixedColorSet("Line5" & Format(PageNo, "0"), True, KHT.かんばん枠色)
            Call .CorFixedColorSet("Line6" & Format(PageNo, "0"), True, KHT.かんばん枠色)
        '--- 背景色
            .CorForm.CrObjects("Box1" & Format(PageNo, "0")).FillColor = KHT.かんばん色
            .CorForm.CrObjects("Box2" & Format(PageNo, "0")).FillColor = KHT.かんばん色
            .CorForm.CrObjects("Box1" & Format(PageNo, "0")).Visible = True
            .CorForm.CrObjects("Box2" & Format(PageNo, "0")).Visible = True
            
        End With
        
    End With
    
    ItemMoveFore = True

End Function

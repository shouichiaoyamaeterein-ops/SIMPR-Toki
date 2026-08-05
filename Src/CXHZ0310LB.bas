Attribute VB_Name = "CXHZ0310LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   [材料]自社伝票（納品書・受領書・支給書）発行
'*
'*   モジュールID  : CXHZ0310LB.BAS
'*
'*                  CXHZ0310LPrint(Integer)
'*
'*   引数 : Integer  =  0  :新規発行
'*                      1  :再発行
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:
'    If Not CXHZ0310LPrint(1) Then Exit Sub
'
'*   作成日  :2004/02/09   By CIS
'*   変更日  :2004/04/06   By CIS - 自社が取引先マスタに未登録時、クリアを行う。
'*   変更日  :2008/04/11   By CIS - 背番号８桁対応
'*   変更日  :2008/06/06   By CIS - タイトル横にロゴ表示
'*   変更日  :2008/06/18   By CIS - 明細１行中の３行目に発注品番追加
'*---------------------------------------------------------------------------------
Global CXHZ0310Sort          As Byte     ' 印刷順
Global CXHZ0310SortStr       As String   ' 印刷順
Public mQRMeker             As QRmaker  ' QR
Global mUserArea            As String
Global mHonkoFlg            As Boolean
Global mNouHenFlg           As Boolean
Global mBarNaiyo            As Boolean
Global mPage                As Boolean

Dim SyoriKb         As Byte
Dim DP_HMoto        As String
Dim DP_JisyaCd      As String
Dim DP_JisyaKj      As String
Dim DP_JisyaNm      As String
Dim DP_JisyaKjNm    As String
Dim DP_Kiban        As String

Dim pZaist(10)      As String
Dim pSunpo(10)      As String
Dim pZaisy(10)      As String
Public Function CXHZ0310LPrint(Optional RePrint As Integer, Optional PrtMsg As Boolean = False)
    CXHZ0310LPrint = False
    App.Title = "自社伝票"
' 印刷進行画面表示ﾌﾗｸﾞｾｯﾄ
    PrtMsg = True   ' ActiveX使用の為、必ずTrue
    pref_MsgDisp = PrtMsg

'帳票区分 （新規発行 OR 再発行 )
    Select Case RePrint
      Case 0:  SyoriKb = 25
      Case 1:  SyoriKb = 125
    End Select
' 出力データ存在確認
    gSL_Select = "SELECT HC.*,ISNULL(SY1.値名称,'') 発注区分名,ISNULL(SY2.値名称,'') 手配区分名,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.略称,'') 仕入先名,ISNULL(TR2.略称,'') 納入先名,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.取引先,'') 仕入先CD,ISNULL(TR1.工場CD,'') 仕入先工場,ISNULL(TR1.工場名,'') 仕入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR2.取引先,'') 納入先CD,ISNULL(TR2.工場CD,'') 納入先工場,ISNULL(TR2.工場名,'') 納入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR3.取引先,'') 契約取引先,ISNULL(TR3.工場CD,'') 契約工場"
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料発注帳票発行テーブル HC"
    gSL_Select = gSL_Select & "   ON  HC.納品書管理NO =CH.管理NO"
    gSL_Select = gSL_Select & "   AND HC.納入番号 = CH.管理文字"
    gSL_Select = gSL_Select & "   AND HC.カード = CH.検索情報"
    gSL_Select = gSL_Select & "   AND HC.頁NO = CH.補助番号"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY1"
    gSL_Select = gSL_Select & "   ON  SY1.区分名称 = '発注区分'"
    gSL_Select = gSL_Select & "   AND SY1.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY1.値 = HC.発注区分"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2"
    gSL_Select = gSL_Select & "   ON  SY2.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "   AND SY2.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY2.値 = HC.手配区分"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR1"
    gSL_Select = gSL_Select & "   ON  TR1.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR1.取引先CD = HC.仕入先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR2"
    gSL_Select = gSL_Select & "   ON  TR2.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR2.取引先CD = HC.納入先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR3"
    gSL_Select = gSL_Select & "   ON  TR3.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR3.取引先CD = HC.契約先"
    gSL_Select = gSL_Select & " WHERE 端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND 帳票区分 = " & SyoriKb
'    gSL_Select = gSL_Select & "   AND HC.納入番号 = 'A1820'"
'    gSL_Select = gSL_Select & "   AND HC.カード = '7'"
    gSL_Select = gSL_Select & " ORDER BY HC.仕入先,HC.納入番号,HC.発行順,HC.頁NO"
    
    If Not ZCHRead(gSL_Select, 0, 1) Then
       Call ZCHClose(1)
       CXHZ0310LPrint = True
       Exit Function
    End If
    If CisDB.RecordCount = 0 Then
       CXHZ0310LPrint = True
       Exit Function
    End If
    pref_TCnt = CisDB.RecordCount
    
'+++++++++++++++++++++++++++++++'
'+  印刷処理
'+++++++++++++++++++++++++++++++'
    Set pref_Form = CXHZ0310
    pref_Msg = "【 自 社 伝 票  印 刷 中 】"
    
    If pref_MsgDisp Then
        pref_Form.Show vbModal
        Unload pref_Form
        Set pref_Form = Nothing
    Else
        Call CoPrintHZ0310
    End If
    
    CXHZ0310LPrint = True
End Function
'+----------------------------------------------------------+
'+                    印  刷  処  理                        ＋
'+----------------------------------------------------------+
Public Function CoPrintHZ0310() As Boolean
    CoPrintHZ0310 = False
    Set CorForm = pref_Form
    
    pref_NCnt = 0                ' 現在発行ページ初期化
    On Error GoTo CoPrintHZ0310_Err
' システム情報(発行元)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 4
    Call IRNRead("", 1)
    DP_HMoto = Mid(IRN.oyian, 1, 7)
' システム情報(自社工場・機番)
    DP_JisyaCd = RTrim(gCompany)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 7
    Call IRNRead("", 1)
    DP_JisyaKj = Mid(IRN.oyian, 1, 1)
    DP_Kiban = Mid(IRN.oyian2, 1, 2)
' 取引先マスタより名称を求める
    gSL_Select = "SELECT * FROM 取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(gCompany) & Left(IRN.oyian, 1) & "'"
    If TRMRead(gSL_Select, 1) Then
        DP_JisyaNm = TRM.略称
        DP_JisyaKjNm = TRM.工場名
'*---------------------------------------------<< 2004/04/06 Insert Start >>
    Else
        DP_JisyaNm = ""
        DP_JisyaKjNm = ""
'*---------------------------------------------<< 2004/04/06 Insert End >>
    End If
'ﾊﾞｰｺｰﾄﾞ初期ｾｯﾄ
    Call GetBarInfo("JISYADEN")
    Erase CoBarInf
    CoBarInf(0) = -1
    CoBarInf(1) = CowBar_Info(0)
    CoBarInf(2) = CowBar_Info(1)
    CoBarInf(3) = CowBar_Info(2)
    CoBarInf(4) = CowBar_Info(3)
    CoBarInf(5) = CowBar_Info(4)
    CorRtn = CoSetBarInfo(CoBarInf(0), WorkBuffer(0))
    
    If Not CoReportProcH0310("O") Then GoTo CoPrintHZ0310_ED

' 明細を読み込んでｾｯﾄする(ﾚｺｰﾄﾞ件数分)
    Do Until Not ZCH_RDSTS
        ' 項目転送
        If Not CoH0310Move Then GoTo CoPrintHZ0310_ED
        ' 印刷開始
        If Not CoReportProcH0310("P") Then GoTo CoPrintHZ0310_ED
        
        pref_NCnt = pref_NCnt + 1
        If pref_MsgDisp Then
            pref_Cnt = Format(pref_NCnt, "###") & " / " & Format(pref_TCnt, "###")
            DoEvents
        End If
        
        ' 材料発注帳票発行テーブル　発行日更新(新規発行のみ)
        If SyoriKb = 25 Then
            gSL_Select = "UPDATE 材料発注帳票発行テーブル SET "
            gSL_Select = gSL_Select & " 発行日 = CONVERT(VARCHAR(8),GETDATE(),112)"
            gSL_Select = gSL_Select & " WHERE 納品書管理NO = " & ZCH.納品書管理NO & ""
            gSL_Select = gSL_Select & "   AND 納入番号 = '" & ZCH.納入番号 & "'"
            gSL_Select = gSL_Select & "   AND 頁NO = " & ZCH.頁NO & ""
            gSL_Select = gSL_Select & "   AND カード = '" & RTrim(ZCH.カード) & "'"
            With CisDB
                .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
                On Error GoTo CoPrintHZ0310_Err
                .SQL = gSL_Select
                .DBExec
                .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
            End With
        
            gSL_Select = "UPDATE 材料発注テーブル SET "
            Select Case RTrim(ZCH.カード)
                Case "6":
                    gSL_Select = gSL_Select & " 受領書発行日 = CONVERT(VARCHAR(8),GETDATE(),112)"
                Case "7":
                    gSL_Select = gSL_Select & " 納品書発行日 = CONVERT(VARCHAR(8),GETDATE(),112)"
                Case "8":
                    gSL_Select = gSL_Select & " 支給書発行日 = CONVERT(VARCHAR(8),GETDATE(),112)"
            End Select
            gSL_Select = gSL_Select & " WHERE 納品書管理NO = " & ZCH.納品書管理NO & ""
            gSL_Select = gSL_Select & "   AND 納入番号 = '" & ZCH.納入番号 & "'"
            gSL_Select = gSL_Select & "   AND 納品書頁 = " & ZCH.頁NO & ""
            With CisDB
                .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
                On Error GoTo CoPrintHZ0310_Err
                .SQL = gSL_Select
                .DBExec
                .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
            End With
        End If
        
        '次のﾚｺｰﾄﾞを読み込む
        Call ZCHReadNext(1)
    Loop

    CoPrintHZ0310 = True

CoPrintHZ0310_ED:
    Call ZCHClose(1)
' CoReport 終了処理
    If pref_TCnt <> 0 Then Call CoReportProcH0310("E")
    
'    Call PrintSet2("", True)
    Exit Function
CoPrintHZ0310_Err:
    CisFun.ErrorBox
'    Call PrintSet2("", True)
    End
End Function
'*------------------------------------------------------------------*
'*                                                                  *
'*      納品書　発行内容　項目転送                                    *
'*                                                                  *
'*------------------------------------------------------------------*
Private Function CoH0310Move() As Boolean
    Dim wBar            As String
    Dim wBarNaiyo       As String
    
    CoH0310Move = True
    
    On Error GoTo CoH0310Move_Err
    With ZCH
'   *------------------*
'   * 『HEAD』SET  *
'   *------------------*
        Select Case .カード
            Case "6": gStr = "受　領　書"
            Case "7": gStr = "納　品　書"
            Case "8": gStr = "支　給　書"
        End Select
        If Not CorData("伝票種類", gStr, "S", 10) Then GoTo CoH0310Move_Err
        If Mid(.納入番号, 5, 1) <> "0" Then
            If Not CorData("分納区分", "分割", "", 6) Then GoTo CoH0310Move_Err
        End If
        If Not CorData("発注区分", .発注区分名, "", 6) Then GoTo CoH0310Move_Err
        If Not CorData("手配区分", .手配区分名, "", 6) Then GoTo CoH0310Move_Err
        If .SK区分 = 0 Then
            If Not CorData("SK区分", "かんばん", "", 8) Then GoTo CoH0310Move_Err
        Else
            If Not CorData("SK区分", "指示部品", "", 8) Then GoTo CoH0310Move_Err
        End If
        If RTrim(.発行日) <> "" Then
            If Not CorSetPattern(9) Then GoTo CoH0310Move_Err
            If Not CorBox(17500, 200, 19350, 1000, 0) Then GoTo CoH0310Move_Err
            If Not CorData("再発行区分", "再発行", "", 6) Then GoTo CoH0310Move_Err
        End If
    ' 発行日 編集
        With CisFun
            .DateE_BefVal = Format(Now(), "yyyymmdd")
            ZCH.発行日 = RTrim(.DateE_BefVal)
            .DateEdit J_YMD
            gStr = .DateE_AfVal
        End With
        If Not CorData("発行日", gStr, "", 14) Then GoTo CoH0310Move_Err
    '納入先
        If RTrim(.手配区分) <> "1" Then
            If Not CorData("納入先", Left(DP_JisyaCd, 5), "", 5) Then GoTo CoH0310Move_Err
            If Not CorData("納入先名称1", CisFun.Left2(DP_JisyaNm, 20), "", 20) Then GoTo CoH0310Move_Err
            If Not CorData("納入先名称2", CisFun.Left2(DP_JisyaKjNm, 12), "", 12) Then GoTo CoH0310Move_Err
            If Not CorData("納入先事業所1", Left(DP_JisyaKj, 2), "", 2) Then GoTo CoH0310Move_Err
            If RTrim(.受入) <> "" Then
                If Not CorData("納入先事業所2", "-", "", 1) Then GoTo CoH0310Move_Err
                If Not CorData("納入先事業所3", .受入, "", 2) Then GoTo CoH0310Move_Err
            End If
            If Not CorData("支給元", .契約先, "", 5) Then GoTo CoH0310Move_Err
            If Not CorData("支給元1", .契約先, "", 5) Then GoTo CoH0310Move_Err
            If Not CorData("支給元2", "-", "", 1) Then GoTo CoH0310Move_Err
        Else
            If Not CorData("納入先", Left(.納入先CD, 5), "", 5) Then GoTo CoH0310Move_Err
            If Not CorData("納入先名称1", CisFun.Left2(.納入先名, 20), "", 20) Then GoTo CoH0310Move_Err
            If Not CorData("納入先名称2", CisFun.Left2(.納入先工場名, 12), "", 12) Then GoTo CoH0310Move_Err
            If Not CorData("納入先事業所1", Left(.納入先工場, 2), "", 2) Then GoTo CoH0310Move_Err
            If RTrim(.納入先受入) <> "" Then
                If Not CorData("納入先事業所2", "-", "", 1) Then GoTo CoH0310Move_Err
                If Not CorData("納入先事業所3", .納入先受入, "", 2) Then GoTo CoH0310Move_Err
            End If
            If Not CorData("支給元", Left(gCompany, 5), "", 5) Then GoTo CoH0310Move_Err
            If Not CorData("支給元1", Left(gCompany, 5), "", 5) Then GoTo CoH0310Move_Err
            If Not CorData("支給元2", "-", "", 1) Then GoTo CoH0310Move_Err
        End If
        
        If Not CorData("納入場所", "", "", 4) Then GoTo CoH0310Move_Err
        With CisFun
            .DateE_BefVal = ZCH.納入日
            .DateEdit J_YMD
            gStr = .DateE_AfVal
        End With
        If Not CorData("納入指示日", gStr, "", 14) Then GoTo CoH0310Move_Err
        If RTrim(.納入便) <> "" Then
            If Not CorData("納入指示便", RTrim(.納入便) & "便", "", 4) Then GoTo CoH0310Move_Err
        End If
        If Not CorData("機番", DP_Kiban, "", 3) Then GoTo CoH0310Move_Err
        If Not CorData("カード", .カード, "", 1) Then GoTo CoH0310Move_Err
'        If Not CorData("担当者", "", "", 10) Then GoTo CoH0310Move_Err
        
        If Not CorData("納品書NO", .納入番号, "", 5) Then GoTo CoH0310Move_Err
        
    '受注者
        If Not CorData("受注者", Left(.仕入先CD, 5), "", 5) Then GoTo CoH0310Move_Err
        If Not CorData("受注者名称1", CisFun.Left2(.仕入先名, 20), "", 20) Then GoTo CoH0310Move_Err
        If Not CorData("受注者名称2", CisFun.Left2(.仕入先工場名, 12), "", 12) Then GoTo CoH0310Move_Err
        If Not CorData("受注者事業所", Left(.仕入先工場, 2), "", 2) Then GoTo CoH0310Move_Err
        If RTrim(.手配区分) <> "1" And RTrim(.納入先) <> "" Then
            If Not CorData("手配商社", Left(.納入先CD, 5), "", 5) Then GoTo CoH0310Move_Err
            If RTrim(.納入先工場) <> "" Then
                If Not CorData("手配商社2", RTrim(.納入先工場), "", 2) Then GoTo CoH0310Move_Err
                If Not CorData("手配商社3", "-", "", 1) Then GoTo CoH0310Move_Err
            End If
            If Not CorData("手配商社名称1", CisFun.Left2(.納入先名, 20), "", 20) Then GoTo CoH0310Move_Err
            If Not CorData("手配商社名称2", CisFun.Left2(.納入先工場名, 12), "", 12) Then GoTo CoH0310Move_Err
                        
            If Not CorData("手配商社M", "手配商社", "", 8) Then GoTo CoH0310Move_Err
            If Not CorData("手配商社名称M", "手配商社名称", "", 12) Then GoTo CoH0310Move_Err
            If Not CorData("手配商社御中M", "御中", "", 4) Then GoTo CoH0310Move_Err
        End If
    'ボックス設定
    If RTrim(.手配区分) = "1" And RTrim(.カード) = "7" Then
        If Not CorSetPattern(1) Then GoTo CoH0310Move_Err
        '納入先
        If Not CorBox(3000, 2100, 3900, 2500, 0) Then GoTo CoH0310Move_Err
        '納入先事業所
        If Not CorBox(3000, 3500, 3350, 3900, 0) Then GoTo CoH0310Move_Err
        If Not CorBox(3800, 3500, 4250, 3900, 0) Then GoTo CoH0310Move_Err
        '機番
        If Not CorBox(3000, 5500, 3700, 5900, 0) Then GoTo CoH0310Move_Err
        'カード
        If Not CorBox(3000, 6000, 3400, 6400, 0) Then GoTo CoH0310Move_Err
        '納品書NO
        If Not CorBox(3000, 8100, 5000, 8900, 0) Then GoTo CoH0310Move_Err
        '受注者
        If Not CorBox(15000, 2200, 16000, 2600, 0) Then GoTo CoH0310Move_Err
        '受注者事業所
        If Not CorBox(15000, 3600, 15400, 4000, 0) Then GoTo CoH0310Move_Err
    End If
    
    If Not CorImage("Logo", gIniExe & "BMP\Logo.bmp") Then GoTo CoH0310Move_Err        ' 2008/06/06 ADD
    
'   *------------------*
'   * 『BODY』SET  *
'   *------------------*
        Erase pZaist: Erase pSunpo: Erase pZaisy
        For gInt = 1 To .明細行数
            gStr = Format(.行(gInt), "##")
            If Not CorListSet("NO", gStr, gInt, 2) Then GoTo CoH0310Move_Err
            '材料マスタ参照
            gSL_Select = "SELECT ZT.背番号,ZT.板厚,ZT.幅,ZT.長さ,ZT.材質,"
            gSL_Select = gSL_Select & " ZR.表示寸法,ZR.材種コード ,ISNULL(HN.表示品番,'') 品番,"
            gSL_Select = gSL_Select & " CASE WHEN ZR.実数区分 = 1 THEN '*' ELSE '' END 備考,"
            gSL_Select = gSL_Select & " ISNULL(SY1.値名称,'')手配単位名,ISNULL(SY2.値名称,'')材種名"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/18 INSERT START
            'gSL_Select = gSL_Select & ",ISNULL(ZR.発注品番, '') 発注品番"
            gSL_Select = gSL_Select & ",ISNULL(ZT.発注品番, '') 発注品番"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/18 INSERT END
            gSL_Select = gSL_Select & " FROM 材料取引先マスタ ZT"
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料マスタ ZR"
            gSL_Select = gSL_Select & "   ON  ZR.材料管理番号 = ZT.材料管理番号"
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料品番マスタ ZN"
            gSL_Select = gSL_Select & "   ON  ZN.材料管理番号 = ZT.材料管理番号"
            gSL_Select = gSL_Select & "   AND ZN.使用品番 = (SELECT MIN(使用品番) FROM 材料品番マスタ HN2"
            gSL_Select = gSL_Select & "                   WHERE HN2.材料管理番号 = ZT.材料管理番号)"
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
            gSL_Select = gSL_Select & "   ON  HN.品番 = ZN.使用品番"
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY1"
            gSL_Select = gSL_Select & "   ON  SY1.区分名称 = '単位区分'"
            gSL_Select = gSL_Select & "   AND SY1.値 = ZR.手配単位"
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2"
            gSL_Select = gSL_Select & "   ON  SY2.区分名称 = '材種区分'"
            gSL_Select = gSL_Select & "   AND SY2.値 = ZR.材種"
            gSL_Select = gSL_Select & " WHERE ZT.材料管理番号 = " & .材料管理番号(gInt)
            gSL_Select = gSL_Select & "   AND ZT.取引先 = '" & .仕入先 & "'"
            gSL_Select = gSL_Select & "   AND ZT.受入 = '" & .受入 & "'"
            If ZTMRead(gSL_Select, 1) Then
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
                'If Not CorListSet("背番号", ZTM.背番号, gInt, 5) Then GoTo CoH0310Move_Err
                If Not CorListSet("背番号", ZTM.背番号, gInt, 8) Then GoTo CoH0310Move_Err
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
                If Not CorListSet("材質", ZTM.材質, gInt, 25) Then GoTo CoH0310Move_Err
                If Not CorListSet("寸法", ZTM.表示寸法, gInt, 23) Then GoTo CoH0310Move_Err
                gStr = RTrim(ZTM.材種コード) & " " & RTrim(ZTM.材種名)
                gStr = CisFun.Left2(Trim(gStr), 28)
                If Not CorListSet("材種", gStr, gInt, 28) Then GoTo CoH0310Move_Err
                If Not CorListSet("使用品番", ZTM.品番, gInt, 25) Then GoTo CoH0310Move_Err
                If Not CorListSet("単位", ZTM.手配単位名, gInt, 4) Then GoTo CoH0310Move_Err
                If Not CorListSet("単位2", ZTM.手配単位名 & ZTM.備考, gInt, 5) Then GoTo CoH0310Move_Err
                
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/18 INSERT START
                If Not CorListSet("発注品番", ZTM.発注品番, gInt, 25) Then GoTo CoH0310Move_Err
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/18 INSERT END
                
                pZaist(gInt) = ZTM.材質
                pSunpo(gInt) = Replace(Format(ZTM.板厚, "0000.00"), ".", "")
                pSunpo(gInt) = pSunpo(gInt) & Replace(Format(ZTM.幅, "0000.00"), ".", "")
                pSunpo(gInt) = pSunpo(gInt) & Replace(Format(ZTM.長さ, "0000.00"), ".", "")
                pZaisy(gInt) = ZTM.材種コード
            End If
            If .収容数(gInt) <> 0 Then
                gStr = Format(.収容数(gInt), "###,##0.000")
'                gStr = Replace(gStr, ".000", "    ")
                If Not CorListSet("収容数", gStr, gInt, 11) Then GoTo CoH0310Move_Err
            End If
            If .収容数(gInt) <> 0 Then
                gStr = Format(.箱数(gInt), "##,###")
                If Not CorListSet("箱数", gStr, gInt, 6) Then GoTo CoH0310Move_Err
            End If
            If .納入数(gInt) <> 0 Then
                gStr = Format(.納入数(gInt), "###,###,##0.000")
'                gStr = Replace(gStr, ".000", "    ")
                If Not CorListSet("納入数", gStr, gInt, 14) Then GoTo CoH0310Move_Err
            End If
            If Not CorListSet("完納予定1", "／", gInt, 2) Then GoTo CoH0310Move_Err
            If Not CorListSet("完納予定2", "便", gInt, 2) Then GoTo CoH0310Move_Err
            
            If Not CorLine(400, (gInt - 1) * 1500 + 12600, 19400, (gInt - 1) * 1500 + 12600) Then GoTo CoH0310Move_Err
            If Not CorSetPattern(1) Then GoTo CoH0310Move_Err
            If RTrim(.手配区分) = "1" And RTrim(.カード) = "7" Then
                If Not CorBox(1000, (gInt - 1) * 1500 + 11200, 6000, (gInt - 1) * 1500 + 11750, 0) Then GoTo CoH0310Move_Err
                If Not CorBox(10900, (gInt - 1) * 1500 + 11200, 13700, (gInt - 1) * 1500 + 11800, 0) Then GoTo CoH0310Move_Err
            End If
        Next gInt
'   *------------------*
'   * 『TAIL』SET  *
'   *------------------*
        gStr = Format(.箱総数, "##,##0")
        If Not CorData("総箱数", gStr, "", 7) Then GoTo CoH0310Move_Err
        If Not CorData("GROUPNO", .グループNO, "", 3) Then GoTo CoH0310Move_Err
        gStr = Format(.頁NO, "#0")
        If Not CorData("帳票現頁", gStr, "", 2) Then GoTo CoH0310Move_Err
        gStr = Format(.総頁数, "#0")
        If Not CorData("帳票総頁", gStr, "", 2) Then GoTo CoH0310Move_Err
        gStr = Format(.物理現頁数, "#0")
        If Not CorData("納番現頁", gStr, "", 2) Then GoTo CoH0310Move_Err
        gStr = Format(.物理総頁数, "#0")
        If Not CorData("納番総頁", gStr, "", 2) Then GoTo CoH0310Move_Err
        Select Case .カード
            Case "6": gStr = "●"
            Case "7": gStr = "■"
            Case "8": gStr = "▲"
            Case Else: gStr = "☆"
            
        End Select
        If Not CorData("帳票マーク", gStr, "", 2) Then GoTo CoH0310Move_Err
        If .END表示区分 <> 0 Then
            If Not CorSetPattern(10) Then GoTo CoH0310Move_Err
            If Not CorBox(18000, 27200, 19550, 27800, 0) Then GoTo CoH0310Move_Err
            If Not CorData("END区分", "END", "", 6) Then GoTo CoH0310Move_Err
        End If
        If .ALLEND表示区分 <> 0 Then
            If Not CorSetPattern(10) Then GoTo CoH0310Move_Err
            If Not CorBox(18000, 27800, 19550, 28300, 0) Then GoTo CoH0310Move_Err
            If Not CorData("ALLEND区分", "ALLEND", "", 6) Then GoTo CoH0310Move_Err
        End If


'   *---------------------*
'   * 『バーコード』SET    *
'   *---------------------*
        If .納品書管理NO <> 0 Then
            gStr = Format(.納品書管理NO, "0000000000") & Format(.頁NO, "00")
            If CorBar("BAR", gStr) = "ERROR" Then GoTo CoH0310Move_Err
            gStr = "*" & gStr & "*"
            If Not CorData("BAR_Naiyo", gStr, "", 14) Then GoTo CoH0310Move_Err
        End If
'   *---------------------*
'   * 『QR』    SET       *
'   *---------------------*
        Call Denpyo_QREdit
        mQRMeker.InputData = RTrim(CoQR)
        mQRMeker.Refresh
        
        mQRMeker.CreateQrMetaFile CorForm.hDC, gIniExe & "QR.BMP", 2
        If Not CorImage("QR", gIniExe & "QR.BMP") Then Exit Function
    End With

    Exit Function
CoH0310Move_Err:
    CoH0310Move = False
End Function
'*==========================================================================*
'*==========================================================================*
'*------------------------------------------------------------------*
'*                                                                  *
'*      CoReport Proccess           (O)Open   (E)Close              *
'*                                                                  *
'*------------------------------------------------------------------*
Public Function CoReportProcH0310(Repart As String)
    CoReportProcH0310 = False
'   *******************************************************************
'   * フォーム変更処理                              ( Repart = "P" )  *
'   *******************************************************************
    If Repart = "C" Then
'       CorPrtForm = "検査指示2.FED"
'       If CorChange(G_Crprint_Path & CorPrtForm) Then CoReportProcH0310 = True
       Exit Function
    End If
'   *******************************************************************
'   * 印刷指示処理                                  ( Repart = "P" )  *
'   *******************************************************************
    If Repart = "P" Then
        If CorData("", "", "E") Then CoReportProcH0310 = True
    End If
'   *******************************************************************
'   * 終了処理                                      ( Repart = "E" )  *
'   *******************************************************************
    If Repart = "E" Then
        If CorOpCl(crPara, "E") Then CoReportProcH0310 = True
        Exit Function
    End If

'   *****************************************************************
'   * オープン処理    <<ﾌﾟﾘﾝﾀﾄﾞﾗｲﾊﾞ設定・初期処理>> ( Repart = "O" )  *
'   *****************************************************************
    If Repart = "O" Then
' CoReport 初期処理

        With crPara
'           * CoStruct Set *
            .cshwnd = CorForm.hWnd
            .cshdc = Printer.hDC
            .csout = OUTPUT_PRNDRIVER
            .csScale = 100&
            .csmx = 0: .csmy = 0
            .cssx = 0: .cssy = 0
            .cswidth = 0: .csheight = 0
'            .csfname = gIniExe & CorPrtForm & ".FED"
            .csoutfname = ""
'           * DevModeStruct Set *　プリンタの設定
            CorPrtForm = "LXHZ0310"
            .csfname = gIniExe & "Prt\" & CorPrtForm & ".FED"
            .csdevm.dmOrientation = DMORIENT_PORTRAIT   'DMORIENT_LANDSCAPE
'                .csdevm.dmOrientation = DMORIENT_PORTRAIT
            .csdevm.dmPaperSize = DMPAPER_A4
            If pref_PrintNo = 0 Then
               .csdevm.dmFields = DM_TEMPORARY Or DM_PAPERSIZE Or DM_ORIENTATION
            Else
               .csdevm.dmDefaultSource = pref_PrintNo
               .csdevm.dmFields = DM_TEMPORARY Or DM_DEFAULTSOURCE Or DM_PAPERSIZE Or DM_ORIENTATION
            End If
        
        End With
        
        If CorOpCl(crPara, "O") Then
            CoReportProcH0310 = True
        End If

    End If
End Function

'*------------------------------------------------------------------*
'*                                                                  *
'*      自社伝票　QR　作成                                           *
'*                                                                  *
'*------------------------------------------------------------------*
Private Function Denpyo_QREdit() As Boolean
    Dim QREdit1      As String * 107
    Dim QREdit2      As String
    
    QREdit1 = Space(107)
    
    With ZCH
        Mid(QREdit1, 1, 7) = Left(DP_HMoto, 7)
        Mid(QREdit1, 8, 5) = Left(gCompany, 5)
        Mid(QREdit1, 13, 1) = "Z"
        Mid(QREdit1, 14, 1) = "J"
        Mid(QREdit1, 15, 1) = RTrim(.カード)
        Mid(QREdit1, 16, 5) = RTrim(.納入番号)
        Mid(QREdit1, 21, 2) = Format(.頁NO, "00")
        Mid(QREdit1, 23, 10) = Format(.納品書管理NO, "0000000000")
        
        Mid(QREdit1, 33, 5) = Left(.仕入先CD, 4)                       '※受注者（仕入先）
        Mid(QREdit1, 38, 2) = Left(.仕入先工場, 1)                     '受注者事業所（仕入先工場）
        If SyoriKb = 5 Then
            Mid(QREdit1, 40, 1) = "0"           '発行区分(新規発行)
        Else
            Mid(QREdit1, 40, 1) = "1"           '発行区分(再発行)
        End If
        If Mid(.納入番号, 5, 1) = "0" Then
            Mid(QREdit1, 41, 1) = "0"           '分割区分（新規）
        Else
            Mid(QREdit1, 41, 1) = "1"           '分割区分(分納）
        End If
        Mid(QREdit1, 42, 1) = RTrim(.発注区分)           '発注区分
        Mid(QREdit1, 43, 1) = RTrim(.手配区分)           '手配区分
        Mid(QREdit1, 44, 1) = Format(.SK区分, "0")       'SK区分
        If RTrim(.手配区分) <> "1" Then
            Mid(QREdit1, 45, 5) = RTrim(.契約取引先)       '支給元(契約先)
            Mid(QREdit1, 50, 2) = RTrim(.契約工場)         '支給元(契約先工場)
        Else
            Mid(QREdit1, 45, 5) = RTrim(DP_JisyaCd)       '支給元(自社)
            Mid(QREdit1, 50, 2) = RTrim(DP_JisyaKj)       '支給元(自社工場)
        End If
        Mid(QREdit1, 52, 2) = RTrim(.受入)                '納入者事業所(自社受入）
        If RTrim(.手配区分) <> "1" Then
            Mid(QREdit1, 54, 5) = RTrim(DP_JisyaCd)       '納入者(自社)
            Mid(QREdit1, 59, 2) = RTrim(DP_JisyaKj)       '納入者(自社工場)
        Else
            Mid(QREdit1, 54, 5) = RTrim(.納入先CD)         '納入者(納入先)
            Mid(QREdit1, 59, 2) = RTrim(.納入先工場)       '納入者事業所(納入先工場)
            Mid(QREdit1, 61, 2) = RTrim(.納入先受入)       '納入者事業所(納入先工場)
        End If
        If RTrim(.手配区分) <> "1" And RTrim(.納入先CD) <> "" Then
            Mid(QREdit1, 63, 5) = RTrim(.納入先CD)         '手配商社(メーカー)
            Mid(QREdit1, 68, 2) = RTrim(.納入先工場)       '手配商社事業所(メーカー工場)
        End If
        Mid(QREdit1, 70, 8) = RTrim(.納入日)               '納入日
        Mid(QREdit1, 78, 2) = RTrim(.納入便)               '納入便
        Mid(QREdit1, 80, 8) = RTrim(.作成日)               '作成日
        Mid(QREdit1, 88, 3) = RTrim(DP_Kiban)             '機番
        Mid(QREdit1, 91, 4) = Format(.箱総数, "0000")      '箱総数
        Mid(QREdit1, 95, 2) = Format(.総頁数, "00")        '総頁数
        Mid(QREdit1, 97, 3) = RTrim(.グループNO)           'グループNO
        Mid(QREdit1, 100, 2) = Format(.物理現頁数, "00")      '物理現頁数
        Mid(QREdit1, 102, 2) = Format(.物理総頁数, "00")      '物理総頁数
        Mid(QREdit1, 104, 1) = Format(.END表示区分, "0")      'END表示区分
        Mid(QREdit1, 105, 1) = Format(.ALLEND表示区分, "0")   'ALLEND表示区分
        Mid(QREdit1, 106, 2) = Format(.明細行数, "00")        '明細行数
        
        For gInt = 1 To .明細行数
            QREdit2 = QREdit2 & RTrim(pZaist(gInt))                         '材質
            QREdit2 = QREdit2 & Space(25 - CisFun.Len2(pZaist(gInt)))
            QREdit2 = QREdit2 & RTrim(pSunpo(gInt))                         '寸法(000000000000000000)
            QREdit2 = QREdit2 & Space(18 - CisFun.Len2(pSunpo(gInt)))
            QREdit2 = QREdit2 & RTrim(pZaisy(gInt))                         '材種コード
            QREdit2 = QREdit2 & Space(15 - CisFun.Len2(pZaisy(gInt)))
            QREdit2 = QREdit2 & Replace(Format(.収容数(gInt), "000000.000"), ".", "")       '収容数(000000000)
            QREdit2 = QREdit2 & Format(.箱数(gInt), "00000")                                '箱数（枚数）
            QREdit2 = QREdit2 & Replace(Format(.端数(gInt), "000000.000"), ".", "")         '端数(000000000)
            QREdit2 = QREdit2 & Replace(Format(.納入数(gInt), "00000000.000"), ".", "")     '納入数(00000000000)
        Next gInt
    End With
    
    CoQR = RTrim(QREdit1) & RTrim(QREdit2)
    
End Function


Attribute VB_Name = "CXJ0140LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   自社伝票（納品書・受領書・支給書）発行
'*
'*   モジュールID  : CXJ0140LB.BAS
'*
'*                  CXH0310LPrint(Integer)
'*
'*   引数 : Integer  =  0  :新規発行
'*   　　               1  :再発行
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*   　　 : False : 印刷データなし
'*
'*   使用例:
'    If Not CXJ0140LPrint(1) Then Exit Sub
'
'*   作成日  :2004/03/25   By CIS
'*   変更日  :2004/04/06   By CIS - Irnaksic未登録時に異常終了する対応
'*   変更日  :2004/04/16   By CIS - 帳票現頁のバグ対応
'*   変更日  :2004/04/21   By CIS - 受注区分文字数を半角８文字に変更
'*   変更日  :2004/06/24   By CIS - 受領書・出荷時発行の対応
'*   変更日  :2008/04/09   By CIS - 背番号8桁
'*
'*---------------------------------------------------------------------------------
Public mQRMeker             As QRmaker  ' QR
'*--------------------------------------------------<< 2004/06/24 Insert Start >>
Global gCHT_DelKbn53        As Byte
'*--------------------------------------------------<< 2004/06/24 Insert End >>

    Dim SyoriKb         As Byte
    Dim DP_HMoto        As String
    Dim DP_JisyaCd      As String
    Dim DP_JisyaKj      As String
    Dim DP_JisyaNm      As String
    Dim DP_JisyaKjNm    As String
    Dim DP_Kiban        As String
    Dim wNouNo          As String   '>納入番号
    Dim wNKanri         As Long     '>納品書管理NO
    Dim wNowPage        As Byte     '>帳票現頁
    Dim wAllPage        As Byte     '>帳票総頁
    Dim wNowPageT       As Byte     '>納番現頁
    Dim wAllPageT       As Byte     '>納番総頁
    Dim wLineNo         As Integer
    Dim wCard           As String   '>ｶｰﾄﾞ#
    Dim wHakoT          As Long     '>総箱数
'   <<< QR >>>
    Dim QREdit1         As String * 100
    Dim QREdit2         As String

Public Function CXJ0140LPrint(Optional RePrint As Integer, Optional PrtMsg As Boolean = False)
    CXJ0140LPrint = False

    App.Title = "自社伝票"
'   >> 印刷進行画面表示ﾌﾗｸﾞｾｯﾄ
    PrtMsg = True           ' ActiveX使用の為、必ずTrue
    pref_MsgDisp = PrtMsg

'   >> 帳票区分（新規発行 OR 再発行 )
    Select Case RePrint
      Case 0:  SyoriKb = 53
      Case 1:  SyoriKb = 153
    End Select
'   >> 出力データ存在確認
    Call SqlSelect_Group
    If Not JNWRead(gSL_Select, 0, 1) Then
       Call JNWClose(1)
       CXJ0140LPrint = True
       Exit Function
    End If
    pref_TCnt = CisDB.RecordCount

'+++++++++++++++++++++++++++++++'
'+  印刷処理
'+++++++++++++++++++++++++++++++'
    Set pref_Form = CXJ0140
    pref_Msg = "【 自 社 伝 票  印 刷 中 】"

    If pref_MsgDisp Then
        pref_Form.Show vbModal
        Unload pref_Form
        Set pref_Form = Nothing
    Else
        Call CoPrintJ0140
    End If

    CXJ0140LPrint = True
End Function
'+----------------------------------------------------------+
'+                    印  刷  処  理                        ＋
'+----------------------------------------------------------+
Public Function CoPrintJ0140() As Boolean
    CoPrintJ0140 = False

    On Error GoTo CoPrintJ0140_Err

    Set CorForm = pref_Form
    pref_NCnt = 0                ' 現在発行ページ初期化

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
'*--------------------------------------------------<< 2004/04/06 Insert Start >>
    Else
        DP_JisyaNm = ""
        DP_JisyaKjNm = ""
'*--------------------------------------------------<< 2004/04/06 Insert End >>
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

    If Not CoReportProcJ0140("O") Then GoTo CoPrintJ0140_Err

'   <><><><><> 納入番号の読込み <><><><><>
    Do Until Not JNW_RDSTS
        With JNW
            wNouNo = .納入番号
            wNKanri = .納品書管理NO
            wNowPage = .納品書頁
            wAllPage = .帳票総頁
            wCard = .帳票種類
        End With
        Call SetPageCount

        Call SqlSelect_Print
        Call JYTRead(gSL_Select, 0, 0)
    '   <><><><><> 明細を読み込んでｾｯﾄする(ﾚｺｰﾄﾞ件数分) <><><><><>
        Do Until Not JYT_RDSTS
        '   >> HEAD部 項目転送
            wHakoT = 0
            If Not HeadMove Then GoTo CoPrintJ0140_Err
        '   <><><> １ページ分実行 <><><>
            While JYT_RDSTS And JYT.納入番号 = wNouNo _
                            And JYT.納品書頁 = wNowPage
            '   >> BODY部 項目転送
                wHakoT = wHakoT + JYT.枚数
                wLineNo = JYT.納品書行
                If Not BodyMove(wLineNo) Then GoTo CoPrintJ0140_Err

            '   >> 次のﾚｺｰﾄﾞを読み込む
                Call JYTReadNext(0)
            Wend
        '   >> TAIL部 項目転送
            If Not TailMove Then GoTo CoPrintJ0140_Err
        '   >> 印刷開始
            If Not CoReportProcJ0140("P") Then GoTo CoPrintJ0140_Err
        Loop
        Call JYTClose(0)
        
        pref_NCnt = pref_NCnt + 1
        If pref_MsgDisp Then
            pref_Cnt = Format(pref_NCnt, "###") & " / " & Format(pref_TCnt, "###")
            DoEvents
        End If

'       >> 次のﾚｺｰﾄﾞを読み込む
        Call JNWReadNext(1)
    Loop
    Call JNWClose(1)

'   >> CoReport 終了処理
    If pref_TCnt <> 0 Then Call CoReportProcJ0140("E")
    
    CoPrintJ0140 = True
    Exit Function
CoPrintJ0140_Err:
    CisFun.ErrorBox
    End
End Function
'+++++++++++++++++++++++++++++++'
'+      ページ数のセット
'+++++++++++++++++++++++++++++++'
Private Sub SetPageCount()
    Dim wDenMax         As Byte
    Dim wDenCnt         As Byte

'*--------------------------------------------------<< 2004/06/24 Update Start >>
'    wDenMax = 0
'    wDenCnt = 0
'    If JNW.受領書形態 = "1" Then
'        If wCard = "6" Then wDenCnt = wDenMax
'        wDenMax = wDenMax + 1
'    End If
'    If JNW.納品書形態 = "1" Then
'        If wCard = "7" Then wDenCnt = wDenMax
'        wDenMax = wDenMax + 1
'    End If
'    If CisFun.Mid2(JNW.納入先, 1, 4) <> CisFun.Mid2(JNW.契約先, 1, 4) Then
'        If JNW.納品書形態 = "1" Then
'            If wCard = "8" Then wDenCnt = wDenMax
'            wDenMax = wDenMax + 1                               '>2004/04/17 Insert
'        End If
''       wDenMax = wDenMax + 1                                   '>2004/04/17 Delete
'    End If
'    wAllPageT = wAllPage * wDenMax
''*--------------------------------------------------<< 2004/04/16 Update Start >>
''   wNowPageT = wAllPage * wDenCnt + 1
'    wNowPageT = wAllPage * wDenCnt + wNowPage
''*--------------------------------------------------<< 2004/04/16 Update End >>
    If wCard = "6" And _
       JNW.受領書形態 = "1" And JNW.受領書発行 = 1 Then
        wDenMax = 1
        wDenCnt = 0
    Else
        wDenMax = 0
        wDenCnt = 0
        If JNW.受領書形態 = "1" And JNW.受領書発行 = 0 Then
            If wCard = "6" Then wDenCnt = wDenMax
            wDenMax = wDenMax + 1
        End If
        If JNW.納品書形態 = "1" Then
            If wCard = "7" Then wDenCnt = wDenMax
            wDenMax = wDenMax + 1
        End If
        If CisFun.Mid2(JNW.納入先, 1, 4) <> CisFun.Mid2(JNW.契約先, 1, 4) Then
            If JNW.納品書形態 = "1" Then
                If wCard = "8" Then wDenCnt = wDenMax
                wDenMax = wDenMax + 1
            End If
        End If
    End If
    wAllPageT = wAllPage * wDenMax
    wNowPageT = wAllPage * wDenCnt + wNowPage
'*--------------------------------------------------<< 2004/06/24 Update End >>

End Sub
'*==========================================================================*
'   *----------------------------*
'   *   納入番号単位用 SQL文
'   *----------------------------*
Private Function SqlSelect_Group() As Boolean
    SqlSelect_Group = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  jyt.納入番号"
    gSL_Select = gSL_Select & ", jyt.納品書管理NO"
    gSL_Select = gSL_Select & ", jyt.納品書頁"
    gSL_Select = gSL_Select & ", jyt.納入先"
    gSL_Select = gSL_Select & ", cht.検索情報  帳票種類"
    gSL_Select = gSL_Select & ", ("
    gSL_Select = gSL_Select & "select max(xxx.納品書頁)"
    gSL_Select = gSL_Select & "  from 受注テーブル xxx"
    gSL_Select = gSL_Select & " where xxx.納品書管理NO = jyt.納品書管理NO"
    gSL_Select = gSL_Select & "   and xxx.納品書発行区分 = 1"
    gSL_Select = gSL_Select & "  ) 帳票総頁"
    gSL_Select = gSL_Select & ", isnull(htm.納品書形態, '') 納品書形態"
    gSL_Select = gSL_Select & ", isnull(htm.受領書形態, '') 受領書形態"
    gSL_Select = gSL_Select & ", isnull(htm.契約先, '') 契約先"
'*--------------------------------------------------<< 2004/06/24 Insert Start >>
    gSL_Select = gSL_Select & ", isnull(htm.受領書発行, 0) 受領書発行"
'*--------------------------------------------------<< 2004/06/24 Insert End >>
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  帳票出力テーブル cht"
    gSL_Select = gSL_Select & ", 受注テーブル jyt"
    gSL_Select = gSL_Select & "  left outer join 品番取引先マスタ htm"
    gSL_Select = gSL_Select & "    on htm.品番 = jyt.品番"
    gSL_Select = gSL_Select & "   and htm.取引先 = jyt.納入先"
    gSL_Select = gSL_Select & "   and htm.受入 = jyt.受入"
    gSL_Select = gSL_Select & "   and htm.品目 = 0"
    gSL_Select = gSL_Select & " where cht.端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and cht.帳票区分 = " & SyoriKb
    gSL_Select = gSL_Select & "   and jyt.納入番号 = cht.管理文字"
    gSL_Select = gSL_Select & "   and jyt.納品書管理NO = cht.管理NO"
    gSL_Select = gSL_Select & "   and jyt.納品書頁 = cht.補助番号"
    gSL_Select = gSL_Select & "   and jyt.納品書行 = 1"
    gSL_Select = gSL_Select & "   and jyt.納品書発行区分 = 1"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  cht.管理文字"      '>納入番号
    gSL_Select = gSL_Select & ", cht.検索情報"      '>伝票種類（受領書・納品書・支給書）
    gSL_Select = gSL_Select & ", cht.補助番号"      '>納品書頁

    SqlSelect_Group = True
End Function
'*==========================================================================*
'   *----------------------------*
'   *   納入番号内印刷用 SQL文
'   *----------------------------*
Private Function SqlSelect_Print() As Boolean
    SqlSelect_Print = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  jyt.*"
    gSL_Select = gSL_Select & ", isnull(hnm.表示品番, jyt.品番) 表示品番"
    gSL_Select = gSL_Select & ", isnull(hnm.品名, '') 品名"
    gSL_Select = gSL_Select & ", isnull(hnm.実数区分, 0) 実数区分"
    gSL_Select = gSL_Select & ", isnull(trm.取引先, '') 取引先CD"
    gSL_Select = gSL_Select & ", isnull(trm.工場CD, '') 工場CD"
    gSL_Select = gSL_Select & ", isnull(trm.略称, '') 納入先名"
    gSL_Select = gSL_Select & ", isnull(trm.工場名, '') 工場名"
    gSL_Select = gSL_Select & ", isnull(htm.契約先, '') 契約先"
    gSL_Select = gSL_Select & ", isnull(htm.SK区分, 0) SK区分"
'>  gSL_Select = gSL_Select & ", isnull(htm.手配区分, '') 手配区分"
    gSL_Select = gSL_Select & ", isnull(sy1.値名称, '') 受注区分名"
'>  gSL_Select = gSL_Select & ", isnull(sy2.値名称, '') 手配区分名"
    gSL_Select = gSL_Select & ", isnull(tr2.取引先, '') 契約取引先"
    gSL_Select = gSL_Select & ", isnull(tr2.工場CD, '') 契約工場"
'*--------------------------------------------------<< 2004/06/24 Insert Start >>
    If wCard = "6" And _
       JNW.受領書形態 = "1" And JNW.受領書発行 = 1 Then
        gSL_Select = gSL_Select & ", skt.出荷数 納入数"
        gSL_Select = gSL_Select & ", skt.出荷枚数 枚数"
    End If
'*--------------------------------------------------<< 2004/06/24 Insert End >>
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  受注テーブル jyt"
    gSL_Select = gSL_Select & "  left outer join 品番マスタ hnm"
    gSL_Select = gSL_Select & "    on hnm.品番 = jyt.品番"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ trm"
    gSL_Select = gSL_Select & "    on trm.取引先CD = jyt.納入先"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 0"
    gSL_Select = gSL_Select & "  left outer join 品番取引先マスタ htm"
    gSL_Select = gSL_Select & "    on htm.品番 = jyt.品番"
    gSL_Select = gSL_Select & "   and htm.取引先 = jyt.納入先"
    gSL_Select = gSL_Select & "   and htm.受入 = jyt.受入"
    gSL_Select = gSL_Select & "   and htm.品目 = 0"
    gSL_Select = gSL_Select & "  left outer join 名称マスタ sy1"
    gSL_Select = gSL_Select & "    on sy1.区分名称 = '受注区分'"
    gSL_Select = gSL_Select & "   and sy1.区分タイプ = ''"
    gSL_Select = gSL_Select & "   and sy1.値 = jyt.受注区分"
'>  gSL_Select = gSL_Select & "  left outer join 名称マスタ sy2"
'>  gSL_Select = gSL_Select & "    on sy2.区分名称 = '手配区分'"
'>  gSL_Select = gSL_Select & "   and sy2.区分タイプ = ''"
'>  gSL_Select = gSL_Select & "   and sy2.値 = htm.手配区分"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ tr2"
    gSL_Select = gSL_Select & "    on tr2.取引先CD = htm.契約先"
    gSL_Select = gSL_Select & "   and tr2.取引先区分 = 0"
'*--------------------------------------------------<< 2004/06/24 Insert Start >>
    If wCard = "6" And _
       JNW.受領書形態 = "1" And JNW.受領書発行 = 1 Then
        gSL_Select = gSL_Select & ", 出荷テーブル skt"
    End If
'*--------------------------------------------------<< 2004/06/24 Insert End >>
    gSL_Select = gSL_Select & " where jyt.納入番号 = '" & wNouNo & "'"
    gSL_Select = gSL_Select & "   and jyt.納品書管理NO = '" & wNKanri & "'"
    gSL_Select = gSL_Select & "   and jyt.納品書頁 = '" & wNowPage & "'"
    gSL_Select = gSL_Select & "   and jyt.納品書発行区分 = 1"
'*--------------------------------------------------<< 2004/06/24 Insert Start >>
    If wCard = "6" And _
       JNW.受領書形態 = "1" And JNW.受領書発行 = 1 Then
        gSL_Select = gSL_Select & "   and skt.出荷指示書NO = jyt.出荷指示書NO"
        gSL_Select = gSL_Select & "   and skt.出荷指示書行NO = jyt.出荷指示書行NO"
    End If
'*--------------------------------------------------<< 2004/06/24 Insert End >>
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  jyt.納品書頁"
    gSL_Select = gSL_Select & ", jyt.納品書行"

    SqlSelect_Print = True
End Function
'*------------------------------------------------------------------*
'*                                                                  *
'*      納品書　発行内容　項目転送                                    *
'*                                                                  *
'*------------------------------------------------------------------*
'   *------------------*
'   *   『HEAD部』SET
'   *------------------*
Private Function HeadMove() As Boolean
    HeadMove = False

    On Error GoTo HeadMove_Err

    With JYT
        Select Case wCard
            Case "6": gStr = "受　領　書"
            Case "7": gStr = "納　品　書"
            Case "8": gStr = "支　給　書"
        End Select
        If Not CorData("伝票種類", gStr, "S", 10) Then Exit Function
        If Mid(.納入番号, 5, 1) <> "0" Then
            If Not CorData("分納区分", "分割", "", 6) Then Exit Function
        End If
        If Not CorData("発注区分", .受注区分名, "", 8) Then Exit Function
'>      If Not CorData("手配区分", .手配区分名, "", 6) Then Exit Function
        If .SK区分 = 0 Then
            If Not CorData("SK区分", "かんばん", "", 8) Then Exit Function
        Else
            If Not CorData("SK区分", "指示部品", "", 8) Then Exit Function
        End If
        If SyoriKb = 153 Then
            If Not CorSetPattern(9) Then Exit Function
            If Not CorBox(17500, 200, 19350, 1000, 0) Then Exit Function
            If Not CorData("再発行区分", "再発行", "", 6) Then Exit Function
        End If
    '   >> 発行日 編集
        With CisFun
            .DateE_BefVal = Format(Date, "yyyymmdd")
            .DateEdit J_YMD
            gStr = .DateE_AfVal
        End With
        If Not CorData("発行日", gStr, "", 14) Then Exit Function
    '   >> 納入先
        If Not CorData("納入先", .取引先CD, "", 5) Then Exit Function
        If Not CorData("納入先名称1", .納入先名, "", 20) Then Exit Function
        If Not CorData("納入先名称2", .工場名, "", 16) Then Exit Function
        If Not CorData("納入先事業所1", .工場CD, "", 2) Then Exit Function
        If RTrim(.受入) <> "" Then
            If Not CorData("納入先事業所2", "-", "", 1) Then Exit Function
            If Not CorData("納入先事業所3", .受入, "", 2) Then Exit Function
        End If
'>      If Not CorData("納入場所", "", "", 4) Then Exit Function
        With CisFun
            .DateE_BefVal = JYT.納入日
            .DateEdit J_YMD
            gStr = .DateE_AfVal
        End With
        If Not CorData("納入指示日", gStr, "", 14) Then Exit Function
        If RTrim(.納入便) <> "" Then
            If Not CorData("納入指示便", RTrim(.納入便) & "便", "", 4) Then Exit Function
        End If
        If Not CorData("機番", DP_Kiban, "", 3) Then Exit Function
        If Not CorData("カード", wCard, "", 1) Then Exit Function
'>      If Not CorData("担当者", "", "", 10) Then Exit Function

        If Not CorData("支給元", .契約先, "", 5) Then Exit Function
        If Not CorData("支給元1", .契約先, "", 5) Then Exit Function
        If Not CorData("支給元2", "-", "", 1) Then Exit Function

        If Not CorData("納品書NO", .納入番号, "", 5) Then Exit Function

    '   >> 受注者
        If Not CorData("受注者", DP_JisyaCd, "", 5) Then Exit Function
        If Not CorData("受注者名称1", DP_JisyaNm, "", 20) Then Exit Function
        If Not CorData("受注者名称2", DP_JisyaKjNm, "", 16) Then Exit Function
        If Not CorData("受注者事業所", DP_JisyaKj, "", 2) Then Exit Function
    '   >> ボックス設定
'>      If RTrim(.手配区分) = "1" And wCard = "7" Then
'>          If Not CorSetPattern(1) Then Exit Function
'>      '   納入先
'>          If Not CorBox(3000, 2100, 3950, 2500, 0) Then Exit Function
'>      '   納入先事業所
'>          If Not CorBox(3000, 3500, 3450, 3900, 0) Then Exit Function
'>          If Not CorBox(3750, 3500, 4200, 3900, 0) Then Exit Function
'>      '   機番
'>          If Not CorBox(3000, 5500, 3700, 5900, 0) Then Exit Function
'>      '   カード
'>          If Not CorBox(3000, 6000, 3400, 6400, 0) Then Exit Function
'>      '   納品書NO
'>          If Not CorBox(3000, 8100, 5000, 8900, 0) Then Exit Function
'>      '   受注者
'>          If Not CorBox(15000, 2200, 16000, 2600, 0) Then Exit Function
'>      '   受注者事業所
'>          If Not CorBox(15000, 3600, 15400, 4000, 0) Then Exit Function
'>      End If
'       <><><><><> 『 バーコード 』SET <><><><><>
        If .納品書管理NO <> 0 Then
            gStr = Format(.納品書管理NO, "0000000000") & Format(.納品書頁, "00")
            If CorBar("BAR", gStr) = "ERROR" Then Exit Function
            gStr = "*" & gStr & "*"
            If Not CorData("BAR_Naiyo", gStr, "", 14) Then Exit Function
        End If
    End With
'   <><><><><> 『 QR 』HEAD SET <><><><><>
    Call QREditHead

    HeadMove = True
    Exit Function
HeadMove_Err:
    CisFun.ErrorBox
    End
End Function
'   *------------------*
'   *   『BODY部』SET
'   *------------------*
Private Function BodyMove(Index As Integer) As Boolean
    BodyMove = False

    On Error GoTo BodyMove_Err

    With JYT
        gStr = Format(.納品書行, "##")
        If Not CorListSet("NO", gStr, Index, 2) Then Exit Function
        If Not CorListSet("品番", .表示品番, Index, 30) Then Exit Function
        If Not CorListSet("品名", .品名, Index, 32) Then Exit Function
        '*-*-*-*-* 2008/04/09 Update Start
        'If Not CorListSet("背番号", .背番号, Index, 5) Then Exit Function
        If Not CorListSet("背番号", .背番号, Index, 8) Then Exit Function
        '*-*-*-*-* 2008/04/09 Update End
                
        gStr = Format(.収容数, "##")
        If Not CorListSet("収容数", gStr, Index, 7) Then Exit Function
        gStr = Format(.枚数, "##")
        If Not CorListSet("箱数", gStr, Index, 6) Then Exit Function
        gStr = Format(.納入数, "##")
        If Not CorListSet("納入数", gStr, Index, 7) Then Exit Function
        Select Case .実数区分
            Case 0:     gStr = ""
            Case Else:  gStr = "*"
        End Select
        If Not CorListSet("分割個数", "個" & gStr, Index, 3) Then Exit Function
        If Not CorListSet("完納予定1", "／", Index, 2) Then Exit Function
        If Not CorListSet("完納予定2", "便", Index, 2) Then Exit Function
'
        If Not CorLine(400, (Index - 1) * 1500 + 12600, 19400, (Index - 1) * 1500 + 12600) Then Exit Function
        If Not CorSetPattern(1) Then Exit Function
'>      If RTrim(.手配区分) = "1" And wCard = "7" Then
'>          If Not CorBox(1000, (Index - 1) * 1500 + 11200, 6900, (Index - 1) * 1500 + 11750, 0) Then Exit Function
'>          If Not CorBox(13300, (Index - 1) * 1500 + 11200, 15000, (Index - 1) * 1500 + 11750, 0) Then Exit Function
'>      End If
    End With
'   <><><><><> 『 QR 』BODY SET <><><><><>
    Call QREditBody

    BodyMove = True
    Exit Function
BodyMove_Err:
    CisFun.ErrorBox
    End
End Function
'   *------------------*
'   *   『TAIL部』SET
'   *------------------*
Private Function TailMove() As Boolean
    TailMove = False

    On Error GoTo TailMove_Err

    gStr = Format(wHakoT, "##,##0")
    If Not CorData("総箱数", gStr, "", 7) Then Exit Function
    gStr = CisFun.Mid2(wNouNo, 2, 3)
    If Not CorData("GROUPNO", gStr, "", 3) Then Exit Function
    gStr = Format(wNowPage, "#0")
    If Not CorData("帳票現頁", gStr, "", 2) Then Exit Function
    gStr = Format(wAllPage, "#0")
    If Not CorData("帳票総頁", gStr, "", 2) Then Exit Function
    gStr = Format(wNowPageT, "#0")
    If Not CorData("納番現頁", gStr, "", 2) Then Exit Function
    gStr = Format(wAllPageT, "#0")
    If Not CorData("納番総頁", gStr, "", 2) Then Exit Function
    Select Case wCard
        Case "6": gStr = "●"
        Case "7": gStr = "■"
        Case "8": gStr = "▲"
        Case Else: gStr = "☆"
    End Select
    If Not CorData("帳票マーク", gStr, "", 2) Then Exit Function
    If wNowPage = wAllPage Then
        If Not CorSetPattern(10) Then Exit Function
        If Not CorBox(18000, 27200, 19550, 27800, 0) Then Exit Function
        If Not CorData("END区分", "END", "", 6) Then Exit Function
    End If
    If wNowPageT = wAllPageT Then
        If Not CorSetPattern(10) Then Exit Function
        If Not CorBox(18000, 27800, 19550, 28300, 0) Then Exit Function
        If Not CorData("ALLEND区分", "ALLEND", "", 6) Then Exit Function
    End If
'   <><><><><> 『 QR 』TAIL SET <><><><><>
    Call QREditTail

'   !!__________ 『 QR 』SET __________!!
    CoQR = RTrim(QREdit1) & RTrim(QREdit2)
    mQRMeker.InputData = RTrim(CoQR)
    mQRMeker.Refresh
    mQRMeker.CreateQrMetaFile CorForm.hDC, gIniExe & "QR.BMP", 2
    If Not CorImage("QR", gIniExe & "QR.BMP") Then Exit Function

    TailMove = True
    Exit Function
TailMove_Err:
    CisFun.ErrorBox
    End
End Function
'*==========================================================================*
'*==========================================================================*
'*------------------------------------------------------------------*
'*                                                                  *
'*      CoReport Proccess           (O)Open   (E)Close              *
'*                                                                  *
'*------------------------------------------------------------------*
Public Function CoReportProcJ0140(Repart As String)
    CoReportProcJ0140 = False
'   *******************************************************************
'   * フォーム変更処理                              ( Repart = "P" )  *
'   *******************************************************************
    If Repart = "C" Then
       'CorPrtForm = "LXJ0140.FED"
       'If CorChange(G_Crprint_Path & CorPrtForm) Then CoReportProcJ0140 = True
        Exit Function
    End If
'   *******************************************************************
'   * 印刷指示処理                                  ( Repart = "P" )  *
'   *******************************************************************
    If Repart = "P" Then
        If CorData("", "", "E") Then CoReportProcJ0140 = True
    End If
'   *******************************************************************
'   * 終了処理                                      ( Repart = "E" )  *
'   *******************************************************************
    If Repart = "E" Then
        If CorOpCl(crPara, "E") Then CoReportProcJ0140 = True
        Exit Function
    End If

'   *****************************************************************
'   * オープン処理    <<ﾌﾟﾘﾝﾀﾄﾞﾗｲﾊﾞ設定・初期処理>> ( Repart = "O" )  *
'   *****************************************************************
    If Repart = "O" Then
'   >> CoReport 初期処理

        With crPara
'           * CoStruct Set *
            .cshwnd = CorForm.hWnd
            .cshdc = Printer.hDC
            .csout = OUTPUT_PRNDRIVER
            .csScale = 100&
            .csmx = 0: .csmy = 0
            .cssx = 0: .cssy = 0
            .cswidth = 0: .csheight = 0
            .csoutfname = ""
'           * DevModeStruct Set *　プリンタの設定
            CorPrtForm = "LXJ0140"
            .csfname = gIniExe & "Prt\" & CorPrtForm & ".FED"
            .csdevm.dmOrientation = DMORIENT_PORTRAIT   'DMORIENT_LANDSCAPE
            .csdevm.dmPaperSize = DMPAPER_A4
            If pref_PrintNo = 0 Then
               .csdevm.dmFields = DM_TEMPORARY Or DM_PAPERSIZE Or DM_ORIENTATION
            Else
               .csdevm.dmDefaultSource = pref_PrintNo
               .csdevm.dmFields = DM_TEMPORARY Or DM_DEFAULTSOURCE Or DM_PAPERSIZE Or DM_ORIENTATION
            End If
        End With

        If CorOpCl(crPara, "O") Then
            CoReportProcJ0140 = True
        End If

    End If
End Function

'*------------------------------------------------------------------*
'*                                                                  *
'*      自社伝票　QR　作成                                           *
'*                                                                  *
'*------------------------------------------------------------------*
'   *------------------*
'   *   『HEAD部』
'   *------------------*
Private Function QREditHead() As Boolean
    QREditHead = False

    On Error GoTo QREditHead_Err

    QREdit1 = Space(100)
    QREdit2 = ""

    With JYT
        Mid(QREdit1, 1, 7) = Left(DP_HMoto, 7)
        Mid(QREdit1, 8, 5) = Left(gCompany, 5)
        Mid(QREdit1, 13, 1) = "D"
        Mid(QREdit1, 14, 1) = "J"
        Mid(QREdit1, 15, 1) = wCard
        Mid(QREdit1, 16, 5) = RTrim(.納入番号)
        Mid(QREdit1, 21, 2) = Format(.納品書頁, "00")
        Mid(QREdit1, 23, 10) = Format(.納品書管理NO, "0000000000")
'
        Mid(QREdit1, 33, 5) = Left(DP_JisyaCd, 4)       '※受注者（自社）
        Mid(QREdit1, 38, 2) = Left(DP_JisyaKj, 1)       '受注者事業所（自社工場）
        Select Case SyoriKb
            Case 53:    Mid(QREdit1, 40, 1) = "0"       '発行区分(新規発行)
            Case 153:   Mid(QREdit1, 40, 1) = "1"       '発行区分(再発行)
        End Select
        Select Case CisFun.Mid2(.納入番号, 5, 1)
            Case "0":   Mid(QREdit1, 41, 1) = "0"       '分割区分(新規)
            Case Else:  Mid(QREdit1, 41, 1) = "1"       '分割区分(分納)
        End Select
        Mid(QREdit1, 42, 1) = RTrim(.受注区分)          '受注区分
        Mid(QREdit1, 43, 1) = "0"                       '手配区分
        Mid(QREdit1, 44, 1) = Format(.SK区分, "0")      'SK区分
        Mid(QREdit1, 45, 5) = RTrim(.契約取引先)        '支給元(契約先)
        Mid(QREdit1, 50, 2) = RTrim(.契約工場)          '支給元(契約先工場)
        Mid(QREdit1, 52, 2) = RTrim(.受入)              '納入者事業所(受入）
        Mid(QREdit1, 54, 5) = RTrim(.取引先CD)          '納入者(納入先)
        Mid(QREdit1, 59, 2) = RTrim(.工場CD)            '納入者事業所(工場)
        Mid(QREdit1, 61, 2) = RTrim(.受入)              '納入者事業所(受入)
        Mid(QREdit1, 63, 8) = RTrim(.納入日)            '納入日
        Mid(QREdit1, 71, 2) = RTrim(.納入便)            '納入便
        Mid(QREdit1, 73, 8) = Format(Date, "yyyymmdd")  '作成日
        Mid(QREdit1, 81, 3) = RTrim(DP_Kiban)           '機番
    End With

    QREditHead = True
    Exit Function
QREditHead_Err:
    CisFun.ErrorBox
    End
End Function
'   *------------------*
'   *   『BODY部』
'   *------------------*
Private Function QREditBody() As Boolean
    QREditBody = False

    On Error GoTo QREditBody_Err

    With JYT
        QREdit2 = QREdit2 & .背番号                             '背番号
        '*-*-*-*-* 2008/04/09 Update Start
        'QREdit2 = QREdit2 & Space(5 - CisFun.Len2(.背番号))
        QREdit2 = QREdit2 & Space(8 - CisFun.Len2(.背番号))
        '*-*-*-*-* 2008/04/09 Update End
        QREdit2 = QREdit2 & .品番                               '品番
        QREdit2 = QREdit2 & Space(25 - CisFun.Len2(.品番))
        QREdit2 = QREdit2 & Format(.収容数, "000000")           '収容数
        QREdit2 = QREdit2 & Format(.枚数, "00000")              '箱数（枚数）
        QREdit2 = QREdit2 & Format(.納入数, "0000000000")       '納入数（発注数）
    End With

    QREditBody = True
    Exit Function
QREditBody_Err:
    CisFun.ErrorBox
    End
End Function
'   *------------------*
'   *   『TAIL部』
'   *------------------*
Private Function QREditTail() As Boolean
    QREditTail = False

    On Error GoTo QREditTail_Err

    Mid(QREdit1, 84, 4) = Format(wHakoT, "0000")        '箱総数
    Mid(QREdit1, 88, 2) = Format(wAllPage, "00")        '総頁数
    Mid(QREdit1, 90, 3) = CisFun.Mid2(wNouNo, 2, 3)     'グループNO
    Mid(QREdit1, 93, 2) = Format(wNowPageT, "00")       '物理現頁数
    Mid(QREdit1, 95, 2) = Format(wAllPageT, "00")       '物理総頁数
'   >> END表示区分
    If wNowPage = wAllPage Then Mid(QREdit1, 97, 1) = "1" Else Mid(QREdit1, 97, 1) = "0"
'   >> ALLEND表示区分
    If wNowPageT = wAllPageT Then Mid(QREdit1, 98, 1) = "1" Else Mid(QREdit1, 98, 1) = "0"
'   >> 明細行数
    Mid(QREdit1, 99, 2) = Format(wLineNo, "00")

    QREditTail = True
    Exit Function
QREditTail_Err:
    CisFun.ErrorBox
    End
End Function


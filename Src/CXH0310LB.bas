Attribute VB_Name = "CXH0310LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   自社伝票（納品書・受領書・支給書）発行
'*
'*   モジュールID  : CXH0310LB.BAS
'*
'*                  CXH0310LPrint(Integer)
'*
'*   引数 : Integer  =  0  :新規発行
'*                      1  :再発行
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:
'    If Not CXH0310LPrint(1) Then Exit Sub
'
'*   作成日  :2004/02/09   By CIS
'*   変更日  :2004/04/06   By CIS - 自社が取引先マスタに未登録時、クリアを行う。
'*   変更日  :2004/05/17   By CIS - 受領印の追加。
'*   変更日  :2005/01/27   By CIS - 自社名称獲得が正しくない
'*   変更日  :2005/05/11   By CIS - 行が抜けた場合、行№を印字する
'*   変更日  :2005/12/03   By CIS   CoReport8.5に変更
'*   変更日  :2008/04/11   By CIS - 背番号８桁対応
'*   変更日  :2008/06/06   By CIS - タイトル横にロゴ表示
'*
'*---------------------------------------------------------------------------------
Global CXH0310Sort          As Byte     ' 印刷順
Global CXH0310SortStr       As String   ' 印刷順
Public mQRMeker             As QRmaker  ' QR
Dim CXH0310CoRpt        As New CoReports85
Dim CXH0310BarCode      As String
Global CXH0310PrtOut    As Integer
Global CXH0310QR        As String

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
Dim DP_JYMD         As String
Dim DP_HankoFlg     As Boolean
Dim DP_HankoUpper   As String
Dim DP_HankoLower   As String
Public Function CXH0310LPrint(Optional RePrint As Integer, Optional PrtMsg As Boolean = False)
    CXH0310LPrint = False
    App.Title = "自社伝票"
' 印刷進行画面表示ﾌﾗｸﾞｾｯﾄ
    PrtMsg = True   ' ActiveX使用の為、必ずTrue
    pref_MsgDisp = PrtMsg

    CXH0310PrtOut = gDemoPrt
'帳票区分 （新規発行 OR 再発行 )
    Select Case RePrint
      Case 0:  SyoriKb = 5
      Case 1:  SyoriKb = 105
    End Select
' 出力データ存在確認
    gSL_Select = "SELECT HC.*,ISNULL(SY1.値名称,'') 発注区分名,ISNULL(SY2.値名称,'') 手配区分名,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.略称,'') 仕入先名,ISNULL(TR2.略称,'') 納入先名,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.取引先,'') 仕入先CD,ISNULL(TR1.工場CD,'') 仕入先工場,ISNULL(TR1.工場名,'') 仕入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR2.取引先,'') 納入先CD,ISNULL(TR2.工場CD,'') 納入先工場,ISNULL(TR2.工場名,'') 納入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR3.取引先,'') 契約取引先,ISNULL(TR3.工場CD,'') 契約工場"
    '=================== *** 2008/11/14 Start
    gSL_Select = gSL_Select & " ,ISNULL(SY3.値名称,'') 自社受入名"
    '=================== *** 2008/11/14 End
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 発注帳票発行テーブル HC"
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
    '=================== *** 2008/11/14 Start
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY3"
    gSL_Select = gSL_Select & "   ON  SY3.区分名称 = '自社工区'"
    gSL_Select = gSL_Select & "   AND SY3.英数字1 = HC.受入"
    gSL_Select = gSL_Select & "   AND SY3.区分タイプ = ''"
    '=================== *** 2008/11/14 End
    gSL_Select = gSL_Select & " WHERE 端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND 帳票区分 = " & SyoriKb
'''    gSL_Select = gSL_Select & "   AND HC.納入番号 = 'B8800'"
'''    gSL_Select = gSL_Select & "   AND HC.カード = '7'"
    gSL_Select = gSL_Select & " ORDER BY HC.仕入先,HC.納入番号,HC.発行順,HC.頁NO"
    
    If Not HCHRead(gSL_Select, 0, 1) Then
       Call HCHClose(1)
       CXH0310LPrint = True
       Exit Function
    End If
    If CisDB.RecordCount = 0 Then
       CXH0310LPrint = True
       Exit Function
    End If
    pref_TCnt = CisDB.RecordCount
    
'+++++++++++++++++++++++++++++++'
'+  印刷処理
'+++++++++++++++++++++++++++++++'
    Set pref_Form = CXH0310
    pref_Msg = "【 自 社 伝 票  印 刷 中 】"
    
    If pref_MsgDisp Then
        pref_Form.Show vbModal
        Unload pref_Form
        Set pref_Form = Nothing
    Else
        Call CoPrintH0310
    End If
    
    CXH0310LPrint = True
End Function
'+----------------------------------------------------------+
'+                    印  刷  処  理                        ＋
'+----------------------------------------------------------+
Public Function CoPrintH0310() As Boolean
    CoPrintH0310 = False
'''    Set CorForm = pref_Form
    
    pref_NCnt = 0                ' 現在発行ページ初期化
    On Error GoTo CoPrintH0310_Err
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
' 受領印
    DP_HankoFlg = False
    DP_JYMD = ""
    DP_HankoUpper = ""
    DP_HankoLower = ""
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 20
    Call IRNRead("", 1)
    If Mid(IRN.oyian, 1, 1) = "1" Then
       DP_HankoFlg = True
       If SyoriKb = 5 Then
          If Mid(IRN.oyian, 2, 1) = "0" Then
             DP_JYMD = Format(Date, "yy.mm.dd")
          End If
       End If
       DP_HankoUpper = CisFun.Mid2(IRN.oyian, 3, 8)
       DP_HankoLower = CisFun.Mid2(IRN.oyian2, 1, 8)
    End If
    
'    DP_HankoFlg = True
'    DP_JYMD = Format(Date, "yy.mm.dd")
'    DP_HankoUpper = "物品受領"
'    DP_HankoLower = "ＣＩＳＸ"
                                    
' 取引先マスタより名称を求める
    gSL_Select = "SELECT * FROM 取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"

''''
''''-------------------------------------------------------< 2005.01.27 MNT START >
''''
''''    gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(gCompany) & Left(IRN.oyian, 1) & "'"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(DP_JisyaCd) & RTrim(DP_JisyaKj) & "'"
''''
''''-------------------------------------------------------< 2005.01.27 MNT  END  >
''''

    If TRMRead(gSL_Select, 1) Then
        DP_JisyaNm = TRM.略称
        DP_JisyaKjNm = TRM.工場名
'*---------------------------------------------<< 2004/04/06 Insert Start >>
    Else
        DP_JisyaNm = ""
        DP_JisyaKjNm = ""
'*---------------------------------------------<< 2004/04/06 Insert End >>
    End If
'-----------------------------'
'   CoReport初期設定
'-----------------------------'
    With CXH0310CoRpt
        .SetDraw = pref_Form.CrDraw1
                
        .SetPrtFormCnt = 1
        .SetPrtForm = "LXH0310"
        .SetPrtPath = gIniExe & "Prt\"
        Select Case CXH0310PrtOut
            Case 0: .SetPrtOut = PrinterX
            Case 1: .SetPrtOut = Preview
                    .SetPreviewForm = frmPreview
            Case 2: .SetPrtOut = PDF
        End Select
        .SetPrtName = "自給チケット"
        If Not .CorOutStart Then Exit Function
        
        .SetFormNo = 1
    End With
' ﾊﾞｰｺｰﾄﾞ初期ｾｯﾄ
    Call GetCryOption("G", "バーコード", "自社伝票", gSYSOp_IniAtai)
    If RTrim(gSYSOp_IniAtai) = "" Then
        gSYSOp_IniAtai = "B1"
        Call GetCryOption("P", "バーコード", "自社伝票", gSYSOp_IniAtai)
        Call GetCryOption("P", "バーコード", "BU", "3,5,6,10,3")
        Call GetCryOption("P", "バーコード", "B0", "Form設定")
        Call GetCryOption("P", "バーコード", "B1", "3,5,6,10,3")
        Call GetCryOption("P", "バーコード", "B2", "2,3,4,6,3")
        Call GetCryOption("P", "バーコード", "B3", "4,8,8,16,6")
        Call GetCryOption("P", "バーコード", "B4", "3,4,6,8,3")
    End If
    If gSYSOp_IniAtai = "BU" Then
        Call GetCryOption("G", "バーコード", "BU", CXH0310BarCode)
        If RTrim(CXH0310BarCode) = "" Then
            gSYSOp_IniAtai = "B1"
        End If
    End If
    
' 明細を読み込んでｾｯﾄする(ﾚｺｰﾄﾞ件数分)
    Do Until Not HCH_RDSTS
        ' 項目転送
        If Not CoH0310Move Then GoTo CoPrintH0310_ED
        ' 印刷開始
        CXH0310CoRpt.CoPrintOut
        
        pref_NCnt = pref_NCnt + 1
        If pref_MsgDisp Then
            pref_Cnt = Format(pref_NCnt, "###") & " / " & Format(pref_TCnt, "###")
            DoEvents
        End If
        
        ' 発注帳票発行テーブル　発行日更新(新規発行のみ)
        If SyoriKb = 5 Then
            gSL_Select = "UPDATE 発注帳票発行テーブル SET "
            gSL_Select = gSL_Select & " 発行日 = CONVERT(VARCHAR(8),GETDATE(),112)"
            gSL_Select = gSL_Select & " WHERE 納品書管理NO = " & HCH.納品書管理NO & ""
            gSL_Select = gSL_Select & "   AND 納入番号 = '" & HCH.納入番号 & "'"
            gSL_Select = gSL_Select & "   AND 頁NO = " & HCH.頁NO & ""
            gSL_Select = gSL_Select & "   AND カード = '" & RTrim(HCH.カード) & "'"
            With CisDB
                .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
                On Error GoTo CoPrintH0310_Err
                .SQL = gSL_Select
                .DBExec
                .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
            End With
        
            gSL_Select = "UPDATE 発注テーブル SET "
            Select Case RTrim(HCH.カード)
                Case "6":
                    gSL_Select = gSL_Select & " 受領書発行日 = CONVERT(VARCHAR(8),GETDATE(),112)"
                Case "7":
                    gSL_Select = gSL_Select & " 納品書発行日 = CONVERT(VARCHAR(8),GETDATE(),112)"
                Case "8":
                    gSL_Select = gSL_Select & " 支給書発行日 = CONVERT(VARCHAR(8),GETDATE(),112)"
            End Select
            gSL_Select = gSL_Select & " WHERE 納品書管理NO = " & HCH.納品書管理NO & ""
            gSL_Select = gSL_Select & "   AND 納入番号 = '" & HCH.納入番号 & "'"
            gSL_Select = gSL_Select & "   AND 納品書頁 = " & HCH.頁NO & ""
            With CisDB
                .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
                On Error GoTo CoPrintH0310_Err
                .SQL = gSL_Select
                .DBExec
                .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
            End With
        End If
        
        '次のﾚｺｰﾄﾞを読み込む
        Call HCHReadNext(1)
    Loop

    CoPrintH0310 = True

CoPrintH0310_ED:
    Call HCHClose(1)
' CoReport 終了処理
    If pref_TCnt <> 0 Then CXH0310CoRpt.CoOutEnd
    
'    Call PrintSet2("", True)
    Exit Function
CoPrintH0310_Err:
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
    With HCH
'   *------------------*
'   * 『HEAD』SET  *
'   *------------------*
        Select Case .カード
            Case "6": gStr = "受 領 書"
            Case "7": gStr = "納 品 書"
            Case "8": gStr = "支 給 書"
        End Select
        If Not CXH0310CoRpt.CorData("伝票種類", gStr, "S", 10) Then GoTo CoH0310Move_Err
        If Mid(.納入番号, 5, 1) <> "0" Then
            If Not CXH0310CoRpt.CorData("分納区分", "分割", "", 6) Then GoTo CoH0310Move_Err
        End If
        If Not CXH0310CoRpt.CorData("発注区分", .発注区分名, "", 6) Then GoTo CoH0310Move_Err
        If Not CXH0310CoRpt.CorData("手配区分", .手配区分名, "", 6) Then GoTo CoH0310Move_Err
        If .SK区分 = 0 Then
            If Not CXH0310CoRpt.CorData("SK区分", "かんばん", "", 8) Then GoTo CoH0310Move_Err
        Else
            If Not CXH0310CoRpt.CorData("SK区分", "指示部品", "", 8) Then GoTo CoH0310Move_Err
        End If
        If RTrim(.発行日) <> "" Then
            Call CXH0310CoRpt.CorEditItemSet("再発行", True)
        Else
            Call CXH0310CoRpt.CorEditItemSet("再発行", False)
        End If
    ' 発行日 編集
        With CisFun
            .DateE_BefVal = Format(Now(), "yyyymmdd")
            HCH.発行日 = RTrim(.DateE_BefVal)
            .DateEdit J_YMD
            gStr = .DateE_AfVal
        End With
        If Not CXH0310CoRpt.CorData("発行日", gStr, "", 14) Then GoTo CoH0310Move_Err
    '納入先
        If RTrim(.手配区分) <> "1" Then
            If Not CXH0310CoRpt.CorData("納入先", Left(DP_JisyaCd, 5), "", 5) Then GoTo CoH0310Move_Err
            If Not CXH0310CoRpt.CorData("納入先名称1", CisFun.Left2(DP_JisyaNm, 20), "", 20) Then GoTo CoH0310Move_Err
            
            '==================== *** 2008/11/14 Start
            'If Not CXH0310CoRpt.CorData("納入先名称2", CisFun.Left2(DP_JisyaKjNm, 16), "", 16) Then GoTo CoH0310Move_Err
            If RTrim(.自社受入名) <> "" Then
                If Not CXH0310CoRpt.CorData("納入先名称2", CisFun.Left2(RTrim(.自社受入名), 16), "", 16) Then GoTo CoH0310Move_Err
            Else
                If Not CXH0310CoRpt.CorData("納入先名称2", CisFun.Left2(DP_JisyaKjNm, 16), "", 16) Then GoTo CoH0310Move_Err
            End If
            '==================== *** 2008/11/14 End
            
            If Not CXH0310CoRpt.CorData("納入先事業所1", Left(DP_JisyaKj, 2), "", 2) Then GoTo CoH0310Move_Err
            If RTrim(.受入) <> "" Then
                If Not CXH0310CoRpt.CorData("納入先事業所2", "-", "", 1) Then GoTo CoH0310Move_Err
                If Not CXH0310CoRpt.CorData("納入先事業所3", .受入, "", 2) Then GoTo CoH0310Move_Err
            End If
            If Not CXH0310CoRpt.CorData("支給元", .契約先, "", 5) Then GoTo CoH0310Move_Err
            If Not CXH0310CoRpt.CorData("支給元1", .契約先, "", 5) Then GoTo CoH0310Move_Err
            If Not CXH0310CoRpt.CorData("支給元2", "-", "", 1) Then GoTo CoH0310Move_Err
        Else
            If Not CXH0310CoRpt.CorData("納入先", Left(.納入先CD, 5), "", 5) Then GoTo CoH0310Move_Err
            If Not CXH0310CoRpt.CorData("納入先名称1", CisFun.Left2(.納入先名, 20), "", 20) Then GoTo CoH0310Move_Err
            If Not CXH0310CoRpt.CorData("納入先名称2", CisFun.Left2(.納入先工場名, 16), "", 16) Then GoTo CoH0310Move_Err
            If Not CXH0310CoRpt.CorData("納入先事業所1", Left(.納入先工場, 2), "", 2) Then GoTo CoH0310Move_Err
            If RTrim(.納入先受入) <> "" Then
                If Not CXH0310CoRpt.CorData("納入先事業所2", "-", "", 1) Then GoTo CoH0310Move_Err
                If Not CXH0310CoRpt.CorData("納入先事業所3", .納入先受入, "", 2) Then GoTo CoH0310Move_Err
            End If
            If Not CXH0310CoRpt.CorData("支給元", Left(gCompany, 5), "", 5) Then GoTo CoH0310Move_Err
            If Not CXH0310CoRpt.CorData("支給元1", Left(gCompany, 5), "", 5) Then GoTo CoH0310Move_Err
            If Not CXH0310CoRpt.CorData("支給元2", "-", "", 1) Then GoTo CoH0310Move_Err
        End If
        
        If Not CXH0310CoRpt.CorData("納入場所", "", "", 4) Then GoTo CoH0310Move_Err
        With CisFun
            .DateE_BefVal = HCH.納入日
            .DateEdit J_YMD
            gStr = .DateE_AfVal
        End With
        If Not CXH0310CoRpt.CorData("納入指示日", gStr, "", 14) Then GoTo CoH0310Move_Err
        If RTrim(.納入便) <> "" Then
            If Not CXH0310CoRpt.CorData("納入指示便", RTrim(.納入便) & "便", "", 4) Then GoTo CoH0310Move_Err
        End If
        If Not CXH0310CoRpt.CorData("機番", DP_Kiban, "", 3) Then GoTo CoH0310Move_Err
        If Not CXH0310CoRpt.CorData("カード", .カード, "", 1) Then GoTo CoH0310Move_Err
'        If Not CXH0310CoRpt.cordata("担当者", "", "", 10) Then GoTo CoH0310Move_Err
        
        If Not CXH0310CoRpt.CorData("納品書NO", .納入番号, "", 5) Then GoTo CoH0310Move_Err
        
    '受注者
        If Not CXH0310CoRpt.CorData("受注者", Left(.仕入先CD, 5), "", 5) Then GoTo CoH0310Move_Err
        If Not CXH0310CoRpt.CorData("受注者名称1", CisFun.Left2(.仕入先名, 20), "", 20) Then GoTo CoH0310Move_Err
        If Not CXH0310CoRpt.CorData("受注者名称2", CisFun.Left2(.仕入先工場名, 16), "", 16) Then GoTo CoH0310Move_Err
        If Not CXH0310CoRpt.CorData("受注者事業所", Left(.仕入先工場, 2), "", 2) Then GoTo CoH0310Move_Err
    'ボックス設定
    If RTrim(.手配区分) = "1" And RTrim(.カード) = "7" Then
        Call CXH0310CoRpt.CorEditItemSet("直送Box1", True)
        Call CXH0310CoRpt.CorEditItemSet("直送Box2", True)
        Call CXH0310CoRpt.CorEditItemSet("直送Box3", True)
        Call CXH0310CoRpt.CorEditItemSet("直送Box4", True)
        Call CXH0310CoRpt.CorEditItemSet("直送Box5", True)
        Call CXH0310CoRpt.CorEditItemSet("直送Box6", True)
        Call CXH0310CoRpt.CorEditItemSet("直送Box7", True)
        Call CXH0310CoRpt.CorEditItemSet("直送Box8", True)
    Else
        Call CXH0310CoRpt.CorEditItemSet("直送Box1", False)
        Call CXH0310CoRpt.CorEditItemSet("直送Box2", False)
        Call CXH0310CoRpt.CorEditItemSet("直送Box3", False)
        Call CXH0310CoRpt.CorEditItemSet("直送Box4", False)
        Call CXH0310CoRpt.CorEditItemSet("直送Box5", False)
        Call CXH0310CoRpt.CorEditItemSet("直送Box6", False)
        Call CXH0310CoRpt.CorEditItemSet("直送Box7", False)
        Call CXH0310CoRpt.CorEditItemSet("直送Box8", False)
    End If
    
    Call CXH0310CoRpt.CorImageSet("Logo", gIniExe & "BMP\", "Logo.jpg") '------------ 2008/06/06 ADD
        
'   *------------------*
'   * 『BODY』SET  *
'   *------------------*
        For gInt = 1 To 10
            Call CXH0310CoRpt.CorEditItemSet("品番Box" & gInt, False)
            Call CXH0310CoRpt.CorEditItemSet("数量Box" & gInt, False)
            Call CXH0310CoRpt.CorEditItemSet("GyoLine" & gInt, False)
        Next
        For gInt = 1 To .明細行数
'2005.05.11            gStr = Format(.行(gInt), "##")
            gStr = Format(gInt, "##")
            If Not CXH0310CoRpt.CorListSet("NO", gStr, gInt - 1, 2) Then GoTo CoH0310Move_Err
            If Not CXH0310CoRpt.CorListSet("品番", .表示品番(gInt), gInt - 1, 30) Then GoTo CoH0310Move_Err
            gStr = CisFun.Left2(.部品名称(gInt), 32)
            If Not CXH0310CoRpt.CorListSet("品名", gStr, gInt - 1, 32) Then GoTo CoH0310Move_Err
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE STTART
            'If Not CXH0310CoRpt.CorListSet("背番号", .背番号(gInt), gInt - 1, 5) Then GoTo CoH0310Move_Err
            If Not CXH0310CoRpt.CorListSet("背番号", .背番号(gInt), gInt - 1, 8) Then GoTo CoH0310Move_Err
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
            gStr = Format(.収容数(gInt), "##")
            If Not CXH0310CoRpt.CorListSet("収容数", gStr, gInt - 1, 7) Then GoTo CoH0310Move_Err
            gStr = Format(.箱数(gInt), "##")
            If Not CXH0310CoRpt.CorListSet("箱数", gStr, gInt - 1, 6) Then GoTo CoH0310Move_Err
            gStr = Format(.納入数(gInt), "##")
            If Not CXH0310CoRpt.CorListSet("納入数", gStr, gInt - 1, 7) Then GoTo CoH0310Move_Err
            '品番マスタ参照
            gSL_Select = "SELECT 実数区分 FROM 品番マスタ"
            gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(.部品番号(gInt)) & "'"
            Call HNMRead(gSL_Select, 1)
            If HNM.実数区分 = 0 Then
                gStr = ""
            Else
                gStr = "*"
            End If
            If Not CXH0310CoRpt.CorListSet("分割個数", "個" & gStr, gInt - 1, 3) Then GoTo CoH0310Move_Err
            If Not CXH0310CoRpt.CorListSet("完納予定1", "／", gInt - 1, 2) Then GoTo CoH0310Move_Err
            If Not CXH0310CoRpt.CorListSet("完納予定2", "便", gInt - 1, 2) Then GoTo CoH0310Move_Err
            
            If RTrim(.手配区分) = "1" And RTrim(.カード) = "7" Then
                Call CXH0310CoRpt.CorEditItemSet("品番Box" & gInt, True)
                Call CXH0310CoRpt.CorEditItemSet("数量Box" & gInt, True)
            End If
            Call CXH0310CoRpt.CorEditItemSet("GyoLine" & gInt, True)
        Next gInt
'   *------------------*
'   * 『TAIL』SET  *
'   *------------------*
        gStr = Format(.箱総数, "####0")
        If Not CXH0310CoRpt.CorData("総箱数", gStr, "", 7) Then GoTo CoH0310Move_Err
        If Not CXH0310CoRpt.CorData("GROUPNO", .グループNO, "", 3) Then GoTo CoH0310Move_Err
        gStr = Format(.頁NO, "#0")
        If Not CXH0310CoRpt.CorData("帳票現頁", gStr, "", 2) Then GoTo CoH0310Move_Err
        gStr = Format(.総頁数, "#0")
        If Not CXH0310CoRpt.CorData("帳票総頁", gStr, "", 2) Then GoTo CoH0310Move_Err
        gStr = Format(.物理現頁数, "#0")
        If Not CXH0310CoRpt.CorData("納番現頁", gStr, "", 2) Then GoTo CoH0310Move_Err
        gStr = Format(.物理総頁数, "#0")
        If Not CXH0310CoRpt.CorData("納番総頁", gStr, "", 2) Then GoTo CoH0310Move_Err
        Select Case .カード
            Case "6": gStr = "●"
            Case "7": gStr = "■"
            Case "8": gStr = "▲"
            Case Else: gStr = "☆"
            
        End Select
        If Not CXH0310CoRpt.CorData("帳票マーク", gStr, "", 2) Then GoTo CoH0310Move_Err
        If .END表示区分 <> 0 Then
            Call CXH0310CoRpt.CorEditItemSet("ALLENDBox", True)
            If Not CXH0310CoRpt.CorData("END区分", "END", "", 6) Then GoTo CoH0310Move_Err
        Else
            Call CXH0310CoRpt.CorEditItemSet("ALLENDBox", False)
        End If
        If .ALLEND表示区分 <> 0 Then
            Call CXH0310CoRpt.CorEditItemSet("ALLENDBox", True)
            If Not CXH0310CoRpt.CorData("ALLEND区分", "ALLEND", "", 6) Then GoTo CoH0310Move_Err
        Else
            Call CXH0310CoRpt.CorEditItemSet("ALLENDBox", False)
        End If


'   *---------------------*
'   * 『受領印』SET    *
'   *---------------------*
        If RTrim(.カード) = "6" Then
            If DP_HankoFlg Then
                gStr = Dir(gIniExe & "BMP\受領印.BMP")
                If RTrim(gStr) <> "" Then
                    If Not CXH0310CoRpt.CorImage("受領印", gIniExe & "BMP\受領印.BMP") Then Exit Function
                    
                    If Not CXH0310CoRpt.CorData("受領日付", DP_JYMD, "", 8) Then GoTo CoH0310Move_Err
                    If Not CXH0310CoRpt.CorData("受領印上", DP_HankoUpper, "", 8) Then GoTo CoH0310Move_Err
                    If Not CXH0310CoRpt.CorData("受領印下", DP_HankoLower, "", 8) Then GoTo CoH0310Move_Err
                End If
            End If
        End If
'   *---------------------*
'   * 『バーコード』SET    *
'   *---------------------*
        If .納品書管理NO <> 0 Then
            gStr = Format(.納品書管理NO, "0000000000") & Format(.頁NO, "00")
            Select Case gSYSOp_IniAtai
                Case "BU": CXH0310CoRpt.SetBar39 = User
                            CXH0310CoRpt.SetBar39User = CXH0310BarCode
                Case "B0": CXH0310CoRpt.SetBar39 = form設定
                Case "B2": CXH0310CoRpt.SetBar39 = パターン2_5_8_13_20_7
                Case "B3": CXH0310CoRpt.SetBar39 = パターン3_4_8_8_16_6
                Case "B4": CXH0310CoRpt.SetBar39 = パターン4_3_4_6_8_3
                Case Else: CXH0310CoRpt.SetBar39 = パターン1_3_5_6_10_3
            End Select
            If Not CXH0310CoRpt.CorData("BAR", gStr, "", 12) Then GoTo CoH0310Move_Err
            
            gStr = "*" & gStr & "*"
            If Not CXH0310CoRpt.CorData("BAR_Naiyo", gStr, "", 14) Then GoTo CoH0310Move_Err
        End If
'   *---------------------*
'   * 『QR』    SET       *
'   *---------------------*
        Call Denpyo_QREdit
        If Not CXH0310CoRpt.CorQRData("QR", CXH0310QR, "", 610) Then Exit Function
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
'''''Public Function CoReportProcH0310(Repart As String)
'''''    CoReportProcH0310 = False
''''''   *******************************************************************
''''''   * フォーム変更処理                              ( Repart = "P" )  *
''''''   *******************************************************************
'''''    If Repart = "C" Then
''''''       CorPrtForm = "検査指示2.FED"
''''''       If CorChange(G_Crprint_Path & CorPrtForm) Then CoReportProcH0310 = True
'''''       Exit Function
'''''    End If
''''''   *******************************************************************
''''''   * 印刷指示処理                                  ( Repart = "P" )  *
''''''   *******************************************************************
'''''    If Repart = "P" Then
'''''        If CorData("", "", "E") Then CoReportProcH0310 = True
'''''    End If
''''''   *******************************************************************
''''''   * 終了処理                                      ( Repart = "E" )  *
''''''   *******************************************************************
'''''    If Repart = "E" Then
'''''        If CorOpCl(crPara, "E") Then CoReportProcH0310 = True
'''''        Exit Function
'''''    End If
'''''
''''''   *****************************************************************
''''''   * オープン処理    <<ﾌﾟﾘﾝﾀﾄﾞﾗｲﾊﾞ設定・初期処理>> ( Repart = "O" )  *
''''''   *****************************************************************
'''''    If Repart = "O" Then
'''''' CoReport 初期処理
'''''
'''''        With crPara
''''''           * CoStruct Set *
'''''            .cshwnd = CorForm.hWnd
'''''            .cshdc = Printer.hDC
'''''            .csout = OUTPUT_PRNDRIVER
'''''            .csScale = 100&
'''''            .csmx = 0: .csmy = 0
'''''            .cssx = 0: .cssy = 0
'''''            .cswidth = 0: .csheight = 0
''''''            .csfname = gIniExe & CorPrtForm & ".FED"
'''''            .csoutfname = ""
''''''           * DevModeStruct Set *　プリンタの設定
'''''            CorPrtForm = "LXH0310"
'''''            .csfname = gIniExe & "Prt\" & CorPrtForm & ".FED"
'''''            .csdevm.dmOrientation = DMORIENT_PORTRAIT   'DMORIENT_LANDSCAPE
''''''                .csdevm.dmOrientation = DMORIENT_PORTRAIT
'''''            .csdevm.dmPaperSize = DMPAPER_A4
'''''            If pref_PrintNo = 0 Then
'''''               .csdevm.dmFields = DM_TEMPORARY Or DM_PAPERSIZE Or DM_ORIENTATION
'''''            Else
'''''               .csdevm.dmDefaultSource = pref_PrintNo
'''''               .csdevm.dmFields = DM_TEMPORARY Or DM_DEFAULTSOURCE Or DM_PAPERSIZE Or DM_ORIENTATION
'''''            End If
'''''
'''''        End With
'''''
'''''        If CorOpCl(crPara, "O") Then
'''''            CoReportProcH0310 = True
'''''        End If
'''''
'''''    End If
'''''End Function

'*------------------------------------------------------------------*
'*                                                                  *
'*      自社伝票　QR　作成                                           *
'*                                                                  *
'*------------------------------------------------------------------*
Private Function Denpyo_QREdit() As Boolean
    Dim QREdit1      As String * 100
    Dim QREdit2      As String
    
    QREdit1 = Space(100)
    
    With HCH
        Mid(QREdit1, 1, 7) = Left(DP_HMoto, 7)
        Mid(QREdit1, 8, 5) = Left(gCompany, 5)
        Mid(QREdit1, 13, 1) = "D"
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
        Mid(QREdit1, 63, 8) = RTrim(.納入日)               '納入日
        Mid(QREdit1, 71, 2) = RTrim(.納入便)               '納入便
        Mid(QREdit1, 73, 8) = RTrim(.作成日)               '作成日
        Mid(QREdit1, 81, 3) = RTrim(DP_Kiban)             '機番
        Mid(QREdit1, 84, 4) = Format(.箱総数, "0000")      '箱総数
        Mid(QREdit1, 88, 2) = Format(.総頁数, "00")        '総頁数
        Mid(QREdit1, 90, 3) = RTrim(.グループNO)           'グループNO
        Mid(QREdit1, 93, 2) = Format(.物理現頁数, "00")      '物理現頁数
        Mid(QREdit1, 95, 2) = Format(.物理総頁数, "00")      '物理総頁数
        Mid(QREdit1, 97, 1) = Format(.END表示区分, "0")      'END表示区分
        Mid(QREdit1, 98, 1) = Format(.ALLEND表示区分, "0")   'ALLEND表示区分
        Mid(QREdit1, 99, 2) = Format(.明細行数, "00")        '明細行数
        
'''        QREdit2 = ""
'''        QREdit2 = QREdit2 & CisFun.Left2(.発注区分名, 6)   '発注区分名
'''        QREdit2 = QREdit2 & Space(6 - CisFun.Len2(CisFun.Left2(.発注区分名, 6)))
'''        QREdit2 = QREdit2 & CisFun.Left2(.手配区分名, 6)   '手配区分名
'''        QREdit2 = QREdit2 & Space(6 - CisFun.Len2(CisFun.Left2(.手配区分名, 6)))
'''        If .SK区分 = 0 Then
'''            QREdit2 = QREdit2 & "かんばん"   'SK区分名(かんばん）
'''        Else
'''            QREdit2 = QREdit2 & "指示部品"   'SK区分名(指示）
'''        End If
'''        QREdit2 = QREdit2 & CisFun.Left2(.仕入先名, 24)   '仕入先名
'''        QREdit2 = QREdit2 & Space(24 - CisFun.Len2(CisFun.Left2(.仕入先名, 24)))
'''        If RTrim(.手配区分) <> "1" Then
'''            QREdit2 = QREdit2 & CisFun.Left2(DP_JisyaNm, 24)  '納入先名
'''            QREdit2 = QREdit2 & Space(24 - CisFun.Len2(CisFun.Left2(DP_JisyaNm, 24)))
'''        Else
'''            QREdit2 = QREdit2 & CisFun.Left2(.納入先名, 24)   '納入先名
'''            QREdit2 = QREdit2 & Space(24 - CisFun.Len2(CisFun.Left2(.納入先名, 24)))
'''        End If
'''        QREdit2 = QREdit2 & Format(.明細行数, "00")      '明細行数
'''
        For gInt = 1 To .明細行数
            QREdit2 = QREdit2 & RTrim(.背番号(gInt))                      '背番号
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
            'QREdit2 = QREdit2 & Space(5 - CisFun.Len2(.背番号(gInt)))
            QREdit2 = QREdit2 & Space(8 - CisFun.Len2(.背番号(gInt)))
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
            QREdit2 = QREdit2 & RTrim(.部品番号(gInt))                    '部品番号
            QREdit2 = QREdit2 & Space(25 - CisFun.Len2(.部品番号(gInt)))
'''            HNM.品番 = RTrim(.部品番号(gInt))
'''            If HNMRead("", 1) Then
'''                QREdit2 = QREdit2 & RTrim(HNM.品番編集)                      '品番編集
'''            End If
'''            QREdit2 = QREdit2 & Space(5 - CisFun.Len2(HNM.品番編集))
'''            QREdit2 = QREdit2 & RTrim(.部品名称(gInt))                    '部品名称
'''            QREdit2 = QREdit2 & Space(40 - CisFun.Len2(.部品名称(gInt)))
            QREdit2 = QREdit2 & Format(.収容数(gInt), "000000")           '収容数
            QREdit2 = QREdit2 & Format(.箱数(gInt), "00000")              '箱数（枚数）
'''            QREdit2 = QREdit2 & Format(.端数(gInt), "000000")             '端数
            QREdit2 = QREdit2 & Format(.納入数(gInt), "0000000000")       '納入数（発注数）
'''
        
        Next gInt
    End With
    CXH0310QR = RTrim(QREdit1) & RTrim(QREdit2)
    
End Function


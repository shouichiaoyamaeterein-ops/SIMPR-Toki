Attribute VB_Name = "CXH0390LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   自社伝票（チケット）発行
'*
'*   モジュールID  : CXH0390LB.BAS
'*
'*                  CXH0390LPrint(Integer)
'*
'*   引数 : Integer  =  0  :新規発行
'*                      1  :再発行
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:
'    If Not CXH0390LPrint(1) Then Exit Sub
'
'*   作成日  :  2005/02/22   By CIS
'*   変更日  :  2005/12/03  CoReport8.5に変更
'*   変更日  :  2006/02/09  実機にて(TOPIX TX-435A)テスト及び、修正
'*   変更日  :  2008/04/11  By CIS - 背番号８桁対応
'*
'*---------------------------------------------------------------------------------
Public mQRMeker390      As QRmaker  ' QR
Dim CXH0390CoRpt        As New CoReports85
Dim CXH0390BarCode      As String
Global CXH0390PrtOut    As Integer
Global CXH0390QR    As String
Global OUT_Printer      As String               '2006/02/14

Dim SaveBun         As String
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
Dim DP_HMaisu       As Long

Public Function CXH0390LPrint(Optional RePrint As Integer, Optional PrtMsg As Boolean = False)
    CXH0390LPrint = False
    App.Title = "自社伝票"
' 印刷進行画面表示ﾌﾗｸﾞｾｯﾄ
    PrtMsg = True   ' ActiveX使用の為、必ずTrue
    pref_MsgDisp = PrtMsg

    CXH0390PrtOut = gDemoPrt
'帳票区分 （新規発行 OR 再発行 )
    Select Case RePrint
      Case 0:  SyoriKb = 17
      Case 1:  SyoriKb = 117
    End Select
' 出力データ存在確認
    gSL_Select = "SELECT Case When Substring(HC.納入番号,5,1) = '0' then '0' else '1' end 分納区分,"
    gSL_Select = gSL_Select & " HC.*,ISNULL(SY1.値名称,'') 発注区分名,ISNULL(SY2.値名称,'') 手配区分名,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.略称,'') 仕入先名,ISNULL(TR2.略称,'') 納入先名,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.取引先,'') 仕入先CD,ISNULL(TR1.工場CD,'') 仕入先工場,ISNULL(TR1.工場名,'') 仕入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR2.取引先,'') 納入先CD,ISNULL(TR2.工場CD,'') 納入先工場,ISNULL(TR2.工場名,'') 納入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR3.取引先,'') 契約取引先,ISNULL(TR3.工場CD,'') 契約工場"
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 発注帳票発行テーブル HC"
    gSL_Select = gSL_Select & "   ON  HC.納品書管理NO =CH.管理NO"
    gSL_Select = gSL_Select & "   AND HC.納入番号 = CH.管理文字"
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
    gSL_Select = gSL_Select & " ORDER BY HC.仕入先,HC.納入番号,HC.発行順,HC.頁NO"
    
    If Not HCHRead(gSL_Select, 0, 1) Then
       Call HCHClose(1)
       CXH0390LPrint = True
       Exit Function
    End If
    If CisDB.RecordCount = 0 Then
       CXH0390LPrint = True
       Exit Function
    End If
    pref_TCnt = CisDB.RecordCount
    
'+++++++++++++++++++++++++++++++'
'+  印刷処理
'+++++++++++++++++++++++++++++++'
    Set pref_Form = CXH0390
    pref_Msg = "【 自社伝票(ﾁｹｯﾄ) 印 刷 中 】"
    
    If pref_MsgDisp Then
        pref_Form.Show vbModal
        Unload pref_Form
        Set pref_Form = Nothing
    Else
        If pref_Form Is Nothing Then
           Set pref_Form = CXH0390
        End If
        Call CoPrintH0390
    End If
    
    CXH0390LPrint = True
End Function
'+----------------------------------------------------------+
'+                    印  刷  処  理                        ＋
'+----------------------------------------------------------+
Public Function CoPrintH0390() As Boolean
    CoPrintH0390 = False
'''''    Set CorForm = pref_Form
    
    pref_NCnt = 0                ' 現在発行ページ初期化
    On Error GoTo CoPrintH0390_Err

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
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(DP_JisyaCd) & RTrim(DP_JisyaKj) & "'"
    If TRMRead(gSL_Select, 1) Then
        DP_JisyaNm = TRM.略称
        DP_JisyaKjNm = TRM.工場名
    Else
        DP_JisyaNm = ""
        DP_JisyaKjNm = ""
    End If
    
'-----------------------------'
'   CoReport初期設定
'-----------------------------'
    With CXH0390CoRpt
        .SetDraw = pref_Form.CrDraw1
                
        .SetPrtFormCnt = 2
        .SetPrtForm = "LXH0390"
        .SetPrtForms("LXH0391") = 2
        .SetPrtPath = gIniExe & "Prt\"
        
'---------------------------------------------------------------------< 2006.02.14 ADD START >
'
        .SetPaperSize = UserSize
        .SetPaperWidth = 20000
        .SetPaperLength = 10000
        .SetPrinter = OUT_Printer
'
'---------------------------------------------------------------------< 2006.02.14 ADD  END  >
        
        Select Case CXH0390PrtOut
            Case 0: .SetPrtOut = PrinterX
            Case 1: .SetPrtOut = Preview
                    .SetPreviewForm = frmPreview
            Case 2: .SetPrtOut = PDF
        End Select
        .SetPrtName = "自給チケット"
        If Not .CorOutStart Then Exit Function
        
    End With
' ﾊﾞｰｺｰﾄﾞ初期ｾｯﾄ
    Call GetCryOption("G", "バーコード", "自給チケット", gSYSOp_IniAtai)
    If RTrim(gSYSOp_IniAtai) = "" Then
        gSYSOp_IniAtai = "B1"
        Call GetCryOption("P", "バーコード", "自給チケット", gSYSOp_IniAtai)
        Call GetCryOption("P", "バーコード", "BU", "3,5,6,10,3")
        Call GetCryOption("P", "バーコード", "B0", "Form設定")
        Call GetCryOption("P", "バーコード", "B1", "3,5,6,10,3")
        Call GetCryOption("P", "バーコード", "B2", "2,3,4,6,3")
        Call GetCryOption("P", "バーコード", "B3", "4,8,8,16,6")
        Call GetCryOption("P", "バーコード", "B4", "3,4,6,8,3")
    End If
    If gSYSOp_IniAtai = "BU" Then
        Call GetCryOption("G", "バーコード", "BU", CXH0390BarCode)
        If RTrim(CXH0390BarCode) = "" Then
            gSYSOp_IniAtai = "B1"
        End If
    End If
    
    SaveBun = HCH.分納区分
    If HCH.分納区分 = "0" Then CXH0390CoRpt.SetFormNo = 1 Else CXH0390CoRpt.SetFormNo = 2
   
' 明細を読み込んでｾｯﾄする(ﾚｺｰﾄﾞ件数分)
    Do Until Not HCH_RDSTS
        If SaveBun <> HCH.分納区分 Then
            If HCH.分納区分 = "0" Then CXH0390CoRpt.SetFormNo = 1 Else CXH0390CoRpt.SetFormNo = 2
           SaveBun = HCH.分納区分
        End If
        
        ' 項目転送
        If Not CoH0390Move Then GoTo CoPrintH0390_ED
        ' 印刷開始
        CXH0390CoRpt.CoPrintOut
        
        pref_NCnt = pref_NCnt + 1
        If pref_MsgDisp Then
            pref_Cnt = Format(pref_NCnt, "###") & " / " & Format(pref_TCnt, "###")
            DoEvents
        End If
        
        ' 発注帳票発行テーブル　発行日更新(新規発行のみ)
        If SyoriKb = 17 Then
            gSL_Select = "UPDATE 発注帳票発行テーブル SET "
            gSL_Select = gSL_Select & " 発行日 = CONVERT(VARCHAR(8),GETDATE(),112)"
            gSL_Select = gSL_Select & " WHERE 納品書管理NO = " & HCH.納品書管理NO & ""
            gSL_Select = gSL_Select & "   AND 納入番号 = '" & HCH.納入番号 & "'"
            gSL_Select = gSL_Select & "   AND 頁NO = " & HCH.頁NO & ""
            With CisDB
                .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
                On Error GoTo CoPrintH0390_Err
                .SQL = gSL_Select
                .DBExec
                .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
            End With
        
            gSL_Select = "UPDATE 発注テーブル SET "
            gSL_Select = gSL_Select & " 納品書発行日 = CONVERT(VARCHAR(8),GETDATE(),112)"
            gSL_Select = gSL_Select & " WHERE 納品書管理NO = " & HCH.納品書管理NO & ""
            gSL_Select = gSL_Select & "   AND 納入番号 = '" & HCH.納入番号 & "'"
            gSL_Select = gSL_Select & "   AND 納品書頁 = " & HCH.頁NO & ""
            With CisDB
                .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
                On Error GoTo CoPrintH0390_Err
                .SQL = gSL_Select
                .DBExec
                .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
            End With
        End If
        
        '次のﾚｺｰﾄﾞを読み込む
        Call HCHReadNext(1)
    Loop

    CoPrintH0390 = True

CoPrintH0390_ED:
    Call HCHClose(1)
' CoReport 終了処理
    If pref_TCnt <> 0 Then CXH0390CoRpt.CoOutEnd
    
'    Call PrintSet2("", True)
    Exit Function
CoPrintH0390_Err:
    CisFun.ErrorBox
'    Call PrintSet2("", True)
    End
End Function
'*------------------------------------------------------------------*
'*                                                                  *
'*      納品書　発行内容　項目転送                                    *
'*                                                                  *
'*------------------------------------------------------------------*
Private Function CoH0390Move() As Boolean
    Dim wBar            As String
    Dim wBarNaiyo       As String
    Dim wHYmd           As String
    Dim wEndGyo         As Integer
    
    CoH0390Move = True
    
    On Error GoTo CoH0390Move_Err
    With HCH
'   *------------------*
'   * 『HEAD』SET  *
'   *------------------*
        gStr = .手配区分名
        If Not CXH0390CoRpt.CorData("Head", gStr & "かんばんチケット", "S", 20) Then GoTo CoH0390Move_Err
        
        If Not CXH0390CoRpt.CorData("仕入先名", .仕入先名, "", 16) Then GoTo CoH0390Move_Err
        gStr = Format(Mid(.納入日, 3, 2), "00")
        If Not CXH0390CoRpt.CorData("納入年", gStr, "", 2) Then GoTo CoH0390Move_Err
        gStr = Format(Mid(.納入日, 5, 2), "#")
        If Not CXH0390CoRpt.CorData("納入月", gStr, "", 2) Then GoTo CoH0390Move_Err
        gStr = Format(Mid(.納入日, 7, 2), "#")
        If Not CXH0390CoRpt.CorData("納入日", gStr, "", 2) Then GoTo CoH0390Move_Err
        If Trim(.納入便) <> "" Then
           If Not CXH0390CoRpt.CorData("納入便", .納入便, "", 2) Then GoTo CoH0390Move_Err
           If Not CXH0390CoRpt.CorData("便", "便", "", 2) Then GoTo CoH0390Move_Err
        End If
        gStr = .納入番号 & "-" & Format(.頁NO, "00") & "/" & Format(.総頁数, "00") & "-" & Format(.分納時初回行, "0")
        If Not CXH0390CoRpt.CorData("納入番号", gStr, "", 13) Then GoTo CoH0390Move_Err
        If Not CXH0390CoRpt.CorData("伝票", .カード, "", 1) Then GoTo CoH0390Move_Err
        gStr = .仕入先
        If .仕入先工場 <> "" Then
            gStr = gStr & "-" & .仕入先工場
        End If
        If Not CXH0390CoRpt.CorData("仕入先", gStr, "", 10) Then GoTo CoH0390Move_Err
        If .手配区分 = "1" Then
           If Not CXH0390CoRpt.CorData("支給先", .納入先, "", 4) Then GoTo CoH0390Move_Err
           If Not CXH0390CoRpt.CorData("仕入先名2", .納入先名, "", 16) Then GoTo CoH0390Move_Err
           If Not CXH0390CoRpt.CorData("支給元", DP_JisyaCd, "", 4) Then GoTo CoH0390Move_Err
        Else
           If Not CXH0390CoRpt.CorData("支給先", DP_JisyaCd, "", 4) Then GoTo CoH0390Move_Err
           If Not CXH0390CoRpt.CorData("仕入先名2", DP_JisyaNm, "", 16) Then GoTo CoH0390Move_Err
        End If
        If Not CXH0390CoRpt.CorData("受入", .受入, "", 2) Then GoTo CoH0390Move_Err
        If Not CXH0390CoRpt.CorData("機番", DP_Kiban, "", 2) Then GoTo CoH0390Move_Err
        If CisFun.Val2(.発注日) <> 0 Then
           gStr = Format(Mid(.発注日, 5, 2), "#")
           If Not CXH0390CoRpt.CorData("発注月", gStr, "", 2) Then GoTo CoH0390Move_Err
           gStr = Format(Mid(.発注日, 7, 2), "#")
           If Not CXH0390CoRpt.CorData("発注日", gStr, "", 2) Then GoTo CoH0390Move_Err
           If Trim(.発注便) <> "" Then
              If Not CXH0390CoRpt.CorData("発注便", .発注便, "", 2) Then GoTo CoH0390Move_Err
              If Not CXH0390CoRpt.CorData("便発注", "便", "", 2) Then GoTo CoH0390Move_Err
           End If
        End If
'   *------------------*
'   * 『BODY』SET  *
'   *------------------*
        DP_HMaisu = 0
        wEndGyo = .明細行数
        If .分納区分 <> "0" Then
           wEndGyo = 1
        End If
        For gInt = 1 To wEndGyo
            If .分納区分 <> "0" Then
               gStr = "1"
            Else
               gStr = Format(.行(gInt), "##")
            End If
            If Not CXH0390CoRpt.CorListSet("No", gStr, gInt - 1, 1) Then GoTo CoH0390Move_Err
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
            'If Not CXH0390CoRpt.CorListSet("背番号", .背番号(gInt), gInt - 1, 5) Then GoTo CoH0390Move_Err
            If Not CXH0390CoRpt.CorListSet("背番号", .背番号(gInt), gInt - 1, 8) Then GoTo CoH0390Move_Err
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
            If Not CXH0390CoRpt.CorListSet("品番", .表示品番(gInt), gInt - 1, 14) Then GoTo CoH0390Move_Err
            gStr = Format(.収容数(gInt), "##")
            If Not CXH0390CoRpt.CorListSet("収容数", gStr, gInt - 1, 6) Then GoTo CoH0390Move_Err
            gStr = Format(.箱数(gInt), "##")
            DP_HMaisu = DP_HMaisu + .箱数(gInt)
            
            If Not CXH0390CoRpt.CorListSet("箱数", gStr, gInt - 1, 3) Then GoTo CoH0390Move_Err
            gStr = Format(.納入数(gInt), "##")
            If Not CXH0390CoRpt.CorListSet("納入数", gStr, gInt - 1, 7) Then GoTo CoH0390Move_Err
            '備考
            gStr = ""
            Select Case .打切区分(gInt)
                Case 1: gStr = "打切り中"
                Case 9: gStr = "打切り完了"
            End Select
            If gStr <> "" Then
               If Not CXH0390CoRpt.CorListSet("備考", gStr, gInt - 1, 14) Then GoTo CoH0390Move_Err
            End If

        Next gInt
'   *------------------*
'   * 『TAIL』SET  *
'   *------------------*
        If Not CXH0390CoRpt.CorData("担当者", .担当者, "", 10) Then GoTo CoH0390Move_Err
        gStr = DP_HMaisu
        If Not CXH0390CoRpt.CorData("総箱数", gStr, "", 4) Then GoTo CoH0390Move_Err
        
        ' 作成日
        wHYmd = Format(Now(), "yyyymmdd")
        gStr = Format(Mid(wHYmd, 3, 2), "00")
        If Not CXH0390CoRpt.CorData("作成年", gStr, "", 2) Then GoTo CoH0390Move_Err
        gStr = Format(Mid(wHYmd, 5, 2), "#")
        If Not CXH0390CoRpt.CorData("作成月", gStr, "", 2) Then GoTo CoH0390Move_Err
        gStr = Format(Mid(wHYmd, 7, 2), "#")
        If Not CXH0390CoRpt.CorData("作成日", gStr, "", 2) Then GoTo CoH0390Move_Err
        gStr = Format(Time, "h")
        If Not CXH0390CoRpt.CorData("作成時", gStr, "", 2) Then GoTo CoH0390Move_Err
        
        gStr = Format(.頁NO, "0") & "/" & Format(.総頁数, "0")
        If Not CXH0390CoRpt.CorData("ページ", gStr, "", 5) Then GoTo CoH0390Move_Err

'   *---------------------*
'   * 『再発行』SET    *
'   *---------------------*
        If SyoriKb = 117 Then
            Call CXH0390CoRpt.CorEditItemSet("再発行", True)
        Else
            Call CXH0390CoRpt.CorEditItemSet("再発行", False)
        End If
'   *---------------------*
'   * 『バーコード』SET    *
'   *---------------------*
        If .納品書管理NO <> 0 Then
            gStr = Format(.納品書管理NO, "0000000000") & Format(.頁NO, "00")
            Select Case gSYSOp_IniAtai
                Case "BU": CXH0390CoRpt.SetBar39 = User
                            CXH0390CoRpt.SetBar39User = CXH0390BarCode
                Case "B0": CXH0390CoRpt.SetBar39 = form設定
                Case "B2": CXH0390CoRpt.SetBar39 = パターン2_5_8_13_20_7
                Case "B3": CXH0390CoRpt.SetBar39 = パターン3_4_8_8_16_6
                Case "B4": CXH0390CoRpt.SetBar39 = パターン4_3_4_6_8_3
                Case Else: CXH0390CoRpt.SetBar39 = パターン1_3_5_6_10_3
            End Select
            If Not CXH0390CoRpt.CorData("BAR", gStr, "", 12) Then GoTo CoH0390Move_Err
            gStr = "*" & gStr & "*"
            If Not CXH0390CoRpt.CorData("BAR_Naiyo", gStr, "", 14) Then GoTo CoH0390Move_Err
        End If
'   *---------------------*
'   * 『QR』    SET       *
'   *---------------------*
        Call Denpyo_QREdit
        If Not CXH0390CoRpt.CorQRData("QR", CXH0390QR, "", 400) Then Exit Function
    End With

    Exit Function
CoH0390Move_Err:
    CoH0390Move = False
End Function
'*==========================================================================*
'*==========================================================================*
'*------------------------------------------------------------------*
'*                                                                  *
'*      CoReport Proccess           (O)Open   (E)Close              *
'*                                                                  *
'*------------------------------------------------------------------*
'''''Public Function CoReportProcH0390(Repart As String)
'''''    CoReportProcH0390 = False
''''''   *******************************************************************
''''''   * フォーム変更処理                              ( Repart = "P" )  *
''''''   *******************************************************************
'''''    If Repart = "C" Then
''''''       CorPrtForm = "検査指示2.FED"
''''''       If CorChange(G_Crprint_Path & CorPrtForm) Then CoReportProcH0390 = True
'''''       Exit Function
'''''    End If
''''''   *******************************************************************
''''''   * 印刷指示処理                                  ( Repart = "P" )  *
''''''   *******************************************************************
'''''    If Repart = "P" Then
'''''        If CorData("", "", "E") Then CoReportProcH0390 = True
'''''    End If
''''''   *******************************************************************
''''''   * 終了処理                                      ( Repart = "E" )  *
''''''   *******************************************************************
'''''    If Repart = "E" Then
'''''        If CorOpCl(crPara, "E") Then CoReportProcH0390 = True
'''''        Exit Function
'''''    End If
'''''
''''''   *****************************************************************
''''''   * オープン処理    <<ﾌﾟﾘﾝﾀﾄﾞﾗｲﾊﾞ設定・初期処理>> ( Repart = "O" )  *
''''''   *****************************************************************
'''''    If Repart = "O" Then
'''''        If SaveBun = "0" Then
'''''           CorPrtForm = "LXH0390"
'''''        Else
'''''           CorPrtForm = "LXH0391"
'''''        End If
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
'''''            .csfname = gIniExe & "Prt\" & CorPrtForm & ".FED"
'''''            .csdevm.dmPaperLength = 1000
'''''            .csdevm.dmPaperWidth = 2000
'''''            .csdevm.dmOrientation = DMORIENT_PORTRAIT   'DMORIENT_LANDSCAPE
'''''            .csdevm.dmPaperSize = DMPAPER_USER
'''''            .csdevm.dmFields = DM_TEMPORARY Or DM_PAPERSIZE Or _
'''''                               DM_PAPERLENGTH Or DM_PAPERWIDTH Or _
'''''                               DM_ORIENTATION
'''''
'''''        End With
'''''
'''''        If CorOpCl(crPara, "O") Then
'''''            CoReportProcH0390 = True
'''''        End If
'''''        'ﾊﾞｰｺｰﾄﾞ初期ｾｯﾄ
'''''        Call GetBarInfo("JISYACHIKET")
'''''        Erase CoBarInf
'''''        CoBarInf(0) = -1
'''''        CoBarInf(1) = CowBar_Info(0)
'''''        CoBarInf(2) = CowBar_Info(1)
'''''        CoBarInf(3) = CowBar_Info(2)
'''''        CoBarInf(4) = CowBar_Info(3)
'''''        CoBarInf(5) = CowBar_Info(4)
'''''        CorRtn = CoSetBarInfo(CoBarInf(0), WorkBuffer(0))
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
        Mid(QREdit1, 14, 1) = "C"
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
        
        For gInt = 1 To .明細行数
            QREdit2 = QREdit2 & RTrim(.背番号(gInt))                      '背番号
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
            'QREdit2 = QREdit2 & Space(5 - CisFun.Len2(.背番号(gInt)))
            QREdit2 = QREdit2 & Space(8 - CisFun.Len2(.背番号(gInt)))
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
            QREdit2 = QREdit2 & RTrim(.部品番号(gInt))                    '部品番号
            QREdit2 = QREdit2 & Space(25 - CisFun.Len2(.部品番号(gInt)))
            QREdit2 = QREdit2 & Format(.収容数(gInt), "000000")           '収容数
            QREdit2 = QREdit2 & Format(.箱数(gInt), "00000")              '箱数（枚数）
            QREdit2 = QREdit2 & Format(.納入数(gInt), "0000000000")       '納入数（発注数）
        Next gInt
    End With
    CXH0390QR = RTrim(QREdit1) & RTrim(QREdit2)
    
End Function


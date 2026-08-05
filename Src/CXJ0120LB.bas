Attribute VB_Name = "CXJ0120LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   出荷指示書 発行
'*
'*   モジュールID  : CXJ0120LB.BAS
'*
'*                  CXJ0120LPrint(Integer)
'*
'*   引数 : Integer  =  0  :新規発行
'*                      1  :再発行
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:
'    If Not CXJ0120LPrint(1) Then Exit Sub
'
'*   作成日  :2004/02/09   By CIS
'*   変更日  :2008/04/09   By CIS 背番号8桁
'*   変更日  :2009/09/21   By CIS 土岐津Ver
'*---------------------------------------------------------------------------------
Dim mPage           As Long

Dim SyoriKb         As Byte
Dim Bk_SjNo         As Long
Dim Ttl_Hako        As Long
Dim Prt_DateTime    As String

Public Function CXJ0120LPrint(Optional RePrint As Integer, Optional PrtMsg As Boolean = False)
    CXJ0120LPrint = False
    App.Title = "出荷指示書"
' 印刷進行画面表示ﾌﾗｸﾞｾｯﾄ
    PrtMsg = True   ' ActiveX使用の為、必ずTrue
    pref_MsgDisp = PrtMsg

'帳票区分 （新規発行 OR 再発行 )
    Select Case RePrint
      Case 0:  SyoriKb = 50
      Case 1:  SyoriKb = 150
    End Select
' 出力データ存在確認
    
    gSL_Select = "Select JY.出荷指示書NO,JY.出荷指示書行NO,JY.納入日,JY.納入便," & vbCr
    gSL_Select = gSL_Select & "JY.納入先,JY.受入," & vbCr
    gSL_Select = gSL_Select & "JY.収容数,JY.背番号,JY.枚数,JY.納入数," & vbCr
    gSL_Select = gSL_Select & "JY.端数,JY.分納区分," & vbCr
    gSL_Select = gSL_Select & "HN.表示品番,HT.SK区分,HT.契約先," & vbCr
    gSL_Select = gSL_Select & "ISNULL(TR1.略称,'') 契約先名," & vbCr
    gSL_Select = gSL_Select & "ISNULL(TR.略称,'') 納入先名," & vbCr
    gSL_Select = gSL_Select & "ISNULL(SY.値名称,'') 受注区分名," & vbCr
    gSL_Select = gSL_Select & "HT.出荷管理部署,ISNULL(TR2.略称,'') 出荷管理部署名" & vbCr
    gSL_Select = gSL_Select & ", ISNULL(HT.置場,'') 工場名 " & vbCr                         ' 2009.11.06 add
    gSL_Select = gSL_Select & ", ISNULL(JY.納入番号,'') 納入番号 " & vbCr                   ' 2009.12.22 add
    gSL_Select = gSL_Select & " From 帳票出力テーブル CH" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 受注テーブル JY" & vbCr
    gSL_Select = gSL_Select & "    ON JY.出荷指示書NO = CH.管理NO" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 品番マスタ HN" & vbCr
    gSL_Select = gSL_Select & "    ON HN.品番 = JY.品番" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 品番取引先マスタ HT" & vbCr
    gSL_Select = gSL_Select & "    ON  HT.品番 = JY.品番" & vbCr
    gSL_Select = gSL_Select & "    And HT.取引先 = JY.納入先" & vbCr
    gSL_Select = gSL_Select & "    And HT.受入 = JY.受入"
    gSL_Select = gSL_Select & "    And HT.品目 = 0" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR1" & vbCr
    gSL_Select = gSL_Select & "    ON TR1.取引先区分 = 0" & vbCr
    gSL_Select = gSL_Select & "    And TR1.取引先CD = HT.契約先" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR" & vbCr
    gSL_Select = gSL_Select & "    ON TR.取引先区分 = 0" & vbCr
    gSL_Select = gSL_Select & "    And TR.取引先CD = JY.納入先" & vbCr
    
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR2" & vbCr
    gSL_Select = gSL_Select & "    ON TR2.取引先区分 = 1" & vbCr
    gSL_Select = gSL_Select & "    And TR2.取引先CD = HT.出荷管理部署" & vbCr
    gSL_Select = gSL_Select & "    And TR2.社内区分 = 1" & vbCr
    
    gSL_Select = gSL_Select & "  Left Outer Join 名称マスタ SY" & vbCr
    gSL_Select = gSL_Select & "    ON  SY.区分名称 = '受注区分'" & vbCr
    gSL_Select = gSL_Select & "    And SY.値 = JY.受注区分" & vbCr
    gSL_Select = gSL_Select & "    And SY.区分ﾀｲﾌﾟ = ''" & vbCr
    gSL_Select = gSL_Select & " Where CH.端末番号 = Host_Name()" & vbCr
    gSL_Select = gSL_Select & "   AND CH.帳票区分 = " & SyoriKb & vbCr
'    gSL_Select = gSL_Select & " Order By JY.納入先,JY.納入日,JY.納入便,HT.SK区分,JY.受注区分,"
    gSL_Select = gSL_Select & "  Order By JY.出荷指示書NO,JY.出荷指示書行NO"
    
    If Not JYTRead(gSL_Select, 0, 1) Then
       Call JYTClose(1)
       CXJ0120LPrint = True
       Exit Function
    End If
    If CisDB.RecordCount = 0 Then
       CXJ0120LPrint = True
       Exit Function
    End If
    pref_TCnt = CisDB.RecordCount
    
'+++++++++++++++++++++++++++++++'
'+  印刷処理
'+++++++++++++++++++++++++++++++'
    Set pref_Form = CXJ0120
    pref_Msg = "【 出 荷 指 示 書 印 刷 中 】"
    
    If pref_MsgDisp Then
        pref_Form.Show vbModal
        Unload pref_Form
        Set pref_Form = Nothing
    Else
        Call CoPrintJ0120
    End If
    
    CXJ0120LPrint = True
End Function
'+----------------------------------------------------------+
'+                    印  刷  処  理                        ＋
'+----------------------------------------------------------+
Public Function CoPrintJ0120() As Boolean
    CoPrintJ0120 = False
    Set CorForm = pref_Form
    
    pref_NCnt = 0                ' 現在発行ページ初期化
    mPage = 1
    On Error GoTo CoPrintJ0120_Err
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
    
    If Not CoReportProcJ0120("O") Then GoTo CoPrintJ0120_ED
    
    Prt_DateTime = Format(Now(), "YYYY/MM/DD hh:nn")
    Bk_SjNo = JYT.出荷指示書NO
    Ttl_Hako = 0
    
' 明細を読み込んでｾｯﾄする(ﾚｺｰﾄﾞ件数分)
    Do Until Not JYT_RDSTS
        
        ' 印刷開始
        If Bk_SjNo <> JYT.出荷指示書NO Then
           mPage = mPage + 1
           If Not CoJ0120MoveT Then GoTo CoPrintJ0120_ED
           If Not CoReportProcJ0120("P") Then GoTo CoPrintJ0120_ED
           Ttl_Hako = 0
        End If
        Bk_SjNo = JYT.出荷指示書NO
        
        ' 項目転送
        If Not CoJ0120Move Then GoTo CoPrintJ0120_ED
        
        Ttl_Hako = Ttl_Hako + JYT.枚数

        pref_NCnt = pref_NCnt + 1
        If pref_MsgDisp Then
            pref_Cnt = Format(pref_NCnt, "###") & " / " & Format(pref_TCnt, "###")
            DoEvents
        End If
        
        ' 受注テーブル　発行日区分更新(新規発行のみ)
        If SyoriKb = 50 Then
            gSL_Select = "UPDATE 受注テーブル SET "
            gSL_Select = gSL_Select & " 指示書発行区分 = 1"
            gSL_Select = gSL_Select & " WHERE 出荷指示書NO = " & JYT.出荷指示書NO
            gSL_Select = gSL_Select & "   AND 出荷指示書行NO = " & JYT.出荷指示書行NO
            With CisDB
                On Error GoTo CoPrintJ0120_Err
                .SQL = gSL_Select
                .DBExec
            End With
        End If
        
        '次のﾚｺｰﾄﾞを読み込む
        Call JYTReadNext(1)
    Loop

    CoPrintJ0120 = True

    If Not CoJ0120MoveT Then GoTo CoPrintJ0120_ED
    If Not CoReportProcJ0120("P") Then GoTo CoPrintJ0120_ED

CoPrintJ0120_ED:
    Call JYTClose(1)
' CoReport 終了処理
    If pref_TCnt <> 0 Then Call CoReportProcJ0120("E")
    
    Exit Function
CoPrintJ0120_Err:
    CisFun.ErrorBox
    End
End Function
'*------------------------------------------------------------------*
'*                                                                  *
'*      出荷指示書　発行内容　項目転送                              *
'*                                                                  *
'*------------------------------------------------------------------*
Private Function CoJ0120Move() As Boolean
    Dim wBar            As String
    Dim wBarNaiyo       As String
    
    CoJ0120Move = True
    
    On Error GoTo CoJ0120Move_Err
    With JYT
'   *------------------*
'   * 『HEAD』SET  *
'   *------------------*
        If .SK区分 = 0 Then
           If Not CorData("SK区分", "かんばん", "S", 8) Then GoTo CoJ0120Move_Err
        Else
           If Not CorData("SK区分", "指    示", "S", 8) Then GoTo CoJ0120Move_Err
        End If
        If Not CorData("受注区分", CisFun.Mid2(.受注区分名, 1, 8), "", 8) Then GoTo CoJ0120Move_Err
        If .分納区分 <> 0 Then
            If Not CorData("分納区分", "分納", "", 6) Then GoTo CoJ0120Move_Err
        End If
            
        gStr = Format(.出荷指示書NO, "0000000000")
        If Not CorData("出荷指示書NO", gStr, "", 10) Then GoTo CoJ0120Move_Err
    
    
    ' 発行日 編集
'        With CisFun
'            .DateE_BefVal = Format(Now(), "yyyymmdd")
'            .DateEdit J_YMD
'            gStr = .DateE_AfVal
'        End With
        If Not CorData("発行日", Prt_DateTime, "", 16) Then GoTo CoJ0120Move_Err
        If Not CorData("PAGE", Format(mPage, "0"), "", 3) Then GoTo CoJ0120Move_Err
    '納入先
        If Not CorData("納入先", .納入先, "", 7) Then GoTo CoJ0120Move_Err
        If Not CorData("納入先名称", CisFun.Left2(.納入先名, 20), "", 20) Then GoTo CoJ0120Move_Err
        If RTrim(.受入) <> "" Then
           If Not CorData("受入", .受入, "", 2) Then GoTo CoJ0120Move_Err
        End If
    '出荷日
        With CisFun
            .DateE_BefVal = JYT.納入日
            .DateEdit J_YMD
            gStr = .DateE_AfVal
        End With
        If Not CorData("納入指示日", gStr, "", 14) Then GoTo CoJ0120Move_Err
        If RTrim(.納入便) <> "" Then
            If Not CorData("納入指示便", RTrim(.納入便) & "便", "", 4) Then GoTo CoJ0120Move_Err
        End If
        
        If Not CorData("契約先", CisFun.Mid2(.契約先名, 1, 20), "", 20) Then GoTo CoJ0120Move_Err

        If Not CorData("出荷管理部署", CisFun.Mid2(.出荷管理部署名, 1, 20), "", 20) Then GoTo CoJ0120Move_Err
'   *------------------*
'   * 『BODY』SET      *
'   *------------------*
        gInt = .出荷指示書行NO
        If gInt > 0 And gInt < 16 Then
           If Not CorListSet("NO", .出荷指示書行NO, gInt, 2) Then GoTo CoJ0120Move_Err
           '*-*-*-*-* 2008/04/09 Update Start
           'If Not CorListSet("背番号", .背番号, gInt, 5) Then GoTo CoJ0120Move_Err
           If Not CorListSet("背番号", .背番号, gInt, 8) Then GoTo CoJ0120Move_Err
           '*-*-*-*-* 2008/04/09 Update End
           
           If Trim(.表示品番) <> "" Then
              If Not CorListSet("品番", .表示品番, gInt, 30) Then GoTo CoJ0120Move_Err
           Else
              If Not CorListSet("品番", .品番, gInt, 30) Then GoTo CoJ0120Move_Err
           End If
           If Not CorListSet("納入番号", Mid(.納入番号, 1, 9), gInt, 9) Then GoTo CoJ0120Move_Err   '　2009.12.22 add
           gStr = Format(.収容数, "##,###")
           If Not CorListSet("収容数", gStr, gInt, 7) Then GoTo CoJ0120Move_Err
           gStr = Format(.枚数, "##,###")
           If Not CorListSet("箱数", gStr, gInt, 6) Then GoTo CoJ0120Move_Err
           gStr = Format(.端数, "##,###")
           If Not CorListSet("端数", gStr, gInt, 7) Then GoTo CoJ0120Move_Err
           gStr = Format(.納入数, "##,###")
           If Not CorListSet("納入数", gStr, gInt, 11) Then GoTo CoJ0120Move_Err
           If Not CorListSet("置場", Mid(.工場名, 1, 10), gInt, 10) Then GoTo CoJ0120Move_Err   '　2009.11.06 add
        End If
    End With

    Exit Function
CoJ0120Move_Err:
    CoJ0120Move = False
End Function

Private Function CoJ0120MoveT() As Boolean
    CoJ0120MoveT = False
'   *------------------*
'   * 『TAIL』SET  *
'   *------------------*
    gStr = Format(Ttl_Hako, "##,###")
    If Not CorData("総箱数", gStr, "", 7) Then GoTo CoJ0120MoveT_Err
    Ttl_Hako = 0
'   *---------------------*
'   * 『バーコード』SET    *
'   *---------------------*
    gStr = Format(Bk_SjNo, "0000000000")
    If CorBar("BAR", gStr) = "ERROR" Then GoTo CoJ0120MoveT_Err
    gStr = "*" & gStr & "*"
    If Not CorData("BAR_Naiyo", gStr, "", 14) Then GoTo CoJ0120MoveT_Err
    
    CoJ0120MoveT = True
    
    Exit Function
CoJ0120MoveT_Err:

End Function
'*==========================================================================*
'*==========================================================================*
'*------------------------------------------------------------------*
'*                                                                  *
'*      CoReport Proccess           (O)Open   (E)Close              *
'*                                                                  *
'*------------------------------------------------------------------*
Public Function CoReportProcJ0120(Repart As String)
    CoReportProcJ0120 = False
'   *******************************************************************
'   * フォーム変更処理                              ( Repart = "P" )  *
'   *******************************************************************
    If Repart = "C" Then
'       CorPrtForm = "検査指示2.FED"
'       If CorChange(G_Crprint_Path & CorPrtForm) Then CoReportProcJ0120 = True
       Exit Function
    End If
'   *******************************************************************
'   * 印刷指示処理                                  ( Repart = "P" )  *
'   *******************************************************************
    If Repart = "P" Then
        If CorData("", "", "E") Then CoReportProcJ0120 = True
    End If
'   *******************************************************************
'   * 終了処理                                      ( Repart = "E" )  *
'   *******************************************************************
    If Repart = "E" Then
        If CorOpCl(crPara, "E") Then CoReportProcJ0120 = True
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
            CorPrtForm = "LXJ0120"
            .csfname = gIniExe & "Prt\" & CorPrtForm & ".FED"
            .csdevm.dmOrientation = DMORIENT_PORTRAIT   'DMORIENT_LANDSCAPE
            .csdevm.dmPaperSize = DMPAPER_A4
            .csdevm.dmFields = DM_TEMPORARY Or DM_PAPERSIZE Or DM_ORIENTATION
        End With
        
        If CorOpCl(crPara, "O") Then
            CoReportProcJ0120 = True
        End If

    End If
End Function

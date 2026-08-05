Attribute VB_Name = "CXH0380TB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   支給チケットファイル出力
'*
'*   モジュールID  : CXH0380TB.BAS
'*
'*                  CXH0380TFile(Integer)
'*
'*   引数 : Integer  =  0  :新規発行
'*                      1  :再発行
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:
'*   If Not CXH0380TFile(1) Then Exit Sub
'*
'*   作成日  : 2005/02/21  By CIS
'*   変更日  : 2006/02/02  By CIS - INIファイル-クライアント対応（TEXT出力先）
'*   変更日  : 2006/02/08  By CIS - 取引先変換マスタ【品番】キー追加に伴う修正
'*   変更日  : 2006/03/09  By CIS - 品番個別設定時、後工程符号参照に間違い
'*   変更日  : 2007/12/18  By CIS - 取引先変換マスタに代行契約先追加
'*   変更日  : 2008/01/16  By CIS - 取引先変換マスタの変更による修正
'*
'*---------------------------------------------------------------------------------
Dim STicket_JisyaCd    As String
Dim STicket_Kiban      As String
Dim STicket_JisyaNm    As String
Dim STicket_OutDev     As String

Dim Jama_JisyaKj    As String
Dim STicket_FileName   As String
Dim STicket_TextFile   As String
Dim STicketRec         As String

Dim SyoriKb         As Byte
'*------------------------------------------------------------------*
'*                                                                  *
'*      支給チケットファイル出力　初期処理                                   *
'*                                                                  *
'*------------------------------------------------------------------*
Private Function STicket_TextInit() As Boolean
    STicket_TextInit = False
'--( 自社コード )
    STicket_JisyaCd = Left(gCompany, 4)
'--( チケット機番 ) [システム情報]
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 15
    Call IRNRead("", 1)
    STicket_Kiban = Mid(IRN.oyian2, 1, 2)
'--( JAMA自社工場・JAMA機番 ) [システム情報]
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 8
    Call IRNRead("", 1)
    Jama_JisyaKj = Mid(IRN.oyian, 1, 1)
'--( 自社名称 )--  取引先マスタより自社名称を求める
    gSL_Select = "SELECT * FROM 取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(gCompany) & "'"
    If TRMRead(gSL_Select, 1) Then
        STicket_JisyaNm = TRM.略称
    Else
        STicket_JisyaNm = gCompName
    End If

'--( TEXT出力先　獲得 )-- [クライアント情報]
    With CisFun
'-----------------------------------------------< 2006/02/02 ADD START >
        .INI_FileID = gIniExe & gwsIniName
'-----------------------------------------------< 2006/02/02 ADD START >
        .INI_Section = "Option"
        .INI_Key = "STICKETFILEOUT"
        .INI_Default = ""
        .GetIni
        STicket_OutDev = RTrim(.INI_String)
        
        If RTrim(.INI_String) = "" Then
            .INI_String = gIniExe
            .PutIni
            STicket_OutDev = .INI_String
        End If
        If Right(RTrim(STicket_OutDev), 1) <> "\" Then
            STicket_OutDev = RTrim(STicket_OutDev) & "\"
        End If
'-----------------------------------------------< 2006/02/02 ADD START >
        .INI_FileID = gIniExe & gIniName
'-----------------------------------------------< 2006/02/02 ADD START >
    End With
    'ディレクトリ存在確認
    With CisFun
        .FE_Path = RTrim(STicket_OutDev)
        .DirCheck Msg_Crt
    End With
'--<< ファイル名　獲得 >>--
    STicket_FileName = "TDZ20101100" & Format(Now(), "yyyymmddhhnnss") & ".HKO"

    STicket_TextInit = True
End Function
Public Function CXH0380TFile(Optional ReCreate As Integer, Optional PrtMsg As Boolean = False)
    CXH0380TFile = False
    
    App.Title = "支給チケットファイル出力"
' 印刷進行画面表示ﾌﾗｸﾞｾｯﾄ
    PrtMsg = True
    pref_MsgDisp = PrtMsg
'帳票区分 （新規発行 OR 再発行 )
    Select Case ReCreate
      Case 0:  SyoriKb = 19
      Case 1:  SyoriKb = 119
    End Select
'--------------'
'   初期設定
'--------------'
    If Not STicket_TextInit Then Exit Function
'-------------------------------'
'   出力データ存在確認
'-------------------------------'
    gSL_Select = "SELECT HC.*,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.取引先,'') 仕入先CD,ISNULL(TR1.工場CD,'') 仕入先工場,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.略称,'') 仕入先名,"
'(↓)--------------------------------------------------<< 2006/02/08 >> Insert&Update Start
'    gSL_Select = gSL_Select & "  CASE WHEN ISNULL(TH.契約取引先,'') = '' THEN TR1.取引先CD "
'    gSL_Select = gSL_Select & "       ELSE ISNULL(TH.契約取引先,'') END 契約取引先,"
'    gSL_Select = gSL_Select & "  CASE WHEN ISNULL(TH.契約工場,'') = '' THEN TR1.工場CD "
'    gSL_Select = gSL_Select & "       ELSE ISNULL(TH.契約工場,'') END 契約工場,"
'    gSL_Select = gSL_Select & "  CASE WHEN ISNULL(TR2.被自社,'') = '' THEN '" & RTrim(STicket_JisyaCd) & "' "
'    gSL_Select = gSL_Select & "       ELSE ISNULL(TR2.被自社,'') END 被自社,"
'    gSL_Select = gSL_Select & "  CASE WHEN ISNULL(TR2.被自社工場,'') = '' THEN '" & RTrim(Jama_JisyaKj) & "' "
'    gSL_Select = gSL_Select & "       ELSE ISNULL(TR2.被自社工場,'') END 被自社工場"
'    gSL_Select = gSL_Select & ", isnull(TH.支給元, '') 支給元"
'    gSL_Select = gSL_Select & ", isnull(TH.支給元2, '') 支給元2"
'    gSL_Select = gSL_Select & ", isnull(TH.代行, '') 代行"
'    gSL_Select = gSL_Select & ", isnull(TH.代行工場, '') 代行工場"
'    gSL_Select = gSL_Select & ", isnull(TH.代行受入, '') 代行受入"
'    gSL_Select = gSL_Select & ", isnull(TH.代行名称, '') 代行名称"
'    gSL_Select = gSL_Select & ", case when isnull(TH.支給元2, '') = ''"
'    gSL_Select = gSL_Select & "       then substring(isnull(SY1.英数字1, ''), 1, 3)"
'    gSL_Select = gSL_Select & "       else substring(isnull(SY2.英数字1, ''), 1, 3)"
'    gSL_Select = gSL_Select & "   end 後工程符号"
    gSL_Select = gSL_Select & "  CASE WHEN ISNULL(TH2.契約先,'') = '' THEN "
    gSL_Select = gSL_Select & "       CASE WHEN ISNULL(TH.契約取引先,'') = '' THEN TR1.取引先CD "
    gSL_Select = gSL_Select & "            ELSE ISNULL(TH.契約取引先,'') END "
    gSL_Select = gSL_Select & "  ELSE "
    gSL_Select = gSL_Select & "       CASE WHEN ISNULL(TH2.契約取引先,'') = '' THEN TR1.取引先CD "
    gSL_Select = gSL_Select & "            ELSE ISNULL(TH2.契約取引先,'') END "
    gSL_Select = gSL_Select & "  END 契約取引先,"
    gSL_Select = gSL_Select & "  CASE WHEN ISNULL(TH2.契約先,'') = '' THEN "
    gSL_Select = gSL_Select & "       CASE WHEN ISNULL(TH.契約工場,'') = '' THEN TR1.工場CD "
    gSL_Select = gSL_Select & "            ELSE ISNULL(TH.契約工場,'') END "
    gSL_Select = gSL_Select & "  ELSE "
    gSL_Select = gSL_Select & "       CASE WHEN ISNULL(TH2.契約工場,'') = '' THEN TR1.工場CD "
    gSL_Select = gSL_Select & "            ELSE ISNULL(TH2.契約工場,'') END "
    gSL_Select = gSL_Select & "  END 契約工場,"
    gSL_Select = gSL_Select & "  CASE WHEN ISNULL(TR2.被自社,'') = '' THEN '" & RTrim(STicket_JisyaCd) & "' "
    gSL_Select = gSL_Select & "       ELSE ISNULL(TR2.被自社,'') END 被自社,"
    gSL_Select = gSL_Select & "  CASE WHEN ISNULL(TR2.被自社工場,'') = '' THEN '" & RTrim(Jama_JisyaKj) & "' "
    gSL_Select = gSL_Select & "       ELSE ISNULL(TR2.被自社工場,'') END 被自社工場"
    gSL_Select = gSL_Select & ", CASE WHEN ISNULL(TH2.契約先,'') = '' THEN "
    gSL_Select = gSL_Select & "       isnull(TH.支給元, '') "
    gSL_Select = gSL_Select & "  ELSE ISNULL(TH2.支給元, '') "
    gSL_Select = gSL_Select & "  END 支給元"
    gSL_Select = gSL_Select & ", CASE WHEN ISNULL(TH2.契約先,'') = '' THEN "
    gSL_Select = gSL_Select & "       isnull(TH.支給元2, '') "
    gSL_Select = gSL_Select & "  ELSE ISNULL(TH2.支給元2, '') "
    gSL_Select = gSL_Select & "  END 支給元2"
    
    gSL_Select = gSL_Select & ", CASE WHEN ISNULL(TH2.契約先,'') = '' THEN "
    gSL_Select = gSL_Select & "       isnull(TH.納入先, '') "
    gSL_Select = gSL_Select & "  ELSE ISNULL(TH2.納入先, '') "
    gSL_Select = gSL_Select & "  END 契約納入先"
    gSL_Select = gSL_Select & ", CASE WHEN ISNULL(TH2.契約先,'') = '' THEN "
    gSL_Select = gSL_Select & "       isnull(TH.納入先工場, '') "
    gSL_Select = gSL_Select & "  ELSE ISNULL(TH2.納入先工場, '') "
    gSL_Select = gSL_Select & "  END 契約納入先工場"
    gSL_Select = gSL_Select & ", CASE WHEN ISNULL(TH2.契約先,'') = '' THEN "
    gSL_Select = gSL_Select & "       isnull(TH.納入先受入, '') "
    gSL_Select = gSL_Select & "  ELSE ISNULL(TH2.納入先受入, '') "
    gSL_Select = gSL_Select & "  END 契約納入先受入"
    gSL_Select = gSL_Select & ", CASE WHEN ISNULL(TH2.契約先,'') = '' THEN "
    gSL_Select = gSL_Select & "       isnull(TH.納入先名称, '') "
    gSL_Select = gSL_Select & "  ELSE ISNULL(TH2.納入先名称, '') "
    gSL_Select = gSL_Select & "  END 契約納入先名称"
    gSL_Select = gSL_Select & ", CASE WHEN ISNULL(TH2.直直区分,0) = 0 THEN "
    gSL_Select = gSL_Select & "       isnull(TH.直直区分,0) "
    gSL_Select = gSL_Select & "  ELSE ISNULL(TH2.直直区分,0) "
    gSL_Select = gSL_Select & "  END 直直区分"
    
'''    gSL_Select = gSL_Select & ", CASE WHEN ISNULL(TH2.契約先,'') = '' THEN "
'''    gSL_Select = gSL_Select & "       isnull(TH.代行, '') "
'''    gSL_Select = gSL_Select & "  ELSE ISNULL(TH2.代行, '') "
'''    gSL_Select = gSL_Select & "  END 代行"
'''    gSL_Select = gSL_Select & ", CASE WHEN ISNULL(TH2.契約先,'') = '' THEN "
'''    gSL_Select = gSL_Select & "       isnull(TH.代行工場, '') "
'''    gSL_Select = gSL_Select & "  ELSE ISNULL(TH2.代行工場, '') "
'''    gSL_Select = gSL_Select & "  END 代行工場"
'''    gSL_Select = gSL_Select & ", CASE WHEN ISNULL(TH2.契約先,'') = '' THEN "
'''    gSL_Select = gSL_Select & "       isnull(TH.代行受入, '') "
'''    gSL_Select = gSL_Select & "  ELSE ISNULL(TH2.代行受入, '') "
'''    gSL_Select = gSL_Select & "  END 代行受入"
'''    gSL_Select = gSL_Select & ", CASE WHEN ISNULL(TH2.契約先,'') = '' THEN "
'''    gSL_Select = gSL_Select & "       isnull(TH.代行名称, '') "
'''    gSL_Select = gSL_Select & "  ELSE ISNULL(TH2.代行名称, '') "
'''    gSL_Select = gSL_Select & "  END 代行名称"
    gSL_Select = gSL_Select & ", CASE WHEN ISNULL(TH2.契約先,'') = '' THEN "
    gSL_Select = gSL_Select & "       case when isnull(TH.支給元2, '') = ''"
    gSL_Select = gSL_Select & "             then substring(isnull(SY1.英数字1, ''), 1, 3)"
    gSL_Select = gSL_Select & "             else substring(isnull(SY2.英数字1, ''), 1, 3)"
    gSL_Select = gSL_Select & "       END "
    gSL_Select = gSL_Select & "  ELSE "
    gSL_Select = gSL_Select & "       case when isnull(TH2.支給元2, '') = ''"
    gSL_Select = gSL_Select & "             then substring(isnull(SY21.英数字1, ''), 1, 3)"
    gSL_Select = gSL_Select & "             else substring(isnull(SY22.英数字1, ''), 1, 3)"
    gSL_Select = gSL_Select & "       END "
    gSL_Select = gSL_Select & "   end 後工程符号"
'(↑)--------------------------------------------------<< 2006/02/08 >> Insert&Update End
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 発注帳票発行テーブル HC"
    gSL_Select = gSL_Select & "   ON  HC.納品書管理NO =CH.管理NO"
    gSL_Select = gSL_Select & "   AND HC.納入番号 = CH.管理文字"
    gSL_Select = gSL_Select & "   AND HC.頁NO = CH.補助番号"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR1"
    gSL_Select = gSL_Select & "   ON  TR1.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR1.取引先CD = HC.仕入先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR2"
    gSL_Select = gSL_Select & "   ON  TR2.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR2.取引先CD = HC.契約先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先変換マスタ暫定 TH"
    gSL_Select = gSL_Select & "   ON  TH.契約先 = HC.契約先"
    gSL_Select = gSL_Select & "   AND TH.代行契約先 = HC.代行契約先"            ' 2007/12/18 ADD
    gSL_Select = gSL_Select & "   AND TH.取引先 = HC.仕入先"
    gSL_Select = gSL_Select & "   AND TH.受入 = HC.受入"
'(↓)--------------------------------------------------<< 2006/02/08 >> Insert Start
    gSL_Select = gSL_Select & "   AND TH.品番 = ''"
'(↑)--------------------------------------------------<< 2006/02/08 >> Insert End
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY1"
    gSL_Select = gSL_Select & "   ON  SY1.区分名称 = '後工程符号管理'"
    gSL_Select = gSL_Select & "   AND SY1.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY1.値 = TH.支給元"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2"
    gSL_Select = gSL_Select & "   ON  SY2.区分名称 = '後工程符号管理'"
    gSL_Select = gSL_Select & "   AND SY2.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY2.値 = TH.支給元2"
'(↓)--------------------------------------------------<< 2006/02/08 >> Insert Start
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先変換マスタ暫定 TH2"
    gSL_Select = gSL_Select & "   ON  TH2.契約先 = HC.契約先"
    gSL_Select = gSL_Select & "   AND TH2.代行契約先 = HC.代行契約先"            ' 2007/12/18 ADD
    gSL_Select = gSL_Select & "   AND TH2.取引先 = HC.仕入先"
    gSL_Select = gSL_Select & "   AND TH2.受入 = HC.受入"
    gSL_Select = gSL_Select & "   AND TH2.品番 = HC.部品番号1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY21"
    gSL_Select = gSL_Select & "   ON  SY21.区分名称 = '後工程符号管理'"
    gSL_Select = gSL_Select & "   AND SY21.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY21.値 = TH2.支給元"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY22"
    gSL_Select = gSL_Select & "   ON  SY22.区分名称 = '後工程符号管理'"
    gSL_Select = gSL_Select & "   AND SY22.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY22.値 = TH2.支給元2"
'(↑)--------------------------------------------------<< 2006/02/08 >> Insert End
    
    gSL_Select = gSL_Select & " WHERE 端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND 帳票区分 = " & SyoriKb
    gSL_Select = gSL_Select & " ORDER BY HC.契約先,HC.仕入先,HC.納入番号,HC.発行順,HC.頁NO"
    
    If Not HCHRead(gSL_Select, 0, 1) Then
       Call HCHClose(1)
       CXH0380TFile = True
       Exit Function
    End If
    If CisDB.RecordCount = 0 Then
       CXH0380TFile = True
       Exit Function
    End If
    pref_TCnt = CisDB.RecordCount
    
'+++++++++++++++++++++++++++++++'
'+  印刷処理
'+++++++++++++++++++++++++++++++'
    Set pref_Form = CXH0380
    pref_Msg = "【 支給チケットファイル  出力中 】"
    
    If pref_MsgDisp Then
        pref_Form.Show vbModal
        Unload pref_Form
        Set pref_Form = Nothing
    Else
        Call STicket_TextCreate
    End If
    
    CXH0380TFile = True
End Function
'+----------------------------------------------------------+
'+                ファイル出力  処理                        ＋
'+----------------------------------------------------------+
Public Function STicket_TextCreate() As Boolean
    STicket_TextCreate = False
    Dim TicketFNo       As Integer
    Dim Key_Keiyaku     As String
    
'    Set CorForm = pref_Form
    
    pref_NCnt = 0                ' 現在発行ページ初期化
    On Error GoTo STicket_TextCreate_Err

'------------------------'
'    出力ﾃｷｽﾄ OPEN
'------------------------'
    STicket_TextFile = RTrim(STicket_OutDev) & Trim(STicket_FileName)
    TicketFNo = FreeFile: gCnt1 = 0
    Open STicket_TextFile For Output As #TicketFNo
' 明細を読み込んでセットする (レコード件数分)
    Do Until Not HCH_RDSTS
        gCnt1 = gCnt1 + 1
    '------------------------'
    '   レコード編集
    '------------------------'
        Call STicket_TextEdit
    '------------------------'
    '   レコード出力
    '------------------------'
        On Error GoTo STicket_TextCreate_Err
        Print #TicketFNo, STicketRec
        
        pref_NCnt = pref_NCnt + 1
        If pref_MsgDisp Then
            pref_Cnt = Format(pref_NCnt, "###") & " / " & Format(pref_TCnt, "###")
            DoEvents
        End If
        Key_Keiyaku = RTrim(HCH.契約先)

    '--------------------------------------------------'
    '   発注帳票発行テーブル　発行日更新(新規発行のみ)
    '--------------------------------------------------'
        If SyoriKb = 19 Then
            gSL_Select = "UPDATE 発注帳票発行テーブル SET "
            gSL_Select = gSL_Select & " 発行日 = CONVERT(VARCHAR(8),GETDATE(),112),"
            gSL_Select = gSL_Select & " TEXTNAME = '" & Trim(STicket_FileName) & "'"
            gSL_Select = gSL_Select & " WHERE 納品書管理NO = " & HCH.納品書管理NO & ""
            gSL_Select = gSL_Select & "   AND 納入番号 = '" & HCH.納入番号 & "'"
            gSL_Select = gSL_Select & "   AND 頁NO = " & HCH.頁NO & ""
            With CisDB
                .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
                On Error GoTo STicket_TextCreate_Err
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
                On Error GoTo STicket_TextCreate_Err
                .SQL = gSL_Select
                .DBExec
                .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
            End With
        End If

        '次のﾚｺｰﾄﾞを読み込む
        Call HCHReadNext(1)
    Loop
'------------------------'
'   (最終)ﾄﾚｲﾗｰ 出力
'------------------------'
    STicketRec = "90" & Format(gCnt1, "0000")
    Print #TicketFNo, STicketRec
'------------------------'
'   出力ﾃｷｽﾄ CLOSE
'------------------------'
    Close #TicketFNo


    STicket_TextCreate = True

STicket_TextCreate_ED:
    Call HCHClose(1)
    
    Exit Function
STicket_TextCreate_Err:
   CisFun.ErrorBox
    End
End Function
'************************************
'*    JAMA帳票発行テキスト編集
'************************************
Private Function STicket_TextEdit() As Boolean
    STicket_TextEdit = False
    
    STicketRec = ""
'>> レコード種別
    If Mid(HCH.納入番号, 5, 1) = "0" Then
        STicketRec = "V"                                '(かんばんチケット）
    Else
        STicketRec = "W"                                '(分割)
    End If
'>> 発行回数フラグ
    If SyoriKb = 19 Then
        STicketRec = STicketRec & "0"                   '(初回）
    Else
        STicketRec = STicketRec & "1"                   '(再発行）
    End If
    With HCH
'   >> 納入先・納入先工場・受入　（契約先から見た自社コード・工場）     2008/01/16変更
        If RTrim$(.契約納入先) = "" Then                                            '※納入先（契約先から見た自社コード）
            If .直直区分 = 0 Then
                STicketRec = STicketRec & Left(.被自社, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.被自社, 1, 4)))         '※納入先（契約先から見た自社コード・工場）
            Else
                STicketRec = STicketRec & Left(.契約先, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.契約先, 1, 4)))         '※納入先（契約先から見た自社コード・工場）
            End If
        Else
            STicketRec = STicketRec & Left(.契約納入先, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.契約納入先, 1, 4)))     '※納入先（契約先から見た自社コード・工場）
        End If
        If RTrim$(.契約納入先工場) = "" Then                                        '※納入先（契約先から見た自社工場）
            STicketRec = STicketRec & Left(.被自社工場, 1)
        Else
            STicketRec = STicketRec & Left(.契約納入先工場, 1)
        End If
        If RTrim$(.契約納入先受入) = "" Then                                        '※納入先（契約先から見た自社受入）
            STicketRec = STicketRec & Left(.受入, 2) & Space(2 - CisFun.Len2(.受入))
        Else
            STicketRec = STicketRec & Left(.契約納入先受入, 2) & Space(2 - CisFun.Len2(.契約納入先受入))
        End If
        
'''        If RTrim$(.代行) = "" Then
'''            If RTrim(.支給元2) = "" Then
'''                STicketRec = STicketRec & Left(.被自社, 4) & Space(4 - CisFun.Len2(.被自社))
'''                STicketRec = STicketRec & Left(.被自社工場, 1) & Space(1 - CisFun.Len2(.被自社工場))
'''            Else
'''                STicketRec = STicketRec & Left(.支給元, 5) & Space(5 - CisFun.Len2(Left(.支給元, 5)))
'''            End If
'''            STicketRec = STicketRec & Left(.受入, 2) & Space(2 - CisFun.Len2(.受入))
'''        Else
''''           代行情報が設定時
'''            If RTrim(.支給元2) = "" Then
'''    '           >> 納入先
'''                STicketRec = STicketRec & CisFun.Mid2(.代行, 1, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.代行, 1, 4)))
'''    '           >> 納入先事業所
'''                STicketRec = STicketRec & CisFun.Mid2(.代行工場, 1, 1) & Space(1 - CisFun.Len2(CisFun.Mid2(.代行工場, 1, 1)))
'''            Else
'''                STicketRec = STicketRec & Left(.支給元, 5) & Space(5 - CisFun.Len2(Left(.支給元, 5)))
'''            End If
'''            STicketRec = STicketRec & .代行受入 & Space(2 - CisFun.Len2(.代行受入))
'''        End If
'   >> 自工程符号
        STicketRec = STicketRec & Space(3)
'   >> 仕入先・仕入先工場
        STicketRec = STicketRec & Left(.契約取引先, 4)                          '※受注者（契約先から見た仕入先コード・工場）
        STicketRec = STicketRec & Space(4 - CisFun.Len2(Left(.契約取引先, 4)))
        STicketRec = STicketRec & Left(.契約工場, 1)
        STicketRec = STicketRec & Space(1 - CisFun.Len2(Left(.契約工場, 1)))
'   >> 仕入先工場　数字変換（1000の時）
        If RTrim(.契約取引先) Like "1000*" Then
            gStr = cCvtX9(.契約工場)
            STicketRec = STicketRec & Left(gStr, 2)
            STicketRec = STicketRec & Space(2 - CisFun.Len2(gStr))
        Else
            STicketRec = STicketRec & Space(2)
        End If
'   >> 出荷場・前工程符号・便区分
        STicketRec = STicketRec & Left(.後工程符号, 3) & Space(3 - CisFun.Len2(Left(.後工程符号, 3)))
        STicketRec = STicketRec & Space(3)
        STicketRec = STicketRec & Space(1)
'   >> 納入番号
        STicketRec = STicketRec & RTrim(.納入番号)
'   >> ページNO
        STicketRec = STicketRec & Format(.頁NO, "00")
'   >> 行連番
        STicketRec = STicketRec & Format(.分納時初回行, "0")
'   >> 発注日・発注便
        STicketRec = STicketRec & RTrim(.発注日)
        STicketRec = STicketRec & Space(8 - CisFun.Len2(.発注日))
        STicketRec = STicketRec & RTrim(.発注便)
        STicketRec = STicketRec & Space(2 - CisFun.Len2(.発注便))
'   >> 納入日・納入便
        STicketRec = STicketRec & RTrim(.納入日)
        STicketRec = STicketRec & Space(8 - CisFun.Len2(.納入日))
        STicketRec = STicketRec & RTrim(.納入便)
        STicketRec = STicketRec & Space(2 - CisFun.Len2(.納入便))
'   >> 発注区分
        STicketRec = STicketRec & Format(.発注区分, "0")
'   >> カード#
        STicketRec = STicketRec & RTrim(.カード)
'   >> ファイルID
        STicketRec = STicketRec & "1"
'   >> 機番
        STicketRec = STicketRec & RTrim(STicket_Kiban) & Space(2 - CisFun.Len2(STicket_Kiban))
'  >> 支給元１・支給元２
        If RTrim(.支給元2) = "" Then
            STicketRec = STicketRec & CisFun.Mid2(.支給元, 1, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.支給元, 1, 4)))
            STicketRec = STicketRec & CisFun.Mid2(.支給元2, 1, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.支給元2, 1, 4)))
        Else
            STicketRec = STicketRec & CisFun.Mid2(.支給元2, 1, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.支給元2, 1, 4)))
            STicketRec = STicketRec & CisFun.Mid2(.支給元, 1, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.支給元, 1, 4)))
        End If
'  >> 担当者
        If RTrim(.担当者) = "" Then
            STicketRec = STicketRec & Space(12)
        Else
            STicketRec = STicketRec & CisFun.Mid2(.担当者, 1, 12) & Space(12 - CisFun.Len2(CisFun.Mid2(.担当者, 1, 12)))
        End If
'  >> 2次元コード区分（固定）
        STicketRec = STicketRec & "2"
'  >> 仕入先名称
        STicketRec = STicketRec & CisFun.Left2(Trim(.仕入先名), 24)                          '受注者名称
        STicketRec = STicketRec & Space(24 - CisFun.Len2(Trim(CisFun.Left2(Trim(.仕入先名), 24))))
'  >> 納入先名称
        If RTrim$(.契約納入先名称) = "" Then
            STicketRec = STicketRec & CisFun.Left2(Trim(STicket_JisyaNm), 22)                      '納入者名称
            STicketRec = STicketRec & Space(22 - CisFun.Len2(Trim(CisFun.Left2(Trim(STicket_JisyaNm), 22))))
        Else
'           代行情報が設定時
            gStr = CisFun.Left2(.契約納入先名称, 22)
            STicketRec = STicketRec & gStr & Space(22 - CisFun.Len2(gStr))
        End If
'  >> 棚№
        STicketRec = STicketRec & Space(4)
'  >> 総箱数
        STicketRec = STicketRec & Format(.箱総数, "0000")
'  >> ページ総数
        STicketRec = STicketRec & Format(.総頁数, "00")
'  >> 集荷便名
        STicketRec = STicketRec & Space(16)
'  >> データ区分（固定）
        STicketRec = STicketRec & "A"
'  >> 区分１（固定）
        STicketRec = STicketRec & "0"
'  >> 区分２（固定）
        STicketRec = STicketRec & "0"
'  >> 区分３（固定）
        STicketRec = STicketRec & "0"
'  >> 明細行数
        STicketRec = STicketRec & Format(.明細行数, "0")
        
        For gInt = 1 To 7
    '  >> 背番号
            STicketRec = STicketRec & Left(.背番号(gInt), 3)
            STicketRec = STicketRec & Space(3 - CisFun.Len2(Left(.背番号(gInt), 3)))
    '  >> 部品番号
            STicketRec = STicketRec & Left(.部品番号(gInt), 12)
            STicketRec = STicketRec & Space(12 - CisFun.Len2(Left(.部品番号(gInt), 12)))
    '  >> 収容数
            If .収容数(gInt) <> 0 Then
                STicketRec = STicketRec & Format(.収容数(gInt), "00000")
            Else
                STicketRec = STicketRec & Space(5)
            End If
    '  >> 箱数
            If .箱数(gInt) <> 0 Then
            STicketRec = STicketRec & Format(.箱数(gInt), "000")
            Else
                STicketRec = STicketRec & Space(3)
            End If
    '  >> 納入数
            If .納入数(gInt) <> 0 Then
                STicketRec = STicketRec & Format(.納入数(gInt), "000000")
            Else
                STicketRec = STicketRec & Space(6)
            End If
    '  >> 打切区分・打切残数・備考
            If .打切区分(gInt) = 9 Then
                STicketRec = STicketRec & "3"
                STicketRec = STicketRec & Space(6)
                STicketRec = STicketRec & "打切完了" & Space(8)
            Else
                STicketRec = STicketRec & "0"
                STicketRec = STicketRec & Space(6)
                STicketRec = STicketRec & Space(16)
            End If
    '  >> 予備
            STicketRec = STicketRec & Space(13)
        Next gInt
'  >> 作成日
        STicketRec = STicketRec & RTrim(.作成日) & Space(8 - CisFun.Len2(.作成日))
'  >> 作成時間
        gStr = Replace(.作成時間, ":", "")
        STicketRec = STicketRec & RTrim(gStr) & Space(6 - CisFun.Len2(gStr))
'   >> ページ連番
        STicketRec = STicketRec & Format(.頁NO, "00")
'   >> 再発行連番
        STicketRec = STicketRec & Space(7)
    End With
    
    STicket_TextEdit = True
End Function
'*--------------------------------------------------------------------------------
'*
'*   文字(1桁)字ー>数字変換(2桁)
'*              A=10,B=11..Z=35
'*   モジュールID    : cCvtX9
'*
'*   初期設定項目    : なし
'*   引数           : Work  --> 変換元値
'*   戻値           : 数値
'*   例   :
'*    I = cCvtX9(Str)
'*
'*   作成日  :1997/12/02   cis K.K
'*   変更日  :
'*---------------------------------------------------------------------------------
Public Function cCvtX9(ByVal work As Variant) As Integer
    Static w9   As Integer
    w9 = 0
    Select Case Trim(work)
           Case "": w9 = 36
           Case "-": w9 = 37
           Case "+": w9 = 38
           Case "#": w9 = 39
           Case "*": w9 = 40
    End Select
    If w9 <> 0 Then
       cCvtX9 = w9
       Exit Function
    End If
    If IsNumeric(work) Then
        ' 数値ならば
        If work <> "" Then
           cCvtX9 = CDbl(work)
        Else
           cCvtX9 = work
        End If
    Else
        ' 数値以外なら
        If work >= "A" And work <= "Z" Then
            cCvtX9 = Asc(work) - 55
        Else
            cCvtX9 = 99
        End If
    End If
End Function


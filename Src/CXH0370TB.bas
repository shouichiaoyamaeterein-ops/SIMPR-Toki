Attribute VB_Name = "CXH0370TB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   JAMAファイル出力
'*
'*   モジュールID  : CXH0370TB.BAS
'*
'*                  CXH0370TFile(Integer)
'*
'*   引数 : Integer  =  0  :新規発行
'*                      1  :再発行
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:
'*   If Not CXH0370TFile(1) Then Exit Sub
'*
'*   作成日  : 2004/02/14  By CIS
'*   変更日  : 2004/09/16  By CIS - 修正（支給元情報＆納入先情報）
'*   変更日  : 2004/09/22  By CIS - 取引先変換マスタのキー変更に伴う修正（受入の追加）
'*   変更日  : 2005/09/27  By CIS - 取引先変換マスタの未登録時、支給元を契約先にする。
'*   変更日  : 2006/02/02  By CIS - INIファイル-クライアント対応（TEXT出力先）
'*   変更日  : 2006/02/08  By CIS - 取引先変換マスタ【品番】キー追加に伴う修正
'*   変更日  : 2006/02/28  By CIS - 直直時の支給基体系修正（支給元１・２が逆）
'*   変更日  : 2006/03/09  By CIS - 品番個別設定時、後工程符号参照に間違い
'*   変更日  : 2006/03/15  By CIS - 被自社＝支給元、支給元＝支給元２に変更
'*   変更日  : 2007/12/18  By CIS - 取引先変換マスタに代行契約先追加
'*   変更日  : 2008/01/15  By CIS - 取引先変換マスタの変更による修正
'*
'*---------------------------------------------------------------------------------
Dim Jama_JisyaCd    As String
Dim Jama_JisyaKj    As String
Dim Jama_JisyaNm    As String
Dim Jama_Kiban      As String
Dim Jama_OutDev     As String
Dim Jama_FileName   As String
Dim Jama_TextFile   As String
Dim JamaRec         As String

Dim SyoriKb         As Byte
'*------------------------------------------------------------------*
'*                                                                  *
'*      JAMAファイル出力　初期処理                                   *
'*                                                                  *
'*------------------------------------------------------------------*
Private Function JAMA_TextInit() As Boolean
    JAMA_TextInit = False
'--( 自社コード )
    Jama_JisyaCd = Left(gCompany, 4)
'--( JAMA自社工場・JAMA機番 ) [システム情報]
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 8
    Call IRNRead("", 1)
    Jama_JisyaKj = Mid(IRN.oyian, 1, 1)
    Jama_Kiban = Mid(IRN.oyian2, 1, 2)
'--( 自社名称 )--  取引先マスタより自社名称を求める
    gSL_Select = "SELECT * FROM 取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(gCompany) & RTrim(Jama_JisyaKj) & "'"
    If TRMRead(gSL_Select, 1) Then
        Jama_JisyaNm = TRM.略称
    Else
        Jama_JisyaNm = gCompName
    End If

'--( TEXT出力先　獲得 )-- [クライアント情報]
    With CisFun
'-----------------------------------------------< 2006/02/02 ADD START >
        .INI_FileID = gIniExe & gwsIniName
'-----------------------------------------------< 2006/02/02 ADD START >
        .INI_Section = "Option"
        .INI_Key = "JAMAFILEOUT"
        .INI_Default = ""
        .GetIni
        Jama_OutDev = RTrim(.INI_String)
        
        If RTrim(.INI_String) = "" Then
            .INI_String = gIniExe
            .PutIni
            Jama_OutDev = .INI_String
        End If
        If Right(RTrim(Jama_OutDev), 1) <> "\" Then
            Jama_OutDev = RTrim(Jama_OutDev) & "\"
        End If
'-----------------------------------------------< 2006/02/02 ADD START >
        .INI_FileID = gIniExe & gIniName
'-----------------------------------------------< 2006/02/02 ADD START >
    End With
    'ディレクトリ存在確認
    With CisFun
        .FE_Path = RTrim(Jama_OutDev)
        .DirCheck Msg_Crt
    End With
'--<< ファイル名　獲得 >>--
    Jama_FileName = "JAMA0101100" & Format(Now(), "yyyymmddhhnnss") & ".HKO"

    JAMA_TextInit = True
End Function
Public Function CXH0370TFile(Optional ReCreate As Integer, Optional PrtMsg As Boolean = False)
    CXH0370TFile = False
    
    App.Title = "JAMAファイル出力"
' 印刷進行画面表示ﾌﾗｸﾞｾｯﾄ
    PrtMsg = True
    pref_MsgDisp = PrtMsg
'帳票区分 （新規発行 OR 再発行 )
    Select Case ReCreate
      Case 0:  SyoriKb = 7
      Case 1:  SyoriKb = 107
    End Select
'--------------'
'   初期設定
'--------------'
    If Not JAMA_TextInit Then Exit Function
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
'    gSL_Select = gSL_Select & "  CASE WHEN ISNULL(TR2.被自社,'') = '' THEN '" & RTrim(Jama_JisyaCd) & "' "
'    gSL_Select = gSL_Select & "       ELSE ISNULL(TR2.被自社,'') END 被自社,"
'    gSL_Select = gSL_Select & "  CASE WHEN ISNULL(TR2.被自社工場,'') = '' THEN '" & RTrim(Jama_JisyaKj) & "' "
'    gSL_Select = gSL_Select & "       ELSE ISNULL(TR2.被自社工場,'') END 被自社工場"
''(↓)--------------------------------------------------<< 2004/09/16 >> Insert Start
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
''(↑)--------------------------------------------------<< 2004/09/16 >> Insert End
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
    gSL_Select = gSL_Select & "  CASE WHEN ISNULL(TR2.被自社,'') = '' THEN '" & RTrim(Jama_JisyaCd) & "' "
    gSL_Select = gSL_Select & "       ELSE ISNULL(TR2.被自社,'') END 被自社,"
    gSL_Select = gSL_Select & "  CASE WHEN ISNULL(TR2.被自社工場,'') = '' THEN '" & RTrim(Jama_JisyaKj) & "' "
    gSL_Select = gSL_Select & "       ELSE ISNULL(TR2.被自社工場,'') END 被自社工場"
    gSL_Select = gSL_Select & ",  CASE WHEN ISNULL(TH2.契約先,'') = '' THEN "
    gSL_Select = gSL_Select & "       CASE WHEN ISNULL(TH.直直区分,0) = 0 THEN "
    gSL_Select = gSL_Select & "                     isnull(TH.支給元, '') "
    gSL_Select = gSL_Select & "       ELSE "
    gSL_Select = gSL_Select & "                     ISNULL(TH.支給元2,'') END"
    gSL_Select = gSL_Select & "  ELSE "
    gSL_Select = gSL_Select & "       CASE WHEN ISNULL(TH2.直直区分,0) = 0 THEN     "
    gSL_Select = gSL_Select & "                     ISNULL(TH2.支給元,'') "
    gSL_Select = gSL_Select & "       ELSE "
    gSL_Select = gSL_Select & "                     ISNULL(TH2.支給元2,'') END "
    gSL_Select = gSL_Select & "  END 支給元"
'[][][][][][][][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][][][] Start
'    gSL_Select = gSL_Select & "       isnull(TH.支給元, '') "
'    gSL_Select = gSL_Select & "  ELSE ISNULL(TH2.支給元,'') "
'''    gSL_Select = gSL_Select & "       CASE WHEN ISNULL(TH.支給元2,'') = '' THEN "
'''    gSL_Select = gSL_Select & "                     isnull(TH.支給元, '') "
'''    gSL_Select = gSL_Select & "       ELSE "
'''    gSL_Select = gSL_Select & "                     ISNULL(TH.被自社,'') END"
'''    gSL_Select = gSL_Select & "  ELSE "
'''    gSL_Select = gSL_Select & "       CASE WHEN ISNULL(TH2.支給元2,'') = '' THEN     "
'''    gSL_Select = gSL_Select & "                     ISNULL(TH2.支給元,'') "
'''    gSL_Select = gSL_Select & "       ELSE "
'''    gSL_Select = gSL_Select & "                     ISNULL(TH2.被自社,'') END "
'[][][][][][][][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][][][] End
    gSL_Select = gSL_Select & ", CASE WHEN ISNULL(TH2.契約先,'') = '' THEN "
    gSL_Select = gSL_Select & "       CASE WHEN ISNULL(TH.直直区分,0) = 0 THEN "
    gSL_Select = gSL_Select & "                     '' "
    gSL_Select = gSL_Select & "       ELSE "
    gSL_Select = gSL_Select & "                     ISNULL(TH.支給元,'') END"
    gSL_Select = gSL_Select & "  ELSE "
    gSL_Select = gSL_Select & "       CASE WHEN ISNULL(TH2.直直区分,0) = 0 THEN     "
    gSL_Select = gSL_Select & "                     '' "
    gSL_Select = gSL_Select & "       ELSE "
    gSL_Select = gSL_Select & "                     ISNULL(TH2.支給元,'') END "
'[][][][][][][][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][][][] Start
'    gSL_Select = gSL_Select & "       isnull(TH.支給元2, '') "
'    gSL_Select = gSL_Select & "  ELSE ISNULL(TH2.支給元2,'') "
'''    gSL_Select = gSL_Select & "       CASE WHEN ISNULL(TH.支給元2,'') = '' THEN "
'''    gSL_Select = gSL_Select & "                     ''  "
'''    gSL_Select = gSL_Select & "       ELSE "
'''    gSL_Select = gSL_Select & "                     ISNULL(TH.支給元,'') END"
'''    gSL_Select = gSL_Select & "  ELSE "
'''    gSL_Select = gSL_Select & "       CASE WHEN ISNULL(TH2.支給元2,'') = '' THEN     "
'''    gSL_Select = gSL_Select & "                     '' "
'''    gSL_Select = gSL_Select & "       ELSE "
'''    gSL_Select = gSL_Select & "                     ISNULL(TH2.支給元,'') END "
'[][][][][][][][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][][][] End
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
    
    gSL_Select = gSL_Select & ", CASE WHEN ISNULL(TH2.契約先,'') = '' THEN "
    gSL_Select = gSL_Select & "       case when isnull(TH.支給元2, '') = ''"
    gSL_Select = gSL_Select & "            then substring(isnull(SY1.英数字1, ''), 1, 3)"
    gSL_Select = gSL_Select & "            else substring(isnull(SY2.英数字1, ''), 1, 3)"
    gSL_Select = gSL_Select & "       END"
    gSL_Select = gSL_Select & "  ELSE"
    gSL_Select = gSL_Select & "       case when isnull(TH2.支給元2, '') = ''"
    gSL_Select = gSL_Select & "            then substring(isnull(SY21.英数字1, ''), 1, 3)"
    gSL_Select = gSL_Select & "            else substring(isnull(SY22.英数字1, ''), 1, 3)"
    gSL_Select = gSL_Select & "       END"
    gSL_Select = gSL_Select & "  END 後工程符号"
'(↑)--------------------------------------------------<< 2006/02/08 >> Insert&Update End

    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 発注帳票発行テーブル HC"
    gSL_Select = gSL_Select & "   ON  HC.納品書管理NO =CH.管理NO"
    gSL_Select = gSL_Select & "   AND HC.納入番号 = CH.管理文字"
    gSL_Select = gSL_Select & "   AND HC.カード = CH.検索情報"
    gSL_Select = gSL_Select & "   AND HC.頁NO = CH.補助番号"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR1"
    gSL_Select = gSL_Select & "   ON  TR1.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR1.取引先CD = HC.仕入先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR2"
    gSL_Select = gSL_Select & "   ON  TR2.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR2.取引先CD = HC.契約先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先変換マスタ TH"
    gSL_Select = gSL_Select & "   ON  TH.契約先 = HC.契約先"
    gSL_Select = gSL_Select & "   AND TH.代行契約先 = HC.代行契約先"           ' 2007/12/18 ADD
    gSL_Select = gSL_Select & "   AND TH.取引先 = HC.仕入先"
'(↓)--------------------------------------------------<< 2004/09/22 >> Insert Start
    gSL_Select = gSL_Select & "   AND TH.受入 = HC.受入"
'(↑)--------------------------------------------------<< 2004/09/22 >> Insert End
'(↓)--------------------------------------------------<< 2006/02/08 >> Insert Start
    gSL_Select = gSL_Select & "   AND TH.品番 = ''"
'(↑)--------------------------------------------------<< 2006/02/08 >> Insert End
'(↓)--------------------------------------------------<< 2004/09/16 >> Insert Start
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY1"
    gSL_Select = gSL_Select & "   ON  SY1.区分名称 = '後工程符号管理'"
    gSL_Select = gSL_Select & "   AND SY1.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY1.値 = TH.支給元"                      '2006/03/15
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2"
    gSL_Select = gSL_Select & "   ON  SY2.区分名称 = '後工程符号管理'"
    gSL_Select = gSL_Select & "   AND SY2.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY2.値 = TH.支給元2"                     '2006/03/15
'(↑)--------------------------------------------------<< 2004/09/16 >> Insert End

'(↓)--------------------------------------------------<< 2006/02/08 >> Insert Start
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先変換マスタ TH2"
    gSL_Select = gSL_Select & "   ON  TH2.契約先 = HC.契約先"
    gSL_Select = gSL_Select & "   AND TH2.代行契約先 = HC.代行契約先"           ' 2007/12/18 ADD
    gSL_Select = gSL_Select & "   AND TH2.取引先 = HC.仕入先"
    gSL_Select = gSL_Select & "   AND TH2.受入 = HC.受入"
    gSL_Select = gSL_Select & "   AND TH2.品番 = HC.部品番号1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY21"
    gSL_Select = gSL_Select & "   ON  SY21.区分名称 = '後工程符号管理'"
    gSL_Select = gSL_Select & "   AND SY21.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY21.値 = TH2.支給元"                    '2006/03/15
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY22"
    gSL_Select = gSL_Select & "   ON  SY22.区分名称 = '後工程符号管理'"
    gSL_Select = gSL_Select & "   AND SY22.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY22.値 = TH2.支給元2"                   '2006/03/15
'(↑)--------------------------------------------------<< 2006/02/08 >> Insert End
    
    gSL_Select = gSL_Select & " WHERE 端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND 帳票区分 = " & SyoriKb
    gSL_Select = gSL_Select & " ORDER BY HC.契約先,HC.仕入先,HC.納入番号,HC.発行順,HC.頁NO"
    
    If Not HCHRead(gSL_Select, 0, 1) Then
       Call HCHClose(1)
       CXH0370TFile = True
       Exit Function
    End If
    If CisDB.RecordCount = 0 Then
       CXH0370TFile = True
       Exit Function
    End If
    pref_TCnt = CisDB.RecordCount
    
'+++++++++++++++++++++++++++++++'
'+  印刷処理
'+++++++++++++++++++++++++++++++'
    Set pref_Form = CXH0370
    pref_Msg = "【 ＪＡＭＡファイル  出力中 】"
    
    If pref_MsgDisp Then
        pref_Form.Show vbModal
        Unload pref_Form
        Set pref_Form = Nothing
    Else
        Call Jama_TextCreate
    End If
    
    CXH0370TFile = True
End Function
'+----------------------------------------------------------+
'+                ファイル出力  処理                        ＋
'+----------------------------------------------------------+
Public Function Jama_TextCreate() As Boolean
    Jama_TextCreate = False
    Dim JamaFNo         As Integer
    Dim Key_Keiyaku     As String
    
'    Set CorForm = pref_Form
    
    pref_NCnt = 0                ' 現在発行ページ初期化
    On Error GoTo Jama_TextCreate_Err

'------------------------'
'    出力ﾃｷｽﾄ OPEN
'------------------------'
    Jama_TextFile = RTrim(Jama_OutDev) & Trim(Jama_FileName)
    JamaFNo = FreeFile: gCnt1 = 0
    Open Jama_TextFile For Output As #JamaFNo
' 明細を読み込んでセットする (レコード件数分)
    Do Until Not HCH_RDSTS
'''        If RTrim(Key_Keiyaku) <> RTrim(HCH.契約先) Then
'''        '(( 契約先変更 ))
'''            If RTrim(Key_Keiyaku) <> "" Then
'''            '------------------------'
'''            '   ﾄﾚｲﾗｰ 出力
'''            '------------------------'
'''                JamaRec = "90" & Format(gCnt1, "0000")
'''                Print #JamaFNo, JamaRec
'''            '------------------------'
'''            '   出力ﾃｷｽﾄ CLOSE
'''            '------------------------'
'''                Close #JamaFNo
'''            End If
'''        '---------------------------------------------'
'''        '   ディレクトリ存在確認（契約先単位フォルダ）
'''        '---------------------------------------------'
'''            With CisFun
'''                .FE_Path = RTrim(Jama_OutDev) & Trim(HCH.契約先)
'''                .DirCheck NoMsg_Crt
'''            End With
'''        '------------------------'
'''        '    出力ﾃｷｽﾄ OPEN
'''        '------------------------'
'''            Jama_TextFile = RTrim(Jama_OutDev) & Trim(HCH.契約先) & "\" & Trim(Jama_FileName)
'''            JamaFNo = FreeFile: gCnt1 = 0
'''            Open Jama_TextFile For Output As #JamaFNo
'''        End If
        gCnt1 = gCnt1 + 1
    '------------------------'
    '   レコード編集
    '------------------------'
        Call Jama_TextEdit
    '------------------------'
    '   レコード出力
    '------------------------'
        On Error GoTo Jama_TextCreate_Err
        Print #JamaFNo, JamaRec
        
        pref_NCnt = pref_NCnt + 1
        If pref_MsgDisp Then
            pref_Cnt = Format(pref_NCnt, "###") & " / " & Format(pref_TCnt, "###")
            DoEvents
        End If
        Key_Keiyaku = RTrim(HCH.契約先)

    '--------------------------------------------------'
    '   発注帳票発行テーブル　発行日更新(新規発行のみ)
    '--------------------------------------------------'
        If SyoriKb = 7 Then
            gSL_Select = "UPDATE 発注帳票発行テーブル SET "
            gSL_Select = gSL_Select & " 発行日 = CONVERT(VARCHAR(8),GETDATE(),112),"
            gSL_Select = gSL_Select & " TEXTNAME = '" & Trim(Jama_FileName) & "'"
            gSL_Select = gSL_Select & " WHERE 納品書管理NO = " & HCH.納品書管理NO & ""
            gSL_Select = gSL_Select & "   AND 納入番号 = '" & HCH.納入番号 & "'"
            gSL_Select = gSL_Select & "   AND 頁NO = " & HCH.頁NO & ""
            gSL_Select = gSL_Select & "   AND カード = '" & RTrim(HCH.カード) & "'"
            With CisDB
                .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
                On Error GoTo Jama_TextCreate_Err
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
                On Error GoTo Jama_TextCreate_Err
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
    JamaRec = "90" & Format(gCnt1, "0000")
    Print #JamaFNo, JamaRec
'------------------------'
'   出力ﾃｷｽﾄ CLOSE
'------------------------'
    Close #JamaFNo


    Jama_TextCreate = True

Jama_TextCreate_ED:
    Call HCHClose(1)
    
    Exit Function
Jama_TextCreate_Err:
    CisFun.ErrorBox
    End
End Function
'************************************
'*    JAMA帳票発行テキスト編集
'************************************
Private Function Jama_TextEdit() As Boolean
    Jama_TextEdit = False
    
    JamaRec = ""
    JamaRec = "J"                                'レコード種別
    If SyoriKb = 7 Then
        JamaRec = JamaRec & "0"                  '発行回数フラグ（初回）
    Else
        JamaRec = JamaRec & "1"                  '発行回数フラグ（再発行）
    End If
    JamaRec = JamaRec & "00001"                   'レイアウトID
    If Mid(HCH.納入番号, 5, 1) = "0" Then
        JamaRec = JamaRec & "0"                   '分割区分(初回）
    Else
        JamaRec = JamaRec & "1"                   '分割区分(分割）
    End If
    With HCH
'(↓)--------------------------------------------------<< 2004/09/16 >> Update Start
'       JamaRec = JamaRec & Left(.被自社, 4) & Space(6)                   '※納入先（契約先から見た自社コード・工場）
'       JamaRec = JamaRec & Left(.被自社工場, 1) & "-" & Left(.受入, 2)               '納入先事業所（工場-受入)
'       JamaRec = JamaRec & Space(2 - CisFun.Len2(.受入)) & Space(6)
'''        If RTrim$(.代行) = "" Then
'''            JamaRec = JamaRec & Left(.被自社, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.被自社, 1, 4))) & Space(6)                   '※納入先（契約先から見た自社コード・工場）
'''            JamaRec = JamaRec & Left(.被自社工場, 1) & "-" & Left(.受入, 2)               '納入先事業所（工場-受入)
'''            JamaRec = JamaRec & Space(2 - CisFun.Len2(.受入))
'''        Else
''''           代行情報が設定時
''''           >> 納入先
'''            JamaRec = JamaRec & CisFun.Mid2(.代行, 1, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.代行, 1, 4)))
'''            JamaRec = JamaRec & Space(6)
''''           >> 納入先事業所
'''            JamaRec = JamaRec & CisFun.Mid2(.代行工場, 1, 1) & Space(1 - CisFun.Len2(CisFun.Mid2(.代行工場, 1, 1)))
'''            JamaRec = JamaRec & "-"
'''            JamaRec = JamaRec & .代行受入 & Space(2 - CisFun.Len2(.代行受入))
'''        End If
        
        If RTrim$(.契約納入先) = "" Then                                            '※納入先（契約先から見た自社コード）
            If .直直区分 = 0 Then
                JamaRec = JamaRec & Left(.被自社, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.被自社, 1, 4))) & Space(6)                   '※納入先（契約先から見た自社コード・工場）
            Else
                JamaRec = JamaRec & Left(.契約先, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.契約先, 1, 4))) & Space(6)                   '※納入先（契約先から見た自社コード・工場）
            End If
        Else
            JamaRec = JamaRec & Left(.契約納入先, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.契約納入先, 1, 4))) & Space(6)                   '※納入先（契約先から見た自社コード・工場）
        End If
        If RTrim$(.契約納入先工場) = "" Then                                        '※納入先（契約先から見た自社工場）
            JamaRec = JamaRec & Left(.被自社工場, 1) & "-"
        Else
            JamaRec = JamaRec & Left(.契約納入先工場, 1) & "-"
        End If
        If RTrim$(.契約納入先受入) = "" Then                                        '※納入先（契約先から見た自社受入）
            JamaRec = JamaRec & Left(.受入, 2) & Space(2 - CisFun.Len2(.受入))
        Else
            JamaRec = JamaRec & Left(.契約納入先受入, 2) & Space(2 - CisFun.Len2(.契約納入先受入))
        End If
'
        If RTrim$(.後工程符号) = "" Then
            JamaRec = JamaRec & Space(6)
        Else
            JamaRec = JamaRec & "-"
            JamaRec = JamaRec & .後工程符号 & Space(3 - CisFun.Len2(.後工程符号))
            JamaRec = JamaRec & Space(2)
        End If
'(↑)--------------------------------------------------<< 2004/09/16 >> Update End
        JamaRec = JamaRec & Space(5)                                                  '納入場所
        JamaRec = JamaRec & Left(.契約取引先, 4)                          '※受注者（契約先から見た仕入先コード・工場）
        JamaRec = JamaRec & Space(4 - CisFun.Len2(Left(.契約取引先, 4))) & Space(6)
        JamaRec = JamaRec & Left(.契約工場, 1)                                        '受注者事業所
        JamaRec = JamaRec & Space(1 - CisFun.Len2(Left(.契約工場, 1))) & Space(9)
        JamaRec = JamaRec & RTrim(.納入番号) & Space(4)                               '納入番号
        JamaRec = JamaRec & Format(.頁NO, "00")                                        'ページNO
        JamaRec = JamaRec & "00"                                                       '行連番
        JamaRec = JamaRec & RTrim(.納入日)                                             '納入日
        JamaRec = JamaRec & RTrim(.納入便)                                             '納入便
        JamaRec = JamaRec & Space(2 - CisFun.Len2(.納入便))
        JamaRec = JamaRec & Format(.発注区分, "0")                                     '発注区分
        JamaRec = JamaRec & RTrim(.カード)                                             'カード#
        JamaRec = JamaRec & RTrim(Jama_Kiban) & Space(3 - CisFun.Len2(Jama_Kiban))     '機番
'(↓)--------------------------------------------------<< 2005/09/27 >> Insert Start
        If RTrim(.支給元) = "" And _
           RTrim(.支給元2) = "" Then
           .支給元 = Left(.契約先, 4)
        End If
'(↑)--------------------------------------------------<< 2004/09/27 >> Insert End
'(↓)--------------------------------------------------<< 2004/09/16,2005/02/23 >> Update Start
'       JamaRec = JamaRec & RTrim(Left(.契約先, 4))                                     '支給元体系
'       JamaRec = JamaRec & Space(4 - CisFun.Len2(RTrim(Left(.契約先, 4))))
'       JamaRec = JamaRec & RTrim(Left(.契約先, 4))                                    '支給元1(契約先）
'       JamaRec = JamaRec & Space(4 - CisFun.Len2(RTrim(Left(.契約先, 4)))) & Space(4)
'       JamaRec = JamaRec & Space(4)                                                   '支給元2(契約先）
'''        If RTrim(.支給元2) = "" Then
    '       >> 支給元体系
            JamaRec = JamaRec & CisFun.Mid2(.支給元, 1, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.支給元, 1, 4)))
    '       >> 支給元１
            JamaRec = JamaRec & CisFun.Mid2(.支給元, 1, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.支給元, 1, 4)))
            JamaRec = JamaRec & Space(4)
    '       >> 支給元２
            JamaRec = JamaRec & CisFun.Mid2(.支給元2, 1, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.支給元2, 1, 4)))
'''        Else
'
'-----     -----     -----     -----     -----     < 2006/02/28 MNT START >     -----     -----     -----     -----
'
'    '       >> 支給元体系
'            JamaRec = JamaRec & CisFun.Mid2(.支給元2, 1, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.支給元2, 1, 4)))
'    '       >> 支給元１
'            JamaRec = JamaRec & CisFun.Mid2(.支給元2, 1, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.支給元2, 1, 4)))
'            JamaRec = JamaRec & Space(4)
'    '       >> 支給元２
'            JamaRec = JamaRec & CisFun.Mid2(.支給元, 1, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.支給元, 1, 4)))
'''    '       >> 支給元体系
'''            JamaRec = JamaRec & CisFun.Mid2(.支給元, 1, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.支給元2, 1, 4)))
'''    '       >> 支給元１
'''            JamaRec = JamaRec & CisFun.Mid2(.支給元, 1, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.支給元2, 1, 4)))
'''            JamaRec = JamaRec & Space(4)
'''    '       >> 支給元２
'''            JamaRec = JamaRec & CisFun.Mid2(.支給元2, 1, 4) & Space(4 - CisFun.Len2(CisFun.Mid2(.支給元, 1, 4)))
'
'-----     -----     -----     -----     -----     < 2006/02/28 MNT  END  >     -----     -----     -----     -----
'
'''        End If
'(↑)--------------------------------------------------<< 2004/09/16,2005/02/23 >> Update End
        JamaRec = JamaRec & Space(12)                                                  '担当者
        JamaRec = JamaRec & CisFun.Left2(Trim(.仕入先名), 24)                          '受注者名称
        JamaRec = JamaRec & Space(24 - CisFun.Len2(Trim(CisFun.Left2(Trim(.仕入先名), 24))))
'(↓)--------------------------------------------------<< 2004/09/16 >> Update Start
'       JamaRec = JamaRec & CisFun.Left2(Trim(Jama_JisyaNm), 24)                      '納入者名称
'       JamaRec = JamaRec & Space(24 - CisFun.Len2(Trim(CisFun.Left2(Trim(Jama_JisyaNm), 24))))
        If RTrim$(.契約納入先名称) = "" Then
            JamaRec = JamaRec & CisFun.Left2(Trim(Jama_JisyaNm), 24)                      '納入者名称
            JamaRec = JamaRec & Space(24 - CisFun.Len2(Trim(CisFun.Left2(Trim(Jama_JisyaNm), 24))))
        Else
'           代行情報が設定時
'           >> 納入先名称
            JamaRec = JamaRec & .契約納入先名称 & Space(24 - CisFun.Len2(.契約納入先名称))
        End If
'(↑)--------------------------------------------------<< 2004/09/13 >> Update End
        JamaRec = JamaRec & Format(.箱総数, "0000")                                    '総箱数
        JamaRec = JamaRec & Format(.総頁数, "00")                                      '総ページ数
        JamaRec = JamaRec & "B"                                                       'データ区分
        JamaRec = JamaRec & RTrim(.作成日)                                             '作成日
        JamaRec = JamaRec & Space(2)                                                   '作成便
        JamaRec = JamaRec & RTrim(.グループNO)                                         'グループNO
        JamaRec = JamaRec & Format(.物理現頁数, "00")                                   '物理現頁数
        JamaRec = JamaRec & Format(.物理総頁数, "00")                                   '物理総頁数
        JamaRec = JamaRec & Format(.END表示区分, "0")                                   'END表示区分
        JamaRec = JamaRec & Format(.ALLEND表示区分, "0")                                'ALLEND表示区分
    
        JamaRec = JamaRec & Chr(&H4) & Space(50)                                      '見出追加QRバーデータ
        JamaRec = JamaRec & Chr(&H4) & Space(29)                                      '追加１次元バーデータ
        JamaRec = JamaRec & "00"                                                       '見出フリー項目数
        
        JamaRec = JamaRec & Format(.明細行数, "00")                                     '明細行数
        For gInt = 1 To .明細行数
            JamaRec = JamaRec & Format(.行(gInt), "00")                                   '行
            JamaRec = JamaRec & RTrim(.背番号(gInt))                                      '背番号
            JamaRec = JamaRec & Space(10 - CisFun.Len2(RTrim(.背番号(gInt))))
            JamaRec = JamaRec & RTrim(.表示品番(gInt))                                    '部品番号
            JamaRec = JamaRec & Space(25 - CisFun.Len2(RTrim(.表示品番(gInt))))
            JamaRec = JamaRec & RTrim(CisFun.Left2(.部品名称(gInt), 20))                  '部品名称
            JamaRec = JamaRec & Space(20 - CisFun.Len2(RTrim(CisFun.Left2(.部品名称(gInt), 20))))
            JamaRec = JamaRec & Format(.収容数(gInt), "0000000000")                       '収容数
            JamaRec = JamaRec & Format(.箱数(gInt), "000")                                '箱数
            JamaRec = JamaRec & Format(.納入数(gInt), "0000000000")                       '納入数
        
            JamaRec = JamaRec & Chr(&H4) & Space(25)                                      '明細追加QRバーデータ
            JamaRec = JamaRec & "0"                                                       '明細フリー項目数
        Next gInt
    End With
    
    Jama_TextEdit = True
End Function


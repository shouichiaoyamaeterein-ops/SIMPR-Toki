Attribute VB_Name = "CXH0400TB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   自給チケットファイル出力
'*
'*   モジュールID  : CXH0400TB.BAS
'*
'*                  CXH0400TFile(Integer)
'*
'*   引数 : Integer  =  0  :新規発行
'*                      1  :再発行
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:
'*   If Not CXH0400TFile(1) Then Exit Sub
'*
'*   作成日  :2005/02/22   By CIS
'*   変更日  :2006/02/02   By CIS - INIファイル-クライアント対応（TEXT出力先）
'*   変更日  :2008/04/11   By CIS - 背番号８桁対応
'*---------------------------------------------------------------------------------
Dim JTicket_JisyaCd      As String
Dim JTicket_JisyaKj      As String
Dim JTicket_JisyaNm      As String
Dim JTicket_JisyaKjNm    As String
Dim JTicket_Kiban        As String
Dim JTicket_OutDev       As String
Dim JTicket_FileName     As String
Dim JTicket_TextFile     As String
Dim JTicketRec           As String

Dim SyoriKb         As Byte
'*------------------------------------------------------------------*
'*                                                                  *
'*      自給チケットファイル出力　初期処理                            *
'*                                                                  *
'*------------------------------------------------------------------*
Private Function Jisya_TextInit() As Boolean
    Jisya_TextInit = False
'--( 自社コード )
    JTicket_JisyaCd = Left(gCompany, 4)
'--( Jisya自社工場・Jisya機番 ) [システム情報]
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 7
    Call IRNRead("", 1)
    JTicket_JisyaKj = Mid(IRN.oyian, 1, 1)
    JTicket_Kiban = Mid(IRN.oyian2, 1, 2)
'--( 自社名称 )--  取引先マスタより自社名称を求める
    gSL_Select = "SELECT * FROM 取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(gCompany) & RTrim(JTicket_JisyaKj) & "'"
    If TRMRead(gSL_Select, 1) Then
        JTicket_JisyaNm = TRM.略称
        JTicket_JisyaKjNm = TRM.工場名
    Else
        JTicket_JisyaNm = gCompName
        JTicket_JisyaKjNm = ""
    End If

'--( TEXT出力先　獲得 )-- [クライアント情報]
    With CisFun
'-----------------------------------------------< 2006/02/02 ADD START >
        .INI_FileID = gIniExe & gwsIniName
'-----------------------------------------------< 2006/02/02 ADD START >
        .INI_Section = "Option"
        .INI_Key = "JTICKETFILEOUT"
        .INI_Default = ""
        .GetIni
        JTicket_OutDev = RTrim(.INI_String)
        
        If RTrim(.INI_String) = "" Then
            .INI_String = gIniExe
            .PutIni
            JTicket_OutDev = .INI_String
        End If
        If Right(RTrim(JTicket_OutDev), 1) <> "\" Then
            JTicket_OutDev = RTrim(JTicket_OutDev) & "\"
        End If
'-----------------------------------------------< 2006/02/02 ADD START >
        .INI_FileID = gIniExe & gIniName
'-----------------------------------------------< 2006/02/02 ADD START >
    End With
    'ディレクトリ存在確認
    With CisFun
        .FE_Path = RTrim(JTicket_OutDev)
        .DirCheck Msg_Crt
    End With
'--<< ファイル名　獲得 >>--
    JTicket_FileName = Format(Now(), "yyyymmddhhnnss") '& ".HKO"

    Jisya_TextInit = True
End Function
Public Function CXH0400TFile(Optional ReCreate As Integer, Optional PrtMsg As Boolean = False)
    CXH0400TFile = False
    
    App.Title = "自給チケットファイル出力"
' 印刷進行画面表示ﾌﾗｸﾞｾｯﾄ
    PrtMsg = True
    pref_MsgDisp = PrtMsg
'帳票区分 （新規発行 OR 再発行 )
    Select Case ReCreate
      Case 0:  SyoriKb = 18
      Case 1:  SyoriKb = 118
    End Select
'--------------'
'   初期設定
'--------------'
    If Not Jisya_TextInit Then Exit Function
'-------------------------------'
'   出力データ存在確認
'-------------------------------'
    gSL_Select = "SELECT HC.*,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.取引先,'') 仕入先CD,ISNULL(TR1.工場CD,'') 仕入先工場,ISNULL(TR1.略称,'') 仕入先名,ISNULL(TR1.工場名,'') 仕入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR2.取引先,'') 契約取引先,ISNULL(TR2.工場CD,'') 契約工場,"
    gSL_Select = gSL_Select & "  ISNULL(TR3.取引先,'') 納入先CD,ISNULL(TR3.工場CD,'') 納入先工場,ISNULL(TR3.略称,'') 納入先名,ISNULL(TR3.工場名,'') 納入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(SY1.値名称,'') 発注区分名,ISNULL(SY2.値名称,'') 手配区分名"
    
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
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR3"
    gSL_Select = gSL_Select & "   ON  TR3.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR3.取引先CD = HC.納入先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY1"
    gSL_Select = gSL_Select & "   ON  SY1.区分名称 = '発注区分'"
    gSL_Select = gSL_Select & "   AND SY1.値 = HC.発注区分"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2"
    gSL_Select = gSL_Select & "   ON  SY2.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "   AND SY2.値 = HC.手配区分"
    gSL_Select = gSL_Select & " WHERE 端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND 帳票区分 = " & SyoriKb
    gSL_Select = gSL_Select & " ORDER BY HC.仕入先,HC.契約先,HC.納入番号,HC.発行順,HC.頁NO"
    
    If Not HCHRead(gSL_Select, 0, 1) Then
       Call HCHClose(1)
       CXH0400TFile = True
       Exit Function
    End If
    If CisDB.RecordCount = 0 Then
       CXH0400TFile = True
       Exit Function
    End If
    pref_TCnt = CisDB.RecordCount
    
'+++++++++++++++++++++++++++++++'
'+  印刷処理
'+++++++++++++++++++++++++++++++'
    Set pref_Form = CXH0400
    pref_Msg = "【 自給チケットファイル  出力中 】"
    
    If pref_MsgDisp Then
        pref_Form.Show vbModal
        Unload pref_Form
        Set pref_Form = Nothing
    Else
        Call JTicket_TextCreate
    End If
    
    CXH0400TFile = True
End Function
'+----------------------------------------------------------+
'+                ファイル出力  処理                        ＋
'+----------------------------------------------------------+
Public Function JTicket_TextCreate() As Boolean
    JTicket_TextCreate = False
    Dim JTicketFNo           As Integer
    Dim Key_Siire     As String
    
'    Set CorForm = pref_Form
    
    pref_NCnt = 0                ' 現在発行ページ初期化
    On Error GoTo JTicket_TextCreate_Err

' 明細を読み込んでセットする (レコード件数分)
    Do Until Not HCH_RDSTS
        If RTrim(Key_Siire) <> RTrim(HCH.仕入先) Then
        '(( 契約先変更 ))
            If RTrim(Key_Siire) <> "" Then
            '------------------------'
            '   ﾄﾚｲﾗｰ 出力
            '------------------------'
                JTicketRec = "JT" & Format(gCnt1, "0000")
                Print #JTicketFNo, JTicketRec
            '------------------------'
            '   出力ﾃｷｽﾄ CLOSE
            '------------------------'
                Close #JTicketFNo
            End If
''        '---------------------------------------------'
''        '   ディレクトリ存在確認（契約先単位フォルダ）
''        '---------------------------------------------'
''            With CisFun
''                .FE_Path = RTrim(JTicket_OutDev) & Trim(HCH.契約先)
''                .DirCheck NoMsg_Crt
''            End With
        '------------------------'
        '    出力ﾃｷｽﾄ OPEN
        '------------------------'
            JTicket_TextFile = RTrim(JTicket_OutDev) & "TD" & JTicket_FileName & RTrim(HCH.仕入先) & ".HKO"
            JTicketFNo = FreeFile: gCnt1 = 0
            Open JTicket_TextFile For Output As #JTicketFNo
        '------------------------'
        '   ヘッダー 出力
        '------------------------'
            If SyoriKb = 18 Then
                JTicketRec = "THB0"
            Else
                JTicketRec = "THB1"
            End If
            JTicketRec = JTicketRec & RTrim(gCompany) & Space(5 - CisFun.Len2(RTrim(gCompany)))
            JTicketRec = JTicketRec & RTrim(HCH.仕入先) & Space(7 - CisFun.Len2(RTrim(HCH.仕入先)))
            JTicketRec = JTicketRec & JTicket_FileName
            Print #JTicketFNo, JTicketRec
        End If
        gCnt1 = gCnt1 + 1
    '------------------------'
    '   レコード編集
    '------------------------'
        Call Jisya_TextEdit
    '------------------------'
    '   レコード出力
    '------------------------'
        On Error GoTo JTicket_TextCreate_Err
        Print #JTicketFNo, JTicketRec
        
        pref_NCnt = pref_NCnt + 1
        If pref_MsgDisp Then
            pref_Cnt = Format(pref_NCnt, "###") & " / " & Format(pref_TCnt, "###")
            DoEvents
        End If
        Key_Siire = RTrim(HCH.仕入先)

    '--------------------------------------------------'
    '   発注帳票発行テーブル　発行日更新(新規発行のみ)
    '--------------------------------------------------'
        If SyoriKb = 18 Then
            gSL_Select = "UPDATE 発注帳票発行テーブル SET "
            gSL_Select = gSL_Select & " 発行日 = CONVERT(VARCHAR(8),GETDATE(),112),"
            gSL_Select = gSL_Select & " TEXTNAME = '" & "JD" & JTicket_FileName & RTrim(HCH.仕入先) & ".HKO" & "'"
            gSL_Select = gSL_Select & " WHERE 納品書管理NO = " & HCH.納品書管理NO & ""
            gSL_Select = gSL_Select & "   AND 納入番号 = '" & HCH.納入番号 & "'"
            gSL_Select = gSL_Select & "   AND 頁NO = " & HCH.頁NO & ""
            gSL_Select = gSL_Select & "   AND カード = '" & RTrim(HCH.カード) & "'"
            With CisDB
                .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
                On Error GoTo JTicket_TextCreate_Err
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
                On Error GoTo JTicket_TextCreate_Err
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
    JTicketRec = "TT" & Format(gCnt1, "0000")
    Print #JTicketFNo, JTicketRec
'------------------------'
'   出力ﾃｷｽﾄ CLOSE
'------------------------'
    Close #JTicketFNo


    JTicket_TextCreate = True

JTicket_TextCreate_ED:
    Call HCHClose(1)
    
    Exit Function
JTicket_TextCreate_Err:
    CisFun.ErrorBox
    End
End Function
'************************************
'*    Jisya帳票発行テキスト編集
'************************************
Private Function Jisya_TextEdit() As Boolean
    Dim JTicketRec1      As String * 89
    Dim JTicketRec2      As String
    Jisya_TextEdit = False
    
    JTicketRec1 = Space(89)
    Mid(JTicketRec1, 1, 1) = "T"                              'レコード種別
    Mid(JTicketRec1, 2, 1) = "B"                              'データ種別
    Mid(JTicketRec1, 3, 3) = Format(gCnt1, "000")             'レコード連番
    With HCH
        Mid(JTicketRec1, 6, 1) = RTrim(.カード)                           'カード#
        Mid(JTicketRec1, 7, 5) = RTrim(.納入番号) & Space(4)              '納入番号
        Mid(JTicketRec1, 12, 2) = Format(.頁NO, "00")                     'ページNO
        Mid(JTicketRec1, 14, 10) = Format(.納品書管理NO, "0000000000")     '納品書管理NO
        Mid(JTicketRec1, 24, 5) = Left(.仕入先CD, 4)                       '※受注者（仕入先）
        Mid(JTicketRec1, 29, 2) = Left(.仕入先工場, 1)                     '受注者事業所（仕入先工場）
        If SyoriKb = 18 Then
            Mid(JTicketRec1, 31, 1) = "0"           '発行区分(新規発行)
        Else
            Mid(JTicketRec1, 31, 1) = "1"           '発行区分(再発行)
        End If
        If Mid(.納入番号, 5, 1) = "0" Then
            Mid(JTicketRec1, 32, 1) = "0"           '分割区分（新規）
        Else
            Mid(JTicketRec1, 32, 1) = "1"           '分割区分(分納）
        End If
        Mid(JTicketRec1, 33, 1) = RTrim(.発注区分)           '発注区分
        Mid(JTicketRec1, 34, 1) = RTrim(.手配区分)           '手配区分
        Mid(JTicketRec1, 35, 1) = Format(.SK区分, "0")       'SK区分
        If RTrim(.手配区分) <> "1" Then
            Mid(JTicketRec1, 36, 5) = RTrim(.契約取引先)       '支給元(契約先)
            Mid(JTicketRec1, 41, 2) = RTrim(.契約工場)         '支給元(契約先工場)
        Else
            Mid(JTicketRec1, 36, 5) = RTrim(JTicket_JisyaCd)       '支給元(自社)
            Mid(JTicketRec1, 41, 2) = RTrim(JTicket_JisyaKj)       '支給元(自社工場)
        End If
        Mid(JTicketRec1, 43, 2) = RTrim(.受入)                '納入者事業所(自社受入）
        If RTrim(.手配区分) <> "1" Then
            Mid(JTicketRec1, 45, 5) = RTrim(JTicket_JisyaCd)       '納入者(自社)
            Mid(JTicketRec1, 50, 2) = RTrim(JTicket_JisyaKj)       '納入者(自社工場)
        Else
            Mid(JTicketRec1, 45, 5) = RTrim(.納入先CD)         '納入者(納入先)
            Mid(JTicketRec1, 50, 2) = RTrim(.納入先工場)       '納入者事業所(納入先工場)
            Mid(JTicketRec1, 52, 2) = RTrim(.納入先受入)       '納入者事業所(納入先工場)
        End If
        Mid(JTicketRec1, 54, 8) = RTrim(.納入日)               '納入日
        Mid(JTicketRec1, 62, 2) = RTrim(.納入便)               '納入便
        Mid(JTicketRec1, 64, 8) = RTrim(.作成日)               '作成日
        Mid(JTicketRec1, 72, 3) = RTrim(JTicket_Kiban)         '機番
        Mid(JTicketRec1, 75, 4) = Format(.箱総数, "0000")      '箱総数
        Mid(JTicketRec1, 79, 2) = Format(.総頁数, "00")        '総頁数
        Mid(JTicketRec1, 81, 3) = RTrim(.グループNO)           'グループNO
        Mid(JTicketRec1, 84, 2) = Format(.物理現頁数, "00")      '物理現頁数
        Mid(JTicketRec1, 86, 2) = Format(.物理総頁数, "00")      '物理総頁数
        Mid(JTicketRec1, 88, 1) = Format(.END表示区分, "0")      'END表示区分
        Mid(JTicketRec1, 89, 1) = Format(.ALLEND表示区分, "0")   'ALLEND表示区分
        
        JTicketRec2 = ""
        '(( 名称 ))
        JTicketRec2 = JTicketRec2 & CisFun.Left2(.発注区分名, 6)   '発注区分名
        JTicketRec2 = JTicketRec2 & Space(6 - CisFun.Len2(CisFun.Left2(.発注区分名, 6)))
        JTicketRec2 = JTicketRec2 & CisFun.Left2(.手配区分名, 6)   '手配区分名
        JTicketRec2 = JTicketRec2 & Space(6 - CisFun.Len2(CisFun.Left2(.手配区分名, 6)))
        If .SK区分 = 0 Then
            JTicketRec2 = JTicketRec2 & "かんばん"   'SK区分名(かんばん）
        Else
            JTicketRec2 = JTicketRec2 & "指示部品"   'SK区分名(指示）
        End If
        JTicketRec2 = JTicketRec2 & CisFun.Left2(.仕入先名, 24)   '仕入先名
        JTicketRec2 = JTicketRec2 & Space(24 - CisFun.Len2(CisFun.Left2(.仕入先名, 24)))
        JTicketRec2 = JTicketRec2 & CisFun.Left2(.仕入先工場名, 16)   '仕入先工場名
        JTicketRec2 = JTicketRec2 & Space(16 - CisFun.Len2(CisFun.Left2(.仕入先工場名, 16)))
        If RTrim(.手配区分) <> "1" Then
            JTicketRec2 = JTicketRec2 & CisFun.Left2(JTicket_JisyaNm, 24)  '納入先名（自社）
            JTicketRec2 = JTicketRec2 & Space(24 - CisFun.Len2(CisFun.Left2(JTicket_JisyaNm, 24)))
            JTicketRec2 = JTicketRec2 & CisFun.Left2(JTicket_JisyaKjNm, 16)  '納入先工場名（自社）
            JTicketRec2 = JTicketRec2 & Space(16 - CisFun.Len2(CisFun.Left2(JTicket_JisyaKjNm, 16)))
        Else
            JTicketRec2 = JTicketRec2 & CisFun.Left2(.納入先名, 24)   '納入先名
            JTicketRec2 = JTicketRec2 & Space(24 - CisFun.Len2(CisFun.Left2(.納入先名, 24)))
            JTicketRec2 = JTicketRec2 & CisFun.Left2(.納入先工場名, 16)   '納入先工場名
            JTicketRec2 = JTicketRec2 & Space(16 - CisFun.Len2(CisFun.Left2(.納入先工場名, 16)))
        End If
        
        JTicketRec2 = JTicketRec2 & Format(.明細行数, "00")      '明細行数
    
        For gInt = 1 To .明細行数
            JTicketRec2 = JTicketRec2 & RTrim(.背番号(gInt))                      '背番号
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
            'JTicketRec2 = JTicketRec2 & Space(5 - CisFun.Len2(.背番号(gInt)))
            JTicketRec2 = JTicketRec2 & Space(8 - CisFun.Len2(.背番号(gInt)))
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
            JTicketRec2 = JTicketRec2 & RTrim(.部品番号(gInt))                    '部品番号
            JTicketRec2 = JTicketRec2 & Space(25 - CisFun.Len2(.部品番号(gInt)))
            HNM.品番 = RTrim(.部品番号(gInt))
            If HNMRead("", 1) Then
                JTicketRec2 = JTicketRec2 & RTrim(HNM.品番編集)                   '品番編集
            End If
            JTicketRec2 = JTicketRec2 & Space(5 - CisFun.Len2(HNM.品番編集))
            JTicketRec2 = JTicketRec2 & RTrim(.部品名称(gInt))                    '部品名称
            JTicketRec2 = JTicketRec2 & Space(40 - CisFun.Len2(.部品名称(gInt)))
            JTicketRec2 = JTicketRec2 & Format(.収容数(gInt), "000000")           '収容数
            JTicketRec2 = JTicketRec2 & Format(.箱数(gInt), "00000")              '箱数（枚数）
            JTicketRec2 = JTicketRec2 & Format(.端数(gInt), "000000")             '端数
            JTicketRec2 = JTicketRec2 & Format(.納入数(gInt), "0000000000")       '納入数（発注数）
        
        Next gInt
        
        JTicketRec = RTrim(JTicketRec1) & RTrim(JTicketRec2)
    End With
    
    Jisya_TextEdit = True
End Function



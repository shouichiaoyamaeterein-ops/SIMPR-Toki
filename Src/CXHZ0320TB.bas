Attribute VB_Name = "CXHZ0320TB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   [材料]自社伝票ファイル出力
'*
'*   モジュールID  : CXHZ0320TB.BAS
'*
'*                  CXHZ0320TFile(Integer)
'*
'*   引数 : Integer  =  0  :新規発行
'*                      1  :再発行
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:
'*   If Not CXHZ0320TFile(1) Then Exit Sub
'*
'*   作成日  :2004/02/23   By CIS
'*   変更日  :2004/04/06   By CIS - 自社が取引先マスタに未登録時、クリアを行う。
'*---------------------------------------------------------------------------------
Dim My_JisyaCd      As String
Dim My_JisyaKj      As String
Dim My_JisyaNm      As String
Dim My_JisyaKjNm    As String
Dim My_Kiban        As String
Dim My_OutDev       As String
Dim My_FileName     As String
Dim My_TextFile     As String
Dim MyRec           As String

Dim SyoriKb         As Byte
'*------------------------------------------------------------------*
'*                                                                  *
'*      Jisyaファイル出力　初期処理                                   *
'*                                                                  *
'*------------------------------------------------------------------*
Private Function Jisya_TextInit() As Boolean
    Jisya_TextInit = False
'--( 自社コード )
    My_JisyaCd = Left(gCompany, 4)
'--( Jisya自社工場・Jisya機番 ) [システム情報]
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 7
    Call IRNRead("", 1)
    My_JisyaKj = Mid(IRN.oyian, 1, 1)
    My_Kiban = Mid(IRN.oyian2, 1, 2)
'--( 自社名称 )--  取引先マスタより自社名称を求める
    gSL_Select = "SELECT * FROM 取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(gCompany) & RTrim(My_JisyaKj) & "'"
    If TRMRead(gSL_Select, 1) Then
        My_JisyaNm = TRM.略称
        My_JisyaKjNm = TRM.工場名
    Else
        My_JisyaNm = gCompName
'*---------------------------------------------<< 2004/04/06 Insert Start >>
        My_JisyaKjNm = ""
'*---------------------------------------------<< 2004/04/06 Insert End >>
    End If

'--( TEXT出力先　獲得 )-- [クライアント情報]
    With CisFun
        .INI_Section = "Option"
        .INI_Key = "DENPYOFILEOUT"
        .INI_Default = ""
        .GetIni
        My_OutDev = RTrim(.INI_String)
        
        If RTrim(.INI_String) = "" Then
            .INI_String = gIniExe
            .PutIni
            My_OutDev = .INI_String
        End If
        If Right(RTrim(My_OutDev), 1) <> "\" Then
            My_OutDev = RTrim(My_OutDev) & "\"
        End If
    End With
    'ディレクトリ存在確認
    With CisFun
        .FE_Path = RTrim(My_OutDev)
        .DirCheck Msg_Crt
    End With
'--<< ファイル名　獲得 >>--
    My_FileName = Format(Now(), "yyyymmddhhnnss") '& ".HKO"

    Jisya_TextInit = True
End Function
Public Function CXHZ0320TFile(Optional ReCreate As Integer, Optional PrtMsg As Boolean = False)
    CXHZ0320TFile = False
    
    App.Title = "[材料]自社伝票ファイル出力"
' 印刷進行画面表示ﾌﾗｸﾞｾｯﾄ
    PrtMsg = True
    pref_MsgDisp = PrtMsg
'帳票区分 （新規発行 OR 再発行 )
    Select Case ReCreate
      Case 0:  SyoriKb = 26
      Case 1:  SyoriKb = 126
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
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料発注帳票発行テーブル HC"
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
    
    If Not ZCHRead(gSL_Select, 0, 1) Then
       Call ZCHClose(1)
       CXHZ0320TFile = True
       Exit Function
    End If
    If CisDB.RecordCount = 0 Then
       CXHZ0320TFile = True
       Exit Function
    End If
    pref_TCnt = CisDB.RecordCount
    
'+++++++++++++++++++++++++++++++'
'+  印刷処理
'+++++++++++++++++++++++++++++++'
    Set pref_Form = CXHZ0320
    pref_Msg = "【 自社伝票ファイル  出力中 】"
    
    If pref_MsgDisp Then
        pref_Form.Show vbModal
        Unload pref_Form
        Set pref_Form = Nothing
    Else
        Call Jisya_TextCreate
    End If
    
    CXHZ0320TFile = True
End Function
'+----------------------------------------------------------+
'+                ファイル出力  処理                        ＋
'+----------------------------------------------------------+
Public Function Jisya_TextCreate() As Boolean
    Jisya_TextCreate = False
    Dim MyFNo           As Integer
    Dim Key_Siire     As String
    
    Set CorForm = pref_Form
    
    pref_NCnt = 0                ' 現在発行ページ初期化
    On Error GoTo Jisya_TextCreate_Err

' 明細を読み込んでセットする (レコード件数分)
    Do Until Not ZCH_RDSTS
        If RTrim(Key_Siire) <> RTrim(ZCH.仕入先) Then
        '(( 契約先変更 ))
            If RTrim(Key_Siire) <> "" Then
            '------------------------'
            '   ﾄﾚｲﾗｰ 出力
            '------------------------'
                MyRec = "JT" & Format(gCnt1, "0000")
                Print #MyFNo, MyRec
            '------------------------'
            '   出力ﾃｷｽﾄ CLOSE
            '------------------------'
                Close #MyFNo
            End If
''        '---------------------------------------------'
''        '   ディレクトリ存在確認（契約先単位フォルダ）
''        '---------------------------------------------'
''            With CisFun
''                .FE_Path = RTrim(My_OutDev) & Trim(zch.契約先)
''                .DirCheck NoMsg_Crt
''            End With
        '------------------------'
        '    出力ﾃｷｽﾄ OPEN
        '------------------------'
            My_TextFile = RTrim(My_OutDev) & "JZ" & My_FileName & RTrim(ZCH.仕入先) & ".HKO"
            MyFNo = FreeFile: gCnt1 = 0
            Open My_TextFile For Output As #MyFNo
        '------------------------'
        '   ヘッダー 出力
        '------------------------'
            If SyoriKb = 26 Then
                MyRec = "JHZ0"
            Else
                MyRec = "JHZ1"
            End If
            MyRec = MyRec & RTrim(gCompany) & Space(5 - CisFun.Len2(RTrim(gCompany)))
            MyRec = MyRec & RTrim(ZCH.仕入先) & Space(7 - CisFun.Len2(RTrim(ZCH.仕入先)))
            MyRec = MyRec & My_FileName
            Print #MyFNo, MyRec
        End If
        gCnt1 = gCnt1 + 1
    '------------------------'
    '   レコード編集
    '------------------------'
        Call Jisya_TextEdit
    '------------------------'
    '   レコード出力
    '------------------------'
        On Error GoTo Jisya_TextCreate_Err
        Print #MyFNo, MyRec
        
        pref_NCnt = pref_NCnt + 1
        If pref_MsgDisp Then
            pref_Cnt = Format(pref_NCnt, "###") & " / " & Format(pref_TCnt, "###")
            DoEvents
        End If
        Key_Siire = RTrim(ZCH.仕入先)

    '--------------------------------------------------'
    '   材料発注帳票発行テーブル　発行日更新(新規発行のみ)
    '--------------------------------------------------'
        If SyoriKb = 26 Then
            gSL_Select = "UPDATE 材料発注帳票発行テーブル SET "
            gSL_Select = gSL_Select & " 発行日 = CONVERT(VARCHAR(8),GETDATE(),112),"
            gSL_Select = gSL_Select & " TEXTNAME = '" & "JZ" & My_FileName & RTrim(ZCH.仕入先) & ".HKO" & "'"
            gSL_Select = gSL_Select & " WHERE 納品書管理NO = " & ZCH.納品書管理NO & ""
            gSL_Select = gSL_Select & "   AND 納入番号 = '" & ZCH.納入番号 & "'"
            gSL_Select = gSL_Select & "   AND 頁NO = " & ZCH.頁NO & ""
            gSL_Select = gSL_Select & "   AND カード = '" & RTrim(ZCH.カード) & "'"
            With CisDB
                .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
                On Error GoTo Jisya_TextCreate_Err
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
                On Error GoTo Jisya_TextCreate_Err
                .SQL = gSL_Select
                .DBExec
                .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
            End With
        End If

        '次のﾚｺｰﾄﾞを読み込む
        Call ZCHReadNext(1)
    Loop
'------------------------'
'   (最終)ﾄﾚｲﾗｰ 出力
'------------------------'
    MyRec = "JT" & Format(gCnt1, "0000")
    Print #MyFNo, MyRec
'------------------------'
'   出力ﾃｷｽﾄ CLOSE
'------------------------'
    Close #MyFNo


    Jisya_TextCreate = True

Jisya_TextCreate_ED:
    Call ZCHClose(1)
    
    Exit Function
Jisya_TextCreate_Err:
    CisFun.ErrorBox
    End
End Function
'************************************
'*    Jisya帳票発行テキスト編集
'************************************
Private Function Jisya_TextEdit() As Boolean
    Dim MyRec1      As String * 89
    Dim MyRec2      As String
    Jisya_TextEdit = False
    
    MyRec1 = Space(89)
    Mid(MyRec1, 1, 1) = "J"                              'レコード種別
    Mid(MyRec1, 2, 1) = "B"                              'データ種別
    Mid(MyRec1, 3, 3) = Format(gCnt1, "000")             'レコード連番
    With ZCH
        Mid(MyRec1, 6, 1) = RTrim(.カード)                           'カード#
        Mid(MyRec1, 7, 5) = RTrim(.納入番号) & Space(4)              '納入番号
        Mid(MyRec1, 12, 2) = Format(.頁NO, "00")                     'ページNO
        Mid(MyRec1, 14, 10) = Format(.納品書管理NO, "0000000000")     '納品書管理NO
        Mid(MyRec1, 24, 5) = Left(.仕入先CD, 4)                       '※受注者（仕入先）
        Mid(MyRec1, 29, 2) = Left(.仕入先工場, 1)                     '受注者事業所（仕入先工場）
        If SyoriKb = 26 Then
            Mid(MyRec1, 31, 1) = "0"           '発行区分(新規発行)
        Else
            Mid(MyRec1, 31, 1) = "1"           '発行区分(再発行)
        End If
        If Mid(.納入番号, 5, 1) = "0" Then
            Mid(MyRec1, 32, 1) = "0"           '分割区分（新規）
        Else
            Mid(MyRec1, 32, 1) = "1"           '分割区分(分納）
        End If
        Mid(MyRec1, 33, 1) = RTrim(.発注区分)           '発注区分
        Mid(MyRec1, 34, 1) = RTrim(.手配区分)           '手配区分
        Mid(MyRec1, 35, 1) = Format(.SK区分, "0")       'SK区分
        If RTrim(.手配区分) <> "1" Then
            Mid(MyRec1, 36, 5) = RTrim(.契約取引先)       '支給元(契約先)
            Mid(MyRec1, 41, 2) = RTrim(.契約工場)         '支給元(契約先工場)
        Else
            Mid(MyRec1, 36, 5) = RTrim(My_JisyaCd)       '支給元(自社)
            Mid(MyRec1, 41, 2) = RTrim(My_JisyaKj)       '支給元(自社工場)
        End If
        Mid(MyRec1, 43, 2) = RTrim(.受入)                '納入者事業所(自社受入）
        If RTrim(.手配区分) <> "1" Then
            Mid(MyRec1, 45, 5) = RTrim(My_JisyaCd)       '納入者(自社)
            Mid(MyRec1, 50, 2) = RTrim(My_JisyaKj)       '納入者(自社工場)
        Else
            Mid(MyRec1, 45, 5) = RTrim(.納入先CD)         '納入者(納入先)
            Mid(MyRec1, 50, 2) = RTrim(.納入先工場)       '納入者事業所(納入先工場)
            Mid(MyRec1, 52, 2) = RTrim(.納入先受入)       '納入者事業所(納入先工場)
        End If
        Mid(MyRec1, 54, 8) = RTrim(.納入日)               '納入日
        Mid(MyRec1, 62, 2) = RTrim(.納入便)               '納入便
        Mid(MyRec1, 64, 8) = RTrim(.作成日)               '作成日
        Mid(MyRec1, 72, 3) = RTrim(My_Kiban)             '機番
        Mid(MyRec1, 75, 4) = Format(.箱総数, "0000")      '箱総数
        Mid(MyRec1, 79, 2) = Format(.総頁数, "00")        '総頁数
        Mid(MyRec1, 81, 3) = RTrim(.グループNO)           'グループNO
        Mid(MyRec1, 84, 2) = Format(.物理現頁数, "00")      '物理現頁数
        Mid(MyRec1, 86, 2) = Format(.物理総頁数, "00")      '物理総頁数
        Mid(MyRec1, 88, 1) = Format(.END表示区分, "0")      'END表示区分
        Mid(MyRec1, 89, 1) = Format(.ALLEND表示区分, "0")   'ALLEND表示区分
        
        MyRec2 = ""
        '(( 名称 ))
        MyRec2 = MyRec2 & CisFun.Left2(.発注区分名, 6)   '発注区分名
        MyRec2 = MyRec2 & Space(6 - CisFun.Len2(CisFun.Left2(.発注区分名, 6)))
        MyRec2 = MyRec2 & CisFun.Left2(.手配区分名, 6)   '手配区分名
        MyRec2 = MyRec2 & Space(6 - CisFun.Len2(CisFun.Left2(.手配区分名, 6)))
        If .SK区分 = 0 Then
            MyRec2 = MyRec2 & "かんばん"   'SK区分名(かんばん）
        Else
            MyRec2 = MyRec2 & "指示部品"   'SK区分名(指示）
        End If
        MyRec2 = MyRec2 & CisFun.Left2(.仕入先名, 24)   '仕入先名
        MyRec2 = MyRec2 & Space(24 - CisFun.Len2(CisFun.Left2(.仕入先名, 24)))
        MyRec2 = MyRec2 & CisFun.Left2(.仕入先工場名, 16)   '仕入先工場名
        MyRec2 = MyRec2 & Space(16 - CisFun.Len2(CisFun.Left2(.仕入先工場名, 16)))
        If RTrim(.手配区分) <> "1" Then
            MyRec2 = MyRec2 & CisFun.Left2(My_JisyaNm, 24)  '納入先名（自社）
            MyRec2 = MyRec2 & Space(24 - CisFun.Len2(CisFun.Left2(My_JisyaNm, 24)))
            MyRec2 = MyRec2 & CisFun.Left2(My_JisyaKjNm, 16)  '納入先工場名（自社）
            MyRec2 = MyRec2 & Space(16 - CisFun.Len2(CisFun.Left2(My_JisyaKjNm, 16)))
        Else
            MyRec2 = MyRec2 & CisFun.Left2(.納入先名, 24)   '納入先名
            MyRec2 = MyRec2 & Space(24 - CisFun.Len2(CisFun.Left2(.納入先名, 24)))
            MyRec2 = MyRec2 & CisFun.Left2(.納入先工場名, 16)   '納入先工場名
            MyRec2 = MyRec2 & Space(16 - CisFun.Len2(CisFun.Left2(.納入先工場名, 16)))
        End If
        If RTrim(.カード) <> "1" And RTrim(.納入先) <> "" Then
            MyRec2 = MyRec2 & CisFun.Left2(.納入先CD, 5)   '手配商社CD(メーカーCD)
            MyRec2 = MyRec2 & Space(5 - CisFun.Len2(CisFun.Left2(.納入先CD, 5)))
            MyRec2 = MyRec2 & CisFun.Left2(.納入先工場, 2)   '手配商社工場CD
            MyRec2 = MyRec2 & Space(2 - CisFun.Len2(CisFun.Left2(.納入先工場, 2)))
            MyRec2 = MyRec2 & CisFun.Left2(.納入先名, 24)   '手配商社名
            MyRec2 = MyRec2 & Space(24 - CisFun.Len2(CisFun.Left2(.納入先名, 24)))
        Else
            MyRec2 = MyRec2 & Space(31)   '手配商社CD、工場CD、名称
        End If
        
        '(( 明細部 ))
        MyRec2 = MyRec2 & Format(.明細行数, "00")      '明細行数
    
        For gInt = 1 To .明細行数
            '材料マスタ参照
            gSL_Select = "SELECT ZT.背番号,ZT.板厚,ZT.幅,ZT.長さ,ZT.材質,"
            gSL_Select = gSL_Select & " ZR.表示寸法,ZR.材種コード,ISNULL(HN.表示品番,'') 品番,"
            gSL_Select = gSL_Select & " CASE WHEN ZR.実数区分 = 1 THEN '1' ELSE '0' END 備考,"
            gSL_Select = gSL_Select & " ISNULL(SY.値名称,'')手配単位名,ISNULL(SY2.値名称,'')材種名"
            gSL_Select = gSL_Select & " FROM 材料取引先マスタ ZT"
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料マスタ ZR"
            gSL_Select = gSL_Select & "   ON  ZR.材料管理番号 = ZT.材料管理番号"
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料品番マスタ ZN"
            gSL_Select = gSL_Select & "   ON  ZN.材料管理番号 = ZT.材料管理番号"
            gSL_Select = gSL_Select & "   AND ZN.使用品番 = (SELECT MIN(使用品番) FROM 材料品番マスタ HN2"
            gSL_Select = gSL_Select & "                   WHERE HN2.材料管理番号 = ZT.材料管理番号)"
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
            gSL_Select = gSL_Select & "   ON  HN.品番 = ZN.使用品番"
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY"
            gSL_Select = gSL_Select & "   ON  SY.区分名称 = '単位区分'"
            gSL_Select = gSL_Select & "   AND SY.値 = ZR.手配単位"
            gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2"
            gSL_Select = gSL_Select & "   ON  SY2.区分名称 = '材種区分'"
            gSL_Select = gSL_Select & "   AND SY2.値 = ZR.材種"
            gSL_Select = gSL_Select & " WHERE ZT.材料管理番号 = " & .材料管理番号(gInt)
            gSL_Select = gSL_Select & "   AND ZT.取引先 = '" & .仕入先 & "'"
            gSL_Select = gSL_Select & "   AND ZT.受入 = '" & .受入 & "'"
            If ZTMRead(gSL_Select, 1) Then
                MyRec2 = MyRec2 & RTrim(ZTM.材質)                           '材質
                MyRec2 = MyRec2 & Space(25 - CisFun.Len2(ZTM.材質))
                MyRec2 = MyRec2 & Replace(Format(ZTM.板厚, "0000.00"), ".", "") '板厚
                MyRec2 = MyRec2 & Replace(Format(ZTM.幅, "0000.00"), ".", "")   '幅
                MyRec2 = MyRec2 & Replace(Format(ZTM.長さ, "0000.00"), ".", "") '長さ
                MyRec2 = MyRec2 & RTrim(ZTM.材種コード)                               '材種
                MyRec2 = MyRec2 & Space(15 - CisFun.Len2(ZTM.材種コード))
                MyRec2 = MyRec2 & CisFun.Left2(ZTM.材種名, 12)                  '材種
                MyRec2 = MyRec2 & Space(12 - CisFun.Len2(CisFun.Left2(ZTM.材種名, 12)))
                MyRec2 = MyRec2 & CisFun.Left2(ZTM.品番, 25)                    '材種
                MyRec2 = MyRec2 & Space(25 - CisFun.Len2(CisFun.Left2(ZTM.品番, 25)))
                MyRec2 = MyRec2 & CisFun.Left2(ZTM.手配単位名, 4)                '単位名
                MyRec2 = MyRec2 & Space(4 - CisFun.Len2(CisFun.Left2(ZTM.手配単位名, 4)))
                MyRec2 = MyRec2 & CisFun.Left2(ZTM.備考, 1)                     '実数区分
            Else
                MyRec2 = MyRec2 & Space(84)
            End If
            MyRec2 = MyRec2 & Replace(Format(.収容数(gInt), "000000.000"), ".", "")     '収容数
            MyRec2 = MyRec2 & Format(.箱数(gInt), "00000")                              '箱数（枚数）
            MyRec2 = MyRec2 & Replace(Format(.端数(gInt), "000000.000"), ".", "")       '端数
            MyRec2 = MyRec2 & Replace(Format(.納入数(gInt), "00000000.000"), ".", "")   '納入数
        
        Next gInt
        
        MyRec = RTrim(MyRec1) & RTrim(MyRec2)
    End With
    
    Jisya_TextEdit = True
End Function


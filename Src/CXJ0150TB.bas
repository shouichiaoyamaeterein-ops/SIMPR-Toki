Attribute VB_Name = "CXJ0150TB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   自社伝票ファイル出力
'*
'*   モジュールID  : CXJ0150TB.BAS
'*
'*                  CXJ0150TFile(Integer)
'*
'*   引数 : Integer  =  0  :新規発行
'*   　　               1  :再発行
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*   　　 : False : 印刷データなし
'*
'*   使用例:
'*   If Not CXJ0150TFile(1) Then Exit Sub
'*
'*   作成日  :2004/03/26   By CIS
'*   変更日  :2004/04/06   By CIS - Irnaksic未登録時の対応
'*   変更日  :2004/04/16   By CIS - 帳票現頁のバグ対応
'*   変更日  :2006/02/02   By CIS - INIファイル-クライアント対応（TEXT出力先）
'*   変更日  :2008/04/10   By CIS - 背番号8桁
'*---------------------------------------------------------------------------------
    Dim My_JisyaCd      As String
    Dim My_JisyaKj      As String
    Dim My_JisyaNm      As String
    Dim My_JisyaKjNm    As String
    Dim My_Kiban        As String
    Dim My_OutDev       As String
    Dim My_FileName     As String
    Dim My_TextFile     As String

    Dim SyoriKb         As Byte
    Dim wNouNo          As String   '>納入番号
    Dim wNKanri         As Long     '>納品書管理NO
    Dim wNowPage        As Byte     '>帳票現頁
    Dim wAllPage        As Byte     '>帳票総頁
    Dim wNowPageT       As Byte     '>納番現頁
    Dim wAllPageT       As Byte     '>納番総頁
    Dim wLineNo         As Integer
    Dim wCard           As String   '>ｶｰﾄﾞ#
    Dim wHakoT          As Long     '>総箱数
'   << TEXT >>
    Dim MyFNo           As Integer
    Dim MyRec           As String
    Dim MyRec1          As String * 89
    Dim MyRec2          As String
    Dim MyRec3          As String
    Dim Key_Torcd       As String
    Dim MyCnt           As Long     '>出力件数
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
'*--------------------------------------------------<< 2004/04/06 Insert Start >>
        My_JisyaKjNm = ""
'*--------------------------------------------------<< 2004/04/06 Insert End >>
    End If

'--( TEXT出力先　獲得 )-- [クライアント情報]
    With CisFun
'-----------------------------------------------< 2006/02/02 ADD START >
        .INI_FileID = gIniExe & gwsIniName
'-----------------------------------------------< 2006/02/02 ADD START >
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
'-----------------------------------------------< 2006/02/02 ADD START >
        .INI_FileID = gIniExe & gIniName
'-----------------------------------------------< 2006/02/02 ADD START >
    End With
    'ディレクトリ存在確認
    With CisFun
        .FE_Path = RTrim(My_OutDev)
        .DirCheck Msg_Crt
    End With
'--<< ファイル名　獲得 >>--
    My_FileName = Format(Now(), "yyyymmddhhnnss")

    Jisya_TextInit = True
End Function
'<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
Public Function CXJ0150TFile(Optional ReCreate As Integer, Optional PrtMsg As Boolean = False)
    CXJ0150TFile = False

    App.Title = "自社伝票ファイル出力"
' 印刷進行画面表示ﾌﾗｸﾞｾｯﾄ
    PrtMsg = True
    pref_MsgDisp = PrtMsg
'帳票区分 （新規発行 OR 再発行 )
    Select Case ReCreate
      Case 0:  SyoriKb = 54
      Case 1:  SyoriKb = 154
    End Select
'--------------'
'   初期設定
'--------------'
    If Not Jisya_TextInit Then Exit Function
'-------------------------------'
'   出力データ存在確認
'-------------------------------'
    Call SqlSelect_Group
    If Not JNWRead(gSL_Select, 0, 1) Then
       Call JNWClose(1)
       CXJ0150TFile = True
       Exit Function
    End If
    If CisDB.RecordCount = 0 Then
       CXJ0150TFile = True
       Exit Function
    End If
    pref_TCnt = CisDB.RecordCount

'+++++++++++++++++++++++++++++++'
'+  印刷処理
'+++++++++++++++++++++++++++++++'
    Set pref_Form = CXJ0150
    pref_Msg = "【 自社伝票ファイル  出力中 】"

    If pref_MsgDisp Then
        pref_Form.Show vbModal
        Unload pref_Form
        Set pref_Form = Nothing
    Else
        Call Jisya_TextCreate
    End If

    CXJ0150TFile = True
End Function
'+----------------------------------------------------------+
'+                ファイル出力  処理                        ＋
'+----------------------------------------------------------+
Public Function Jisya_TextCreate() As Boolean
    Jisya_TextCreate = False

    On Error GoTo Jisya_TextCreate_Err

    Set CorForm = pref_Form
    pref_NCnt = 0                ' 現在発行ページ初期化

'   <><><><><> 納入番号の読込み <><><><><>
    Do Until Not JNW_RDSTS

        Key_Torcd = RTrim(JNW.納入先)
        '------------------------'
        '   出力ﾃｷｽﾄ OPEN
        '------------------------'
        My_TextFile = RTrim(My_OutDev) & "JD" & My_FileName & Key_Torcd & ".HKO"
        MyFNo = FreeFile
        MyCnt = 0
        Open My_TextFile For Output As #MyFNo
        '------------------------'
        '   ヘッダー 出力
        '------------------------'
        Select Case SyoriKb
            Case 54:    MyRec = "JHB0"
            Case Else:  MyRec = "JHB1"
        End Select
        MyRec = MyRec & RTrim(gCompany) & Space(5 - CisFun.Len2(RTrim(gCompany)))
        MyRec = MyRec & Key_Torcd & Space(7 - CisFun.Len2(Key_Torcd))
        MyRec = MyRec & My_FileName
        Print #MyFNo, MyRec

    '   <><><> 同一納入先の実行 <><><>
        While JNW_RDSTS And JNW.納入先 = Key_Torcd
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
            '------------------------'
            '   レコード編集
            '------------------------'
            '   >> HEAD部 項目転送
                wHakoT = 0
                Call TextEditHead
            '   <><><> １ページ分実行 <><><>
                While JYT_RDSTS And JYT.納入番号 = wNouNo _
                                And JYT.納品書頁 = wNowPage
                '   >> BODY部 項目転送
                    wHakoT = wHakoT + JYT.枚数
                    wLineNo = JYT.納品書行
                    Call TextEditBody

                '   >> 次のﾚｺｰﾄﾞを読み込む
                    Call JYTReadNext(0)
                Wend
            '   >> TAIL部 項目転送
                Call TextEditTail
            '------------------------'
            '   レコード出力
            '------------------------'
                Print #MyFNo, MyRec
                MyCnt = MyCnt + 1

                pref_NCnt = pref_NCnt + 1
                If pref_MsgDisp Then
                    pref_Cnt = Format(pref_NCnt, "###") & " / " & Format(pref_TCnt, "###")
                    DoEvents
                End If
            Loop
            Call JYTClose(0)

    '       >> 次のﾚｺｰﾄﾞを読み込む
            Call JNWReadNext(1)
        Wend

        '------------------------'
        '   ﾄﾚｲﾗｰ 出力
        '------------------------'
        MyRec = "JT" & Format(MyCnt, "0000")
        Print #MyFNo, MyRec
        '------------------------'
        '   出力ﾃｷｽﾄ CLOSE
        '------------------------'
        Close #MyFNo

    Loop
    Call JNWClose(1)

    Jisya_TextCreate = True
    Exit Function
Jisya_TextCreate_Err:
    CisFun.ErrorBox
    End
End Function
'+++++++++++++++++++++++++++++++'
'+      ページ数のセット
'+++++++++++++++++++++++++++++++'
Private Sub SetPageCount()
    Dim wDenMax         As Byte
    Dim wDenCnt         As Byte

    wDenMax = 0
    wDenCnt = 0
    If JNW.受領書形態 = "1" Then
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
            wDenMax = wDenMax + 1                               '>2004/04/17 Insert
        End If
'       wDenMax = wDenMax + 1                                   '>2004/04/17 Delete
    End If
    wAllPageT = wAllPage * wDenMax
'*--------------------------------------------------<< 2004/04/16 Update Start >>
'   wNowPageT = wAllPage * wDenCnt + 1
    wNowPageT = wAllPage * wDenCnt + wNowPage
'*--------------------------------------------------<< 2004/04/16 Update End >>

End Sub
'+==========================================================+
'+              ファイル編集  処理
'+==========================================================+
'   *---------------------------------*
'   *   帳票発行テキスト編集 - HEAD部
'   *---------------------------------*
Private Function TextEditHead() As Boolean
    TextEditHead = False

    MyRec1 = Space(89)
    MyRec2 = ""
    MyRec3 = ""

    Mid(MyRec1, 1, 1) = "J"                              'レコード種別
    Mid(MyRec1, 2, 1) = "B"                              'データ種別
    Mid(MyRec1, 3, 3) = Format(gCnt1, "000")             'レコード連番
    With JYT
        Mid(MyRec1, 6, 1) = wCard                                   'カード#
        Mid(MyRec1, 7, 5) = RTrim(.納入番号) & Space(4)             '納入番号
        Mid(MyRec1, 12, 2) = Format(.納品書頁, "00")                'ページNO
        Mid(MyRec1, 14, 10) = Format(.納品書管理NO, "0000000000")   '納品書管理NO
        Mid(MyRec1, 24, 5) = My_JisyaCd                             '受注者
        Mid(MyRec1, 29, 2) = My_JisyaKj                             '受注者事業所
        Select Case SyoriKb
            Case 54:    Mid(MyRec1, 31, 1) = "0"            '発行区分(新規発行)
            Case Else:  Mid(MyRec1, 31, 1) = "1"            '発行区分(再発行)
        End Select
        Select Case CisFun.Mid2(.納入番号, 5, 1)
            Case "0":   Mid(MyRec1, 32, 1) = "0"            '分割区分(新規)
            Case Else:  Mid(MyRec1, 32, 1) = "1"            '分割区分(分納）
        End Select
        Mid(MyRec1, 33, 1) = RTrim(.受注区分)               '発注区分
        Mid(MyRec1, 34, 1) = "0"                            '手配区分
        Mid(MyRec1, 35, 1) = Format(.SK区分, "0")           'SK区分
        Mid(MyRec1, 36, 5) = RTrim(.契約取引先)             '支給元(契約先)
        Mid(MyRec1, 41, 2) = RTrim(.契約工場)               '支給元(契約先工場)
        Mid(MyRec1, 43, 2) = RTrim(.受入)                   '納入者事業所(受入）
        Mid(MyRec1, 45, 5) = RTrim(.取引先CD)               '納入者(納入先)
        Mid(MyRec1, 50, 2) = RTrim(.工場CD)                 '納入者事業所(納入先工場)
        Mid(MyRec1, 52, 2) = RTrim(.受入)                   '納入者事業所(納入先受入)
        Mid(MyRec1, 54, 8) = RTrim(.納入日)                 '納入日
        Mid(MyRec1, 62, 2) = RTrim(.納入便)                 '納入便
        Mid(MyRec1, 64, 8) = Format(Date, "yyyymmdd")       '作成日
        Mid(MyRec1, 72, 3) = RTrim(My_Kiban)                '機番
'       <><><><><> 名称 <><><><><>
        MyRec2 = MyRec2 & CisFun.Left2(.受注区分名, 6)                  '発注区分名
        MyRec2 = MyRec2 & Space(6 - CisFun.Len2(CisFun.Left2(.受注区分名, 6)))
        MyRec2 = MyRec2 & Space(6)                          '手配区分名
        Select Case .SK区分
            Case 0:    MyRec2 = MyRec2 & "かんばん"         'SK区分名(かんばん）
            Case Else: MyRec2 = MyRec2 & "指示部品"         'SK区分名(指示）
        End Select
        MyRec2 = MyRec2 & CisFun.Left2(My_JisyaNm, 24)                  '仕入先名（自社）
        MyRec2 = MyRec2 & Space(24 - CisFun.Len2(CisFun.Left2(My_JisyaNm, 24)))
        MyRec2 = MyRec2 & CisFun.Left2(My_JisyaKjNm, 16)                '仕入先工場名（自社）
        MyRec2 = MyRec2 & Space(16 - CisFun.Len2(CisFun.Left2(My_JisyaKjNm, 16)))
        MyRec2 = MyRec2 & CisFun.Left2(.納入先名, 24)                   '納入先名
        MyRec2 = MyRec2 & Space(24 - CisFun.Len2(CisFun.Left2(.納入先名, 24)))
        MyRec2 = MyRec2 & CisFun.Left2(.工場名, 16)                     '納入先工場名
        MyRec2 = MyRec2 & Space(16 - CisFun.Len2(CisFun.Left2(.工場名, 16)))
    End With

    TextEditHead = True
End Function
'   *---------------------------------*
'   *   帳票発行テキスト編集 - BODY部
'   *---------------------------------*
Private Function TextEditBody() As Boolean
    TextEditBody = False

    With JYT
        MyRec3 = MyRec3 & .背番号                           '背番号
        '*-*-*-*-* 2008/04/10 Update Start
        'MyRec3 = MyRec3 & Space(5 - CisFun.Len2(.背番号))
        MyRec3 = MyRec3 & Space(8 - CisFun.Len2(.背番号))
        '*-*-*-*-* 2008/04/10 Update End
        MyRec3 = MyRec3 & .品番                             '部品番号
        MyRec3 = MyRec3 & Space(25 - CisFun.Len2(.品番))
        MyRec3 = MyRec3 & .品番編集                         '品番編集
        MyRec3 = MyRec3 & Space(5 - CisFun.Len2(.品番編集))
        MyRec3 = MyRec3 & .品名                             '部品名称
        MyRec3 = MyRec3 & Space(40 - CisFun.Len2(.品名))
        MyRec3 = MyRec3 & Format(.収容数, "000000")         '収容数
        MyRec3 = MyRec3 & Format(.枚数, "00000")            '箱数（枚数）
        MyRec3 = MyRec3 & Format(.端数, "000000")           '端数
        MyRec3 = MyRec3 & Format(.納入数, "0000000000")     '納入数
    End With

    TextEditBody = True
End Function
'   *---------------------------------*
'   *   帳票発行テキスト編集 - TAIL部
'   *---------------------------------*
Private Function TextEditTail() As Boolean
    TextEditTail = False

    Mid(MyRec1, 75, 4) = Format(wHakoT, "0000")         '箱総数
    Mid(MyRec1, 79, 2) = Format(wAllPage, "00")         '総頁数
    Mid(MyRec1, 81, 3) = CisFun.Mid2(wNouNo, 2, 3)      'グループNO
    Mid(MyRec1, 84, 2) = Format(wNowPageT, "00")        '物理現頁数
    Mid(MyRec1, 86, 2) = Format(wAllPageT, "00")        '物理総頁数
'   >> END表示区分
    If wNowPage = wAllPage Then Mid(MyRec1, 88, 1) = "1" Else Mid(MyRec1, 88, 1) = "0"
'   >> ALLEND表示区分
    If wNowPageT = wAllPageT Then Mid(MyRec1, 89, 1) = "1" Else Mid(MyRec1, 89, 1) = "0"
'   >> 明細行数
    MyRec2 = MyRec2 & Format(wLineNo, "00")

'   <><><><><> 結合 <><><><><>
    MyRec = MyRec1 & MyRec2 & MyRec3

    TextEditTail = True
End Function
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
    gSL_Select = gSL_Select & ", isnull(hnm.品番編集, '') 品番編集"
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
    gSL_Select = gSL_Select & " where jyt.納入番号 = '" & wNouNo & "'"
    gSL_Select = gSL_Select & "   and jyt.納品書管理NO = '" & wNKanri & "'"
    gSL_Select = gSL_Select & "   and jyt.納品書頁 = '" & wNowPage & "'"
    gSL_Select = gSL_Select & "   and jyt.納品書発行区分 = 1"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  jyt.納品書頁"
    gSL_Select = gSL_Select & ", jyt.納品書行"

    SqlSelect_Print = True
End Function


Attribute VB_Name = "CXM0000LB"
Option Explicit

'*--------------------------------------------------------------------------------
'*
'*   環境設定 発行
'*
'*   モジュールID  : CXM0000LB.bas
'*
'*                   CXM0000LPrint(Integer)
'*
'*   初期設定値 : CXM0000_PrtKB :  0/ﾌﾟﾘﾝﾀｰ,1/ﾌﾟﾚﾋﾞｭｰ,2/Mail送信,3/ﾌﾟﾚﾌﾟﾘﾝﾄ
'*
'*   引数 : なし
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:
'    If Not CXM0000LPrint(0) Then Exit Sub
'
'*   作成日  : 2006/04/03   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Global CXM0000_PrtKB            As Byte                 '  0/ﾌﾟﾘﾝﾀｰ,1/ﾌﾟﾚﾋﾞｭｰ,2/Mail送信
Global CXM0000_CoReport         As New CoReports85      ' クラス定義
Global CXM0000_PrintName        As String
Public CXM0000_MailSes          As Object   ' OutLook(MSMAPI32.ocx)ｵﾌﾞｼﾞｪｸﾄ(ｾｯｼｮﾝ
Public CXM0000_MailMsg          As Object   ' OutLook(MSMAPI32.ocx)ｵﾌﾞｼﾞｪｸﾄ

Dim Rg_On           As String
Dim Rg_Off          As String

Dim Ma_ConString    As String
Dim Ma_Use          As String
Dim Ma_Port         As String
Dim Ma_Speed        As String
Dim Ma_Parity       As String
Dim Ma_DataLen      As String
Dim Ma_StopLen      As String
Dim Ma_HandShak     As String

Dim Ma_EditStr      As String

'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/
'_/             印　　刷　　処　　理
'_/
'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Function CXM0000LPrint()
    CXM0000LPrint = False
    App.Title = "環境設定"
    
' 印刷進行画面表示ﾌﾗｸﾞｾｯﾄ
    pref_MsgDisp = True
'+++++++++++++++++++++++++++++++'
'+  印刷処理
'+++++++++++++++++++++++++++++++'
    Set pref_Form = CXM0000L
    pref_Msg = "【 環 境 設 定 印 刷 中 】"
    pref_Form.Show vbModal
    Unload pref_Form
    Set pref_Form = Nothing

    CXM0000LPrint = True
    
CXM0000LPrint_ED:
    
End Function

'+----------------------------------------------------------+
'+                    印  刷  処  理                        ＋
'+----------------------------------------------------------+
Public Function CoPrintM0000() As Boolean
    Dim PageCnt     As Long
    
    CoPrintM0000 = False
    On Error GoTo CoPrintM0000_Er
    
    pref_NCnt = 0                ' 現在発行ページ初期化
    
    With CXM0000_CoReport
        .SetDraw = pref_Form.CrDraw1

        .SetPrtForm = "LXM0000"        ' ﾌｫｰﾑIDをｾｯﾄ
        .SetPrtPath = gIniExe & "Prt\" ' ﾌｫｰﾑ格納先
        
        Select Case CXM0000_PrtKB
            Case 0: .SetPrtOut = PrinterX            ' ﾌﾟﾘﾝﾀｰ
            Case 1: .SetPrtOut = Preview            ' ﾌﾟﾚﾋﾞｭｰ
                    .SetPreviewForm = frmPreview    '   ﾌﾟﾚﾋﾞｭｰ用ﾌｫｰﾑ(frm)
            Case 2: .SetPrtOut = PDF                ' PDF出力
                    .SetPDFName = "SIMPRA環境設定"
        End Select
        .SetPrtName = "環境設定"                      ' 帳票名
        .SetPaperSize = A4
        .SetPaperOrientation = 縦
        If Trim(CXM0000_PrintName) = "" Then
           .SetPrinter = CXM0000_PrintName
        End If
        PageCnt = 0
        
        '------ 印刷処理開始
        If Not .CorOutStart Then GoTo CoPrintM0000_Er
        '------ 項目転送
        Call ItemMoveFore
        '------ 印刷
        If Not .CoPrintOut Then GoTo CoPrintM0000_Er
        '------ 印刷終了
        If Not .CoOutEnd Then GoTo CoPrintM0000_Er
    End With
    
    If CXM0000_PrtKB = 2 Then
    '   < メッセージ作成用ワークの初期化 >
        Call gMapi_ItemClear

        gMapi_MailTO = "simpr-a@cisweb.co.jp"      ' 送信先(宛先)
        gMapi_MailCC = ""              ' CC(カーボンコピー)
        gMapi_MailSBJ = ""             ' 件名
        gMapi_MailMSG = ""             ' 本文
    
        gMapi_AttcBIN(0) = gIniExe & "Prt\SIMPRA環境設定.pdf"
        
        DoEvents
    '    ' 添付ファイル
        If RTrim$(Dir(gMapi_AttcBIN(0))) = "" Then
    '       gStr = "  【エラー】  添付ファイルなし ファイル名= " & cSend_MailDir & cSend_MailFile
           Exit Function
        End If
    
        If Not SndMail(CXM0000_MailSes, CXM0000_MailMsg, False) Then Exit Function
    End If
    
    CoPrintM0000 = True
    Exit Function
CoPrintM0000_Er:
    Call CisFun.ErrorBox("印刷エラー")
End Function
'*------------------------------------------------------------------*
'*                                                                  *
'*      発行内容　項目転送                                          *
'*                                                                  *
'*------------------------------------------------------------------*
Private Function ItemMoveFore() As Boolean

    ItemMoveFore = False

    With CXM0000_CoReport
        If CXM0000_PrtKB = 3 Then
            '------ まるを全て非表示
            For Each gObj In .CorForm.CrObjects
                If gObj.ObjectType = 3 Then
                   gObj.Visible = False
                End If
            Next gObj
                '------ フィールドを全てクリアします
            Call .CorForm.ClearAllFields
            ItemMoveFore = True
            Exit Function
        End If
    
    
    '--------------------------------
    '--     基本設定
    '--------------------------------
    '---------------- システム環境
        If Not .CorData("発行日", Format(Date, "YYYY.MM.DD"), "S", 0) Then Exit Function
        If Not .CorData("発行端末", CisFun.WSID, "", 0) Then Exit Function
        If Not .CorData("VERSION", "Ver " & RTrim(gSysVer), "", 0) Then Exit Function
        If Not .CorData("UPDATE", "最終更新日 " & RTrim(gSysUpDate), "", 0) Then Exit Function
        If Not .CorData("システム名", gSysName, "", 0) Then Exit Function
        If Not .CorData("USERCD", gCompany, "", 0) Then Exit Function
        If Not .CorData("USERNAME", gCompName, "", 0) Then Exit Function
        If gLogoDisp Then
           gStr = "表示"
        Else
           gStr = "非表示"
        End If
        If Not .CorData("LOGODISP", gStr, "", 0) Then Exit Function
        If Not .CorData("LOGO", gLogo, "", 0) Then Exit Function
    '---------------- クライアント環境
        If Not .CorData("EXEDIR", gIniExe, "", 0) Then Exit Function
        If Not .CorData("DBServer", CisDB.Server, "", 0) Then Exit Function
        If Not .CorData("DBName", CisDB.DBName, "", 0) Then Exit Function
        If Not .CorData("DBUser", CisDB.User, "", 0) Then Exit Function
        
        With CisDB
            gStr = ""
            .SQL = "SELECT filename FROM master.dbo.sysdatabases Where name = '" & CisDB.DBName & "'"
            If .DBRead(1) Then
               If Not IsNull(.Recordset.Fields.Item(0).Value) Then gStr = .Recordset.Fields.Item(0).Value
            End If
        End With
        If Not .CorData("DBPATH", gStr, "", 0) Then Exit Function
        
        With CisDB
            gStr = ""
            .SQL = "SELECT filename FROM sysfiles Where fileid = 2"
            If .DBRead(1) Then
               If Not IsNull(.Recordset.Fields.Item(0).Value) Then gStr = .Recordset.Fields.Item(0).Value
            End If
        End With
        If Not .CorData("LOGPATH", gStr, "", 0) Then Exit Function
    '--------------------------------
    '--     オプション設定
    '--------------------------------
    '---------------- システム共通設定
        '--------------------------------
        '---------------- 受注関連
        '--------------------------------
        ' 売上処理
        If Not CircleDisp(50, "受注_検収", "受注_売上") Then Exit Function
        ' 受注かんばん読取りエラーリスト発行
        If Not CircleDisp(51, "受注_ﾘｽﾄ有", "受注_ﾘｽﾄなし") Then Exit Function
        '--------------------------------
        '---------------- 発注関連
        '--------------------------------
        ' 発行元
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 4
        If Not IRNRead("", 1) Then Call ItemsClearIRN
        gStr = Mid(IRN.oyian, 1, 7)
        If Not .CorData("発行元", gStr, "", 0) Then Exit Function
        ' かんばん流動履歴日数
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 32
        If Not IRNRead("", 1) Then IRN.oyian = "020"
        gStr = CisFun.Mid2(IRN.oyian, 1, 3)
        If Not .CorData("流動履歴", gStr, "", 0) Then Exit Function
        ' かんばん滞留期間日数
        gSL_Select = " SELECT * FROM 名称マスタ "
        gSL_Select = gSL_Select & " WHERE 区分名称 = 'データ管理'"
        gSL_Select = gSL_Select & "   AND 区分タイプ = '' "
        gSL_Select = gSL_Select & "   AND 値 = 'KAN' "
        If Not SYMRead(gSL_Select, 1) Then Call ItemsClearSYM
        gStr = SYM.数字1
        If Not .CorData("滞留日数", gStr, "", 0) Then Exit Function
        ' 未流動かんばん
        If Not CircleDisp(37, "発注_未流動_有効", "発注_未流動_無効") Then Exit Function
        ' かんばん読取り 即時発注
        If Not CircleDisp(5, "発注_読取_指定", "発注_読取_即時") Then Exit Function
        ' 稼働日カレンダー
        If Not CircleDisp(33, "発注_稼働日_自社", "発注_稼働日_仕入先") Then Exit Function
        ' 納入日・便算出方法
        If Not CircleDisp(14, "発注_納入日_時間", "発注_納入日_テーブル") Then Exit Function
        ' 完納予定日・便算出方法
        If Not CircleDisp(16, "発注_完納日_手動", "発注_完納日_自動") Then Exit Function
        ' 平準化月超過発注
        If Not .CorEditItemSet("発注_平準化_回収", False) Then Exit Function
        If Not .CorEditItemSet("発注_平準化_発注", False) Then Exit Function
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 6
        If Not IRNRead("", 1) Then Call ItemsClearIRN
        If CisFun.Val2(Mid(IRN.oyian, 1, 1)) = 0 Then
           If Not .CorData("発注_平準化_出力", "ＣＳＶ", "", 0) Then Exit Function
        Else
           If Not .CorData("発注_平準化_出力", "帳票", "", 0) Then Exit Function
        End If
        If CisFun.Val2(Mid(IRN.oyian, 2, 1)) = 0 Then
           If Not .CorEditItemSet("発注_平準化_回収", True) Then Exit Function
        Else
           If Not .CorEditItemSet("発注_平準化_発注", True) Then Exit Function
        End If
        ' 自社伝票発行用(工場・機番)
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 7
        If Not IRNRead("", 1) Then Call ItemsClearIRN: IRN.oyian = "0": IRN.oyian2 = "00"
        ' 工場
        If Not .CorData("発注_自社_工場", RTrim(IRN.oyian), "", 0) Then Exit Function
        ' 機番
        If Not .CorData("発注_自社_機番", RTrim(IRN.oyian2), "", 0) Then Exit Function
        ' 受領印
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 20
        If Not IRNRead("", 1) Then Call ItemsClearIRN: IRN.oyian = "00受領"
        If CisFun.Val2(Mid(IRN.oyian, 1, 1)) = 0 Then
           If Not .CorData("発注_受領書_印字", "なし", "", 0) Then Exit Function
           If Not .CorData("発注_受領書_日付", "", "", 0, False) Then Exit Function
           If Not .CorData("発注_受領書_上段", "", "", 0, False) Then Exit Function
           If Not .CorData("発注_受領書_下段", "", "", 0, False) Then Exit Function
        Else
           If Not .CorData("発注_受領書_印字", "あり", "", 0) Then Exit Function
           If CisFun.Val2(Mid(IRN.oyian, 2, 1)) = 0 Then
              If Not .CorData("発注_受領書_日付", "受領日あり", "", 0) Then Exit Function
           Else
              If Not .CorData("発注_受領書_日付", "受領日なし", "", 0) Then Exit Function
           End If
           If Not .CorData("発注_受領書_上段", CisFun.Mid2(IRN.oyian, 3, 8), "", 0) Then Exit Function
           If Not .CorData("発注_受領書_下段", CisFun.Mid2(IRN.oyian2, 1, 8), "", 0) Then Exit Function
        End If
        ' JAMA伝票発行用(工場・機番)
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 8
        If Not IRNRead("", 1) Then Call ItemsClearIRN: IRN.oyian = "0": IRN.oyian2 = "00"
        ' 工場
        If Not .CorData("発注_JAMA_工場", RTrim(IRN.oyian), "", 0) Then Exit Function
        ' 機番
        If Not .CorData("発注_JAMA_機番", RTrim(IRN.oyian2), "", 0) Then Exit Function
        ' チケット伝票発行用(機番)
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 15
        If Not IRNRead("", 1) Then Call ItemsClearIRN: IRN.oyian2 = "00"
        If Not .CorData("発注_チケット_機番", RTrim(IRN.oyian2), "", 0) Then Exit Function
        ' 臨時かんばん発行
        If Not CircleDisp(1, "発注_臨時_有", "発注_臨時_なし") Then Exit Function
        ' 現品票発行
        If Not CircleDisp(2, "発注_現品票_有", "発注_現品票_なし") Then Exit Function
        ' 端数ラベル発行
        If Not CircleDisp(3, "発注_端数_有", "発注_端数_なし") Then Exit Function
        ' 発注明細控
        If Not CircleDisp(12, "発注_明細_控_有", "発注_明細_控_なし") Then Exit Function
        ' 発注明細空
        If Not CircleDisp(13, "発注_明細_空_有", "発注_明細_空_なし") Then Exit Function
        ' 発注かんばん読取りエラーリスト発行
        If Not CircleDisp(35, "発注_かんばんエラー_有", "発注_かんばんエラー_なし") Then Exit Function
        ' 発注打切りリスト発行
        If Not CircleDisp(36, "発注_打切_有", "発注_打切_なし") Then Exit Function
        ' 材料発注管理オプション
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 10
        If Not IRNRead("", 1) Then Call ItemsClearIRN: IRN.oyian = "0"
        If IRN.oyian = "1" Then
            Rg_On = "登録済":   Rg_Off = "登録解除"
        Else
            Rg_On = "登録":   Rg_Off = "未登録"
        End If
        If CisFun.Val2(IRN.oyian) = 1 Then
           If Not .CorData("発注_材料発注管理", Rg_On, "", 0) Then Exit Function
        Else
           If Not .CorData("発注_材料発注管理", Rg_Off, "", 0) Then Exit Function
        End If
        ' 材料発注　内示対象データ
        If Not .CorEditItemSet("発注_材料発注_当月", False) Then Exit Function
        If Not .CorEditItemSet("発注_材料発注_翌月", False) Then Exit Function
        If Not .CorEditItemSet("発注_材料発注_翌々月", False) Then Exit Function
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 11
        If Not IRNRead("", 1) Then Call ItemsClearIRN: IRN.oyian = "0"
        If CisFun.Val2(Left(IRN.oyian, 1)) = 0 Then
           If Not .CorEditItemSet("発注_材料発注_当月", True) Then Exit Function
        End If
        If CisFun.Val2(Left(IRN.oyian, 1)) = 2 Then
           If Not .CorEditItemSet("発注_材料発注_翌月", True) Then Exit Function
        End If
        If CisFun.Val2(Left(IRN.oyian, 1)) = 3 Then
           If Not .CorEditItemSet("発注_材料発注_翌々月", True) Then Exit Function
        End If
    '---------------- クライアント設定
        '--------------------------------
        '---------------- フォルダ指定
        '--------------------------------
        ' TEXT変換ﾛｸﾞ保持日数
        gStr = "60"
        Call GetClientOption("G", "TextConvert", "LogDelDay", gStr)
        If Not .CorData("TEXTログ", gStr, "", 0) Then Exit Function
        ' 自社ｺｰﾄﾞﾁｪｯｸ
        gStr = "No"
        Call GetClientOption("G", "TextConvert", "Jisya", gStr)
        If UCase(gStr) = "Y" Or UCase(gStr) = "YES" Then
           If Not .CorData("自社コードチェック", "あり", "", 0) Then Exit Function
        Else
           If Not .CorData("自社コードチェック", "なし", "", 0) Then Exit Function
        End If
        ' 自社伝票ファイル出力先
        gStr = gIniExe
        Call GetClientOption("G", "Option", "DENPYOFILEOUT", gStr)
        If Not .CorListSet("DIR", gStr, 0, 75) Then Exit Function
        ' JAMA伝票ファイル出力先
        gStr = gIniExe
        Call GetClientOption("G", "Option", "JAMAFILEOUT", gStr)
        If Not .CorListSet("DIR", gStr, 1, 75) Then Exit Function
        ' 平準化月超過ファイル出力先
        gStr = gIniExe
        Call GetClientOption("G", "Option", "HEIFILEOUT", gStr)
        If Not .CorListSet("DIR", gStr, 2, 75) Then Exit Function
        ' 変換元フォルダ
        gStr = gIniExe & "Text\"
        Call GetClientOption("G", "TextConvert", "InputDir", gStr)
        If Not .CorListSet("DIR", gStr, 3, 75) Then Exit Function
        ' イメージ格納先
        gStr = gIniExe
        Call GetClientOption("G", "Option", "Image", gStr)
        If Not .CorListSet("DIR", gStr, 4, 75) Then Exit Function
        ' 自給チケットファイル出力先
        gStr = gIniExe
        Call GetClientOption("G", "Option", "JTICKETFILEOUT", gStr)
        If Not .CorListSet("DIR", gStr, 5, 75) Then Exit Function
        ' 支給チケットファイル出力先
        gStr = gIniExe
        Call GetClientOption("G", "Option", "STICKETFILEOUT", gStr)
        If Not .CorListSet("DIR", gStr, 6, 75) Then Exit Function
    
        '--------------------------------
        '---------------- 接続機器関連
        '--------------------------------
    
        Call Get_Connection("5")
        Call CircleDisp(0, "接続_キーボード_使用", "接続_キーボード_未使用", Ma_Use)
        ' かんばんソーター
        Call Get_Connection("1")
        Call CircleDisp(0, "接続_使用1", "接続_未使用1", Ma_Use)
        If Not .CorListSet("接続情報", Ma_EditStr, 0, 0) Then Exit Function
        ' かんばん手差しﾘｰﾀﾞｰ(QR)
        Call Get_Connection("2")
        Call CircleDisp(0, "接続_使用2", "接続_未使用2", Ma_Use)
        If Not .CorListSet("接続情報", Ma_EditStr, 1, 0) Then Exit Function
        ' QRﾊﾝﾄﾞｽｷｬﾅ
        Call Get_Connection("4")
        Call CircleDisp(0, "接続_使用3", "接続_未使用3", Ma_Use)
        If Not .CorListSet("接続情報", Ma_EditStr, 2, 0) Then Exit Function
        ' かんばん手差しﾘｰﾀﾞｰ(縦ﾊﾞｰ)
        Call Get_Connection("6")
        Call CircleDisp(0, "接続_使用4", "接続_未使用4", Ma_Use)
        If Not .CorListSet("接続情報", Ma_EditStr, 3, 0) Then Exit Function
        ' QRﾎｯﾊﾟﾘｰﾀﾞｰ
        Call Get_Connection("7")
        Call CircleDisp(0, "接続_使用5", "接続_未使用5", Ma_Use)
        If Not .CorListSet("接続情報", Ma_EditStr, 4, 0) Then Exit Function
        
        'かんばん印刷先
        Call GetClientOption("G", "かんばん印刷先", CisFun.WSID, gStr)
        If Not .CorData("かんばん印刷先", gStr, "", 0) Then Exit Function
        'チケット印刷先
        Call GetClientOption("G", "自給チケット印刷先", CisFun.WSID, gStr)
        If Not .CorData("チケット印刷先", gStr, "", 0) Then Exit Function
    '---------------- バックアップ情報
        '--------------------------------
        '---------------- バックアップ情報取得
        '--------------------------------
        Call sql_job_info_Get
        
        For gInt = 0 To sql_job_max - 1
            If Trim(sql_Job(gInt).JobName) <> "" Then
               If Not .CorListSet("JOB_NAME", sql_Job(gInt).JobName, gInt, 0) Then Exit Function
               If Not .CorListSet("JOB_実行", sql_Job(gInt).実行, gInt, 0) Then Exit Function
               If Not .CorListSet("JOB_間隔", sql_Job(gInt).間隔, gInt, 0) Then Exit Function
               If Trim(sql_Job(gInt).active_start_time) <> "" Then
                  gStr = Mid(sql_Job(gInt).active_start_time, 1, 2) & ":" & _
                         Mid(sql_Job(gInt).active_start_time, 3, 2)
                  If Not .CorListSet("JOB_開始時刻", gStr, gInt, 0) Then Exit Function
               End If
               If Trim(sql_Job(gInt).next_run_date) <> "" Then
                  gStr = Mid(sql_Job(gInt).next_run_date, 5, 2) & "/" & _
                         Mid(sql_Job(gInt).next_run_date, 7, 2)
                  If Not .CorListSet("JOB_次回", gStr, gInt, 0) Then Exit Function
               End If
               If CisFun.Val2(sql_Job(gInt).EndProcYMD) <> 0 Then
                  gStr = Mid(sql_Job(gInt).EndProcYMD, 5, 2) & "/" & _
                         Mid(sql_Job(gInt).EndProcYMD, 7, 2)
                  If Not .CorListSet("JOB_最終", gStr, gInt, 0) Then Exit Function
               End If
               If Not .CorListSet("JOB_DIR", sql_Job(gInt).BackPath, gInt, 0) Then Exit Function
            End If
        Next gInt
    
    End With
End Function
'+-----------------------------------+
'+      まる表示非表示
'+-----------------------------------+
Private Function CircleDisp(kb As Long, BaseDisp As String, SelectDisp As String, Optional BaseFlg As String = "Y") As Boolean
    CircleDisp = False
    With CXM0000_CoReport
        If Not .CorEditItemSet(BaseDisp, False) Then Exit Function
        If Not .CorEditItemSet(SelectDisp, False) Then Exit Function
        If kb <> 0 Then
            Call ItemsClearIRN
            IRN.bk = "Option": IRN.rno = kb
            If Not IRNRead("", 1) Then Call ItemsClearIRN
            If CisFun.Val2(Left(IRN.oyian, 1)) = 0 Then
               If Not .CorEditItemSet(BaseDisp, True) Then Exit Function
            Else
               If Not .CorEditItemSet(SelectDisp, True) Then Exit Function
            End If
        Else
            If BaseFlg = "Y" Then
               If Not .CorEditItemSet(BaseDisp, True) Then Exit Function
            Else
               If Not .CorEditItemSet(SelectDisp, True) Then Exit Function
            End If
        End If
    End With
    CircleDisp = True
End Function
'+-----------------------------------+
'+      接続機器情報取得
'+-----------------------------------+
Private Sub Get_Connection(ConnectionNo As String)
    Dim wStr()      As String
    
    Ma_EditStr = ""
    Ma_ConString = ""
    Ma_Use = ""
    Ma_Port = ""
    Ma_Speed = ""
    Ma_Parity = ""
    Ma_DataLen = ""
    Ma_StopLen = ""
    Ma_HandShak = ""
    
    gStr = ""
    Call GetCryOption("G", "Apparatus", ConnectionNo, gStr)
    If Trim(gStr) = "" Then Exit Sub
    Ma_ConString = gStr
    ' 使用･未使用
    Call GetClientOption("G", Ma_ConString, "Use", gStr)
    If UCase(Trim(gStr)) = "N" Or UCase(Trim(gStr)) = "NO" Then
       Ma_Use = "N"
    Else
       Ma_Use = "Y"
    End If
    ' ポート
    Call GetClientOption("G", Ma_ConString, "CommPort", gStr)
    Ma_Port = RTrim(gStr)
    ' 速度/ﾊﾟﾘﾃｨ/ｽﾀｰﾄ･ｽﾄｯﾌﾟ/ﾃﾞｰﾀ長
    Call GetCryOption("G", Ma_ConString, "Setting", gStr)
    If RTrim(gStr) <> "" Then
       wStr = Split(gStr, ",")
       Ma_Speed = wStr(0)
       Ma_Parity = wStr(1)
       Ma_DataLen = wStr(2)
       Ma_StopLen = wStr(3)
    End If
    ' ﾊﾝﾄﾞｼｪｲｸ
    Call GetCryOption("G", Ma_ConString, "HandShaking", gStr)
    Ma_HandShak = RTrim(gStr)


   ' 接続情報編集
    'Port:1 Speed:19200 Pty:偶数 Data:8 Start:1 Hand:3:XON/XOFF･RTS/CTS
    Ma_EditStr = ""
    If Ma_Port = "" Then Exit Sub
    
    Ma_EditStr = "Port=" & Ma_Port
    gStr = ""
    Select Case UCase(Ma_Parity)
        Case "N":   gStr = "なし"
        Case "E":   gStr = "偶数"
        Case "O":   gStr = "奇数"
        Case "S":   gStr = "ｽﾍﾟｰｽ"
        Case "M":   gStr = "ﾏｰｸ"
    End Select
    Ma_EditStr = Ma_EditStr & " ,ﾊﾟﾘﾃｨ=" & gStr
    Ma_EditStr = Ma_EditStr & " ,ﾃﾞｰﾀ長=" & Ma_DataLen
    Ma_EditStr = Ma_EditStr & " ,ｽﾄｯﾌﾟ長=" & Ma_StopLen
    gStr = ""
    Select Case Ma_HandShak
        Case 0:   gStr = "0:なし"
        Case 1:   gStr = "1:XON/XOFF"
        Case 2:   gStr = "2:RTS/CTS"
        Case 3:   gStr = "3:XON/XOFF･RTS/CTS"
    End Select
    Ma_EditStr = Ma_EditStr & " ,ﾊﾝﾄﾞｼｪｲｸ=" & gStr
    
    
End Sub

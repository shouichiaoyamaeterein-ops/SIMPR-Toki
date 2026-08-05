Attribute VB_Name = "TNSSENDC04"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   TNS送信データ作成
'*          直送支給品(C04)
'*
'*   モジュールID  : TNSSENDC40.Bas
'*
'*
'*   作成日  : 2005/09/02   By CIS
'*   変更日  : 2007/02/08   By ｲﾝﾀﾞｽﾄﾘｰより複写し変更
'*   変更日  : 2007/03/07   By CIS 品番工場符号削除（ナルコ対応）
'*   変更日  : 2007/03/08   By CIS 品番チェック機能追加
'*                                 R07(TNS)対応(トヨタ紡織)
'*   変更日  : 2007/03/13   By CIS 品番工場符号削除（ナルコ対応）処理　削除
'*   変更日  : 2007/05/07   By CIS 契約先４桁以上対応(データ読込時のみ)
'*   変更日  :
'*
'*---------------------------------------------------------------------------------
'---------------------------------------- 2007/03/09 Insert
Public C04_HinErrCnt    As Long
Public C04_HeadKey      As String
Dim C04_OutFlg          As Boolean
Dim C04_SaveHinbn       As String
'---------------------------------------- 2007/03/09 Insert End
Dim Brk_KeyBefore       As String
Dim Brk_KeyAfter        As String
Dim C04_DataCnt         As Long
Dim PCnt                As Long
Dim Save_STorcd         As String


Public Function TNSSEND_C04() As Boolean
    
    On Error GoTo TNSSEND_C04_ER
    
    TNSSEND_C04 = False
    
    Ts_SendCnt = 0
    
'---------------------------------------------------------- 2007/03/13 Delete
'    '------------------------------------------------------------------ 2007/03/06 Insert
'    ReDim Ts_KojyoFix(1)
'    gSL_Select = "Select * From 名称マスタ"
'    gSL_Select = gSL_Select & " Where 区分名称 = '自社工区'"
'    gSL_Select = gSL_Select & " And   区分タイプ = ''"
'    gSL_Select = gSL_Select & " Order By 値"
'
'    If SYMRead(gSL_Select) Then
'       ReDim Ts_KojyoFix(CisDB.RecordCount)
'       gLong = 0
'       Do Until Not SYM_RDSTS
'
'            Ts_KojyoFix(gLong) = RTrim(SYM.英数字1)
'            gLong = gLong + 1
'
'            Call SYMReadNext
'       Loop
'    End If
'    Call SYMClose
'
'    '------------------------------------------------------------------ 2007/03/06 Insert End
'---------------------------------------------------------- 2007/03/13 Delete End
'
    gStr = "【" & TSS.ファイル名称 & "】 データ抽出 "
    Call TNSSend_MsgDisp(gStr, gStr)
    
    Call CisDB.DBTran(TransBegin, 1)
    
    ' 処理対象データ入力
    gSL_Select = "Select CASE WHEN NY.支給元2 = '' THEN SY.数字2 Else SY.数字1 END カード,"
    gSL_Select = gSL_Select & "CASE WHEN NY.支給先CD = '' THEN '" & Ts_Jisya & "' ELSE NY.支給先CD END 支給先,"
    gSL_Select = gSL_Select & "CASE WHEN NY.支給先工場CD = '' THEN '" & Ts_JisyaKj & "' ELSE NY.支給先工場CD END 支給先工場,"
'    gSL_Select = gSL_Select & "     ELSE CASE WHEN NY.支給先工場CD = '' THEN 支給先CD + '" & RTrim(Ts_JisyaKj) & "' Else NY.支給先CD END END 支給先,"
    gSL_Select = gSL_Select & "NY.*"
    gSL_Select = gSL_Select & "    From TNS支給入荷ﾋﾞｭｰ NY"
    gSL_Select = gSL_Select & "    Inner Join  名称マスタ SY"
    gSL_Select = gSL_Select & "        ON SY.区分名称 = 'TNS作成ﾀｲﾌﾟ'"
    gSL_Select = gSL_Select & "        And SY.値 = '" & TSS.処理区分 & "'"
'-----   -----   -----   -----   -----   -----   -----   -----   -----< 2007/08/03 ADD >
    gSL_Select = gSL_Select & " LEFT OUTER JOIN TNS送信対象取引先マスタ TT"
    gSL_Select = gSL_Select & "     ON TT.契約先 = NY.契約先"
    gSL_Select = gSL_Select & "     AND TT.取引先 = NY.自社仕入先"
    gSL_Select = gSL_Select & "     AND TT.受入 = NY.自社受入"
'-----   -----   -----   -----   -----   -----   -----   -----   -----< 2007/08/03 ADD END >
    If Ts_ReSendNo = 0 Then
        If Ts_SendKb = 0 Then
            gSL_Select = gSL_Select & " Where IsNull(NY.TNS作成区分, 0) = 0"
        Else
            gSL_Select = gSL_Select & " Where IsNull(NY.TNS作成区分2, 0) = 0"
        End If
    Else
        If Ts_SendKb = 0 Then
            gSL_Select = gSL_Select & " Where IsNull(NY.TNS作成区分, 0) = 1"
            gSL_Select = gSL_Select & "   And IsNull(NY.TNS作成管理NO, 0) = " & Ts_ReSendNo
        Else
            gSL_Select = gSL_Select & " Where IsNull(NY.TNS作成区分2, 0) = 1"
            gSL_Select = gSL_Select & "   And IsNull(NY.TNS作成管理NO2, 0) = " & Ts_ReSendNo
        End If
    End If

'-----   -----   -----   -----   -----   -----   -----   -----   -----< 2007/05/07 MNT >
'    gSL_Select = gSL_Select & "   And Substring(IsNull(NY.契約先, ''),1,4) = '" & TSS.送信先 & "'"
    gSL_Select = gSL_Select & "   And Substring( IsNull(NY.契約先, ''), 1, " & Len(TSS.送信先) & " ) = '" & TSS.送信先 & "'"
'-----   -----   -----   -----   -----   -----   -----   -----   -----< 2007/05/07 MNT END >
'-----   -----   -----   -----   -----   -----   -----   -----   -----< 2007/12/27 ADD >
    If Ts_ReSendNo = 0 Then
        gSL_Select = gSL_Select & "   And NY.入荷日 >= '" & TSS.送信開始日 & "'"
    End If
    If Ts_SendKb <> 0 Then
        gSL_Select = gSL_Select & "   And Substring( IsNull(NY.代行契約先, ''), 1, " & Len(TSS.代行送信先) & " ) = '" & TSS.代行送信先 & "'"
    End If
'-----   -----   -----   -----   -----   -----   -----   -----   -----< 2007/12/27 ADD END >

    gSL_Select = gSL_Select & " Order By カード,NY.仕入先,NY.支給先受入,NY.入荷日,NY.納入日,NY.納入便,NY.品番"
    If Not TSVRead(gSL_Select, , 1) Then
       gStr = "該当データなし "
       Call TNSSend_MsgDisp(gStr, gStr, 20)
       If Not Middle_TSRUpDate("03", "対象データなし") Then Exit Function
       
       If TSS.ゼロ件データ = 1 Then
          gStr = "ゼロ件データ作成なし "
          Call TNSSend_MsgDisp(gStr, gStr, 20)
          Call TSVClose(1)
          Call CisDB.DBTran(TransCommit, 1)
          TNSSEND_C04 = True
          Exit Function
       End If
       'ﾍｯﾀﾞｰ情報作成
       Call TNSSend_HeaderCrt
       Ts_SendCnt = Ts_SendCnt + 1
       
       GoTo TNSSEND_C04_ED
    End If
    
    Ts_DataCnt = CisDB.RecordCount
    
    If Not Ts_BarObj Is Nothing Then
        With Ts_FromObj
            Set .MeterPic = Ts_BarObj
            .MeterMax = Ts_DataCnt
        End With
    End If
    
    gStr = "処理対象データ件数 = " & Format(Ts_DataCnt, "#,##0")
    Call TNSSend_MsgDisp(gStr, gStr, 20)
    TSR.データ件数 = Ts_DataCnt
    If Not Middle_TSRUpDate("UPD") Then Exit Function
    
'------------------------------------------------------------ 2007/03/09 Insert
    '------------------ 品番ﾁｪｯｸありの時、対象品番のHEadkey(TNS)を求める
    C04_HeadKey = ""
    If TSS.品番CHK = 1 Then
       gStr = "品番チェック対象データ取得 "
       Call TNSSend_MsgDisp(gStr, gStr, 20)
       gSL_Select = "SELECT MAX(HEADKEY) HEADKEY FROM TNS支給送信品番確認テーブル "
       gSL_Select = gSL_Select & "   Where 契約先 = '" & TSS.送信先 & "'"
       gSL_Select = gSL_Select & " GROUP BY HEADKEY "
       If TSKRead(gSL_Select, 1) Then
          C04_HeadKey = TSK.HEADKEY
       Else
          gStr = "//*****************************************//"
          Call TNSSend_MsgDisp(gStr, gStr, 20)
          gStr = "//      品番チェック対象データ未登録       //"
          Call TNSSend_MsgDisp(gStr, gStr, 20)
          gStr = "//*****************************************//"
          Call TNSSend_MsgDisp(gStr, gStr, 20)
       End If
    End If
'------------------------------------------------------------ 2007/03/09 Insert
    
    'ﾍｯﾀﾞｰ情報作成
    Call TNSSend_HeaderCrt
    Ts_SendCnt = Ts_SendCnt + 1
    ' データレコード作成
    Brk_KeyAfter = ""
    C04_DataCnt = 0
    PCnt = 0
    Do Until Not TSV_RDSTS
    
       'ﾌﾞﾚｲｸｷｰ作成
       With TSV
'---------------------------------------------------------- 2007/03/13 Delete
''---------------------------------------------------------- 2007/03/06 Insert
'            C04_SaveHinbn = .品番
'            For gLong = 0 To UBound(Ts_KojyoFix) - 1
'                .品番 = Replace(.品番, Ts_KojyoFix(gLong), "")
'            Next gLong
''---------------------------------------------------------- 2007/03/06 Insert End
'---------------------------------------------------------- 2007/03/13 Delete End
            Ts_FixStr = .カード
            Brk_KeyBefore = Mid(Ts_FixStr, 1, 1)
            Ts_FixStr = Mid(.支給先, 1, 4) & Mid(.支給先工場, 1, 1)
            Ts_FixStr = Replace(Ts_FixStr, " ", "0")
            Brk_KeyBefore = Brk_KeyBefore & Mid(Ts_FixStr, 1, 5)
            Ts_FixStr = ""
            Brk_KeyBefore = Brk_KeyBefore & Mid(Ts_FixStr, 1, 1)
            Ts_FixStr = .支給先受入
            Brk_KeyBefore = Brk_KeyBefore & Mid(Ts_FixStr, 1, 2)
            Ts_FixStr = .仕入先
            Brk_KeyBefore = Brk_KeyBefore & Mid(Ts_FixStr, 1, 5)
            Ts_FixStr = .納入番号
            Brk_KeyBefore = Brk_KeyBefore & Mid(Ts_FixStr, 1, 5)
            Ts_FixStr = .入荷日
            Brk_KeyBefore = Brk_KeyBefore & Mid(Ts_FixStr, 1, 8)
            Ts_FixStr = .納入日
            Brk_KeyBefore = Brk_KeyBefore & Mid(Ts_FixStr, 1, 8)
            Ts_FixStr = .納入便
            Brk_KeyBefore = Brk_KeyBefore & Mid(Ts_FixStr, 1, 2)
       
       End With
       
    '----------------------------------------- 2007/03/09 Insert
       C04_OutFlg = True
       If TSS.品番CHK = 1 Then C04_OutFlg = HinbnChk

       If C04_OutFlg Then
    '----------------------------------------- 2007/03/09 Insert End
        
       Call CrtData(False)
       '--------------------> 入荷テーブル送信済み　更新
       If Ts_ReSendNo = 0 Then
          With CisDB
                If TSV.データ種別 = 1 Then
                    .SQL = "Update 入荷テーブル Set "
                    
                    If Ts_SendKb = 0 Then
                        .SQL = .SQL & "TNS作成区分 = 1,"
                        .SQL = .SQL & "TNS作成管理NO = " & TSR.履歴管理NO & ","
                        .SQL = .SQL & "TNS作成日 = '" & Format(Now(), "YYYYMMDD") & "',"
                        .SQL = .SQL & "TNS作成処理日 = Getdate(),"
                    Else
                        .SQL = .SQL & "TNS作成区分2 = 1,"
                        .SQL = .SQL & "TNS作成管理NO2 = " & TSR.履歴管理NO & ","
                        .SQL = .SQL & "TNS作成日2 = '" & Format(Now(), "YYYYMMDD") & "',"
                        .SQL = .SQL & "TNS作成処理日2 = Getdate(),"
                    End If
                    .SQL = .SQL & "更新日 = GETDATE(),"
                    .SQL = .SQL & "更新端末 = SubString(Host_Name(),1,20),"
                    .SQL = .SQL & "更新者 = '" & RTrim(gTanto) & "'"
                    .SQL = .SQL & " Where 発注管理NO = " & TSV.発注管理NO & ""
                Else
                    .SQL = "Update 受給検収テーブル Set "
                    If Ts_SendKb = 0 Then
                        .SQL = .SQL & "TNS作成区分 = 1,"
                        .SQL = .SQL & "TNS作成管理NO = " & TSR.履歴管理NO & ","
                        .SQL = .SQL & "TNS作成日 = '" & Format(Now(), "YYYYMMDD") & "',"
                        .SQL = .SQL & "TNS作成処理日 = Getdate(),"
                    Else
                        .SQL = .SQL & "TNS作成区分2 = 1,"
                        .SQL = .SQL & "TNS作成管理NO2 = " & TSR.履歴管理NO & ","
                        .SQL = .SQL & "TNS作成日2 = '" & Format(Now(), "YYYYMMDD") & "',"
                        .SQL = .SQL & "TNS作成処理日2 = Getdate(),"
                    End If
                    .SQL = .SQL & "更新日 = GETDATE(),"
                    .SQL = .SQL & "更新端末 = SubString(Host_Name(),1,20),"
                    .SQL = .SQL & "更新者 = '" & RTrim(gTanto) & "'"
                    .SQL = .SQL & " Where 受給管理NO = " & TSV.発注管理NO & ""
                End If
                 
               If Not .DBExec Then
                  gStr = "入荷ﾃｰﾌﾞﾙ更新ｴﾗｰ(21)"
                  Call TNSSend_MsgDisp(gStr, gStr, 20)
                  Call Middle_TSRUpDate("21", "入荷ﾃｰﾌﾞﾙ更新ｴﾗｰ")
                  Exit Function
               End If
          End With
        End If
    '----------------------------------------- 2007/03/09 Insert
       End If
    '----------------------------------------- 2007/03/09 Insert End
           
       If Not Ts_BarObj Is Nothing Then
          PCnt = PCnt + 1
          With Ts_FromObj
              .MeterUpdate PCnt
          End With
       End If
       
       Call TSVReadNext(1)
    Loop
    
    
    Call CrtData(True)
    
TNSSEND_C04_ED:
    Call TSVClose(1)
    Call CisDB.DBTran(TransCommit, 1)
    
    'ﾄﾚｰﾗｰ情報作成
    Ts_SendCnt = Ts_SendCnt + 1
    Call TNSSend_TrailerCrt
        
    'TNS送信ﾃﾞｰﾀ出力
    If Not TNSSend_DataPut Then
       Exit Function
    End If
    
    gStr = "送信データ件数     = " & Format(Ts_SendCnt, "#,##0")
    Call TNSSend_MsgDisp(gStr, gStr, 20)
    gStr = "履歴管理番号       = " & TSR.履歴管理NO
    Call TNSSend_MsgDisp(gStr, gStr, 20)
    gStr = "送信管理番号       = " & TSR.送信管理NO
    Call TNSSend_MsgDisp(gStr, gStr, 20)
    gStr = "送信対象ファイル   = " & Ts_SendBackName
    Call TNSSend_MsgDisp(gStr, gStr, 20)
    
    TSR.送信件数 = Ts_SendCnt
    TSR.対象ファイル = Ts_SendBackName
    TSR.ERR件数 = C04_HinErrCnt
    If Not Middle_TSRUpDate("UPD") Then Exit Function
    
    
    
    TNSSEND_C04 = True
    Exit Function
TNSSEND_C04_ER:
    gStr = "TNS送信データ作成ｴﾗｰ" & vbCrLf & _
           "ｴﾗｰ番号 " & Str(ERR.Number) & vbCrLf & _
            ERR.Source & " でｴﾗｰが発生しました。" & vbCrLf & _
            ERR.Description
    Call TNSSend_MsgDisp("", gStr, 10, "ER")
    Call CisDB.DBTran(TransRollback, 1)
End Function
'+--------------------------------------------------------+
'+        データ部作成
'+--------------------------------------------------------+
Private Sub CrtData(PutFlg As Boolean)
    
    If Brk_KeyAfter <> Brk_KeyBefore Or C04_DataCnt >= 11 Or PutFlg Then
       If Brk_KeyAfter <> "" Then
          If C04_DataCnt < 11 Then
             Ts_Data = Ts_Data & Space(18 * (11 - C04_DataCnt))
          End If
          Ts_FixStr = Save_STorcd
          Ts_Data = Ts_Data & Mid(Ts_FixStr, 1, 4)
          
          Ts_FixStr = Ts_Data
          Ts_Data = Mid(Ts_FixStr, 1, Ts_Length)
        
          Ts_TNSPutData = Ts_TNSPutData & Ts_Data
          
          Ts_SendCnt = Ts_SendCnt + 1
          C04_DataCnt = 0
       End If
       Ts_Data = "D"
       Ts_FixStr = Brk_KeyBefore
       Ts_Data = Ts_Data & Mid(Ts_FixStr, 1, 37)
       
       Brk_KeyAfter = Brk_KeyBefore
       Save_STorcd = TSV.支給元
    End If
    
    Ts_FixStr = TSV.品番
    Ts_Data = Ts_Data & Mid(Ts_FixStr, 1, 12)
    Ts_FixStr = Format(TSV.入荷数, "000000")
    Ts_Data = Ts_Data & Mid(Ts_FixStr, 1, 6)
    
    C04_DataCnt = C04_DataCnt + 1

End Sub
''**********************************************
''*  入荷実績品番と取引先納品書データとの照合
''**********************************************
Private Function HinbnChk() As Boolean
    HinbnChk = False
    
    gSL_Select = " SELECT "
    gSL_Select = gSL_Select & " 品番 "
    gSL_Select = gSL_Select & " FROM TNS支給送信品番確認テーブル"
    gSL_Select = gSL_Select & " WHERE 契約先 = '" & TSS.送信先 & "'"
    gSL_Select = gSL_Select & "   AND HEADKEY = '" & C04_HeadKey & "'"
    gSL_Select = gSL_Select & "   AND 品番 = '" & TSV.品番 & "'"

    If Not TSKRead(gSL_Select, 1) Then
       C04_HinErrCnt = C04_HinErrCnt + 1
'
       With THE
            .履歴管理NO = TSR.履歴管理NO
            .契約先 = TSS.送信先
            .FCD = TSS.FCODE
            .HEADKEY = C04_HeadKey
'            .品番 = C04_SaveHinbn          ' 2007/03/13 Delete
            .品番 = TSV.品番
            .仕入先 = TSV.自社仕入先
            .受入 = TSV.自社受入
            .REC = ""
            .エラー = 2
            .発注管理NO = TSV.発注管理NO
            .納入番号 = RTrim(TSV.納入番号)
            .納入日 = TSV.納入日
            .変換日付 = Format(Now, "YYYYMMDD")
            .変換時刻 = Format(Now, "HHNN")
            .変換端末 = CisFun.WSID
       End With
       Call THEInsert
'
       Exit Function
    End If
    
    HinbnChk = True
    
End Function
'○－－－－－－－－－－－－－－－－－－－－－－－－○
'○     変　換　エ　ラ　ー　リ　ス　ト　出　力
'○－－－－－－－－－－－－－－－－－－－－－－－－○
Public Sub C04_ErrList(Optional DispFlg As Boolean = False)
    Static wsCnt        As Long
    wsCnt = 0

    On Error GoTo C04_ErrList_Er:

    gStr = "//*****************************************//"
    Call TNSSend_MsgDisp(gStr, gStr, 20)
    gStr = "//      品番チェックエラーリスト発行       //"
    Call TNSSend_MsgDisp(gStr, gStr, 20)
    gStr = "//*****************************************//"
    Call TNSSend_MsgDisp(gStr, gStr, 20)

'   < 印刷データ読み込み (送信先未登録)>
    gSL_Select = "Select TE.契約先,TE.FCD,TE.仕入先,TE.受入,TE.品番,HN.表示品番,HN.品名,HT.背番号,HN.車型,TE.納入番号,"
    gSL_Select = gSL_Select & " SUBSTRING(HT.サイクル,1,1) + '-' + SUBSTRING(HT.サイクル,2,2) + '-' + SUBSTRING(HT.サイクル,4,2) サイクル,"
    gSL_Select = gSL_Select & " CASE WHEN HT.変更日 < '" & Format(Now, "YYYYMMDD") & "'" & " THEN HT.収容数1 "
    gSL_Select = gSL_Select & "      ELSE HT.収容数2 END 収容数,"
    gSL_Select = gSL_Select & " TR.略称 仕入先名,TE.納入日 "
    gSL_Select = gSL_Select & " From TNS支給送信変換エラーテーブル TE"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "     ON  HN.品番 = TE.品番 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT "
    gSL_Select = gSL_Select & "     ON  HT.品番 = TE.品番 "
    gSL_Select = gSL_Select & "     AND HT.取引先 = TE.仕入先 "
    gSL_Select = gSL_Select & "     AND HT.受入 = TE.受入 "
    gSL_Select = gSL_Select & "     AND HT.品目 = 1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "     ON  TR.取引先CD = TE.仕入先 "
    gSL_Select = gSL_Select & "     AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & " WHERE 履歴管理NO = " & TSR.履歴管理NO
    gSL_Select = gSL_Select & " AND   エラー = 2"
    gSL_Select = gSL_Select & " ORDER BY TE.仕入先,HT.背番号,TE.品番 "
    If THERead(gSL_Select) Then
'
        gStr = "           エラー件数 ＝ " & Format(CisDB.RecordCount, "#,###")
        Call TNSSend_MsgDisp(gStr, gStr, 20)
        
        Dim CisSimplePrint As New CisSimplePrint

        With CisSimplePrint
            .PrintTitle = "※ (" & THE.契約先 & "-" & THE.FCD & ")【入荷実績送信】エラーリスト(取引先未登録) ※"
            .Preservation = App.Title
            .PrintSize = A4
            .PrintOrientation = Landscape
            .PrintHeadLine = あり
            .PrintBodyGyo = 11
            
            .MaxCount = CisDB.RecordCount
            
            .pHeadItemSet 0, 左, "№"
            .pHeadItemSet 1, 中, "仕入先"
            .pHeadItemSet 2, 中, ""
            .pHeadItemSet 3, 中, "受"
            .pHeadItemSet 4, 中, "背番号"
            .pHeadItemSet 5, 中, "品  番"
            .pHeadItemSet 6, 中, "品  名"
            .pHeadItemSet 7, 中, "車型"
            .pHeadItemSet 8, 中, "サイクル"
            .pHeadItemSet 9, 中, "収容数"
            .pHeadItemSet 10, 中, "納番"
            .pHeadItemSet 11, 中, "納入日"

            .pBodyItemSet 0, 2, 右, "##"
            .pBodyItemSet 1, 6
            .pBodyItemSet 2, 20
            .pBodyItemSet 3, 2
            .pBodyItemSet 4, 6
            .pBodyItemSet 5, 25
            .pBodyItemSet 6, 30
            .pBodyItemSet 7, 4
            .pBodyItemSet 8, 8
            .pBodyItemSet 9, 6, 右
            .pBodyItemSet 10, 5
            .pBodyItemSet 11, 8

            .BodyInit
            .PrintStatus = 1

            If DispFlg Then
               If Not .CheckPrtCsv Then
                  Call THEClose
                  Exit Sub
               End If
            End If

            Do Until (Not THE_RDSTS)
                wsCnt = wsCnt + 1
                .pBodyValueSet 0, Format(wsCnt, "0")
                .pBodyValueSet 1, THE.仕入先
                .pBodyValueSet 2, THE.仕入先名
                .pBodyValueSet 3, THE.受入
                .pBodyValueSet 4, THE.背番号
                .pBodyValueSet 5, THE.表示品番
                .pBodyValueSet 6, CisFun.Mid2(THE.品名, 1, 30)
                .pBodyValueSet 7, THE.車型
                .pBodyValueSet 8, THE.サイクル
                .pBodyValueSet 9, Format(THE.収容数, "##,###")
                .pBodyValueSet 10, THE.納入番号
                .pBodyValueSet 11, Mid(THE.納入日, 3, 2) & "/" & Mid(THE.納入日, 5, 2) & "/" & Mid(THE.納入日, 7, 2)
                
                .BodySet

                Call THEReadNext
            Loop

            .PrintStart

        End With
        Call THEClose

    End If
    Exit Sub
C04_ErrList_Er:
    gStr = "品番チェックエラーリスト発行エラー" & vbCrLf & _
           "ｴﾗｰ番号 " & Str(ERR.Number) & vbCrLf & _
            ERR.Source & " でｴﾗｰが発生しました。" & vbCrLf & _
            ERR.Description
    Call TNSSend_MsgDisp("", gStr, 10, "ER")
    
End Sub


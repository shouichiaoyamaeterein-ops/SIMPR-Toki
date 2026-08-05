Attribute VB_Name = "TextCnvNaiji"
Option Explicit
'+----------------------------------------------------------+
'+
'+       モジュール名  :テキスト変換 内示モジュール
'+       モジュールID  :TextCnvNaiji.BAS
'+       作成日        :2003/12/11   CIS
'+       更新日        :2004/11/10   CIS 複数契約先対応
'+       更新日        :2004/12/13   CIS 納入先工場を取引先受入マスタより求める
'+       更新日        :2005/12/12   CIS トヨタ内示レイアウト変更による修正
'+                      2006/01/24   CIS 文言データの処理一部修正
'+       更新日         2008/04/16   CIS オプション４に同一データ存在時処理（0:上書き 1:数量加算)
'+      オプション2 : データ保持期間
'+      オプション3 : 内示処理区分　(0:初回,1:計変)
'+      オプション4 : 同一データ存在時処理(0:上書き 1:数量加算)         2008/04/16
'+       更新日        :2008/05/09   CIS 車体精工対応
'+       更新日        :2009/01/12   CIS 受注取引先CD変換ﾏｽﾀ 追加 → 取引先CD/受入のｺｰﾄﾞ変換
'+       更新日        :2010/02/15   CIS 愛知機械対応
'+
'+----------------------------------------------------------+
Public nj_YYMM          As String
Public nj_Renban        As Long

Dim bNJPutFlg           As Boolean

Public nj_CTorcd()      As String
Public nj_CYYMM()       As String
Public nj_HKaisu        As Long
Dim nj_ZnjYM()          As String
Dim TorFlg              As Boolean

Global FirstFlg         As Boolean                      '2008/04/16

Public Function NaijiConVert(lErr As String) As Boolean
    NaijiConVert = False
    
    lErr = ""
    On Error GoTo NaijiConVert_Er:
    
    ' 内示年月（当月）を求める
    If nj_YYMM = "" Then
        ReDim nj_CTorcd(1)
        ReDim nj_CYYMM(1)
        ReDim nj_ZnjYM(10)
        
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/09 UPDATE START
'        gStr = Mid(tx_HeadYmd, 1, 6) & "01"
'        With CisFun
'            .Date_RsFormat = fYM
'            .Date_Add gStr, M, 1, LastDate
'            nj_YYMM = .Date_Result
'        End With
'        nj_ZnjYM(0) = nj_YYMM
        '====================================== 2010/02/15 Start
'        If TKT.TEXTレイアウト = 6 Then
'            nj_YYMM = "20" & Mid(tx_DataStr, 40, 4)
'            nj_ZnjYM(0) = nj_YYMM
'        Else
'            gStr = Mid(tx_HeadYmd, 1, 6) & "01"
'            With CisFun
'                .Date_RsFormat = fYM
'                .Date_Add gStr, M, 1, LastDate
'                nj_YYMM = .Date_Result
'            End With
'            nj_ZnjYM(0) = nj_YYMM
'        End If
        Select Case TKT.TEXTレイアウト
               Case 6
                        nj_YYMM = "20" & Mid(tx_DataStr, 40, 4)
                        nj_ZnjYM(0) = nj_YYMM
               Case 7
                        nj_YYMM = Mid(tx_DataStr, 157, 6)
                        nj_ZnjYM(0) = nj_YYMM
               Case Else
                        gStr = Mid(tx_HeadYmd, 1, 6) & "01"
                        With CisFun
                            .Date_RsFormat = fYM
                            .Date_Add gStr, M, 1, LastDate
                            nj_YYMM = .Date_Result
                        End With
                        nj_ZnjYM(0) = nj_YYMM
        End Select
        '====================================== 2010/02/15 End
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/09 UPDATE END
        With CisDB
            .SQL = "DELETE From 内示TEXTテーブル"
            .SQL = .SQL & " Where FILECODE = '" & TKT.FILECD & "'"
            .SQL = .SQL & " And   HEAD作成日 = '" & tx_HeadYmd & "'"
            .SQL = .SQL & " And   HEADSEQ = '" & tx_HeadSeq & "'"
            Call .DBExec
            
            Call ProcLogPut("内示TEXTテーブル 削除 契約先=" & TKT.契約先 & ",FILECODE = " & TKT.FILECD & ",HEAD作成日 = " & tx_HeadYmd & ",HEADSEQ = '" & tx_HeadSeq & ",削除件数=" & CisDB.RecordCount)

            If TKT.オプション2 <> 0 Then
                With CisFun
                    .Date_RsFormat = fYM
                    .Date_Add nj_YYMM, M, -3, LastDate
                    gStr = .Date_Result
                End With
                .SQL = "DELETE From 内示TEXTテーブル"
                .SQL = .SQL & " Where FILECODE = '" & TKT.FILECD & "'"
                .SQL = .SQL & " And   対象年月 <= '" & gStr & "'"
                Call .DBExec
                Call ProcLogPut("内示TEXTテーブル 削除 契約先=" & TKT.契約先 & ",対象年月 <= " & gStr & ",削除件数=" & CisDB.RecordCount)
                If TKT.オプション3 = 0 Then    ' 初回のみ
'                   .SQL = "DELETE From 内示テーブル"
'                   .SQL = .SQL & " Where 内示年月 <= '" & gStr & "'"
'                   .SQL = .SQL & "   And 契約先 = '" & TKT.契約先 & "'"
'                   Call .DBExec
'                   Call ProcLogPut("内示テーブル 削除 契約先=" & TKT.契約先 & ",内示年月 <= " & gStr & ",削除件数=" & CisDB.RecordCount)
'
'                   .SQL = "DELETE From 内示前回テーブル"
'                   .SQL = .SQL & " Where 内示年月 <= '" & gStr & "'"
'                   .SQL = .SQL & "   And 契約先 = '" & TKT.契約先 & "'"
'                   Call .DBExec
'                   Call ProcLogPut("内示前回テーブル 削除 契約先=" & TKT.契約先 & "内示年月 <= " & gStr & ",削除件数=" & CisDB.RecordCount)
                End If
            End If
'            ' 変更回数最大を求める
'            nj_HKaisu = 0
'            If TKT.オプション3 <> 0 Then    ' 計変のみ
'                gSL_Select = "Select Max(Isnull(変更回数,0)) From 内示TEXTテーブル"
'                gSL_Select = gSL_Select & " Where 内示年月 = '" & nj_YYMM & "'"
'                gSL_Select = gSL_Select & "   And 契約先 = '" & TKT.契約先 & "'"
'                If Not NTTRead(gSL_Select, 1) Then
'                   nj_HKaisu = NTT.変更回数
'                End If
'                nj_HKaisu = nj_HKaisu + 1
'            End If
        End With
    End If
    ' 文言情報件数加算
    tx_MonFlg = 0
    With typConvert
        If .kMonCnt <> 0 Then
            For txCnt1 = 1 To .kMonCnt
                If Mid(tx_DataStr, .mStart(txCnt1), .mLeng(txCnt1)) = .mChar(txCnt1) Then
                   tx_MonFlg = txCnt1
                   Exit For
                End If
            Next txCnt1
        End If
    End With
    If tx_MonFlg <> 0 Then
       tx_BunCnt = tx_BunCnt + 1
       
       NaijiConVert = True                  '2006/01/24
       Exit Function                        '2006/01/24
    End If
    
    ' 項目定義の連番を求める
    With typConvert
        If Not txDataNOGet Then
           lErr = "S1"
           Call TERPut("S1", "テキスト設定エラー", "テキスト設定情報を確認して下さい")
           Exit Function
        End If
    ' カレンダ情報
'        If Mid(tx_DataStr, 2, 1) = "C" Then            '2005/12/12
        If Mid(tx_DataStr, 2, 1) = "C" Or Mid(tx_DataStr, 7, 1) = "C" Then
            Call ItemsClearNTT
                
            For txCnt1 = 1 To .dDataMax
                
                If Not txDataSetGet(txCnt1) Then Exit For
                
                Select Case .dDSetData(tx_DataNo, txCnt1)
                    Case "納入先": NTT.納入先 = tx_SetData
                    Case "内示年月": NTT.内示年月 = tx_SetData
                End Select
            Next txCnt1
            
'            If Trim(NTT.内示年月) <> "" Then               '2005/1212
            If Trim(NTT.内示年月) <> "" And Len(NTT.内示年月) < 6 Then
               NTT.内示年月 = "20" & NTT.内示年月
            End If
            
            For txCnt1 = 0 To UBound(nj_CTorcd) - 1
                If Trim(nj_CTorcd(txCnt1)) = "" Then
                   nj_CTorcd(txCnt1) = NTT.納入先
                   nj_CYYMM(txCnt1) = NTT.内示年月
                   NTT.納入先 = ""
                   NTT.内示年月 = ""
                End If
            Next txCnt1
            
            If Trim(NTT.納入先) <> "" Then
               txCnt2 = UBound(nj_CTorcd)
               ReDim Preserve nj_CTorcd(txCnt2 + 1)
               ReDim Preserve nj_CYYMM(txCnt2 + 1)
               nj_CTorcd(txCnt2) = NTT.納入先
               nj_CYYMM(txCnt2) = NTT.内示年月
               ' 前回内示年月セット
               For txCnt1 = 0 To UBound(nj_ZnjYM) - 1
                   If Trim(nj_ZnjYM(txCnt1)) = NTT.内示年月 Then Exit For
                   If Trim(nj_ZnjYM(txCnt1)) = "" Then
                      nj_ZnjYM(txCnt1) = NTT.内示年月
                      Exit For
                   End If
               Next txCnt1
            End If
            
            On Error GoTo 0
            NaijiConVert = True
            Exit Function
        End If
            
    ' 前回内示情報作成
        If Trim(nj_ZnjYM(0)) <> "" Then
           For txCnt1 = 0 To UBound(nj_ZnjYM) - 1
               If Trim(nj_ZnjYM(txCnt1)) = "" Then Exit For
                With CisDB
                     .SQL = "delete 内示前回テーブル"
                     .SQL = .SQL & " Where 内示年月 = '" & nj_ZnjYM(txCnt1) & "'"
                     .SQL = .SQL & "   And 契約先 = '" & TKT.契約先 & "'"
                     .DBExec

                     .SQL = "Insert into 内示前回テーブル "
                     .SQL = .SQL & " Select * From 内示テーブル "
                     .SQL = .SQL & "     Where 内示年月 = '" & nj_ZnjYM(txCnt1) & "'"
                     .SQL = .SQL & "       And 契約先 = '" & TKT.契約先 & "'"
                     .DBExec
                End With
           Next txCnt1
           nj_ZnjYM(0) = ""
        End If
        ' ｶﾚﾝﾀﾞｰ情報ありでｶﾚﾝﾀﾞｰ情報未登録時は、ｶﾚﾝﾀﾞｰ情報の１件目の年月とする。
        If Trim(nj_CYYMM(0)) <> "" Then
            nj_YYMM = nj_CYYMM(0)
        End If
    
    ' SQL文作成
        If tx_DataNo = .kStartRecNo Then
           Call ItemsClearNTT
           NTT.FILECODE = TKT.FILECD
           NTT.HEAD作成日 = tx_HeadYmd
           NTT.HEADSEQ = tx_HeadSeq
           NTT.対象年月 = nj_YYMM
        End If
        TorFlg = False
        
        For txCnt1 = 1 To .dDataMax
            
            If Not txDataSetGet(txCnt1) Then Exit For
            
            Select Case .dDSetData(tx_DataNo, txCnt1)
                Case "状態区分": NTT.状態区分 = tx_SetData
                Case "納品書不作成区分": NTT.納品書不作成区分 = tx_SetData
                Case "仕入先工区": NTT.仕入先工区 = tx_SetData
                Case "品番": NTT.品番 = tx_SetData
                Case "層別区分": NTT.層別区分 = tx_SetData
                                If NTT.層別区分 <> "4" Then
                                   NTT.号補区分 = "G"
                                Else
                                   NTT.号補区分 = "H"
                                End If
                Case "管理NO": NTT.管理NO = tx_SetData
                Case "納入先"
                     If Trim(tx_SetData) = "" Then
                '+--------------------------------------- 2004/12/13 Insert
                        NTT.納入先 = NonyuGet(TKT.契約先, Trim(NTT.受入))
                '+--------------------------------------- 2004/12/13 Insert End
                '+--------------------------------------- 2004/12/13 Delete
'                        NTT.納入先 = TKT.契約先
                '+--------------------------------------- 2004/12/13 Delete End
                     Else
                        NTT.納入先 = tx_SetData
                        NTT.納入先 = Replace(NTT.納入先, " ", "")                   '2005/12/12
                        TorFlg = True
                     End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/09 INSERT START
                Case "工区": NTT.納入先 = TKT.車体精工コード + tx_SetData
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/09 INSERT END
                Case "持込先": NTT.持込先 = tx_SetData
                Case "持込先工場": NTT.持込先工場 = tx_SetData
                Case "支給先受入": NTT.支給先受入 = tx_SetData
                Case "出庫部署": NTT.出庫部署 = tx_SetData
                Case "支給先担当者": NTT.支給先担当者 = tx_SetData
                Case "支給係": NTT.支給係 = tx_SetData
                Case "受入": NTT.受入 = tx_SetData
                Case "整備": NTT.整備 = tx_SetData
                Case "担当者": NTT.担当者 = tx_SetData
                Case "収容数": NTT.収容数 = tx_SetDataLong
                Case "背番号": NTT.背番号 = tx_SetData
                Case "サイクル": NTT.サイクル = tx_SetData
                Case "納品書発行枚数": NTT.納品書発行枚数 = tx_SetDataLong
                Case "かんばん回転枚数": NTT.かんばん回転枚数 = tx_SetDataLong
                Case "かんばん増減符号": NTT.かんばん増減符号 = tx_SetData
                Case "かんばん増減枚数": NTT.かんばん増減枚数 = tx_SetDataLong
                Case "適用開始年月日"
                     gStr = Trim(tx_SetData)
'2005/12/12 Start
'                     If gStr = "000000" Then gStr = ""
'                     If gStr <> "" Then
'                        If gStr = "999999" Then
'                           NTT.適用開始年月日 = ""
'                        Else
'                           If Mid(gStr, 1, 2) < "90" Then
'                              NTT.適用開始年月日 = "20" & tx_SetData
'                           Else
'                              NTT.適用開始年月日 = "19" & tx_SetData
'                           End If
'                       End If
'                     End If
                     If gStr = "000000" Or gStr = "00000000" Then gStr = ""
                     If gStr <> "" Then
                        If gStr = "999999" Or gStr = "99999999" Then
                           NTT.適用開始年月日 = ""
                        Else
                           If Len(gStr) < 7 Then
                              If Mid(gStr, 1, 2) < "90" Then
                                 NTT.適用開始年月日 = "20" & tx_SetData
                              Else
                                 NTT.適用開始年月日 = "19" & tx_SetData
                              End If
                           Else
                              NTT.適用開始年月日 = tx_SetData
                           End If
                       End If
                     End If
'2005/12/12 End
                Case "適用終了年月日"
                     gStr = Trim(tx_SetData)
'2005/12/12 Start
'                     If gStr = "000000" Then gStr = ""
'                     If gStr <> "" Then
'                        If gStr = "999999" Then
'                           NTT.適用終了年月日 = ""
'                        Else
'                           If Mid(gStr, 1, 2) < "90" Then
'                              NTT.適用終了年月日 = "20" & tx_SetData
'                           Else
'                              NTT.適用終了年月日 = "19" & tx_SetData
'                           End If
'                       End If
'                     End If
                     If gStr = "000000" Or gStr = "00000000" Then gStr = ""
                     If gStr <> "" Then
                        If gStr = "999999" Or gStr = "99999999" Then
                           NTT.適用終了年月日 = ""
                        Else
                           If Len(gStr) < 7 Then
                              If Mid(gStr, 1, 2) < "90" Then
                                 NTT.適用終了年月日 = "20" & tx_SetData
                              Else
                                 NTT.適用終了年月日 = "19" & tx_SetData
                              End If
                           Else
                              NTT.適用終了年月日 = tx_SetData
                           End If
                       End If
                     End If
'2005/12/12 End
                Case "内示数": NTT.内示数 = tx_SetDataLong
                Case "翌月内示数": NTT.翌月内示数 = tx_SetDataLong
                Case "翌々月内示数": NTT.翌々月内示数 = tx_SetDataLong
                Case "当初内示数": NTT.当初内示数 = tx_SetDataLong
                Case "変更数符号": NTT.変更数符号 = tx_SetData
                Case "変更数": NTT.変更数 = tx_SetDataLong
                Case "調整数符号": NTT.調整数符号 = tx_SetData
                Case "調整数": NTT.調整数 = tx_SetDataLong
                Case "納入区分": NTT.納入区分 = tx_SetData

                Case "内外区分": NTT.内外区分 = tx_SetData
                Case "現旧区分": NTT.現旧区分 = tx_SetData
                Case "管理区分": NTT.管理区分 = tx_SetData
                Case "生産区分": NTT.生産区分 = tx_SetData
                Case "AB区分": NTT.AB区分 = tx_SetData
                Case "品名": NTT.品名 = tx_SetData
                Case Else
                    gStr = .dDSetData(tx_DataNo, txCnt1)
                    If Mid(gStr, 1, 4) = "内示数量" Then
                       NTT.内示数量(CisFun.Val2(Mid(gStr, 5, 2))) = tx_SetDataLong
                    End If
            End Select
            
        Next txCnt1
    
        If tx_DataNo = .kEndRecNo Then
           nj_Renban = nj_Renban + 1
           With NTT
                .連番 = nj_Renban
                .変更回数 = nj_HKaisu
                If .層別区分 = "2" Or .層別区分 = "3" Or .層別区分 = "5" Then
                   .支給先受入 = .受入
                   .担当 = .支給先担当者
                   .支給先担当者 = .担当者
                   .内外区分 = ""
                   .現旧区分 = ""
                End If
                If .層別区分 = "5" Then
                    .支給係 = ""
                Else
                    .AB区分 = ""
                End If
                
                ' 内示年月 取得
                For txCnt1 = 0 To UBound(nj_CTorcd) - 1
                    If Trim(nj_CTorcd(txCnt1)) = "" Then Exit For
                    If Trim(nj_CTorcd(txCnt1)) = Trim(Mid(NTT.納入先, 1, 4)) Then
                       NTT.内示年月 = nj_CYYMM(txCnt1)
                       Exit For
                    End If
                Next txCnt1
                If Trim(NTT.内示年月) = "" Then
                   NTT.内示年月 = nj_YYMM
                End If
                ' 工区変換
                .置換工区 = .仕入先工区
                For txCnt1 = 0 To UBound(tx_NjKouku) - 1
                    If RTrim(TKT.契約先) & RTrim(.仕入先工区) = tx_NjKouku(txCnt1) Then
                       .置換工区 = tx_NjKoukuX(txCnt1)
                       Exit For
                    End If
                Next txCnt1
                
                .受入 = Trim(.受入)
                If Len(.受入) > 2 Then
                   .納入先 = .受入
                   .受入 = ""
                End If
                .契約先 = TKT.契約先
           End With
           
'============================================================= 2009/01/12 Start
           If RTrim(NTT.納入先) <> "" Then
              '①取引先CD/受入にて変換マスタを検索  → 取引先CD/受入 を変換
              gSL_Select = "SELECT MIN(変換先CD) 変換先CD,MIN(変換先受入) 変換先受入 "
              gSL_Select = gSL_Select & " FROM 受注取引先CD変換マスタ"
              gSL_Select = gSL_Select & " WHERE 変換元CD = '" & RTrim(NTT.納入先) & "'"
              gSL_Select = gSL_Select & " AND   変換元受入 = '" & RTrim(NTT.受入) & "'"
              Call JTHRead(gSL_Select, 1)
              If RTrim(JTH.変換先CD) = "" Then
                 '②取引先CDだけで変換マスタを検索  → 取引先CD を変換
                 gSL_Select = "SELECT MIN(変換先CD) 変換先CD "
                 gSL_Select = gSL_Select & " From 受注取引先CD変換マスタ"
                 gSL_Select = gSL_Select & " WHERE 変換元CD = '" & RTrim(NTT.納入先) & "'"
                 Call JTHRead(gSL_Select, 1)
                 If RTrim(JTH.変換先CD) <> "" Then
                    NTT.納入先 = JTH.変換先CD
                 End If
              Else
                 NTT.納入先 = JTH.変換先CD
                 NTT.受入 = JTH.変換先受入
              End If
           End If
           If RTrim(NTT.品番) <> "" Then
              gSL_Select = "SELECT 品番 FROM 品番マスタ "
              gSL_Select = gSL_Select & " WHERE REPLACE(品番,'-','') = '" & RTrim(NTT.品番) & "'"
              If HNMRead(gSL_Select, 1) Then
                 NTT.品番 = HNM.品番
              End If
           End If
'============================================================= 2009/01/12 End
           
           Call NTTInsert
           
           tx_OutCnt = tx_OutCnt + 1    ' 出力件数加算
           
'           ' 内示テーブル　更新
'[][][][][][][][][][] 2006/04/18 Add Start [][][][][][][][][][]
           If FirstFlg = False Then
              With CisDB
                   .SQL = "delete 内示テーブル"
                   .SQL = .SQL & " Where 内示年月 = '" & NTT.対象年月 & "'"
                   .SQL = .SQL & "   And 契約先 = '" & TKT.契約先 & "'"
                   .SQL = .SQL & "   And 登録FCODE = '" & TKT.FILECD & "'"
                   .DBExec
              End With
              FirstFlg = True
           End If
'           If Not TorFlg Then
'''           If Not TorFlg And TKT.オプション4 = 0 Then                       '2008/04/16 DEL
'''              gSL_Select = "Delete 内示テーブル "
'''              gSL_Select = gSL_Select & " Where 内示年月 = '" & RTrim(NTT.内示年月) & "'"
'''              gSL_Select = gSL_Select & " And   品番 = '" & RTrim(NTT.品番) & "'"
'''              gSL_Select = gSL_Select & " And   SUBSTRING(納入先,1,4) = '" & CisFun.Mid2(NTT.納入先, 1, 4) & "'"
'''              gSL_Select = gSL_Select & " And   受入 = '" & RTrim(NTT.受入) & "'"
'''              gSL_Select = gSL_Select & " And   内示区分 = ''"
'''              CisDB.SQL = gSL_Select
'''              Call CisDB.DBExec
'''           End If
'[][][][][][][][][][] 2006/04/18 Add End   [][][][][][][][][][]
           
           bNJPutFlg = False
           gSL_Select = "SELECT * From 内示テーブル "
           gSL_Select = gSL_Select & " Where 内示年月 = '" & RTrim(NTT.内示年月) & "'"
           gSL_Select = gSL_Select & " And   品番 = '" & RTrim(NTT.品番) & "'"
           gSL_Select = gSL_Select & " And   納入先 = '" & RTrim(NTT.納入先) & "'"
           gSL_Select = gSL_Select & " And   受入 = '" & RTrim(NTT.受入) & "'"
           gSL_Select = gSL_Select & " And   内示区分 = ''"
           If Not NJTRead(gSL_Select, 1) Then
              Call ItemsClearNJT
              With NJT
                  .内示年月 = NTT.内示年月
                  .品番 = NTT.品番
                  .納入先 = NTT.納入先
                  .受入 = NTT.受入
              End With
              bNJPutFlg = True
           End If

           With NJT
                .契約先 = TKT.契約先
                .工区 = NTT.置換工区
                .層別区分 = NTT.層別区分
                .号補区分 = NTT.号補区分
                .適用開始年月日 = NTT.適用開始年月日
                .適用終了年月日 = NTT.適用終了年月日
'                .内示数 = NTT.内示数
'                .翌月内示数 = NTT.翌月内示数
'                .翌々月内示数 = NTT.翌々月内示数
                .管理NO = NTT.管理NO
                .収容数 = NTT.収容数
                .背番号 = NTT.背番号
'[][][][][][][][][][] 2008/04/16 Add Start [][][][][][][][][][]
                If Not bNJPutFlg And TKT.オプション4 = 1 Then
                    .内示数 = .内示数 + NTT.内示数
                    .翌月内示数 = .翌月内示数 + NTT.翌月内示数
                    .翌々月内示数 = .翌々月内示数 + NTT.翌々月内示数
                    For txCnt1 = 1 To 31
                        .内示数量(txCnt1) = .内示数量(txCnt1) + NTT.内示数量(txCnt1)
                    Next txCnt1
                    If NTT.内示数 = 0 Then
                        For txCnt1 = 1 To 31
                            .内示数 = .内示数 + NTT.内示数量(txCnt1)
                        Next txCnt1
                    End If
                Else
                    .内示数 = NTT.内示数
                    .翌月内示数 = NTT.翌月内示数
                    .翌々月内示数 = NTT.翌々月内示数
                    For txCnt1 = 1 To 31
                        .内示数量(txCnt1) = NTT.内示数量(txCnt1)
                    Next txCnt1
                    If NTT.内示数 = 0 Then
                        For txCnt1 = 1 To 31
                            .内示数 = .内示数 + NTT.内示数量(txCnt1)
                        Next txCnt1
                    End If
                End If
'[][][][][][][][][][] 2008/04/16 Add End [][][][][][][][][][]
                
                .サイクル = NTT.サイクル
                If .サイクル <> "" Then
                   .サイクル = Replace(.サイクル, " ", "0")
                End If
                .担当 = NTT.担当者
                .整備 = NTT.整備
                .納入区分 = NTT.納入区分
                .納品書発行枚数 = NTT.納品書発行枚数
                .かんばん回転枚数 = NTT.かんばん回転枚数
                .かんばん増減符号 = NTT.かんばん増減符号
                .かんばん増減枚数 = NTT.かんばん増減枚数
                .変更数符号 = NTT.変更数符号
                .変更数 = NTT.変更数
                .調整数符号 = NTT.調整数符号
                .調整数 = NTT.調整数
                .処理区分 = 0
                .変更回数 = NTT.変更回数

                .登録FCODE = NTT.FILECODE
                .登録H作成日 = NTT.HEAD作成日
                .登録HSEQ = NTT.HEADSEQ
                .登録連番 = NTT.連番

           End With

           If bNJPutFlg Then
              Call NJTInsert
           Else
              Call NJTUpdate
           End If
        End If
    End With
    On Error GoTo 0
    NaijiConVert = True
    Exit Function
NaijiConVert_Er:
    Static ER_MSG1 As String
    Static ER_MSG2 As String
    
    ER_MSG1 = "         納入内示データ変換処理      " & vbCrLf & " " & vbCrLf
    ER_MSG1 = ER_MSG1 & "【 " & Err.Description & " 】"

    Call ProcLogPut(ER_MSG1)

    ER_MSG2 = Chr(13) & Space(3) & _
               "上記エラーにより、以降の処理を続行出来ません" & _
               Chr(10) & Chr(13) & Space(8) & _
               "内容を確認して再度実行して下さい"
    lErr = "N1"
    Call TERPut("N1", ER_MSG1, ER_MSG2)
    
End Function
'+--------------------------------------- 2004/12/13 Insert
'+--------------------------------------------+
'+      取引先・受入より取引先工場を求める
'+--------------------------------------------+
Private Function NonyuGet(Torcd As String, Ukeir As String) As String
    NonyuGet = Torcd
    
    gSL_Select = "Select 取引先CD From 取引先受入マスタ "
    gSL_Select = gSL_Select & " Where 取引先CD Like '" & RTrim(Torcd) & "%'"
    gSL_Select = gSL_Select & " And   取引先区分 = 0"
    gSL_Select = gSL_Select & " And   受入 = '" & RTrim(Ukeir) & "'"
    gSL_Select = gSL_Select & " Order By 取引先CD"
    
    If Not TUMRead(gSL_Select, 1) Then Exit Function
    
    NonyuGet = TUM.取引先CD
End Function
'+--------------------------------------- 2004/12/13 Insert End


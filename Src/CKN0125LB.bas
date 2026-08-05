Attribute VB_Name = "CKN0125LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   発注テーブル一覧表　発行
'*
'*   モジュールID  : CKN0125LB
'*
'*   作成日  :2004/05/11   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCKN0125        As New CISFormPrint
'印刷用ワーク
Public CKN0125_PrintOut     As Integer
Public CKN0125_MaxKaisu     As Integer
'
Public Function CKN0125LPrint(Optional PrintOut As Integer = 0) As Boolean
    CKN0125LPrint = False

' 出力データ存在確認
    gSL_Select = " SELECT HT.取引先 仕入先,HT.受入, "
    gSL_Select = gSL_Select & " SUBSTRING(HT.サイクル,1,1) + '-' + SUBSTRING(HT.サイクル,2,2) + '-' + SUBSTRING(HT.サイクル,4,2) 編集サイクル,"
    gSL_Select = gSL_Select & " HT.サイクル,ISNULL(TR.略称,'')仕入先名,"
    gSL_Select = gSL_Select & " ISNULL(HC.最終発注日,'')最終発注日,ISNULL(HC.最終発注便,'')最終発注便,ISNULL(HC.振出便,'')振出便,"
    gSL_Select = gSL_Select & " ISNULL(HC.発注処理回数,0)発注処理回数,ISNULL(HC.発行加算日,0)発行加算日,ISNULL(TU.曜日発注,0) 曜日発注"
    gSL_Select = gSL_Select & " FROM 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 発注サイクルマスタ HC"
    gSL_Select = gSL_Select & "   ON  HC.仕入先 = HT.取引先"
    gSL_Select = gSL_Select & "   AND HC.サイクル = HT.サイクル"
    gSL_Select = gSL_Select & "   AND HC.受入 = HT.受入"
    gSL_Select = gSL_Select & "   AND HC.行識別 = 0"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON  TR.取引先CD = HT.取引先"
    gSL_Select = gSL_Select & "   AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先受入マスタ TU"
    gSL_Select = gSL_Select & "   ON  TU.取引先CD = HT.取引先"
    gSL_Select = gSL_Select & "   AND TU.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TU.受入 = HT.受入"
    gSL_Select = gSL_Select & " WHERE HT.品目 = 1"
    gSL_Select = gSL_Select & "   AND TR.社内区分 = 0"
    gSL_Select = gSL_Select & "   AND HT.サイクル <> '' AND HT.サイクル <> '00000'"
    gSL_Select = gSL_Select & " GROUP BY HT.取引先,HT.受入,HT.サイクル,TR.略称,HC.振出便,HC.発注処理回数,HC.最終発注日,HC.最終発注便,HC.発行加算日,TU.曜日発注"
    gSL_Select = gSL_Select & " ORDER BY HT.取引先,HT.受入,HT.サイクル,TR.略称,HC.振出便,HC.発注処理回数,HC.最終発注日,HC.最終発注便,HC.発行加算日"
    If Not HCMRead(gSL_Select, 0, 1) Then
       Call HCMClose
       CKN0125LPrint = True
       Exit Function
    End If

    With CisPrtCKN0125
'       件数退避
        .cDataMaxCnt = CisDB.RecordCount
'
        If CKN0125_MaxKaisu > 8 Then
            Set .iPrtForm = CKN0126L
        Else
            Set .iPrtForm = CKN0125L
        End If
        .dpMSG = "発注テーブル一覧表　印刷中"
        If PrintOut = 0 Then
            .iPreview = pv_Print
        Else
            .iPreview = pv_Preview
        End If
        .dpMSGScreen = モードレス
        .PrintStart
    End With

    CKN0125LPrint = True
End Function


Attribute VB_Name = "CKH0910LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   かんばん読取エラーリスト
'*
'*   モジュールID  : CKH0910LB
'*
'*                  CKH0910LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not CKH0910LPrint Then Exit Sub
'
'*   作成日  :2004/01/06   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCKH0910        As New CISFormPrint
Public CKH0910_PrtCnt       As Long


Public Function CKH0910LPrint() As Boolean
    CKH0910LPrint = False
' 出力データ存在確認
    
    gSL_Select = "SELECT ISNULL(TR.略称,'') 取引先名,"
    gSL_Select = gSL_Select & " CASE WHEN Isnull(HN.表示品番,'') = '' THEN KE.品番 Else HN.表示品番 End 表示品番,"
    gSL_Select = gSL_Select & "  RTRIM(ISNULL(SY.値名称,'')) 手配区分名,"
    gSL_Select = gSL_Select & "  KE.* From 帳票出力テーブル CH"
    gSL_Select = gSL_Select & "  Left Outer Join かんばん読取エラーテーブル KE"
    gSL_Select = gSL_Select & "      ON  KE.エラー管理番号 = CH.管理NO"
    gSL_Select = gSL_Select & "      And KE.エラー種類 = CH.帳票区分"
    gSL_Select = gSL_Select & "  Left Outer Join 品番マスタ HN"
    gSL_Select = gSL_Select & "      ON  HN.品番 = KE.品番"
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "      ON  TR.取引先CD = KE.取引先"
    gSL_Select = gSL_Select & "      And TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "  Left Outer Join 名称マスタ SY"
    gSL_Select = gSL_Select & "      ON  SY.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "      And SY.区分タイプ = ''"
    gSL_Select = gSL_Select & "      And SY.値 = KE.手配区分"
    gSL_Select = gSL_Select & " Where CH.端末番号 = Host_Name()"
    gSL_Select = gSL_Select & "   And CH.帳票区分 = 1"
    
    gSL_Select = gSL_Select & " Order By"
    gSL_Select = gSL_Select & " KE.エラー要因,KE.取引先,KE.受入,KE.背番号,"
    gSL_Select = gSL_Select & " KE.品番,KE.手配区分,KE.収容数,KE.発行NO"
    
    If Not KRERead(gSL_Select, 0, 1) Then
       Call KREClose
       CKH0910LPrint = True
       Exit Function
    End If
    
    CKH0910_PrtCnt = CisDB.RecordCount
    
    With CisPrtCKH0910
         Set .iPrtForm = CKH0910L
         .dpMSG = "かんばん読取エラーリスト　印刷中"
         .iPreview = gDemoPrt
         .dpMSGScreen = 画面なし
         .PrintStart
    End With
    
    CKH0910LPrint = True
End Function




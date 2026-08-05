Attribute VB_Name = "CXJ0031LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   受注かんばん読取エラーリスト
'*
'*   モジュールID  : CXJ0031LB
'*
'*                  CXJ0031LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not CXJ0031LPrint Then Exit Sub
'
'*   作成日  :2004/01/06   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCXJ0031        As New CISFormPrint
Public CXJ0031_PrtCnt       As Long


Public Function CXJ0031LPrint() As Boolean
    CXJ0031LPrint = False
' 出力データ存在確認
    
    gSL_Select = "SELECT ISNULL(TR.略称,'') 取引先名,"
    gSL_Select = gSL_Select & " CASE WHEN Isnull(HN.表示品番,'') = '' THEN KE.品番 Else HN.表示品番 End 表示品番,"
    gSL_Select = gSL_Select & "  KE.* From 帳票出力テーブル CH"
    gSL_Select = gSL_Select & "  Left Outer Join かんばん読取エラーテーブル KE"
    gSL_Select = gSL_Select & "      ON  KE.エラー管理番号 = CH.管理NO"
    gSL_Select = gSL_Select & "      And KE.エラー種類 = CH.帳票区分"
    gSL_Select = gSL_Select & "  Left Outer Join 品番マスタ HN"
    gSL_Select = gSL_Select & "      ON  HN.品番 = KE.品番"
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "      ON  TR.取引先CD = KE.取引先"
    gSL_Select = gSL_Select & "      And TR.取引先区分 = 0"
    gSL_Select = gSL_Select & " Where CH.端末番号 = Host_Name()"
    gSL_Select = gSL_Select & "   And CH.帳票区分 = 41"
    
    gSL_Select = gSL_Select & " Order By"
    gSL_Select = gSL_Select & " KE.エラー要因,KE.取引先,KE.受入,KE.背番号,"
    gSL_Select = gSL_Select & " KE.品番,KE.収容数,KE.発行NO"
    
    If Not KRERead(gSL_Select, 0, 1) Then
       Call KREClose
       CXJ0031LPrint = True
       Exit Function
    End If
    
    CXJ0031_PrtCnt = CisDB.RecordCount
    
    With CisPrtCXJ0031
         Set .iPrtForm = CXJ0031L
         .dpMSG = "受注かんばん読取ｴﾗｰﾘｽﾄ　印刷中"
         .iPreview = gDemoPrt
         .dpMSGScreen = 画面なし
         .PrintStart
    End With
    
    CXJ0031LPrint = True
End Function




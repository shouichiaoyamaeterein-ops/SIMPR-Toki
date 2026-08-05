Attribute VB_Name = "CXH0040LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   発注打切完了済リスト
'*
'*   モジュールID  : CXH0040LB
'*
'*                  CXH0040LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not CXH0040LPrint Then Exit Sub
'
'*   作成日  :2004/03/26   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCXH0040        As New CISFormPrint
Public CXH0040_PrtCnt       As Long

Public Function CXH0040LPrint() As Boolean
    CXH0040LPrint = False
    
' 出力データ存在確認
    gSL_Select = "SELECT * FROM 帳票出力テーブル"
    gSL_Select = gSL_Select & " Where 端末番号 = Host_Name()"
    gSL_Select = gSL_Select & "   And 帳票区分 = 13"
    
    If Not CHTRead(gSL_Select, 1) Then
       CXH0040LPrint = True
       Exit Function
    End If
    
    gStr = "SELECT ISNULL(TR.略称,'') 取引先名,"
    gStr = gStr & " CASE WHEN Isnull(HN.表示品番,'') = '' THEN KE.品番 Else HN.表示品番 End 表示品番,"
    gSL_Select = gStr
    gSL_Select = gSL_Select & "KE.取引先,KE.受入,KE.背番号,KE.品番,KE.発行NO,KE.収容数,"
    gSL_Select = gSL_Select & "KE.打切種別,KE.打切区分,KE.起算日,KE.起算便,KE.完了日,KE.完了便,KE.打切日,"
    gSL_Select = gSL_Select & "KE.打切設定数,KE.今回発注数,KE.今回発注枚数,KE.今回発注端数"
    gSL_Select = gSL_Select & "  From 帳票出力テーブル CH"
    gSL_Select = gSL_Select & "  Left Outer Join かんばん読取エラーテーブル KE"
    gSL_Select = gSL_Select & "      ON  KE.エラー管理番号 = CH.管理NO"
    gSL_Select = gSL_Select & "      And KE.エラー種類 = CH.帳票区分"
    gSL_Select = gSL_Select & "  Left Outer Join 品番マスタ HN"
    gSL_Select = gSL_Select & "      ON  HN.品番 = KE.品番"
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "      ON  TR.取引先CD = KE.取引先"
    gSL_Select = gSL_Select & "      And TR.取引先区分 = 1"
    gSL_Select = gSL_Select & " Where CH.端末番号 = Host_Name()"
    gSL_Select = gSL_Select & "   And CH.帳票区分 = 13"
    
    gSL_Select = gSL_Select & " Order By KE.取引先,KE.受入,KE.品番,KE.発行NO"
    
    If Not KRERead(gSL_Select, 0, 1) Then
       Call KREClose
       CXH0040LPrint = True
       Exit Function
    End If
    
    CXH0040_PrtCnt = CisDB.RecordCount
    
    With CisPrtCXH0040
         Set .iPrtForm = CXH0040L
         .dpMSG = "発注打切完了済リスト　印刷中"
         .iPreview = gDemoPrt
         .dpMSGScreen = モーダル
         .PrintStart
    End With
    
    CXH0040LPrint = True
End Function




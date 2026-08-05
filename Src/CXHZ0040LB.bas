Attribute VB_Name = "CXHZ0040LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   発注打切完了済リスト
'*
'*   モジュールID  : CXHZ0040LB
'*
'*                  CXHZ0040LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not CXHZ0040LPrint Then Exit Sub
'
'*   作成日  :2004/03/26   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCXHZ0040        As New CISFormPrint
Public CXHZ0040_PrtCnt       As Long

Public Function CXHZ0040LPrint() As Boolean
    CXHZ0040LPrint = False
    
' 出力データ存在確認
    gSL_Select = "SELECT * FROM 帳票出力テーブル"
    gSL_Select = gSL_Select & " Where 端末番号 = Host_Name()"
    gSL_Select = gSL_Select & "   And 帳票区分 = 33"
    
    If Not CHTRead(gSL_Select, 1) Then
       CXHZ0040LPrint = True
       Exit Function
    End If
    
    gSL_Select = "SELECT ISNULL(TR.略称,'') 取引先名,"
    gSL_Select = gSL_Select & "KE.取引先,KE.受入,KE.背番号,KE.材質,HN.表示寸法,KE.発行NO,KE.収容数,KE.材料管理番号,"
    gSL_Select = gSL_Select & "KE.打切種別,KE.打切区分,KE.起算日,KE.起算便,KE.完了日,KE.完了便,KE.打切日,"
    gSL_Select = gSL_Select & "KE.打切設定数,KE.今回発注数,KE.今回発注枚数,KE.今回発注端数"
    gSL_Select = gSL_Select & "  From 帳票出力テーブル CH"
    gSL_Select = gSL_Select & "  Left Outer Join 材料かんばん読取エラーテーブル KE"
    gSL_Select = gSL_Select & "      ON  KE.エラー管理番号 = CH.管理NO"
    gSL_Select = gSL_Select & "      And KE.エラー種類 = CH.帳票区分"
    gSL_Select = gSL_Select & "  Left Outer Join 材料マスタ HN"
    gSL_Select = gSL_Select & "      ON  HN.材料管理番号 = KE.材料管理番号"
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "      ON  TR.取引先CD = KE.取引先"
    gSL_Select = gSL_Select & "      And TR.取引先区分 = 1"
    gSL_Select = gSL_Select & " Where CH.端末番号 = Host_Name()"
    gSL_Select = gSL_Select & "   And CH.帳票区分 = 33"
    
    gSL_Select = gSL_Select & " Order By KE.取引先,KE.受入,KE.材質,KE.板厚,KE.幅,KE.長さ,KE.発行NO"
    
    If Not ZRERead(gSL_Select, 0, 1) Then
       Call ZREClose
       CXHZ0040LPrint = True
       Exit Function
    End If
    
    CXHZ0040_PrtCnt = CisDB.RecordCount
    
    With CisPrtCXHZ0040
         Set .iPrtForm = CXHZ0040L
         .dpMSG = "材料発注打切完了済リスト　印刷中"
         .iPreview = gDemoPrt
         .dpMSGScreen = モーダル
         .PrintStart
    End With
    
    CXHZ0040LPrint = True
End Function




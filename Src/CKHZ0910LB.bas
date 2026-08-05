Attribute VB_Name = "CKHZ0910LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   材料かんばん読取エラーリスト
'*
'*   モジュールID  : CKHZ0910LB
'*
'*                  CKHZ0910LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not CKHZ0910LPrint Then Exit Sub
'
'*   作成日  :2004/02/21   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCKHZ0910        As New CISFormPrint
Public CKHZ0910_PrtCnt       As Long


Public Function CKHZ0910LPrint() As Boolean
    CKHZ0910LPrint = False
' 出力データ存在確認
    
    gSL_Select = "SELECT ISNULL(TR.略称,'') 取引先名,ZM.SIZE サイズ,"
    gSL_Select = gSL_Select & "  RTRIM(ISNULL(SY.値名称,'')) 手配区分名,ZM.表示寸法,"
    gSL_Select = gSL_Select & "  KE.* From 帳票出力テーブル CH"
    gSL_Select = gSL_Select & "  Left Outer Join 材料かんばん読取エラーテーブル KE"
    gSL_Select = gSL_Select & "      ON  KE.エラー管理番号 = CH.管理NO"
    gSL_Select = gSL_Select & "      And KE.エラー種類 = CH.帳票区分"
    gSL_Select = gSL_Select & "  Left Outer Join 材料マスタ ZM"
    gSL_Select = gSL_Select & "      ON  ZM.材料管理番号 = KE.材料管理番号"
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "      ON  TR.取引先CD = KE.取引先"
    gSL_Select = gSL_Select & "      And TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "  Left Outer Join 名称マスタ SY"
    gSL_Select = gSL_Select & "      ON  SY.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "      And SY.区分タイプ = ''"
    gSL_Select = gSL_Select & "      And SY.値 = KE.手配区分"
    gSL_Select = gSL_Select & " Where CH.端末番号 = Host_Name()"
    gSL_Select = gSL_Select & "   And CH.帳票区分 = 21"
    
    gSL_Select = gSL_Select & " Order By"
    gSL_Select = gSL_Select & " KE.エラー要因,KE.取引先,KE.受入,KE.背番号,"
    gSL_Select = gSL_Select & " KE.材質,ZM.SIZE,KE.手配区分,KE.収容数,KE.発行NO"
    
    If Not ZRERead(gSL_Select, 0, 1) Then
       Call ZREClose
       CKHZ0910LPrint = True
       Exit Function
    End If
    
    CKHZ0910_PrtCnt = CisDB.RecordCount
    
    With CisPrtCKHZ0910
         Set .iPrtForm = CKHZ0910L
         .dpMSG = "材料かんばん読取ｴﾗｰﾘｽﾄ　印刷中"
         .iPreview = gDemoPrt
         .dpMSGScreen = 画面なし
         .PrintStart
    End With
    
    CKHZ0910LPrint = True
End Function




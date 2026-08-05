Attribute VB_Name = "CXHZ0330LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   [材料]発注明細書発行
'*
'*   モジュールID  : CXHZ0330LB
'*
'*                  CXHZ0330LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not CXHZ0330LPrint Then Exit Sub
'
'*   作成日  :2004/02/23   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCXHZ0330        As New CISFormPrint
Public CXHZ0330_PrtCnt       As Long
Global CXHZ0330_PrintKb      As Byte

Public Function CXHZ0330LPrint(PrintKB As Byte, Optional PrintOut As Integer = 0) As Boolean
    CXHZ0330LPrint = False

    If PrintKB = 0 Then
        CXHZ0330_PrintKb = 28
    Else
        CXHZ0330_PrintKb = 128
    End If

' 出力データ存在確認
    gSL_Select = "SELECT HC.*,HN.表示寸法,HN.材種コード,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.略称,'') 仕入先名,ISNULL(TR1.工場名,'') 仕入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.取引先,'') 仕入先CD,ISNULL(TR1.工場CD,'') 仕入先工場CD,"
    gSL_Select = gSL_Select & "  ISNULL(TR2.略称,'') 納入先名,ISNULL(TR2.工場名,'') 納入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR2.取引先,'') 納入先CD,ISNULL(TR2.工場CD,'') 納入先工場CD,"
    gSL_Select = gSL_Select & "  ISNULL(SY1.値名称,'') 発注区分名,ISNULL(SY2.値名称,'') 手配区分名,"
    gSL_Select = gSL_Select & "  ISNULL(SY3.値名称,'') 手配単位名,ISNULL(SY4.値名称,'') 材種名"
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料発注テーブル HC"
    gSL_Select = gSL_Select & "   ON HC.明細書NO = CH.管理NO"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料マスタ HN"
    gSL_Select = gSL_Select & "   ON HN.材料管理番号 = HC.材料管理番号"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料取引先マスタ HT"
    gSL_Select = gSL_Select & "   ON HT.材料管理番号 = HC.材料管理番号"
    gSL_Select = gSL_Select & "   AND HT.取引先 = HC.仕入先"
    gSL_Select = gSL_Select & "   AND HT.受入 = HC.受入"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR1"
    gSL_Select = gSL_Select & "   ON TR1.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR1.取引先CD = HC.仕入先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR2"
    gSL_Select = gSL_Select & "   ON TR2.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR2.取引先CD = HC.納入先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY1"
    gSL_Select = gSL_Select & "   ON SY1.区分名称 = '発注区分'"
    gSL_Select = gSL_Select & "   AND SY1.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY1.値 = HC.発注区分"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2"
    gSL_Select = gSL_Select & "   ON SY2.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "   AND SY2.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY2.値 = HC.手配区分"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY3"
    gSL_Select = gSL_Select & "   ON SY3.区分名称 = '単位区分'"
    gSL_Select = gSL_Select & "   AND SY3.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY3.値 = HN.手配単位"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY4"
    gSL_Select = gSL_Select & "   ON SY4.区分名称 = '材種区分'"
    gSL_Select = gSL_Select & "   AND SY4.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY4.値 = HN.材種"
    gSL_Select = gSL_Select & " WHERE CH.端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND CH.帳票区分 = " & CXHZ0330_PrintKb
    gSL_Select = gSL_Select & "   AND CH.管理文字 = '1'"
'    gSL_Select = gSL_Select & "   AND HC.明細書NO = 15803"
    
    gSL_Select = gSL_Select & " Order By HC.明細書NO"
    
    If Not ZCTRead(gSL_Select, 0, 1) Then
       Call ZCTClose
       CXHZ0330LPrint = True
       Exit Function
    End If
    
    CXHZ0330_PrtCnt = CisDB.RecordCount
    
    With CisPrtCXHZ0330
         Set .iPrtForm = CXHZ0330L
         .dpMSG = "[材料]発注明細書　印刷中"
         If PrintOut = 0 Then
           .iPreview = pv_Print
         Else
           .iPreview = pv_Preview
         End If
         .dpMSGScreen = モードレス
         .PrintStart
    End With
    
    CXHZ0330LPrint = True
End Function




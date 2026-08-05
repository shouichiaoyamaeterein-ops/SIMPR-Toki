Attribute VB_Name = "CXHZ0340LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   [材料]臨時かんばん発行
'*
'*   モジュールID  : CXHZ0340LB
'*
'*                  CXHZ0340LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not CXHZ0340LPrint Then Exit Sub
'
'*   作成日  :2004/01/30   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCXHZ0340        As New CISFormPrint
Public CXHZ0340_PrtCnt       As Long


Public Function CXHZ0340LPrint(PrintKB As Byte, Optional PrintOut As Integer = 0) As Boolean
    CXHZ0340LPrint = False

    If PrintKB = 0 Then
        PrintKB = 29
    Else
        PrintKB = 129
    End If

' 出力データ存在確認
    gSL_Select = "SELECT HC.*,HN.表示寸法,HN.材料名,HN.材種コード,HT.置場,"
    gSL_Select = gSL_Select & "  CASE WHEN HT.変更日 <> '' AND HT.変更日 + HT.変更便 <= HC.納入日 + HC.納入便"
    gSL_Select = gSL_Select & "  THEN 収容器2 ELSE 収容器1 END 収容器,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.略称,'') 仕入先名,ISNULL(TR1.工場名,'') 仕入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.取引先,'') 仕入先CD,ISNULL(TR1.工場CD,'') 仕入先工場CD,"
    gSL_Select = gSL_Select & "  ISNULL(TR2.略称,'') 納入先名,ISNULL(TR2.工場名,'') 納入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR2.取引先,'') 納入先CD,ISNULL(TR2.工場CD,'') 納入先工場CD,"
    gSL_Select = gSL_Select & "  ISNULL(SY1.値名称,'') 発注区分名,ISNULL(SY2.値名称,'') 手配区分名,"
    gSL_Select = gSL_Select & "  ISNULL(SY3.値名称,'') 手配単位名,ISNULL(SY4.値名称,'') 材種名"
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料発注テーブル HC"
    gSL_Select = gSL_Select & "   ON HC.発注管理NO = CH.管理NO"
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
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY3"
    gSL_Select = gSL_Select & "   ON SY3.区分名称 = '単位区分'"
    gSL_Select = gSL_Select & "   AND SY3.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY3.値 = HN.手配単位"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2"
    gSL_Select = gSL_Select & "   ON SY2.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "   AND SY2.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY2.値 = HC.手配区分"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY4"
    gSL_Select = gSL_Select & "   ON SY4.区分名称 = '材種区分'"
    gSL_Select = gSL_Select & "   AND SY4.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY4.値 = HN.材種"
    gSL_Select = gSL_Select & " WHERE 端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND 帳票区分 = " & PrintKB
    
    gSL_Select = gSL_Select & " Order By"
    gSL_Select = gSL_Select & " HC.仕入先,HC.受入,HC.納入番号,HC.納品書頁,HC.納品書行"
    
    If Not ZCTRead(gSL_Select, 0, 1) Then
       Call ZCTClose
       CXHZ0340LPrint = True
       Exit Function
    End If
   '枚数獲得
    gSL_Select = "SELECT SUM(CASE WHEN HC.端数 <> 0 AND HC.端数区分 <> 0 THEN HC.枚数 - 1 ELSE HC.枚数 END) 件数 "
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料発注テーブル HC"
    gSL_Select = gSL_Select & "   ON HC.発注管理NO = CH.管理NO"
    gSL_Select = gSL_Select & " WHERE 端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND 帳票区分 = " & PrintKB
    With CisDB
        .SQL = gSL_Select
        .ReadStatus = True
        If .DBRead(1, 0) Then
            CXHZ0340_PrtCnt = CisDB.Recordset(0)
        End If
    End With
    
    With CisPrtCXHZ0340
         Set .iPrtForm = CXHZ0340L
         .dpMSG = "臨時かんばん　印刷中"
         If PrintOut = 0 Then
           .iPreview = pv_Print
         Else
           .iPreview = pv_Preview
         End If
         .dpMSGScreen = モードレス
         .PrintStart
    End With
    
    CXHZ0340LPrint = True
End Function




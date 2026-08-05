Attribute VB_Name = "CXH0360LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   端数ラベル発行
'*
'*   モジュールID  : CXH0360LB
'*
'*                  CXH0360LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not CXH0360LPrint Then Exit Sub
'
'*   作成日  :2004/01/30   By CIS
'*   変更日  :2007/10/30   By CIS   帳票発行後にFax処理を行うと正常に動作しない点を修正
'*---------------------------------------------------------------------------------
Public CisPrtCXH0360        As New CISFormPrint
Public CXH0360_PrtCnt       As Long


Public Function CXH0360LPrint(PrintKB As Byte, Optional PrintOut As Integer = 0) As Boolean
    CXH0360LPrint = False

    If PrintKB = 0 Then
        PrintKB = 11
    Else
        PrintKB = 111
    End If

' 出力データ存在確認
    gSL_Select = "SELECT HC.*,HN.表示品番,HN.車型,HN.品名,HT.置場,"
    gSL_Select = gSL_Select & "  CASE WHEN HT.変更日 <> '' AND HT.変更日 + HT.変更便 <= HC.納入日 + HC.納入便"
    gSL_Select = gSL_Select & "  THEN 収容器2 ELSE 収容器1 END 収容器,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.略称,'') 仕入先名,ISNULL(TR1.工場名,'') 仕入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.取引先,'') 仕入先CD,ISNULL(TR1.工場CD,'') 仕入先工場CD,"
    gSL_Select = gSL_Select & "  ISNULL(TR2.略称,'') 納入先名,ISNULL(TR2.工場名,'') 納入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR2.取引先,'') 納入先CD,ISNULL(TR2.工場CD,'') 納入先工場CD,"
    gSL_Select = gSL_Select & "  ISNULL(SY1.値名称,'') 発注区分名,ISNULL(SY2.値名称,'') 手配区分名"
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 発注テーブル HC"
    gSL_Select = gSL_Select & "   ON HC.発注管理NO = CH.管理NO"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "   ON HN.品番 = HC.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT"
    gSL_Select = gSL_Select & "   ON HT.品番 = HC.品番"
    gSL_Select = gSL_Select & "   AND HT.取引先 = HC.仕入先"
    gSL_Select = gSL_Select & "   AND HT.受入 = HC.受入"
    gSL_Select = gSL_Select & "   AND HT.品目 = 1"
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
    gSL_Select = gSL_Select & " WHERE 端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND 帳票区分 = " & PrintKB
    
    gSL_Select = gSL_Select & " Order By"
    gSL_Select = gSL_Select & " HC.仕入先,HC.受入,HC.納入番号,HC.納品書頁,HC.納品書行"
    
    If Not HCTRead(gSL_Select, 0, 1) Then
       Call HCTClose
       CXH0360LPrint = True
       Exit Function
    End If
    '枚数退避
    CXH0360_PrtCnt = CisDB.RecordCount
    
    With CisPrtCXH0360
         Set .iPrtForm = CXH0360L
         .dpMSG = "端数ラベル　印刷中"
         If PrintOut = 0 Then
           .iPreview = pv_Print
         Else
           .iPreview = pv_Preview
         End If
         .dpMSGScreen = モードレス
         .PrintStart
    End With
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2007/10/30 INSERT START
    Set CisPrtCXH0360 = Nothing
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2007/10/30 INSERT END

    CXH0360LPrint = True
End Function




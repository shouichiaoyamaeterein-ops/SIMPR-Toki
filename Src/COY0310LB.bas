Attribute VB_Name = "COY0310LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   売上予想表　印刷
'*
'*   モジュールID  : COY0310.Bas
'*
'*   戻値 : なし
'*
'*   作成日  :2004/11/05   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrt           As New CISFormPrint

Global UY_NJYm          As String
Global UY_Torcd         As String
Global UY_Kouku         As String
Global UY_HinKeta       As Byte

Public Function SQLEdit0310()
    
'    gSL_Select = "SELECT 内示年月,UY.工区,UY.契約先,ISNULL(TR.略称,'') 契約先名,UY.車型,"
    gSL_Select = "SELECT 内示年月,UY.工区,UY.契約先,ISNULL(TR.略称,'') 契約先名,"
    gSL_Select = gSL_Select & " SUM(数量1) 数量1,SUM(金額1) 金額1,SUM(前月比1) 前月比1,"
    gSL_Select = gSL_Select & " SUM(数量2) 数量2,SUM(金額2) 金額2,SUM(前月比2) 前月比2,"
    gSL_Select = gSL_Select & " SUM(数量3) 数量3,SUM(金額3) 金額3"
    gSL_Select = gSL_Select & "  FROM 売上予想テーブル UY"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "     ON TR.取引先CD = UY.契約先"
    gSL_Select = gSL_Select & "     AND TR.取引先区分 = 0"
    gSL_Select = gSL_Select & " WHERE UY.内示年月 = '" & RTrim(UY_NJYm) & "'"
    If RTrim(UY_Torcd) <> "" Then
        gSL_Select = gSL_Select & "   AND UY.契約先 = '" & RTrim(UY_Torcd) & "'"
    End If
    If RTrim(UY_Kouku) <> "" Then
        gSL_Select = gSL_Select & "   AND UY.工区 = '" & RTrim(UY_Kouku) & "'"
    End If
'    gSL_Select = gSL_Select & " GROUP BY UY.内示年月,UY.工区,UY.契約先,UY.車型,TR.略称"
'    gSL_Select = gSL_Select & " ORDER BY UY.内示年月,UY.工区,UY.契約先,UY.車型,TR.略称"
    gSL_Select = gSL_Select & " GROUP BY UY.内示年月,UY.工区,UY.契約先,TR.略称"
    gSL_Select = gSL_Select & " ORDER BY UY.内示年月,UY.工区,UY.契約先,TR.略称"
    
End Function
Public Function SQLEdit0311(SelKbn As Byte)
    
    If SelKbn = 0 Then
'        gSL_Select = "SELECT 内示年月,UY.工区,UY.契約先,ISNULL(TR.略称,'') 契約先名,UY.車型,UY.品番,UY.品番編集,"
        gSL_Select = "SELECT 内示年月,UY.工区,UY.契約先,ISNULL(TR.略称,'') 契約先名,HT.背番号 車型,UY.品番,UY.品番編集,"
        gSL_Select = gSL_Select & " 単価1,仮区分1,数量1,金額1,前月比1,"
        gSL_Select = gSL_Select & " 単価2,仮区分2,数量2,金額2,前月比2,単価3,仮区分3,数量3,金額3"
        gSL_Select = gSL_Select & "  FROM 売上予想テーブル UY"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
        gSL_Select = gSL_Select & "     ON TR.取引先CD = UY.契約先"
        gSL_Select = gSL_Select & "     AND TR.取引先区分 = 0"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN ( SELECT 品番,契約先,MIN(背番号) 背番号 FROM 品番取引先マスタ"
        gSL_Select = gSL_Select & "                                 GROUP BY 品番,契約先 ) HT        "
        gSL_Select = gSL_Select & " ON   HT.品番 = UY.品番 "
        gSL_Select = gSL_Select & " AND  HT.契約先 = UY.契約先"
        gSL_Select = gSL_Select & " WHERE UY.内示年月 = '" & RTrim(UY_NJYm) & "'"
        If RTrim(UY_Torcd) <> "" Then
            gSL_Select = gSL_Select & "   AND UY.契約先 = '" & RTrim(UY_Torcd) & "'"
        End If
        If RTrim(UY_Kouku) <> "" Then
            gSL_Select = gSL_Select & "   AND UY.工区 = '" & RTrim(UY_Kouku) & "'"
        End If
        gSL_Select = gSL_Select & " ORDER BY UY.内示年月,UY.工区,UY.契約先,UY.車型,UY.品番"
    Else
'        gSL_Select = "SELECT 内示年月,UY.工区,UY.契約先,ISNULL(TR.略称,'') 契約先名,UY.車型,UY.集約品番 品番,UY.品番編集,"
        gSL_Select = "SELECT 内示年月,UY.工区,UY.契約先,ISNULL(TR.略称,'') 契約先名,UY.集約品番 品番,UY.品番編集,"
        gSL_Select = gSL_Select & " 単価1,仮区分1,SUM(数量1)数量1,SUM(金額1)金額1,SUM(前月比1)前月比1,"
        gSL_Select = gSL_Select & " 単価2,仮区分2,SUM(数量2)数量2,SUM(金額2)金額2,SUM(前月比2)前月比2,"
        gSL_Select = gSL_Select & " 単価3,仮区分3,SUM(数量3)数量3,SUM(金額3)金額3"
        gSL_Select = gSL_Select & "  FROM 売上予想テーブル UY"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
        gSL_Select = gSL_Select & "     ON TR.取引先CD = UY.契約先"
        gSL_Select = gSL_Select & "     AND TR.取引先区分 = 0"
        gSL_Select = gSL_Select & " WHERE UY.内示年月 = '" & RTrim(UY_NJYm) & "'"
        If RTrim(UY_Torcd) <> "" Then
            gSL_Select = gSL_Select & "   AND UY.契約先 = '" & RTrim(UY_Torcd) & "'"
        End If
        If RTrim(UY_Kouku) <> "" Then
            gSL_Select = gSL_Select & "   AND UY.工区 = '" & RTrim(UY_Kouku) & "'"
        End If
'        gSL_Select = gSL_Select & " GROUP BY UY.内示年月,UY.工区,UY.契約先,UY.車型,UY.集約品番,UY.品番編集,"
        gSL_Select = gSL_Select & " GROUP BY UY.内示年月,UY.工区,UY.契約先,UY.集約品番,UY.品番編集,"
        gSL_Select = gSL_Select & " 単価1,仮区分1,単価2,仮区分2,単価3,仮区分3,TR.略称"
'        gSL_Select = gSL_Select & " ORDER BY UY.内示年月,UY.工区,UY.契約先,UY.車型,UY.集約品番"
        gSL_Select = gSL_Select & " ORDER BY UY.内示年月,UY.工区,UY.契約先,UY.集約品番"
    End If
End Function
Public Function SQLEdit0312()
    
'    gSL_Select = "SELECT 開始年月,UY.予想年月,UY.契約先,ISNULL(TR.略称,'') 契約先名,UY.車型,"
    gSL_Select = "SELECT 開始年月,UY.予想年月,UY.契約先,ISNULL(TR.略称,'') 契約先名,"
    gSL_Select = gSL_Select & " SUM(金額1)金額1,SUM(金額2)金額2,SUM(金額3)金額3,SUM(金額3)金額4,"
    gSL_Select = gSL_Select & " SUM(金額5)金額5,SUM(金額6)金額6,SUM(金額7)金額7,SUM(金額8)金額8,"
    gSL_Select = gSL_Select & " SUM(金額9)金額9,SUM(金額10)金額10,SUM(金額11)金額11,SUM(金額12)金額12,"
    gSL_Select = gSL_Select & " SUM(金額1+金額2+金額3+金額4+金額5+金額6) 上期計,"
    gSL_Select = gSL_Select & " SUM(金額7+金額8+金額9+金額10+金額11+金額12) 下期計,"
    gSL_Select = gSL_Select & " SUM(金額1+金額2+金額3+金額4+金額5+金額6+金額7+金額8+金額9+金額10+金額11+金額12) 期合計"
    gSL_Select = gSL_Select & "  FROM 売上予想型費テーブル UY"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "     ON TR.取引先CD = UY.契約先"
    gSL_Select = gSL_Select & "     AND TR.取引先区分 = 0"
    gSL_Select = gSL_Select & " WHERE UY.予想年月 = '" & RTrim(UY_NJYm) & "'"
    If RTrim(UY_Torcd) <> "" Then
        gSL_Select = gSL_Select & "   AND UY.契約先 = '" & RTrim(UY_Torcd) & "'"
    End If
'    gSL_Select = gSL_Select & " GROUP BY UY.予想年月,UY.契約先,UY.車型,UY.開始年月,TR.略称"
'    gSL_Select = gSL_Select & " ORDER BY UY.予想年月,UY.契約先,UY.車型"
    gSL_Select = gSL_Select & " GROUP BY UY.予想年月,UY.契約先,UY.開始年月,TR.略称"
    gSL_Select = gSL_Select & " ORDER BY UY.予想年月,UY.契約先"
End Function
Public Function SQLEdit0313(SelKbn As Byte)
    
    If SelKbn = 0 Then
'        gSL_Select = "SELECT 開始年月,UY.予想年月,UY.契約先,ISNULL(TR.略称,'') 契約先名,UY.車型,UY.品番,UY.品番編集,"
        gSL_Select = "SELECT 開始年月,UY.予想年月,UY.契約先,ISNULL(TR.略称,'') 契約先名,UY.品番,UY.品番編集,"
        gSL_Select = gSL_Select & " 金額1,金額2,金額3,金額4,金額5,金額6,金額7,金額8,金額9,金額10,金額11,金額12,"
        gSL_Select = gSL_Select & " 金額1+金額2+金額3+金額4+金額5+金額6 上期計,"
        gSL_Select = gSL_Select & " 金額7+金額8+金額9+金額10+金額11+金額12 下期計,"
        gSL_Select = gSL_Select & " 金額1+金額2+金額3+金額4+金額5+金額6+金額7+金額8+金額9+金額10+金額11+金額12 期合計"
        gSL_Select = gSL_Select & "  FROM 売上予想型費テーブル UY"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
        gSL_Select = gSL_Select & "     ON TR.取引先CD = UY.契約先"
        gSL_Select = gSL_Select & "     AND TR.取引先区分 = 0"
        gSL_Select = gSL_Select & " WHERE UY.予想年月 = '" & RTrim(UY_NJYm) & "'"
        If RTrim(UY_Torcd) <> "" Then
            gSL_Select = gSL_Select & "   AND UY.契約先 = '" & RTrim(UY_Torcd) & "'"
        End If
'        If RTrim(UY_Kouku) <> "" Then
'            gSL_Select = gSL_Select & "   AND UY.工区 = '" & RTrim(UY_Kouku) & "'"
'        End If
        gSL_Select = gSL_Select & " ORDER BY UY.予想年月,UY.契約先,UY.車型,UY.品番"
    Else
'        gSL_Select = "SELECT 開始年月,UY.予想年月,UY.契約先,ISNULL(TR.略称,'') 契約先名,UY.車型,UY.集約品番 品番,UY.品番編集,"
        gSL_Select = "SELECT 開始年月,UY.予想年月,UY.契約先,ISNULL(TR.略称,'') 契約先名,UY.集約品番 品番,UY.品番編集,"
        gSL_Select = gSL_Select & " SUM(金額1)金額1,SUM(金額2)金額2,SUM(金額3)金額3,SUM(金額3)金額4,"
        gSL_Select = gSL_Select & " SUM(金額5)金額5,SUM(金額6)金額6,SUM(金額7)金額7,SUM(金額8)金額8,"
        gSL_Select = gSL_Select & " SUM(金額9)金額9,SUM(金額10)金額10,SUM(金額11)金額11,SUM(金額12)金額12,"
        gSL_Select = gSL_Select & " SUM(金額1+金額2+金額3+金額4+金額5+金額6) 上期計,"
        gSL_Select = gSL_Select & " SUM(金額7+金額8+金額9+金額10+金額11+金額12) 下期計,"
        gSL_Select = gSL_Select & " SUM(金額1+金額2+金額3+金額4+金額5+金額6+金額7+金額8+金額9+金額10+金額11+金額12) 期合計"
        gSL_Select = gSL_Select & "  FROM 売上予想型費テーブル UY"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
        gSL_Select = gSL_Select & "     ON TR.取引先CD = UY.契約先"
        gSL_Select = gSL_Select & "     AND TR.取引先区分 = 0"
        gSL_Select = gSL_Select & " WHERE UY.予想年月 = '" & RTrim(UY_NJYm) & "'"
        If RTrim(UY_Torcd) <> "" Then
            gSL_Select = gSL_Select & "   AND UY.契約先 = '" & RTrim(UY_Torcd) & "'"
        End If
'        If RTrim(UY_Kouku) <> "" Then
'            gSL_Select = gSL_Select & "   AND UY.工区 = '" & RTrim(UY_Kouku) & "'"
'        End If
'        gSL_Select = gSL_Select & " GROUP BY UY.予想年月,UY.契約先,UY.車型,UY.集約品番,UY.開始年月,UY.品番編集,TR.略称"
'        gSL_Select = gSL_Select & " ORDER BY UY.予想年月,UY.契約先,UY.車型,UY.集約品番"
        gSL_Select = gSL_Select & " GROUP BY UY.予想年月,UY.契約先,UY.集約品番,UY.開始年月,UY.品番編集,TR.略称"
        gSL_Select = gSL_Select & " ORDER BY UY.予想年月,UY.契約先,UY.集約品番"
    End If
End Function


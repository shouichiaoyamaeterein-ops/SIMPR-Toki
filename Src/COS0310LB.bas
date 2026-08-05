Attribute VB_Name = "COS0310LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   買掛金明細リスト　印刷
'*
'*   モジュールID  : COS0310.Bas
'*
'*   戻値 : なし
'*
'*   作成日  :2004/11/05   By CIS
'*   変更日  :2005/03/31   By CIS 印刷項目追加(複数) 重量
'*                                印刷順変更(単一) 検収日/品番 ⇒ 品番/検収日
'*---------------------------------------------------------------------------------
Public CisPrt           As New CISFormPrint

Global S310_Ym          As String
Global S310_Torcd       As String
Global S310_FCD         As String
Global S310_SCD         As String
Global S310_PrtSyu      As Byte
Global S310_GKbn        As String
Global S310_TotalCnt    As Long
Global S310_SelCnt      As Long
Global S310_TtlKin      As Currency
Global S310_TtlSu       As Long

Global PrtSelT()  As String
Global PrtSelF()  As String
Global PrtTorcd()  As String

'*-------------------------*
'*      印刷データ読込（複数）
'*-------------------------*
Public Function DataReadF()
    DataReadF = False
    
    gSL_Select = "SELECT 対象年月,FCD,識別子,契約先,品番,受入,単価,SUBSTRING(検収日,1,6) 検収日,"
    gSL_Select = gSL_Select & " 管理番号,訂正理由,カード,SUM(数量)数量,SUM(金額)金額,"
    gSL_Select = gSL_Select & " ISNULL(SY1.値名称,'') 契約先名,ISNULL(SY2.値名称,'') 識別子名"
    '+----------------------------------------- 2005/03/31 Insert
    gSL_Select = gSL_Select & " ,重量 "
    '+----------------------------------------- 2005/03/31 Insert End
    gSL_Select = gSL_Select & " FROM TEXT買掛金テーブル TW"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY1"
    gSL_Select = gSL_Select & "   ON  SY1.区分名称 = '照合対象契約先'"
    gSL_Select = gSL_Select & "   AND SY1.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY1.値 = 契約先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2"
    gSL_Select = gSL_Select & "   ON  SY2.区分名称 = 'TEXT買掛金'"
    gSL_Select = gSL_Select & "   AND SY2.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY2.値 = 契約先 + FCD + 識別子"
    gSL_Select = gSL_Select & PrtSelF(S310_SelCnt)
    gSL_Select = gSL_Select & " GROUP BY 対象年月,FCD,識別子,契約先,品番,受入,SUBSTRING(検収日,1,6),単価,"
    gSL_Select = gSL_Select & "          管理番号,訂正理由,カード,SY1.値名称,SY2.値名称"
    '+----------------------------------------- 2005/03/31 Insert
    gSL_Select = gSL_Select & "         ,重量 "
    '+----------------------------------------- 2005/03/31 Insert End
    gSL_Select = gSL_Select & " ORDER BY 対象年月,FCD,識別子,契約先,品番,受入,SUBSTRING(検収日,1,6),単価,"
    gSL_Select = gSL_Select & "          管理番号,訂正理由,カード"
    
    If Not TKKRead(gSL_Select, 0) Then
        Call TKKClose
        Exit Function
    End If
    CisPrt.cDataMaxCnt = CisDB.RecordCount
    S310_TotalCnt = S310_TotalCnt + CisDB.RecordCount
    
    DataReadF = True
End Function
'*-------------------------*
'*      印刷データ読込（単一）
'*-------------------------*
Public Function DataReadT()
    DataReadT = False
    
    gSL_Select = "SELECT TW.*,"
    gSL_Select = gSL_Select & " ISNULL(SY1.値名称,'') 契約先名,ISNULL(SY2.値名称,'') 識別子名"
    gSL_Select = gSL_Select & " FROM TEXT買掛金テーブル TW"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY1"
    gSL_Select = gSL_Select & "   ON  SY1.区分名称 = '照合対象契約先'"
    gSL_Select = gSL_Select & "   AND SY1.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY1.値 = 契約先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2"
    gSL_Select = gSL_Select & "   ON  SY2.区分名称 = 'TEXT買掛金'"
    gSL_Select = gSL_Select & "   AND SY2.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY2.値 = 契約先 + FCD + 識別子"
    gSL_Select = gSL_Select & PrtSelT(S310_SelCnt)
    '+---------------------------------------------- 2005/03/31 Delete
'    gSL_Select = gSL_Select & " ORDER BY 対象年月,契約先,FCD,識別子,検収日,品番,納入番号,受入"
    '+---------------------------------------------- 2005/03/31 Delete End
    '+---------------------------------------------- 2005/03/31 Insert
    gSL_Select = gSL_Select & " ORDER BY 対象年月,契約先,FCD,識別子,品番,検収日,納入番号,受入"
    '+---------------------------------------------- 2005/03/31 Insert End
    
    If Not TKKRead(gSL_Select, 0) Then
        Call TKKClose
        Exit Function
    End If
    S310_TotalCnt = S310_TotalCnt + CisDB.RecordCount
    
    DataReadT = True
End Function


Attribute VB_Name = "CSJ0100LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   プレス製造指示書　発行
'*
'*   モジュールID  : CSJ0100LB
'*
'*   作成日  :2010/01/27   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCSJ0100        As New CISFormPrint
Global gSYMD                As String

'---------------------------------------------------'
'       親品番抽出                                  '
'---------------------------------------------------'
Public Function GetSD(HinBN As String) As Boolean
    With CisDB
        .SQL = "DELETE FROM 構成問合せワーク WHERE 処理端末 = HOST_NAME()"
        .DBExec

        .SQL = "構成問合せ情報抽出"
        .StoadoCount = 4
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = Input用
        .ParaNo = 3: .ParaIO = Input用
        .ParaNo = 4: .ParaIO = Input用

        .ParaNo = 1: .ParaValue = Trim(HinBN)
        .ParaNo = 2: .ParaValue = 0     ' 品番／材料　：品番
        .ParaNo = 3: .ParaValue = 1     ' 正　／逆    ：逆
        .ParaNo = 4: .ParaValue = 1     ' 直前／製品　：製品

        If Not .DBStored Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     構成マスタ情報の抽出に失敗しました。     "
                .MB_Title = "抽出処理"
                .MB_Button = Error
                If .MBOX Then Exit Function
            End With
        End If
    End With
End Function
'****************************************
'*          材料在庫量算出              *
'****************************************   ' Based PXZ0020
Public Function GetZaikoSu(ZaiNo As Long) As Boolean
    GetZaikoSu = False
'    Dim wKanriNo    As Integer: wKanriNo = 0                       '2010/08/27
    Dim wKanriNo    As Long: wKanriNo = 0                           '2010/08/27
    Dim wZaikoSu    As Long:    wZaikoSu = 0
    Dim wMasterSu   As Long:    wMasterSu = 0
    Dim YM0         As String:  YM0 = Format(Now(), "yyyymm")

    If ZaiNo = 0 Then
        HKT.材料在庫量 = 0
        GetZaikoSu = True
        Exit Function
    End If
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT * FROM 材料在庫マスタ"
    gSL_Select = gSL_Select & " WHERE 材料管理番号 = " & CStr(ZaiNo)
    gSL_Select = gSL_Select & "   AND 年月度 = '" & RTrim(YM0) & "'"
    
    If Not ZZMRead(gSL_Select) Then
        wMasterSu = 0
        '---在庫マスタ：無
        gSL_Select = ""
        gSL_Select = gSL_Select & ""
        gSL_Select = gSL_Select & "SELECT SUBSTRING(SN.入荷日,1,6) AS 入荷日, SN.材質, SN.板厚, SN.幅, SN.長さ"
        gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) + ISNULL(JK1.実績数,0) - ISNULL(ZS1.使用量,0) - ISNULL(JK2.実績数,0) AS 入荷数"
        gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) AS 入荷数入, ISNULL(JK1.実績数,0) AS 実績数入"
        gSL_Select = gSL_Select & "      ,ISNULL(JK2.実績数,0) AS 実績数出, ISNULL(ZS1.使用量,0) AS 使用量出"
        gSL_Select = gSL_Select & " FROM 材料入荷テーブル SN"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(S1.入荷日,1,6) AS 入荷日, S1.材質, S1.板厚, S1.幅, S1.長さ, SUM(S1.入荷数) AS 入荷数"
        gSL_Select = gSL_Select & "                         FROM  材料入荷テーブル S1"
        gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(S1.入荷日,1,6), S1.材質, S1.板厚, S1.幅, S1.長さ) SN1"
        gSL_Select = gSL_Select & "        ON  SUBSTRING(SN1.入荷日,1,6) = SUBSTRING(SN.入荷日,1,6)"
        gSL_Select = gSL_Select & "        AND SN1.材質     = SN.材質"
        gSL_Select = gSL_Select & "        AND SN1.板厚     = SN.板厚"
        gSL_Select = gSL_Select & "        AND SN1.幅       = SN.幅"
        gSL_Select = gSL_Select & "        AND SN1.長さ     = SN.長さ"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J1.実績日,1,6) AS 実績日, J1.材質, J1.板厚, J1.幅, J1.長さ, SUM(J1.実績数) AS 実績数"
        gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J1"
        gSL_Select = gSL_Select & "                         WHERE J1.実績区分 = 1"
        gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J1.実績日,1,6), J1.材質, J1.板厚, J1.幅, J1.長さ) JK1"
        gSL_Select = gSL_Select & "        ON  SUBSTRING(JK1.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
        gSL_Select = gSL_Select & "        AND JK1.材質     = SN.材質"
        gSL_Select = gSL_Select & "        AND JK1.板厚     = SN.板厚"
        gSL_Select = gSL_Select & "        AND JK1.幅       = SN.幅"
        gSL_Select = gSL_Select & "        AND JK1.長さ     = SN.長さ"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J2.実績日,1,6) AS 実績日, J2.材質, J2.板厚, J2.幅, J2.長さ, SUM(J2.実績数) AS 実績数"
        gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J2"
        gSL_Select = gSL_Select & "                         WHERE J2.実績区分 <> 1 AND J2.入力種類 <> 'A'"
        gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J2.実績日,1,6), J2.材質, J2.板厚, J2.幅, J2.長さ) JK2"
        gSL_Select = gSL_Select & "        ON  SUBSTRING(JK2.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
        gSL_Select = gSL_Select & "        AND JK2.材質     = SN.材質"
        gSL_Select = gSL_Select & "        AND JK2.板厚     = SN.板厚"
        gSL_Select = gSL_Select & "        AND JK2.幅       = SN.幅"
        gSL_Select = gSL_Select & "        AND JK2.長さ     = SN.長さ"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(Z1.作業日,1,6) AS 作業日, Z1.材質, Z1.板厚, Z1.幅, Z1.長さ, SUM(Z1.使用量) AS 使用量"
        gSL_Select = gSL_Select & "                         FROM  材料使用量テーブル Z1"
        gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(z1.作業日,1,6), Z1.材質, Z1.板厚, Z1.幅, Z1.長さ) ZS1"
        gSL_Select = gSL_Select & "        ON  SUBSTRING(ZS1.作業日,1,6) = SUBSTRING(SN.入荷日,1,6)"
        gSL_Select = gSL_Select & "        AND ZS1.材質     = SN.材質"
        gSL_Select = gSL_Select & "        AND ZS1.板厚     = SN.板厚"
        gSL_Select = gSL_Select & "        AND ZS1.幅       = SN.幅"
        gSL_Select = gSL_Select & "        AND ZS1.長さ     = SN.長さ"
        gSL_Select = gSL_Select & " WHERE SUBSTRING(SN.入荷日,1,6) >= '" & Trim(YM0) & "'"
        gSL_Select = gSL_Select & "   AND SN.材料管理番号 = " & CStr(ZaiNo)
''''''''gSL_Select = gSL_Select & "   AND SN.材質 = '" & Trim(iH1_Zaist) & "'"
''''''''gSL_Select = gSL_Select & "   AND SN.板厚 =  " & Trim(iH1_Itatu) & " "
''''''''gSL_Select = gSL_Select & "   AND SN.幅   =  " & Trim(iH1_Width) & " "
''''''''gSL_Select = gSL_Select & "   AND SN.長さ =  " & Trim(iH1_Long) & " "
        gSL_Select = gSL_Select & " GROUP BY SUBSTRING(SN.入荷日,1,6), SN.材質,   SN.板厚,   SN.幅, SN.長さ"
        gSL_Select = gSL_Select & "         ,SN1.入荷数, JK1.実績数  , JK2.実績数,ZS1.使用量"
        gSL_Select = gSL_Select & " ORDER BY SUBSTRING(SN.入荷日,1,6)"
    Else
        '---在庫マスタ：有
        wKanriNo = ZZM.材料管理番号
        wMasterSu = ZZM.在庫数
        If Trim(ZZM.実棚日) = "" Then
            '---実棚日：空白
            gSL_Select = ""
            gSL_Select = gSL_Select & ""
            gSL_Select = gSL_Select & "SELECT SUBSTRING(SN.入荷日,1,6) AS 入荷日, SN.材質, SN.板厚, SN.幅, SN.長さ"
            gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) + ISNULL(JK1.実績数,0) - ISNULL(ZS1.使用量,0) - ISNULL(JK2.実績数,0) AS 入荷数"
            gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) AS 入荷数入, ISNULL(JK1.実績数,0) AS 実績数入"
            gSL_Select = gSL_Select & "      ,ISNULL(JK2.実績数,0) AS 実績数出, ISNULL(ZS1.使用量,0) AS 使用量出"
            gSL_Select = gSL_Select & " FROM 材料入荷テーブル SN"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(S1.入荷日,1,6) AS 入荷日, S1.材質, S1.板厚, S1.幅, S1.長さ, SUM(S1.入荷数) AS 入荷数"
            gSL_Select = gSL_Select & "                         FROM  材料入荷テーブル S1"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(S1.入荷日,1,6), S1.材質, S1.板厚, S1.幅, S1.長さ) SN1"
            gSL_Select = gSL_Select & "        ON  SUBSTRING(SN1.入荷日,1,6) = SUBSTRING(SN.入荷日,1,6)"
            gSL_Select = gSL_Select & "        AND SN1.材質     = SN.材質"
            gSL_Select = gSL_Select & "        AND SN1.板厚     = SN.板厚"
            gSL_Select = gSL_Select & "        AND SN1.幅       = SN.幅"
            gSL_Select = gSL_Select & "        AND SN1.長さ     = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J1.実績日,1,6) AS 実績日, J1.材質, J1.板厚, J1.幅, J1.長さ, SUM(J1.実績数) AS 実績数"
            gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J1"
            gSL_Select = gSL_Select & "                         WHERE J1.実績区分 = 1"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J1.実績日,1,6), J1.材質, J1.板厚, J1.幅, J1.長さ) JK1"
            gSL_Select = gSL_Select & "        ON  SUBSTRING(JK1.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
            gSL_Select = gSL_Select & "        AND JK1.材質     = SN.材質"
            gSL_Select = gSL_Select & "        AND JK1.板厚     = SN.板厚"
            gSL_Select = gSL_Select & "        AND JK1.幅       = SN.幅"
            gSL_Select = gSL_Select & "        AND JK1.長さ     = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J2.実績日,1,6) AS 実績日, J2.材質, J2.板厚, J2.幅, J2.長さ, SUM(J2.実績数) AS 実績数"
            gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J2"
            gSL_Select = gSL_Select & "                         WHERE J2.実績区分 <> 1 AND J2.入力種類 <> 'A'"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J2.実績日,1,6), J2.材質, J2.板厚, J2.幅, J2.長さ) JK2"
            gSL_Select = gSL_Select & "        ON  SUBSTRING(JK2.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
            gSL_Select = gSL_Select & "        AND JK2.材質     = SN.材質"
            gSL_Select = gSL_Select & "        AND JK2.板厚     = SN.板厚"
            gSL_Select = gSL_Select & "        AND JK2.幅       = SN.幅"
            gSL_Select = gSL_Select & "        AND JK2.長さ     = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(Z1.作業日,1,6) AS 作業日, Z1.材質, Z1.板厚, Z1.幅, Z1.長さ, SUM(Z1.使用量) AS 使用量"
            gSL_Select = gSL_Select & "                         FROM  材料使用量テーブル Z1"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(z1.作業日,1,6), Z1.材質, Z1.板厚, Z1.幅, Z1.長さ) ZS1"
            gSL_Select = gSL_Select & "        ON  SUBSTRING(ZS1.作業日,1,6) = SUBSTRING(SN.入荷日,1,6)"
            gSL_Select = gSL_Select & "        AND ZS1.材質     = SN.材質"
            gSL_Select = gSL_Select & "        AND ZS1.板厚     = SN.板厚"
            gSL_Select = gSL_Select & "        AND ZS1.幅       = SN.幅"
            gSL_Select = gSL_Select & "        AND ZS1.長さ     = SN.長さ"
            gSL_Select = gSL_Select & " WHERE SUBSTRING(SN.入荷日,1,6) >= '" & Trim(YM0) & "'"
            gSL_Select = gSL_Select & "   AND SN.材料管理番号 = " & Trim(Str(wKanriNo))
            gSL_Select = gSL_Select & " GROUP BY SUBSTRING(SN.入荷日,1,6), SN.材質,   SN.板厚,   SN.幅, SN.長さ"
            gSL_Select = gSL_Select & "         ,SN1.入荷数, JK1.実績数  , JK2.実績数,ZS1.使用量"
            gSL_Select = gSL_Select & " ORDER BY SUBSTRING(SN.入荷日,1,6)"
        Else
            '---実棚日：入力済
            gSL_Select = ""
            gSL_Select = gSL_Select & ""
            gSL_Select = gSL_Select & "SELECT SUBSTRING(SN.入荷日,1,6) AS 入荷日, SN.材質, SN.板厚, SN.幅, SN.長さ"
            gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) + ISNULL(JK1.実績数,0) - ISNULL(ZS1.使用量,0) - ISNULL(JK2.実績数,0) AS 入荷数"
            gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) AS 入荷数入, ISNULL(JK1.実績数,0) AS 実績数入"
            gSL_Select = gSL_Select & "      ,ISNULL(JK2.実績数,0) AS 実績数出, ISNULL(ZS1.使用量,0) AS 使用量出"
            gSL_Select = gSL_Select & " FROM 材料入荷テーブル SN"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(S1.入荷日,1,6) AS 入荷日, S1.材質, S1.板厚, S1.幅, S1.長さ, SUM(S1.入荷数) AS 入荷数"
            gSL_Select = gSL_Select & "                         FROM  材料入荷テーブル S1"
            gSL_Select = gSL_Select & "                         WHERE S1.入荷日 > '" & Trim(ZZM.実棚日) & "'"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(S1.入荷日,1,6), S1.材質, S1.板厚, S1.幅, S1.長さ) SN1"
          ''gSL_Select = gSL_Select & "      --ON  SUBSTRING(SN1.入荷日,1,6) = SUBSTRING(SN.入荷日,1,6)"
          ''gSL_Select = gSL_Select & "      --AND"
            gSL_Select = gSL_Select & "        ON  SN1.材質 = SN.材質 AND SN1.板厚 = SN.板厚 AND SN1.幅 = SN.幅 AND SN1.長さ = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J1.実績日,1,6) AS 実績日, J1.材質, J1.板厚, J1.幅, J1.長さ, SUM(J1.実績数) AS 実績数"
            gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J1"
            gSL_Select = gSL_Select & "                         WHERE J1.実績区分 = 1 AND J1.実績日 > '" & Trim(ZZM.実棚日) & "'"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J1.実績日,1,6), J1.材質, J1.板厚, J1.幅, J1.長さ) JK1"
          ''gSL_Select = gSL_Select & "      --ON  SUBSTRING(JK1.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
          ''gSL_Select = gSL_Select & "      --AND"
            gSL_Select = gSL_Select & "        ON  JK1.材質 = SN.材質 AND JK1.板厚 = SN.板厚 AND JK1.幅 = SN.幅 AND JK1.長さ = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J2.実績日,1,6) AS 実績日, J2.材質, J2.板厚, J2.幅, J2.長さ, SUM(J2.実績数) AS 実績数"
            gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J2"
            gSL_Select = gSL_Select & "                         WHERE J2.実績区分 <> 1 AND J2.入力種類 <> 'A' AND J2.実績日 > '" & Trim(ZZM.実棚日) & "'"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J2.実績日,1,6), J2.材質, J2.板厚, J2.幅, J2.長さ) JK2"
          ''gSL_Select = gSL_Select & "      --ON  SUBSTRING(JK2.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
          ''gSL_Select = gSL_Select & "      --AND"
            gSL_Select = gSL_Select & "        ON  JK2.材質 = SN.材質 AND JK2.板厚 = SN.板厚 AND JK2.幅 = SN.幅 AND JK2.長さ = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(Z1.作業日,1,6) AS 作業日, Z1.材質, Z1.板厚, Z1.幅, Z1.長さ, SUM(Z1.使用量) AS 使用量"
            gSL_Select = gSL_Select & "                         FROM  材料使用量テーブル Z1"
            gSL_Select = gSL_Select & "                         WHERE Z1.作業日   > '" & Trim(ZZM.実棚日) & "'"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(z1.作業日,1,6), Z1.材質, Z1.板厚, Z1.幅, Z1.長さ) ZS1"
          ''gSL_Select = gSL_Select & "      --ON  SUBSTRING(ZS1.作業日,1,6) = SUBSTRING(SN.入荷日,1,6)"
          ''gSL_Select = gSL_Select & "      --AND"
            gSL_Select = gSL_Select & "        ON  ZS1.材質 = SN.材質 AND ZS1.板厚 = SN.板厚 AND ZS1.幅 = SN.幅 AND ZS1.長さ = SN.長さ"
            gSL_Select = gSL_Select & " WHERE SUBSTRING(SN.入荷日,1,6) >= '" & Trim(YM0) & "' AND SN.材料管理番号 = " & Trim(Str(ZZM.材料管理番号))
            gSL_Select = gSL_Select & " GROUP BY SUBSTRING(SN.入荷日,1,6), SN.材質,   SN.板厚,   SN.幅, SN.長さ"
            gSL_Select = gSL_Select & "         ,SN1.入荷数, JK1.実績数  , JK2.実績数,ZS1.使用量"
            gSL_Select = gSL_Select & " ORDER BY SUBSTRING(SN.入荷日,1,6)"
        End If
    End If
    
    If ZYTRead(gSL_Select) Then
        Do While ZYT_RDSTS
            wZaikoSu = wZaikoSu + ZYT.入荷数
            Call ZYTReadNext
        Loop
    End If
    
    wZaikoSu = wMasterSu + wZaikoSu
    HKT.材料在庫量 = wZaikoSu

    GetZaikoSu = True
End Function
'---------------------------------------------------'
'       構成問合せワーク                            '
'---------------------------------------------------'
Public Function SelectKOL(HinBN As String) As Boolean

    gSL_Select = "SELECT MAX(KTW.親品番) AS 親品番, HTM.背番号 AS 表示品番"
    gSL_Select = gSL_Select & " FROM  構成問合せワーク KTW"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT 品番, MAX(背番号) AS 背番号"
    gSL_Select = gSL_Select & "                         From 品番取引先マスタ"
    gSL_Select = gSL_Select & "                         Where 品目 = 0"
    gSL_Select = gSL_Select & "                         GROUP BY 品番) HTM"
    gSL_Select = gSL_Select & "        ON  HTM.品番 = KTW.親品番"
    gSL_Select = gSL_Select & " WHERE 処理端末 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND 子品番   = '" & Trim(HinBN) & "'"
    gSL_Select = gSL_Select & " GROUP BY HTM.背番号"
    
End Function

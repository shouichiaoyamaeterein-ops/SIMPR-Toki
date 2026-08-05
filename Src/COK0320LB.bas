Attribute VB_Name = "COK0320LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   支払金額連絡表　印刷
'*
'*   モジュールID  : COK0320LB.Bas
'*
'*   戻値 : なし
'*
'*   作成日  :  2004/12/09   By CIS
'*   変更日  :  2008/01/24  By CIS - 受給追加
'*   変更日  :  2009/02/10  By CIS
'*   変更日  :  2009/05/12  By CIS 区切り線、背番号追加
'*---------------------------------------------------------------------------------
Public CisPrtCOK0320        As New CISFormPrint
Public CisFrmCOK0320        As New CISFormContorl
'>> Print
Public pTitle               As String
'>> CSV
Public pCsvDir              As String
Public pCsvKey              As String
'>> DBNo
Public DBNoCOK0320          As Integer
'
'>> 自社名称
Public pJisyaNm             As String
'>> 帳票可変項目テーブル
Public pFreeMsg(4)          As String
'>> 帳票ID
Public pPrtID(5)            As String
'
Global PrtFlagHyosi         As Boolean      '>支払金額連絡書
Global PrtFlagSire          As Boolean      '>仕入明細
Global PrtFlagSkyu          As Boolean      '>支給明細
Global PrtYM                As String
Global PrtTorcd()           As String       '>請求先
Global PrtDidcd()           As String       '>識別子
'
'>> CSV出力用
    Dim OutFNo              As Integer
    Dim OutStr              As String
    Dim OutCnt              As Long
    Dim OutSeq              As Long
'
'+++++++++++++++++++++++++++++++'
'+      印　刷　処　理
'+++++++++++++++++++++++++++++++'
Public Function COK0320LPrint(Optional PrintOut As Integer = 0 _
                            , Optional DBNo As Integer = 0) As Boolean
    COK0320LPrint = False
'
    Call DBSelect_TRM(1)
    Call DBSelect_CKT(1)
    Call DBSelect_SYM(1)
    Call SqlSelect
    If Not KSHRead(gSL_Select, , DBNo) Then
        Call KSHClose(DBNo)
        Exit Function
    End If
    DBNoCOK0320 = DBNo
'
    If PrintOut = 2 Then
        Call COK0320CsvPut(DBNoCOK0320)
    Else
        With CisPrtCOK0320
            .cDataMaxCnt = CisDB.RecordCount
            Set .iPrtForm = COK0320L
            .dpMSG = pTitle & "　印刷中"
            Select Case PrintOut
                Case 0: .iPreview = pv_Print
                Case 1: .iPreview = pv_Preview
            End Select
            .dpMSGScreen = モードレス
            .PrintStart
        End With
    End If
'
    COK0320LPrint = True
End Function
'*-----------------------------------*
'*      テーブル読込用SQL文
'*-----------------------------------*
Private Sub SqlSelect()
Dim wIndex      As Integer
Dim wConma      As String

    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  ksh.*"
    gSL_Select = gSL_Select & ", isnull(trm.正式名称, '') 契約先名"
    gSL_Select = gSL_Select & ", isnull(sy1.値名称, '') 支払区分名1"
    gSL_Select = gSL_Select & ", isnull(sy2.値名称, '') 支払区分名2"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  買掛支払テーブル ksh"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ trm"
    gSL_Select = gSL_Select & "    on trm.取引先CD = ksh.契約先"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 1"
    gSL_Select = gSL_Select & "  left outer join 名称マスタ sy1"
    gSL_Select = gSL_Select & "    on sy1.区分名称 = '買掛支払区分'"
    gSL_Select = gSL_Select & "   and sy1.区分タイプ = ''"
    gSL_Select = gSL_Select & "   and sy1.値 = ksh.支払区分1"
    gSL_Select = gSL_Select & "  left outer join 名称マスタ sy2"
    gSL_Select = gSL_Select & "    on sy2.区分名称 = '買掛支払区分'"
    gSL_Select = gSL_Select & "   and sy2.区分タイプ = ''"
    gSL_Select = gSL_Select & "   and sy2.値 = ksh.支払区分2"
    gSL_Select = gSL_Select & " where ksh.実績年月 = '" & Trim$(PrtYM) & "'"
    If 1 < UBound(PrtTorcd) Then
        gSL_Select = gSL_Select & "   and ksh.契約先 in"
        gSL_Select = gSL_Select & " ("
        wConma = " "
        For wIndex = 1 To UBound(PrtTorcd)
            gSL_Select = gSL_Select & wConma & "'" & RTrim$(PrtTorcd(wIndex)) & "'"
            wConma = ","
        Next wIndex
        gSL_Select = gSL_Select & " )"
    End If
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  ksh.契約先"
End Sub
'*-----------------------------------*
'+      自社名称の索引処理
'*-----------------------------------*
Private Sub DBSelect_TRM(Optional DBNo As Integer = 1)
    gSL_Select = ""
    gSL_Select = gSL_Select & "select 正式名称"
    gSL_Select = gSL_Select & "  from 取引先マスタ"
    gSL_Select = gSL_Select & " where 取引先CD = '" & RTrim$(gCompany) & "'"
    gSL_Select = gSL_Select & "   and 取引先区分 = 1"
    If TRMRead(gSL_Select, 1, DBNo) Then
        pJisyaNm = TRM.正式名称
    Else
        pJisyaNm = ""
    End If
End Sub
'*-----------------------------------*
'*      実績テーブル読込用SQL文
'*-----------------------------------*
Public Sub SqlSelect_Jiseki()
Dim wIndex          As Integer
Dim wSel_Didcd      As String
'
'   >> 契約先より抽出対象の識別子のｾｯﾄ
    wSel_Didcd = ""
    For wIndex = 1 To UBound(PrtTorcd)
        If RTrim$(KSH.契約先) = RTrim$(PrtTorcd(wIndex)) Then
            wSel_Didcd = RTrim$(PrtDidcd(wIndex))
            Exit For
        End If
    Next wIndex
'
    gSL_Select = ""
    If PrtFlagSire = True Then
        gSL_Select = gSL_Select & "select" & vbCrLf
        gSL_Select = gSL_Select & "  kjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.品番" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.検収日" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.納入番号" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.数量" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.単価" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.仮単価区分" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.金額" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.計算方法" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.検収管理NO" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(sy1.値名称, '') 識別子名" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.表示品番, kjs.品番) 表示品番" & vbCrLf
        '============================================================= *** 2008/06/16 Start
        gSL_Select = gSL_Select & ", '' 部品材料区分 " & vbCrLf
        gSL_Select = gSL_Select & ", '' 表示寸法 " & vbCrLf
        gSL_Select = gSL_Select & ", '' 材質 " & vbCrLf
        gSL_Select = gSL_Select & ", 0 材料管理番号" & vbCrLf
        '============================================================= *** 2008/06/16 End
'''        gSL_Select = gSL_Select & ", kjs.背番号" & vbCrLf             '*** 2009/05/12
        gSL_Select = gSL_Select & " from" & vbCrLf
        gSL_Select = gSL_Select & "  購入実績テーブル kjs" & vbCrLf
        gSL_Select = gSL_Select & "  left outer join 名称マスタ sy1" & vbCrLf
        gSL_Select = gSL_Select & "    on sy1.区分名称 = '買掛識別子'" & vbCrLf
        gSL_Select = gSL_Select & "   and sy1.区分タイプ = ''" & vbCrLf
        gSL_Select = gSL_Select & "   and sy1.値 = kjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & "  left outer join 品番マスタ hnm" & vbCrLf
        gSL_Select = gSL_Select & "    on hnm.品番 = kjs.品番" & vbCrLf
        gSL_Select = gSL_Select & " where kjs.実績年月 = '" & Trim$(KSH.実績年月) & "'" & vbCrLf
        gSL_Select = gSL_Select & "   and kjs.契約先 = '" & RTrim$(KSH.契約先) & "'" & vbCrLf
        If wSel_Didcd <> "" Then
            gSL_Select = gSL_Select & "   and kjs.識別子 in (" & wSel_Didcd & ")" & vbCrLf
        End If
        gSL_Select = gSL_Select & "   and not (kjs.識別子 = '01' and kjs.計算方法 = 0)" & vbCrLf
        gSL_Select = gSL_Select & "   and not (kjs.識別子 = '41' and kjs.計算方法 = 0)" & vbCrLf
        gSL_Select = gSL_Select & " union all " & vbCrLf
        gSL_Select = gSL_Select & "select" & vbCrLf
        gSL_Select = gSL_Select & "  kjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.品番" & vbCrLf
        gSL_Select = gSL_Select & ", '' 検収日" & vbCrLf
        gSL_Select = gSL_Select & ", '' 納入番号" & vbCrLf
        gSL_Select = gSL_Select & ", sum(kjs.数量) 金額" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.単価" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.仮単価区分" & vbCrLf
        gSL_Select = gSL_Select & ", sum(kjs.金額) 金額" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.計算方法" & vbCrLf
        gSL_Select = gSL_Select & ", 0 検収管理NO" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(sy1.値名称, '') 識別子名" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.表示品番, kjs.品番) 表示品番" & vbCrLf
        '============================================================= *** 2008/06/16 Start
        gSL_Select = gSL_Select & ", '' 部品材料区分 " & vbCrLf
        gSL_Select = gSL_Select & ", '' 表示寸法 " & vbCrLf
        gSL_Select = gSL_Select & ", '' 材質 " & vbCrLf
        gSL_Select = gSL_Select & ", 0 材料管理番号" & vbCrLf
        '============================================================= *** 2008/06/16 End
'''        gSL_Select = gSL_Select & ", kjs.背番号" & vbCrLf             '*** 2009/05/12
        gSL_Select = gSL_Select & " from" & vbCrLf
        gSL_Select = gSL_Select & "  購入実績テーブル kjs" & vbCrLf
        gSL_Select = gSL_Select & "  left outer join 名称マスタ sy1" & vbCrLf
        gSL_Select = gSL_Select & "    on sy1.区分名称 = '買掛識別子'" & vbCrLf
        gSL_Select = gSL_Select & "   and sy1.区分タイプ = ''" & vbCrLf
        gSL_Select = gSL_Select & "   and sy1.値 = kjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & "  left outer join 品番マスタ hnm" & vbCrLf
        gSL_Select = gSL_Select & "    on hnm.品番 = kjs.品番" & vbCrLf
        gSL_Select = gSL_Select & " where kjs.実績年月 = '" & Trim$(KSH.実績年月) & "'" & vbCrLf
        gSL_Select = gSL_Select & "   and kjs.契約先 = '" & RTrim$(KSH.契約先) & "'" & vbCrLf
        If wSel_Didcd <> "" Then
            gSL_Select = gSL_Select & "   and kjs.識別子 in (" & wSel_Didcd & ")" & vbCrLf
        End If
        gSL_Select = gSL_Select & "   and ((kjs.識別子 = '01' and kjs.計算方法 = 0)" & vbCrLf
        gSL_Select = gSL_Select & "   OR (kjs.識別子 = '41' and kjs.計算方法 = 0))" & vbCrLf
        gSL_Select = gSL_Select & " group by" & vbCrLf
        gSL_Select = gSL_Select & "  kjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.品番" & vbCrLf
'''        gSL_Select = gSL_Select & ", kjs.背番号" & vbCrLf               '*** 2009/05/12
        gSL_Select = gSL_Select & ", kjs.単価" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.仮単価区分" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.計算方法" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(sy1.値名称, '')" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.表示品番, kjs.品番)" & vbCrLf
        
        '==================================================== *** 2008/06/16 Start
        gSL_Select = gSL_Select & " union all " & vbCrLf
        gSL_Select = gSL_Select & "select" & vbCrLf
        gSL_Select = gSL_Select & "  kjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & ", '' 品番 " & vbCrLf
        gSL_Select = gSL_Select & ", kjs.検収日" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.納入番号" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.数量" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.単価" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.仮単価区分" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.金額" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.計算方法" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.検収管理NO" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(sy1.値名称, '') 識別子名" & vbCrLf
        gSL_Select = gSL_Select & ", '' 表示品番" & vbCrLf
        gSL_Select = gSL_Select & ", 'Z' 部品材料区分 " & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.表示寸法,'') 表示寸法 " & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.材質,'') 材質 " & vbCrLf
        gSL_Select = gSL_Select & ",  kjs.材料管理番号" & vbCrLf
'''        gSL_Select = gSL_Select & ", '' 背番号" & vbCrLf             '*** 2009/05/12
        gSL_Select = gSL_Select & " from" & vbCrLf
        gSL_Select = gSL_Select & "  材料購入実績テーブル kjs" & vbCrLf
        gSL_Select = gSL_Select & "  left outer join 名称マスタ sy1" & vbCrLf
        gSL_Select = gSL_Select & "    on sy1.区分名称 = '買掛識別子'" & vbCrLf
        gSL_Select = gSL_Select & "   and sy1.区分タイプ = ''" & vbCrLf
        gSL_Select = gSL_Select & "   and sy1.値 = kjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & "  left outer join 材料マスタ hnm" & vbCrLf
        gSL_Select = gSL_Select & "    on hnm.材料管理番号 = kjs.材料管理番号" & vbCrLf
        gSL_Select = gSL_Select & " where kjs.実績年月 = '" & Trim$(KSH.実績年月) & "'" & vbCrLf
        gSL_Select = gSL_Select & "   and kjs.契約先 = '" & RTrim$(KSH.契約先) & "'" & vbCrLf
        If wSel_Didcd <> "" Then
            gSL_Select = gSL_Select & "   and kjs.識別子 in (" & wSel_Didcd & ")" & vbCrLf
        End If
        gSL_Select = gSL_Select & "   and not (kjs.識別子 = '51' and kjs.計算方法 = 0)" & vbCrLf
        gSL_Select = gSL_Select & " union all " & vbCrLf
        
        gSL_Select = gSL_Select & "select" & vbCrLf
        gSL_Select = gSL_Select & "  kjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & ", '' 品番" & vbCrLf
        gSL_Select = gSL_Select & ", '' 検収日" & vbCrLf
        gSL_Select = gSL_Select & ", '' 納入番号" & vbCrLf
        gSL_Select = gSL_Select & ", sum(kjs.数量) 金額" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.単価" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.仮単価区分" & vbCrLf
        gSL_Select = gSL_Select & ", sum(kjs.金額) 金額" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.計算方法" & vbCrLf
        gSL_Select = gSL_Select & ", 0 検収管理NO" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(sy1.値名称, '') 識別子名" & vbCrLf
        gSL_Select = gSL_Select & ", '' 表示品番" & vbCrLf
        gSL_Select = gSL_Select & ", 'Z' 部品材料区分 " & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.表示寸法,'') 表示寸法 " & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.材質,'') 材質 " & vbCrLf
        gSL_Select = gSL_Select & ", kjs.材料管理番号" & vbCrLf
'''        gSL_Select = gSL_Select & ", '' 背番号" & vbCrLf             '*** 2009/05/12
        gSL_Select = gSL_Select & " from" & vbCrLf
        gSL_Select = gSL_Select & "  材料購入実績テーブル kjs" & vbCrLf
        gSL_Select = gSL_Select & "  left outer join 名称マスタ sy1" & vbCrLf
        gSL_Select = gSL_Select & "    on sy1.区分名称 = '買掛識別子'" & vbCrLf
        gSL_Select = gSL_Select & "   and sy1.区分タイプ = ''" & vbCrLf
        gSL_Select = gSL_Select & "   and sy1.値 = kjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & "  left outer join 材料マスタ hnm" & vbCrLf
        gSL_Select = gSL_Select & "    on hnm.材料管理番号 = kjs.材料管理番号" & vbCrLf
        gSL_Select = gSL_Select & " where kjs.実績年月 = '" & Trim$(KSH.実績年月) & "'" & vbCrLf
        gSL_Select = gSL_Select & "   and kjs.契約先 = '" & RTrim$(KSH.契約先) & "'" & vbCrLf
        If wSel_Didcd <> "" Then
            gSL_Select = gSL_Select & "   and kjs.識別子 in (" & wSel_Didcd & ")" & vbCrLf
        End If
        gSL_Select = gSL_Select & "   and (kjs.識別子 = '51' and kjs.計算方法 = 0)" & vbCrLf
        gSL_Select = gSL_Select & " group by" & vbCrLf
        gSL_Select = gSL_Select & "  kjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.材料管理番号" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.単価" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.仮単価区分" & vbCrLf
        gSL_Select = gSL_Select & ", kjs.計算方法" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(sy1.値名称, '')" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.表示寸法, '')" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.材質, '')" & vbCrLf
        
        '==================================================== *** 2008/06/16 End
    End If
    
    If PrtFlagSire = True And _
       PrtFlagSkyu = True Then
        gSL_Select = gSL_Select & " union all " & vbCrLf
    End If
    If PrtFlagSkyu = True Then
        gSL_Select = gSL_Select & "select" & vbCrLf
        gSL_Select = gSL_Select & "  sjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.品番" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.検収日" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.納入番号" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.数量" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.単価" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.仮単価区分" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.金額" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.計算方法" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.検収管理NO" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(sy2.値名称, '') 識別子名" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.表示品番, sjs.品番) 表示品番" & vbCrLf
        '============================================================= *** 2008/06/16 Start
        gSL_Select = gSL_Select & ", '' 部品材料区分 " & vbCrLf
        gSL_Select = gSL_Select & ", '' 表示寸法 " & vbCrLf
        gSL_Select = gSL_Select & ", '' 材質 " & vbCrLf
        gSL_Select = gSL_Select & ", 0 材料管理番号" & vbCrLf
        '============================================================= *** 2008/06/16 End
'''        gSL_Select = gSL_Select & ", sjs.背番号" & vbCrLf             '*** 2009/05/12
        gSL_Select = gSL_Select & " from" & vbCrLf
        gSL_Select = gSL_Select & "  支給実績テーブル sjs" & vbCrLf
        gSL_Select = gSL_Select & "  left outer join 名称マスタ sy2" & vbCrLf
        gSL_Select = gSL_Select & "    on sy2.区分名称 = '買掛識別子'" & vbCrLf
        gSL_Select = gSL_Select & "   and sy2.区分タイプ = ''" & vbCrLf
        gSL_Select = gSL_Select & "   and sy2.値 = sjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & "  left outer join 品番マスタ hnm" & vbCrLf
        gSL_Select = gSL_Select & "    on hnm.品番 = sjs.品番" & vbCrLf
        gSL_Select = gSL_Select & " where sjs.実績年月 = '" & Trim$(KSH.実績年月) & "'" & vbCrLf
        gSL_Select = gSL_Select & "   and sjs.契約先 = '" & RTrim$(KSH.契約先) & "'" & vbCrLf
        If wSel_Didcd <> "" Then
            gSL_Select = gSL_Select & "   and sjs.識別子 in (" & wSel_Didcd & ")" & vbCrLf
        End If
        gSL_Select = gSL_Select & "   and not (sjs.識別子 = '21' and sjs.計算方法 = 0)" & vbCrLf
        gSL_Select = gSL_Select & " union all " & vbCrLf
        gSL_Select = gSL_Select & "select" & vbCrLf
        gSL_Select = gSL_Select & "  sjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.品番" & vbCrLf
        gSL_Select = gSL_Select & ", '' 検収日" & vbCrLf
        gSL_Select = gSL_Select & ", '' 納入番号" & vbCrLf
        gSL_Select = gSL_Select & ", sum(sjs.数量) 数量" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.単価" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.仮単価区分" & vbCrLf
        gSL_Select = gSL_Select & ", sum(sjs.金額) 金額" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.計算方法" & vbCrLf
        gSL_Select = gSL_Select & ", 0 検収管理NO" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(sy2.値名称, '') 識別子名" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.表示品番, sjs.品番) 表示品番" & vbCrLf
        '============================================================= *** 2008/06/16 Start
        gSL_Select = gSL_Select & ", '' 部品材料区分 " & vbCrLf
        gSL_Select = gSL_Select & ", '' 表示寸法 " & vbCrLf
        gSL_Select = gSL_Select & ", '' 材質 " & vbCrLf
        gSL_Select = gSL_Select & ", 0 材料管理番号" & vbCrLf
        '============================================================= *** 2008/06/16 End
'''        gSL_Select = gSL_Select & ", '' 背番号" & vbCrLf             '*** 2009/05/12
        gSL_Select = gSL_Select & " from" & vbCrLf
        gSL_Select = gSL_Select & "  支給実績テーブル sjs" & vbCrLf
        gSL_Select = gSL_Select & "  left outer join 名称マスタ sy2" & vbCrLf
        gSL_Select = gSL_Select & "    on sy2.区分名称 = '買掛識別子'" & vbCrLf
        gSL_Select = gSL_Select & "   and sy2.区分タイプ = ''" & vbCrLf
        gSL_Select = gSL_Select & "   and sy2.値 = sjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & "  left outer join 品番マスタ hnm" & vbCrLf
        gSL_Select = gSL_Select & "    on hnm.品番 = sjs.品番" & vbCrLf
        gSL_Select = gSL_Select & " where sjs.実績年月 = '" & Trim$(KSH.実績年月) & "'" & vbCrLf
        gSL_Select = gSL_Select & "   and sjs.契約先 = '" & RTrim$(KSH.契約先) & "'" & vbCrLf
        If wSel_Didcd <> "" Then
            gSL_Select = gSL_Select & "   and sjs.識別子 in (" & wSel_Didcd & ")" & vbCrLf
        End If
        gSL_Select = gSL_Select & "   and (sjs.識別子 = '21' and sjs.計算方法 = 0)" & vbCrLf
        gSL_Select = gSL_Select & " group by" & vbCrLf
        gSL_Select = gSL_Select & "  sjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.品番" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.単価" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.仮単価区分" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.計算方法" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(sy2.値名称, '')" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.表示品番, sjs.品番)" & vbCrLf
        
        gSL_Select = gSL_Select & " union all " & vbCrLf
        '==================================================== *** 2008/06/16 Start
        gSL_Select = gSL_Select & "select" & vbCrLf
        gSL_Select = gSL_Select & "  sjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & ", '' 品番" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.検収日" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.納入番号" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.数量" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.単価" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.仮単価区分" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.金額" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.計算方法" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.検収管理NO" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(sy2.値名称, '') 識別子名" & vbCrLf
        gSL_Select = gSL_Select & ", '' 表示品番" & vbCrLf
        gSL_Select = gSL_Select & ", 'Z' 部品材料区分 " & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.表示寸法,'') 表示寸法 " & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.材質,'') 材質 " & vbCrLf
        gSL_Select = gSL_Select & ",  sjs.材料管理番号" & vbCrLf
'''        gSL_Select = gSL_Select & ", '' 背番号" & vbCrLf             '*** 2009/05/12
        gSL_Select = gSL_Select & " from" & vbCrLf
        gSL_Select = gSL_Select & "  材料支給実績テーブル sjs" & vbCrLf
        gSL_Select = gSL_Select & "  left outer join 名称マスタ sy2" & vbCrLf
        gSL_Select = gSL_Select & "    on sy2.区分名称 = '買掛識別子'" & vbCrLf
        gSL_Select = gSL_Select & "   and sy2.区分タイプ = ''" & vbCrLf
        gSL_Select = gSL_Select & "   and sy2.値 = sjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & "  left outer join 材料マスタ hnm" & vbCrLf
        gSL_Select = gSL_Select & "    on hnm.材料管理番号 = sjs.材料管理番号" & vbCrLf
        gSL_Select = gSL_Select & " where sjs.実績年月 = '" & Trim$(KSH.実績年月) & "'" & vbCrLf
        gSL_Select = gSL_Select & "   and sjs.契約先 = '" & RTrim$(KSH.契約先) & "'" & vbCrLf
        If wSel_Didcd <> "" Then
            gSL_Select = gSL_Select & "   and sjs.識別子 in (" & wSel_Didcd & ")" & vbCrLf
        End If
        gSL_Select = gSL_Select & "   and not (sjs.識別子 = '71' and sjs.計算方法 = 0)" & vbCrLf
        gSL_Select = gSL_Select & " union all " & vbCrLf
        gSL_Select = gSL_Select & "select" & vbCrLf
        gSL_Select = gSL_Select & "  sjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & ", '' 品番" & vbCrLf
        gSL_Select = gSL_Select & ", '' 検収日" & vbCrLf
        gSL_Select = gSL_Select & ", '' 納入番号" & vbCrLf
        gSL_Select = gSL_Select & ", sum(sjs.数量) 数量" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.単価" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.仮単価区分" & vbCrLf
        gSL_Select = gSL_Select & ", sum(sjs.金額) 金額" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.計算方法" & vbCrLf
        gSL_Select = gSL_Select & ", 0 検収管理NO" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(sy2.値名称, '') 識別子名" & vbCrLf
        gSL_Select = gSL_Select & ", '' 表示品番" & vbCrLf
        gSL_Select = gSL_Select & ", 'Z' 部品材料区分 " & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.表示寸法,'') 表示寸法 " & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.材質,'') 材質 " & vbCrLf
        gSL_Select = gSL_Select & ",  sjs.材料管理番号" & vbCrLf
'''        gSL_Select = gSL_Select & ", '' 背番号" & vbCrLf             '*** 2009/05/12
        gSL_Select = gSL_Select & " from" & vbCrLf
        gSL_Select = gSL_Select & "  材料支給実績テーブル sjs" & vbCrLf
        gSL_Select = gSL_Select & "  left outer join 名称マスタ sy2" & vbCrLf
        gSL_Select = gSL_Select & "    on sy2.区分名称 = '買掛識別子'" & vbCrLf
        gSL_Select = gSL_Select & "   and sy2.区分タイプ = ''" & vbCrLf
        gSL_Select = gSL_Select & "   and sy2.値 = sjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & "  left outer join 材料マスタ hnm" & vbCrLf
        gSL_Select = gSL_Select & "    on hnm.材料管理番号 = sjs.材料管理番号" & vbCrLf
        gSL_Select = gSL_Select & " where sjs.実績年月 = '" & Trim$(KSH.実績年月) & "'" & vbCrLf
        gSL_Select = gSL_Select & "   and sjs.契約先 = '" & RTrim$(KSH.契約先) & "'" & vbCrLf
        If wSel_Didcd <> "" Then
            '2009/02/10 Update
            'gSL_Select = gSL_Select & "   and ( sjs.識別子 in (" & wSel_Didcd & ")" & vbCrLf
            'gSL_Select = gSL_Select & "       or (sjs.識別子 = '71' and sjs.計算方法 = 0) )" & vbCrLf
                        
            gSL_Select = gSL_Select & "   and ( sjs.識別子 in (" & wSel_Didcd & ") )" & vbCrLf
            gSL_Select = gSL_Select & "   and (sjs.識別子 = '71' and sjs.計算方法 = 0)" & vbCrLf
        Else
            gSL_Select = gSL_Select & "   and (sjs.識別子 = '71' and sjs.計算方法 = 0)" & vbCrLf
        End If
        gSL_Select = gSL_Select & " group by" & vbCrLf
        gSL_Select = gSL_Select & "  sjs.識別子" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.材料管理番号" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.単価" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.仮単価区分" & vbCrLf
        gSL_Select = gSL_Select & ", sjs.計算方法" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(sy2.値名称, '')" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.表示寸法, '')" & vbCrLf
        gSL_Select = gSL_Select & ", isnull(hnm.材質, '')" & vbCrLf
        '==================================================== *** 2008/06/16 End
    
    End If
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  1"
    gSL_Select = gSL_Select & ", 2"
    gSL_Select = gSL_Select & ", 3"
    gSL_Select = gSL_Select & ", 4"
    '==================================================== *** 2008/06/16 Start
    gSL_Select = gSL_Select & ", 5"
    gSL_Select = gSL_Select & ", 6"
    gSL_Select = gSL_Select & ", 7"
    gSL_Select = gSL_Select & ", 8"
    '==================================================== *** 2008/06/16 End
End Sub
'*-----------------------------------*
'*      明細テーブル読込用SQL文
'*-----------------------------------*
Public Sub SqlSelect_Meisai()
    gSL_Select = ""
    If RTrim$(KJS.識別子) = "01" Or RTrim$(KJS.識別子) = "41" Then
        gSL_Select = gSL_Select & "select"
        gSL_Select = gSL_Select & "  kms.検収日"
        gSL_Select = gSL_Select & ", kms.納入番号"
        gSL_Select = gSL_Select & ", kms.数量"
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  購入明細テーブル kms"
        If KJS.計算方法 = 1 Then
            gSL_Select = gSL_Select & " where kms.集約検収管理NO = " & KJS.検収管理NO
        Else
            gSL_Select = gSL_Select & " where kms.実績年月 = '" & Trim$(KSH.実績年月) & "'"
            gSL_Select = gSL_Select & "   and kms.契約先 = '" & RTrim$(KSH.契約先) & "'"
            gSL_Select = gSL_Select & "   and kms.識別子 = '" & RTrim$(KJS.識別子) & "'"
            gSL_Select = gSL_Select & "   and kms.品番 = '" & RTrim$(KJS.品番) & "'"
            gSL_Select = gSL_Select & "   and kms.単価 = '" & RTrim$(KJS.単価) & "'"
        End If
    Else
        gSL_Select = gSL_Select & "select"
        gSL_Select = gSL_Select & "  sms.検収日"
        gSL_Select = gSL_Select & ", sms.納入番号"
        gSL_Select = gSL_Select & ", sms.数量"
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  支給明細テーブル sms"
        If KJS.計算方法 = 1 Then
            gSL_Select = gSL_Select & " where sms.集約検収管理NO = " & KJS.検収管理NO
        Else
            gSL_Select = gSL_Select & " where sms.実績年月 = '" & Trim$(KSH.実績年月) & "'"
            gSL_Select = gSL_Select & "   and sms.契約先 = '" & RTrim$(KSH.契約先) & "'"
            gSL_Select = gSL_Select & "   and sms.識別子 = '" & RTrim$(KJS.識別子) & "'"
            gSL_Select = gSL_Select & "   and sms.品番 = '" & RTrim$(KJS.品番) & "'"
            gSL_Select = gSL_Select & "   and sms.単価 = '" & RTrim$(KJS.単価) & "'"
        End If
    End If
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  1"
    gSL_Select = gSL_Select & ", 2"
End Sub
'==================================================== *** 2008/06/16 Start
'*-----------------------------------*
'*      明細テーブル読込用SQL文(材料)
'*-----------------------------------*
Public Sub SqlSelect_Meisai_Z()
    gSL_Select = ""
    If RTrim$(KJS.識別子) = "51" Then
        gSL_Select = gSL_Select & "select"
        gSL_Select = gSL_Select & "  kms.検収日"
        gSL_Select = gSL_Select & ", kms.納入番号"
        gSL_Select = gSL_Select & ", kms.数量"
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  材料購入明細テーブル kms"
        If KJS.計算方法 = 1 Then
            gSL_Select = gSL_Select & " where kms.集約検収管理NO = " & KJS.検収管理NO
        Else
            gSL_Select = gSL_Select & " where kms.実績年月 = '" & Trim$(KSH.実績年月) & "'"
            gSL_Select = gSL_Select & "   and kms.契約先 = '" & RTrim$(KSH.契約先) & "'"
            gSL_Select = gSL_Select & "   and kms.識別子 = '" & RTrim$(KJS.識別子) & "'"
            gSL_Select = gSL_Select & "   and kms.材料管理番号 = '" & RTrim$(KJS.材料管理番号) & "'"
            gSL_Select = gSL_Select & "   and kms.単価 = '" & RTrim$(KJS.単価) & "'"
        End If
    Else
        gSL_Select = gSL_Select & "select"
        gSL_Select = gSL_Select & "  sms.検収日"
        gSL_Select = gSL_Select & ", sms.納入番号"
        gSL_Select = gSL_Select & ", sms.数量"
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  材料支給明細テーブル sms"
        If KJS.計算方法 = 1 Then
            gSL_Select = gSL_Select & " where sms.集約検収管理NO = " & KJS.検収管理NO
        Else
            gSL_Select = gSL_Select & " where sms.実績年月 = '" & Trim$(KSH.実績年月) & "'"
            gSL_Select = gSL_Select & "   and sms.契約先 = '" & RTrim$(KSH.契約先) & "'"
            gSL_Select = gSL_Select & "   and sms.識別子 = '" & RTrim$(KJS.識別子) & "'"
            gSL_Select = gSL_Select & "   and sms.材料管理番号 = '" & RTrim$(KJS.材料管理番号) & "'"
            gSL_Select = gSL_Select & "   and sms.単価 = '" & RTrim$(KJS.単価) & "'"
        End If
    End If
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  1"
    gSL_Select = gSL_Select & ", 2"
End Sub
'==================================================== *** 2008/06/16 End

'***********************************************************
'
'               Ｃ　Ｓ　Ｖ　出　力　処　理
'
'***********************************************************
'+++++++++++++++++++++++++++++++'
'+      出力ﾌｧｲﾙ存在確認
'+++++++++++++++++++++++++++++++'
Public Function COK0320CsvCheck() As Boolean
    COK0320CsvCheck = False
'
    On Error GoTo COK0320CsvCheck_Err
'
    If Dir(pCsvDir) <> "" Then
        DoEvents
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    " & pCsvDir & "　ファイルは既に存在します。            "
            .MB_MSG(3) = "    " & "上書きしますか？        "
            .MB_Title = "上書き確認"
            .MB_Button = Yes_No
            If Not .MBOX Then Exit Function
        End With
    End If
'
    COK0320CsvCheck = True
    Exit Function
COK0320CsvCheck_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------------+
'+      CSVﾌｧｲﾙ出力処理
'+----------------------------+
Private Function COK0320CsvPut(Optional DBNo As Integer = 0) As Boolean
    COK0320CsvPut = False
'
    On Error GoTo COK0320CsvPut_Err
'
    CisFrmCOK0320.MeterMax = CisDB.RecordCount

'   >> 出力ﾃｷｽﾄ OPEN
    OutFNo = FreeFile
    Open pCsvDir For Output As #OutFNo
'   >> HEAD部・編集＆出力
    If PrtFlagHyosi = True Then
        If PrtFlagSire = False And _
           PrtFlagSkyu = False Then Call CsvEdit_Head(OutStr): Print #OutFNo, OutStr
    Else
        Call CsvPut_Head
        OutSeq = 0
    End If
'   >> ｸﾘｱ
    OutCnt = 0
    Do Until Not KSH_RDSTS
    '   >> ｶｳﾝﾄ
        OutCnt = OutCnt + 1
    '   ## 表紙の処理 ##
        If PrtFlagHyosi = True Then
        '   >> Head部 編集＆出力
            If PrtFlagSire = True Or _
               PrtFlagSkyu = True Then Call CsvEdit_Head(OutStr): Print #OutFNo, OutStr
        '   >> Body部 編集＆出力
            Call CsvEdit_Body(OutStr, OutCnt)
            Print #OutFNo, OutStr
        End If
    '   ## 明細の処理 ##
        If PrtFlagSire = True Or _
           PrtFlagSkyu = True Then Call CsvPut_Jiseki
    '   >> 進行ﾒｰﾀｰ表示
        CisFrmCOK0320.MeterUpdate OutCnt

    '   >> 次のﾚｺｰﾄﾞを読み込む
        Call KSHReadNext(DBNo)
    Loop
    Call KSHClose(DBNo)
'   >> 出力ﾃｷｽﾄ Close
    Close #OutFNo

    SaveSetting appname:=gRegKey, section:="CSV", Key:=pCsvKey, setting:=pCsvDir
'
    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "    【  CSV出力 正常終了 】            "
        .MB_MSG(3) = "    出 力 先  = " & pCsvDir & "　　　　　"
        .MB_MSG(4) = "    出力件数 = " & Format(OutCnt, "#,##0") & " 件"
        .MB_Title = "【確認】"
        .MB_Button = OK
        .MBOX
    End With
'
    COK0320CsvPut = True
    Exit Function
COK0320CsvPut_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------------+
'+      Head部　編集処理
'+----------------------------+
Private Sub CsvEdit_Head(CsvStr As String)
    CsvStr = "№"
    CsvStr = CsvStr & ",実績年月"
    CsvStr = CsvStr & ",契約先,契約先名"
    CsvStr = CsvStr & ",請求書№"
    CsvStr = CsvStr & ",前月繰越"
    CsvStr = CsvStr & ",仕入金額,仕入消費税"
    CsvStr = CsvStr & ",相殺金額,相殺消費税"
    CsvStr = CsvStr & ",請求金額"
    CsvStr = CsvStr & ",支払予定日１,支払方法１,支払金額１"
    CsvStr = CsvStr & ",支払予定日２,支払方法２,支払金額２"
End Sub
'+----------------------------+
'+      Body部　編集処理
'+----------------------------+
Private Sub CsvEdit_Body(CsvStr As String, CsvCnt As Long)
    With KSH
        CsvStr = Format(CsvCnt, "###0")
        CsvStr = CsvStr & "," & Format(.実績年月, "@@@@/@@")
        CsvStr = CsvStr & "," & RTrim$(.契約先)
        CsvStr = CsvStr & "," & RTrim$(.契約先名)
        CsvStr = CsvStr & "," & .請求書NO
        CsvStr = CsvStr & "," & .前月繰越
        CsvStr = CsvStr & "," & .仕入金額
        CsvStr = CsvStr & "," & .仕入消費税
        CsvStr = CsvStr & "," & .相殺金額
        CsvStr = CsvStr & "," & .相殺消費税
        CsvStr = CsvStr & "," & .請求金額
        If RTrim$(.支払区分1) = "" Then
            CsvStr = CsvStr & ",,,"
        Else
            CsvStr = CsvStr & "," & Format(.支払日1, "@@@@/@@/@@")
            CsvStr = CsvStr & "," & RTrim$(.支払区分名1)
            CsvStr = CsvStr & "," & .支払金額1
        End If
        If RTrim$(.支払区分2) = "" Then
            CsvStr = CsvStr & ",,,"
        Else
            CsvStr = CsvStr & "," & Format(.支払日2, "@@@@/@@/@@")
            CsvStr = CsvStr & "," & RTrim$(.支払区分名2)
            CsvStr = CsvStr & "," & .支払金額2
        End If
    End With
End Sub
'******************************
'*      明細表の出力処理
'******************************
Private Function CsvPut_Jiseki(Optional DBNo As Integer = 1) As Boolean
    CsvPut_Jiseki = False
'
    Call SqlSelect_Jiseki
    If KJSRead(gSL_Select, , DBNo) Then
        If PrtFlagHyosi = True Then
            OutStr = ""
            Print #OutFNo, OutStr
            Call CsvPut_Head
            OutSeq = 0
        End If
'
        Do Until Not KJS_RDSTS
            Select Case RTrim$(KJS.識別子)
                Case "01", "21": Call CsvPut_Meisai
                Case Else:       Call CsvPut_Body3
            End Select
'
            Call KJSReadNext(DBNo)
        Loop
    End If
    Call KJSClose(DBNo)
    If PrtFlagHyosi = True Then
        OutStr = ""
        Print #OutFNo, OutStr
    End If
'
    CsvPut_Jiseki = True
End Function
'*---------------------*
'*  明細行の出力処理
'*---------------------*
Private Function CsvPut_Meisai(Optional DBNo As Integer = 2) As Boolean
    CsvPut_Meisai = False
'
    Call SqlSelect_Meisai
    If SJSRead(gSL_Select, , DBNo) Then
        Do Until Not SJS_RDSTS
            Call CsvPut_Body2
'
            Call SJSReadNext(DBNo)
        Loop
    End If
    Call SJSClose(DBNo)
'
    CsvPut_Meisai = True
End Function
'+----------------------------+
'+      Head部  出力処理
'+----------------------------+
Private Sub CsvPut_Head()
    OutStr = "№"
    OutStr = OutStr & ",実績年月"
    OutStr = OutStr & ",契約先,契約先名"
    OutStr = OutStr & ",識別子,識別子名"
    OutStr = OutStr & ",品番"
    OutStr = OutStr & ",検収日"
    OutStr = OutStr & ",納入番号"
    OutStr = OutStr & ",数量"
    OutStr = OutStr & ",単価"
    OutStr = OutStr & ",仮"
    OutStr = OutStr & ",金額"
    Print #OutFNo, OutStr
End Sub
'+----------------------------+
'+      Body部 2 出力処理
'+----------------------------+
Private Sub CsvPut_Body2()
    OutSeq = OutSeq + 1
    With KSH
        OutStr = Format(OutSeq, "###0")
        OutStr = OutStr & "," & Format(.実績年月, "@@@@/@@")
        OutStr = OutStr & "," & RTrim$(.契約先)
        OutStr = OutStr & "," & RTrim$(.契約先名)
    End With
    With KJS
        OutStr = OutStr & "," & RTrim$(.識別子)
        OutStr = OutStr & "," & RTrim$(.識別子名)
        OutStr = OutStr & "," & RTrim$(.表示品番)
        Select Case CisFun.Len2(SJS.検収日)
            Case 6:    OutStr = OutStr & "," & Format(SJS.検収日, "@@@@/@@")
            Case 8:    OutStr = OutStr & "," & Format(SJS.検収日, "@@@@/@@/@@")
            Case Else: OutStr = OutStr & "," & SJS.検収日
        End Select
        OutStr = OutStr & "," & RTrim$(SJS.納入番号)
        OutStr = OutStr & "," & SJS.数量
        OutStr = OutStr & "," & .単価
        Select Case RTrim$(.仮単価区分)
            Case "1":  OutStr = OutStr & ",*"
            Case Else: OutStr = OutStr & ","
        End Select
        OutStr = OutStr & "," & .金額
    End With
    Print #OutFNo, OutStr
End Sub
'+----------------------------+
'+      Body部 3 出力処理
'+----------------------------+
Private Sub CsvPut_Body3()
    OutSeq = OutSeq + 1
    With KSH
        OutStr = Format(OutSeq, "###0")
        OutStr = OutStr & "," & Format(.実績年月, "@@@@/@@")
        OutStr = OutStr & "," & RTrim$(.契約先)
        OutStr = OutStr & "," & RTrim$(.契約先名)
    End With
    With KJS
        OutStr = OutStr & "," & RTrim$(.識別子)
        OutStr = OutStr & "," & RTrim$(.識別子名)
        OutStr = OutStr & "," & RTrim$(.表示品番)
        Select Case CisFun.Len2(.検収日)
            Case 6:    OutStr = OutStr & "," & Format(.検収日, "@@@@/@@")
            Case 8:    OutStr = OutStr & "," & Format(.検収日, "@@@@/@@/@@")
            Case Else: OutStr = OutStr & "," & .検収日
        End Select
        OutStr = OutStr & "," & RTrim$(.納入番号)
        OutStr = OutStr & "," & .数量
        OutStr = OutStr & "," & .単価
        Select Case RTrim$(.仮単価区分)
            Case "1":  OutStr = OutStr & ",*"
            Case Else: OutStr = OutStr & ","
        End Select
        OutStr = OutStr & "," & .金額
    End With
    Print #OutFNo, OutStr
End Sub

'***********************************************************
'
'       帳　票　可　変　項　目　テ　ー　ブ　ル　処　理
'
'***********************************************************
'+-------------------------------+
'+  帳票可変項目テーブル追加・処理
'+-------------------------------+
Public Function DBInsert_CKT(Optional DBNo As Integer = 1) As Boolean
    Dim w_RenBan        As Integer
    Dim w_CNT           As String
'
    DBInsert_CKT = False
'
    On Error GoTo DBInsert_CKT_Err
    Call CisDB.DBTran(TransBegin, DBNo)     ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
'
    gSL_Select = ""
    gSL_Select = gSL_Select & "select *"
    gSL_Select = gSL_Select & "  from 帳票可変項目テーブル"
    gSL_Select = gSL_Select & " where 帳票番号 = 2"
    If Not CKTRead(gSL_Select, 1) Then
        Call ItemsClearCKT
        w_RenBan = 0
        Do Until w_RenBan = 4
            w_RenBan = w_RenBan + 1
            With CKT
                .帳票番号 = 2
                .連番 = w_RenBan
                .帳票名 = "支払金額連絡表"
                Select Case w_RenBan
                    Case 1: .可変項目 = "連絡先"
                    Case 2: .可変項目 = "　　　　担当者"
                    Case 3: .可変項目 = "　　　　　TEL:"
                    Case 4: .可変項目 = "　　　　　FAX:"
                End Select
                .作成者 = gTanto
            End With
            Call CKTInsert
        Loop
        '+---------------------+
        '+  ｺﾝﾎﾞﾎﾞｯｸｽ 初期設定
        '+---------------------+
        With CisFun
            .INI_Section = "帳票管理"
            .INI_FileID = gIniExe & "Combo.ini"
            .INI_Default = ""
            .INI_Key = "CNT"
            .GetIni
            w_CNT = Format(CisFun.Val2(.INI_String) + 1, "#")
            If w_CNT = "1" Then
               .INI_Key = "TYPE"
               .INI_String = "DATA"
               .PutIni
            End If
            .INI_Key = "CNT"
            .INI_String = w_CNT
            .PutIni
            .INI_Key = "DATA" & w_CNT
            .INI_String = "2,支払金額連絡表"
            .PutIni
        End With
    End If
'
    Call CisDB.DBTran(TransCommit, DBNo)    ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    On Error GoTo 0
'
    DBInsert_CKT = True
    Exit Function
DBInsert_CKT_Err:
    CisFun.ErrorBox
    End
End Function
'+-------------------------------+
'+  帳票可変項目テーブル索引処理
'+-------------------------------+
Private Sub DBSelect_CKT(Optional DBNo As Integer = 1)
    Dim w_RenBan        As Integer
'
'   >> ｸﾘｱ
    Erase pFreeMsg
'
    gSL_Select = ""
    gSL_Select = gSL_Select & "select 可変項目"
    gSL_Select = gSL_Select & "  from 帳票可変項目テーブル"
    gSL_Select = gSL_Select & " where 帳票番号 = 2"
    gSL_Select = gSL_Select & " order by 連番"
    If CKTRead(gSL_Select, , DBNo) Then
        w_RenBan = 0
        Do Until Not CKT_RDSTS
            w_RenBan = w_RenBan + 1
            pFreeMsg(w_RenBan) = RTrim$(CKT.可変項目)
            Call CKTReadNext(DBNo)
        Loop
    End If
    Call CKTClose(DBNo)
End Sub
'***********************************************************
'
'       帳　票　管　理　（名称マスタ）　処　理
'
'***********************************************************
'+-------------------------------+
'+  名称マスタ追加・処理
'+-------------------------------+
Public Function DBInsert_SYM(Optional DBNo As Integer = 1) As Boolean
    DBInsert_SYM = False
'
    On Error GoTo DBInsert_SYM_Err
    Call CisDB.DBTran(TransBegin, DBNo)     ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
'
'   >> 帳票管理
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "帳票管理"
        .区分桁数 = 10
        .英数字1 = ""
        .作成者 = gTanto
        .値 = "LOK0320A": .値名称 = "支払金額連絡表": Call SYMInitCreate(False)
        .値 = "LOK0320B": .値名称 = "仕入明細表":     Call SYMInitCreate
        .値 = "LOK0320C": .値名称 = "支給明細表":     Call SYMInitCreate
        .値 = "LOK0320D": .値名称 = "仕入明細表":     Call SYMInitCreate
        .値 = "LOK0320E": .値名称 = "支給明細表":     Call SYMInitCreate
    End With
'
    Call CisDB.DBTran(TransCommit, DBNo)    ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    On Error GoTo 0
'
    DBInsert_SYM = True
    Exit Function
DBInsert_SYM_Err:
    CisFun.ErrorBox
    End
End Function
'+-------------------------------+
'+  名称マスタ索引処理
'+-------------------------------+
Private Sub DBSelect_SYM(Optional DBNo As Integer = 1)
'   >> ｸﾘｱ
    Erase pPrtID
'
    gSL_Select = ""
    gSL_Select = gSL_Select & "select 値, 英数字1"
    gSL_Select = gSL_Select & "  from 名称マスタ"
    gSL_Select = gSL_Select & " where 区分名称 = '帳票管理'"
    gSL_Select = gSL_Select & "   and 区分タイプ = ''"
    gSL_Select = gSL_Select & "   and 値 like 'LOK0320%'"
    If SYMRead(gSL_Select, , DBNo) Then
        Do Until Not SYM_RDSTS
            Select Case RTrim$(SYM.値)
                Case "LOK0320A": pPrtID(1) = RTrim$(SYM.英数字1)
                Case "LOK0320B": pPrtID(2) = RTrim$(SYM.英数字1)
                Case "LOK0320C": pPrtID(3) = RTrim$(SYM.英数字1)
                Case "LOK0320D": pPrtID(4) = RTrim$(SYM.英数字1)
                Case "LOK0320E": pPrtID(5) = RTrim$(SYM.英数字1)
            End Select
            Call SYMReadNext(DBNo)
        Loop
    End If
    Call SYMClose(DBNo)
End Sub


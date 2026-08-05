Attribute VB_Name = "CXJ0130LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ノーかんばんエフ発行
'*
'*   モジュールID  : CXJ0130LB
'*
'*                  CXJ0130LPrint(PrintKB)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*   　　 : False : 印刷データなし
'*
'*   使用例:
'
'    If Not CXJ0130LPrint(0) Then Exit Sub
'
'*   作成日  :2004/03/24   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCXJ0130        As New CISFormPrint
Public CXJ0130_PrtCnt       As Long
Global CXJ0130_PrintKb      As Byte
Global CXJ0130_PrintOut     As Integer

Public Function CXJ0130LPrint(PrintKB As Byte, Optional PrintOut As Integer = 0) As Boolean
    CXJ0130LPrint = False

    Select Case PrintKB
        Case 0: CXJ0130_PrintKb = 51
        Case 1: CXJ0130_PrintKb = 151
    End Select
    CXJ0130_PrintOut = PrintOut

'   >> 出力データ存在確認
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  jyt.出荷予定日"
    gSL_Select = gSL_Select & ", jyt.出荷予定便"
    gSL_Select = gSL_Select & ", jyt.背番号"
    gSL_Select = gSL_Select & ", jyt.受入"
    gSL_Select = gSL_Select & ", jyt.収容数"
    gSL_Select = gSL_Select & ", jyt.枚数"
    gSL_Select = gSL_Select & ", jyt.端数"
    gSL_Select = gSL_Select & ", jyt.納入数"
    gSL_Select = gSL_Select & ", isnull(sym.値名称, '') 受注区分名"
    gSL_Select = gSL_Select & ", isnull(hnm.表示品番, jyt.品番) 表示品番"
    gSL_Select = gSL_Select & ", isnull(hnm.社内背番, '') 社内背番"
    gSL_Select = gSL_Select & ", isnull(hnm.品名, '') 品名"
    gSL_Select = gSL_Select & ", isnull(trm.略称, '') 納入先名"
    gSL_Select = gSL_Select & ", isnull(trm.工場名, '') 工場名"
'''    gSL_Select = gSL_Select & ", HT.出荷管理部署,ISNULL(TR2.略称,'') 出荷管理部署名" & vbCr
    gSL_Select = gSL_Select & ", HT.出荷管理部署,ISNULL(TR2.工場名,'') 出荷管理部署名" & vbCr
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  帳票出力テーブル cht"
    gSL_Select = gSL_Select & ", 受注テーブル jyt"
    gSL_Select = gSL_Select & "  left outer join 名称マスタ sym"
    gSL_Select = gSL_Select & "    on sym.区分名称 = '受注区分'"
    gSL_Select = gSL_Select & "   and sym.区分タイプ = ''"
    gSL_Select = gSL_Select & "   and sym.値 = jyt.受注区分"
    gSL_Select = gSL_Select & "  left outer join 品番マスタ hnm"
    gSL_Select = gSL_Select & "    on hnm.品番 = jyt.品番"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ trm"
    gSL_Select = gSL_Select & "    on trm.取引先CD = jyt.納入先"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 0"
    gSL_Select = gSL_Select & "  Left Outer Join 品番取引先マスタ HT" & vbCr
    gSL_Select = gSL_Select & "    ON  HT.品番 = jyt.品番" & vbCr
    gSL_Select = gSL_Select & "    And HT.取引先 = jyt.納入先" & vbCr
    gSL_Select = gSL_Select & "    And HT.受入 = jyt.受入"
    gSL_Select = gSL_Select & "    And HT.品目 = 0" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR2" & vbCr
    gSL_Select = gSL_Select & "    ON TR2.取引先区分 = 1" & vbCr
    gSL_Select = gSL_Select & "    And TR2.取引先CD = HT.出荷管理部署" & vbCr
    gSL_Select = gSL_Select & "    And TR2.社内区分 = 1" & vbCr
    gSL_Select = gSL_Select & " where cht.端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and cht.帳票区分 = " & CXJ0130_PrintKb
    gSL_Select = gSL_Select & "   and jyt.出荷指示書NO = cht.管理NO"
    gSL_Select = gSL_Select & "   and jyt.出荷指示書行NO = cht.管理行NO"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  jyt.出荷予定日"
    gSL_Select = gSL_Select & ", jyt.出荷予定便"
    gSL_Select = gSL_Select & ", jyt.納入先"
    gSL_Select = gSL_Select & ", jyt.受入"
    gSL_Select = gSL_Select & ", jyt.品番"
    
    Call Str_Monitor(gSL_Select)            ''  DEBUG
    
    If Not JYTRead(gSL_Select, 0, 1) Then
        Call JYTClose(1)
        CXJ0130LPrint = True
        Exit Function
    End If
'   >> 件数
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  isnull(sum(case when jyt.収容数 = 0 then 1 else jyt.枚数 end), 0) 件数"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  帳票出力テーブル cht"
    gSL_Select = gSL_Select & ", 受注テーブル jyt"
    gSL_Select = gSL_Select & " where cht.端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and cht.帳票区分 = " & CXJ0130_PrintKb
    gSL_Select = gSL_Select & "   and jyt.出荷指示書NO = cht.管理NO"
    gSL_Select = gSL_Select & "   and jyt.出荷指示書行NO = cht.管理行NO"
    With CisDB
        .SQL = gSL_Select
        .ReadStatus = True
        If .DBRead(1, 0) Then
            CXJ0130_PrtCnt = CisDB.Recordset(0)
        End If
    End With
'   <><><><><> ノーかんばんエフ　確認画面 <><><><><>
    If CXJ0130_PrintKb = 51 Then
        CXJ0130.Show vbModal
        Unload CXJ0130
        Set CXJ0130 = Nothing
    End If
'   <><><><><><><><><><><><><><><><><><><><><><><><>

    With CisPrtCXJ0130
        Set .iPrtForm = CXJ0130L
        Select Case CXJ0130_PrintOut
            Case 0: .iPreview = pv_Print
            Case 1: .iPreview = pv_Preview
            Case Else
                Call JYTClose(1)
                CXJ0130LPrint = True
                Exit Function
        End Select
        .dpMSG = "ノーかんばんエフ　印刷中"
        .dpMSGScreen = モードレス
        .PrintStart
    End With

    CXJ0130LPrint = True
End Function

'===================================================================================' 2019/08/02 ADD START
'+++++++++++++++++++++++++++++++++++++++++++++++++++    '++ デバッグ用です +++
'+                                                 +
'+  文字列をイミディエイトウィンドウに表示する     +
'+                                                 +
'+++++++++++++++++++++++++++++++++++++++++++++++++++
Public Sub Str_Monitor(ByVal wStr As String)
    Debug.Print vbCrLf
    Debug.Print Format(Now, "yyyy/MM/dd HH:mm:ss") + _
                " .................................................." + _
                vbCrLf + wStr
End Sub
'===================================================================================' 2019/08/02 ADD END





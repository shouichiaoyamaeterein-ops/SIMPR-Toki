Attribute VB_Name = "CKN0110LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   発注サイクル一覧表　発行
'*
'*   モジュールID  : CKN0110LB
'*
'*   作成日  :2004/05/11   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCKN0110        As New CISFormPrint
'印刷用ワーク
Public CKN0110_PrintKB      As Byte
Public CKN0110_PrintOut     As Integer
'
Public Function CKN011nLPrint(PrintKB As Byte, Optional PrintOut As Integer = 0) As Boolean
    CKN011nLPrint = False

' 出力データ存在確認
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  hcm.仕入先"
    gSL_Select = gSL_Select & ", hcm.サイクル"
    gSL_Select = gSL_Select & ", hcm.受入"
    gSL_Select = gSL_Select & ", hcm.振出便"
    gSL_Select = gSL_Select & ", hcm.昼夜区分"
    gSL_Select = gSL_Select & ", hcm.振出時間"
    gSL_Select = gSL_Select & ", hcm.発注加算日"
    gSL_Select = gSL_Select & ", hcm.発行時間"
    gSL_Select = gSL_Select & ", hcm.回収時間"
    gSL_Select = gSL_Select & ", hcm.入荷時間"
    gSL_Select = gSL_Select & ", trm.略称 仕入先名"
    gSL_Select = gSL_Select & ", trm.取引先"
    gSL_Select = gSL_Select & ", trm.工場CD"
    gSL_Select = gSL_Select & ", hcm.最終発注日"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  発注サイクルマスタ hcm"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ trm"
    gSL_Select = gSL_Select & "    on trm.取引先CD = hcm.仕入先"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 1"
    gSL_Select = gSL_Select & " where hcm.行識別 = 0"
    gSL_Select = gSL_Select & " order by"
    If PrintKB = 1 Then
        gSL_Select = gSL_Select & "  hcm.仕入先"
        gSL_Select = gSL_Select & ", hcm.サイクル"
        gSL_Select = gSL_Select & ", hcm.受入"
        gSL_Select = gSL_Select & ", hcm.振出便"
    Else
        gSL_Select = gSL_Select & "  hcm.発行時間"
        gSL_Select = gSL_Select & ", hcm.仕入先"
        gSL_Select = gSL_Select & ", hcm.サイクル"
        gSL_Select = gSL_Select & ", hcm.受入"
        gSL_Select = gSL_Select & ", hcm.振出便"
    End If
    If Not HCMRead(gSL_Select, 0, 1) Then
       Call HCMClose
       CKN011nLPrint = True
       Exit Function
    End If

    With CisPrtCKN0110
'       件数退避
        .cDataMaxCnt = CisDB.RecordCount
'
        If PrintKB = 1 Then
            Set .iPrtForm = CKN0111L
        Else
            Set .iPrtForm = CKN0112L
        End If
        .dpMSG = "発注サイクル一覧表　印刷中"
        If PrintOut = 0 Then
            .iPreview = pv_Print
        Else
            .iPreview = pv_Preview
        End If
        .dpMSGScreen = モードレス
        .PrintStart
    End With

    CKN011nLPrint = True
End Function


Attribute VB_Name = "CXH0065LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   発注予定一覧表　発行
'*
'*   モジュールID  : CXH0065LB
'*
'*   作成日  :2004/05/11   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCXH0065        As New CISFormPrint
'印刷用ワーク
Public CXH0065_PrintOut     As Integer
Global CXH0065_HYmd         As String
Public CXH0065_UkeirUmu     As Integer
'
Public Function CXH0065LPrint(Optional PrintOut As Integer = 0) As Boolean
    CXH0065LPrint = False
'----------------------------------------'
'   発行用データ作成
'----------------------------------------'
    With CisDB
        .SQL = "かんばん発注予定発行抽出"
        .StoadoCount = 3
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = Input用
        .ParaNo = 3: .ParaIO = OutPut用
    '
        .ParaNo = 1: .ParaValue = CXH0065_HYmd
        .ParaNo = 2: .ParaValue = 0
    '
        If Not .DBStored Then
           With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    かんばん発注予定発行抽出エラー            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
           End With
           Exit Function
        End If

        .ParaNo = 0: gCnt2 = .ParaValue
        If gCnt2 <> 0 Then
           With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    かんばん発注予定発行抽出エラー            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
           End With
           Exit Function
        End If
    End With
' 出力データ存在確認
    gSL_Select = " SELECT 発注時間,発注便,仕入先,ISNULL(TR.略称,'')仕入先名,サイクル,"
    If CXH0065_UkeirUmu <> 0 Then
        gSL_Select = gSL_Select & " 受入,"
    End If
    gSL_Select = gSL_Select & "  最終発注日,最終発注便,発行日,発注日,納入日,納入便,発行加算日,曜日発注"
    gSL_Select = gSL_Select & "  FROM 発注予定発行テーブル HY"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "     ON TR.取引先CD = HY.仕入先"
    gSL_Select = gSL_Select & "     AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & " Where 端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND 発注日 = 発注予定日"
    If CXH0065_UkeirUmu = 0 Then
        gSL_Select = gSL_Select & " GROUP BY 発注時間,発注便,仕入先,納入日,納入便,サイクル,最終発注日,最終発注便,発行日,発注日,納入日,納入便,発行加算日,TR.略称,曜日発注"
    End If
    gSL_Select = gSL_Select & " ORDER BY 発注時間,発注便,仕入先,納入日,納入便,サイクル"
    If Not HYHRead(gSL_Select, 0, 1) Then
       Call HYHClose
       CXH0065LPrint = True
       Exit Function
    End If

    With CisPrtCXH0065
'       件数退避
        .cDataMaxCnt = CisDB.RecordCount
'
        Set .iPrtForm = CXH0065L
        .dpMSG = "発注予定一覧表　印刷中"
        If PrintOut = 0 Then
            .iPreview = pv_Print
        Else
            .iPreview = pv_Preview
        End If
        .dpMSGScreen = モードレス
        .PrintStart
    End With

    CXH0065LPrint = True
End Function


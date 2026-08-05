Attribute VB_Name = "COK0010LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   支払一覧表
'*
'*   モジュールID  : COK0010LB
'*
'*                   COK0010LPrint(Ukbn, Hinbn, Torcd)
'*
'*   引数 : Ukbn     : 打切種別
'*   　　   Hinbn    : 品番
'*   　　   Torcd    : 取引先
'*   　　   PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  :
'*   　　   False : 印刷データ無し
'*
'*   使用例:
'
'    If Not COK0010LPrint Then Exit Sub
'
'*   作成日  :2004/02/12   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCOK0010        As New CISFormPrint
Public CisSimplePrint       As New CisSimplePrint
Public cpPrtCnt             As Long
Public cpPrtKb              As Byte
Public cpYm                 As String
Public cpGokei(8)           As Currency
'
Public Function COK0010LPrint() As Boolean
    COK0010LPrint = False

'   品番取引先マスタ
    gSL_Select = "SELECT SH.*,ISNULL(TR.略称,'') 契約先名,"
    gSL_Select = gSL_Select & " ISNULL(SY1.値名称,'') 支払区分名1,ISNULL(SY2.値名称,'') 支払区分名2"
    gSL_Select = gSL_Select & " FROM 買掛支払テーブル SH"
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 名称マスタ SY1"
    gSL_Select = gSL_Select & "    ON  SY1.区分名称 = '買掛支払区分'"
    gSL_Select = gSL_Select & "    AND SY1.区分タイプ = ''"
    gSL_Select = gSL_Select & "    AND SY1.値 = SH.支払区分1"
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 名称マスタ SY2"
    gSL_Select = gSL_Select & "    ON  SY2.区分名称 = '買掛支払区分'"
    gSL_Select = gSL_Select & "    AND SY2.区分タイプ = ''"
    gSL_Select = gSL_Select & "    AND SY2.値 = SH.支払区分2"
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "    ON  TR.取引先CD = SH.契約先"
    gSL_Select = gSL_Select & "    AND TR.取引先区分 = 1"
    gSL_Select = gSL_Select & " WHERE SH.実績年月 = '" & RTrim(cpYm) & "'"
    gSL_Select = gSL_Select & " ORDER BY SH.契約先"

    If Not KSHRead(gSL_Select, 0, 1) Then
        Call KSHClose
        Exit Function
    End If
    cpPrtCnt = CisDB.RecordCount
    With CisSimplePrint
        .MaxCount = CisDB.RecordCount
        .ProgressInit
    End With

    If cpPrtKb <> 2 Then
        With CisPrtCOK0010
             Set .iPrtForm = COK0010L
             .dpMSG = "支払一覧表　印刷中"
             Select Case cpPrtKb
                Case 0: .iPreview = pv_Print
                Case 1: .iPreview = pv_Preview
             End Select
             .dpMSGScreen = モーダル
             .PrintStart
        End With
    Else
        Call COK0010_CsvOut
    End If

    COK0010LPrint = True
End Function
'+----------------------------------------------------------+
'+         CSVファイル出力  処理                        ＋
'+----------------------------------------------------------+
Public Function COK0010_CsvOut() As Boolean
    COK0010_CsvOut = False
    Dim cFNo           As Integer
    Dim PrtHead         As String
    Dim CsvOutName      As String
    
'    pref_NCnt = 0                ' 現在発行ページ初期化
    On Error GoTo COK0010_CsvOut_Err
'------------------------'
'    出力ﾃｷｽﾄ OPEN
'------------------------'
'    CsvOutName = GetSetting(gRegKey, "CXH00050", "Csv", "")
    CsvOutName = CisSimplePrint.GetCsvFileName
    cFNo = FreeFile: gCnt1 = 0
    Open CsvOutName For Output As #cFNo

    On Error GoTo COK0010_CsvOut_Err
    
    PrtHead = "№,支払年月,契約先,契約先名,請求書№,前月繰越,仕入金額,仕入消費税,"
    PrtHead = PrtHead & "相殺金額,相殺消費税,請求金額,支払金額,繰越残高,支払区分1,支払日1,支払金額1,支払区分2,支払日2,支払金額2"
    On Error GoTo COK0010_CsvOut_Err
    Print #cFNo, PrtHead
' 明細を読み込んでセットする (レコード件数分)
    Do Until Not KSH_RDSTS
        With KSH
        '------------------------'
        '   レコード編集
        '------------------------'
            gCnt1 = gCnt1 + 1
            gStr = ""
            gStr = Format(gCnt1, "###0") & ","
            gStr = gStr & Left(.実績年月, 4) & "/" & Mid(.実績年月, 5, 2) & ","
            gStr = gStr & RTrim(.契約先) & ","
            gStr = gStr & RTrim(.契約先名) & ","
            gStr = gStr & RTrim(.請求書NO) & ","
            gStr = gStr & .前月繰越 & ","
            gStr = gStr & .仕入金額 & ","
            gStr = gStr & .仕入消費税 & ","
            gStr = gStr & .相殺金額 & ","
            gStr = gStr & .相殺消費税 & ","
            gStr = gStr & .請求金額 & ","
            gStr = gStr & .支払金額 & ","
            gStr = gStr & .繰越金額 & ","
            gStr = gStr & RTrim(.支払区分名1) & ","
            gStr = gStr & RTrim(.支払日1) & ","
            gStr = gStr & RTrim(.支払金額1) & ","
            gStr = gStr & RTrim(.支払区分名2) & ","
            gStr = gStr & RTrim(.支払日2) & ","
            gStr = gStr & RTrim(.支払金額2) & ","
        '------------------------'
        '   レコード出力
        '------------------------'
            On Error GoTo COK0010_CsvOut_Err
            Print #cFNo, gStr
        End With
        
        With CisSimplePrint
            .ProgressCnt = gCnt1
            .ProgressSet
        End With
        '次のﾚｺｰﾄﾞを読み込む
        Call KSHReadNext(1)
    Loop
'------------------------'
'   出力ﾃｷｽﾄ CLOSE
'------------------------'
    Close #cFNo

    COK0010_CsvOut = True

COK0010_CsvOut_Ed:
    Call KSHClose(1)
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "    【  CSV出力 正常終了 】            "
        .MB_MSG(3) = "    出 力 先  = " & RTrim(CsvOutName) & "　　　　　"
        .MB_MSG(4) = "    出力件数 = " & Format(gCnt1, "#,##0") & " 件"
        .MB_Title = "【確認】"
        .MB_Button = OK
        .MBOX
    End With
    Exit Function
COK0010_CsvOut_Err:
    CisFun.ErrorBox
    End
End Function




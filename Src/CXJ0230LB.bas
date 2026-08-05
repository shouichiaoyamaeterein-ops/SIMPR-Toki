Attribute VB_Name = "CXJ0230LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   出荷プルーフリスト発行
'*
'*   モジュールID  : CXJ0230LB
'*
'*                  CXJ0230LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not CXJ0230LPrint Then Exit Sub
'
'*   作成日  :2004/02/10   By CIS
'*   変更日  :2008/04/10   By CIS 入荷→出荷
'*---------------------------------------------------------------------------------
Public CisPrtCXJ0230        As New CISFormPrint
Public CisSimplePrint       As New CisSimplePrint
Public CXJ0230_PrtCnt       As Long
Global CXJ0230_PrtOutKb     As Byte

Public Function CXJ0230LPrint(Optional PrintOut As Integer = 0) As Boolean
    CXJ0230LPrint = False

    CXJ0230_PrtOutKb = PrintOut
' 出力データ存在確認
    gSL_Select = "SELECT NW.*,HN.表示品番,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.略称,'') 納入先名,ISNULL(TR1.工場名,'') 納入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.取引先,'') 納入先CD,ISNULL(TR1.工場CD,'') 納入先工場CD,"
    gSL_Select = gSL_Select & "  ISNULL(SY1.値名称,'') 受注区分名"
    gSL_Select = gSL_Select & " FROM 出荷画面入力ワーク NW"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "   ON HN.品番 = NW.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT"
    gSL_Select = gSL_Select & "   ON HT.品番 = NW.品番"
    gSL_Select = gSL_Select & "   AND HT.取引先 = NW.納入先"
    gSL_Select = gSL_Select & "   AND HT.受入 = NW.受入"
    gSL_Select = gSL_Select & "   AND HT.品目 = 0"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR1"
    gSL_Select = gSL_Select & "   ON TR1.取引先区分 = 0"
    gSL_Select = gSL_Select & "   AND TR1.取引先CD = NW.納入先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY1"
    gSL_Select = gSL_Select & "   ON SY1.区分名称 = '受注区分'"
    gSL_Select = gSL_Select & "   AND SY1.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY1.値 = NW.受注区分"
    gSL_Select = gSL_Select & " WHERE NW.処理端末 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND NW.処理区分 = 2"
    
    gSL_Select = gSL_Select & " Order By NW.納入先,NW.契約先,納入日,納入便,納入番号,納品書頁,納品書行"
    
    If Not SGWRead(gSL_Select, 0, 1) Then
       Call SGWClose
       CXJ0230LPrint = True
       Exit Function
    End If
    
    CXJ0230_PrtCnt = CisDB.RecordCount
    With CisSimplePrint
        .MaxCount = CisDB.RecordCount
        .ProgressInit
    End With
    
    If CXJ0230_PrtOutKb <> 2 Then
        With CisPrtCXJ0230
             Set .iPrtForm = CXJ0230L
             '*-*-*-*-* 2008/04/10 Update Start
             '.dpMSG = "入荷プルーフリスト　印刷中"
             .dpMSG = "出荷プルーフリスト　印刷中"
             '*-*-*-*-* 2008/04/10 Update End
             If PrintOut = 0 Then
               .iPreview = pv_Print
             Else
               .iPreview = pv_Preview
             End If
             .dpMSGScreen = モードレス
             .PrintStart
        End With
    Else
        Call CXJ0230_CsvOut
    End If
    
    CXJ0230LPrint = True
End Function
'+----------------------------------------------------------+
'+         CSVファイル出力  処理                        ＋
'+----------------------------------------------------------+
Public Function CXJ0230_CsvOut() As Boolean
    CXJ0230_CsvOut = False
    Dim cFNo           As Integer
    Dim PrtHead         As String
    Dim CsvOutName      As String
    Dim Key_Torcd       As String
    
'    pref_NCnt = 0                ' 現在発行ページ初期化
    On Error GoTo CXJ0230_CsvOut_Err
'------------------------'
'    出力ﾃｷｽﾄ OPEN
'------------------------'
    CsvOutName = CisSimplePrint.GetCsvFileName
    cFNo = FreeFile: gCnt1 = 0
    Open CsvOutName For Output As #cFNo

    '*-*-*-*-* 2008/04/10 Update Start
    'PrtHead = "№,納入先,納入先名,納入先工場,契約先,入荷日,入荷便,納入番号,受入,受注区分,"
    PrtHead = "№,納入先,納入先名,納入先工場,契約先,出荷日,出荷便,納入番号,受入,受注区分,"
    '*-*-*-*-* 2008/04/10 Update End
    PrtHead = PrtHead & "品番,背番号,収容数,箱数,端数,入荷数"
    On Error GoTo CXJ0230_CsvOut_Err
    Print #cFNo, PrtHead
    
    Key_Torcd = RTrim(SGW.納入先)
' 明細を読み込んでセットする (レコード件数分)
    Do Until Not SGW_RDSTS
        If RTrim(Key_Torcd) <> RTrim(SGW.納入先) Then
        '------------------------'
        '   ﾄﾚｲﾗｰ 出力
        '------------------------'
            On Error GoTo CXJ0230_CsvOut_Err
            gStr = ","
            Print #cFNo, gStr
            Print #cFNo, PrtHead
        End If
    '------------------------'
    '   レコード編集
    '------------------------'
        gCnt1 = gCnt1 + 1
        gStr = ""
        gStr = Format(gCnt1, "###0") & ","
        gStr = gStr & "'" & RTrim(SGW.納入先) & ","
        gStr = gStr & RTrim(SGW.納入先名) & ","
        gStr = gStr & RTrim(SGW.納入先工場名) & ","
        gStr = gStr & RTrim(SGW.契約先) & ","
        With CisFun
            .DateE_BefVal = RTrim(SGW.納入日)
            .DateEdit Slash
            gStr = gStr & RTrim(.DateE_AfVal) & ","
        End With
        gStr = gStr & RTrim(SGW.納入便) & ","
        gStr = gStr & RTrim(SGW.納入番号) & ","
        gStr = gStr & "'" & RTrim(SGW.受入) & ","
        gStr = gStr & RTrim(SGW.受注区分名) & ","
        gStr = gStr & RTrim(SGW.表示品番) & ","
        gStr = gStr & RTrim(SGW.背番号) & ","
        gStr = gStr & SGW.収容数 & ","
        gStr = gStr & SGW.枚数 & ","
        gStr = gStr & SGW.端数 & ","
        gStr = gStr & SGW.納入数 & ","
    '------------------------'
    '   レコード出力
    '------------------------'
        On Error GoTo CXJ0230_CsvOut_Err
        Print #cFNo, gStr
        
        Key_Torcd = RTrim(SGW.納入先)
        
        With CisSimplePrint
            .ProgressCnt = gCnt1
            .ProgressSet
        End With
        '次のﾚｺｰﾄﾞを読み込む
        Call SGWReadNext(1)
    Loop
'------------------------'
'   出力ﾃｷｽﾄ CLOSE
'------------------------'
    Close #cFNo

    CXJ0230_CsvOut = True

CXJ0230_CsvOut_Ed:
    Call SGWClose(1)
    
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
CXJ0230_CsvOut_Err:
    CisFun.ErrorBox
    End
End Function





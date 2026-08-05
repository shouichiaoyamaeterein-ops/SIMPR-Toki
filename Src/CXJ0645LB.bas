Attribute VB_Name = "CXJ0645LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   分納伝票読取リスト発行
'*
'*   モジュールID  : CXJ0645LB
'*
'*                  CXJ0645LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not CXJ0645LPrint Then Exit Sub
'
'*   作成日  :2009/12/22   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCXJ0645        As New CISFormPrint
Public CisSimplePrint       As New CisSimplePrint
Public CXJ0645_PrtCnt       As Long
Global CXJ0645_PrtOutKb     As Byte

Public Function CXJ0645LPrint(Optional PrintOut As Integer = 0) As Boolean
    CXJ0645LPrint = False

    CXJ0645_PrtOutKb = PrintOut
' 出力データ存在確認
    gSL_Select = "SELECT JB.*,HN.表示品番,"
    gSL_Select = gSL_Select & "  ISNULL(TR.正式名称,'') 納入先名"
    gSL_Select = gSL_Select & " FROM 受注SIM伝票分納ワーク JB"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "   ON HN.品番 = JB.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON TR.取引先区分 = 0"
    gSL_Select = gSL_Select & "   AND TR.取引先CD = JB.納入先"
    gSL_Select = gSL_Select & " WHERE JB.端末番号 = HOST_NAME()"
    'gSL_Select = gSL_Select & " ORDER BY 納入先,受入,納入日,納入便,納入番号,納品書頁,納品書行"
    
    If Not JSBRead(gSL_Select, 0, 1) Then
       Call JSBClose
       CXJ0645LPrint = True
       Exit Function
    End If
    
    CXJ0645_PrtCnt = CisDB.RecordCount
    With CisSimplePrint
        .MaxCount = CisDB.RecordCount
        .ProgressInit
    End With
    
    If CXJ0645_PrtOutKb <> 2 Then
        With CisPrtCXJ0645
             Set .iPrtForm = CXJ0645L
             .dpMSG = "分納伝票読取リスト　印刷中"
             If PrintOut = 0 Then
               .iPreview = pv_Print
             Else
               .iPreview = pv_Preview
             End If
             .dpMSGScreen = モードレス
             .PrintStart
        End With
    Else
        Call CXJ0645_CsvOut
    End If
    
    CXJ0645LPrint = True
End Function
'+----------------------------------------------------------+
'+         CSVファイル出力  処理                        ＋
'+----------------------------------------------------------+
Public Function CXJ0645_CsvOut() As Boolean
    CXJ0645_CsvOut = False
    Dim cFNo           As Integer
    Dim PrtHead         As String
    Dim CsvOutName      As String
    Dim Key_Torcd       As String
    
'    pref_NCnt = 0                ' 現在発行ページ初期化
    On Error GoTo CXJ0645_CsvOut_Err
'------------------------'
'    出力ﾃｷｽﾄ OPEN
'------------------------'
    CsvOutName = CisSimplePrint.GetCsvFileName
    cFNo = FreeFile: gCnt1 = 0
    Open CsvOutName For Output As #cFNo

    PrtHead = "№,納入先,納入先名,受入,納入日,納入便,納入番号,"
    PrtHead = PrtHead & "品番,納入数"
    On Error GoTo CXJ0645_CsvOut_Err
    Print #cFNo, PrtHead
    
    Key_Torcd = RTrim(JSB.納入先)
' 明細を読み込んでセットする (レコード件数分)
    Do Until Not JSB_RDSTS
        If RTrim(Key_Torcd) <> RTrim(JSB.納入先) Then
        '------------------------'
        '   ﾄﾚｲﾗｰ 出力
        '------------------------'
            On Error GoTo CXJ0645_CsvOut_Err
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
        gStr = gStr & "'" & RTrim(JSB.納入先) & ","
        gStr = gStr & RTrim(JSB.納入先名) & ","
        gStr = gStr & RTrim(JSB.受入) & ","
        With CisFun
            .DateE_BefVal = RTrim(JSB.納入日)
            .DateEdit Slash
            gStr = gStr & RTrim(.DateE_AfVal) & ","
        End With
        gStr = gStr & RTrim(JSB.納入便) & ","
        gStr = gStr & RTrim(JSB.納入番号) & ","
        gStr = gStr & RTrim(JSB.表示品番) & ","
        gStr = gStr & JSB.納入数 & ","
    '------------------------'
    '   レコード出力
    '------------------------'
        On Error GoTo CXJ0645_CsvOut_Err
        Print #cFNo, gStr
        
        Key_Torcd = RTrim(JSB.納入先)
        
        With CisSimplePrint
            .ProgressCnt = gCnt1
            .ProgressSet
        End With
        '次のﾚｺｰﾄﾞを読み込む
        Call JSBReadNext(1)
    Loop
'------------------------'
'   出力ﾃｷｽﾄ CLOSE
'------------------------'
    Close #cFNo

    CXJ0645_CsvOut = True

CXJ0645_CsvOut_Ed:
    Call JSBClose(1)
    
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
CXJ0645_CsvOut_Err:
    CisFun.ErrorBox
    End
End Function





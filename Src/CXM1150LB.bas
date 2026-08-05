Attribute VB_Name = "CXM1150LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   構成マスタ照会一覧 発行
'*
'*   モジュールID  : CXM1150LB
'*
'*                  CXM1150LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   GyakKB : 0-正展開 1-逆展開  HZKB : 0-品番 1-材料  TSKB : 0-直前 1-製品
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:
'*
'    If Not CXM1150LPrint Then Exit Sub
'
'*   作成日  :2009/03/13   By CIS
'*   変更日  :2011/04/04   By CIS  支給は表示しない
'*   変更日  :2012/07/03   By CIS  号補区分表示無
'
'*   作成日  :2019/07/30   By CIS
'
'*---------------------------------------------------------------------------------
Public CisPrtCXM1150        As New CISFormPrint
Public CisSimplePrint       As New CisSimplePrint
Public CXM1150_PrtCnt       As Long
Global CXM1150_PrtOutKb     As Byte
Global L_wsHKbn             As Byte                     '2009/11/10

Global mCSVDir              As String                   '2017/11/08


'Public Function CXM1150LPrint(Optional PrintOut As Integer = 0) As Boolean         '2009/11/10
Public Function CXM1150LPrint(Optional PrintOut As Integer = 0, Optional GyakKB As Byte = 0, Optional HZKB As Byte = 0, Optional TSKB As Byte = 0) As Boolean
    CXM1150LPrint = False

    '=========================================== 2009/11/10 Start
    L_wsHKbn = 0
    If HZKB <> 0 Then
        L_wsHKbn = 2
    Else
        If GyakKB <> 0 Then
            L_wsHKbn = 1
        End If
    End If
    '=========================================== 2009/11/10 End
    
    CXM1150_PrtOutKb = PrintOut
    
    
    GoTo SKIPLB11                       ' 2019/07/26 ADD
    
    
    
' 出力データ存在確認
    If L_wsHKbn = 0 Then                            '2009/11/10
        gSL_Select = ""
        gSL_Select = gSL_Select & "SELECT KW.*,"
        gSL_Select = gSL_Select & "       CASE WHEN HTM.号補区分 = ''  THEN ''"
        gSL_Select = gSL_Select & "            WHEN HTM.号補区分 = 'G' THEN '号口'"
        gSL_Select = gSL_Select & "            WHEN HTM.号補区分 = 'H' THEN '補給' END 号補区分,"
        gSL_Select = gSL_Select & "       CASE WHEN ISNULL(HTM1.取引先, '') <> '' THEN" & vbCrLf
        gSL_Select = gSL_Select & "                 CASE WHEN ISNULL(TRM1.略称, '') <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                      THEN HTM1.取引先 + ':' + TRM1.略称" & vbCrLf
        gSL_Select = gSL_Select & "                      ELSE HTM1.取引先 END" & vbCrLf
        gSL_Select = gSL_Select & "            WHEN ISNULL(HTM2.取引先, '') <> '' THEN" & vbCrLf
        gSL_Select = gSL_Select & "                 CASE WHEN ISNULL(TRM2.略称, '') <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                      THEN HTM2.取引先 + ':' + TRM2.略称" & vbCrLf
        gSL_Select = gSL_Select & "                      ELSE HTM2.取引先 END" & vbCrLf
        gSL_Select = gSL_Select & "            WHEN ISNULL(ZTM1.取引先, '') <> '' THEN" & vbCrLf
        gSL_Select = gSL_Select & "                 CASE WHEN ISNULL(TRM3.略称, '') <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                      THEN ZTM1.取引先 + ':' + TRM3.略称" & vbCrLf
        gSL_Select = gSL_Select & "                      ELSE ZTM1.取引先 END" & vbCrLf
        gSL_Select = gSL_Select & "            WHEN ISNULL(ZTM2.取引先, '') <> '' THEN" & vbCrLf
        gSL_Select = gSL_Select & "                 CASE WHEN ISNULL(TRM4.略称, '') <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                      THEN ZTM2.取引先 + ':' + TRM4.略称" & vbCrLf
        gSL_Select = gSL_Select & "                      ELSE ZTM2.取引先 END" & vbCrLf
        gSL_Select = gSL_Select & "            ELSE '' END 手配先" & vbCrLf
        
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert Start
        gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(HTM1.取引先, '') <> '' THEN HTM1.取引先" & vbCrLf
        gSL_Select = gSL_Select & "            WHEN ISNULL(HTM2.取引先, '') <> '' THEN HTM2.取引先" & vbCrLf
        gSL_Select = gSL_Select & "            WHEN ISNULL(ZTM1.取引先, '') <> '' THEN ZTM1.取引先" & vbCrLf
        gSL_Select = gSL_Select & "            WHEN ISNULL(ZTM2.取引先, '') <> '' THEN ZTM2.取引先" & vbCrLf
        gSL_Select = gSL_Select & "            ELSE '' END 取引先" & vbCrLf
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert End
        
        gSL_Select = gSL_Select & "  FROM 構成問合せワーク KW "
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(号補区分) 号補区分, 品番"
        gSL_Select = gSL_Select & "                          FROM 品番取引先マスタ"
        gSL_Select = gSL_Select & "                         GROUP BY 品番) HTM"
        gSL_Select = gSL_Select & "            ON HTM.品番 = KW.子品番"
        gSL_Select = gSL_Select & "           AND KW.区分 = ''"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(取引先) 取引先, 品番" & vbCrLf
        gSL_Select = gSL_Select & "                          FROM 品番取引先マスタ" & vbCrLf
'======================= 2011/04/04 Start
'        gSL_Select = gSL_Select & "                         WHERE 手配区分 <> '2' AND 品目 = 1 GROUP BY 品番) HTM1" & vbCrLf
        gSL_Select = gSL_Select & "                         WHERE ( 手配区分 <> '2' AND 手配区分 <> '4' )  AND 品目 = 1 GROUP BY 品番) HTM1" & vbCrLf
'======================= 2011/04/04 End
        gSL_Select = gSL_Select & "            ON HTM1.品番 = KW.子品番" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(取引先) 取引先, 品番" & vbCrLf
        gSL_Select = gSL_Select & "                          FROM 品番取引先マスタ" & vbCrLf
        gSL_Select = gSL_Select & "                         WHERE 手配区分 = '2' AND 品目 = 1 GROUP BY 品番) HTM2" & vbCrLf
        gSL_Select = gSL_Select & "            ON HTM2.品番 = KW.子品番" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(取引先) 取引先, 材料管理番号" & vbCrLf
        gSL_Select = gSL_Select & "                          FROM 材料取引先マスタ" & vbCrLf
'======================= 2011/04/04 Start
'        gSL_Select = gSL_Select & "                         WHERE 手配区分 <> '2' GROUP BY 材料管理番号) ZTM1" & vbCrLf
        gSL_Select = gSL_Select & "                         WHERE ( 手配区分 <> '2' AND 手配区分 <> '4' ) GROUP BY 材料管理番号) ZTM1" & vbCrLf
'======================= 2011/04/04 End
        gSL_Select = gSL_Select & "            ON CAST(ZTM1.材料管理番号 AS VARCHAR) = KW.材料管理番号" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(取引先) 取引先, 材料管理番号" & vbCrLf
        gSL_Select = gSL_Select & "                          FROM 材料取引先マスタ" & vbCrLf
        gSL_Select = gSL_Select & "                         WHERE 手配区分 = '2' GROUP BY 材料管理番号) ZTM2" & vbCrLf
        gSL_Select = gSL_Select & "            ON CAST(ZTM2.材料管理番号 AS VARCHAR) = KW.材料管理番号" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TRM1" & vbCrLf
        gSL_Select = gSL_Select & "            ON HTM1.取引先 = TRM1.取引先CD" & vbCrLf
        gSL_Select = gSL_Select & "           AND TRM1.取引先区分 = 1" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TRM2" & vbCrLf
        gSL_Select = gSL_Select & "            ON HTM2.取引先 = TRM2.取引先CD" & vbCrLf
        gSL_Select = gSL_Select & "           AND TRM2.取引先区分 = 1" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TRM3" & vbCrLf
        gSL_Select = gSL_Select & "            ON ZTM1.取引先 = TRM3.取引先CD" & vbCrLf
        gSL_Select = gSL_Select & "           AND TRM3.取引先区分 = 1" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TRM4" & vbCrLf
        gSL_Select = gSL_Select & "            ON ZTM2.取引先 = TRM4.取引先CD" & vbCrLf
        gSL_Select = gSL_Select & "           AND TRM4.取引先区分 = 1" & vbCrLf
        gSL_Select = gSL_Select & "   WHERE KW.処理端末 = HOST_NAME() "
        gSL_Select = gSL_Select & " ORDER BY KW.順位,KW.子品番 "
     End If                             '2009/11/10
    
    '================================================ 2009/11/10 Start
    If L_wsHKbn = 1 Then
        gSL_Select = "SELECT KW.親品番,KW.品名,KW.表示品番,KW.子品番,"
        gSL_Select = gSL_Select & "       MAX(CASE WHEN ISNULL(HTM1.取引先, '') <> '' THEN" & vbCrLf
        gSL_Select = gSL_Select & "                     CASE WHEN ISNULL(TRM1.略称, '') <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                          THEN HTM1.取引先 + ':' + TRM1.略称" & vbCrLf
        gSL_Select = gSL_Select & "                          ELSE HTM1.取引先 END" & vbCrLf
        gSL_Select = gSL_Select & "                WHEN ISNULL(HTM2.取引先, '') <> '' THEN" & vbCrLf
        gSL_Select = gSL_Select & "                     CASE WHEN ISNULL(TRM2.略称, '') <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                          THEN HTM2.取引先 + ':' + TRM2.略称" & vbCrLf
        gSL_Select = gSL_Select & "                          ELSE HTM2.取引先 END" & vbCrLf
        gSL_Select = gSL_Select & "                ELSE '' END) 手配先" & vbCrLf
        
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert Start
        gSL_Select = gSL_Select & "      ,MAX(CASE WHEN ISNULL(HTM1.取引先, '') <> '' THEN HTM1.取引先" & vbCrLf
        gSL_Select = gSL_Select & "                WHEN ISNULL(HTM2.取引先, '') <> '' THEN HTM2.取引先" & vbCrLf
        gSL_Select = gSL_Select & "                ELSE '' END) 取引先" & vbCrLf
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert End
        
        gSL_Select = gSL_Select & "  FROM 構成問合せワーク KW "
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(取引先) 取引先, 品番" & vbCrLf
        gSL_Select = gSL_Select & "                          FROM 品番取引先マスタ" & vbCrLf
'======================= 2011/04/04 Start
'        gSL_Select = gSL_Select & "                         WHERE 手配区分 <> '2' AND 品目 = 1 GROUP BY 品番) HTM1" & vbCrLf
        gSL_Select = gSL_Select & "                         WHERE ( 手配区分 <> '2' AND 手配区分 <> '4' ) AND 品目 = 1 GROUP BY 品番) HTM1" & vbCrLf
'======================= 2011/04/04 End
        gSL_Select = gSL_Select & "            ON HTM1.品番 = KW.親品番" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(取引先) 取引先, 品番" & vbCrLf
        gSL_Select = gSL_Select & "                          FROM 品番取引先マスタ" & vbCrLf
        gSL_Select = gSL_Select & "                         WHERE 手配区分 = '2' AND 品目 = 1 GROUP BY 品番) HTM2" & vbCrLf
        gSL_Select = gSL_Select & "            ON HTM2.品番 = KW.親品番" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TRM1" & vbCrLf
        gSL_Select = gSL_Select & "            ON HTM1.取引先 = TRM1.取引先CD" & vbCrLf
        gSL_Select = gSL_Select & "           AND TRM1.取引先区分 = 1" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TRM2" & vbCrLf
        gSL_Select = gSL_Select & "            ON HTM2.取引先 = TRM2.取引先CD" & vbCrLf
        gSL_Select = gSL_Select & "           AND TRM2.取引先区分 = 1" & vbCrLf
        gSL_Select = gSL_Select & " WHERE KW.処理端末 = HOST_NAME()"
        gSL_Select = gSL_Select & " GROUP BY KW.親品番,KW.品名,KW.表示品番,KW.子品番"
        gSL_Select = gSL_Select & " ORDER BY KW.親品番"
    End If
     
     If L_wsHKbn = 2 Then
        gSL_Select = "SELECT KW.親品番,KW.品名,KW.表示寸法,KW.表示品番,KW.子品番,"
        gSL_Select = gSL_Select & "       MAX(CASE WHEN ISNULL(HTM1.取引先, '') <> '' THEN" & vbCrLf
        gSL_Select = gSL_Select & "                     CASE WHEN ISNULL(TRM1.略称, '') <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                          THEN HTM1.取引先 + ':' + TRM1.略称" & vbCrLf
        gSL_Select = gSL_Select & "                          ELSE HTM1.取引先 END" & vbCrLf
        gSL_Select = gSL_Select & "                WHEN ISNULL(HTM2.取引先, '') <> '' THEN" & vbCrLf
        gSL_Select = gSL_Select & "                     CASE WHEN ISNULL(TRM2.略称, '') <> ''" & vbCrLf
        gSL_Select = gSL_Select & "                          THEN HTM2.取引先 + ':' + TRM2.略称" & vbCrLf
        gSL_Select = gSL_Select & "                          ELSE HTM2.取引先 END" & vbCrLf
        gSL_Select = gSL_Select & "                ELSE '' END) 手配先" & vbCrLf
        
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert Start
        gSL_Select = gSL_Select & "      ,MAX(CASE WHEN ISNULL(HTM1.取引先, '') <> '' THEN HTM1.取引先" & vbCrLf
        gSL_Select = gSL_Select & "                WHEN ISNULL(HTM2.取引先, '') <> '' THEN HTM2.取引先" & vbCrLf
        gSL_Select = gSL_Select & "                ELSE '' END) 取引先" & vbCrLf
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert End
        
        gSL_Select = gSL_Select & "  FROM 構成問合せワーク KW "
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(取引先) 取引先, 品番" & vbCrLf
        gSL_Select = gSL_Select & "                          FROM 品番取引先マスタ" & vbCrLf
'======================= 2011/04/04 Start
'        gSL_Select = gSL_Select & "                         WHERE 手配区分 <> '2' AND 品目 = 1 GROUP BY 品番) HTM1" & vbCrLf
        gSL_Select = gSL_Select & "                         WHERE ( 手配区分 <> '2' AND 手配区分 <> '4' ) AND 品目 = 1 GROUP BY 品番) HTM1" & vbCrLf
'======================= 2011/04/04 End
        gSL_Select = gSL_Select & "            ON HTM1.品番 = KW.親品番" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT MAX(取引先) 取引先, 品番" & vbCrLf
        gSL_Select = gSL_Select & "                          FROM 品番取引先マスタ" & vbCrLf
        gSL_Select = gSL_Select & "                         WHERE 手配区分 = '2' AND 品目 = 1 GROUP BY 品番) HTM2" & vbCrLf
        gSL_Select = gSL_Select & "            ON HTM2.品番 = KW.親品番" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TRM1" & vbCrLf
        gSL_Select = gSL_Select & "            ON HTM1.取引先 = TRM1.取引先CD" & vbCrLf
        gSL_Select = gSL_Select & "           AND TRM1.取引先区分 = 1" & vbCrLf
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TRM2" & vbCrLf
        gSL_Select = gSL_Select & "            ON HTM2.取引先 = TRM2.取引先CD" & vbCrLf
        gSL_Select = gSL_Select & "           AND TRM2.取引先区分 = 1" & vbCrLf
        gSL_Select = gSL_Select & " WHERE KW.処理端末 = HOST_NAME() "
        gSL_Select = gSL_Select & " GROUP BY KW.親品番,KW.品名,KW.表示寸法,KW.表示品番,KW.子品番"
        gSL_Select = gSL_Select & " ORDER BY KW.表示寸法,KW.親品番"
    End If
    '================================================ 2009/11/10 End
   Call Str_Monitor(gSL_Select)
   If Not KOLRead(gSL_Select, 0, 1) Then
       Call KOLClose
       CXM1150LPrint = True
       Exit Function
    End If
    
    
    
SKIPLB11:                                   ' 2019/07/26 ADD
    CisDB.RecordCount = iRdCnt              ' 2019/07/26 ADD

    CXM1150_PrtCnt = CisDB.RecordCount
    With CisSimplePrint
        .MaxCount = CisDB.RecordCount
        .ProgressInit
    End With
    
    
    If CXM1150_PrtOutKb <> 2 Then
        With CisPrtCXM1150
             Set .iPrtForm = CXM1150L
             .dpMSG = "構成マスタ照会一覧　印刷中"
             If PrintOut = 0 Then
               .iPreview = pv_Print
             Else
               .iPreview = pv_Preview
             End If
             .dpMSGScreen = モードレス
             .PrintStart
        End With
    Else
        Call CXM1150_CsvOut
    End If
    
    CXM1150LPrint = True
End Function
'+----------------------------------------------------------+
'+         CSVファイル出力  処理                        ＋
'+----------------------------------------------------------+
Public Function CXM1150_CsvOut() As Boolean
    CXM1150_CsvOut = False
    Dim cFNo           As Integer
    Dim PrtHead         As String
    Dim CsvOutName      As String
    Dim Key_Torcd       As String
    
'    pref_NCnt = 0                ' 現在発行ページ初期化
    On Error GoTo CXM1150_CsvOut_Err
'------------------------'
'    出力ﾃｷｽﾄ OPEN
'------------------------'
    CsvOutName = CisSimplePrint.GetCsvFileName
    
    cFNo = FreeFile: gCnt1 = 0
    Open CsvOutName For Output As #cFNo

    PrtHead = "レベル,手配先,区分,品番/材質,寸法,品名/材料名,背番号,収容数,使用数,号補区分,展開区分" 'Update 2011/03/01
    'PrtHead = "レベル,手配先,区分,品番/材質,寸法,品名/材料名,使用数,号補区分,展開区分"
    On Error GoTo CXM1150_CsvOut_Err
    Print #cFNo, PrtHead
    
    
    '===================================================================================================' 2019/07/26 ADD START
    Dim iCnt As Long
    For iCnt = 1 To iRdCnt
        With CisDB
            .SQL = "構成問合せ情報抽出"
            .StoadoCount = 4
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = Input用
            .ParaNo = 2: .ParaIO = Input用
            .ParaNo = 3: .ParaIO = Input用
            .ParaNo = 4: .ParaIO = Input用
            .ParaNo = 1: .ParaValue = wHinbn(iCnt)
            .ParaNo = 2: .ParaValue = bStaPara2
            .ParaNo = 3: .ParaValue = bStaPara3
            .ParaNo = 4: .ParaValue = bStaPara4
        
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
    
        If wPrtHKbn = 0 Then
            gSL_Select = "SELECT KW.* FROM 構成問合せワーク KW "
            gSL_Select = gSL_Select & " WHERE KW.処理端末 = HOST_NAME() "
            gSL_Select = gSL_Select & " ORDER BY KW.順位,KW.子品番 "
        Else
            gSL_Select = "SELECT KW.親品番,KW.品名,KW.表示品番 FROM 構成問合せワーク KW "
            gSL_Select = gSL_Select & " WHERE KW.処理端末 = HOST_NAME() "
            gSL_Select = gSL_Select & " GROUP BY KW.親品番,KW.品名,KW.表示品番"
            gSL_Select = gSL_Select & " ORDER BY KW.親品番"
        End If
        If Not KOLRead(gSL_Select) Then
            GoTo SKIPLB01
        End If
        Dim ChkCnt As Long
        ChkCnt = CisDB.RecordCount          ' 抽出件数(合計)退避
    '===================================================================================================' 2019/07/26 ADD END
    
' 明細を読み込んでセットする (レコード件数分)
    Do Until Not KOL_RDSTS
    '------------------------'
    '   レコード編集
    '------------------------'
        gCnt1 = gCnt1 + 1
        gStr = ""
        Select Case KOL.レベル
            Case 0: gStr = gStr & "0,"
            Case 1: gStr = gStr & "*1,"
            Case 2: gStr = gStr & "**2,"
            Case 3: gStr = gStr & "***3,"
            Case 4: gStr = gStr & "****4,"
            Case Else
                gStr = gStr & "*****" & Format(KOL.レベル, "##") & ","
        End Select
        gStr = gStr & RTrim(KOL.手配先) & ","
        If KOL.区分 <> "" Then
           gStr = gStr & "材料,"
        Else
           gStr = gStr & ","
        End If
        gStr = gStr & RTrim(KOL.表示品番) & ","
        gStr = gStr & RTrim(KOL.表示寸法) & ","
        gStr = gStr & RTrim(KOL.品名) & ","
        
        
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert Start
        Select Case L_wsHKbn
            Case 0
                If KOL.区分 = "" Then
                    gSL_Select = ""
                    gSL_Select = gSL_Select & "SELECT HT1.背番号"
                    gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(HT1.変更日,'') <> '' AND CONVERT(VARCHAR(8),GETDATE(),112) >= HT1.変更日 THEN HT1.収容数2 ELSE HT1.収容数1 END AS 収容数1"
                    gSL_Select = gSL_Select & "  FROM 品番取引先マスタ HT1"
                    gSL_Select = gSL_Select & " INNER JOIN (SELECT MAX(受入) 受入"
                    gSL_Select = gSL_Select & "               FROM 品番取引先マスタ"
                    gSL_Select = gSL_Select & "              WHERE 取引先 = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "                AND 品番   = '" & KOL.子品番 & "') HT2"
                    gSL_Select = gSL_Select & "       ON  HT1.取引先 = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "       AND HT1.品番   = '" & KOL.子品番 & "'"
                    gSL_Select = gSL_Select & "       AND HT1.受入   = HT2.受入"
                Else
                    gSL_Select = ""
                    gSL_Select = gSL_Select & "SELECT ZT1.背番号"
                    gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(ZT1.変更日,'') <> '' AND CONVERT(VARCHAR(8),GETDATE(),112) >= ZT1.変更日 THEN ZT1.収容数2 ELSE ZT1.収容数1 END AS 収容数1"
                    gSL_Select = gSL_Select & "  FROM 材料取引先マスタ ZT1"
                    gSL_Select = gSL_Select & " INNER JOIN (SELECT MAX(受入) 受入"
                    gSL_Select = gSL_Select & "               FROM 材料取引先マスタ"
                    gSL_Select = gSL_Select & "              WHERE 取引先       = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "                AND 材料管理番号 =  " & KOL.材料管理番号 & ") ZT2"
                    gSL_Select = gSL_Select & "       ON  ZT1.取引先       = '" & KOL.取引先 & "'"
                    gSL_Select = gSL_Select & "       AND ZT1.材料管理番号 = '" & KOL.材料管理番号 & "'"
                    gSL_Select = gSL_Select & "       AND ZT1.受入         = ZT2.受入"
                End If
            Case 1, 2
                gSL_Select = ""
                gSL_Select = gSL_Select & "SELECT HT1.背番号"
                gSL_Select = gSL_Select & "      ,CASE WHEN ISNULL(HT1.変更日,'') <> '' AND CONVERT(VARCHAR(8),GETDATE(),112) >= HT1.変更日 THEN HT1.収容数2 ELSE HT1.収容数1 END AS 収容数1"
                gSL_Select = gSL_Select & "  FROM 品番取引先マスタ HT1"
                gSL_Select = gSL_Select & " INNER JOIN (SELECT MAX(受入) 受入"
                gSL_Select = gSL_Select & "               FROM 品番取引先マスタ"
                gSL_Select = gSL_Select & "              WHERE 取引先 = '" & KOL.取引先 & "'"
                gSL_Select = gSL_Select & "                AND 品番   = '" & KOL.子品番 & "') HT2"
                gSL_Select = gSL_Select & "       ON  HT1.取引先 = '" & KOL.取引先 & "'"
                gSL_Select = gSL_Select & "       AND HT1.品番   = '" & KOL.子品番 & "'"
                gSL_Select = gSL_Select & "       AND HT1.受入   = HT2.受入"
        End Select
        Call Str_Monitor(gSL_Select)        ' 2019/07/26 ADD：確認のため追加
        
        Call HTMRead(gSL_Select, , 2)

        ' 背番号
        gStr = gStr & RTrim(HTM.背番号) & ","
        ' 収容数
        gStr = gStr & HTM.収容数1 & ","
        '☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★ 2011/02/28 Insert End
        
        
        
        gStr = gStr & KOL.構成数 & ","
'        gStr = gStr & RTrim(KOL.号補区分) & ","            '2012/07/03
        gStr = gStr & "," & ","                             '2012/07/03
        Select Case KOL.展開区分
            Case 0: gStr = gStr & ""
            Case 1: gStr = gStr & "まで対象"
            Case 2: gStr = gStr & "以降対象外"
            Case 3: gStr = gStr & "のみ対象外"
        End Select
        On Error GoTo CXM1150_CsvOut_Err
        Print #cFNo, gStr
        
        With CisSimplePrint
            .ProgressCnt = gCnt1
            .ProgressSet
        End With
        '次のﾚｺｰﾄﾞを読み込む
      ''Call KOLReadNext(1)         ' 2019/07/26 UPD：初回時に指定なしでNext時に番号を指定するとエラーになる
        Call KOLReadNext            ' 2019/07/26 UPD
    Loop
    
SKIPLB01:                           ' 2019/07/26 ADD
    Next                            ' 2019/07/26 ADD
    
'------------------------'
'   出力ﾃｷｽﾄ CLOSE
'------------------------'
    Close #cFNo

    CXM1150_CsvOut = True

CXM1150_CsvOut_Ed:
    Call KOLClose(1)
    
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
CXM1150_CsvOut_Err:
    CisFun.ErrorBox
    End
End Function





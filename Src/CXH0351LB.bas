Attribute VB_Name = "CXH0351LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   現品票 発行
'*
'*   モジュールID  : CXH0351LB.BAS
'*
'*                  CXH0351LPrint(Integer)
'*
'*   引数 : Integer  =  0  :新規発行
'*                      1  :再発行
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:
'    If Not CXH0351LPrint(1) Then Exit Sub
'
'*   作成日  :  2007/09/20   By CIS     CoReport8.5
'*   変更日  :  2008/04/11   By CIS     背番号８桁対応,車型５桁対応
'*   変更日  :
'*
'*---------------------------------------------------------------------------------
Dim CXH0351CoRpt        As New CoReports85
Public CXH0351_PrtCnt       As Long

Dim DP_HMoto        As String
Dim DP_JisyaCd      As String
Dim DP_JisyaKj      As String
Dim DP_JisyaNm      As String
Dim DP_JisyaKjNm    As String
Dim DP_TMaisu       As Long
Dim DP_GMaisu       As Long
Dim DP_PrtDate      As String
Dim DP_PrtTime      As String

Global CXH0351PrtOut    As Integer
Global CXH0351QR    As String
Global CXH0351_Printer      As String

Public mQRMeker351      As QRmaker  ' QR

Public Function CXH0351LPrint(Optional RePrint As Byte, Optional PrintOut As Integer, Optional PrtMsg As Boolean = False)
    CXH0351LPrint = False
    App.Title = "現品票"
    DP_PrtDate = Format(Now(), "yyyy/mm/dd")             '発行日
    DP_PrtTime = Format(Now(), "hh:nn")                  '発行時間
' 印刷進行画面表示ﾌﾗｸﾞｾｯﾄ
    PrtMsg = True   ' ActiveX使用の為、必ずTrue
    pref_MsgDisp = PrtMsg

    CXH0351PrtOut = gDemoPrt
'帳票区分 （新規発行 OR 再発行 )
    If RePrint = 0 Then
        RePrint = 10
    Else
        RePrint = 110
    End If
    CXH0351PrtOut = PrintOut
' 出力データ存在確認
    gSL_Select = "SELECT HC.*,HN.表示品番,HN.車型,HN.品名,HT.置場,"
    gSL_Select = gSL_Select & "  CASE WHEN HT.変更日 <> '' AND HT.変更日 + HT.変更便 <= HC.納入日 + HC.納入便"
    gSL_Select = gSL_Select & "  THEN 収容器2 ELSE 収容器1 END 収容器,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.略称,'') 仕入先名,ISNULL(TR1.工場名,'') 仕入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.取引先,'') 仕入先CD,ISNULL(TR1.工場CD,'') 仕入先工場CD,"
    gSL_Select = gSL_Select & "  ISNULL(TR2.略称,'') 納入先名,ISNULL(TR2.工場名,'') 納入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR2.取引先,'') 納入先CD,ISNULL(TR2.工場CD,'') 納入先工場CD,"
    gSL_Select = gSL_Select & "  ISNULL(SY1.値名称,'') 発注区分名,ISNULL(SY2.値名称,'') 手配区分名"
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 発注テーブル HC"
    gSL_Select = gSL_Select & "   ON HC.発注管理NO = CH.管理NO"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "   ON HN.品番 = HC.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT"
    gSL_Select = gSL_Select & "   ON HT.品番 = HC.品番"
    gSL_Select = gSL_Select & "   AND HT.取引先 = HC.仕入先"
    gSL_Select = gSL_Select & "   AND HT.受入 = HC.受入"
    gSL_Select = gSL_Select & "   AND HT.品目 = 1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR1"
    gSL_Select = gSL_Select & "   ON TR1.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR1.取引先CD = HC.仕入先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR2"
    gSL_Select = gSL_Select & "   ON TR2.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR2.取引先CD = HC.納入先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY1"
    gSL_Select = gSL_Select & "   ON SY1.区分名称 = '発注区分'"
    gSL_Select = gSL_Select & "   AND SY1.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY1.値 = HC.発注区分"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2"
    gSL_Select = gSL_Select & "   ON SY2.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "   AND SY2.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY2.値 = HC.手配区分"
    gSL_Select = gSL_Select & " WHERE 端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND 帳票区分 = " & RePrint
    
    gSL_Select = gSL_Select & " Order By"
    gSL_Select = gSL_Select & " HC.仕入先,HC.受入,HC.納入番号,HC.納品書頁,HC.納品書行"
    
    If Not HCTRead(gSL_Select, 0, 1) Then
       Call HCTClose(1)
       CXH0351LPrint = True
       Exit Function
    End If
    If CisDB.RecordCount = 0 Then
       CXH0351LPrint = True
       Exit Function
    End If
   '枚数獲得
    gSL_Select = "SELECT SUM(CASE WHEN HC.収容数 <> 0 THEN HC.枚数 ELSE 1 END) 件数 "
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 発注テーブル HC"
    gSL_Select = gSL_Select & "   ON HC.発注管理NO = CH.管理NO"
    gSL_Select = gSL_Select & " WHERE 端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND 帳票区分 = " & RePrint
    With CisDB
        .SQL = gSL_Select
        .ReadStatus = True
        If .DBRead(1, 0) Then
            CXH0351_PrtCnt = CisDB.Recordset(0)
        End If
    End With
    pref_TCnt = CXH0351_PrtCnt
    
'+++++++++++++++++++++++++++++++'
'+  印刷処理
'+++++++++++++++++++++++++++++++'
    Set pref_Form = CXH0351
    pref_Msg = "【 現品票 印刷 】"
    
    If pref_MsgDisp Then
        pref_Form.Show vbModal
        Unload pref_Form
        Set pref_Form = Nothing
    Else
        If pref_Form Is Nothing Then
           Set pref_Form = CXH0351
        End If
        Call CoPrintH0351
    End If
    
    CXH0351LPrint = True
End Function
'+----------------------------------------------------------+
'+                    印  刷  処  理                        ＋
'+----------------------------------------------------------+
Public Function CoPrintH0351() As Boolean
    CoPrintH0351 = False
'''''    Set CorForm = pref_Form
    
    pref_NCnt = 0                ' 現在発行ページ初期化
    On Error GoTo CoPrintH0351_Err

' システム情報(発行元)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 4
    Call IRNRead("", 1)
    DP_HMoto = Mid(IRN.oyian, 1, 7)
' システム情報(自社工場・機番)
    DP_JisyaCd = RTrim(gCompany)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 7
    Call IRNRead("", 1)
    DP_JisyaKj = Mid(IRN.oyian, 1, 1)
                                    
' 取引先マスタより名称を求める
    gSL_Select = "SELECT * FROM 取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(DP_JisyaCd) & RTrim(DP_JisyaKj) & "'"
    If TRMRead(gSL_Select, 1) Then
        DP_JisyaNm = TRM.略称
        DP_JisyaKjNm = TRM.工場名
    Else
        DP_JisyaNm = ""
        DP_JisyaKjNm = ""
    End If
    
'-----------------------------'
'   CoReport初期設定
'-----------------------------'
    With CXH0351CoRpt
        .SetDraw = pref_Form.CrDraw1
                
        .SetPrtFormCnt = 1
        .SetPrtForm = "LXH0351"
        .SetPrtPath = gIniExe & "Prt\"
'---------------------------------------------------------------------< 2006.02.14 ADD START >
'        .SetPaperSize = UserSize
'        .SetPaperWidth = 20000
'        .SetPaperLength = 10000
        .SetPaperSize = A4
'        .SetPaperWidth = 20000
'        .SetPaperLength = 10000
        .SetPrinter = CXH0351_Printer
'---------------------------------------------------------------------< 2006.02.14 ADD  END  >
        
        Select Case CXH0351PrtOut
            Case 0: .SetPrtOut = PrinterX
            Case 1: .SetPrtOut = Preview
                    .SetPreviewForm = frmPreview
            Case 2: .SetPrtOut = PDF
        End Select
        .SetPrtName = "現品票"
        If Not .CorOutStart Then Exit Function
        
    End With
   
' 明細を読み込んでｾｯﾄする(ﾚｺｰﾄﾞ件数分)
    Do Until Not HCT_RDSTS
        With HCT
            '収容数のある時は枚数、無い時は発注数
            If .収容数 <> 0 Then
                DP_TMaisu = .枚数
            Else
                DP_TMaisu = 1
            End If
        '*---------------------------------*
        '*  Body行　印刷（枚数分繰返し）
        '*---------------------------------*
            For DP_GMaisu = 1 To DP_TMaisu
        
                ' 項目転送
                If Not CoH0351Move Then GoTo CoPrintH0351_ED
                ' 印刷開始
                CXH0351CoRpt.CoPrintOut
                
                pref_NCnt = pref_NCnt + 1
                If pref_MsgDisp Then
                    pref_Cnt = Format(pref_NCnt, "###") & " / " & Format(pref_TCnt, "###")
                    DoEvents
                End If
            Next DP_GMaisu
        
        End With
        
        '次のﾚｺｰﾄﾞを読み込む
        Call HCTReadNext(1)
    Loop

    CoPrintH0351 = True

CoPrintH0351_ED:
    Call HCTClose(1)
' CoReport 終了処理
    If pref_TCnt <> 0 Then CXH0351CoRpt.CoOutEnd
    
'    Call PrintSet2("", True)
    Exit Function
CoPrintH0351_Err:
    CisFun.ErrorBox
'    Call PrintSet2("", True)
    End
End Function
'*------------------------------------------------------------------*
'*                                                                  *
'*      現品票　発行内容　項目転送                                    *
'*                                                                  *
'*------------------------------------------------------------------*
Private Function CoH0351Move() As Boolean
    Dim wBar            As String
    Dim wBarNaiyo       As String
    Dim wHYmd           As String
    Dim wEndGyo         As Integer
    
    CoH0351Move = True
    
    On Error GoTo CoH0351Move_Err
    With HCT
'   *------------------*
'   * 『HEAD』SET  *
'   *------------------*
'現品票タイトル
        If .発注区分 = 3 Then
            gStr = "現 品 票(臨時)"
        Else
            gStr = "現 品 票"
        End If
        If Not CXH0351CoRpt.CorData("現品票T", gStr, "S", 32) Then GoTo CoH0351Move_Err
'手配区分
        If Not CXH0351CoRpt.CorData("手配区分", .手配区分名, "", 4) Then GoTo CoH0351Move_Err
'自社名
        gStr = CisFun.Left2(DP_JisyaNm, 20)
        If Not CXH0351CoRpt.CorData("自社名", gStr, "", 20) Then GoTo CoH0351Move_Err
        gStr = CisFun.Left2(DP_JisyaKjNm, 20)
        If Not CXH0351CoRpt.CorData("自社工場", gStr, "", 20) Then GoTo CoH0351Move_Err
        
'仕入先名
        gStr = RTrim(.仕入先CD)
        If RTrim(.仕入先工場CD) <> "" Then
            gStr = gStr & "-" & RTrim(.仕入先工場CD)
        End If
        If Not CXH0351CoRpt.CorData("仕入先", gStr, "", 8) Then GoTo CoH0351Move_Err
        gStr = CisFun.Left2(.仕入先名, 20)
        If Not CXH0351CoRpt.CorData("仕入先名", gStr, "", 20) Then GoTo CoH0351Move_Err
        gStr = CisFun.Left2(.仕入先工場名, 20)
        If Not CXH0351CoRpt.CorData("仕入先工場", gStr, "", 20) Then GoTo CoH0351Move_Err
'納入先名
        If RTrim(.納入先) <> "" Then
            If Not CXH0351CoRpt.CorData("納入先T", "納入先", "", 6) Then GoTo CoH0351Move_Err
            gStr = RTrim(.納入先CD)
            If RTrim(.納入先工場CD) <> "" Then
                gStr = gStr & "-" & RTrim(.納入先工場CD)
            End If
            If Not CXH0351CoRpt.CorData("納入先", gStr, "", 8) Then GoTo CoH0351Move_Err
            gStr = CisFun.Left2(.納入先名, 20)
            If Not CXH0351CoRpt.CorData("納入先名", gStr, "", 20) Then GoTo CoH0351Move_Err
            gStr = CisFun.Left2(.納入先工場名, 20)
            If Not CXH0351CoRpt.CorData("納入先工場", gStr, "", 20) Then GoTo CoH0351Move_Err
        End If
'納入日・便
        gStr = Format(Mid(.納入日, 1, 4), "0000")
        If Not CXH0351CoRpt.CorData("納入年", gStr, "", 4) Then GoTo CoH0351Move_Err
        gStr = Format(Mid(.納入日, 5, 2), "#")
        If Not CXH0351CoRpt.CorData("納入月", gStr, "", 2) Then GoTo CoH0351Move_Err
        gStr = Format(Mid(.納入日, 7, 2), "#")
        If Not CXH0351CoRpt.CorData("納入日", gStr, "", 2) Then GoTo CoH0351Move_Err
        If Trim(.納入便) <> "" Then
           If Not CXH0351CoRpt.CorData("納入便", .納入便, "", 2) Then GoTo CoH0351Move_Err
           If Not CXH0351CoRpt.CorData("納入便T", "便", "", 2) Then GoTo CoH0351Move_Err
        End If
'頁
        gStr = "( " & Format(DP_GMaisu, "#,##0") & " / " & Format(DP_TMaisu, "#,##0") & " )"
        If Not CXH0351CoRpt.CorData("頁", gStr, "", 13) Then GoTo CoH0351Move_Err

'品番
        If CisFun.Len2(RTrim(.表示品番)) <= 22 Then
            If Not CXH0351CoRpt.CorData("品番2", .表示品番, "", 22) Then GoTo CoH0351Move_Err
        Else
            If Not CXH0351CoRpt.CorData("品番", .表示品番, "", 30) Then GoTo CoH0351Move_Err
        End If
'品名
        If Not CXH0351CoRpt.CorData("品名", .品名, "", 40) Then GoTo CoH0351Move_Err
'収容数
        If .収容数 <> 0 Then
            If Not CXH0351CoRpt.CorData("端数T", "端数", "", 4) Then GoTo CoH0351Move_Err
            gStr = Format(.収容数, "#,###")
            If Not CXH0351CoRpt.CorData("収容数", gStr, "", 6) Then GoTo CoH0351Move_Err
            
            If .端数 <> 0 And DP_TMaisu = DP_GMaisu Then
                gStr = Format(.端数, "#,###")
                If Not CXH0351CoRpt.CorData("端数", gStr, "", 6) Then GoTo CoH0351Move_Err
                gStr = "数量[ " & Format(.端数, "#,##0") & " / " & Format(.発注数, "#,##0") & " ]"
                If Not CXH0351CoRpt.CorData("頁数量", gStr, "", 21) Then GoTo CoH0351Move_Err
            Else
                gStr = "数量[ " & Format(.収容数, "#,##0") & " / " & Format(.発注数, "#,##0") & " ]"
                If Not CXH0351CoRpt.CorData("頁数量", gStr, "", 21) Then GoTo CoH0351Move_Err
            End If
        Else
            If Not CXH0351CoRpt.CorData("端数T", "数量", "", 4) Then GoTo CoH0351Move_Err
            
            gStr = Format(.発注数, "#,###")
            If Not CXH0351CoRpt.CorData("端数", gStr, "", 6) Then GoTo CoH0351Move_Err
            
            gStr = "数量[ " & Format(.発注数, "#,##0") & " / " & Format(.発注数, "#,##0") & " ]"
            If Not CXH0351CoRpt.CorData("頁数量", gStr, "", 21) Then GoTo CoH0351Move_Err
        End If
        
'置場
        If Not CXH0351CoRpt.CorData("置場", .置場, "", 10) Then GoTo CoH0351Move_Err
'背番号
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
        'If Not CXH0351CoRpt.CorData("背番号", .背番号, "", 5) Then GoTo CoH0351Move_Err
        If Not CXH0351CoRpt.CorData("背番号", .背番号, "", 8) Then GoTo CoH0351Move_Err
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
'収容器
        If Not CXH0351CoRpt.CorData("収容器", .収容器, "", 10) Then GoTo CoH0351Move_Err
'車型
        If Not CXH0351CoRpt.CorData("車型", .車型, "", 5) Then GoTo CoH0351Move_Err
        
'受入
        If Not CXH0351CoRpt.CorData("受入", .受入, "", 2) Then GoTo CoH0351Move_Err
        
'   *------------------*
'   * 『TAIL』SET  *
'   *------------------*
'発行日
        If Not CXH0351CoRpt.CorData("発行日", DP_PrtDate, "", 10) Then GoTo CoH0351Move_Err

'   *---------------------*
'   * 『QR』    SET       *
'   *---------------------*
        Call Denpyo_QREdit
        If Not CXH0351CoRpt.CorQRData("QR", CXH0351QR, "", 150) Then Exit Function
    End With

    Exit Function
CoH0351Move_Err:
    CoH0351Move = False
End Function
'*------------------------------------------------------------------*
'*                                                                  *
'*      現品票　QR　作成                                           *
'*                                                                  *
'*------------------------------------------------------------------*
Private Function Denpyo_QREdit() As Boolean
    Dim QREdit1      As String * 150
    
    QREdit1 = Space(150)
    
    Mid(QREdit1, 1, 7) = Left(DP_HMoto, 7)              '発行元
    Mid(QREdit1, 8, 5) = gCompany                       '自社コード
    Mid(QREdit1, 13, 1) = "S"                           'QR識別(R:臨時かんばん、S:現品票、H:端数ラベル)
    Mid(QREdit1, 14, 25) = HCT.品番                     '品番
    Mid(QREdit1, 39, 1) = "1"                           '品目(0:受注、1:発注)
    Mid(QREdit1, 40, 7) = HCT.仕入先                    '取引先コード
    Mid(QREdit1, 47, 2) = HCT.受入                      '取引先受入
    Mid(QREdit1, 49, 1) = Format(HCT.手配区分, "0")     '手配区分
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
'    Mid(QREdit1, 50, 5) = HCT.背番号                    '背番号
'    Mid(QREdit1, 55, 5) = HCT.納入番号                  '納入番号
'    Mid(QREdit1, 60, 8) = HCT.納入日                    '納入日
'    Mid(QREdit1, 68, 2) = HCT.納入便                    '納入便
'    Mid(QREdit1, 70, 7) = HCT.納入先                    '納入先コード
'    Mid(QREdit1, 77, 2) = HCT.納入先受入                '納入先受入
'    Mid(QREdit1, 79, 6) = Format(HCT.収容数, "000000")  '収容数
'    Mid(QREdit1, 85, 10) = HCT.収容器                   '収容器
'    If HCT.収容数 <> 0 Then
'        If HCT.端数 <> 0 And DP_TMaisu = DP_GMaisu Then
'            Mid(QREdit1, 95, 6) = Format(HCT.端数, "000000")  '数量
'        Else
'            Mid(QREdit1, 95, 6) = Format(HCT.収容数, "000000")  '数量
'        End If
'    Else
'        Mid(QREdit1, 95, 6) = Format(HCT.発注数, "000000")  '数量
'    End If
'    Mid(QREdit1, 101, 3) = Format(DP_GMaisu, "000")          '現枚数
'    Mid(QREdit1, 104, 3) = Format(DP_TMaisu, "000")        '総枚数
'    CisFun.DateE_AfVal = DP_PrtDate
'    CisFun.DateEditUn
'    Mid(QREdit1, 107, 8) = CisFun.DateE_BefVal          '発行日
'    Mid(QREdit1, 115, 5) = DP_PrtTime                   '発行時間
'    Mid(QREdit1, 120, 5) = HCT.サイクル                 'サイクル
'    Mid(QREdit1, 125, 10) = HCT.置場                    '置場
'    Mid(QREdit1, 135, 4) = HCT.車型                     '車型
'    Mid(QREdit1, 139, 1) = Format(HCT.発注区分, "0")    '発注区分
    Mid(QREdit1, 50, 8) = HCT.背番号                    '背番号
    Mid(QREdit1, 58, 5) = HCT.納入番号                  '納入番号
    Mid(QREdit1, 63, 8) = HCT.納入日                    '納入日
    Mid(QREdit1, 71, 2) = HCT.納入便                    '納入便
    Mid(QREdit1, 73, 7) = HCT.納入先                    '納入先コード
    Mid(QREdit1, 80, 2) = HCT.納入先受入                '納入先受入
    Mid(QREdit1, 82, 6) = Format(HCT.収容数, "000000")  '収容数
    Mid(QREdit1, 88, 10) = HCT.収容器                   '収容器
    If HCT.収容数 <> 0 Then
        If HCT.端数 <> 0 And DP_TMaisu = DP_GMaisu Then
            Mid(QREdit1, 98, 6) = Format(HCT.端数, "000000")  '数量
        Else
            Mid(QREdit1, 98, 6) = Format(HCT.収容数, "000000")  '数量
        End If
    Else
        Mid(QREdit1, 98, 6) = Format(HCT.発注数, "000000")  '数量
    End If
    Mid(QREdit1, 104, 3) = Format(DP_GMaisu, "000")          '現枚数
    Mid(QREdit1, 107, 3) = Format(DP_TMaisu, "000")        '総枚数
    CisFun.DateE_AfVal = DP_PrtDate
    CisFun.DateEditUn
    Mid(QREdit1, 110, 8) = CisFun.DateE_BefVal          '発行日
    Mid(QREdit1, 118, 5) = DP_PrtTime                   '発行時間
    Mid(QREdit1, 123, 5) = HCT.サイクル                 'サイクル
    Mid(QREdit1, 128, 10) = HCT.置場                    '置場
    Mid(QREdit1, 138, 5) = HCT.車型                     '車型
    Mid(QREdit1, 143, 1) = Format(HCT.発注区分, "0")    '発注区分
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
    
    CXH0351QR = RTrim(QREdit1)
    
End Function


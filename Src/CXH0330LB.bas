Attribute VB_Name = "CXH0330LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   発注明細書発行
'*
'*   モジュールID  : CXH0330LB
'*
'*                  CXH0330LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not CXH0330LPrint Then Exit Sub
'
'*   作成日  :2004/02/10   By CIS
'*   変更日  :2004/05/10   By CIS FAX送信機能追加
'*---------------------------------------------------------------------------------
'Public CisPrtCXH0330        As New CISFormPrint
Public CisPrtCXH0330        As CISFormPrint
Public CXH0330_PrtCnt       As Long
Global CXH0330_PrintKb      As Byte
Global CXH0330_PrintKb2     As Byte
Public CXH0330_PrtOut       As String       ' 2004/05/10 Add
Public CXH0330_FaxObj       As Object 'MfRemote
Public CXH0330_Frm          As CISFormContorl
Public CXH0330_PrtChg       As Boolean
Public CXH0330_Hikae        As Boolean
Public CXH0330_Kara         As Boolean
Public CXH0330_FaxConnect   As Boolean
Public CXH0330_FaxProc      As Boolean
Public CXH0330_FaxServer    As String
Public CXH0330_BasePrinter  As String

    'SendFrom(電話番号,氏名,会社名,所属,住所1,住所2,FAX番号,リモートID,携帯電話(30),メールアドレス(40))
Dim jw_Tel      As String
Dim jw_Name     As String
Dim jw_Cmpany   As String
Dim jw_Busyo    As String
Dim jw_Fax      As String
Dim Fax_Shigai  As String
Dim Fax_No      As String
Dim TextRec     As String
Dim wTorcd()    As String
Dim wTId        As Long
Dim blnReturn   As Boolean

Public Function CXH0330LPrint(PrintKB As Byte, Optional PrintOut As Integer = 0) As Boolean
    CXH0330LPrint = False

    If PrintKB = 0 Then
        CXH0330_PrintKb = 8
        CXH0330_PrintKb2 = 15
    Else
        CXH0330_PrintKb = 108
        CXH0330_PrintKb2 = 115
    End If
    
    If CXH0330_PrtOut <> "1" Then
       CXH0330_PrtOut = "1"
    End If
    ' Fax使用　有無
    CXH0330_FaxProc = False
    CXH0330_FaxConnect = False
    CXH0330_FaxServer = ""
    With CisFun
        .INI_Section = "Fax"
        .INI_FileID = gIniExe & gIniName
        .INI_Default = ""
        .INI_Key = .WSID
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "No"
           .PutIni
        End If
        gStr = Trim(.INI_String)
        If UCase(gStr) = "Y" Or UCase(gStr) = "YES" Then
           CXH0330_FaxProc = True
        End If
        .INI_Key = "Server"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "Server"
           .PutIni
        End If
        CXH0330_FaxServer = Trim(.INI_String)
    End With
' 出力データ存在確認
    gSL_Select = "SELECT HC.*,HN.表示品番,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.略称,'') 仕入先名,ISNULL(TR1.工場名,'') 仕入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.取引先,'') 仕入先CD,ISNULL(TR1.工場CD,'') 仕入先工場CD,"
    gSL_Select = gSL_Select & "  ISNULL(SY1.値名称,'') 発注区分名,ISNULL(SY2.値名称,'') 手配区分名"
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 発注テーブル HC"
    gSL_Select = gSL_Select & "   ON HC.明細書NO = CH.管理NO"
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
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY1"
    gSL_Select = gSL_Select & "   ON SY1.区分名称 = '発注区分'"
    gSL_Select = gSL_Select & "   AND SY1.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY1.値 = HC.発注区分"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2"
    gSL_Select = gSL_Select & "   ON SY2.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "   AND SY2.区分タイプ = ''"
    gSL_Select = gSL_Select & "   AND SY2.値 = HC.手配区分"
    gSL_Select = gSL_Select & " WHERE CH.端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND CH.帳票区分 = " & CXH0330_PrintKb
    If CXH0330_Hikae Then
       gSL_Select = gSL_Select & "   AND CH.管理文字 IN ('1','3','4')"
       gSL_Select = gSL_Select & "   AND HT.明細書発行 IN ('1','3','4')"
    Else
       If Not CXH0330_FaxProc Then
          gSL_Select = gSL_Select & "   AND CH.管理文字 IN ('1','3')"
          gSL_Select = gSL_Select & "   AND HT.明細書発行 IN ('1','3')"
       Else
          gSL_Select = gSL_Select & "   AND CH.管理文字 = '1'"
          gSL_Select = gSL_Select & "   AND HT.明細書発行 = '1'"
       End If
    End If
    
    gSL_Select = gSL_Select & " Order By HC.明細書NO,HC.納入番号,HC.納品書頁,HC.納品書行"
    
    If HCTRead(gSL_Select, 0, 1) Then
       CXH0330_PrtCnt = CisDB.RecordCount
       CXH0330_Kara = False
       Set CisPrtCXH0330 = New CISFormPrint
       With CisPrtCXH0330
            Set .iPrtForm = CXH0330L
            .dpMSG = "発注明細書　印刷中"
            If PrintOut = 0 Then
              .iPreview = pv_Print
            Else
              .iPreview = pv_Preview
            End If
            .dpMSGScreen = モードレス
            .PrintStart
       End With
       Set CisPrtCXH0330 = Nothing
    Else
       Call HCTClose
    End If
    
    gSL_Select = " SELECT HC.*,'' 表示品番,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.略称,'') 仕入先名,ISNULL(TR1.工場名,'') 仕入先工場名,"
    gSL_Select = gSL_Select & "  ISNULL(TR1.取引先,'') 仕入先CD,ISNULL(TR1.工場CD,'') 仕入先工場CD, "
    gSL_Select = gSL_Select & "  '' 発注区分名,'' 手配区分名"
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 空明細書テーブル HC"
    gSL_Select = gSL_Select & "   ON HC.明細書NO = CH.管理NO"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR1"
    gSL_Select = gSL_Select & "   ON TR1.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR1.取引先CD = HC.仕入先"
    gSL_Select = gSL_Select & " WHERE CH.端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND CH.帳票区分 = " & CXH0330_PrintKb2
    If CXH0330_Hikae Then
       gSL_Select = gSL_Select & "   AND CH.管理文字 IN ('1','3','4')"
       gSL_Select = gSL_Select & "   AND TR1.明細書発行 IN ('1','3','4')"
    Else
       If Not CXH0330_FaxProc Then
          gSL_Select = gSL_Select & "   AND CH.管理文字 IN ('1','3')"
          gSL_Select = gSL_Select & "   AND TR1.明細書発行 IN ('1','3')"
       Else
          gSL_Select = gSL_Select & "   AND CH.管理文字 = '1'"
          gSL_Select = gSL_Select & "   AND TR1.明細書発行 = '1'"
       End If
    End If
    gSL_Select = gSL_Select & " Order By HC.明細書NO,HC.仕入先,HC.受入"
    
    If HCTRead(gSL_Select, 0, 1) Then
       CXH0330_PrtCnt = CisDB.RecordCount
       CXH0330_Kara = True
       
       Set CisPrtCXH0330 = New CISFormPrint
       
       With CisPrtCXH0330
            Set .iPrtForm = CXH0330L
            .dpMSG = "空明細書　印刷中"
            If PrintOut = 0 Then
              .iPreview = pv_Print
            Else
              .iPreview = pv_Preview
            End If
            .dpMSGScreen = モードレス
            .PrintStart
       End With
       Set CisPrtCXH0330 = Nothing
    Else
       Call HCTClose
       CXH0330LPrint = True
       Exit Function
    End If
    
    CXH0330LPrint = True
End Function
'////////////////////////////////////////
'//     FAX送信
'////////////////////////////////////////
'+--------------------------------------+
'+      まいとーく確認
'+--------------------------------------+
Private Function MytalkChk() As Boolean
    MytalkChk = False
    On Error GoTo MytalkChk_Er
    Set CXH0330_FaxObj = CreateObject("MFREMOTE.MfRemoteCtrl.1")
    
    MytalkChk = True
    Exit Function
MytalkChk_Er:

End Function
'+-------------------------------------------------------------------------------------------+
'+      FAX送信
'+-------------------------------------------------------------------------------------------+
Public Function CXH0330LPrintFax(PrintKB As Byte, Optional PrintOut As Integer = 0) As Boolean
    
    CXH0330LPrintFax = False

    CXH0330_BasePrinter = Printer.DeviceName
    ' Fax使用　有無
    CXH0330_FaxProc = False
    CXH0330_FaxConnect = False
    CXH0330_FaxServer = ""
    With CisFun
        .INI_Section = "Fax"
        .INI_FileID = gIniExe & gIniName
        .INI_Default = ""
        .INI_Key = .WSID
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "No"
           .PutIni
        End If
        gStr = Trim(.INI_String)
        If UCase(gStr) = "Y" Or UCase(gStr) = "YES" Then
           CXH0330_FaxProc = True
        End If
        .INI_Key = "Server"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "Server"
           .PutIni
        End If
        CXH0330_FaxServer = Trim(.INI_String)
    
    End With
    
    If Not CXH0330_FaxProc Then
       CXH0330LPrintFax = True
       Exit Function
    End If
    
    If Not MytalkChk Then
       CXH0330_FaxProc = False
    End If
    
    If PrintKB = 0 Then
        CXH0330_PrintKb = 8
        CXH0330_PrintKb2 = 15
    Else
        CXH0330_PrintKb = 108
        CXH0330_PrintKb2 = 115
    End If
    If CXH0330_PrtOut <> "3" Then
       CXH0330_PrtOut = "3"
    End If
    
' システム情報(自社工場)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 7
    Call IRNRead("", 1, 1)
' 取引先マスタより名称を求める
    jw_Tel = ""
    jw_Name = ""
    jw_Cmpany = ""
    jw_Busyo = ""
    jw_Fax = ""
    gSL_Select = "SELECT * FROM 取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(gCompany) & Left(IRN.oyian, 1) & "'"
    If TRMRead(gSL_Select, 1, 1) Then
       jw_Tel = TRM.TEL1
       jw_Name = TRM.担当者
       jw_Cmpany = TRM.略称
       jw_Busyo = TRM.部署名1
       jw_Fax = TRM.FAX
       If Trim(jw_Busyo) = "" Then
          jw_Busyo = jw_Cmpany
       End If
    End If
    ' FAX出力
    If Not FaxPrint Then Exit Function
    ' FAX出力(空明細書)
    If Not FaxPrintKara Then Exit Function
    
    ' まいとーくFAX終了
    If CXH0330_FaxServer <> Trim(CisFun.WSID) Then
       On Error Resume Next
       CXH0330_FaxObj.AppExit
       On Error GoTo 0
    End If
    
'    If CXH0330_BasePrinter <> Printer.DeviceName Then
'       With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "          プリンタの変更が出来ませんでした。                  "
'            .MB_MSG(4) = "      通常使うプリンタの変更を手動にて行って下さい。          "
'            .MB_Title = ""
'            .MB_Button = OK
'            .MBOX
'        End With
'    End If
    CXH0330LPrintFax = True
End Function
'+------------------------------------+
'+      FAX出力（通常)
'+------------------------------------+
Private Function FaxPrint() As Boolean
    FaxPrint = False
' 出力データ存在確認
    gSL_Select = "Select 仕入先 "
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 発注テーブル HC"
    gSL_Select = gSL_Select & "   ON HC.明細書NO = CH.管理NO"
    gSL_Select = gSL_Select & " WHERE CH.端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND CH.帳票区分 = " & CXH0330_PrintKb
    gSL_Select = gSL_Select & "   AND CH.管理文字 = '3'"
    gSL_Select = gSL_Select & " Group by 仕入先"
    If Not HCTRead(gSL_Select, 0, 1) Then
       Call HCTClose
       FaxPrint = True
       Exit Function
    End If
    
    CXH0330_Hikae = False
' まいとーくFAX起動 & 初期設定
    CXH0330_PrtChg = False
    If CXH0330_FaxProc Then
        Dim i   As Integer
        Dim lpData  As String
        '「まいと～くFAX 」の起動
        i = CXH0330_FaxObj.Connect
        If i = 1 Or i = 2 Then
            CXH0330_FaxConnect = True
            'プリンタドライバ名を取得
            lpData = CXH0330_FaxObj.GetPrinterName
            '送付状を添付しない
            blnReturn = CXH0330_FaxObj.CoverPage("")
            
            '使用するプリンタドライバを「まいと～くFAX 」のものに
            For i = 0 To Printers.Count - 1
                If Printers(i).DeviceName = lpData Then
                    Set Printer = Printers(i)
                    Call CXH0330_Frm.PrintSet(Printer.DeviceName)
                    CXH0330_PrtChg = True
                    Exit For
                End If
            Next i
        End If
    End If
    
    ReDim wTorcd(CisDB.RecordCount)
    wTId = 0
    Do Until Not HCT_RDSTS
        wTorcd(wTId) = HCT.仕入先
        wTId = wTId + 1

        Call HCTReadNext(1)
    Loop

    Call HCTClose(1)
    For wTId = 0 To UBound(wTorcd)
        If Trim(wTorcd(wTId)) = "" Then Exit For

        If CXH0330_FaxConnect Then
            '念のためデータのクリア
            CXH0330_FaxObj.ClearSendFax
            '送付状を添付しない
            blnReturn = CXH0330_FaxObj.CoverPage("")
            
            Call FaxSendSet(Trim(wTorcd(wTId)))
        End If
        
        gSL_Select = "SELECT HC.*,HN.表示品番,"
        gSL_Select = gSL_Select & "  ISNULL(TR1.略称,'') 仕入先名,ISNULL(TR1.工場名,'') 仕入先工場名,"
        gSL_Select = gSL_Select & "  ISNULL(TR1.取引先,'') 仕入先CD,ISNULL(TR1.工場CD,'') 仕入先工場CD,"
        gSL_Select = gSL_Select & "  ISNULL(SY1.値名称,'') 発注区分名,ISNULL(SY2.値名称,'') 手配区分名"
        gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 発注テーブル HC"
        gSL_Select = gSL_Select & "   ON HC.明細書NO = CH.管理NO"
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
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY1"
        gSL_Select = gSL_Select & "   ON SY1.区分名称 = '発注区分'"
        gSL_Select = gSL_Select & "   AND SY1.区分タイプ = ''"
        gSL_Select = gSL_Select & "   AND SY1.値 = HC.発注区分"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2"
        gSL_Select = gSL_Select & "   ON SY2.区分名称 = '手配区分'"
        gSL_Select = gSL_Select & "   AND SY2.区分タイプ = ''"
        gSL_Select = gSL_Select & "   AND SY2.値 = HC.手配区分"
        gSL_Select = gSL_Select & " WHERE CH.端末番号 = HOST_NAME()"
        gSL_Select = gSL_Select & "   AND CH.帳票区分 = " & CXH0330_PrintKb
        gSL_Select = gSL_Select & "   AND CH.管理文字 = '3'"
        gSL_Select = gSL_Select & "   AND HT.明細書発行 = '3'"
        gSL_Select = gSL_Select & "   AND HC.仕入先 = '" & Trim(wTorcd(wTId)) & "'"
        
        gSL_Select = gSL_Select & " Order By HC.明細書NO,HC.納入番号,HC.納品書頁,HC.納品書行"
    
        If HCTRead(gSL_Select, 0, 1) Then
           CXH0330_PrtCnt = CisDB.RecordCount
           CXH0330_Kara = False
           Set CisPrtCXH0330 = New CISFormPrint
           With CisPrtCXH0330
                Set .iPrtForm = CXH0330L
                .dpMSG = "発注明細書(FAX)印刷中"
                .iPreview = pv_Print
                .dpMSGScreen = モードレス
                .PrintStart
           End With
           Set CisPrtCXH0330 = Nothing
        End If
    Next wTId
    
    If CXH0330_PrtChg Then
       Call CXH0330_Frm.PrintSet("")
    End If
    
    If CXH0330_FaxProc Then
        CXH0330_FaxObj.DisConnect
    End If
    
    FaxPrint = True

End Function
'+------------------------------------+
'+      FAX出力（空明細)
'+------------------------------------+
Private Function FaxPrintKara() As Boolean
    FaxPrintKara = False
' 出力データ存在確認
    gSL_Select = "Select 仕入先 "
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 空明細書テーブル HC"
    gSL_Select = gSL_Select & "   ON HC.明細書NO = CH.管理NO"
    gSL_Select = gSL_Select & " WHERE CH.端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND CH.帳票区分 = " & CXH0330_PrintKb2
    gSL_Select = gSL_Select & "   AND CH.管理文字 = '3'"
    gSL_Select = gSL_Select & " Group by 仕入先"
    If Not HCTRead(gSL_Select, 0, 1) Then
       Call HCTClose
       FaxPrintKara = True
       Exit Function
    End If
    
    CXH0330_Hikae = False
' まいとーくFAX起動 & 初期設定
    CXH0330_PrtChg = False
    If CXH0330_FaxProc Then
        Dim i   As Integer
        Dim lpData  As String
        '「まいと～くFAX 」の起動
        i = CXH0330_FaxObj.Connect
        If i = 1 Or i = 2 Then
            CXH0330_FaxConnect = True
            'プリンタドライバ名を取得
            lpData = CXH0330_FaxObj.GetPrinterName
            '送付状を添付しない
            blnReturn = CXH0330_FaxObj.CoverPage("")
            
            '使用するプリンタドライバを「まいと～くFAX 」のものに
            For i = 0 To Printers.Count - 1
                If Printers(i).DeviceName = lpData Then
                    Set Printer = Printers(i)
                    Call CXH0330_Frm.PrintSet(Printer.DeviceName)
                    CXH0330_PrtChg = True
                    Exit For
                End If
            Next i
        End If
    End If

    ReDim wTorcd(CisDB.RecordCount)
    wTId = 0
    Do Until Not HCT_RDSTS
        wTorcd(wTId) = HCT.仕入先
        wTId = wTId + 1

        Call HCTReadNext(1)
    Loop

    Call HCTClose(1)
    For wTId = 0 To UBound(wTorcd)
        If Trim(wTorcd(wTId)) = "" Then Exit For

        If CXH0330_FaxConnect Then
            '念のためデータのクリア
            CXH0330_FaxObj.ClearSendFax
            '送付状を添付しない
            blnReturn = CXH0330_FaxObj.CoverPage("")
            
            Call FaxSendSet(Trim(wTorcd(wTId)))
        End If
        
        gSL_Select = "SELECT HC.*,'' 表示品番,"
        gSL_Select = gSL_Select & "  ISNULL(TR1.略称,'') 仕入先名,ISNULL(TR1.工場名,'') 仕入先工場名,"
        gSL_Select = gSL_Select & "  ISNULL(TR1.取引先,'') 仕入先CD,ISNULL(TR1.工場CD,'') 仕入先工場CD "
        gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 空明細書テーブル HC"
        gSL_Select = gSL_Select & "   ON HC.明細書NO = CH.管理NO"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR1"
        gSL_Select = gSL_Select & "   ON TR1.取引先区分 = 1"
        gSL_Select = gSL_Select & "   AND TR1.取引先CD = HC.仕入先"
        gSL_Select = gSL_Select & " WHERE CH.端末番号 = HOST_NAME()"
        gSL_Select = gSL_Select & "   AND CH.帳票区分 = " & CXH0330_PrintKb2
        gSL_Select = gSL_Select & "   AND CH.管理文字 = '3'"
        gSL_Select = gSL_Select & "   AND TR1.明細書発行 = '3'"
        gSL_Select = gSL_Select & "   AND HC.仕入先 = '" & Trim(wTorcd(wTId)) & "'"
        
        gSL_Select = gSL_Select & " Order By HC.明細書NO,HC.仕入先,HC.受入"
    
        If HCTRead(gSL_Select, 0, 1) Then
           CXH0330_PrtCnt = CisDB.RecordCount
           CXH0330_Kara = True
           Set CisPrtCXH0330 = New CISFormPrint
           With CisPrtCXH0330
                Set .iPrtForm = CXH0330L
                .dpMSG = "空明細書(FAX)印刷中"
                .iPreview = pv_Print
                .dpMSGScreen = モードレス
                .PrintStart
           End With
           Set CisPrtCXH0330 = Nothing
        End If
    Next wTId
    
    If CXH0330_PrtChg Then
       Call CXH0330_Frm.PrintSet("")
    End If
    
    If CXH0330_FaxProc Then
        CXH0330_FaxObj.DisConnect
    End If

    FaxPrintKara = True

End Function
'+--------------------------------------+
'+      送信先情報  設定
'+--------------------------------------+
Private Sub FaxSendSet(Torcd As String)
    Dim intReceiptNumber    As Integer
    Dim blnReturn           As Boolean
    Dim Send_Torcd(4)       As String
    Dim Send_Tel(4)         As String
    Dim Send_Name(4)        As String
    Dim Send_Cmpany(4)      As String
    Dim Send_Busyo(4)       As String
    Dim Send_Fax(4)         As String

    If Not CXH0330_FaxConnect Then Exit Sub

    '自局情報を設定する設定になっている
    'SendFrom(電話番号,氏名,会社名,所属,住所1,住所2,FAX番号,リモートID,携帯電話(30),メールアドレス(40))
    blnReturn = CXH0330_FaxObj.SendFrom(jw_Tel, _
        jw_Name, jw_Cmpany, jw_Busyo, "", "", jw_Fax, "", "", "")
    '送信先情報を設定
    gSL_Select = "SELECT * FROM 取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(Torcd) & "'"
    If TRMRead(gSL_Select, 1, 1) Then
       If Trim(TRM.FAX) <> "" Then
          Send_Torcd(0) = Torcd
          Send_Tel(0) = TRM.TEL1
          Send_Name(0) = TRM.担当者
          Send_Cmpany(0) = TRM.略称
          Send_Busyo(0) = TRM.部署名1
          Send_Fax(0) = TRM.FAX
          If Trim(Send_Name(0)) = "" Then
             Send_Name(0) = Send_Cmpany(0)
          End If
       End If
       Send_Torcd(1) = TRM.FAX取引先1
       Send_Torcd(2) = TRM.FAX取引先2
       Send_Torcd(3) = TRM.FAX取引先3
    End If
    
    For gCnt1 = 1 To 3
        If Trim(Send_Torcd(gCnt1)) = "" Then Exit For
        gSL_Select = "SELECT * FROM 取引先マスタ"
        gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
        gSL_Select = gSL_Select & "   AND 取引先CD = '" & RTrim(Send_Torcd(gCnt1)) & "'"
        If TRMRead(gSL_Select, 1, 1) Then
           If Trim(TRM.FAX) <> "" Then
              Send_Torcd(gCnt1) = Send_Torcd(gCnt1)
              Send_Tel(gCnt1) = TRM.TEL1
              Send_Name(gCnt1) = TRM.担当者
              Send_Cmpany(gCnt1) = TRM.略称
              Send_Busyo(gCnt1) = TRM.部署名1
              Send_Fax(gCnt1) = TRM.FAX
              If Trim(Send_Name(gCnt1)) = "" Then
                 Send_Name(gCnt1) = Send_Cmpany(gCnt1)
              End If
           End If
        End If
    Next gCnt1
    
    For gCnt1 = 0 To 3
        If Trim(Send_Torcd(gCnt1)) = "" Then Exit For
        If Trim(Send_Fax(gCnt1)) <> "" Then
            'SendToV7(国番号,
            '         FAX市外局番,FAX番号,
            '         メールアドレス,インターネットFAX,氏名,敬称,
            '         会社名,所属,役職,
            '         電話市外局番,電話番号,Fコード)
            Call FaxNoEdit(Send_Fax(gCnt1))
            intReceiptNumber = CXH0330_FaxObj.SendToV7(81, _
                Fax_Shigai, Fax_No, _
                "", False, Send_Name(gCnt1), "", _
                Send_Cmpany(gCnt1), Send_Busyo(gCnt1), "", _
                "", "", "")
        End If
    Next gCnt1
End Sub
'+----------------------------------------+
'+      Fax番号編集  市外局番+FAX番号
'+----------------------------------------+
Private Function FaxNoEdit(FaxNo As String) As Boolean
    Dim wPos    As Long
    Dim wFax()  As String
    
    FaxNoEdit = False
    
    Fax_Shigai = ""
    Fax_No = ""
    
    If Trim(FaxNo) = "" Then Exit Function
    
    wFax = Split(FaxNo, "-")
    If UBound(wFax) > 1 Then
       If Trim(wFax(2)) <> "" Then
          Fax_Shigai = wFax(0)
          Fax_No = wFax(1) & "-" & wFax(2)
          FaxNoEdit = True
          Exit Function
       End If
    End If
    If UBound(wFax) > 0 Then
       If Trim(wFax(1)) <> "" Then
          Fax_No = wFax(0) & "-" & wFax(1)
          FaxNoEdit = True
          Exit Function
       End If
    End If
    Fax_No = wFax(0)
    FaxNoEdit = True
End Function
'+-------------------------------------------------------------------------------------------+
'+      e-Mail送信
'+-------------------------------------------------------------------------------------------+
Public Function CXH0330LPrintMail(PrintKB As Byte, Optional PrintOut As Integer = 0) As Boolean
    Dim wFPath          As String
    Dim wFName          As String
    Dim wFNo            As Integer
    Dim wTenpName       As String
    Dim wShiire         As String
    
    CXH0330LPrintMail = False

    If PrintKB = 0 Then
        CXH0330_PrintKb = 8
        CXH0330_PrintKb2 = 15
    Else
        CXH0330_PrintKb = 108
        CXH0330_PrintKb2 = 115
    End If
    
    If CXH0330_PrtOut <> "4" Then
       CXH0330_PrtOut = "4"
    End If
' 出力データ存在確認
    gSL_Select = "SELECT HC.*,Isnull(HN.表示品番,'') 表示品番,"
    gSL_Select = gSL_Select & "  ISNULL(HT.備考,'') 手配品番管理コード, "
    gSL_Select = gSL_Select & " 'A' 作成種別,"
    gSL_Select = gSL_Select & " CONVERT(VARCHAR(8),HC.作成日,112)+ CONVERT(VARCHAR(2),HC.作成日,114)+ SUBSTRING(CONVERT(VARCHAR(5),HC.作成日,114),4,2) 作成日付,"
    gSL_Select = gSL_Select & " CONVERT(VARCHAR(8),HC.更新日,112)+ CONVERT(VARCHAR(2),HC.更新日,114)+ SUBSTRING(CONVERT(VARCHAR(5),HC.更新日,114),4,2) 更新日付 "
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 発注テーブル HC"
    gSL_Select = gSL_Select & "   ON HC.明細書NO = CH.管理NO"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "   ON HN.品番 = HC.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT"
    gSL_Select = gSL_Select & "   ON HT.品番 = HC.品番"
    gSL_Select = gSL_Select & "   AND HT.取引先 = HC.仕入先"
    gSL_Select = gSL_Select & "   AND HT.受入 = HC.受入"
    gSL_Select = gSL_Select & "   AND HT.品目 = 1"
    gSL_Select = gSL_Select & " WHERE CH.端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND CH.帳票区分 = " & CXH0330_PrintKb
    gSL_Select = gSL_Select & "   AND CH.管理文字 = '4'"
    gSL_Select = gSL_Select & "   AND HT.明細書発行 = '4'"
    gSL_Select = gSL_Select & " Order By HC.仕入先,HC.明細書NO,HC.納入番号,HC.納品書頁,HC.納品書行"
    
    If HCTRead(gSL_Select, 0, 1) Then
       CXH0330_PrtCnt = CisDB.RecordCount
       CXH0330_Kara = False
       CXH0333.Show vbModal
       
       Unload CXH0333
       Set CXH0333 = Nothing
    Else
       Call HCTClose
    End If
    ' 空明細書
    gSL_Select = "SELECT HC.*,"
    gSL_Select = gSL_Select & " 'A' 作成種別,"
    gSL_Select = gSL_Select & " CONVERT(VARCHAR(8),HC.作成日,112)+ CONVERT(VARCHAR(2),HC.作成日,114)+ SUBSTRING(CONVERT(VARCHAR(5),HC.作成日,114),4,2) 作成日付 "
    gSL_Select = gSL_Select & " FROM 帳票出力テーブル CH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 空明細書テーブル HC"
    gSL_Select = gSL_Select & "   ON HC.明細書NO = CH.管理NO"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR1"
    gSL_Select = gSL_Select & "   ON TR1.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR1.取引先CD = HC.仕入先"
    gSL_Select = gSL_Select & " WHERE CH.端末番号 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND CH.帳票区分 = " & CXH0330_PrintKb2
    gSL_Select = gSL_Select & "   AND CH.管理文字 = '4'"
    gSL_Select = gSL_Select & "   AND TR1.明細書発行 = '4'"
    gSL_Select = gSL_Select & " Order By HC.仕入先,HC.受入"
    
    If HCTRead(gSL_Select, 0, 1) Then
       CXH0330_PrtCnt = CisDB.RecordCount
       CXH0330_Kara = True
        
       CXH0333.Show vbModal
       Unload CXH0333
       Set CXH0333 = Nothing
    Else
       Call HCTClose
    End If
'
    CXH0330LPrintMail = True
End Function

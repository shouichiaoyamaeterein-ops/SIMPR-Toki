Attribute VB_Name = "CKN1240LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   発注かんばん発行
'*
'*   モジュールID  : CKN1240LB.BAS
'*
'*                  CKN1240LPrint(Integer)
'*
'*   引数 : Integer  =  0  :新規発行
'*                      1  :再発行
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:
'    If Not CKN1240LPrint(1) Then Exit Sub
'
'*   作成日  : 2004/09/02   By CIS  ｾﾝﾄﾛﾆｸｽ対応
'*   変更日  :  2006/12/20  版数（1～99） 99を超えた時、1に戻る
'*   変更日  :  2008/04/11  By CIS 背番号８桁対応、車型５桁対応
'*---------------------------------------------------------------------------------
Global mBarNaiyo        As Boolean
Public mQRMeker         As QRmaker  ' QR
Global mQRTec           As Boolean

Global kn_Call          As String
Global kn_PrtKB         As Byte
Global Kn_PCnt          As Long

Global Kn_Rtn           As Boolean
Global Kn_Jisya         As String
Global Kn_JisyaNM       As String
Global Kn_JisyaKJNM     As String

Global kBas_Proc        As String

Global Kn_Page          As Boolean
Global Kn_PageS         As Long
Global Kn_PageE         As Long

Global Kn_Image         As String

Global s_Clip           As Boolean

Dim Brk_Torcd           As String
Dim Brk_Tornm           As String
Dim Brk_HMaisu          As Long

Dim mPrtSy              As String
Dim mShikiri            As Integer
Dim mShikiriAdd         As Long

Public Function CKN1240LPrint(Optional PrtSy As Integer = 0, Optional PrtMsg As Boolean = False)
    CKN1240LPrint = False
    App.Title = "発注かんばん"
' 印刷進行画面表示ﾌﾗｸﾞｾｯﾄ
    PrtMsg = True   ' ActiveX使用の為、必ずTrue
    pref_MsgDisp = PrtMsg
    
    If PrtSy = 0 Then
       mPrtSy = ""          ' 新規発行
    Else
       mPrtSy = "X"         ' 再発行
    End If
    
    gSL_Select = "Select * From かんばん発行テーブル "
    gSL_Select = gSL_Select & " Where かんばん種類 = " & kn_PrtKB
    gSL_Select = gSL_Select & "   And 発行済 = " & PrtSy
    gSL_Select = gSL_Select & "   And 発行FLG = '" & mPrtSy & "'"
    gSL_Select = gSL_Select & " Order By 発行種類,取引先,受入,背番号,品番,収容区分,発行NO"
    
    If Not KHTRead(gSL_Select) Then
        Exit Function
    End If

    pref_TCnt = CisDB.RecordCount
    
    Qr_Jisya = ""
    Qr_HMoto = ""
    ' 自社コード
    Qr_Jisya = RTrim(gCompany)
    '発行元
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 4
    Call IRNRead("", 1, 1)
    Qr_HMoto = RTrim(Mid(IRN.oyian, 1, 7))
    If Trim(Qr_HMoto) = "" Then Qr_HMoto = "SIMPR-A"
' かんばん設定　仕切かんばん
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 31
    Call IRNRead("", 1, 1)
    mShikiri = CisFun.Val2(Mid(IRN.oyian, 1, 1))
    mShikiriAdd = CisFun.Val2(Mid(IRN.oyian, 2, 2))
    If mShikiriAdd <> 0 Then
       mShikiriAdd = mShikiriAdd * 10
    End If
    ' 発行用設定項目
    With CisFun
        .INI_Section = "かんばん"
        .INI_FileID = gIniExe & gIniName
        ' QR発行種類
        .INI_Default = ""
        .INI_Key = "QRPrint"
        .GetIni
        gStr = RTrim(.INI_String)
        If Trim(gStr) = "" Then
           .INI_String = "QRMaker"
           .PutIni
        End If
        If UCase(gStr) = "TEC" Then
           mQRTec = True
        Else
           mQRTec = False
        End If
    End With
    
'+++++++++++++++++++++++++++++++'
'+  印刷処理
'+++++++++++++++++++++++++++++++'

    Set pref_Form = CKN1240
    Select Case kn_PrtKB
        Case 1: CorPrtForm = "LKN1241"
             pref_Msg = "【 工 程 か ん ば ん 印 刷 中 】"
        Case 2: CorPrtForm = "LKN1242"
             pref_Msg = "【 材 料 か ん ば ん 印 刷 中 】"
        Case 3: CorPrtForm = "LKN1243"
             pref_Msg = "【 受 注 裏 か ん ば ん 印 刷 中 】"
        Case Else
             CorPrtForm = "LKN1240"
             pref_Msg = "【 発 注 か ん ば ん 印 刷 中 】"
    End Select
    CorPrtForm2 = "LKN1249"     ' 仕切りかんばん
        
    pref_Form.Show vbModal
    Unload pref_Form
    Set pref_Form = Nothing
    
    CKN1240LPrint = True
End Function
'+----------------------------------------------------------+
'+                    印  刷  処  理                        ＋
'+----------------------------------------------------------+
Public Function CoPrintN1240() As Boolean
    Dim wPrtFlg     As Boolean
    
    CoPrintN1240 = False
    Set CorForm = pref_Form
    
    On Error GoTo CoPrintN1240_Err
    ' ｲﾒｰｼﾞ格納先取得
    With CisFun
        .INI_Section = "Option"
        .INI_Key = "Image":       .INI_Default = gIniExe
        .GetIni
        Kn_Image = RTrim(.INI_String)
    End With
    If Right(Kn_Image, 1) <> "\" Then Kn_Image = Kn_Image & "\"
    
    pref_NCnt = 0                ' 現在発行ページ初期化
        
    If Not CoReportProcN1240("O") Then GoTo CoPrintN1240_ED

    Brk_Torcd = KHT.取引先
    Brk_Tornm = KHT.取引先名
    Brk_HMaisu = 0
    
'ﾊﾞｰｺｰﾄﾞ初期ｾｯﾄ
'    Call GetBarInfo("JISYADEN")
'    Erase CoBarInf
'    CoBarInf(0) = -1
'    CoBarInf(1) = CowBar_Info(0)
'    CoBarInf(2) = CowBar_Info(1)
'    CoBarInf(3) = CowBar_Info(2)
'    CoBarInf(4) = CowBar_Info(3)
'    CoBarInf(5) = CowBar_Info(4)
'    CorRtn = CoSetBarInfo(CoBarInf(0), WorkBuffer(0))
    
' 明細を読み込んでｾｯﾄする(ﾚｺｰﾄﾞ件数分)
    
    gLong = 0
    Do Until Not KHT_RDSTS
'        DoEvents
'        wPrtFlg = True
        
        If mPrtSy = "" Then
           If kn_PrtKB = 0 Then
              gSL_Select = "Select * From かんばん流動管理テーブル"
              gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(KHT.取引先) & "'"
              gSL_Select = gSL_Select & " And   受入 = '" & RTrim(KHT.受入) & "'"
              gSL_Select = gSL_Select & " And   品番 = '" & RTrim(KHT.品番) & "'"
              gSL_Select = gSL_Select & " And   発行NO = " & KHT.発行NO & ""
              If Not KRTRead(gSL_Select, 1, 1) Then
                 Call ItemsClearKRT
                 With KRT
                      .取引先 = KHT.取引先
                      .受入 = KHT.受入
                      .品番 = KHT.品番
                      .発行NO = KHT.発行NO
                      .版数 = 1
                      .発行区分 = 0
                      .収容区分 = KHT.収容区分
                 End With
                 Call KRTInsert(1)
               Else
                 With KRT
                      If .版数 + 1 > 99 Then
                        .版数 = 1
                      Else
                        .版数 = .版数 + 1
                      End If
                      .発行区分 = 0
                 End With
                 Call KRTUpdate(1)
               End If
           End If
           If kn_PrtKB = 2 Then
              gSL_Select = "Select * From かんばん材料管理テーブル"
              gSL_Select = gSL_Select & " Where 材料管理番号 = " & KHT.材料管理番号
              gSL_Select = gSL_Select & " And   取引先 = '" & RTrim(KHT.取引先) & "'"
              gSL_Select = gSL_Select & " And   受入 = '" & RTrim(KHT.受入) & "'"
              gSL_Select = gSL_Select & " And   発行NO = " & KHT.発行NO & ""
              If Not KZKRead(gSL_Select, 1, 1) Then
                 Call ItemsClearKZK
                 With KZK
                      .材料管理番号 = KHT.材料管理番号
                      .取引先 = KHT.取引先
                      .受入 = KHT.受入
                      .発行NO = KHT.発行NO
                      .版数 = 1
                      .発行区分 = 0
                      .収容区分 = KHT.収容区分
                 End With
                 Call KZKInsert(1)
              Else
                 With KZK
                      If .版数 + 1 > 99 Then
                        .版数 = 1
                      Else
                        .版数 = .版数 + 1
                      End If
                      .発行区分 = 0
                 End With
                 Call KZKUpdate(1)
              End If
          End If
        End If
        ' QRエリア転送
        Call QRAreaClear
        With QRArea
            .品番 = KHT.品番
            .品目 = KHT.品目
            .取引先 = KHT.取引先
            .受入 = KHT.受入
            .手配区分 = KHT.手配区分
            .版数 = KRT.版数
            .発行日 = Format(Now(), "YYYYMMDD")
            .背番号 = KHT.背番号
            .連番 = KHT.発行NO
            .収容数 = KHT.収容数
            .サイクル = KHT.サイクル
            .納入先 = KHT.納入先
            .納入先受入 = KHT.納入先受入
            .工程 = KHT.工程
            .次工程 = KHT.次工程
            .編集品番 = KHT.表示品番
            .手配区分名 = KHT.手配区分名
            .アドレス = KHT.アドレス
            If kn_PrtKB = 2 Then
               .品目 = 2
               .品番 = KHT.材質
               .板厚 = KHT.板厚
               .幅 = KHT.幅
               .長さ = KHT.長さ
               .寸法区分 = KHT.寸法区分
               .材料管理番号 = KHT.材料管理番号
               .材料収容数 = KHT.収容数
            End If
        End With
        ' 発行
        ' 項目転送
        If Not CoN1240Move Then GoTo CoPrintN1240_ED
        ' 印刷開始
        If Not CoReportProcN1240("P") Then GoTo CoPrintN1240_ED
        
        With CisDB
            .ConnectNo = 1
            .SQL = "Update かんばん発行テーブル Set 発行済 = 1"
            If mPrtSy = "" Then
                .SQL = .SQL & " ,版数 = " & KRT.版数
            End If
            .SQL = .SQL & " Where かんばん種類 = " & kn_PrtKB
            .SQL = .SQL & " And   背番号 = '" & RTrim(KHT.背番号) & "'"
            .SQL = .SQL & " And   品番 = '" & RTrim(KHT.品番) & "'"
            .SQL = .SQL & " And   取引先 = '" & RTrim(KHT.取引先) & "'"
            .SQL = .SQL & " And   受入 = '" & RTrim(KHT.受入) & "'"
            .SQL = .SQL & " And   発行NO = " & KHT.発行NO
            .DBExec
        End With
        ' 進捗ゲージ更新
        pref_NCnt = pref_NCnt + 1
        gLong = gLong + 1
        If pref_MsgDisp Then
            pref_Cnt = Format(pref_NCnt, "###") & " / " & Format(pref_TCnt, "###")
            DoEvents
            '中断処理
            If pref_Cansel_Flg Then
                If MsgBox("処理を中断します。よろしいですか？", 33, "印刷") = 1 Then
                   GoTo CoPrintN1240_ED
                Else
                   pref_Cansel_Flg = False
                End If
            End If
        End If
        
        Call KHTReadNext
        
        ' 仕切りかんばん発行
        If mShikiri = 1 Then
           If Not PrintShikiri Then GoTo CoPrintN1240_ED
        End If
    Loop

    CoPrintN1240 = True

CoPrintN1240_ED:
    Call KHTClose
' CoReport 終了処理
    If pref_TCnt <> 0 And mShikiri = 0 Then Call CoReportProcN1240("E")
    Exit Function
CoPrintN1240_Err:
    gStr = CisDB.ErrCode & CisDB.ErrMsg
    CisFun.ErrorBox
    End
End Function
'*------------------------------------------------------------------*
'*                                                                  *
'*      かんばん　発行内容　項目転送                                *
'*                                                                  *
'*------------------------------------------------------------------*
Private Function CoN1240Move() As Boolean
    Dim wBar            As String
    Dim wBarNaiyo       As String
    Dim wHIn1           As String
    Dim wHin2           As String
    
    CoN1240Move = True
    
    
    On Error GoTo CoN1240Move_Err
    With KHT
        
        gStr = RTrim(.背番号)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
'        If CisFun.Len2(gStr) < 5 Then
'           gStr = gStr & Space(5 - CisFun.Len2(gStr))
'        End If
'        If Not CorData("SEBAN", gStr, "S", 5) Then GoTo CoN1240Move_Err
        If CisFun.Len2(gStr) < 8 Then
           gStr = gStr & Space(8 - CisFun.Len2(gStr))
        End If
        If Not CorData("SEBAN", gStr, "S", 8) Then GoTo CoN1240Move_Err
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
        If Not CorData("サイクル", .サイクル, "", 7) Then GoTo CoN1240Move_Err
        If Not CorData("置場", .置場, "", 10) Then GoTo CoN1240Move_Err
        If Not CorData("箱種", .収容器, "", 10) Then GoTo CoN1240Move_Err
        ' 部品かんばん・材料かんばん共通
        If kn_PrtKB = 0 Or kn_PrtKB = 2 Then
            If Not CorData("仕入先", .取引先CD, "", 8) Then GoTo CoN1240Move_Err
            If Not CorData("仕入先名", .取引先名, "", 20) Then GoTo CoN1240Move_Err
            If Not CorData("仕入先工場名", .取引先工場名, "", 20) Then GoTo CoN1240Move_Err
            
            If Not CorData("直送先", .納入先CD, "", 8) Then GoTo CoN1240Move_Err
            If Not CorData("直送先名", .納入先名, "", 20) Then GoTo CoN1240Move_Err
            If Not CorData("直送先工場名", .納入先工場名, "", 20) Then GoTo CoN1240Move_Err
            
            If Not CorData("自社名", .自社名1, "", 20) Then GoTo CoN1240Move_Err
            If Not CorData("自社工場", .自社名2, "", 20) Then GoTo CoN1240Move_Err
                    
            If Not CorData("受入", .受入, "", 2) Then GoTo CoN1240Move_Err
        End If
        ' 部品かんばん・工程かんばん・裏かんばん共通
        If kn_PrtKB = 0 Or kn_PrtKB = 1 Or kn_PrtKB = 3 Then
           If Trim(.品番30) = "" Then
              If Not CorData("品番14", .品番14, "", 14) Then GoTo CoN1240Move_Err
           Else
              If Not CorData("品番", .品番30, "", 30) Then GoTo CoN1240Move_Err
           End If
           gStr = Format(.収容数, "#,###")
           If Not CorData("収容数", gStr, "", 7) Then GoTo CoN1240Move_Err
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
           'If Not CorData("車型", .車型, "", 4) Then GoTo CoN1240Move_Err
           If Not CorData("車型", .車型, "", 5) Then GoTo CoN1240Move_Err
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
           If Not CorData("品名", .品名, "", 40) Then GoTo CoN1240Move_Err
        End If
        ' 部品かんばん
        If kn_PrtKB = 0 Then
           If Not CorData("直送先固定", .納入先固定, "", 6) Then GoTo CoN1240Move_Err
        End If
        ' 工程かんばん
        If kn_PrtKB = 1 Then
           If Not CorData("工程", .工程名, "", 16) Then GoTo CoN1240Move_Err
           If Not CorData("後工程", .次工程名, "", 16) Then GoTo CoN1240Move_Err
           gStr = Format(.生産ロット, "#,###")
           If Not CorData("生産ロット", gStr, "", 6) Then GoTo CoN1240Move_Err
           If .基準工数 <> 0 Then
              gStr = Format(.基準工数, "##0.0")
              If Not CorData("基準工数", gStr, "", 5) Then GoTo CoN1240Move_Err
           End If
           If Not CorData("基準生産日", .基準生産日, "", 14) Then GoTo CoN1240Move_Err
        End If
        ' 材料かんばん
        If kn_PrtKB = 2 Then
           If Not CorData("直送先固定", .納入先固定, "", 8) Then GoTo CoN1240Move_Err
           If Not CorData("材種コード", .材種コード, "", 15) Then GoTo CoN1240Move_Err
           gStr = Format(.収容数, "#,##0.000")
           If Not CorData("収容数", gStr, "", 11) Then GoTo CoN1240Move_Err
           If Trim(.材質) = "" Then
              If Not CorData("品番14", .材質18, "", 18) Then GoTo CoN1240Move_Err
           Else
              If Not CorData("品番", .材質, "", 30) Then GoTo CoN1240Move_Err
           End If
           If Not CorData("品名", .表示品番, "", 23) Then GoTo CoN1240Move_Err
        '   *---------------------*
        '   * 『材料』  SET *
        '   *---------------------*
            On Error Resume Next
            gStr = gIniExe & "BMP\材料.BMP"
            If Dir(gStr) <> "" Then
               If Not CorImage("ZAIRYO", gStr) Then Exit Function
            End If
            On Error GoTo 0
        End If
        ' 裏かんばん
        If kn_PrtKB = 3 Then
            If Not CorData("仕入先", .取引先CD, "", 8) Then GoTo CoN1240Move_Err
            If Not CorData("仕入先名", .取引先名, "", 16) Then GoTo CoN1240Move_Err
            
            If Not CorData("自社名", .自社名1, "", 20) Then GoTo CoN1240Move_Err
                    
            If Not CorData("受入", .受入, "", 2) Then GoTo CoN1240Move_Err
            If Not CorData("SSEBAN", .材質, "", 5) Then GoTo CoN1240Move_Err
        
        End If
    ' 発行日 編集
        With CisFun
            .DateE_BefVal = Format(Now(), "yyyymmdd")
            .DateEdit Dot
            gStr = .DateE_AfVal
        End With
        If Not CorData("発行日", gStr, "", 14) Then GoTo CoN1240Move_Err
        gStr = Format(.発行NO, "000")
        If Not CorData("連番", gStr, "", 3) Then GoTo CoN1240Move_Err

'   *---------------------*
'   * 『バーコード』SET    *
'   *---------------------*
'        If Trim(.アドレス) <> "" Then
'            gStr = Format(CisFun.Val2(.アドレス), "0000000") & Format(.発行NO, "000")
'            If CorBar("BAR", gStr) = "ERROR" Then GoTo CoN1240Move_Err
'            If Not CorData("BAR_Naiyo", gStr, "", 10) Then GoTo CoN1240Move_Err
'        End If
'   *---------------------*
'   * 『イメージ』  SET   *
'   *---------------------*
        On Error Resume Next
'        If Trim(.IMAGE) <> "" Then
'           If Dir(Kn_Image & .IMAGE) <> "" Then
'              If Not CorImage("IMAGE", Kn_Image & .IMAGE) Then GoTo CoN1240Move_Err
'           End If
'        End If
        On Error GoTo 0
'   *---------------------*
'   * 『QR』  SET         *
'   *---------------------*
        gStr = QR_OutCreate
        
        If s_Clip Then Clipboard.SetText Clipboard.GetText & gStr & vbCrLf
        
        If mQRTec Then
           If Not CorData("TECQR", gStr, "", 300) Then GoTo CoN1240Move_Err
        Else
           mQRMeker.InputData = gStr
           mQRMeker.Refresh
    
           mQRMeker.CreateQrMetaFile CorForm.hDC, gIniExe & "QR.BMP", 2
           If Not CorImage("QR", gIniExe & "QR.BMP") Then Exit Function
        End If
'   *---------------------*
'   * 『手配区分名』  SET *
'   *---------------------*
        If kn_PrtKB <> 3 Then       ' 裏かんばん以外
            On Error Resume Next
            gStr = Trim(.手配区分名)
            gStr = gIniExe & "BMP\" & gStr & ".BMP"
            If Dir(gStr) <> "" Then
               If Not CorImage("KUBUN", gStr) Then Exit Function
            End If
            On Error GoTo 0
        End If
    End With

    Exit Function
CoN1240Move_Err:
    CoN1240Move = False
End Function
'+-------------------------------------+
'+      仕切りかんばん発行
'+-------------------------------------+
Private Function PrintShikiri() As Boolean
    Dim WaitCnt     As Long
    PrintShikiri = False
    
    Brk_HMaisu = Brk_HMaisu + 1
    
    If Brk_Torcd <> KHT.取引先 Or Not KHT_RDSTS Then
       If Not CoReportProcN1240("E") Then GoTo PrintShikiri_Er
       
       WaitCnt = 0
       Do Until WaitCnt > 200000: WaitCnt = WaitCnt + 1: Loop
       
       If Not CoReportProcN1240("O", True) Then GoTo PrintShikiri_Er
        
       If Not CorData("取引先固定", "仕入先：", "S", 8) Then GoTo PrintShikiri_Er
       If Not CorData("取引先名", Brk_Tornm, "", 20) Then GoTo PrintShikiri_Er
       gStr = Format(Brk_HMaisu, "#,###")
       If Not CorData("発行枚数", gStr, "", 5) Then GoTo PrintShikiri_Er
       gStr = Format(Now(), "YYYY/MM/DD hh:nn")
       If Not CorData("発行日時", gStr, "", 16) Then GoTo PrintShikiri_Er
    
       If Not CoReportProcN1240("P") Then GoTo PrintShikiri_Er
       If Not CoReportProcN1240("E") Then GoTo PrintShikiri_Er
       
       If KHT_RDSTS Then
          If Not CoReportProcN1240("O") Then GoTo PrintShikiri_Er
       End If
       
       Brk_Torcd = KHT.取引先
       Brk_Tornm = KHT.取引先名
       Brk_HMaisu = 0
    
    End If
    
    
    PrintShikiri = True
    Exit Function
PrintShikiri_Er:

End Function
'*==========================================================================*
'*==========================================================================*
'*------------------------------------------------------------------*
'*                                                                  *
'*      CoReport Proccess           (O)Open   (E)Close              *
'*                                                                  *
'*------------------------------------------------------------------*
Public Function CoReportProcN1240(Repart As String, Optional PrtChg As Boolean = False)
    CoReportProcN1240 = False
'   *******************************************************************
'   * 印刷指示処理                                  ( Repart = "P" )  *
'   *******************************************************************
    If Repart = "P" Then
        If CorData("", "", "E") Then CoReportProcN1240 = True
    End If
'   *******************************************************************
'   * 終了処理                                      ( Repart = "E" )  *
'   *******************************************************************
    If Repart = "E" Then
        If CorOpCl(crPara, "E") Then CoReportProcN1240 = True
        Exit Function
    End If

'   *****************************************************************
'   * オープン処理    <<ﾌﾟﾘﾝﾀﾄﾞﾗｲﾊﾞ設定・初期処理>> ( Repart = "O" )  *
'   *****************************************************************
    If Repart = "O" Then
' CoReport 初期処理

        With crPara
'           * CoStruct Set *
            .cshwnd = CorForm.hWnd
            .cshdc = Printer.hDC
            .csout = OUTPUT_PRNDRIVER
            .csScale = 100&
            .csmx = 0: .csmy = 0
            .cssx = 0: .cssy = 0
            .cswidth = 0: .csheight = 0
            .csoutfname = ""
'           * DevModeStruct Set *　プリンタの設定
            If PrtChg Then
               .csfname = gIniExe & "Prt\" & CorPrtForm2 & ".FED"
               .csdevm.dmPaperLength = 850 + mShikiriAdd
               .csdevm.dmPaperWidth = 2000
            Else
               .csfname = gIniExe & "Prt\" & CorPrtForm & ".FED"
               .csdevm.dmPaperLength = 850
               .csdevm.dmPaperWidth = 2000
            End If
            .csdevm.dmOrientation = DMORIENT_PORTRAIT   'DMORIENT_LANDSCAPE
            .csdevm.dmPaperSize = DMPAPER_USER
            .csdevm.dmFields = DM_TEMPORARY Or DM_PAPERSIZE Or _
                               DM_PAPERLENGTH Or DM_PAPERWIDTH Or _
                               DM_ORIENTATION
        
        End With
        
        If CorOpCl(crPara, "O") Then
            CoReportProcN1240 = True
        End If

    End If
End Function

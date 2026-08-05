Attribute VB_Name = "CKN3210LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   発注かんばん発行(購入・社内:最終工程)
'*
'*   モジュールID  : CKN3210LB.BAS
'*
'*                   CKN3210LPrint(Integer)
'*
'*   初期設定値 : CKN3210_PrtKB(かんばん種類) :  0/購入,1/社内(最終工程)
'*                CKN3210_PrtSy(発行種類)     :  0 固定
'*
'*   引数 : PrtSy    =  0  :新規発行
'*                      1  :再発行
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:
'    If Not CKN3210LPrint(0) Then Exit Sub
'
'*   作成日  : 2005/12/06   By CIS  テクノエイト
'*   変更日  : 2005/12/13   By CIS - INIファイル-クライアント対応（TEXT出力先）
'*   変更日  : 2006/02/14   By CIS  印刷先ﾌﾟﾘﾝﾀ設定に対応
'*   変更日  : 2006/02/15   By CIS  『ポイントコード』出力中止
'*   変更日  : 2006/02/28   By CIS  バーコード初期値セット変更
'*   変更日  : 2006/03/24   By CIS  マルチバー３本目の品番だけを１０桁のみで発行
'*   変更日  : 2007/06/22   By CIS  Code39文字チェック追加
'*
'*---------------------------------------------------------------------------------
Global CKN3210_PrtKB            As Byte
Global CKN3210_PrtSy            As Byte
Global CKN3210_ImagePath        As String
Global CKN3210_CoReport         As New CoReports85        ' クラス定義

Global CKN3210_ColorCode(9)     As String
Global CKN3210_Color(9)         As Long
Global CKN3210_ColorSetFlg      As Boolean


Dim CorItem                     As CrObject
Dim User_BarCode                As String
Dim SetColor                    As Long
Dim EndNo                       As Long

Dim bCnt                        As Long
Dim bStr                        As String

Dim BarHeght                    As Long


Dim OutPrinter                  As String           ' 2006/02/14 Insert
'+--------------------------------------------+
'+      かんばん色初期内容セット
'+--------------------------------------------+
Public Sub CKN3210ColorSet()
' かんばん色取得
    Erase CKN3210_ColorCode
    Erase CKN3210_Color
    
    gSL_Select = "Select * from 名称マスタ"
    gSL_Select = gSL_Select & " Where 区分名称 = '色区分_発注'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & " Order By 値"
    
    If SYMRead(gSL_Select, , 1) Then
       gInt = 0
       Do Until Not SYM_RDSTS
            
          With SYM
                CKN3210_ColorCode(gInt) = .値
                CKN3210_Color(gInt) = .数字1
          End With
          gInt = gInt + 1
          
          Call SYMReadNext(1)
       Loop
    End If
    
    Call SYMClose(1)
    
    CKN3210_ColorSetFlg = True
End Sub
'+--------------------------------------------+
'+      かんばん色取得
'+--------------------------------------------+
Public Function CKN3210ColorGet(Torcd As String) As Long

    bStr = Format(Torcd, "@@@@")
    bStr = Replace(bStr, " ", "0")
    
    For bCnt = 1 To 9
        If Trim(CKN3210_ColorCode(bCnt)) = "" Then Exit For
        If CKN3210_ColorCode(bCnt) = Mid(bStr, 1, 1) Then
           CKN3210ColorGet = CKN3210_Color(bCnt)
           Exit Function
        End If
    Next bCnt
    
    CKN3210ColorGet = CKN3210_Color(0)

End Function
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/
'_/             印　　刷　　処　　理
'_/
'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
Public Function CKN3210LPrint(Optional PrtSy As Integer = 0)
    CKN3210LPrint = False
    If CKN3210_PrtKB = 0 Then
       App.Title = "発注かんばん"
    Else
       App.Title = "社内(最終)かんばん"
    End If
    
' 印刷進行画面表示ﾌﾗｸﾞｾｯﾄ
    pref_MsgDisp = True

    gSL_Select = "Select TR.取引先 取引先C,TR.工場CD 工場C,TR.略称 取引先名,TR.工場名 工場名,"
    gSL_Select = gSL_Select & " HT.アドレス,HT.IMAGE 品番イメージ,HN.チェック1,HN.チェック2,"
    gSL_Select = gSL_Select & " HT.納入先,HT.手配区分,HT.最終発行連番,"
    gSL_Select = gSL_Select & " ISNULL(KB.イメージ,'') 機番イメージ, "
    gSL_Select = gSL_Select & " ISNULL(PT.イメージ,'') ポイントイメージ,HN.ポイント, "
    gSL_Select = gSL_Select & " ISNULL(TR1.略称,'') 行先名1,HT.行先1,"
    gSL_Select = gSL_Select & " ISNULL(TR2.略称,'') 行先名2,HT.行先2,"
    gSL_Select = gSL_Select & " ISNULL(TR3.略称,'') 行先名3,HT.行先3,"
    gSL_Select = gSL_Select & " KW.* From かんばん発行ワーク KW "
    gSL_Select = gSL_Select & " Left Outer Join 品番マスタ HN"
    gSL_Select = gSL_Select & "      On HN.品番 = KW.品番"
    gSL_Select = gSL_Select & " Left Outer Join 取引先マスタ TR "
    gSL_Select = gSL_Select & "    ON  TR.取引先CD = KW.取引先"
    gSL_Select = gSL_Select & "    And TR.取引先区分 = 1"
    gSL_Select = gSL_Select & " Left Outer Join 品番取引先マスタ HT "
    gSL_Select = gSL_Select & "    On  HT.品番 = KW.品番"
    gSL_Select = gSL_Select & "    And HT.取引先 = KW.取引先"
    gSL_Select = gSL_Select & "    And HT.受入   = KW.受入"
    gSL_Select = gSL_Select & " Left Outer Join 機番マスタ KB"
    gSL_Select = gSL_Select & "   On KB.機番 = HT.機番"
    gSL_Select = gSL_Select & " Left Outer Join ポイントマスタ PT"
    gSL_Select = gSL_Select & "   On PT.ポイント = HN.ポイント"
    gSL_Select = gSL_Select & " Left Outer Join 取引先マスタ TR1"
    gSL_Select = gSL_Select & "   On  TR1.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR1.取引先CD = HT.行先1"
    gSL_Select = gSL_Select & " Left Outer Join 取引先マスタ TR2"
    gSL_Select = gSL_Select & "   On  TR2.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR2.取引先CD = HT.行先2"
    gSL_Select = gSL_Select & " Left Outer Join 取引先マスタ TR3"
    gSL_Select = gSL_Select & "   On  TR3.取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND TR3.取引先CD = HT.行先3"
    
    
    gSL_Select = gSL_Select & " Where KW.処理端末 = Host_Name()"
    gSL_Select = gSL_Select & "   And KW.かんばん種類 = " & CKN3210_PrtKB
    gSL_Select = gSL_Select & "   And KW.発行種類 = " & CKN3210_PrtSy
    gSL_Select = gSL_Select & "   And KW.発行済 = " & PrtSy
    gSL_Select = gSL_Select & " Order By KW.発行種類,KW.取引先,KW.受入,KW.背番号,KW.品番,KW.発行NO"

    If Not KHWRead(gSL_Select) Then
        Exit Function
    End If

    pref_TCnt = CisDB.RecordCount

'+++++++++++++++++++++++++++++++'
'+  印刷処理
'+++++++++++++++++++++++++++++++'
    Set pref_Form = CKN3210L
    If CKN3210_PrtKB = 0 Then
       pref_Msg = "【 発 注 か ん ば ん 印 刷 中 】"
    Else
       pref_Msg = "【 社内かんばん(最終)印 刷 中 】"
    End If
    pref_Form.Show vbModal
    Unload pref_Form
    Set pref_Form = Nothing

    CKN3210LPrint = True
End Function
'+----------------------------------------------------------+
'+                    印  刷  処  理                        ＋
'+----------------------------------------------------------+
Public Function CoPrintN3210() As Boolean
    Dim PageCnt     As Long
    
    CoPrintN3210 = False
    On Error GoTo CoPrintN3210_Er
    
' イメージ格納先
    With CisFun
'-----------------------------------------------< 2005/12/13 ADD START >
        .INI_FileID = gIniExe & gwsIniName
'-----------------------------------------------< 2005/12/13 ADD START >
        .INI_Section = "Option"
        .INI_Key = "Image":       .INI_Default = gIniExe
        .GetIni
        If Right(RTrim(.INI_String), 1) <> "\" Then
            CKN3210_ImagePath = RTrim(.INI_String) & "\"
        Else
            CKN3210_ImagePath = RTrim(.INI_String)
        End If
'----------------------------------------------- 2006/02/14 Insert
    ' かんばん印刷先
        OutPrinter = ""
        .INI_Section = "かんばん印刷先"
        .INI_Default = ""
        .INI_Key = CisFun.WSID
        .GetIni
        If Trim(.INI_String) <> "" Then
           OutPrinter = RTrim(.INI_String)
        End If
'----------------------------------------------- 2006/02/14 Insert End
'-----------------------------------------------< 2005/12/13 ADD START >
        .INI_FileID = gIniExe & gIniName
'-----------------------------------------------< 2005/12/13 ADD START >
    End With
' かんばん色設定
    If Not CKN3210_ColorSetFlg Then
       Call CKN3210ColorSet
    End If
    
' ﾊﾞｰｺｰﾄﾞ初期ｾｯﾄ
    Call GetClientOption("G", "バーコード", "かんばん", gSYSOp_IniAtai)
    If RTrim(gSYSOp_IniAtai) = "" Then

'-----     -----     -----     -----     -----     -----     < 2006/02/28 MNT START >
'        gSYSOp_IniAtai = "B0"
'        Call GetClientOption("P", "バーコード", "BU", "3,5,6,10,3")
        gSYSOp_IniAtai = "BU"
        Call GetClientOption("P", "バーコード", "BU", "5,8,13,20,7")
'-----     -----     -----     -----     -----     -----     < 2006/02/28 MNT  END  >

        Call GetClientOption("P", "バーコード", "B0", "Form設定")
        Call GetClientOption("P", "バーコード", "B1", "3,5,6,10,3")
        Call GetClientOption("P", "バーコード", "B2", "2,3,4,6,3")
        Call GetClientOption("P", "バーコード", "B3", "4,8,8,16,6")
        Call GetClientOption("P", "バーコード", "B4", "3,4,6,8,3")
        Call GetClientOption("P", "バーコード", "かんばん高さ", "0")
        Call GetClientOption("P", "バーコード", "かんばん", gSYSOp_IniAtai)
    End If
    If gSYSOp_IniAtai = "BU" Then
        Call GetClientOption("G", "バーコード", "BU", User_BarCode)
        If RTrim(User_BarCode) = "" Then
            gSYSOp_IniAtai = "B0"
        End If
    End If
    Call GetClientOption("G", "バーコード", "かんばん高さ", gStr)
    BarHeght = CisFun.Val2(gStr)
    
    pref_NCnt = 0                ' 現在発行ページ初期化
    
    With CKN3210_CoReport
        .SetDraw = pref_Form.CrDraw1

        .SetPrtFormCnt = 2
        .SetPrtForm = "LKN3210"        ' ﾌｫｰﾑIDをｾｯﾄ
       ' 複数ﾌｫｰﾑIDｾｯﾄ
        .SetFormNo = 2: .SetPrtForm = "LKN3210B"
        
        .SetPrtPath = gIniExe & "Prt\" ' ﾌｫｰﾑ格納先
        .SetPrtOut = PrinterX          ' ﾌﾟﾘﾝﾀｰ
'        Select Case mPrintOut
'            Case 0: .SetPrtOut = Printer            ' ﾌﾟﾘﾝﾀｰ
'            Case 1: .SetPrtOut = Preview            ' ﾌﾟﾚﾋﾞｭｰ
'                    .SetPreviewForm = frmPreview    '   ﾌﾟﾚﾋﾞｭｰ用ﾌｫｰﾑ(frm)
'            Case 2: .SetPrtOut = PDF                ' PDF出力
'        End Select
        If CKN3210_PrtKB = 0 Then
           .SetPrtName = "発注かんばん"                      ' 帳票名
        Else
           .SetPrtName = "社内(最終)かんばん"                      ' 帳票名
        End If
        .SetDuplex = 2
        PageCnt = 0
'------------------------------------------------------ 2006/02/14 Insert
        .SetPrinter = OutPrinter
        .SetPaperSize = A4
        .SetPaperOrientation = 縦
'------------------------------------------------------ 2006/02/14 Insert End
        
        '------ 印刷処理開始
        If Not .CorOutStart Then GoTo CoPrintN3210_Er
        
        ' ﾌｫｰﾑ項目非表示
        Call PrtItemHeide(1)
        Call PrtItemHeide(2)
                
        Do Until Not KHW_RDSTS
            PageCnt = PageCnt + 1
        '-----------------------------------------------    最終発行NOを求める(新規発行時)
            EndNo = 0
            If KHW.発行区分 = 0 And KHW.発行済 = 0 Then
               EndNo = KHW.最終発行連番
               If CKN3210_PrtKB = 0 Then
                  gSL_Select = "Select Max(発行NO) 発行NO From かんばん流動管理テーブル "
                  gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(KHW.取引先) & "'"
                  gSL_Select = gSL_Select & " And   受入 = '" & RTrim(KHW.受入) & "'"
                  gSL_Select = gSL_Select & " And   品番 = '" & RTrim(KHW.品番) & "'"
                  If KRTRead(gSL_Select, 1, 1) Then
                     If KRT.発行NO > KHW.最終発行連番 Then
                        EndNo = KRT.発行NO
                     End If
                  End If
               Else
                  gSL_Select = "Select Max(発行NO) 発行NO From かんばん流動社内テーブル "
                  gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(KHW.取引先) & "'"
                  gSL_Select = gSL_Select & " And   受入 = '" & RTrim(KHW.受入) & "'"
                  gSL_Select = gSL_Select & " And   品番 = '" & RTrim(KHW.品番) & "'"
                  If KRSRead(gSL_Select, 1, 1) Then
                     If KRS.発行NO > KHW.最終発行連番 Then
                        EndNo = KRS.発行NO
                     End If
                  End If
               End If
               KHW.発行NO = EndNo + 1
            End If
            '------ かんばん色設定
            SetColor = CKN3210ColorGet(KHW.取引先C)
            
            '処理ﾌｫｰﾑ選択
            .SetFormNo = 1
            '------ 項目転送(おもて)
            Call ItemMoveFore(PageCnt)
            '------ 印刷(１ﾍﾟｰｼﾞ分)
            If PageCnt >= 3 Then
               If Not .CoPrintOut Then GoTo CoPrintN3210_Er
               Call PrtItemHeide(1)
            End If

            '処理ﾌｫｰﾑ選択
            .SetFormNo = 2
            '------ 項目転送(うら）
            Call ItemMoveBack(PageCnt)
            '------ 印刷(１ﾍﾟｰｼﾞ分)
            If PageCnt >= 3 Then
               If Not .CoPrintOut Then GoTo CoPrintN3210_Er
               PageCnt = 0
               Call PrtItemHeide(2)
            End If
            '------ かんばん流動管理テーブル更新
            If CKN3210_PrtKB = 0 Then
               gSL_Select = "Select * From かんばん流動管理テーブル"
               gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(KHW.取引先) & "'"
               gSL_Select = gSL_Select & " And   受入 = '" & RTrim(KHW.受入) & "'"
               gSL_Select = gSL_Select & " And   品番 = '" & RTrim(KHW.品番) & "'"
               gSL_Select = gSL_Select & " And   発行NO = " & KHW.発行NO & ""
               If Not KRTRead(gSL_Select, 1, 1) Then
                  Call ItemsClearKRT
                  With KRT
                       .取引先 = KHW.取引先
                       .受入 = KHW.受入
                       .品番 = KHW.品番
                       .発行NO = KHW.発行NO
                       .発行区分 = 0
                       .収容区分 = KHW.収容区分
                       .所番地区分 = KHW.所番地区分
                       .所番地 = KHW.所番地
                       .収容数 = KHW.収容数
                       .箱種 = KHW.収容器
                   End With
                   Call KRTInsert(1)
               Else
                   With KRT
                       .版数 = .版数 + 1
                       .発行区分 = 0
                       .収容区分 = KHW.収容区分
                       .所番地区分 = KHW.所番地区分
                       .所番地 = KHW.所番地
                       .収容数 = KHW.収容数
                       .箱種 = KHW.収容器
                   End With
                   Call KRTUpdate(1)
              End If
            Else
               gSL_Select = "Select * From かんばん流動社内テーブル"
               gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(KHW.取引先) & "'"
               gSL_Select = gSL_Select & " And   受入 = '" & RTrim(KHW.受入) & "'"
               gSL_Select = gSL_Select & " And   品番 = '" & RTrim(KHW.品番) & "'"
               gSL_Select = gSL_Select & " And   発行NO = " & KHW.発行NO & ""
               If Not KRSRead(gSL_Select, 1, 1) Then
                  Call ItemsClearKRS
                  With KRS
                       .取引先 = KHW.取引先
                       .受入 = KHW.受入
                       .品番 = KHW.品番
                       .発行NO = KHW.発行NO
                       .発行区分 = 0
                       .収容区分 = KHW.収容区分
                       .所番地区分 = KHW.所番地区分
                       .所番地 = KHW.所番地
                       .収容数 = KHW.収容数
                       .箱種 = KHW.収容器
                   End With
                   Call KRSInsert(1)
               Else
                   With KRS
                       .版数 = .版数 + 1
                       .発行区分 = 0
                       .収容区分 = KHW.収容区分
                       .所番地区分 = KHW.所番地区分
                       .所番地 = KHW.所番地
                       .収容数 = KHW.収容数
                       .箱種 = KHW.収容器
                   End With
                   Call KRSUpdate(1)
              End If
            End If
           '--------------------- かんばん発行ワーク更新
            With CisDB
                .ConnectNo = 1
                .SQL = "Update かんばん発行ワーク Set 発行済 = 2"
                If KHW.発行区分 = 0 And KHW.発行済 = 0 Then
                   .SQL = .SQL & " , 発行NO = " & KHW.発行NO
                End If
                .SQL = .SQL & " Where 処理端末 = Host_Name()"
                .SQL = .SQL & " And   かんばん種類 = " & CKN3210_PrtKB
                .SQL = .SQL & " And   品番 = '" & RTrim(KHW.品番) & "'"
                .SQL = .SQL & " And   取引先 = '" & RTrim(KHW.取引先) & "'"
                .SQL = .SQL & " And   受入 = '" & RTrim(KHW.受入) & "'"
                .SQL = .SQL & " And   連番 = " & KHW.連番
                If Not .DBExec Then GoTo CoPrintN3210_Er
            
            '--------------------- 品番取引先マスタ　最終発行連番 更新
                If KHW.発行区分 = 0 And KHW.発行済 = 0 Then
                   .SQL = "Update 品番取引先マスタ Set  "
                   .SQL = .SQL & "最終発行連番 = Case When 最終発行連番 > " & KHW.発行NO & " THEN 最終発行連番 Else " & KHW.発行NO & " End,"
                   .SQL = .SQL & "更新日 = GETDATE(),"
                   .SQL = .SQL & "更新者 = '" & gTanto & "',"
                   .SQL = .SQL & "更新端末 = SubString(Host_Name(),1,20)"
                   .SQL = .SQL & " Where 品番 = '" & RTrim(KHW.品番) & "'"
                   .SQL = .SQL & "   And 取引先 = '" & RTrim(KHW.取引先) & "'"
                   .SQL = .SQL & "   And 受入 = '" & RTrim(KHW.受入) & "'"
                   If Not .DBExec Then GoTo CoPrintN3210_Er
                End If
            End With
            
            ' 進捗ゲージ更新
            pref_NCnt = pref_NCnt + 1
            pref_Cnt = Format(pref_NCnt, "###") & " / " & Format(pref_TCnt, "###")
            DoEvents
            '中断処理
            If pref_Cansel_Flg Then
                If MsgBox("処理を中断します。よろしいですか？", 33, "印刷") = 1 Then
                   GoTo CoPrintN3210_ED
                Else
                   pref_Cansel_Flg = False
                End If
            End If
            
            Call KHWReadNext
        Loop
        
        If PageCnt <> 0 Then
            .SetFormNo = 1
           If Not .CoPrintOut Then GoTo CoPrintN3210_Er
            .SetFormNo = 2
           If Not .CoPrintOut Then GoTo CoPrintN3210_Er
        End If
        
CoPrintN3210_ED:
        Call KHWClose
        
        '------ 印刷終了
        If Not .CoOutEnd Then GoTo CoPrintN3210_Er
    End With
    
    CoPrintN3210 = True
    Exit Function
CoPrintN3210_Er:
    Call CisFun.ErrorBox("印刷エラー")
End Function
'+--------------------------------------------------+
'+      フォーム項目非表示 & 初期化
'+--------------------------------------------------+
Private Sub PrtItemHeide(FormNo As Integer)
    With CKN3210_CoReport
        '処理ﾌｫｰﾑ選択
        .SetFormNo = FormNo
        '------ 全て非表示
        For Each CorItem In .CorForm.CrObjects
            CorItem.Visible = False
        Next CorItem
        '------ フィールドを全てクリアします
        .CorForm.ClearAllFields
    End With
End Sub
'*------------------------------------------------------------------*
'*                                                                  *
'*      かんばん　発行内容　項目転送   (おもて)                     *
'*                                                                  *
'*------------------------------------------------------------------*
Private Sub ItemMoveFore(PageNo As Long)
'-- ﾊﾞｰｺｰﾄﾞ
    Dim wBarArea    As String * 15
    
    Dim wBarWork    As String * 20
    Dim wBarCD      As String * 1

    With CKN3210_CoReport
        
        '------------------ 仕入先情報
        gStr = Format(KHW.取引先C, "@@@@")  '右詰
        If Trim(KHW.工場C) <> "" Then
           gStr = gStr & "-" & KHW.工場C
        End If
        If Not .CorData("仕入先CD" & Format(PageNo, "0"), gStr, "", 7) Then Exit Sub
        gStr = CisFun.Mid2(KHW.取引先名, 1, 10)
        If Not .CorData("仕入先名1" & Format(PageNo, "0"), gStr, "", 10) Then Exit Sub
        gStr = CisFun.Mid2(KHW.取引先名, 11, 10)
        If Not .CorData("仕入先名2" & Format(PageNo, "0"), gStr, "", 10) Then Exit Sub
        gStr = CisFun.Mid2(KHW.工場名, 1, 10)
        If Not .CorData("工場名" & Format(PageNo, "0"), gStr, "", 10) Then Exit Sub
            '--- 罫線・固定項目
        Call FixedColorSet("仕入先BOX" & Format(PageNo, "0"))
        Call FixedColorSet("lb_仕入先" & Format(PageNo, "0"))
        Call FixedColorSet("lb_仕入先名" & Format(PageNo, "0"))
                
        '------------------ サイクル～作成日
        If Trim(KHW.曜日) <> "" Then
           If Not .CorData("サイクル" & Format(PageNo, "0"), KHW.曜日, "", 8) Then Exit Sub
        Else
            If Trim(KHW.サイクル) <> "" Then
               gStr = Mid(KHW.サイクル, 1, 1)
               gStr = gStr & "-" & Format(Mid(KHW.サイクル, 2, 2), "0")
               gStr = gStr & "-" & Format(Mid(KHW.サイクル, 4, 2), "0")
               If Not .CorData("サイクル" & Format(PageNo, "0"), gStr, "", 8) Then Exit Sub
            End If
        End If
        gStr = Format(KHW.発行NO, "000")
        If Not .CorData("連番" & Format(PageNo, "0"), gStr, "", 3) Then Exit Sub
        If Not .CorData("アドレス" & Format(PageNo, "0"), KHW.アドレス, "", 5) Then Exit Sub
        gStr = Format(Now(), "YYYY/MM/DD")
        If Not .CorData("作成日" & Format(PageNo, "0"), gStr, "", 10) Then Exit Sub
            '--- 罫線・固定項目
        Call FixedColorSet("ｻｲｸﾙBOX" & Format(PageNo, "0"))
        Call FixedColorSet("lb_ｻｲｸﾙ" & Format(PageNo, "0"))
        Call FixedColorSet("lb_発行連番" & Format(PageNo, "0"))
        Call FixedColorSet("lb_ADDR" & Format(PageNo, "0"))
        Call FixedColorSet("lb_作成日" & Format(PageNo, "0"))
        '------------------ 置場
        gStr = Mid(KHW.所番地, 1, 4)
        If Not .CorData("置場1" & Format(PageNo, "0"), gStr, "", 4) Then Exit Sub
        gStr = Mid(KHW.所番地, 5, 4)
        If Not .CorData("置場2" & Format(PageNo, "0"), gStr, "", 4) Then Exit Sub
            '--- 罫線・固定項目
        Call FixedColorSet("所番地BOX" & Format(PageNo, "0"))
        Call FixedColorSet("所番地ライン" & Format(PageNo, "0"))
        Call FixedColorSet("lb_置場" & Format(PageNo, "0"), False)
        '------------------ 品番
        gStr = Mid(KHW.表示品番, 1, 14)
        If Not .CorData("品番" & Format(PageNo, "0"), gStr, "", 14) Then Exit Sub
            '--- 罫線・固定項目
        Call FixedColorSet("品番BOX" & Format(PageNo, "0"))
        Call FixedColorSet("lb_品番" & Format(PageNo, "0"))
        '------------------ 箱種
        gStr = Mid(KHW.収容器, 1, 10)
        If Not .CorData("箱種" & Format(PageNo, "0"), gStr, "", 10) Then Exit Sub
            '--- 罫線・固定項目
        Call FixedColorSet("箱種BOX" & Format(PageNo, "0"))
        Call FixedColorSet("lb_箱種" & Format(PageNo, "0"))
        '------------------ 収容数
        gStr = Format(KHW.収容数, "#")
        If Not .CorData("収容数" & Format(PageNo, "0"), gStr, "", 6) Then Exit Sub
            '--- 罫線・固定項目
        Call FixedColorSet("収容数BOX" & Format(PageNo, "0"))
        Call FixedColorSet("lb_収容数" & Format(PageNo, "0"))
        '------------------ 背番号
        gStr = Mid(KHW.背番号, 1, 4)
        If Not .CorData("背番号" & Format(PageNo, "0"), gStr, "", 4) Then Exit Sub
            '--- 罫線・固定項目
        Call FixedColorSet("背番号BOX" & Format(PageNo, "0"))
        Call FixedColorSet("背番号BOX2" & Format(PageNo, "0"))
        Call FixedColorSet("lb_背番号" & Format(PageNo, "0"), False)
        '------------------ 機番(ｲﾒｰｼﾞ)
        Call ImageSet("機番イメージ" & Format(PageNo, "0"), KHW.機番イメージ)
        Call FixedColorSet("機番BOX" & Format(PageNo, "0"))
        '------------------ 自社(ｲﾒｰｼﾞ)
        Call ImageSet("ロゴ" & Format(PageNo, "0"), "Logo.jpg")
        '------------------ 自社(ｲﾒｰｼﾞ)
        Call ImageSet("COMPANY" & Format(PageNo, "0"), "company.jpg")
        '------------------ ポイント
        
'-----     -----     -----     -----     -----     -----     -----     -----     < 2006/02/15 DEL >
'
'        If Not .CorData("ポイント" & Format(PageNo, "0"), KHW.ポイント, "", 3) Then Exit Sub
'
'-----     -----     -----     -----     -----     -----     -----     -----     < 2006/02/15 DEL >
        
        Call ImageSet("ポイントイメージ" & Format(PageNo, "0"), KHW.ポイントイメージ)
        Call FixedColorSet("ポイントBOX" & Format(PageNo, "0"))
        Call FixedColorSet("lb_ポイント" & Format(PageNo, "0"))
        '------------------ 行先
        gStr = CisFun.Mid2(KHW.行先名1, 1, 10)
        If Not .CorData("行先1" & Format(PageNo, "0"), gStr, "", 10) Then Exit Sub
        gStr = CisFun.Mid2(KHW.行先名2, 1, 10)
        If Not .CorData("行先2" & Format(PageNo, "0"), gStr, "", 10) Then Exit Sub
        gStr = CisFun.Mid2(KHW.行先名3, 1, 10)
        If Not .CorData("行先3" & Format(PageNo, "0"), gStr, "", 10) Then Exit Sub
        Call FixedColorSet("行先BOX" & Format(PageNo, "0"))
        Call FixedColorSet("lb_行先" & Format(PageNo, "0"))
        Call FixedColorSet("lb_行先1" & Format(PageNo, "0"))
        Call FixedColorSet("lb_行先2" & Format(PageNo, "0"))
        Call FixedColorSet("lb_行先3" & Format(PageNo, "0"))
        '------------------ 受入
        If Not .CorData("受入" & Format(PageNo, "0"), KHW.受入, "", 2) Then Exit Sub
            '--- 罫線・固定項目
        Call FixedColorSet("受入BOX" & Format(PageNo, "0"))
        Call FixedColorSet("lb_受入" & Format(PageNo, "0"))
    
    
    '   *---------------------*
    '   * 『バーコード』SET    *
    '   *---------------------*
        
        Select Case gSYSOp_IniAtai
            Case "BU": .SetBar39 = User
                       .SetBar39User = User_BarCode
            Case "B0": .SetBar39 = form設定
            Case "B2": .SetBar39 = パターン2_2_3_4_6_3
            Case "B3": .SetBar39 = パターン3_4_8_8_16_6
            Case "B4": .SetBar39 = パターン4_3_4_6_8_3
            Case Else: .SetBar39 = パターン1_3_5_6_10_3
        End Select
        .SetBar39Height = BarHeght
        
        '------------------------------
        
'　　　"０"（固定）＋仕入先＋納入先＋かんばん種別＋かん改訂＋ＣＤ
'　　　　　（１） 　（６）　（６）　　　（１）　　　（１）　（１）
        wBarArea = ""
        wBarWork = "0"
        wBarArea = Mid(wBarWork, 1, 1)
        gStr = Format(KHW.取引先C, "@@@@")  '右詰
        
        If Trim(KHW.工場C) <> "" Then
           gStr = gStr & KHW.工場C
        End If
        wBarWork = gStr
        wBarArea = Mid(wBarArea, 1, 1) & Mid(wBarWork, 1, 6)
        If KHW.手配区分 = "1" Then
           wBarWork = KHW.納入先
        Else
           wBarWork = ""
        End If
        wBarArea = Mid(wBarArea, 1, 7) & Mid(wBarWork, 1, 6)
        wBarArea = Mid(wBarArea, 1, 13) & "10"
        wBarCD = cCheck43(wBarArea, 15)
        
        If Not .CorData("Bar0" & Format(PageNo, "0"), wBarArea & wBarCD, "", 16) Then Exit Sub
        gStr = "*" & wBarArea & wBarCD & "*"
        If Not .CorData("BarData0" & Format(PageNo, "0"), gStr, "", 18) Then Exit Sub
    
'       '1' + 連番[3] + アドレス[5] + ?
'　　　"１"（固定）＋発行枝番＋かんＳＥＬＦ＋ブランク＋ＣＤ
'　　　　　（１）　　　（４）　　（５）　　　（５）　 （１）
        wBarArea = ""
        wBarWork = "1"
        wBarArea = Mid(wBarWork, 1, 1)
        wBarWork = Format(KHW.発行NO, "0000")
        wBarArea = Mid(wBarArea, 1, 1) & Mid(wBarWork, 1, 4)
        wBarWork = KHW.アドレス
        wBarArea = Mid(wBarArea, 1, 5) & Mid(wBarWork, 1, 5)
        wBarCD = cCheck43(wBarArea, 15)
        
        If Not .CorData("Bar1" & Format(PageNo, "0"), wBarArea & wBarCD, "", 16) Then Exit Sub
        gStr = "*" & wBarArea & wBarCD & "*"
        If Not .CorData("BarData1" & Format(PageNo, "0"), gStr, "", 18) Then Exit Sub

'　　　"２"（固定）＋品番＋ＣＤ
'　 　（１）　     （１４）（１）
        wBarArea = ""
        wBarWork = "2"
        wBarArea = Mid(wBarWork, 1, 1)
        
'-----     -----     -----     -----     -----     -----     -----     < 2006/03/24 MNT >
'        wBarWork = KHW.品番
        wBarWork = Left$(KHW.品番, 10)
'-----     -----     -----     -----     -----     -----     -----     < 2006/03/24 MNT END >
        
'---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   < 2007/06/22 MNT >
'        wBarArea = Mid(wBarArea, 1, 1) & Mid(wBarWork, 1, 14)
'        wBarCD = cCheck43(wBarArea, 15)
'
'        If Not .CorData("Bar2" & Format(PageNo, "0"), wBarArea & wBarCD, "", 16) Then Exit Sub
'        gStr = "*" & wBarArea & wBarCD & "*"
'        If Not .CorData("BarData2" & Format(PageNo, "0"), gStr, "", 18) Then Exit Sub
        
        If ChrCheck_Code39(wBarWork) Then
            wBarArea = Mid(wBarArea, 1, 1) & Mid(wBarWork, 1, 14)
            wBarCD = cCheck43(wBarArea, 15)
            
            If Not .CorData("Bar2" & Format(PageNo, "0"), wBarArea & wBarCD, "", 16) Then Exit Sub
            gStr = "*" & wBarArea & wBarCD & "*"
            If Not .CorData("BarData2" & Format(PageNo, "0"), gStr, "", 18) Then Exit Sub
        End If
'---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   < 2007/06/22 MNT END >

'　　　"＋"（固定）＋ストアアドレス＋ブランク＋収容数＋ＣＤ
'　　　　　（１）　　　　（８）　　　（１）　　 （５）（１）
        wBarArea = ""
        wBarWork = "+"
        wBarArea = Mid(wBarWork, 1, 1)
        wBarWork = Mid(KHW.所番地, 1, 8)
        wBarArea = Mid(wBarArea, 1, 1) & Mid(wBarWork, 1, 8)
        wBarWork = ""
        wBarArea = Mid(wBarArea, 1, 9) & Mid(wBarWork, 1, 1)
        wBarWork = Format(KHW.収容数, "00000")
        wBarArea = Mid(wBarArea, 1, 10) & Mid(wBarWork, 1, 5)
        wBarCD = cCheck43(wBarArea, 15)
        
'---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   < 2007/06/22 MNT >
'        If Not .CorData("Bar3" & Format(PageNo, "0"), wBarArea & wBarCD, "", 16) Then Exit Sub
'        gStr = "*" & wBarArea & wBarCD & "*"
'        If Not .CorData("BarData3" & Format(PageNo, "0"), gStr, "", 18) Then Exit Sub

        If ChrCheck_Code39(wBarArea) Then
            If Not .CorData("Bar3" & Format(PageNo, "0"), wBarArea & wBarCD, "", 16) Then Exit Sub
            gStr = "*" & wBarArea & wBarCD & "*"
            If Not .CorData("BarData3" & Format(PageNo, "0"), gStr, "", 18) Then Exit Sub
        End If
'---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   < 2007/06/22 MNT END >
        
'------------------------------------------------------------  QR
        Call QRAreaClear
        With QRArea
            .連番 = KHW.発行NO
            .FMTID = "K01"
            gStr = Format(KHW.取引先C, "@@@@")  '右詰
            If Trim(KHW.工場C) <> "" Then
               gStr = gStr & KHW.工場C
            End If
            .取引先 = gStr
            If KHW.手配区分 = "1" Then
               .納入先 = KHW.納入先
            End If
            .かんばん種類 = "1"
            .アドレス = KHW.アドレス
            .収容数 = KHW.収容数
            .受入 = KHW.受入
            .背番号 = KHW.背番号
            .品番 = KHW.表示品番
            .箱種 = KHW.収容器
            .所番地 = KHW.所番地
            .ポイント = ""
            .行先1 = KHW.行先1
            .行先2 = KHW.行先2
            .行先3 = KHW.行先3
        End With
        gStr = QR_KanbanCreate
        If Not .CorQRData("QR" & Format(PageNo, "0"), gStr, "", 150) Then Exit Sub
    
    
        '------------------ 区切り線
        With CKN3210_CoReport
            .CorForm.CrObjects("ライン1" & Format(PageNo, "0")).Visible = True
            .CorForm.CrObjects("ライン2" & Format(PageNo, "0")).Visible = True
        End With
        
    End With
End Sub
'*------------------------------------------------------------------*
'*                                                                  *
'*      かんばん　発行内容　項目転送   (うら)                       *
'*                                                                  *
'*------------------------------------------------------------------*
Private Sub ItemMoveBack(PageNo As Long)

    With CKN3210_CoReport
        '------------------ 品番
        If Not .CorData("品番" & Format(PageNo, "0"), KHW.表示品番, "", 14) Then Exit Sub
        Call FixedColorSet("lb_品番" & Format(PageNo, "0"))
        '------------------ 品番(ｲﾒｰｼﾞ)
        Call ImageSet("品番イメージ" & Format(PageNo, "0"), KHW.品番イメージ)
        '------------------ 箱種
        If Not .CorData("箱種" & Format(PageNo, "0"), KHW.収容器, "", 10) Then Exit Sub
        Call FixedColorSet("lb_箱種" & Format(PageNo, "0"))
        '------------------ チェック
        If Not .CorData("チェック1" & Format(PageNo, "0"), KHW.チェック1, "", 10) Then Exit Sub
        If Not .CorData("チェック2" & Format(PageNo, "0"), KHW.チェック2, "", 10) Then Exit Sub
        Call FixedColorSet("lb_チェック" & Format(PageNo, "0"))
        '------------------ 発行日
        gStr = Format(Now(), "yyyy年m月d日")
        If Not .CorData("発行日" & Format(PageNo, "0"), gStr, "", 14) Then Exit Sub
        Call FixedColorSet("lb_発行" & Format(PageNo, "0"))
    
        '------------------ 罫線
        Call FixedColorSet("BOX" & Format(PageNo, "0"))
        Call FixedColorSet("縦線" & Format(PageNo, "0"))
        Call FixedColorSet("横線1" & Format(PageNo, "0"))
        Call FixedColorSet("横線2" & Format(PageNo, "0"))
        Call FixedColorSet("横線3" & Format(PageNo, "0"))
    
        '------------------ 区切り線
        With CKN3210_CoReport
            .CorForm.CrObjects("ライン" & Format(PageNo, "0")).Visible = True
            .CorForm.CrObjects("ライン2" & Format(PageNo, "0")).Visible = True
        End With
    End With
End Sub
'+--------------------------------------------+
'+      固定項目 表示 & 色設定
'+--------------------------------------------+
Private Sub FixedColorSet(ItemName As String, Optional ColorFlg As Boolean = True)
    With CKN3210_CoReport
        .CorForm.CrObjects(ItemName).Visible = True
        
        If Not ColorFlg Then Exit Sub
        
        If .CorForm.CrObjects(ItemName).ObjectType = corBox Then
           .CorForm.CrObjects(ItemName).LineColor = SetColor
        End If
        If .CorForm.CrObjects(ItemName).ObjectType = corLine Then
           .CorForm.CrObjects(ItemName).LineColor = SetColor
        End If
        If .CorForm.CrObjects(ItemName).ObjectType = corLabel Then
           .CorForm.CrObjects(ItemName).TextColor = SetColor
        End If
    End With
End Sub
'+--------------------------------------------+
'+      イメージ項目 表示 & 画像セット
'+--------------------------------------------+
Private Sub ImageSet(ItemName As String, ImageName As String)
    Dim ImageCheckFlg   As Boolean
    
    With CKN3210_CoReport
        .CorForm.CrObjects(ItemName).Visible = False
    End With
    
    If RTrim(ImageName) = "" Then Exit Sub
    
    On Error GoTo ImageSet_Ex
    ImageCheckFlg = False
    gStr = RTrim(CKN3210_ImagePath) & ImageName
    If Dir(gStr) <> "" Then
        If Not CKN3210_CoReport.CorImage(ItemName, gStr) Then Exit Sub
        ImageCheckFlg = True
    End If
ImageSet_Ex:
    If ImageCheckFlg Then Exit Sub
'    gStr = gIniExe & ImageName
'    If Dir(gStr) <> "" Then
'        If Not CKN3210_CoReport.CorImage(ItemName, gStr) Then Exit Sub
'    End If
End Sub
'*--------------------------------------------------------------------------------
'*
'*   チェックCD 計算(モジュラス４３）
'*
'*   モジュールID    : cCheck43
'*
'*   初期設定項目    : なし
'*   引数           : StrIN  --> ﾁｪｯｸｻﾑ対象文字列
'*                  : InLng  --> ﾁｪｯｸｻﾑ対象文字数
'*   戻値           : 文字（１桁）
'*   例   :
'*    Str = cCheck43(StrIN,10)
'*
'*   作成日  :2005/04/20   cis K.K
'*   変更日  :
'*---------------------------------------------------------------------------------
Private Function cCheck43(StrIN As String, InLng As Integer) As String
    Static CDInt    As Integer
    Static CDComp   As Integer
    Static CDNum    As Integer
    Static CDSum    As String * 1
    
    cCheck43 = " "
    CDComp = 0
    CDSum = " "
    
    For CDInt = 1 To InLng
         CDNum = 0
         Select Case UCase(Mid(StrIN, CDInt, 1))
            Case "0": CDNum = 0
            Case "1": CDNum = 1
            Case "2": CDNum = 2
            Case "3": CDNum = 3
            Case "4": CDNum = 4
            Case "5": CDNum = 5
            Case "6": CDNum = 6
            Case "7": CDNum = 7
            Case "8": CDNum = 8
            Case "9": CDNum = 9
            Case "-": CDNum = 36
            Case ".": CDNum = 37
            Case " ": CDNum = 38
            Case "$": CDNum = 39
            Case "/": CDNum = 40
            Case "+": CDNum = 41
            Case "%": CDNum = 42
            Case "A" To "Z"
                 CDNum = Asc(Mid(StrIN, CDInt, 1)) - 55
        End Select
        CDComp = CDComp + CDNum
    Next CDInt
    CDNum = CDComp Mod 43
    CDSum = ""
    Select Case CDNum
       Case 0: CDSum = "0"
       Case 1: CDSum = "1"
       Case 2: CDSum = "2"
       Case 3: CDSum = "3"
       Case 4: CDSum = "4"
       Case 5: CDSum = "5"
       Case 6: CDSum = "6"
       Case 7: CDSum = "7"
       Case 8: CDSum = "8"
       Case 9: CDSum = "9"
       Case 36: CDSum = "-"
       Case 37: CDSum = "."
       Case 38: CDSum = " "
       Case 39: CDSum = "$"
       Case 40: CDSum = "/"
       Case 41: CDSum = "+"
       Case 42: CDSum = "%"
       Case 10 To 35
            CDSum = Chr(CDNum + 55)
   End Select
   cCheck43 = CDSum
End Function
'---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   < 2007/06/22 ADD >
'*--------------------------------------------------------------------------------
'*
'*   Code 39 文字チェック
'*
'*   初期設定項目    : なし
'*   引数           : pVal  -   チェック対象文字列
'*   戻値           : True  -   Check OK,   False -   Check NG
'*   例   :
'*    Str = cCheck43(StrIN,10)
'*
'*   作成日  :  2007/06/22  By CIS
'*   変更日  :
'*
'*---------------------------------------------------------------------------------
Private Function ChrCheck_Code39(pVal As String) As Boolean
    Static swI      As Byte
    Static swF      As Boolean
    Static swInt    As Integer
    Static swVal    As String
    
    For swI = 1 To Len(RTrim$(pVal))
        swVal = Mid$(pVal, swI, 1)
        swInt = Asc(swVal)
        
        swF = (swInt >= 48 And swInt <= 57) _
           Or (swInt >= 65 And swInt <= 90) _
           Or swVal = "-" Or swVal = " " Or swVal = "$" Or swVal = "/" _
           Or swVal = "+" Or swVal = "%" Or swVal = "*" Or swVal = "."
    
        If Not swF Then Exit For
    Next swI
    
    ChrCheck_Code39 = swF
End Function
'---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   ---   < 2007/06/22 ADD END >


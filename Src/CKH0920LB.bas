Attribute VB_Name = "CKH0920LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   平準化エラーリスト
'*
'*   モジュールID  : CKH0920LB
'*
'*                  CKH0920LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not CKH0920LPrint Then Exit Sub
'
'*   作成日  :2004/01/07   By CIS
'*   変更日  :2006/02/02   By CIS - INIファイル-クライアント対応（TEXT出力先）
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCKH0920        As New CISFormPrint
Public CKH0920_PrtCnt       As Long
Public CKH0920_HeiOut       As String
Public CKH0920_HeiOutDir    As String
Public CKH0920_HeiOutFile   As String

Dim bSL_Select              As String

Public Function CKH0920LPrint() As Boolean
    CKH0920LPrint = False
    
    ' 平準化月超過発注
    CKH0920_HeiOut = ""
    CKH0920_HeiOutDir = ""
    
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 6
    Call IRNRead("", 1, 1)
    If CisFun.Val2(Mid(IRN.oyian, 2, 1)) = 0 Then
       CKH0920_HeiOut = "X"
       
       With CisFun
'-----------------------------------------------< 2006/02/02 MNT START >
'            .INI_FileID = gIniExe & gIniName
            .INI_FileID = gIniExe & gwsIniName
'-----------------------------------------------< 2006/02/02 MNT START >
            .INI_Section = "Option"
            '平準化月超過ファイル出力先
            .INI_Key = "HEIFILEOUT": .INI_Default = gIniExe
            .GetIni
            CKH0920_HeiOutDir = RTrim(.INI_String)
            '平準化月超過ファイル名
            .INI_Key = "HEIFILEName": .INI_Default = "平準化月超過"
            .GetIni
            CKH0920_HeiOutFile = RTrim(.INI_String)
            CKH0920_HeiOutFile = CKH0920_HeiOutFile & Format(Date, "YYYYMM") & ".csv"
'-----------------------------------------------< 2006/02/02 ADD START >
            .INI_FileID = gIniExe & gIniName
'-----------------------------------------------< 2006/02/02 ADD START >
       End With
       If Right(CKH0920_HeiOutDir, 1) <> "\" Then CKH0920_HeiOutDir = CKH0920_HeiOutDir & "\"
       CKH0920_HeiOutFile = CKH0920_HeiOutDir & CKH0920_HeiOutFile
    End If
    
    bSL_Select = "SELECT ISNULL(TR.略称,'') 取引先名,"
    bSL_Select = bSL_Select & " CASE WHEN Isnull(HN.表示品番,'') = '' THEN KE.品番 Else HN.表示品番 End 表示品番,"
    bSL_Select = bSL_Select & "  RTRIM(ISNULL(SY.値名称,'')) 手配区分名,"
    bSL_Select = bSL_Select & "  KE.* From 帳票出力テーブル CH"
    bSL_Select = bSL_Select & "  Left Outer Join かんばん読取エラーテーブル KE"
    bSL_Select = bSL_Select & "      ON  KE.エラー管理番号 = CH.管理NO"
    bSL_Select = bSL_Select & "      And KE.エラー種類 = CH.帳票区分"
    bSL_Select = bSL_Select & "  Left Outer Join 品番マスタ HN"
    bSL_Select = bSL_Select & "      ON  HN.品番 = KE.品番"
    bSL_Select = bSL_Select & "  Left Outer Join 取引先マスタ TR"
    bSL_Select = bSL_Select & "      ON  TR.取引先CD = KE.取引先"
    bSL_Select = bSL_Select & "      And TR.取引先区分 = 1"
    bSL_Select = bSL_Select & "  Left Outer Join 名称マスタ SY"
    bSL_Select = bSL_Select & "      ON  SY.区分名称 = '手配区分'"
    bSL_Select = bSL_Select & "      And SY.区分タイプ = ''"
    bSL_Select = bSL_Select & "      And SY.値 = KE.手配区分"
    bSL_Select = bSL_Select & " Where CH.端末番号 = Host_Name()"
    bSL_Select = bSL_Select & "   And CH.帳票区分 = 2"
    
    If CKH0920_HeiOut = "X" Then
       gSL_Select = bSL_Select & " And KE.平準化区分 = 3"
       gSL_Select = gSL_Select & " Order By"
       gSL_Select = gSL_Select & " KE.取引先,KE.受入,KE.背番号,"
       gSL_Select = gSL_Select & " KE.品番,KE.手配区分,KE.収容数,"
       gSL_Select = gSL_Select & " KE.平準化区分,KE.発行NO"
        
       If KRERead(gSL_Select, 0, 1) Then
          Call PutCSV
          Call KREClose(1)
       End If
       gSL_Select = bSL_Select & " And KE.平準化区分 <> 3"
    Else
       gSL_Select = bSL_Select
    End If
    
    gSL_Select = gSL_Select & " Order By"
    gSL_Select = gSL_Select & " KE.取引先,KE.受入,KE.背番号,"
    gSL_Select = gSL_Select & " KE.品番,KE.手配区分,KE.収容数,"
    gSL_Select = gSL_Select & " KE.平準化区分,KE.発行NO"
    
    If Not KRERead(gSL_Select, 0, 1) Then
       Call KREClose(1)
       CKH0920LPrint = True
       Exit Function
    End If
    
    CKH0920_PrtCnt = CisDB.RecordCount
    
    With CisPrtCKH0920
         Set .iPrtForm = CKH0920L
         .dpMSG = "平準化エラーリスト　印刷中"
         .iPreview = gDemoPrt
         .dpMSGScreen = モーダル

         .PrintStart
    End With
    
    CKH0920LPrint = True
End Function
'*---------------------*
'*　  CSV出力処理
'*---------------------*
Private Function PutCSV()
    Dim wFNo            As Integer
    Dim OutStr          As String
    Dim FileFlg         As Boolean
    Dim BK_Torcd        As String
    Dim BK_Ukeir        As String       '受入
    Dim BK_Seban        As String       '背番号
    Dim BK_Hinbn        As String       '品番
    Dim BK_Syuyo        As Long         '収容数
    Dim wKanma          As String
    
    FileFlg = False
    If Dir(CKH0920_HeiOutFile) = "" Then
       FileFlg = True
    End If
    wFNo = FreeFile
    
    Open CKH0920_HeiOutFile For Append As #wFNo
    
    If FileFlg Then
        OutStr = "仕入先コード,仕入先名,受入,背番号,品番,手配区分,"
        OutStr = OutStr & "収容数,設定枚数,発生日,発行NO"
        Print #wFNo, OutStr
    End If
    BK_Torcd = ""
    
    Do Until Not KRE_RDSTS
        With KRE
            If .取引先 <> BK_Torcd Or _
               .受入 <> BK_Ukeir Or _
               .背番号 <> BK_Seban Or _
               .品番 <> BK_Hinbn Or _
               .収容数 <> BK_Syuyo Then
               
                OutStr = .取引先 & ","
                OutStr = OutStr & .取引先名 & ","
                OutStr = OutStr & .受入 & ","
                OutStr = OutStr & .背番号 & ","
                OutStr = OutStr & .表示品番 & ","
                OutStr = OutStr & .手配区分名 & ","
                OutStr = OutStr & .収容数 & ","
                OutStr = OutStr & .平準化設定枚数 & ","
                OutStr = OutStr & Format(Date, "yyyy/mm/dd") & ","
                OutStr = OutStr & Format(.発行NO, "000")
                If BK_Torcd <> "" Then Print #wFNo, OutStr
                wKanma = ","
            Else
                OutStr = OutStr & wKanma
                OutStr = OutStr & Format(.発行NO, "000") & ","
                wKanma = ""
            End If
        '*---------------------------------*
        '*  比較項目　退避
        '*---------------------------------*
            BK_Torcd = .取引先
            BK_Ukeir = .受入
            BK_Seban = .背番号
            BK_Hinbn = .品番
            BK_Syuyo = .収容数
            
            Call KREReadNext(1)
        End With
    Loop
    
    Print #wFNo, OutStr
    
    Close #wFNo
    
End Function




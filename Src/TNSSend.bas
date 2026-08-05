Attribute VB_Name = "TNSSend"
Option Explicit
'+----------------------------------------------------------+
'+
'+       モジュール名  : ＴＮＳ送信用　共通モジュール
'+       モジュールID  : TNSSend.BAS
'+       作成日        : 2005/09/02   CIS
'+       更新日        : 2006/02/08   By CIS  自動化対応
'+                       PARA - /T契約先ｺｰﾄﾞ And /Fﾌｧｲﾙｺｰﾄﾞ  :自動処理
'+                              /D or /Disp    :画面表示あり(自動処理時のみ有効）
'+
'+       変更日        : 2007/03/07   By CIS 品番工場符号削除（ナルコ対応）
'+       変更日        :
'+----------------------------------------------------------+
'---------------------------------------------------- 2006/02/08 Insert
Global Ts_Auto          As Boolean
Global Ts_AutoDisp      As Boolean
Global Ts_ATorcd        As String
Global Ts_AFCode        As String
Global Ts_AEda          As Byte
'---------------------------------------------------- 2006/02/08 Insert End


Global Ts_ProcType      As Byte     ' 送信処理ﾀｲﾌﾟ 0:送信,1:再送信,2:再処理(作成･送信),3:再処理(作成)
Global Ts_Grid          As Object   ' 送信処理ﾀｲﾌﾟ=0(送信)の時、セット
Global Ts_ProcErr       As Boolean  ' 一括処理フラグ
Global Ts_ReSendNo      As Long     ' 再処理用送信管理NO

Global Ts_SendProc      As Boolean  ' 送信処理ﾌﾗｸﾞ True:正常終了,False:異常終了
Global Ts_SndTorcd()    As String   ' 送信先
Global Ts_SndFCode()    As String   ' ﾌｧｲﾙｺｰﾄﾞ
Global Ts_SndKTorcd()   As String   ' 契約先        2007/12/26 ADD
Global Ts_SndSendKb()   As Byte     ' 送信先区分    2007/12/26 ADD
Global Ts_SndRow()      As Long     ' 送信処理ﾀｲﾌﾟ=0(送信)の時のｸﾞﾘｯﾄﾞ行番号
Global Ts_ColResult     As Integer
Global Ts_ColResultCD   As Integer
Global Ts_ColProcData   As Integer
Global Ts_ColSendCnt    As Integer
Global Ts_ColDataCnt    As Integer
Global Ts_ColErrCode    As Integer
Global Ts_ColRNo        As Integer
Global Ts_ColSelect     As Integer
Global Ts_ErrFName      As String

Global Ts_TextProcObj   As Object
Global Ts_ProcListObj   As ListBox
Global Ts_BarObj        As Object
Global Ts_FromObj       As CISFormContorl


Global Ts_SendPath      As String   ' TNS送信ﾃﾞｰﾀ作成先
Global Ts_SendBackPath  As String   ' TNS送信ﾃﾞｰﾀﾊﾞｯｸｱｯﾌﾟ先
Global Ts_SendBackName  As String   ' TNS送信ﾃﾞｰﾀﾊﾞｯｸｱｯﾌﾟﾌｧｲﾙ名(ﾌﾙﾊﾟｽ)
Global Ts_SendFName     As String   ' TNS送信ﾃﾞｰﾀﾌｧｲﾙ名
Global Ts_Jisya         As String   ' 自社コード + 自社工場
Global Ts_JisyaCd       As String   ' 自社コード
Global Ts_JisyaKj       As String   ' 自社工場

Global Ts_Torcd         As String
Global Ts_FCode         As String
Global Ts_KTorcd        As String   ' 契約先        2007/12/26 ADD
Global Ts_SendKb        As Byte     ' 送信区分      2007/12/26 ADD
Global Ts_Length        As Long

Global Ts_DelDay        As Long     ' ログ削除日数
Global Ts_Prefix        As String   ' TNS送信ﾃﾞｰﾀﾌｧｲﾙﾌﾟﾚﾌｨｯｸｽ
Global Ts_TNSPakege     As String   ' TNS通信ﾊﾟｯｹｰｼﾞ
Global Ts_TNSPakegeFlg  As Boolean  ' TNS通信ﾊﾟｯｹｰｼﾞﾌﾗｸﾞ
Global Ts_EDIstart      As String   ' TNSｼｽﾃﾑ起動PG
Global Ts_EDIstartFlg   As Boolean  ' TNSｼｽﾃﾑ起動PG存在確認 True:あり,False:なし
Global Ts_EDIEndPG      As String   ' TNSｼｽﾃﾑ終了PG
Global Ts_EDIEndPGFlg   As Boolean  ' TNSｼｽﾃﾑ終了PG存在確認 True:あり,False:なし

Global Ts_Header        As String
Global Ts_Trailer       As String
Global Ts_Data          As String
Global Ts_FixStr        As String * 1000
Global Ts_TNSPutData    As String
Global Ts_SendCnt       As Long
Global Ts_DataCnt       As Long
Global Ts_FNo           As Integer
Global Ts_FNo2          As Integer

Global Ts_SeqNoAdd      As Byte     ' 送信連番 0:加算有、1:加算無     2008/01/09 ADD
'+-------------------------------------------------- 2007/03/06 Insert
Global Ts_KojyoFix()    As String
'+-------------------------------------------------- 2007/03/06 Insert End


Dim TsStr               As String

'+-------- TNSAPI用ｴﾘｱ
Global Ts_TNSReturn     As String
Global Ts_TNSfunc_seq   As String
Global Ts_TNSFName      As String
Global Ts_Atesaki       As String
Global Ts_Shikibetsu    As String

Public Declare Function File_Send_Api Lib "EDI_USER" (ByVal rc_code As String, _
                                                    ByVal func_seq As String, _
                                                    ByVal async_code As String, _
                                                    ByVal apl_cnv As String, _
                                                    ByVal file_code As String, _
                                                    ByVal atesaki_name As String, _
                                                    ByVal send_mode As String, _
                                                    ByVal file_name As String, _
                                                    ByVal file_ext As String, _
                                                    ByVal file_type As String, _
                                                    ByVal compact As String, _
                                                    ByVal a_code As String, _
                                                    ByVal code_conv As String, _
                                                    ByVal apl_len As String, _
                                                    ByVal user_id As String, _
                                                    ByVal yobi As String _
                                                    ) As Long

Public Declare Function File_Send_Ex_Api Lib "EDI_USER" (ByVal rc_code As String, _
                                                    ByVal func_seq As String, _
                                                    ByVal async_code As String, _
                                                    ByVal apl_cnv As String, _
                                                    ByVal file_code As String, _
                                                    ByVal atesaki_name As String, _
                                                    ByVal send_mode As String, _
                                                    ByVal file_name As String, _
                                                    ByVal file_ext As String, _
                                                    ByVal file_type As String, _
                                                    ByVal compact As String, _
                                                    ByVal a_code As String, _
                                                    ByVal code_conv As String, _
                                                    ByVal apl_len As String, _
                                                    ByVal user_id As String, _
                                                    ByVal yobi As String _
                                                    ) As Long
'---------------------------------------------------- 2006/02/08 Insert
'+---------------------------------+
'+    初期処理
'+---------------------------------+
Sub Main()
'   #-------------------#
'   # 二 重 起 動 防 止 #
'   #-------------------#
    If Not Dupli_Start Then End
'   #-------------------#
'   # 初 期 内 容 取 得 #
'   #-------------------#
    If Not IniGet Then End
    
    Ts_Auto = False
    Ts_AutoDisp = False
    Ts_ATorcd = ""
    Ts_AFCode = ""
    Ts_AEda = 0
    
    If InStr(1, UCase(Command), "/D") > 0 Or _
       InStr(1, UCase(Command), "/DISP") > 0 Then
       Ts_AutoDisp = True
    End If
    
    '--------------- 契約先取得
    gCnt1 = InStr(1, UCase(Command), "/T")
    If gCnt1 > 0 Then
       gCnt2 = InStr(gCnt1, UCase(Command), " ")
       If gCnt2 = 0 Then
          gCnt2 = Len(Command)
       End If
       gStr = Mid(UCase(Command), gCnt1, gCnt2)
       Ts_ATorcd = RTrim(Replace(gStr, "/T", ""))
    End If
    '--------------  ﾌｧｲﾙｺｰﾄﾞ取得
    gCnt1 = InStr(1, UCase(Command), "/F")
    If gCnt1 > 0 Then
       gCnt2 = InStr(gCnt1, UCase(Command), " ")
       If gCnt2 = 0 Then
          gCnt2 = Len(Command) - gCnt1 + 2
       Else
          gCnt2 = gCnt2 + 1 - gCnt1
       End If
       gStr = Mid(UCase(Command), gCnt1, gCnt2)
       Ts_AFCode = RTrim(Replace(gStr, "/F", ""))
    End If
    '-------------- TNS設定確認
    If Trim(Ts_ATorcd) <> "" And Trim(Ts_AFCode) <> "" Then
       
       gSL_Select = "Select * From TNS送信設定テーブル"
       gSL_Select = gSL_Select & " Where 送信先 = '" & RTrim(Ts_ATorcd) & "'"
       gSL_Select = gSL_Select & " And   FCODE = '" & RTrim(Ts_AFCode) & "'"
       gSL_Select = gSL_Select & " And   枝番 = " & Ts_AEda
       If TSSRead(gSL_Select, 1) Then
          Ts_Auto = True
       End If
    End If
    '------------------------------------------------------------------ 2007/03/06 Insert
    ReDim Ts_KojyoFix(1)
    gSL_Select = "Select * From 名称マスタ"
    gSL_Select = gSL_Select & " Where 区分名称 = '自社工区'"
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " Order By 値"

    If SYMRead(gSL_Select) Then
       ReDim Ts_KojyoFix(CisDB.RecordCount)
       gLong = 0
       Do Until Not SYM_RDSTS
                        
            Ts_KojyoFix(gLong) = RTrim(SYM.英数字1)
            gLong = gLong + 1
            
            Call SYMReadNext
       Loop
    End If
    Call SYMClose
    
    '------------------------------------------------------------------ 2007/03/06 Insert End
    '------------- 自動処理
    If Ts_Auto Then
       Ts_ProcType = 0
       Ts_ReSendNo = 0
       Set Ts_Grid = Nothing
       ReDim Ts_SndTorcd(1)
       ReDim Ts_SndFCode(1)
       ReDim Ts_SndRow(1)

       Ts_SndTorcd(0) = RTrim(Ts_ATorcd)
       Ts_SndFCode(0) = RTrim(Ts_AFCode)
    '
       Ts_ProcErr = False
        
       CXT0020.Show vbModal
       Unload CXT0020
       Set CXT0020 = Nothing
    '
'    '------------------------------------------------------------------------------- 2007/03/09 Insert
    '+-----------> 品番ﾁｪｯｸｴﾗｰﾘｽﾄ発行
        If TSS.TNSタイプ = 1 Or TSS.TNSタイプ = 2 Then
           If TSS.品番CHK = 1 And TSS.品番ERRLIST = 1 And C04_HinErrCnt <> 0 Then
              Call C04_ErrList
           End If
        End If
'    '------------------------------------------------------------------------------- 2007/03/09 Insert End
        
       Call CisDB.DBDISConnect
       End
    End If
    
    '------------- 手動処理
    CXT0010.Show
    
End Sub
'---------------------------------------------------- 2006/02/08 Insert End
'+---------------------------------+
'+    初期処理
'+---------------------------------+
Public Sub TnsSend_Init(Optional StartMsg As Boolean = True)
    With CisFun
        ' 変換ログ削除日数
        .INI_Section = "TNS送信"
        .INI_Default = "10"
        .INI_Key = "LogDelDay"
        .GetIni
        Ts_DelDay = .Val2(.INI_String)
        ' 変換ログ出力ファイル設定
        .Log_PutFlg = True
        .Log_FileInitial = "TNS_Send"
        .Log_FilePath = gIniExe & "Log\"
        .Log_DataHD = ""
        .Log_DataHDTime = あり
        .Log_DelDay = Ts_DelDay

        .LogPut ""
        ' TNSｼｽﾃﾑ起動PG
        .INI_Section = "TNS送信"
        .INI_Default = ""
        .INI_Key = "Edistart"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "C:\Edi\Exe\Edistart.exe"
           .PutIni
        End If
        Ts_EDIstart = Trim(.INI_String)
        
        If Dir(Ts_EDIstart) = "" Then
           Ts_EDIstartFlg = False
        Else
           Ts_EDIstartFlg = True
        End If
        
        ' TNSｼｽﾃﾑ終了PG
        .INI_Section = "TNS送信"
        .INI_Default = ""
        .INI_Key = "EDIEndPG"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "C:\Program Files\TNS\EDIG07.exe"
           .PutIni
        End If
        Ts_EDIEndPG = Trim(.INI_String)
        If Dir(Ts_EDIEndPG) = "" Then
           Ts_EDIEndPGFlg = False
        Else
           Ts_EDIEndPGFlg = True
        End If
        ' TNS通信ﾊﾟｯｹｰｼﾞ
        .INI_Section = "TNS送信"
        .INI_Default = ""
        .INI_Key = "Package"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = .WinRoot & "Edi_user.dll"
           .PutIni
        End If
        Ts_TNSPakege = Trim(.INI_String)
        If Dir(Ts_TNSPakege) = "" Then
           Ts_TNSPakegeFlg = False
        Else
           Ts_TNSPakegeFlg = True
        End If
        
        ' 開始ログ出力
        If StartMsg Then
           Call TNSSend_LogPut("S")
        End If
    End With
'--( 自社コード )
    Ts_JisyaCd = Left(gCompany, 4)
'--( JAMA自社工場・JAMA機番 ) [システム情報]
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 8
    Call IRNRead("", 1)
    Ts_JisyaKj = Mid(IRN.oyian, 1, 1)


    Ts_Jisya = Ts_JisyaCd & Ts_JisyaKj
End Sub
'+---------------------------------------------------+
'+
'+      TNS送信　開始・終了　ログ出力
'+
'+---------------------------------------------------+
Public Sub TNSSend_LogPut(PutType As String, Optional PutMsg As String, Optional SpaceCnt As Long = 0)

    With CisFun
        ' 開始ログ
        If PutType = "S" Then
           TsStr = "↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓": .LogPut TsStr
           TsStr = "↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓": .LogPut TsStr
           .LogPut "|"
           TsStr = "/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_": .LogPut TsStr
           TsStr = "/_/_/                                       /_/_/": .LogPut TsStr
           TsStr = "/_/_/     ＴＮＳ送信データ作成処理 開始       /_/_/": .LogPut TsStr
           
           If Ts_Auto Then
              TsStr = "/_/_/                                       /_/_/": .LogPut TsStr
              TsStr = "/_/_/          自 動 処 理                  /_/_/": .LogPut TsStr
              TsStr = "/_/_/                                       /_/_/": .LogPut TsStr
           End If
           
           Select Case Ts_ProcType
                Case 2
                     TsStr = "/_/_/           【 再 送 信 】               /_/_/": .LogPut TsStr
                Case 3
                     TsStr = "/_/_/       【 再処理(作成・送信) 】          /_/_/": .LogPut TsStr
                Case 4
                     TsStr = "/_/_/          【 再処理(作成) 】            /_/_/": .LogPut TsStr
           End Select
           TsStr = "/_/_/                                       /_/_/": .LogPut TsStr
           TsStr = "/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_": .LogPut TsStr
           .LogPut "|"
           If Not Ts_EDIstartFlg Then
                TsStr = " ＴＮＳシステム起動ＰＧ未登録 ": .LogPut TsStr
           End If
           If Not Ts_EDIEndPGFlg Then
                TsStr = " ＴＮＳシステム終了ＰＧ未登録 ": .LogPut TsStr
           End If
           .LogPut "|"
           Exit Sub
        End If
        If PutType = "E" Then
           .LogPut "|"
           TsStr = "/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_": .LogPut TsStr
           TsStr = "/_/_/                                       /_/_/": .LogPut TsStr
           TsStr = "/_/_/     ＴＮＳ送信データ作成処理 終了       /_/_/": .LogPut TsStr
           Select Case Ts_ProcType
                Case 2
                     TsStr = "/_/_/           【 再 送 信 】               /_/_/": .LogPut TsStr
                Case 3
                     TsStr = "/_/_/       【 再処理(作成・送信) 】          /_/_/": .LogPut TsStr
                Case 4
                     TsStr = "/_/_/          【 再処理(作成) 】            /_/_/": .LogPut TsStr
           End Select
           TsStr = "/_/_/                                       /_/_/": .LogPut TsStr
           TsStr = "/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_": .LogPut TsStr
           .LogPut "|"
           TsStr = "↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑": .LogPut TsStr
           TsStr = "↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑": .LogPut TsStr
           Exit Sub
        End If
        If PutType = "ER" Then
           .LogPut "|"
           TsStr = "* Err * Err * Err * Err * Err * Err * Err * Err * Err": .LogPut TsStr
           .LogPut "|"
        End If
        If SpaceCnt <> 0 Then
           .LogPut Space(SpaceCnt) & PutMsg
        Else
           .LogPut PutMsg
        End If
    End With
End Sub
'+---------------------------------------------------+
'+
'+      処理進捗　メッセージ表示
'+
'+---------------------------------------------------+
Public Sub TNSSend_MsgDisp(ProcMsg As String, DispMsg As String, Optional SpaceCnt As Long = 0, Optional PutType As String = "")
    
    If Trim(ProcMsg) <> "" Then
       Call TNSSend_LogPut(PutType, ProcMsg, SpaceCnt)
       If Not Ts_TextProcObj Is Nothing Then
          Ts_TextProcObj = ProcMsg
          DoEvents
       End If
    End If
    
    If Trim(DispMsg) <> "" Then
       If ProcMsg <> DispMsg Then
          Call TNSSend_LogPut(PutType, DispMsg, SpaceCnt)
       End If
       If Not Ts_ProcListObj Is Nothing Then
          Ts_ProcListObj.AddItem DispMsg
          DoEvents
       End If
    End If
End Sub
'*-----------------------------------------------------------------------------------*
'*
'*      TNS送信バックアップ削除
'*
'*              SendTorcd  : 送信先
'*              SendFcode  : ファイルコード
'*              BackPath   : バックアップ格納先
'*              DelDay     : 削除日数
'*
'*-----------------------------------------------------------------------------------*
Public Function TNSSend_BackKill(SendTorcd As String, SendKTorcd As String, SendFcode As String, DelDay As Byte) As Boolean
    TNSSend_BackKill = False
    
    Static sw_YMD               As String
    Static sw_DIR               As String
    Static sw_FILE              As String
    Static sw_FILEYMD           As String
    
    On Error GoTo TNSSend_BackKill_Er
    
    With CisFun
        Ts_Prefix = SendTorcd
        If RTrim(SendKTorcd) <> "" And RTrim(SendTorcd) <> RTrim(SendKTorcd) Then
            Ts_Prefix = Ts_Prefix & "_" & SendKTorcd
        End If
        Ts_Prefix = Ts_Prefix & "_" & SendFcode
        
        .INI_FileID = gIniExe & gwsIniName
        ' TNS送信ﾃﾞｰﾀ作成先ﾌｫﾙﾀﾞを求める
        .INI_Section = "TNSSENDDIR"
        .INI_Default = ""
        If TSS.作成先区分1 = 0 Then
           .INI_Key = "COMMON"
        Else
           .INI_Key = Ts_Prefix
        End If
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = gIniExe & "TNS\Send\"
        End If
        Ts_SendPath = Trim(.INI_String)
        ' TNS送信ﾃﾞｰﾀﾊﾞｯｸｱｯﾌﾟ作成先ﾌｫﾙﾀﾞを求める
        .INI_Section = "TNSBACKDIR"
        .INI_Default = ""
        If TSS.作成先区分2 = 0 Then
           .INI_Key = "COMMON"
        Else
           .INI_Key = Ts_Prefix
        End If
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = gIniExe & "TNS\Back\"
        End If
        Ts_SendBackPath = Trim(.INI_String)
        ' フォルダ存在確認　存在しない時、自動作成
        .FE_Path = Ts_SendPath
        .DirCheck NoMsg_Crt
        
        .FE_Path = Ts_SendBackPath
        .DirCheck NoMsg_Crt
        
        .INI_FileID = gIniExe & gIniName

    End With
    
    If DelDay = 0 Then
       TNSSend_BackKill = True
       Exit Function
    End If
    
    If Right(Ts_SendBackPath, 1) <> "\" Then
       Ts_SendBackPath = Ts_SendBackPath & "\"
    End If
    
    sw_YMD = Format(DateAdd("D", DelDay * -1, Now()), "YYYYMMDD")
    sw_DIR = Ts_SendBackPath & Ts_Prefix & "*.txt"
        
    Call TNSSend_LogPut("", "削除フォルダ = " & Ts_SendBackPath, 10)
    Call TNSSend_LogPut("", "削除日付 = " & sw_YMD & "より前の日付(開始日付)", 10)
        
    On Error Resume Next
    sw_FILE = RTrim$(Dir(sw_DIR))
    Do While RTrim$(sw_FILE) <> ""
        sw_FILEYMD = Replace(sw_FILE, Ts_Prefix & "_", "")
        
        If Mid$(sw_FILEYMD, 1, 8) < sw_YMD Then
           Kill Ts_SendBackPath & sw_FILE
        End If
       
       sw_FILE = RTrim$(Dir)
    Loop
    
    With CisDB
        .SQL = "Delete TNS支給送信変換エラーテーブル"
        .SQL = .SQL & " FROM TNS送信履歴テーブル TE "
        .SQL = .SQL & " Where TNS支給送信変換エラーテーブル.履歴管理NO = TE.履歴管理NO"
        .SQL = .SQL & "   And TE.送信先 = '" & RTrim(SendKTorcd) & "'"
        .SQL = .SQL & "   And TE.FCODE = '" & RTrim(SendFcode) & "'"
        .SQL = .SQL & "   AND TE.HTYPE = ''"
        .SQL = .SQL & "   AND Convert(char(8),TE.開始日時,112) < '" & sw_YMD & "'"
        If Not .DBExec Then
           Call TNSSend_MsgDisp("", "NS支給送信変換ｴﾗｰﾃｰﾌﾞﾙ 削除ｴﾗｰ" & vbCrLf & .ErrMsg, 10, "ER")
           Exit Function
        End If
    
        .SQL = "Delete TNS送信履歴テーブル "
        .SQL = .SQL & " Where 送信先 = '" & RTrim(SendTorcd) & "'"
        .SQL = .SQL & "   And FCODE = '" & RTrim(SendFcode) & "'"
        .SQL = .SQL & "   And 契約先 = '" & RTrim(SendKTorcd) & "'"
        .SQL = .SQL & "   AND HTYPE = ''"
        .SQL = .SQL & "   AND Convert(char(8),開始日時,112) < '" & sw_YMD & "'"
        .ConnectNo = 0
        If Not .DBExec Then
           Call TNSSend_MsgDisp("", "TNS送信履歴テーブル 削除エラー" & vbCrLf & .ErrMsg, 10, "ER")
           Exit Function
        End If
    End With
    
    TNSSend_BackKill = True
    Exit Function
TNSSend_BackKill_Er:
    gStr = "TNS送信バックアップ削除エラー" & vbCrLf & _
           "ｴﾗｰ番号 " & Str(ERR.Number) & vbCrLf & _
            ERR.Source & " でｴﾗｰが発生しました。" & vbCrLf & _
            ERR.Description
    Call TNSSend_MsgDisp("", gStr, 10, "ER")
End Function
'*-----------------------------------------------------------------------------------*
'*
'*      TNS送信ヘッダーレコード作成
'*
'*-----------------------------------------------------------------------------------*
Public Sub TNSSend_HeaderCrt()
'--------------------------------------- 送信連番 加算      2008/01/09MNT
    If Ts_SeqNoAdd = 0 Then
        TSR.送信連番 = TSR_SEQNO + 1
    Else
        TSR.送信連番 = TSR_SEQNO
    End If
'---------------------------------------
   
    Ts_Header = "H"
    Ts_FixStr = TSS.宛先
    Ts_Header = Ts_Header & Mid(Ts_FixStr, 1, 5)
    Ts_FixStr = TSS.発信元
    Ts_Header = Ts_Header & Mid(Ts_FixStr, 1, 5)
    Ts_FixStr = TSS.FCODE
    Ts_Header = Ts_Header & Mid(Ts_FixStr, 1, 3)
    Ts_FixStr = Format(Now(), "YYMMDD")
    Ts_Header = Ts_Header & Mid(Ts_FixStr, 1, 6)
    Ts_FixStr = Format(TSR.送信連番, "00")
    Ts_Header = Ts_Header & Mid(Ts_FixStr, 1, 2)
    Ts_FixStr = Format(TSS.レコード長, "00000")
    Ts_Header = Ts_Header & Mid(Ts_FixStr, 1, 5)
    If TSS.TNSフォーマット = 1 Then
       Ts_Header = Ts_Header & Space(53) & "0000000    200"
       Ts_Length = 1000
    Else
       Ts_Length = 250
    End If
    Ts_FixStr = Ts_Header
    Ts_Header = Mid(Ts_FixStr, 1, Ts_Length)
    
    Ts_TNSPutData = Ts_Header
    
End Sub
'*-----------------------------------------------------------------------------------*
'*
'*      TNS送信トレーラーレコード作成
'*
'*-----------------------------------------------------------------------------------*
Public Sub TNSSend_TrailerCrt()
    
    Ts_Trailer = "T"
    If TSS.TNSフォーマット = 1 Then
        Ts_FixStr = Format(Ts_SendCnt, "0000000")
        Ts_Trailer = Ts_Trailer & Mid(Ts_FixStr, 1, 7)
    Else
        Ts_FixStr = Format(Ts_SendCnt, "00000")
        Ts_Trailer = Ts_Trailer & Mid(Ts_FixStr, 1, 5)
    End If
    Ts_FixStr = Ts_Trailer
    Ts_Trailer = Mid(Ts_FixStr, 1, Ts_Length)
    
    Ts_TNSPutData = Ts_TNSPutData & Ts_Trailer
    
End Sub
'*-----------------------------------------------------------------------------------*
'*
'*      TNS送信データ出力
'*      2006/03/09 ファイル名空白時、ﾌｧｲﾙｺｰﾄﾞ+年月日字分秒　をファイル名とする
'*-----------------------------------------------------------------------------------*
Public Function TNSSend_DataPut() As Boolean
    TNSSend_DataPut = False
    On Error GoTo TNSSend_DataPut_Er
    

    If Right(Ts_SendPath, 1) <> "\" Then
       Ts_SendPath = Ts_SendPath & "\"
    End If
    If Right(Ts_SendBackPath, 1) <> "\" Then
       Ts_SendBackPath = Ts_SendBackPath & "\"
    End If

    Ts_SendBackName = Ts_SendBackPath & Ts_Prefix & Format(Now(), "_YYYYMMDDhhnnss.txt")
    
    '[][][][][][][][][][]2006/03/09[][][][][][][][][][][][][][][][][][][][][][][][][][][][]
    If Trim(Ts_SendFName) = "" Then
       Ts_SendFName = Ts_FCode & Format(Now(), "_YYYYMMDDhhnnss") & ".txt"
    End If
    '[][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][]
    
    On Error Resume Next

    Kill Ts_SendPath & Ts_SendFName

    On Error GoTo TNSSend_DataPut_Er
    Ts_FNo = FreeFile

    Open Ts_SendPath & Ts_SendFName For Binary As #Ts_FNo
    
    Ts_FNo2 = FreeFile
    
    Open Ts_SendBackName For Binary As #Ts_FNo2
    
    Put #Ts_FNo, , Ts_TNSPutData
    Put #Ts_FNo2, , Ts_TNSPutData
    
    Close #Ts_FNo
    Close #Ts_FNo2
    
    TNSSend_DataPut = True
    Exit Function

TNSSend_DataPut_Er:
    gStr = "TNS送信データ出力ｴﾗｰ" & vbCrLf & _
           "ｴﾗｰ番号 " & Str(ERR.Number) & vbCrLf & _
            ERR.Source & " でｴﾗｰが発生しました。" & vbCrLf & _
            ERR.Description
    Call TNSSend_MsgDisp("", gStr, 10, "ER")
End Function
'
''+---------------------------------------+
''+    TNS送信処理呼出
''+---------------------------------------+
'Public Function TnsSendProc() As Boolean
'
'    TnsSendProc = False
'
'    If Not TsEdiFlg Then
'        Call TsFileCtl.cLogPut("新TNSｼｽﾃﾑ 通信パッケージ 未登録")
'        ReDim gMBMsg(5)
'        gMBMsg(1) = "                 新TNSｼｽﾃﾑ 通信パッケージ 未登録       "
'        gMBMsg(3) = "  通信パッケージをインストール後、再度処理を行って下さい！！      "
'        gMBTitle = ""
'        gMBButton = "D"
'        Call MBOX
'
'        Exit Function
'    End If
'
'    CTNS0700.Show vbModal
'    Unload CTNS0700
'    Set CTNS0700 = Nothing
'
'    TnsSendProc = True
'
'End Function
'
'*******************************************************************************
'
'   ＜概要＞
'       ・このモジュールは、ＴＮＳバッチシステムで提供するＡＰＩの呼び出し方法を
'       　Visual Basicでコーディングを行ったサンプルです。
'
'   ＜注意＞
'       ・このモジュール内でパラメータに設定した値は、あくまでも例であり固定的な
'       　値では有りません。
'       ・ＡＰＩの機能及び使用方法は、ＡＰＩの解説書をご覧ください。
'
'
'*******************************************************************************

'*******************************************************************************
'
'   ファイル送信
'
'*******************************************************************************
Public Function TNS_File_Send_Api() As Boolean
    Dim ret             As Boolean          '戻り値
    Dim api_ret         As Long             'ＡＰＩ戻り値

    Dim rc_code         As String * 6       '処理結果
    Dim func_seq        As String * 8       '処理シーケンス
    Dim async_code      As String * 1       '同期／非同期
    Dim apl_cnv         As String * 1       'ヘッダー作成
    Dim file_code       As String * 8       'メッセージ識別子
    Dim atesaki_name    As String * 23      '宛先
    Dim file_name       As String * 128     '送信ファイル名
    Dim send_mode       As String * 1       '上書き送信
    Dim file_ext        As String * 3       '送信ファイル拡張子
    Dim file_type       As String * 1       'ファイル種別
    Dim compact         As String * 1       '圧縮
    Dim a_code          As String * 1       '暗号化
    Dim code_cnv        As String * 1       'コード変換
    Dim apl_len         As String * 5       'ＡＰＬ有効長
    Dim user_id         As String * 16      'ユーザーＩＤ
    Dim yobi            As String * 32      '予備

    TNS_File_Send_Api = False
    Ts_TNSReturn = ""
    Ts_TNSfunc_seq = ""
    On Error GoTo TNS_File_Send_Api_Error

'****************************************
'   変数初期化
'****************************************
    ret = True
    rc_code = ""

'****************************************
'   パラメータ設定
'****************************************

    '========================================
    '処理シーケンス
    '   ・ユニークな英数字
    '========================================
    func_seq = ""

    '========================================
    '同期／非同期
    '   '1' ･･･ 非同期＆通常
    '   '2' ･･･ 同期＆通常
    '   '3' ･･･ 非同期＆優先
    '   '4' ･･･ 同期＆優先
    '========================================
    async_code = "4"

    '========================================
    'ヘッダー作成
    '   'Y' ･･･ ヘッダーを作成する
    '   'N' ･･･ ヘッダーを作成しない
    '========================================
    apl_cnv = "N"

    '========================================
    'メッセージ識別子
    '   ・最大桁数に満たない場合は、NULLをセット
    '========================================
    file_code = RTrim(Ts_Shikibetsu) & Chr(0)
    '========================================
    '宛先
    '   ・最大桁数に満たない場合は、NULLをセット
    '========================================
    atesaki_name = RTrim(Ts_Atesaki) & Chr(0)

    '========================================
    '上書き送信
    '   'Y' ･･･ 上書き送信
    '   'N' ･･･ 通常送信
    '========================================
    send_mode = "N"

    '========================================
    '送信ファイル名
    '   ・最大桁数に満たない場合は、NULLをセット
    '========================================
    file_name = Ts_TNSFName & Chr(0)
    '========================================
    '送信ファイル拡張子
    '   ・最大桁数に満たない場合は、NULLをセット
    '           省略時、転送制御マスタよりセット
    '========================================
    file_ext = ""

    '========================================
    'ファイル種別
    '   '1' ･･･ ＴＮＳ形式
    '   '2' ･･･ ＴＮＳ拡張形式
    '   '3' ･･･ 標準形式
    '   '4' ･･･ 自由形式
    '           省略時、転送制御マスタよりセット
    '========================================
    file_type = ""

    '========================================
    '圧縮
    '   'Y' ･･･ 圧縮
    '   'N' ･･･ 圧縮なし
    '           省略時、転送制御マスタよりセット
    '========================================
    compact = "N"

    '========================================
    '暗号化
    '   'Y' ･･･ 暗号化
    '   'N' ･･･ 暗号化なし
    '           省略時、転送制御マスタよりセット
    '========================================
    a_code = ""

    '========================================
    'コード変換
    '   'Y' ･･･ コード変換
    '   'N' ･･･ コード変換不要
    '           省略時、転送制御マスタよりセット
    '========================================
    code_cnv = ""

    '========================================
    'ＡＰＬ有効長
    '   ・５桁の領域必須
    '           省略時、転送制御マスタよりセット
    '========================================
    apl_len = ""

    '========================================
    'ユーザーＩＤ
    '========================================
    user_id = ""

    '========================================
    '予備
    '========================================
    yobi = ""

'****************************************
'   ＡＰＩ呼び出し
'****************************************
    api_ret = File_Send_Api(rc_code, _
                            func_seq, _
                            async_code, _
                            apl_cnv, _
                            file_code, _
                            atesaki_name, _
                            send_mode, _
                            file_name, _
                            file_ext, _
                            file_type, _
                            compact, _
                            a_code, _
                            code_cnv, _
                            apl_len, _
                            user_id, _
                            yobi)

    If api_ret <> 0& Then ret = False


    If ret = False Or rc_code <> "000000" Then
       Ts_TNSReturn = rc_code
'        MsgBox "ＡＰＩ呼び出しエラー"
    Else
       Ts_TNSfunc_seq = func_seq
    End If

    TNS_File_Send_Api = ret

    Exit Function

TNS_File_Send_Api_Error:
    ret = False

End Function

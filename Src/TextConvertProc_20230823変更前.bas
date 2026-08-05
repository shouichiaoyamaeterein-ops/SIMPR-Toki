Attribute VB_Name = "TextConvert"
Option Explicit
'***********************************************************************************
'       モジュール名  : ﾃｷｽﾄ連携連携
'       モジュールID  : JTextProc.BAS
'       使用法        :
'       作成日        : 2009/09/11   CIS
'       変更日        : 2010/03/17   CIS  愛知機械:出荷(Prefix 2)を追加
'       変更日        : 2010/04/19   CIS  ヨロズ愛知:受注(Prefix 3)を追加
'       変更日        : 2010/07/23   CIS  リスト再発行の条件を修正
'
'***********************************************************************************
Public Text_Proc           As Byte     '0:受注[愛知機械],1:受注[テクノエイト],2:出荷[愛知機械],3:受注[ヨロズ愛知]
Public Text_ProcKB         As Byte     '0:自動,1:手動
Public Text_ProcDisp       As Byte     '0:画面表示なし,1:画面表示あり

Public Text_Prefix(3)      As String   ' ﾌﾟﾚﾌｨｯｸｽ (0:JY(愛知機械),1:JY(ﾃｸﾉｴｲﾄ),2:SK(愛知機械),3:JY(ﾖﾛｽﾞ愛知) )
Public Text_PrefixNm(3)    As String   ' ﾌﾟﾚﾌｨｯｸｽ名 (0:受注,1:受注,2:出荷)
Public Text_InDir(3)       As String   ' 端末ﾃｷｽﾄ格納先
Public Text_TextName(3)    As String   ' 端末ﾃｷｽﾄ名(ﾜｲﾙﾄﾞｶｰﾄﾞ指定可能)
Public Text_OutDir(3)      As String   ' SIMPR-A ﾃﾞｰﾀ退避先
Public Text_DelDay(3)      As Long     ' 保持期間(ﾃｷｽﾄﾃﾞｰﾀ･ﾛｸﾞ･履歴情報)
Public Text_CvtList(3)     As Boolean  ' 変換ﾘｽﾄ発行ﾌﾗｸﾞ(True:発行あり,False:発行なし)
Public Text_CvtErr(3)      As Byte     '   ｴﾗｰ内容    0:対象、1:対象外
Public Text_CvtDbl(3)      As Byte     '   重複       0:対象、1:対象外
Public Text_CvtKari(3)     As Byte     '   仮登録内容 0:対象、1:対象外
Public Text_CvtEtc(3)      As Byte     '   その他     0:対象、1:対象外

Public Text_Orderer(3)     As String   ' 受注先（愛知機械/ﾃｸﾉｴｲﾄ/ﾖﾛｽﾞ愛知）

'--------------------------------------------- 有効時間チェック
Public Text_TimeChk        As Boolean  ' 有効時間設定ﾌﾗｸﾞ(True:行う,False:行わない)
Public Text_SHHMM          As String   ' 開始時刻
Public Text_EHHMM          As String   ' 終了時刻
Public Text_HHMM           As String   ' 現在時刻

Public Text_ErrMsg         As String
Public Text_DirMaxCnt      As Long
Public Text_ProcYMD        As String
Public Text_ProcCNT        As Long
Public Text_Rireki         As Long     ' 履歴管理番号(再処理用)

Public Text_MaeKoteiFile   As String   '前工程情報出力ファイル名
Public Text_MaeKoteiFld    As String   '前工程情報出力フォルダ

Public Text_SBasyoCD       As String   '設置場所コード
Public Text_MachineNo      As String   'マシンNO

Public Text_Fun            As New CISFunction
Public Text_Frm            As CISFormContorl
Public Text_FunFlg         As Boolean

Public DispMsg()            As String

Public Text_Prt            As New CISFormPrint

Public Text_FNo            As Integer
Public Text_Rec            As String

' 文字列取出
Global gString()        As String

'出荷日
Public Text_YMD            As String                '2010/05/31

'+-----------------------------------------+
'+  ﾃｷｽﾄ連携有効時間チェック
'+-----------------------------------------+
Public Function JTextYukou() As Boolean
    JTextYukou = False

    With CisFun
        .INI_Section = "JText"
        .INI_FileID = gIniExe & gIniName
        '-------- 有効時間設定
        .INI_Key = "有効時間"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "Yes"
           .PutIni
        End If
        '----------------- 有効時間ﾁｪｯｸ 行わない
        If UCase(Trim(.INI_String)) <> "Y" And UCase(Trim(.INI_String)) <> "YES" Then
           JTextYukou = True
           Exit Function
        End If
        '------------------------------------- 現在時刻セット
        Text_HHMM = Format(Time(), "hhmm")
        
        '-------- 有効時間(開始)
        .INI_Key = "有効時間_開始_時"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "5"
           .PutIni
        End If
        Text_SHHMM = Format(CisFun.Val2(.INI_String), "00")
        .INI_Key = "有効時間_開始_分"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "00"
           .PutIni
        End If
        Text_SHHMM = Text_SHHMM & Format(CisFun.Val2(.INI_String), "00")
        
        '-------- 有効時間(終了)
        .INI_Key = "有効時間_終了_時"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "5"
           .PutIni
        End If
        Text_EHHMM = Format(CisFun.Val2(.INI_String), "00")
        .INI_Key = "有効時間_終了_分"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "00"
           .PutIni
        End If
        Text_EHHMM = Text_EHHMM & Format(CisFun.Val2(.INI_String), "00")
    End With

    '---------------------- 判定
    ' 開始 > 終了
    If Text_SHHMM > Text_EHHMM Then
       If Text_EHHMM < Text_HHMM And _
          Text_SHHMM > Text_HHMM Then
       Else
          JTextYukou = True
       End If
       Exit Function
    End If
    ' 開始 < 終了
    If Text_SHHMM <= Text_HHMM And _
       Text_EHHMM >= Text_HHMM Then
       JTextYukou = True
       Exit Function
    End If
    

End Function
'+-----------------------------------------+
'+  ﾃｷｽﾄ連携環境初期情報取得
'+-----------------------------------------+
Public Function JTextIniGet() As Boolean

    JTextIniGet = False
    ' 処理名をセット
    Text_Prefix(0) = "JY": Text_PrefixNm(0) = "受注"
    Text_Prefix(1) = "JY": Text_PrefixNm(1) = "受注"
    Text_Prefix(2) = "SK": Text_PrefixNm(2) = "出荷"
    Text_Prefix(3) = "JY": Text_PrefixNm(3) = "受注"
    
    ' 受注先をセット
    Text_Orderer(0) = "_愛知機械"
    Text_Orderer(1) = "_テクノエイト"
    Text_Orderer(2) = "_愛知機械"
    Text_Orderer(3) = "_ヨロズ愛知"
    
    Text_YMD = ""                                           '2010/05/31
    
    ' ログ情報
    With CisFun
        .Log_PutFlg = True
        .Log_FilePath = gIniExe & "Log"
        .Log_FileInitial = "Text_" & Text_Prefix(Text_Proc) & "_"
        .Log_DataHDTime = あり
    End With
    
    ' ﾃｷｽﾄ連携設定環境取得
    With CisFun
        .INI_Section = "JText"

'------------------- ｸﾗｲｱﾝﾄ個別設定
        .INI_FileID = gIniExe & gwsIniName
'
        '-------- ﾃｷｽﾄ連携ﾃｷｽﾄ格納先
        .INI_Key = Text_PrefixNm(Text_Proc) & Text_Orderer(Text_Proc) & "_ﾌｧｲﾙ格納先"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = " "
           .PutIni
        End If
        Text_InDir(Text_Proc) = Trim(.INI_String)
        
        If Trim(Text_InDir(Text_Proc)) <> "" Then
           If Right(Text_InDir(Text_Proc), 1) <> "\" Then
              Text_InDir(Text_Proc) = Text_InDir(Text_Proc) & "\"
           End If
        End If
        '-------- ﾃｷｽﾄ連携ﾃｷｽﾄ退避先(SIMPR-A)
        .INI_Key = Text_PrefixNm(Text_Proc) & Text_Orderer(Text_Proc) & "_ﾌｧｲﾙ退避先"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           gStr = Replace(UCase(gIniExe), "EXE\", "JText\Back\")
           .INI_String = gStr & Text_Orderer(Text_Proc) & "_" & Text_PrefixNm(Text_Proc) & "\"
           .PutIni
        End If
        Text_OutDir(Text_Proc) = Trim(.INI_String)
        If Trim(Text_OutDir(Text_Proc)) <> "" Then
           If Right(Text_OutDir(Text_Proc), 1) <> "\" Then
              Text_OutDir(Text_Proc) = Text_OutDir(Text_Proc) & "\"
           End If
        End If
        
'------------------- ｼｽﾃﾑ共通設定
        .INI_FileID = gIniExe & gIniName
        
        '-------- ﾃｷｽﾄ連携ﾃｷｽﾄﾌｧｲﾙ名
        .INI_Key = Text_PrefixNm(Text_Proc) & Text_Orderer(Text_Proc) & "_ﾌｧｲﾙ名"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           If Text_Proc = 0 Then
              .INI_String = "MFDFILE.FIX"
           Else
              If Text_Proc = 1 Then
                 .INI_String = "HACHU97200*.*"
              Else
                 .INI_String = "MFDFILE.FIX"
              End If
           End If
           .PutIni
        End If
        Text_TextName(Text_Proc) = Trim(.INI_String)
        '-------- 保持期間
        .INI_Key = Text_PrefixNm(Text_Proc) & Text_Orderer(Text_Proc) & "_保持期間"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "10"
           .PutIni
        End If
        Text_DelDay(Text_Proc) = CisFun.Val2(.INI_String)
        If Text_DelDay(Text_Proc) = 0 Then
           Text_DelDay(Text_Proc) = 10
        End If
        '-------- 変換リスト
        .INI_Key = Text_PrefixNm(Text_Proc) & Text_Orderer(Text_Proc) & "_変換リスト"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "No"
           .PutIni
        End If
        If UCase(Trim(.INI_String)) = "Y" Or UCase(Trim(.INI_String)) = "YES" Then
           Text_CvtList(Text_Proc) = True
        Else
           Text_CvtList(Text_Proc) = False
        End If
        '-------- 変換リスト(エラー内容)
        .INI_Key = Text_PrefixNm(Text_Proc) & Text_Orderer(Text_Proc) & "_変換リスト_エラー出力"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "Yes"
           .PutIni
        End If
        If UCase(Trim(.INI_String)) = "Y" Or UCase(Trim(.INI_String)) = "YES" Then
           Text_CvtErr(Text_Proc) = 0
        Else
           Text_CvtErr(Text_Proc) = 1
        End If
        '-------- 変換リスト(重複)
        .INI_Key = Text_PrefixNm(Text_Proc) & Text_Orderer(Text_Proc) & "_変換リスト_重複"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "Yes"
           .PutIni
        End If
        If UCase(Trim(.INI_String)) = "Y" Or UCase(Trim(.INI_String)) = "YES" Then
           Text_CvtDbl(Text_Proc) = 0
        Else
           Text_CvtDbl(Text_Proc) = 1
        End If
        '-------- 変換リスト(仮登録内容)
        .INI_Key = Text_PrefixNm(Text_Proc) & Text_Orderer(Text_Proc) & "_変換リスト_仮登録出力"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "Yes"
           .PutIni
        End If
        If UCase(Trim(.INI_String)) = "Y" Or UCase(Trim(.INI_String)) = "YES" Then
           Text_CvtKari(Text_Proc) = 0
        Else
           Text_CvtKari(Text_Proc) = 1
        End If
        '-------- 変換リスト(その他)
        .INI_Key = Text_PrefixNm(Text_Proc) & Text_Orderer(Text_Proc) & "_変換リスト_その他出力"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "Yes"
           .PutIni
        End If
        If UCase(Trim(.INI_String)) = "Y" Or UCase(Trim(.INI_String)) = "YES" Then
           Text_CvtEtc(Text_Proc) = 0
        Else
           Text_CvtEtc(Text_Proc) = 1
        End If
    
    End With
    JTextIniGet = True
End Function
'+-----------------------------------------+
'+  ﾃｷｽﾄ連携環境情報確認
'+-----------------------------------------+
Public Function JTextIniCheck() As Boolean
    JTextIniCheck = False
    Text_ErrMsg = ""
    
    On Error GoTo JTextIniCheck_Er:
    
    '-------- ﾃｷｽﾄ連携ﾃｷｽﾄ格納先
    If Trim(Text_InDir(Text_Proc)) = "" Then
       Text_ErrMsg = "ﾃｷｽﾄ連携ﾃｷｽﾄ格納先未設定"
       GoTo JTextIniCheck_Er
    End If
    If Text_ProcKB <> 0 Then       ' 手動のみチェック
        If Dir(Text_InDir(Text_Proc), vbDirectory) = "" Then
           Text_ErrMsg = "ﾃｷｽﾄ連携ﾃｷｽﾄ格納先(接続エラー）" & vbCrLf & _
                          Space(40) & " 接続先：" & Text_InDir(Text_Proc)
           GoTo JTextIniCheck_Er
        End If
    End If
    '-------- ﾃｷｽﾄ連携ﾃｷｽﾄﾌｧｲﾙ名
    If Trim(Text_TextName(Text_Proc)) = "" Then
       Text_ErrMsg = "ﾃｷｽﾄ連携ﾃｷｽﾄ名未設定"
       GoTo JTextIniCheck_Er
    End If
    '-------- ﾃｷｽﾄ連携ﾃｷｽﾄ退避先(SIMPR-A)
    If Trim(Text_OutDir(Text_Proc)) = "" Then
       Text_ErrMsg = "ﾃｷｽﾄ連携ﾃｷｽﾄ退避先未設定"
       GoTo JTextIniCheck_Er
    End If
    If Dir(Text_OutDir(Text_Proc), vbDirectory) = "" Then
       With CisFun
            .FE_Path = Text_OutDir(Text_Proc)
            .DirCheck NoMsg_Crt
       End With
       Call CisFun.LogPut("'")
       gStr = Space(10) & "ﾃｷｽﾄ連携ﾃｷｽﾄ退避先　自動作成 作成先：" & Text_OutDir(Text_Proc)
       Call CisFun.LogPut(gStr)
       Call CisFun.LogPut("'")
    End If

    JTextIniCheck = True
    Exit Function
JTextIniCheck_Er:
    If ERR.Number = 52 Then
       With CisFun
            .MB_Lines = 6
            .MB_MSG(0) = "    【" & Text_PrefixNm(Text_Proc) & "】ﾃｷｽﾄ連携端末接続エラー"
            .MB_MSG(2) = "    Error No    = " & ERR.Number
            .MB_MSG(3) = "    Error Msg   = " & ERR.Description
            .MB_MSG(5) = "      接続先 = " & Text_InDir(Text_Proc)
            .MB_Title = "": .MB_Button = Error
            If Text_ProcKB <> 0 Then
               Call .MBOX
            End If
        End With
        gStr = "＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／"
        Call Text_LogPut(gStr, , 1)
        gStr = "'"
        Call Text_LogPut(gStr, , 1)
       gStr = "＿／       【" & Text_PrefixNm(Text_Proc) & "】ﾃｷｽﾄ連携端末接続エラー"
        Call Text_LogPut(gStr)
       gStr = "'"
        Call Text_LogPut(gStr, , 1)
       gStr = "＿／             接続先 = " & Text_InDir(Text_Proc)
        Call Text_LogPut(gStr)
       gStr = "'"
        Call Text_LogPut(gStr, , 1)
       gStr = "＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／"
        Call Text_LogPut(gStr, , 1)
       gStr = "'"
        Call Text_LogPut(gStr, , 1)
    End If

    If Trim(Text_ErrMsg) <> "" Then
       Call Text_LogPut(Text_ErrMsg, "ERR")
    End If

End Function
'+------------------------------------------------------------------------+
'+
'+      ﾃｷｽﾄ連携格納ファイル数を求める
'+
'+------------------------------------------------------------------------+
Public Function Text_GetFile_Count() As Long
    Dim wCount      As Long
    Dim wStrFName   As String
    wCount = 0
    
    On Error GoTo Text_GetFile_Count_er:

    If Text_Rireki = 0 Then
        wStrFName = Dir(Text_InDir(Text_Proc) & Text_TextName(Text_Proc))      ' 最初のファイル名を返します。
    Else
        wStrFName = Dir(TRI.処理退避先 & Text_TextName(Text_Proc))      ' 最初のファイル名を返します。
    End If
    Do While wStrFName <> ""                ' ループを開始します。
        wCount = wCount + 1

        wStrFName = Dir                     ' 次のファイル名を返します。
    Loop


    Text_GetFile_Count = wCount
    Exit Function
Text_GetFile_Count_er:
    Text_GetFile_Count = 0
End Function
'*-------------------------------------------------------------------------*
'*-------------------------------------------------------------------------*
'*-------------------------------------------------------------------------*
'*---------------                                           ---------------*
'*---------------     ﾃｷｽﾄ連携退避先　取得&作成        　      ---------------*
'*---------------                                           ---------------*
'*-------------------------------------------------------------------------*
'*-------------------------------------------------------------------------*
Public Function JTextBackGet() As Boolean
    Dim wJTextErr   As Boolean
    
    JTextBackGet = False
    
    On Error GoTo JTextBackGet_er:
    If Text_Rireki = 0 Then
    '   #-----------------------------#
    '   #     ﾃｷｽﾄ連携ﾃﾞｰﾀ連携元確認     #
    '   #-----------------------------#
        TRI.ERR = "10"      'ﾃｷｽﾄ連携端末接続ｴﾗｰ"
        If Not Middle_TRIUpDate Then Exit Function
        wJTextErr = False
        If Dir(Text_InDir(Text_Proc), vbDirectory) = "" Then
           wJTextErr = True
           GoTo JTextBackGet_er:
        End If
    
        TRI.ERR = "11"      'ﾃｷｽﾄ連携退避先取得ｴﾗｰ"
        If Not Middle_TRIUpDate Then Exit Function
        ' ログ出力
        gStr = "【 ﾃｷｽﾄ連携" & Text_PrefixNm(Text_Proc) & "データ　退避先取得&作成 処理開始 】"
        Call Text_LogPut(gStr, "G", 1)
    ' 退避先パスを求める
        TRI.退避先 = Text_OutDir(Text_Proc) & Format(Now(), "YYYYMMDD")
        With CisFun
            .FE_FullPath = Text_OutDir(Text_Proc) & Format(Now(), "YYYYMMDD")
            .GetUNCPath
            TRI.退避先UNC = .FE_UNCPath
        End With
        If InStr(1, TRI.退避先UNC, CisFun.WSID) <> 0 Then
           TRI.処理退避先 = TRI.退避先
        Else
           TRI.処理退避先 = TRI.退避先UNC
        End If
        If Right(TRI.退避先, 1) <> "\" Then TRI.退避先 = TRI.退避先 & "\"
        If Right(TRI.退避先UNC, 1) <> "\" Then TRI.退避先UNC = TRI.退避先UNC & "\"
        
        With CisFun
             .FE_Path = TRI.処理退避先
             If Not .DirCheck(NoMsg_Crt) Then
                ' フォルダなし
                TRI.ERR = "12"      '退避先ﾌｫﾙﾀﾞ作成ｴﾗｰ(日付)"
                If Not Middle_TRIUpDate Then Exit Function
           
                gStr = "退避先パス = " & TRI.処理退避先
                Call Text_LogPut(gStr)
                With CisFun
                     .MB_Lines = 5
                     .MB_MSG(0) = " 【 ﾃｷｽﾄ連携提供ファイル" & Text_PrefixNm(Text_Proc) & " 】の退避先取得&作成　エラー"
                     .MB_MSG(2) = "    退避先ﾌｫﾙﾀﾞ作成ｴﾗｰ"
                     .MB_MSG(4) = "    退避先パス = " & TRI.処理退避先
                End With
                GoTo JTextBackGet_e2
             End If
        End With
    ' 退避先最大回数を求める
        If Right(TRI.処理退避先, 1) <> "\" Then TRI.処理退避先 = TRI.処理退避先 & "\"
        Text_DirMaxCnt = 0
        gStr = Dir(TRI.処理退避先, vbDirectory)
        Do While gStr <> ""
            If gStr <> "." And gStr <> ".." Then
                If Text_DirMaxCnt < CisFun.Val2(gStr) Then
                   Text_DirMaxCnt = CisFun.Val2(gStr)
                End If
            End If
            gStr = Dir
        Loop
        Text_DirMaxCnt = Text_DirMaxCnt + 1
        TRI.処理退避先 = TRI.処理退避先 & Format(Text_DirMaxCnt, "000") & "\"
        TRI.退避先 = TRI.退避先 & Format(Text_DirMaxCnt, "000") & "\"
        TRI.退避先UNC = TRI.退避先UNC & Format(Text_DirMaxCnt, "000") & "\"
        ' ログ出力
        TRI.ERR = "13"      '退避先ﾌｫﾙﾀﾞ作成ｴﾗｰ(回数)"
        If Not Middle_TRIUpDate Then Exit Function
        
        gStr = "退避先パス = " & TRI.処理退避先
        Call Text_LogPut(gStr)
        
        MkDir (TRI.処理退避先)
        
        ' ログ出力
        gStr = "【 ﾃｷｽﾄ連携" & Text_PrefixNm(Text_Proc) & "データ　退避先取得&作成 処理終了 】"
        Call Text_LogPut(gStr, "G", 1)
    End If
    
    ' ログ情報
    With Text_Fun
        .Log_PutFlg = True
        .Log_FilePath = TRI.処理退避先
        .Log_FileInitial = Text_Prefix(Text_Proc) & "_" & Format(Text_DirMaxCnt, "000") & "_"
        .Log_DataHDTime = あり
    End With
    
    With CisFun
        .Log_DataHD = Format(Text_DirMaxCnt, "000")
    End With
    ' 履歴情報内容更新
    TRI.退避連番 = Text_DirMaxCnt
    
    Text_FunFlg = True
    
    JTextBackGet = True
    Exit Function
JTextBackGet_er:
    If ERR.Number = 52 Or wJTextErr Then
       With CisFun
            .MB_Lines = 6
            .MB_MSG(0) = "    【" & Text_PrefixNm(Text_Proc) & "】ﾃｷｽﾄ連携端末接続エラー"
            .MB_MSG(2) = "    Error No    = " & ERR.Number
            .MB_MSG(3) = "    Error Msg   = " & ERR.Description
            .MB_MSG(5) = "      接続先 = " & Text_InDir(Text_Proc)
            .MB_Title = "": .MB_Button = Error
            If Text_ProcKB <> 0 Then
               Call .MBOX
            End If
        End With
        gStr = "＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／"
        Call Text_LogPut(gStr, , 1)
        gStr = "'"
        Call Text_LogPut(gStr, , 1)
        gStr = "＿／       【" & Text_PrefixNm(Text_Proc) & "】ﾃｷｽﾄ連携端末接続エラー"
        Call Text_LogPut(gStr)
        gStr = "'"
        Call Text_LogPut(gStr, , 1)
        gStr = "＿／             接続先 = " & Text_InDir(Text_Proc)
        Call Text_LogPut(gStr)
        gStr = "'"
        Call Text_LogPut(gStr, , 1)
        gStr = "＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／"
        Call Text_LogPut(gStr, , 1)
        gStr = "'"
        Call Text_LogPut(gStr, , 1)
        Exit Function
    End If
    
    ReDim DispMsg(5)
    DispMsg(0) = " 【 ﾃｷｽﾄ連携提供ファイル" & Text_PrefixNm(Text_Proc) & " 】の退避先取得&作成　エラー"
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description
    DispMsg(4) = "    退避先パス = " & TRI.処理退避先
JTextBackGet_e2:
    Call Text_LogPut(DispMsg(0))
    Call Text_LogPut(DispMsg(1))
    Call Text_LogPut(DispMsg(2))
    Call Text_LogPut(DispMsg(3))
    Call Text_LogPut(DispMsg(4))
    gStr = "【 ﾃｷｽﾄ連携" & Text_PrefixNm(Text_Proc) & "データ　退避先取得&作成 異常終了 】"
    Call Text_LogPut(gStr, "G", 1)
    
    If Text_ProcKB <> 0 Then
       With CisFun
            .MB_Lines = 5
            .MB_MSG(0) = DispMsg(0)
            .MB_MSG(1) = DispMsg(1)
            .MB_MSG(2) = DispMsg(2)
            .MB_MSG(3) = DispMsg(3)
            .MB_MSG(4) = DispMsg(4)
            .MB_Title = ""
            .MB_Button = Error
            Call .MBOX
       End With
    End If

End Function


'===========================================================================
'===========================================================================
'=======                                                             =======
'=======        【 読 込 テ キ ス ト 関 係 処 理 】                  =======
'=======                                                             =======
'=======        ﾃｷｽﾄ連携側データのコピー＆読込用テキストパスセット      =======
'=======        読込テキストのバックアップ                           =======
'=======                                                             =======
'===========================================================================
'===========================================================================
'*----------------------------------------------------*
'*  ﾃｷｽﾄ連携側データのコピー＆読込用テキストパスセット   *
'*----------------------------------------------------*
Public Function Text_TextMove(FRM As Form) As Boolean
    Text_TextMove = False
    
    gStr = "【 ﾃｷｽﾄ連携" & Text_PrefixNm(Text_Proc) & "データ　移動 開始 】"
    Call Text_LogPut(gStr, "G", 1)

    TRI.ERR = "20"      'ﾃｷｽﾄ連携TEXT移動ｴﾗｰ"
    If Not Middle_TRIUpDate Then Exit Function
'   *------------------------------------*
'   * 自社バックアップフォルダへのコピー *
'   *------------------------------------*
    ReDim destFName(1)
    destFName(0) = Text_InDir(Text_Proc) & Text_TextName(Text_Proc)
    ' ログ出力
    gStr = "ﾃｷｽﾄ連携→TNS　移動 (" & destFName(0) & "→" & TRI.処理退避先 & ")"
    Call Text_LogPut(gStr)
    TRI.ERR = "21"      'ﾃｷｽﾄ連携ﾌｧｲﾙ移動ｴﾗｰ"
    If Not Middle_TRIUpDate Then Exit Function
    
    If Text_ProcKB = 0 And Text_ProcDisp = 0 Then
       gInt = 4
    Else
       gInt = 2
    End If
    If Not gFileDestroy(FRM, gInt, TRI.処理退避先) Then
       gStr = "ファイル又はディスクが読み取れません。"
       Call Text_LogPut(gStr)
       RmDir TRI.退避先
       TRI.退避先 = ""
       GoTo Move_Err
    End If
    ' ログ出力
    gStr = "ﾃｷｽﾄ連携→TNS　移動終了"
    Call Text_LogPut(gStr)
    gStr = "【 ﾃｷｽﾄ連携" & Text_PrefixNm(Text_Proc) & "データ　移動 正常終了 】"
    Call Text_LogPut(gStr, "G", 1)
'
    On Error GoTo 0
    Text_TextMove = True
    
    Exit Function

Move_Err:
    ReDim DispMsg(5)
    DispMsg(0) = " 【 ﾃｷｽﾄ連携提供ファイル" & Text_PrefixNm(Text_Proc) & " 】の移動　エラー"
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description
    DispMsg(4) = "    提供ファイル = " & Text_TextName(Text_Proc)
    
    Call Text_LogPut(DispMsg(0))
    Call Text_LogPut(DispMsg(1))
    Call Text_LogPut(DispMsg(2))
    Call Text_LogPut(DispMsg(3))
    Call Text_LogPut(DispMsg(4))
    
    gStr = "【 ﾃｷｽﾄ連携" & Text_PrefixNm(Text_Proc) & "データ　移動 異常終了 】"
    Call Text_LogPut(gStr, "G", 1)
    
    If Text_ProcKB <> 0 Then
       With CisFun
            .MB_Lines = 5
            .MB_MSG(0) = DispMsg(0)
            .MB_MSG(1) = DispMsg(1)
            .MB_MSG(2) = DispMsg(2)
            .MB_MSG(3) = DispMsg(3)
            .MB_MSG(4) = DispMsg(4)
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
       End With
    End If
    
End Function

'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
'*---------------                                                                             ---------------*
'*---------------     テ　キ　ス　ト　デ　ー　タ　テ　ー　ブ　ル　コ　ン　バ　ー　ト          ---------------*
'*---------------                                                                             ---------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
Public Function Text_BULK_Insert(Optional FPath As String = "") As Boolean
    Dim wPathName       As String
    Dim wFileName       As String
    Dim wCnt            As Long
    Dim wFName()        As String
    Dim FIdx            As Long
    Dim FIdx2           As Long
    Dim BULKPath        As String
    
    Erase wFName
    ReDim wFName(1)
    
    Text_BULK_Insert = False
    
    On Error GoTo Text_BULK_Insert_Err:
    
    ' ログ出力
    gStr = "【 ﾃｷｽﾄ連携" & Text_PrefixNm(Text_Proc) & "テキストデータ　コンバート 開始 】"
    Call Text_LogPut(gStr, "G", 1)
    
    TRI.ERR = "41"      'ﾃｷｽﾄ連携ﾃｷｽﾄｲﾝﾎﾟｰﾄｴﾗｰ
    If Not Middle_TRIUpDate Then Exit Function
    
    
    If FPath = "" Then
       FPath = TRI.処理退避先
    End If
    
    BULKPath = TRI.退避先UNC
        
    If UCase(CisFun.WSID) = UCase(Trim(CisDB.Server)) Then
        If InStr(1, TRI.退避先UNC, CisFun.WSID) <> 0 Then
           BULKPath = TRI.退避先
        Else
           BULKPath = TRI.退避先UNC
        End If
    End If
'    With CisFun
'        .FE_FullPath = FPath
'        .GetUNCPath
'        wPathName = .FE_UNCPath
'    End With
'    If Right(wPathName, 1) <> "\" Then
'       wPathName = wPathName & "\"
'    End If
    If Right(FPath, 1) <> "\" Then
       FPath = FPath & "\"
    End If
    If Right(BULKPath, 1) <> "\" Then
       BULKPath = BULKPath & "\"
    End If
    
    ' ファイル名を求める
    wFileName = Dir(FPath & Text_TextName(Text_Proc))
    If wFileName <> "" Then
       gStr = "コンバートパス名　　 ： " & FPath
       Call Text_LogPut(gStr, "G")
    End If
    FIdx = 0
    ' ファイル名を配列に格納
    Do While wFileName <> ""
        FIdx = FIdx + 1
        ReDim Preserve wFName(FIdx)
        wFName(FIdx - 1) = wFileName
    
        wFileName = Dir
    Loop
    
    ' 昇順
    wFileName = ""
    For FIdx = 0 To UBound(wFName) - 1
        If Trim(wFName(FIdx)) = "" Then Exit For
        For FIdx2 = FIdx + 1 To UBound(wFName) - 1
            If Trim(wFName(FIdx2)) = "" Then Exit For
        
            If Trim(wFName(FIdx)) > Trim(wFName(FIdx2)) Then
               wFileName = wFName(FIdx2)
               wFName(FIdx2) = wFName(FIdx)
               wFName(FIdx) = wFileName
            End If
        Next FIdx2
    Next FIdx
    
    wCnt = 0

    For FIdx = 0 To UBound(wFName) - 1
        wFileName = wFName(FIdx)
        wCnt = wCnt + 1
        gStr = "ファイル名 ＝ " & wFileName
        Call Text_LogPut(gStr, , 10)
        '----------------------------------------- テーブルインサート
'        If Not Text_Bulk_Proc(BULKPath & wFileName, wFileName) Then
        If Not Text_Text_Proc(BULKPath & wFileName, wFileName) Then
           gStr = "【 ﾃｷｽﾄ連携" & Text_PrefixNm(Text_Proc) & "テキストデータ　コンバート 異常終了 】"
           Call Text_LogPut(gStr, "G", 1)
           Exit Function
        End If
        
        If Not Text_Frm Is Nothing Then
           Text_Frm.MeterUpdate wCnt
        End If
    Next FIdx
    
    If wCnt = 0 Then
        TRI.ERR = "46"      'ｲﾝﾎﾟｰﾄ対象ﾌｧｲﾙなし
        If Not Middle_TRIUpDate Then Exit Function
        ReDim DispMsg(4)
        DispMsg(2) = "   コンバート対象ファイルなし"
        DispMsg(3) = "   "
        GoTo Text_BULK_Insert_ErrDisp:
    End If

    ' ログ出力
    gStr = "【 ﾃｷｽﾄ連携" & Text_PrefixNm(Text_Proc) & "テキストデータ　コンバート 終了 】"
    Call Text_LogPut(gStr, "G", 1)

    Text_BULK_Insert = True
    Exit Function
Text_BULK_Insert_Err:
    ReDim DispMsg(4)
    DispMsg(0) = " 【 テキストデータコンバート 】エラー"
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description
    
Text_BULK_Insert_ErrDisp:
    
    Call Text_LogPut(DispMsg(0))
    Call Text_LogPut(DispMsg(1))
    Call Text_LogPut(DispMsg(2))
    Call Text_LogPut(DispMsg(3))
    
    gStr = "【 ﾃｷｽﾄ連携" & Text_PrefixNm(Text_Proc) & "テキストデータ　コンバート 異常終了 】"
    Call Text_LogPut(gStr, "G", 1)
    
    If Text_ProcKB <> 0 Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(0) = DispMsg(0)
            .MB_MSG(1) = DispMsg(1)
            .MB_MSG(2) = DispMsg(2)
            .MB_MSG(3) = DispMsg(3)
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
       End With
    End If
End Function

Public Function Text_Bulk_Proc(wFullFile As String, wFileName As String) As Boolean
    Text_Bulk_Proc = False
    Static sw_Table         As String
    Static sw_Stoado        As String

    On Error GoTo Text_Bulk_Proc_Err
    
    If Text_Proc = 0 Then
       sw_Table = "テキスト受注データ"
       sw_Stoado = "テキスト受注TEXT変換"
    Else
       sw_Table = "テキスト出荷データ"
       sw_Stoado = "テキスト出荷TEXT変換"
    End If

'[ データ削除 ]
    TRI.ERR = "43"      'ﾃｷｽﾄ連携ｲﾝﾎﾟｰﾄ削除ｴﾗｰ
    If Not Middle_TRIUpDate Then Exit Function
    With CisDB
        .SQL = "TRUNCATE TABLE " & sw_Table
        If Not .DBExec Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call Text_LogPut(gStr)
           GoTo Text_Bulk_Proc_Err
        End If
    End With
'[ BULK INSERT ]
    TRI.ERR = "44"      'ﾃｷｽﾄ連携ｲﾝﾎﾟｰﾄ処理ｴﾗｰ
    If Not Middle_TRIUpDate Then Exit Function
    
    gSL_Select = "BULK INSERT " & sw_Table
    gSL_Select = gSL_Select & " From '" & wFullFile & "'"
    gSL_Select = gSL_Select & " With ("
    gSL_Select = gSL_Select & " ROWTERMINATOR = '\n'"
    gSL_Select = gSL_Select & " )"
    With CisDB
        .SQL = gSL_Select
        If Not .DBExec Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call Text_LogPut(gStr)
           GoTo Text_Bulk_Proc_Err
        End If
        gStr = "　　　　　処理件数 ＝ " & Format(.RecordCount, "#,##0")
        Call Text_LogPut(gStr, , 40)
    End With
    
'[TEXT変換]
    TRI.ERR = "45"      'ﾃｷｽﾄ連携TEXT変換ｴﾗｰ
    If Not Middle_TRIUpDate Then Exit Function
    With CisDB
        .SQL = sw_Stoado
        .StoadoCount = 5
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
        .ParaNo = 3:    .ParaIO = Input用
        .ParaNo = 4:    .ParaIO = Input用
        .ParaNo = 5:    .ParaIO = OutPut用
    
        .ParaNo = 1: .ParaValue = Format(Now(), "YYYYMMDD")
        .ParaNo = 2: .ParaValue = TRI.処理回数
        .ParaNo = 3: .ParaValue = TRI.履歴管理番号
        
        gCnt1 = InStr(1, wFileName, ".")
        If gCnt1 > 0 Then
           gStr = Mid(wFileName, 1, gCnt1 - 1)
        Else
           gStr = wFileName
        End If
        .ParaNo = 4: .ParaValue = Mid(gStr, 1, 50)
    
        If Not .DBStored Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call Text_LogPut(gStr)
           GoTo Text_Bulk_Proc_Err
        End If
        .ParaNo = 5
        If .ParaValue <> 0 Then
            ReDim DispMsg(4)
            DispMsg(0) = "   " & sw_Stoado & "エラー"
            DispMsg(2) = "   Err Code = " & .ParaValue
            DispMsg(3) = "    File = " & wFullFile
            GoTo Text_Bulk_Proc_ErrDisp
        End If
    End With
    
'------------ ﾃｷｽﾄ連携変換履歴テーブル 再入力
    gLong = TRI.履歴管理番号
    gSL_Select = "Select * From ﾃｷｽﾄ変換履歴テーブル"
    gSL_Select = gSL_Select & " Where 履歴管理番号 = " & TRI.履歴管理番号 & ""
    If Not TRIRead(gSL_Select, 1) Then                                              ' EHR-->TRI
        ReDim DispMsg(4)
        DispMsg(0) = "      " & sw_Stoado & "エラー"
        DispMsg(2) = "    ﾃｷｽﾄ連携変換履歴テーブル異常            "
        DispMsg(3) = "    履歴管理番号 = " & gLong
        GoTo Text_Bulk_Proc_ErrDisp
    End If

    Text_Bulk_Proc = True
    Exit Function
Text_Bulk_Proc_Err:
    ReDim DispMsg(4)
    DispMsg(0) = "   Insert Err"
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description

Text_Bulk_Proc_ErrDisp:

    Call Text_LogPut(DispMsg(0))
    Call Text_LogPut(DispMsg(1))
    Call Text_LogPut(DispMsg(2))
    Call Text_LogPut(DispMsg(3))

    If Text_ProcKB <> 0 Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(0) = DispMsg(0)
            .MB_MSG(1) = DispMsg(1)
            .MB_MSG(2) = DispMsg(2)
            .MB_MSG(3) = DispMsg(3)
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
       End With
    End If

End Function
Public Function Text_Text_Proc(wFullFile As String, wFileName As String) As Boolean
    Text_Text_Proc = False
    Static sw_Table         As String
    Static sw_Stoado        As String
    Static sw_TE_Text       As String

    On Error GoTo Text_Text_Proc_Err
    
'    If Text_Proc = 0 Then
'       sw_Table = "テキスト受注データ"
'       sw_Stoado = "テキスト受注TEXT変換"
'    Else
'       sw_Table = "テキスト出荷データ"
'       sw_Stoado = "テキスト出荷TEXT変換"
'    End If
    Select Case Text_Proc
           Case 0: sw_Table = "テキスト受注データ"
                   sw_Stoado = "テキスト受注TEXT変換"
           Case 1: sw_Table = "テキスト受注データ"
                   sw_Stoado = "テキスト受注TEXT変換_テクノエイト"
           Case 3: sw_Table = "テキスト受注データ"
                   sw_Stoado = "テキスト受注TEXT変換_ヨロズ愛知"
           Case Else: sw_Table = "テキスト出荷データ"
                      sw_Stoado = "テキスト出荷TEXT変換"
    End Select

'[ データ削除 ]
    TRI.ERR = "43"      'ﾃｷｽﾄ連携ｲﾝﾎﾟｰﾄ削除ｴﾗｰ
    If Not Middle_TRIUpDate Then Exit Function
    With CisDB
        .SQL = "TRUNCATE TABLE " & sw_Table
        If Not .DBExec Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call Text_LogPut(gStr)
           GoTo Text_Text_Proc_Err
        End If
    End With
'[ BULK INSERT ]
    TRI.ERR = "44"      'ﾃｷｽﾄ連携ｲﾝﾎﾟｰﾄ処理ｴﾗｰ
    If Not Middle_TRIUpDate Then Exit Function
    
    Text_FNo = FreeFile
    gInt = 0
    Open wFullFile For Input As #Text_FNo
    Do While Not EOF(Text_FNo)
        Line Input #Text_FNo, Text_Rec
        gInt = gInt + 1
        
        If Text_Proc = 0 Or Text_Proc = 3 Then
           gSL_Select = "INSERT INTO テキスト受注データ (REC1)"
           gSL_Select = gSL_Select & " VALUES ('" & CisFun.Mid2(Text_Rec, 1) & "')"
        End If
        If Text_Proc = 1 And gInt <> 1 Then     '１行目は見出しの為処理行わない
           sw_TE_Text = ""
           Call GetString(Text_Rec, ",")        'テクノエイトはカンマ編集のCSV
           '仕入先
           sw_TE_Text = sw_TE_Text & RTrim(CisFun.Mid2(gString(0), 1, 7)) & Space(7 - CisFun.Len2(RTrim(CisFun.Mid2(gString(0), 1, 7))))
           '受入
           sw_TE_Text = sw_TE_Text & RTrim(CisFun.Mid2(gString(1), 1, 2)) & Space(2 - CisFun.Len2(RTrim(CisFun.Mid2(gString(1), 1, 2))))
           '納入日
           sw_TE_Text = sw_TE_Text & Space(4 - CisFun.Len2(CisFun.Mid2(gString(5), 1, 4))) & CisFun.Mid2(gString(5), 1, 4) & Space(2 - CisFun.Len2(CisFun.Mid2(gString(5), 6, 2))) & CisFun.Mid2(gString(5), 6, 2) & Space(2 - CisFun.Len2(CisFun.Mid2(gString(5), 9, 2))) & CisFun.Mid2(gString(5), 9, 2)
           '納入便
           sw_TE_Text = sw_TE_Text & Space(2 - CisFun.Len2(CisFun.Mid2(gString(6), 1, 2))) & CisFun.Mid2(gString(6), 1, 2)
           '納入番号
           sw_TE_Text = sw_TE_Text & RTrim(CisFun.Mid2(gString(7), 1, 5)) & Space(5 - CisFun.Len2(RTrim(CisFun.Mid2(gString(7), 1, 5))))
           '納品書頁
           sw_TE_Text = sw_TE_Text & Space(2 - CisFun.Len2(CisFun.Mid2(gString(8), 1, 2))) & CisFun.Mid2(gString(8), 1, 2)
           '品番
           sw_TE_Text = sw_TE_Text & RTrim(CisFun.Mid2(gString(9), 1, 30)) & Space(30 - CisFun.Len2(RTrim(CisFun.Mid2(gString(9), 1, 30))))
           '背番号
           sw_TE_Text = sw_TE_Text & RTrim(CisFun.Mid2(gString(10), 1, 5)) & Space(5 - CisFun.Len2(RTrim(CisFun.Mid2(gString(10), 1, 5))))
           '収容数
           sw_TE_Text = sw_TE_Text & Space(6 - CisFun.Len2(CisFun.Mid2(gString(11), 1, 6))) & CisFun.Mid2(gString(11), 1, 6)
           '箱数
           sw_TE_Text = sw_TE_Text & Space(3 - CisFun.Len2(CisFun.Mid2(gString(12), 1, 3))) & CisFun.Mid2(gString(12), 1, 3)
           '端数
           sw_TE_Text = sw_TE_Text & Space(6 - CisFun.Len2(CisFun.Mid2(gString(13), 1, 6))) & CisFun.Mid2(gString(13), 1, 6)
           '発注数
           sw_TE_Text = sw_TE_Text & Space(6 - CisFun.Len2(CisFun.Mid2(gString(14), 1, 6))) & CisFun.Mid2(gString(14), 1, 6)
           
           
           
           gSL_Select = "INSERT INTO テキスト受注データ (REC1)"
           gSL_Select = gSL_Select & " VALUES ('" & CisFun.Mid2(sw_TE_Text, 1) & "')"
        End If
        If Text_Proc <> 0 And Text_Proc <> 1 And Text_Proc <> 3 Then
           gSL_Select = "INSERT INTO テキスト出荷データ (REC1)"
           gSL_Select = gSL_Select & " VALUES ('" & CisFun.Mid2(Text_Rec, 1) & "')"
        End If
        With CisDB
            If (Text_Proc = 0) Or (Text_Proc = 1 And gInt <> 1) Or (Text_Proc = 2) Or (Text_Proc = 3) Then
                .SQL = gSL_Select
                If Not .DBExec Then
                   gStr = .ErrCode & ":" & .ErrMsg
                   Call Text_LogPut(gStr)
                   GoTo Text_Text_Proc_Err
                End If
            End If
            
            If Text_Proc = 0 Or Text_Proc = 3 Then
                gStr = " " & Format(Format(gInt, "#,###"), "@@@@@") & "："
                gStr = gStr & CisFun.Mid2(Text_Rec, 24, 6) & " " _
                            & CisFun.Mid2(Text_Rec, 31, 3) & " "
                gStr = gStr & "20" & CisFun.Mid2(Text_Rec, 14, 6) & "01" _
                            & " " & CisFun.Mid2(Text_Rec, 6, 7) & " "
                gStr = gStr & CisFun.Mid2(Text_Rec, 35, 12)
            End If
            If Text_Proc = 1 And gInt <> 1 Then
                gStr = " " & Format(Format(gInt, "#,###"), "@@@@@") & "："
                gStr = gStr & gString(0) & " " _
                            & gString(1) & " "
                gStr = gStr & gString(5) & "-" & gString(6) _
                            & " " & gString(7) & " "
                gStr = gStr & gString(9)
            End If
            If Text_Proc <> 0 And Text_Proc <> 1 And Text_Proc <> 3 Then
                gStr = " " & Format(Format(gInt, "#,###"), "@@@@@") & "："
                gStr = gStr & CisFun.Mid2(Text_Rec, 24, 6) & " " _
                            & CisFun.Mid2(Text_Rec, 31, 3) & " "
                gStr = gStr & "20" & CisFun.Mid2(Text_Rec, 14, 6) & "01" _
                            & " " & CisFun.Mid2(Text_Rec, 6, 7) & " "
                gStr = gStr & CisFun.Mid2(Text_Rec, 35, 12)
            End If
            Call Text_LogPut(gStr, , 5)
        End With
        
    Loop

    Close #Text_FNo
        
'''    gSL_Select = "BULK INSERT " & sw_Table
'''    gSL_Select = gSL_Select & " From '" & wFullFile & "'"
'''    gSL_Select = gSL_Select & " With ("
'''    gSL_Select = gSL_Select & " ROWTERMINATOR = '\n'"
'''    gSL_Select = gSL_Select & " )"
'''    With CisDB
'''        .SQL = gSL_Select
'''        If Not .DBExec Then
'''           gStr = .ErrCode & ":" & .ErrMsg
'''           Call Text_LogPut(gStr)
'''           GoTo Text_Text_Proc_Err
'''        End If
'''        gStr = "　　　　　処理件数 ＝ " & Format(.RecordCount, "#,##0")
'''        Call Text_LogPut(gStr, , 40)
'''    End With
    
'[TEXT変換]
    TRI.ERR = "45"      'ﾃｷｽﾄ連携TEXT変換ｴﾗｰ
    If Not Middle_TRIUpDate Then Exit Function
    With CisDB
        .SQL = sw_Stoado
        .StoadoCount = 5
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
        .ParaNo = 3:    .ParaIO = Input用
        .ParaNo = 4:    .ParaIO = Input用
        .ParaNo = 5:    .ParaIO = OutPut用
    
        .ParaNo = 1: .ParaValue = Format(Now(), "YYYYMMDD")
        .ParaNo = 2: .ParaValue = TRI.処理回数
        .ParaNo = 3: .ParaValue = TRI.履歴管理番号
        
        gCnt1 = InStr(1, wFileName, ".")
        If gCnt1 > 0 Then
           gStr = Mid(wFileName, 1, gCnt1 - 1)
        Else
           gStr = wFileName
        End If
        .ParaNo = 4: .ParaValue = Mid(gStr, 1, 50)

        If Not .DBStored Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call Text_LogPut(gStr)
           GoTo Text_Text_Proc_Err
        End If
        .ParaNo = 5
        If .ParaValue <> 0 Then
            ReDim DispMsg(4)
            DispMsg(0) = "   " & sw_Stoado & "エラー"
            DispMsg(2) = "   Err Code = " & .ParaValue
            DispMsg(3) = "    File = " & wFullFile
            GoTo Text_Text_Proc_ErrDisp
        End If
    End With
    
'------------ ﾃｷｽﾄ連携変換履歴テーブル 再入力
    gLong = TRI.履歴管理番号
    gSL_Select = "Select * From ﾃｷｽﾄ変換履歴テーブル"
    gSL_Select = gSL_Select & " Where 履歴管理番号 = " & TRI.履歴管理番号 & ""
    If Not TRIRead(gSL_Select, 1) Then
        ReDim DispMsg(4)
        DispMsg(0) = "      " & sw_Stoado & "エラー"
        DispMsg(2) = "    ﾃｷｽﾄ連携変換履歴テーブル異常            "
        DispMsg(3) = "    履歴管理番号 = " & gLong
        GoTo Text_Text_Proc_ErrDisp
    End If

    Text_Text_Proc = True
    Exit Function
Text_Text_Proc_Err:
    ReDim DispMsg(4)
    DispMsg(0) = "   Insert Err"
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description

Text_Text_Proc_ErrDisp:

    Call Text_LogPut(DispMsg(0))
    Call Text_LogPut(DispMsg(1))
    Call Text_LogPut(DispMsg(2))
    Call Text_LogPut(DispMsg(3))

    If Text_ProcKB <> 0 Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(0) = DispMsg(0)
            .MB_MSG(1) = DispMsg(1)
            .MB_MSG(2) = DispMsg(2)
            .MB_MSG(3) = DispMsg(3)
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
       End With
    End If

End Function
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
'*---------------                                                                             ---------------*
'*---------------     ﾃｷｽﾄ連携　更　新　処　理                                                   ---------------*
'*---------------                                                                             ---------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
Public Function Text_UpdateProc() As Boolean
    On Error GoTo Text_UpdateProc_Er:
    
    Text_UpdateProc = False
    
    ' ログ出力
    gStr = "【 ﾃｷｽﾄ連携" & Text_PrefixNm(Text_Proc) & "更新処理 開始 】"
    Call Text_LogPut(gStr, "G", 1)
    
    TRI.ERR = "51"      'ﾃｷｽﾄ連携更新ｴﾗｰ
    If Not Middle_TRIUpDate Then Exit Function
    With CisDB
        .SQL = "テキスト受注出荷更新"
'        .StoadoCount = 5
        .StoadoCount = 6                                '2010/05/31
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
        .ParaNo = 3:    .ParaIO = Input用
        .ParaNo = 4:    .ParaIO = Input用
        .ParaNo = 5:    .ParaIO = OutPut用
        .ParaNo = 6:    .ParaIO = Input用               '2010/05/31
    
''        .ParaNo = 1: .ParaValue = Text_Proc
        If Text_Proc = 0 Or Text_Proc = 1 Or Text_Proc = 3 Then
           .ParaNo = 1: .ParaValue = 0
        Else
           .ParaNo = 1: .ParaValue = 1
        End If
        
        .ParaNo = 2: .ParaValue = Format(Now(), "YYYYMMDD")
        .ParaNo = 3: .ParaValue = Text_ProcCNT
        .ParaNo = 4: .ParaValue = TRI.履歴管理番号
        
        .ParaNo = 6: .ParaValue = Text_YMD              '2010/05/31
    
        If Not .DBStored Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call Text_LogPut(gStr)
           GoTo Text_UpdateProc_Er
        End If
        .ParaNo = 5
        If .ParaValue <> 0 Then
            ReDim DispMsg(4)
            DispMsg(0) = "   ﾃｷｽﾄ連携更新エラー"
            DispMsg(2) = "   Err Code = " & .ParaValue
            GoTo Text_UpdateProc_ErDisp
        End If
    End With
    
'------------ ﾃｷｽﾄ連携変換履歴テーブル 再入力
    gLong = TRI.履歴管理番号
    gSL_Select = "Select * From テキスト変換履歴テーブル"
    gSL_Select = gSL_Select & " Where 履歴管理番号 = " & TRI.履歴管理番号 & ""
    If Not TRIRead(gSL_Select, 1) Then
        ReDim DispMsg(4)
        DispMsg(0) = "         ﾃｷｽﾄ連携更新エラー                 "
        DispMsg(2) = "    ﾃｷｽﾄ連携変換履歴テーブル異常            "
        DispMsg(3) = "    履歴管理番号 = " & gLong
        GoTo Text_UpdateProc_ErDisp
    End If
    
    gStr = "対象件数         =" & Format(TRI.入力件数, "#,##0"): Call Text_LogPut(gStr, "", 20)
    gStr = "出力件数         =" & Format(TRI.変換件数, "#,##0"): Call Text_LogPut(gStr, "", 20)
    gStr = "重複件数         =" & Format(TRI.重複, "#,##0"): Call Text_LogPut(gStr, "", 20)
    If Text_Proc = 0 Then
'''       gStr = "出荷指示書空件数 =" & Format(TRI.空, "#,##0"): Call Text_LogPut(gStr, "", 20)
    Else
       gStr = "受注なし件数     =" & Format(TRI.ERR受注, "#,##0"): Call Text_LogPut(gStr, "", 20)
    End If
    gStr = "仮登録件数　　　 :(品番ﾏｽﾀ)=" & Format(TRI.仮品番, "#,##0") & _
                             " (品番取引先ﾏｽﾀ)=" & Format(TRI.仮品取, "#,##0") & _
                             " (取引先ﾏｽﾀ)=" & Format(TRI.仮取引先, "#,##0") & _
                             " (取引先受入ﾏｽﾀ)=" & Format(TRI.仮受入, "#,##0")
    Call Text_LogPut(gStr, "", 20)
    
    ' ログ出力
    gStr = "【 ﾃｷｽﾄ連携" & Text_PrefixNm(Text_Proc) & "更新処理 終了 】"
    Call Text_LogPut(gStr, "G", 1)
    
    Text_UpdateProc = True
    Exit Function
Text_UpdateProc_Er:
    ReDim DispMsg(4)
    DispMsg(0) = "   ﾃｷｽﾄ連携 Update Err"
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description

Text_UpdateProc_ErDisp:
    Call Text_LogPut(DispMsg(0))
    Call Text_LogPut(DispMsg(1))
    Call Text_LogPut(DispMsg(2))
    Call Text_LogPut(DispMsg(3))

    If Text_ProcKB <> 0 Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(0) = DispMsg(0)
            .MB_MSG(1) = DispMsg(1)
            .MB_MSG(2) = DispMsg(2)
            .MB_MSG(3) = DispMsg(3)
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
       End With
    End If

    ' ログ出力
    gStr = "【 ﾃｷｽﾄ連携" & Text_PrefixNm(Text_Proc) & "更新処理 異常終了 】"
    Call Text_LogPut(gStr, "G", 1)
    
End Function

'*-------------------------------------------------------------------------*
'*-------------------------------------------------------------------------*
'*-------------------------------------------------------------------------*
'*---------------                                           ---------------*
'*---------------     ﾊﾞｯｸｱｯﾌﾟデータ削除                    ---------------*
'*---------------                                           ---------------*
'*-------------------------------------------------------------------------*
'*-------------------------------------------------------------------------*
'*-------------------------------------------------------------------------*
Public Function Text_BackUpDel() As Boolean
    Dim wDate()     As String
    Dim wCnt        As Long
    Dim wCnt2       As Long
    Dim wDirName    As String
    Dim wSubDir()   As String
    Dim wDelDay     As String
    Dim wDelTable   As String
    
    Text_BackUpDel = False
    
    If Text_Rireki <> 0 Then
       Text_BackUpDel = True
       Exit Function
    End If
    
    On Error GoTo BackUpDel_Er:
    '+---------------------------------+
    '+  ﾃｷｽﾄ連携ﾊﾞｯｸｱｯﾌﾟ削除
    '+---------------------------------+
    CisFun.Log_DataHD = ""
    
    gStr = "【 ﾃｷｽﾄ連携" & Text_PrefixNm(Text_Proc) & "バックアップデータ削除 開始 】"
    ' ログ出力
    Call Text_LogPut(gStr, "G", 1)
    
    TRI.ERR = "31"      'ﾃｷｽﾄ連携ﾊﾞｯｸｱｯﾌﾟ削除ｴﾗｰ"
    If Not Middle_TRIUpDate Then Exit Function
    
' 日付ﾌｫﾙﾀﾞ名を求める
    wCnt = 0
    wDelDay = Format(Now(), "YYYYMMDD")
    With CisFun
        .Date_RsFormat = fYMD
        .Date_Add wDelDay, D, Text_DelDay(Text_Proc) * -1
        wDelDay = .Date_Result
    End With
    wDirName = Dir(Text_OutDir(Text_Proc), vbDirectory)
    Do While wDirName <> ""
        If wDirName <> "." And wDirName <> ".." And wDirName <> "FPD" And _
           Mid(wDirName, 1, 8) < wDelDay Then
            wCnt = wCnt + 1
            ReDim Preserve wDate(wCnt)
            wDate(wCnt - 1) = wDirName
        End If
        wDirName = Dir
    Loop
    If wCnt = 0 Then
       gStr = "削除対象データなし"
       Call Text_LogPut(gStr)
       GoTo BackUpDel_Ed
    End If
    For wCnt2 = 0 To UBound(wDate) - 1
        ' 日付毎のｻﾌﾞﾌｫﾙﾀﾞを求める
        wCnt = 0
        Erase wSubDir
        
        wDirName = Dir(Text_OutDir(Text_Proc) & wDate(wCnt2) & "\", vbDirectory)
        Do While wDirName <> ""
            If wDirName <> "." And wDirName <> ".." Then
                wCnt = wCnt + 1
                ReDim Preserve wSubDir(wCnt)
                wSubDir(wCnt - 1) = wDirName
            End If
            wDirName = Dir
        Loop
        ' 日付毎のｻﾌﾞﾌｫﾙﾀﾞ内ﾌｧｲﾙ削除
        ' ログ出力
        gStr = "日付毎のｻﾌﾞﾌｫﾙﾀﾞ内ﾌｧｲﾙ削除"
        Call Text_LogPut(gStr, "G")
        For wCnt = 0 To UBound(wSubDir) - 1
            On Error Resume Next
            Kill Text_OutDir(Text_Proc) & wDate(wCnt2) & "\" & wSubDir(wCnt) & "\*.*"
            ' ログ出力
            gStr = "削除内容 = " & Text_OutDir(Text_Proc) & wDate(wCnt2) & "\" & wSubDir(wCnt) & "\*.*"
            CisFun.Log_DataHD = wDate(wCnt2) & "-" & wSubDir(wCnt)
            Call Text_LogPut(gStr, , 10)
            On Error GoTo 0
            ' 日付毎のｻﾌﾞﾌｫﾙﾀﾞ   削除
            RmDir (Text_OutDir(Text_Proc) & wDate(wCnt2) & "\" & wSubDir(wCnt))
        Next wCnt
        ' 日付毎のﾌｫﾙﾀﾞ   削除
        RmDir (Text_OutDir(Text_Proc) & wDate(wCnt2))
        CisFun.Log_DataHD = wDate(wCnt2)
        gStr = "日付毎のﾌｫﾙﾀﾞ削除 = " & (Text_OutDir(Text_Proc) & wDate(wCnt2))
        Call Text_LogPut(gStr, "G", 5)
    Next wCnt2
    
BackUpDel_Ed:
'--------------------------------- ﾃｷｽﾄ連携変換履歴テーブル削除
    wDelTable = "テキスト変換履歴テーブル"
    gStr = wDelTable & " 削除 (処理日付 < " & wDelDay & ")  開始"
    Call Text_LogPut(gStr, "G")
    
    TRI.ERR = "32"      'ﾃｷｽﾄ連携変換履歴削除ｴﾗｰ
    If Not Middle_TRIUpDate Then Exit Function
    With CisDB
        .SQL = "DELETE " & wDelTable
        .SQL = .SQL & " Where 処理日付 < '" & wDelDay & "'"
        '.SQL = .SQL & "   And 受注出荷区分 IN ( " & Text_Proc & "," & Text_Proc + 2 & ")"
'        If Text_Proc = 0 Or Text_Proc = 1 Then
        If Text_Proc = 0 Or Text_Proc = 1 Or Text_Proc = 3 Then
           If Text_Proc = 0 Then
              .SQL = .SQL & "   And 受注出荷区分 IN ( 0,2 )"
           Else
              .SQL = .SQL & "   And 受注出荷区分 IN ( 1,3 )"
           End If
        Else
           .SQL = .SQL & "   And 受注出荷区分 IN ( " & Text_Proc & "," & Text_Proc + 2 & ")"
        End If

        If Not .DBExec Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call Text_LogPut(gStr)
           GoTo BackUpDel_Er
        End If
        gStr = "　　　　　処理件数 ＝ " & Format(.RecordCount, "#,##0")
        Call Text_LogPut(gStr, , 40)
    End With
    gStr = wDelTable & " 削除 (処理日付 < " & wDelDay & ")  終了"
    Call Text_LogPut(gStr, "G")

'--------------------------------- ﾃｷｽﾄ連携変換テーブル削除
    wDelTable = "テキスト変換テーブル"
    gStr = wDelTable & " 削除 (処理日 < " & wDelDay & ")  開始"
    Call Text_LogPut(gStr, "G")
    
    TRI.ERR = "33"      'ﾃｷｽﾄ連携変換削除ｴﾗｰ
    If Not Middle_TRIUpDate Then Exit Function
    With CisDB
        .SQL = "DELETE " & wDelTable
        .SQL = .SQL & " Where 処理日 < '" & wDelDay & "'"
        .SQL = .SQL & "   And 処理区分 = " & Text_Proc
        If Not .DBExec Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call Text_LogPut(gStr)
           GoTo BackUpDel_Er
        End If
        gStr = "　　　　　処理件数 ＝ " & Format(.RecordCount, "#,##0")
        Call Text_LogPut(gStr, , 40)
    End With
    gStr = wDelTable & " 削除 (処理日 < " & wDelDay & ")  終了"
    Call Text_LogPut(gStr, "G")
    
    ' ログ出力
    CisFun.Log_DataHD = ""
    gStr = "【 TNS" & Text_PrefixNm(Text_Proc) & "バックアップデータ削除 正常終了 】"
    Call Text_LogPut(gStr, "G", 1)
    
    Text_BackUpDel = True
    
    Exit Function
BackUpDel_Er:
    ReDim DispMsg(4)
    DispMsg(0) = " 【 ﾃｷｽﾄ連携ﾊﾞｯｸｱｯﾌﾟ" & Text_PrefixNm(Text_Proc) & " 】の削除　エラー"
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description
    
    Call Text_LogPut(DispMsg(0))
    Call Text_LogPut(DispMsg(1))
    Call Text_LogPut(DispMsg(2))
    Call Text_LogPut(DispMsg(3))
    
    ' ログ出力
    CisFun.Log_DataHD = ""
    gStr = "【 TNS" & Text_PrefixNm(Text_Proc) & "バックアップデータ削除 異常終了 】"
    Call Text_LogPut(gStr, "G", 1)
    
    If Text_ProcKB <> 0 Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(0) = DispMsg(0)
            .MB_MSG(1) = DispMsg(1)
            .MB_MSG(2) = DispMsg(2)
            .MB_MSG(3) = DispMsg(3)
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
       End With
    End If
    
End Function
'*-------------------------------------------------------------------------*
'*-------------------------------------------------------------------------*
'*-------------------------------------------------------------------------*
'*---------------                                           ---------------*
'*---------------     ﾃｷｽﾄ連携変換リスト抽出文作成(Select)     ---------------*
'*---------------                                           ---------------*
'*-------------------------------------------------------------------------*
'*-------------------------------------------------------------------------*
'*-------------------------------------------------------------------------*
Public Sub Text_Cvt_List()
    Dim ErrStr  As String
    Dim KariStr As String
    
    
    gStr = "エラー内容印刷："
    If Text_CvtErr(Text_Proc) = 0 Then
       gStr = gStr & "○"
    Else
       gStr = gStr & "×"
    End If
    gStr = gStr & "仮登録内容印刷："
    If Text_CvtKari(Text_Proc) = 0 Then
       gStr = gStr & "○"
    Else
       gStr = gStr & "×"
    End If
    gStr = gStr & "重複印刷："
    If Text_CvtDbl(Text_Proc) = 0 Then
       gStr = gStr & "○"
    Else
       gStr = gStr & "×"
    End If
    gStr = gStr & "その他印刷："
    If Text_CvtEtc(Text_Proc) = 0 Then
       gStr = gStr & "○"
    Else
       gStr = gStr & "×"
    End If
    Call Text_LogPut(gStr)
    
    
' * 変換エラーリスト
    gSL_Select = "Select TH.*, "
    gSL_Select = gSL_Select & " ISNULL(HN.表示品番,'') 表示品番,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(TR.略称,'') = '' THEN TH.納入先名 Else TR.略称 End 納入先名称 "
    gSL_Select = gSL_Select & " From テキスト変換テーブル TH "
    gSL_Select = gSL_Select & " Left Outer Join 品番マスタ HN "
    gSL_Select = gSL_Select & "     ON   HN.品番 = TH.品番"
    gSL_Select = gSL_Select & " Left Outer Join 取引先マスタ TR "
    gSL_Select = gSL_Select & "     ON   TR.取引先CD = TH.納入先"
    gSL_Select = gSL_Select & "     And  TR.取引先区分 = 0"
    gSL_Select = gSL_Select & " Where TH.処理日 = '" & Text_ProcYMD & "'"
    gSL_Select = gSL_Select & " And   TH.処理回数 = " & Text_ProcCNT & ""
    If Text_Proc <> 0 And Text_Proc <> 1 And Text_Proc <> 3 Then
       If Text_Proc <> 2 Then
          gSL_Select = gSL_Select & " And   TH.処理区分 = " & Text_Proc & ""
       Else
          gSL_Select = gSL_Select & " And   TH.処理区分 = 1"
       End If
    Else
'==================================== 2010/07/23 Start
''       gSL_Select = gSL_Select & " And   TH.処理区分 = 0 "
       gSL_Select = gSL_Select & " And   TH.処理区分 = " & Text_Proc & ""
'==================================== 2010/07/23 End
    End If
    
    gStr = ""
    ErrStr = " ERR収容数 + ERR背番号 "
    If Text_Proc <> 0 And Text_Proc <> 1 And Text_Proc <> 3 Then
       ErrStr = ErrStr & " + ERR受注 "
    End If
    KariStr = " 品番仮 + 品取仮 + 取引先仮 + 受入仮 "
    
    gCnt1 = Text_CvtErr(Text_Proc) + Text_CvtDbl(Text_Proc) + _
            Text_CvtKari(Text_Proc) + Text_CvtEtc(Text_Proc)
    
    If gCnt1 <> 4 And gCnt1 <> 0 Then
        If Text_CvtErr(Text_Proc) = 0 Then
           gStr = " AND ( ( " & ErrStr & " ) <> 0 "
        End If
        If Text_CvtDbl(Text_Proc) = 0 Then
           If gStr = "" Then
              gStr = " AND (  ERR重複  <> 0 "
           Else
              gStr = gStr & " Or  ERR重複 <> 0 "
           End If
        End If
        If Text_CvtKari(Text_Proc) = 0 Then
           If gStr = "" Then
              gStr = " AND ( ( " & KariStr & " ) <> 0 "
           Else
              gStr = gStr & " Or ( " & KariStr & " )  <> 0 "
           End If
        End If
        If Text_CvtEtc(Text_Proc) = 0 Then
           If gStr = "" Then
              gStr = " AND ( ( " & ErrStr & " + " & KariStr & " + ERR重複 ) = 0 "
           Else
              gStr = gStr & " Or ( " & ErrStr & " + " & KariStr & " + ERR重複 )  = 0 "
           End If
        End If
        
        If gStr <> "" Then
           gSL_Select = gSL_Select & gStr & " ) "
        End If
    End If
    
    gSL_Select = gSL_Select & " Order By TH.ファイル名,TH.受注管理番号,TH.受注管理行番号"
    
End Sub
'+------------------------------------------------------------------------+
'+
'+      ログ出力
'+
'+------------------------------------------------------------------------+
Public Sub Text_LogPut(LogMsg As String, Optional PutType As String = "", Optional SpCount As Long = 10)
    
    Call Text_LogProc(CisFun, LogMsg, PutType, SpCount)
    
    If Text_FunFlg Then
       Call Text_LogProc(Text_Fun, LogMsg, PutType, SpCount)
    End If
    
End Sub
Public Sub Text_LogProc(LogFun As CISFunction, LogMsg As String, Optional PutType As String = "", Optional SpCount As Long = 10)
    Dim wMsg        As String
    
    
    wMsg = Space(10) & "【 ﾃｷｽﾄ連携" & Text_PrefixNm(Text_Proc) & "データ変換処理"
    If Text_Rireki = 0 Then
        If Text_ProcKB = 0 Then
            wMsg = wMsg & "(自動)"
        Else
            wMsg = wMsg & "(手動)"
        End If
    Else
        wMsg = wMsg & "(再処理)"
    End If
    With LogFun
        If PutType = "START" Or PutType = "S" Then
            .Log_DataHD = ""
            ' ログ出力開始
            gStr = String(40, "★")
            Call .LogPut(gStr)
            Call .LogPut("☆")
            wMsg = "★" & wMsg & " ☆ 処 理 開 始 ☆ 】" & CisFun.WSID
            Call .LogPut(wMsg)
            Call .LogPut("☆")
            gStr = String(40, "★")
            Call .LogPut(gStr)
            Call .LogPut("'")
            Call .LogPut("'")
            Call .LogPut("'")
            
            Exit Sub
        End If
        If PutType = "END" Or PutType = "E" Then
            .Log_DataHD = ""
            ' ログ出力終了
            Call .LogPut("'")
            Call .LogPut("'")
            Call .LogPut("'")
            gStr = String(40, "☆")
            Call .LogPut(gStr)
            Call .LogPut("★")
            wMsg = "☆" & wMsg & " ☆ 処 理 終 了 ☆ 】" & CisFun.WSID
            Call .LogPut(wMsg)
            Call .LogPut("★")
            gStr = String(40, "☆")
            Call .LogPut(gStr)
            
            Exit Sub
        End If
        If PutType = "ABORT" Or PutType = "A" Then
            .Log_DataHD = ""
            ' ログ出力終了
            Call .LogPut("'")
            Call .LogPut("'")
            Call .LogPut("'")
            gStr = String(40, "☆")
            Call .LogPut(gStr)
            Call .LogPut("★")
            wMsg = "☆" & wMsg & " ☆ 処 理 異 常 終 了 ☆ 】" & CisFun.WSID
            Call .LogPut(wMsg)
            Call .LogPut("★")
            gStr = String(40, "☆")
            Call .LogPut(gStr)
            
            Exit Sub
        End If
        ' エラー出力
        If PutType = "ERR" Then
            wMsg = "*Err" & Space(10) & LogMsg
            Call .LogPut("'")
            gStr = "*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err"
            Call .LogPut(gStr)
            Call .LogPut("*Err")
            Call .LogPut(wMsg)
            Call .LogPut("*Err")
            gStr = "*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err*Err"
            Call .LogPut(gStr)
        End If
        If PutType = "G" Then
            Call .LogPut("'")
            wMsg = Space(SpCount) & LogMsg
            Call .LogPut(wMsg)
            Call .LogPut("'")
        End If
        If PutType = "" Then
            wMsg = Space(SpCount) & LogMsg
            Call .LogPut(wMsg)
        End If
    End With
End Sub
'*--------------------------------------------------------------------------------
'*
'*   ﾃﾞﾘﾐﾀで区切られた文字列の個々の要素を取り出し配列ヘ格納する
'*
'*   モジュールID    : GetString
'*
'*   初期設定項目    : なし
'*   引数           : Strg    -> 入力文字列
'*                  : Delimit -> ﾃﾞﾘﾐﾀ（区切り符合）
'*   戻値           : gString(??)
'*   例   :
'*    Call GetString("123,456",",")
'*
'*    結果 : gString(0) -> "123"
'*          gString(1) -> "456"
'*
'*---------------------------------------------------------------------------------
Public Sub GetString(Str As String, Demilit As String)
    Dim GetPos As Integer
    Dim GetI   As Integer
    Dim wStr   As String
    
    GetPos = 1 '最初は 1
    Do
      wStr = GetStr(Str, GetPos, Demilit)
      GetPos = GetPos + Len(wStr) + 1
      GetI = GetI + 1
      ReDim Preserve gString(GetI)
      gString(GetI - 1) = wStr
      
    Loop While GetPos <= Len(Str)
    
End Sub
'*--------------------------------------------------------------------------------
'*
'*   ﾃﾞﾘﾐﾀで区切ら
'れた文字列の個々の要素を取り出す
'*
'*   モジュールID    : GetStr
'*
'*   初期設定項目    : なし
'*   引数           : Strg    -> 入力文字列
'*                  : Pos     -> 取り出す文字位置
'*                  : Delimit -> ﾃﾞﾘﾐﾀ（区切り符合）
'*   戻値           : 数値
'*   例   :
'*  Pos = 1 '最初は 1
'*  Do
'*    B(I) = GetStr(Strg, Pos, ",")
'*    Pos = Pos + Len(B(I)) + 1
'*    I = I + 1
'*  Loop While Pos <= Len(Strg)
'*
'*   作成日  :1997/01/17   Cis K.K
'*   変更日  :
'*---------------------------------------------------------------------------------
Public Function GetStr(Strg As String, POS As Integer, Delimit As String) As String
  Dim GetI As Integer, GetL As Integer, GetA As String, GetB As String
  Dim GetJ As Integer
  
  GetL = Len(Strg)
  If POS < 1 Then POS = 1
  If POS > GetL Then
    GetStr = ""
    Exit Function
  End If

  GetB = ""
  For GetI = POS To GetL
    GetA = Mid(Strg, GetI, 1)
    If GetA = Chr(34) Then
       GetJ = InStr(GetI + 1, Strg, Chr(34), 1)
       GetStr = Mid(Strg, GetI + 1, GetJ)
      Exit Function
    End If
    If GetA = Delimit Then
      GetStr = GetB
      Exit Function
    Else
      GetB = GetB & GetA
    End If
  Next GetI
  
  GetStr = GetB  '最後の文字までデリミタがない場合

End Function


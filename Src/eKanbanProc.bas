Attribute VB_Name = "eKanbanProc"
Option Explicit
'***********************************************************************************
'       モジュール名  : 新eかんばん連携
'       モジュールID  : eKanbanProc.BAS
'       使用法        :
'       作成日        : 2008/07/14   CIS
'       変更日        :
'
'***********************************************************************************
Public eKanban_Proc           As Byte     '0:受注,1:出荷
Public eKanban_ProcKB         As Byte     '0:自動,1:手動
Public eKanban_ProcDisp       As Byte     '0:画面表示なし,1:画面表示あり

Public eKanban_Prefix(1)      As String   ' ﾌﾟﾚﾌｨｯｸｽ (0:JY,1:SK)
Public eKanban_PrefixNm(1)    As String   ' ﾌﾟﾚﾌｨｯｸｽ名 (0:受注,1:出荷)
Public eKanban_InDir(1)       As String   ' 新eかんばん端末ﾃｷｽﾄ格納先
Public eKanban_TextName(1)    As String   ' 新eかんばん端末ﾃｷｽﾄ名(ﾜｲﾙﾄﾞｶｰﾄﾞ指定可能)
Public eKanban_OutDir(1)      As String   ' SIMPR-A 新eかんばんﾃﾞｰﾀ退避先
Public eKanban_DelDay(1)      As Long     ' 保持期間(新eかんばんﾃｷｽﾄﾃﾞｰﾀ･ﾛｸﾞ･履歴情報)
Public eKanban_CvtList(1)     As Boolean  ' 変換ﾘｽﾄ発行ﾌﾗｸﾞ(True:発行あり,False:発行なし)
Public eKanban_CvtErr(1)      As Byte     '   ｴﾗｰ内容    0:対象、1:対象外
Public eKanban_CvtDbl(1)      As Byte     '   重複       0:対象、1:対象外
Public eKanban_CvtKari(1)     As Byte     '   仮登録内容 0:対象、1:対象外
Public eKanban_CvtEtc(1)      As Byte     '   その他     0:対象、1:対象外

'--------------------------------------------- 有効時間チェック
Public eKanban_TimeChk        As Boolean  ' 有効時間設定ﾌﾗｸﾞ(True:行う,False:行わない)
Public eKanban_SHHMM          As String   ' 開始時刻
Public eKanban_EHHMM          As String   ' 終了時刻
Public eKanban_HHMM           As String   ' 現在時刻

Public eKanban_ErrMsg         As String
Public eKanban_DirMaxCnt      As Long
Public eKanban_ProcYMD        As String
Public eKanban_ProcCNT        As Long
Public eKanban_Rireki         As Long     ' 履歴管理番号(再処理用)

Public eKanban_MaeKoteiFile   As String   '前工程情報出力ファイル名
Public eKanban_MaeKoteiFld    As String   '前工程情報出力フォルダ

Public eKanban_SBasyoCD       As String   '設置場所コード
Public eKanban_MachineNo      As String   'マシンNO

Public eKanban_Fun            As New CISFunction
Public eKanban_Frm            As CISFormContorl
Public eKanban_FunFlg         As Boolean

Public DispMsg()            As String

Public eKanban_Prt            As New CISFormPrint

Public eKanban_FNo            As Integer
Public eKanban_Rec            As String

'+-----------------------------------------+
'+  新eかんばん有効時間チェック
'+-----------------------------------------+
Public Function eKanbanYukou() As Boolean
    eKanbanYukou = False

    With CisFun
        .INI_Section = "eKanban"
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
           eKanbanYukou = True
           Exit Function
        End If
        '------------------------------------- 現在時刻セット
        eKanban_HHMM = Format(Time(), "hhmm")
        
        '-------- 有効時間(開始)
        .INI_Key = "有効時間_開始_時"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "5"
           .PutIni
        End If
        eKanban_SHHMM = Format(CisFun.Val2(.INI_String), "00")
        .INI_Key = "有効時間_開始_分"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "00"
           .PutIni
        End If
        eKanban_SHHMM = eKanban_SHHMM & Format(CisFun.Val2(.INI_String), "00")
        
        '-------- 有効時間(終了)
        .INI_Key = "有効時間_終了_時"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "5"
           .PutIni
        End If
        eKanban_EHHMM = Format(CisFun.Val2(.INI_String), "00")
        .INI_Key = "有効時間_終了_分"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "00"
           .PutIni
        End If
        eKanban_EHHMM = eKanban_EHHMM & Format(CisFun.Val2(.INI_String), "00")
    End With

    '---------------------- 判定
    ' 開始 > 終了
    If eKanban_SHHMM > eKanban_EHHMM Then
       If eKanban_EHHMM < eKanban_HHMM And _
          eKanban_SHHMM > eKanban_HHMM Then
       Else
          eKanbanYukou = True
       End If
       Exit Function
    End If
    ' 開始 < 終了
    If eKanban_SHHMM <= eKanban_HHMM And _
       eKanban_EHHMM >= eKanban_HHMM Then
       eKanbanYukou = True
       Exit Function
    End If
    

End Function
'+-----------------------------------------+
'+  新eかんばん環境初期情報取得
'+-----------------------------------------+
Public Function eKanbanIniGet() As Boolean

    eKanbanIniGet = False
    ' 処理名をセット
    eKanban_Prefix(0) = "JY": eKanban_PrefixNm(0) = "受注"
    eKanban_Prefix(1) = "SK": eKanban_PrefixNm(1) = "出荷"
    
    ' ログ情報
    With CisFun
        .Log_PutFlg = True
        .Log_FilePath = gIniExe & "Log"
        .Log_FileInitial = "eKanban_" & eKanban_Prefix(eKanban_Proc) & "_"
        .Log_DataHDTime = あり
    End With
    
    ' 新eかんばん設定環境取得
    With CisFun
        .INI_Section = "eKanban"

'-----     -----     -----     -----     -----     -----     < 2006/03/09 MNT START >
'        .INI_FileID = gIniExe & gIniName
'
'------------------- ｸﾗｲｱﾝﾄ個別設定
        .INI_FileID = gIniExe & gwsIniName
'
'-----     -----     -----     -----     -----     -----     < 2006/03/09 MNT  END  >

        '-------- 新eかんばんﾃｷｽﾄ格納先
        .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_ﾌｧｲﾙ格納先"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = " "
           .PutIni
        End If
        eKanban_InDir(eKanban_Proc) = Trim(.INI_String)
        
        If Trim(eKanban_InDir(eKanban_Proc)) <> "" Then
           If Right(eKanban_InDir(eKanban_Proc), 1) <> "\" Then
              eKanban_InDir(eKanban_Proc) = eKanban_InDir(eKanban_Proc) & "\"
           End If
        End If
        '-------- 新eかんばんﾃｷｽﾄ退避先(SIMPR-A)
        .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_ﾌｧｲﾙ退避先"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           gStr = Replace(UCase(gIniExe), "EXE\", "eKanban\Back\")
           .INI_String = gStr & eKanban_Prefix(eKanban_Proc) & "\"
           .PutIni
        End If
        eKanban_OutDir(eKanban_Proc) = Trim(.INI_String)
        If Trim(eKanban_OutDir(eKanban_Proc)) <> "" Then
           If Right(eKanban_OutDir(eKanban_Proc), 1) <> "\" Then
              eKanban_OutDir(eKanban_Proc) = eKanban_OutDir(eKanban_Proc) & "\"
           End If
        End If
        
'-----     -----     -----     -----     -----     -----     < 2006/03/09 ADD START >
'------------------- ｼｽﾃﾑ共通設定
        .INI_FileID = gIniExe & gIniName
'-----     -----     -----     -----     -----     -----     < 2006/03/09 MNT START >
        
        '-------- 新eかんばんﾃｷｽﾄﾌｧｲﾙ名
        .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_ﾌｧｲﾙ名"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           If eKanban_Proc = 0 Then
              .INI_String = "HSTEIKYO*"
           Else
              .INI_String = "SYUKKA*"
           End If
           .PutIni
        End If
        eKanban_TextName(eKanban_Proc) = Trim(.INI_String)
        '-------- 保持期間
        .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_保持期間"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "10"
           .PutIni
        End If
        eKanban_DelDay(eKanban_Proc) = CisFun.Val2(.INI_String)
        If eKanban_DelDay(eKanban_Proc) = 0 Then
           eKanban_DelDay(eKanban_Proc) = 10
        End If
        '-------- 変換リスト
        .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_変換リスト"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "No"
           .PutIni
        End If
        If UCase(Trim(.INI_String)) = "Y" Or UCase(Trim(.INI_String)) = "YES" Then
           eKanban_CvtList(eKanban_Proc) = True
        Else
           eKanban_CvtList(eKanban_Proc) = False
        End If
        '-------- 変換リスト(エラー内容)
        .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_変換リスト_エラー出力"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "Yes"
           .PutIni
        End If
        If UCase(Trim(.INI_String)) = "Y" Or UCase(Trim(.INI_String)) = "YES" Then
           eKanban_CvtErr(eKanban_Proc) = 0
        Else
           eKanban_CvtErr(eKanban_Proc) = 1
        End If
        '-------- 変換リスト(重複)
        .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_変換リスト_重複"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "Yes"
           .PutIni
        End If
        If UCase(Trim(.INI_String)) = "Y" Or UCase(Trim(.INI_String)) = "YES" Then
           eKanban_CvtDbl(eKanban_Proc) = 0
        Else
           eKanban_CvtDbl(eKanban_Proc) = 1
        End If
        '-------- 変換リスト(仮登録内容)
        .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_変換リスト_仮登録出力"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "Yes"
           .PutIni
        End If
        If UCase(Trim(.INI_String)) = "Y" Or UCase(Trim(.INI_String)) = "YES" Then
           eKanban_CvtKari(eKanban_Proc) = 0
        Else
           eKanban_CvtKari(eKanban_Proc) = 1
        End If
        '-------- 変換リスト(その他)
        .INI_Key = eKanban_PrefixNm(eKanban_Proc) & "_変換リスト_その他出力"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "Yes"
           .PutIni
        End If
        If UCase(Trim(.INI_String)) = "Y" Or UCase(Trim(.INI_String)) = "YES" Then
           eKanban_CvtEtc(eKanban_Proc) = 0
        Else
           eKanban_CvtEtc(eKanban_Proc) = 1
        End If
    
    End With
    eKanbanIniGet = True
End Function
'+-----------------------------------------+
'+  新eかんばん前工程情報出力環境初期情報取得
'+-----------------------------------------+
Public Function eKanban_Maekotei_IniGet() As Boolean
    eKanban_Maekotei_IniGet = False

    With CisFun
        .INI_Section = "eKanban"
    
'------------------- ｼｽﾃﾑ共通設定
        .INI_FileID = gIniExe & gIniName
        
        .INI_Key = "前工程情報_ﾌｧｲﾙ名"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "Maejh.txt"
           .PutIni
        End If
        eKanban_MaeKoteiFile = Trim(.INI_String)
    
        .INI_Key = "設置場所コード"
        .INI_Default = ""
        .GetIni
        eKanban_SBasyoCD = Trim(.INI_String)

        .INI_Key = "マシンNO"
        .INI_Default = ""
        .GetIni
        eKanban_MachineNo = Trim(.INI_String)

'------------------- ｸﾗｲｱﾝﾄ個別設定
        .INI_FileID = gIniExe & gwsIniName
        
        .INI_Key = "前工程情報_ﾌｧｲﾙ格納先"
        .INI_Default = ""
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = gIniExe
           .PutIni
        End If
        eKanban_MaeKoteiFld = Trim(.INI_String)
        
        If Trim(eKanban_MaeKoteiFld) <> "" Then
           If Right(eKanban_MaeKoteiFld, 1) <> "\" Then
              eKanban_MaeKoteiFld = eKanban_MaeKoteiFld & "\"
           End If
        End If
    End With
    
    eKanban_Maekotei_IniGet = True
End Function
'+-----------------------------------------+
'+  新eかんばん環境情報確認
'+-----------------------------------------+
Public Function eKanbanIniCheck() As Boolean
    eKanbanIniCheck = False
    eKanban_ErrMsg = ""
    
    On Error GoTo eKanbanIniCheck_Er:
    
    '-------- 新eかんばんﾃｷｽﾄ格納先
    If Trim(eKanban_InDir(eKanban_Proc)) = "" Then
       eKanban_ErrMsg = "新eかんばんﾃｷｽﾄ格納先未設定"
       GoTo eKanbanIniCheck_Er
    End If
    If eKanban_ProcKB <> 0 Then       ' 手動のみチェック
        If Dir(eKanban_InDir(eKanban_Proc), vbDirectory) = "" Then
           eKanban_ErrMsg = "新eかんばんﾃｷｽﾄ格納先(接続エラー）" & vbCrLf & _
                          Space(40) & " 接続先：" & eKanban_InDir(eKanban_Proc)
           GoTo eKanbanIniCheck_Er
        End If
    End If
    '-------- 新eかんばんﾃｷｽﾄﾌｧｲﾙ名
    If Trim(eKanban_TextName(eKanban_Proc)) = "" Then
       eKanban_ErrMsg = "新eかんばんﾃｷｽﾄ名未設定"
       GoTo eKanbanIniCheck_Er
    End If
    '-------- 新eかんばんﾃｷｽﾄ退避先(SIMPR-A)
    If Trim(eKanban_OutDir(eKanban_Proc)) = "" Then
       eKanban_ErrMsg = "新eかんばんﾃｷｽﾄ退避先未設定"
       GoTo eKanbanIniCheck_Er
    End If
    If Dir(eKanban_OutDir(eKanban_Proc), vbDirectory) = "" Then
       With CisFun
            .FE_Path = eKanban_OutDir(eKanban_Proc)
            .DirCheck NoMsg_Crt
       End With
       Call CisFun.LogPut("'")
       gStr = Space(10) & "新eかんばんﾃｷｽﾄ退避先　自動作成 作成先：" & eKanban_OutDir(eKanban_Proc)
       Call CisFun.LogPut(gStr)
       Call CisFun.LogPut("'")
    End If

    eKanbanIniCheck = True
    Exit Function
eKanbanIniCheck_Er:
    If ERR.Number = 52 Then
       With CisFun
            .MB_Lines = 6
            .MB_MSG(0) = "    【" & eKanban_PrefixNm(eKanban_Proc) & "】新eかんばん端末接続エラー"
            .MB_MSG(2) = "    Error No    = " & ERR.Number
            .MB_MSG(3) = "    Error Msg   = " & ERR.Description
            .MB_MSG(5) = "      接続先 = " & eKanban_InDir(eKanban_Proc)
            .MB_Title = "": .MB_Button = Error
            If eKanban_ProcKB <> 0 Then
               Call .MBOX
            End If
        End With
        gStr = "＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／"
        Call eKanban_LogPut(gStr, , 1)
        gStr = "'"
        Call eKanban_LogPut(gStr, , 1)
       gStr = "＿／       【" & eKanban_PrefixNm(eKanban_Proc) & "】新eかんばん端末接続エラー"
        Call eKanban_LogPut(gStr)
       gStr = "'"
        Call eKanban_LogPut(gStr, , 1)
       gStr = "＿／             接続先 = " & eKanban_InDir(eKanban_Proc)
        Call eKanban_LogPut(gStr)
       gStr = "'"
        Call eKanban_LogPut(gStr, , 1)
       gStr = "＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／"
        Call eKanban_LogPut(gStr, , 1)
       gStr = "'"
        Call eKanban_LogPut(gStr, , 1)
    End If

    If Trim(eKanban_ErrMsg) <> "" Then
       Call eKanban_LogPut(eKanban_ErrMsg, "ERR")
    End If

End Function
'+------------------------------------------------------------------------+
'+
'+      新eかんばん格納ファイル数を求める
'+
'+------------------------------------------------------------------------+
Public Function eKanban_GetFile_Count() As Long
    Dim wCount      As Long
    Dim wStrFName   As String
    wCount = 0
    
    On Error GoTo eKanban_GetFile_Count_er:

    If eKanban_Rireki = 0 Then
        wStrFName = Dir(eKanban_InDir(eKanban_Proc) & eKanban_TextName(eKanban_Proc))      ' 最初のファイル名を返します。
    Else
        wStrFName = Dir(EHR.処理退避先 & eKanban_TextName(eKanban_Proc))      ' 最初のファイル名を返します。
    End If
    Do While wStrFName <> ""                ' ループを開始します。
        wCount = wCount + 1

        wStrFName = Dir                     ' 次のファイル名を返します。
    Loop


    eKanban_GetFile_Count = wCount
    Exit Function
eKanban_GetFile_Count_er:
    eKanban_GetFile_Count = 0
End Function
'*-------------------------------------------------------------------------*
'*-------------------------------------------------------------------------*
'*-------------------------------------------------------------------------*
'*---------------                                           ---------------*
'*---------------     新eかんばん退避先　取得&作成        　      ---------------*
'*---------------                                           ---------------*
'*-------------------------------------------------------------------------*
'*-------------------------------------------------------------------------*
Public Function eKanbanBackGet() As Boolean
    Dim weKanbanErr   As Boolean
    
    eKanbanBackGet = False
    
    On Error GoTo eKanbanBackGet_er:
    If eKanban_Rireki = 0 Then
    '   #-----------------------------#
    '   #     新eかんばんﾃﾞｰﾀ連携元確認     #
    '   #-----------------------------#
        EHR.ERR = "10"      '新eかんばん端末接続ｴﾗｰ"
        If Not Middle_EHRUpDate Then Exit Function
        weKanbanErr = False
        If Dir(eKanban_InDir(eKanban_Proc), vbDirectory) = "" Then
           weKanbanErr = True
           GoTo eKanbanBackGet_er:
        End If
    
        EHR.ERR = "11"      '新eかんばん退避先取得ｴﾗｰ"
        If Not Middle_EHRUpDate Then Exit Function
        ' ログ出力
        gStr = "【 新eかんばん" & eKanban_PrefixNm(eKanban_Proc) & "データ　退避先取得&作成 処理開始 】"
        Call eKanban_LogPut(gStr, "G", 1)
    ' 退避先パスを求める
        EHR.退避先 = eKanban_OutDir(eKanban_Proc) & Format(Now(), "YYYYMMDD")
        With CisFun
            .FE_FullPath = eKanban_OutDir(eKanban_Proc) & Format(Now(), "YYYYMMDD")
            .GetUNCPath
            EHR.退避先UNC = .FE_UNCPath
        End With
        If InStr(1, EHR.退避先UNC, CisFun.WSID) <> 0 Then
           EHR.処理退避先 = EHR.退避先
        Else
           EHR.処理退避先 = EHR.退避先UNC
        End If
        If Right(EHR.退避先, 1) <> "\" Then EHR.退避先 = EHR.退避先 & "\"
        If Right(EHR.退避先UNC, 1) <> "\" Then EHR.退避先UNC = EHR.退避先UNC & "\"
        
        With CisFun
             .FE_Path = EHR.処理退避先
             If Not .DirCheck(NoMsg_Crt) Then
                ' フォルダなし
                EHR.ERR = "12"      '退避先ﾌｫﾙﾀﾞ作成ｴﾗｰ(日付)"
                If Not Middle_EHRUpDate Then Exit Function
           
                gStr = "退避先パス = " & EHR.処理退避先
                Call eKanban_LogPut(gStr)
                With CisFun
                     .MB_Lines = 5
                     .MB_MSG(0) = " 【 新eかんばん提供ファイル" & eKanban_PrefixNm(eKanban_Proc) & " 】の退避先取得&作成　エラー"
                     .MB_MSG(2) = "    退避先ﾌｫﾙﾀﾞ作成ｴﾗｰ"
                     .MB_MSG(4) = "    退避先パス = " & EHR.処理退避先
                End With
                GoTo eKanbanBackGet_e2
             End If
        End With
    ' 退避先最大回数を求める
        If Right(EHR.処理退避先, 1) <> "\" Then EHR.処理退避先 = EHR.処理退避先 & "\"
        eKanban_DirMaxCnt = 0
        gStr = Dir(EHR.処理退避先, vbDirectory)
        Do While gStr <> ""
            If gStr <> "." And gStr <> ".." Then
                If eKanban_DirMaxCnt < CisFun.Val2(gStr) Then
                   eKanban_DirMaxCnt = CisFun.Val2(gStr)
                End If
            End If
            gStr = Dir
        Loop
        eKanban_DirMaxCnt = eKanban_DirMaxCnt + 1
        EHR.処理退避先 = EHR.処理退避先 & Format(eKanban_DirMaxCnt, "000") & "\"
        EHR.退避先 = EHR.退避先 & Format(eKanban_DirMaxCnt, "000") & "\"
        EHR.退避先UNC = EHR.退避先UNC & Format(eKanban_DirMaxCnt, "000") & "\"
        ' ログ出力
        EHR.ERR = "13"      '退避先ﾌｫﾙﾀﾞ作成ｴﾗｰ(回数)"
        If Not Middle_EHRUpDate Then Exit Function
        
        gStr = "退避先パス = " & EHR.処理退避先
        Call eKanban_LogPut(gStr)
        
        MkDir (EHR.処理退避先)
        
        ' ログ出力
        gStr = "【 新eかんばん" & eKanban_PrefixNm(eKanban_Proc) & "データ　退避先取得&作成 処理終了 】"
        Call eKanban_LogPut(gStr, "G", 1)
    End If
    
    ' ログ情報
    With eKanban_Fun
        .Log_PutFlg = True
        .Log_FilePath = EHR.処理退避先
        .Log_FileInitial = eKanban_Prefix(eKanban_Proc) & "_" & Format(eKanban_DirMaxCnt, "000") & "_"
        .Log_DataHDTime = あり
    End With
    
    With CisFun
        .Log_DataHD = Format(eKanban_DirMaxCnt, "000")
    End With
    ' 履歴情報内容更新
    EHR.退避連番 = eKanban_DirMaxCnt
    
    eKanban_FunFlg = True
    
    eKanbanBackGet = True
    Exit Function
eKanbanBackGet_er:
    If ERR.Number = 52 Or weKanbanErr Then
       With CisFun
            .MB_Lines = 6
            .MB_MSG(0) = "    【" & eKanban_PrefixNm(eKanban_Proc) & "】新eかんばん端末接続エラー"
            .MB_MSG(2) = "    Error No    = " & ERR.Number
            .MB_MSG(3) = "    Error Msg   = " & ERR.Description
            .MB_MSG(5) = "      接続先 = " & eKanban_InDir(eKanban_Proc)
            .MB_Title = "": .MB_Button = Error
            If eKanban_ProcKB <> 0 Then
               Call .MBOX
            End If
        End With
        gStr = "＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／"
        Call eKanban_LogPut(gStr, , 1)
        gStr = "'"
        Call eKanban_LogPut(gStr, , 1)
        gStr = "＿／       【" & eKanban_PrefixNm(eKanban_Proc) & "】新eかんばん端末接続エラー"
        Call eKanban_LogPut(gStr)
        gStr = "'"
        Call eKanban_LogPut(gStr, , 1)
        gStr = "＿／             接続先 = " & eKanban_InDir(eKanban_Proc)
        Call eKanban_LogPut(gStr)
        gStr = "'"
        Call eKanban_LogPut(gStr, , 1)
        gStr = "＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／＿／"
        Call eKanban_LogPut(gStr, , 1)
        gStr = "'"
        Call eKanban_LogPut(gStr, , 1)
        Exit Function
    End If
    
    ReDim DispMsg(5)
    DispMsg(0) = " 【 新eかんばん提供ファイル" & eKanban_PrefixNm(eKanban_Proc) & " 】の退避先取得&作成　エラー"
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description
    DispMsg(4) = "    退避先パス = " & EHR.処理退避先
eKanbanBackGet_e2:
    Call eKanban_LogPut(DispMsg(0))
    Call eKanban_LogPut(DispMsg(1))
    Call eKanban_LogPut(DispMsg(2))
    Call eKanban_LogPut(DispMsg(3))
    Call eKanban_LogPut(DispMsg(4))
    gStr = "【 新eかんばん" & eKanban_PrefixNm(eKanban_Proc) & "データ　退避先取得&作成 異常終了 】"
    Call eKanban_LogPut(gStr, "G", 1)
    
    If eKanban_ProcKB <> 0 Then
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
'=======        新eかんばん側データのコピー＆読込用テキストパスセット      =======
'=======        読込テキストのバックアップ                           =======
'=======                                                             =======
'===========================================================================
'===========================================================================
'*----------------------------------------------------*
'*  新eかんばん側データのコピー＆読込用テキストパスセット   *
'*----------------------------------------------------*
Public Function eKanban_TextMove(Frm As Form) As Boolean
    eKanban_TextMove = False
    
    gStr = "【 新eかんばん" & eKanban_PrefixNm(eKanban_Proc) & "データ　移動 開始 】"
    Call eKanban_LogPut(gStr, "G", 1)

    EHR.ERR = "20"      '新eかんばんTEXT移動ｴﾗｰ"
    If Not Middle_EHRUpDate Then Exit Function
'   *------------------------------------*
'   * 自社バックアップフォルダへのコピー *
'   *------------------------------------*
    ReDim destFName(1)
    destFName(0) = eKanban_InDir(eKanban_Proc) & eKanban_TextName(eKanban_Proc)
    ' ログ出力
    gStr = "新eかんばん→TNS　移動 (" & destFName(0) & "→" & EHR.処理退避先 & ")"
    Call eKanban_LogPut(gStr)
    EHR.ERR = "21"      '新eかんばんﾌｧｲﾙ移動ｴﾗｰ"
    If Not Middle_EHRUpDate Then Exit Function
    
    If eKanban_ProcKB = 0 And eKanban_ProcDisp = 0 Then
       gInt = 4
    Else
       gInt = 2
    End If
    If Not gFileDestroy(Frm, gInt, EHR.処理退避先) Then
       gStr = "ファイル又はディスクが読み取れません。"
       Call eKanban_LogPut(gStr)
       RmDir EHR.退避先
       EHR.退避先 = ""
       GoTo Move_Err
    End If
    ' ログ出力
    gStr = "新eかんばん→TNS　移動終了"
    Call eKanban_LogPut(gStr)
    gStr = "【 新eかんばん" & eKanban_PrefixNm(eKanban_Proc) & "データ　移動 正常終了 】"
    Call eKanban_LogPut(gStr, "G", 1)
'
    On Error GoTo 0
    eKanban_TextMove = True
    
    Exit Function

Move_Err:
    ReDim DispMsg(5)
    DispMsg(0) = " 【 新eかんばん提供ファイル" & eKanban_PrefixNm(eKanban_Proc) & " 】の移動　エラー"
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description
    DispMsg(4) = "    提供ファイル = " & eKanban_TextName(eKanban_Proc)
    
    Call eKanban_LogPut(DispMsg(0))
    Call eKanban_LogPut(DispMsg(1))
    Call eKanban_LogPut(DispMsg(2))
    Call eKanban_LogPut(DispMsg(3))
    Call eKanban_LogPut(DispMsg(4))
    
    gStr = "【 新eかんばん" & eKanban_PrefixNm(eKanban_Proc) & "データ　移動 異常終了 】"
    Call eKanban_LogPut(gStr, "G", 1)
    
    If eKanban_ProcKB <> 0 Then
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
Public Function eKanban_BULK_Insert(Optional FPath As String = "") As Boolean
    Dim wPathName       As String
    Dim wFileName       As String
    Dim wCnt            As Long
    Dim wFName()        As String
    Dim FIdx            As Long
    Dim FIdx2           As Long
    Dim BULKPath        As String
    
    Erase wFName
    ReDim wFName(1)
    
    eKanban_BULK_Insert = False
    
    On Error GoTo eKanban_BULK_Insert_Err:
    
    ' ログ出力
    gStr = "【 新eかんばん" & eKanban_PrefixNm(eKanban_Proc) & "テキストデータ　コンバート 開始 】"
    Call eKanban_LogPut(gStr, "G", 1)
    
    EHR.ERR = "41"      '新eかんばんﾃｷｽﾄｲﾝﾎﾟｰﾄｴﾗｰ
    If Not Middle_EHRUpDate Then Exit Function
    
    
    If FPath = "" Then
       FPath = EHR.処理退避先
    End If
    
    BULKPath = EHR.退避先UNC
        
    If UCase(CisFun.WSID) = UCase(Trim(CisDB.Server)) Then
        If InStr(1, EHR.退避先UNC, CisFun.WSID) <> 0 Then
           BULKPath = EHR.退避先
        Else
           BULKPath = EHR.退避先UNC
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
    wFileName = Dir(FPath & eKanban_TextName(eKanban_Proc))
    If wFileName <> "" Then
       gStr = "コンバートパス名　　 ： " & FPath
       Call eKanban_LogPut(gStr, "G")
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
        Call eKanban_LogPut(gStr, , 10)
        '----------------------------------------- テーブルインサート
'        If Not eKanban_Bulk_Proc(BULKPath & wFileName, wFileName) Then
        If Not eKanban_Text_Proc(BULKPath & wFileName, wFileName) Then
           gStr = "【 新eかんばん" & eKanban_PrefixNm(eKanban_Proc) & "テキストデータ　コンバート 異常終了 】"
           Call eKanban_LogPut(gStr, "G", 1)
           Exit Function
        End If
        
        If Not eKanban_Frm Is Nothing Then
           eKanban_Frm.MeterUpdate wCnt
        End If
    Next FIdx
    
    If wCnt = 0 Then
        EHR.ERR = "46"      'ｲﾝﾎﾟｰﾄ対象ﾌｧｲﾙなし
        If Not Middle_EHRUpDate Then Exit Function
        ReDim DispMsg(4)
        DispMsg(2) = "   コンバート対象ファイルなし"
        DispMsg(3) = "   "
        GoTo eKanban_BULK_Insert_ErrDisp:
    End If

    ' ログ出力
    gStr = "【 新eかんばん" & eKanban_PrefixNm(eKanban_Proc) & "テキストデータ　コンバート 終了 】"
    Call eKanban_LogPut(gStr, "G", 1)

    eKanban_BULK_Insert = True
    Exit Function
eKanban_BULK_Insert_Err:
    ReDim DispMsg(4)
    DispMsg(0) = " 【 テキストデータコンバート 】エラー"
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description
    
eKanban_BULK_Insert_ErrDisp:
    
    Call eKanban_LogPut(DispMsg(0))
    Call eKanban_LogPut(DispMsg(1))
    Call eKanban_LogPut(DispMsg(2))
    Call eKanban_LogPut(DispMsg(3))
    
    gStr = "【 新eかんばん" & eKanban_PrefixNm(eKanban_Proc) & "テキストデータ　コンバート 異常終了 】"
    Call eKanban_LogPut(gStr, "G", 1)
    
    If eKanban_ProcKB <> 0 Then
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

Public Function eKanban_Bulk_Proc(wFullFile As String, wFileName As String) As Boolean
    '2008/07/18 eKanban_Text_Proc に変更
    eKanban_Bulk_Proc = False
    Static sw_Table         As String
    Static sw_Stoado        As String

    On Error GoTo eKanban_Bulk_Proc_Err
    
    If eKanban_Proc = 0 Then
       sw_Table = "eかんばん受注データ"
       sw_Stoado = "eかんばん受注TEXT変換"
    Else
       sw_Table = "eかんばん出荷データ"
       sw_Stoado = "eかんばん出荷TEXT変換"
    End If

'[ データ削除 ]
    EHR.ERR = "43"      '新eかんばんｲﾝﾎﾟｰﾄ削除ｴﾗｰ
    If Not Middle_EHRUpDate Then Exit Function
    With CisDB
        .SQL = "TRUNCATE TABLE " & sw_Table
        If Not .DBExec Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call eKanban_LogPut(gStr)
           GoTo eKanban_Bulk_Proc_Err
        End If
    End With
'[ BULK INSERT ]
    EHR.ERR = "44"      '新eかんばんｲﾝﾎﾟｰﾄ処理ｴﾗｰ
    If Not Middle_EHRUpDate Then Exit Function
    
    gSL_Select = "BULK INSERT " & sw_Table
    gSL_Select = gSL_Select & " From '" & wFullFile & "'"
    gSL_Select = gSL_Select & " With ("
    gSL_Select = gSL_Select & " ROWTERMINATOR = '\n'"
    gSL_Select = gSL_Select & " )"
    With CisDB
        .SQL = gSL_Select
        If Not .DBExec Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call eKanban_LogPut(gStr)
           GoTo eKanban_Bulk_Proc_Err
        End If
        gStr = "　　　　　処理件数 ＝ " & Format(.RecordCount, "#,##0")
        Call eKanban_LogPut(gStr, , 40)
    End With
    
'[TEXT変換]
    EHR.ERR = "45"      '新eかんばんTEXT変換ｴﾗｰ
    If Not Middle_EHRUpDate Then Exit Function
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
        .ParaNo = 2: .ParaValue = EHR.処理回数
        .ParaNo = 3: .ParaValue = EHR.履歴管理番号
        
        gCnt1 = InStr(1, wFileName, ".")
        If gCnt1 > 0 Then
           gStr = Mid(wFileName, 1, gCnt1 - 1)
        Else
           gStr = wFileName
        End If
        .ParaNo = 4: .ParaValue = Mid(gStr, 1, 50)
    
        If Not .DBStored Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call eKanban_LogPut(gStr)
           GoTo eKanban_Bulk_Proc_Err
        End If
        .ParaNo = 5
        If .ParaValue <> 0 Then
            ReDim DispMsg(4)
            DispMsg(0) = "   " & sw_Stoado & "エラー"
            DispMsg(2) = "   Err Code = " & .ParaValue
            DispMsg(3) = "    File = " & wFullFile
            GoTo eKanban_Bulk_Proc_ErrDisp
        End If
    End With
    
'------------ 新eかんばん変換履歴テーブル 再入力
    gLong = EHR.履歴管理番号
    gSL_Select = "Select * From eかんばん変換履歴テーブル"
    gSL_Select = gSL_Select & " Where 履歴管理番号 = " & EHR.履歴管理番号 & ""
    If Not EHRRead(gSL_Select, 1) Then
        ReDim DispMsg(4)
        DispMsg(0) = "      " & sw_Stoado & "エラー"
        DispMsg(2) = "    新eかんばん変換履歴テーブル異常            "
        DispMsg(3) = "    履歴管理番号 = " & gLong
        GoTo eKanban_Bulk_Proc_ErrDisp
    End If

    eKanban_Bulk_Proc = True
    Exit Function
eKanban_Bulk_Proc_Err:
    ReDim DispMsg(4)
    DispMsg(0) = "   Insert Err"
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description

eKanban_Bulk_Proc_ErrDisp:

    Call eKanban_LogPut(DispMsg(0))
    Call eKanban_LogPut(DispMsg(1))
    Call eKanban_LogPut(DispMsg(2))
    Call eKanban_LogPut(DispMsg(3))

    If eKanban_ProcKB <> 0 Then
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
Public Function eKanban_Text_Proc(wFullFile As String, wFileName As String) As Boolean
    eKanban_Text_Proc = False
    Static sw_Table         As String
    Static sw_Stoado        As String

    On Error GoTo eKanban_Text_Proc_Err
    
    If eKanban_Proc = 0 Then
       sw_Table = "eかんばん受注データ"
       sw_Stoado = "eかんばん受注TEXT変換"
    Else
       sw_Table = "eかんばん出荷データ"
       sw_Stoado = "eかんばん出荷TEXT変換"
    End If

'[ データ削除 ]
    EHR.ERR = "43"      '新eかんばんｲﾝﾎﾟｰﾄ削除ｴﾗｰ
    If Not Middle_EHRUpDate Then Exit Function
    With CisDB
        .SQL = "TRUNCATE TABLE " & sw_Table
        If Not .DBExec Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call eKanban_LogPut(gStr)
           GoTo eKanban_Text_Proc_Err
        End If
    End With
'[ BULK INSERT ]
    EHR.ERR = "44"      '新eかんばんｲﾝﾎﾟｰﾄ処理ｴﾗｰ
    If Not Middle_EHRUpDate Then Exit Function
    
    eKanban_FNo = FreeFile
    gInt = 0
    Open wFullFile For Input As #eKanban_FNo
    Do While Not EOF(eKanban_FNo)
        Line Input #eKanban_FNo, eKanban_Rec

        gInt = gInt + 1
        If eKanban_Proc = 0 Then
           gSL_Select = "INSERT INTO eかんばん受注データ (REC1,REC2,REC3,REC4,REC5,REC6,REC7)"
           gSL_Select = gSL_Select & " VALUES ('" & CisFun.Mid2(eKanban_Rec, 1, 323) & "',"
           gSL_Select = gSL_Select & "'" & CisFun.Mid2(eKanban_Rec, 324, 260) & "',"
           gSL_Select = gSL_Select & "'" & CisFun.Mid2(eKanban_Rec, 584, 162) & "',"
           gSL_Select = gSL_Select & "'" & CisFun.Mid2(eKanban_Rec, 746, 68) & "',"
           gSL_Select = gSL_Select & "'" & CisFun.Mid2(eKanban_Rec, 814, 692) & "',"
           gSL_Select = gSL_Select & "'" & CisFun.Mid2(eKanban_Rec, 1506, 100) & "',"
           gSL_Select = gSL_Select & "'" & CisFun.Mid2(eKanban_Rec, 1606, 23) & "')"
        Else
           gSL_Select = "INSERT INTO eかんばん出荷データ (REC1,REC2,REC3)"
           gSL_Select = gSL_Select & " VALUES ('" & CisFun.Mid2(eKanban_Rec, 1, 356) & "',"
           gSL_Select = gSL_Select & "'" & CisFun.Mid2(eKanban_Rec, 357, 100) & "',"
           gSL_Select = gSL_Select & "'" & CisFun.Mid2(eKanban_Rec, 457, 32) & "')"
        End If
        With CisDB
            .SQL = gSL_Select
            If Not .DBExec Then
               gStr = .ErrCode & ":" & .ErrMsg
               Call eKanban_LogPut(gStr)
               GoTo eKanban_Text_Proc_Err
            End If
            If eKanban_Proc = 0 Then
                gStr = " " & Format(Format(gInt, "#,###"), "@@@@@") & "：" & CisFun.Mid2(eKanban_Rec, 16, 7) & CisFun.Mid2(eKanban_Rec, 45, 17)
                gStr = gStr & CisFun.Mid2(eKanban_Rec, 155, 6) & CisFun.Mid2(eKanban_Rec, 584, 41)
            Else
                gStr = " " & Format(Format(gInt, "#,###"), "@@@@@") & "：" & CisFun.Mid2(eKanban_Rec, 16, 7) & CisFun.Mid2(eKanban_Rec, 45, 17)
                gStr = gStr & CisFun.Mid2(eKanban_Rec, 66, 10) & CisFun.Mid2(eKanban_Rec, 77, 6)
                gStr = gStr & CisFun.Mid2(eKanban_Rec, 88, 30) & CisFun.Mid2(eKanban_Rec, 14, 7)
            End If
            Call eKanban_LogPut(gStr, , 5)
        End With
        
    Loop

    Close #eKanban_FNo
        
'''    gSL_Select = "BULK INSERT " & sw_Table
'''    gSL_Select = gSL_Select & " From '" & wFullFile & "'"
'''    gSL_Select = gSL_Select & " With ("
'''    gSL_Select = gSL_Select & " ROWTERMINATOR = '\n'"
'''    gSL_Select = gSL_Select & " )"
'''    With CisDB
'''        .SQL = gSL_Select
'''        If Not .DBExec Then
'''           gStr = .ErrCode & ":" & .ErrMsg
'''           Call eKanban_LogPut(gStr)
'''           GoTo eKanban_Text_Proc_Err
'''        End If
'''        gStr = "　　　　　処理件数 ＝ " & Format(.RecordCount, "#,##0")
'''        Call eKanban_LogPut(gStr, , 40)
'''    End With
    
'[TEXT変換]
    EHR.ERR = "45"      '新eかんばんTEXT変換ｴﾗｰ
    If Not Middle_EHRUpDate Then Exit Function
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
        .ParaNo = 2: .ParaValue = EHR.処理回数
        .ParaNo = 3: .ParaValue = EHR.履歴管理番号
        
        gCnt1 = InStr(1, wFileName, ".")
        If gCnt1 > 0 Then
           gStr = Mid(wFileName, 1, gCnt1 - 1)
        Else
           gStr = wFileName
        End If
        .ParaNo = 4: .ParaValue = Mid(gStr, 1, 50)
    
        If Not .DBStored Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call eKanban_LogPut(gStr)
           GoTo eKanban_Text_Proc_Err
        End If
        .ParaNo = 5
        If .ParaValue <> 0 Then
            ReDim DispMsg(4)
            DispMsg(0) = "   " & sw_Stoado & "エラー"
            DispMsg(2) = "   Err Code = " & .ParaValue
            DispMsg(3) = "    File = " & wFullFile
            GoTo eKanban_Text_Proc_ErrDisp
        End If
    End With
    
'------------ 新eかんばん変換履歴テーブル 再入力
    gLong = EHR.履歴管理番号
    gSL_Select = "Select * From eかんばん変換履歴テーブル"
    gSL_Select = gSL_Select & " Where 履歴管理番号 = " & EHR.履歴管理番号 & ""
    If Not EHRRead(gSL_Select, 1) Then
        ReDim DispMsg(4)
        DispMsg(0) = "      " & sw_Stoado & "エラー"
        DispMsg(2) = "    新eかんばん変換履歴テーブル異常            "
        DispMsg(3) = "    履歴管理番号 = " & gLong
        GoTo eKanban_Text_Proc_ErrDisp
    End If

    eKanban_Text_Proc = True
    Exit Function
eKanban_Text_Proc_Err:
    ReDim DispMsg(4)
    DispMsg(0) = "   Insert Err"
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description

eKanban_Text_Proc_ErrDisp:

    Call eKanban_LogPut(DispMsg(0))
    Call eKanban_LogPut(DispMsg(1))
    Call eKanban_LogPut(DispMsg(2))
    Call eKanban_LogPut(DispMsg(3))

    If eKanban_ProcKB <> 0 Then
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
'*---------------     新eかんばん　更　新　処　理                                                   ---------------*
'*---------------                                                                             ---------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
'*-----------------------------------------------------------------------------------------------------------*
Public Function eKanban_UpdateProc() As Boolean
    On Error GoTo eKanban_UpdateProc_Er:
    
    eKanban_UpdateProc = False
    
    ' ログ出力
    gStr = "【 新eかんばん" & eKanban_PrefixNm(eKanban_Proc) & "更新処理 開始 】"
    Call eKanban_LogPut(gStr, "G", 1)
    
    EHR.ERR = "51"      '新eかんばん更新ｴﾗｰ
    If Not Middle_EHRUpDate Then Exit Function
    With CisDB
        .SQL = "eかんばん受注出荷更新"
        .StoadoCount = 5
        .ParaNo = 0:    .ParaIO = Return用
        .ParaNo = 1:    .ParaIO = Input用
        .ParaNo = 2:    .ParaIO = Input用
        .ParaNo = 3:    .ParaIO = Input用
        .ParaNo = 4:    .ParaIO = Input用
        .ParaNo = 5:    .ParaIO = OutPut用
    
        .ParaNo = 1: .ParaValue = eKanban_Proc
        .ParaNo = 2: .ParaValue = Format(Now(), "YYYYMMDD")
        .ParaNo = 3: .ParaValue = eKanban_ProcCNT
        .ParaNo = 4: .ParaValue = EHR.履歴管理番号
    
        If Not .DBStored Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call eKanban_LogPut(gStr)
           GoTo eKanban_UpdateProc_Er
        End If
        .ParaNo = 5
        If .ParaValue <> 0 Then
            ReDim DispMsg(4)
            DispMsg(0) = "   新eかんばん更新エラー"
            DispMsg(2) = "   Err Code = " & .ParaValue
            GoTo eKanban_UpdateProc_ErDisp
        End If
    End With
    
'------------ 新eかんばん変換履歴テーブル 再入力
    gLong = EHR.履歴管理番号
    gSL_Select = "Select * From eかんばん変換履歴テーブル"
    gSL_Select = gSL_Select & " Where 履歴管理番号 = " & EHR.履歴管理番号 & ""
    If Not EHRRead(gSL_Select, 1) Then
        ReDim DispMsg(4)
        DispMsg(0) = "         新eかんばん更新エラー                 "
        DispMsg(2) = "    新eかんばん変換履歴テーブル異常            "
        DispMsg(3) = "    履歴管理番号 = " & gLong
        GoTo eKanban_UpdateProc_ErDisp
    End If
    
    gStr = "対象件数         =" & Format(EHR.入力件数, "#,##0"): Call eKanban_LogPut(gStr, "", 20)
    gStr = "出力件数         =" & Format(EHR.変換件数, "#,##0"): Call eKanban_LogPut(gStr, "", 20)
    gStr = "重複件数         =" & Format(EHR.重複, "#,##0"): Call eKanban_LogPut(gStr, "", 20)
    If eKanban_Proc = 0 Then
       gStr = "出荷指示書空件数 =" & Format(EHR.空, "#,##0"): Call eKanban_LogPut(gStr, "", 20)
    Else
       gStr = "受注なし件数     =" & Format(EHR.ERR受注, "#,##0"): Call eKanban_LogPut(gStr, "", 20)
    End If
    gStr = "仮登録件数　　　 :(品番ﾏｽﾀ)=" & Format(EHR.仮品番, "#,##0") & _
                             " (品番取引先ﾏｽﾀ)=" & Format(EHR.仮品取, "#,##0") & _
                             " (取引先ﾏｽﾀ)=" & Format(EHR.仮取引先, "#,##0") & _
                             " (取引先受入ﾏｽﾀ)=" & Format(EHR.仮受入, "#,##0")
    Call eKanban_LogPut(gStr, "", 20)
    
    ' ログ出力
    gStr = "【 新eかんばん" & eKanban_PrefixNm(eKanban_Proc) & "更新処理 終了 】"
    Call eKanban_LogPut(gStr, "G", 1)
    
    eKanban_UpdateProc = True
    Exit Function
eKanban_UpdateProc_Er:
    ReDim DispMsg(4)
    DispMsg(0) = "   新eかんばん Update Err"
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description

eKanban_UpdateProc_ErDisp:
    Call eKanban_LogPut(DispMsg(0))
    Call eKanban_LogPut(DispMsg(1))
    Call eKanban_LogPut(DispMsg(2))
    Call eKanban_LogPut(DispMsg(3))

    If eKanban_ProcKB <> 0 Then
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
    gStr = "【 新eかんばん" & eKanban_PrefixNm(eKanban_Proc) & "更新処理 異常終了 】"
    Call eKanban_LogPut(gStr, "G", 1)
    
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
Public Function eKanban_BackUpDel() As Boolean
    Dim wDate()     As String
    Dim wCnt        As Long
    Dim wCnt2       As Long
    Dim wDirName    As String
    Dim wSubDir()   As String
    Dim wDelDay     As String
    Dim wDelTable   As String
    
    eKanban_BackUpDel = False
    
    If eKanban_Rireki <> 0 Then
       eKanban_BackUpDel = True
       Exit Function
    End If
    
    On Error GoTo BackUpDel_Er:
    '+---------------------------------+
    '+  新eかんばんﾊﾞｯｸｱｯﾌﾟ削除
    '+---------------------------------+
    CisFun.Log_DataHD = ""
    
    gStr = "【 新eかんばん" & eKanban_PrefixNm(eKanban_Proc) & "バックアップデータ削除 開始 】"
    ' ログ出力
    Call eKanban_LogPut(gStr, "G", 1)
    
    EHR.ERR = "31"      '新eかんばんﾊﾞｯｸｱｯﾌﾟ削除ｴﾗｰ"
    If Not Middle_EHRUpDate Then Exit Function
    
' 日付ﾌｫﾙﾀﾞ名を求める
    wCnt = 0
    wDelDay = Format(Now(), "YYYYMMDD")
    With CisFun
        .Date_RsFormat = fYMD
        .Date_Add wDelDay, D, eKanban_DelDay(eKanban_Proc) * -1
        wDelDay = .Date_Result
    End With
    wDirName = Dir(eKanban_OutDir(eKanban_Proc), vbDirectory)
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
       Call eKanban_LogPut(gStr)
       GoTo BackUpDel_Ed
    End If
    For wCnt2 = 0 To UBound(wDate) - 1
        ' 日付毎のｻﾌﾞﾌｫﾙﾀﾞを求める
        wCnt = 0
        Erase wSubDir
        
        wDirName = Dir(eKanban_OutDir(eKanban_Proc) & wDate(wCnt2) & "\", vbDirectory)
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
        Call eKanban_LogPut(gStr, "G")
        For wCnt = 0 To UBound(wSubDir) - 1
            On Error Resume Next
            Kill eKanban_OutDir(eKanban_Proc) & wDate(wCnt2) & "\" & wSubDir(wCnt) & "\*.*"
            ' ログ出力
            gStr = "削除内容 = " & eKanban_OutDir(eKanban_Proc) & wDate(wCnt2) & "\" & wSubDir(wCnt) & "\*.*"
            CisFun.Log_DataHD = wDate(wCnt2) & "-" & wSubDir(wCnt)
            Call eKanban_LogPut(gStr, , 10)
            On Error GoTo 0
            ' 日付毎のｻﾌﾞﾌｫﾙﾀﾞ   削除
            RmDir (eKanban_OutDir(eKanban_Proc) & wDate(wCnt2) & "\" & wSubDir(wCnt))
        Next wCnt
        ' 日付毎のﾌｫﾙﾀﾞ   削除
        RmDir (eKanban_OutDir(eKanban_Proc) & wDate(wCnt2))
        CisFun.Log_DataHD = wDate(wCnt2)
        gStr = "日付毎のﾌｫﾙﾀﾞ削除 = " & (eKanban_OutDir(eKanban_Proc) & wDate(wCnt2))
        Call eKanban_LogPut(gStr, "G", 5)
    Next wCnt2
    
BackUpDel_Ed:
'--------------------------------- 新eかんばん変換履歴テーブル削除
    wDelTable = "eかんばん変換履歴テーブル"
    gStr = wDelTable & " 削除 (処理日付 < " & wDelDay & ")  開始"
    Call eKanban_LogPut(gStr, "G")
    
    EHR.ERR = "32"      '新eかんばん変換履歴削除ｴﾗｰ
    If Not Middle_EHRUpDate Then Exit Function
    With CisDB
        .SQL = "DELETE " & wDelTable
        .SQL = .SQL & " Where 処理日付 < '" & wDelDay & "'"
        .SQL = .SQL & "   And 受注出荷区分 IN ( " & eKanban_Proc & "," & eKanban_Proc + 2 & ")"
        If Not .DBExec Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call eKanban_LogPut(gStr)
           GoTo BackUpDel_Er
        End If
        gStr = "　　　　　処理件数 ＝ " & Format(.RecordCount, "#,##0")
        Call eKanban_LogPut(gStr, , 40)
    End With
    gStr = wDelTable & " 削除 (処理日付 < " & wDelDay & ")  終了"
    Call eKanban_LogPut(gStr, "G")

'--------------------------------- 新eかんばん変換テーブル削除
    wDelTable = "eかんばん変換テーブル"
    gStr = wDelTable & " 削除 (処理日 < " & wDelDay & ")  開始"
    Call eKanban_LogPut(gStr, "G")
    
    EHR.ERR = "33"      '新eかんばん変換削除ｴﾗｰ
    If Not Middle_EHRUpDate Then Exit Function
    With CisDB
        .SQL = "DELETE " & wDelTable
        .SQL = .SQL & " Where 処理日 < '" & wDelDay & "'"
        .SQL = .SQL & "   And 処理区分 = " & eKanban_Proc
        If Not .DBExec Then
           gStr = .ErrCode & ":" & .ErrMsg
           Call eKanban_LogPut(gStr)
           GoTo BackUpDel_Er
        End If
        gStr = "　　　　　処理件数 ＝ " & Format(.RecordCount, "#,##0")
        Call eKanban_LogPut(gStr, , 40)
    End With
    gStr = wDelTable & " 削除 (処理日 < " & wDelDay & ")  終了"
    Call eKanban_LogPut(gStr, "G")
    
    ' ログ出力
    CisFun.Log_DataHD = ""
    gStr = "【 TNS" & eKanban_PrefixNm(eKanban_Proc) & "バックアップデータ削除 正常終了 】"
    Call eKanban_LogPut(gStr, "G", 1)
    
    eKanban_BackUpDel = True
    
    Exit Function
BackUpDel_Er:
    ReDim DispMsg(4)
    DispMsg(0) = " 【 新eかんばんﾊﾞｯｸｱｯﾌﾟ" & eKanban_PrefixNm(eKanban_Proc) & " 】の削除　エラー"
    DispMsg(2) = "    Error No    = " & ERR
    DispMsg(3) = "    Error Msg   = " & ERR.Description
    
    Call eKanban_LogPut(DispMsg(0))
    Call eKanban_LogPut(DispMsg(1))
    Call eKanban_LogPut(DispMsg(2))
    Call eKanban_LogPut(DispMsg(3))
    
    ' ログ出力
    CisFun.Log_DataHD = ""
    gStr = "【 TNS" & eKanban_PrefixNm(eKanban_Proc) & "バックアップデータ削除 異常終了 】"
    Call eKanban_LogPut(gStr, "G", 1)
    
    If eKanban_ProcKB <> 0 Then
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
'*---------------     新eかんばん変換リスト抽出文作成(Select)     ---------------*
'*---------------                                           ---------------*
'*-------------------------------------------------------------------------*
'*-------------------------------------------------------------------------*
'*-------------------------------------------------------------------------*
Public Sub eKanban_Cvt_List()
    Dim ErrStr  As String
    Dim KariStr As String
    
    
    gStr = "エラー内容印刷："
    If eKanban_CvtErr(eKanban_Proc) = 0 Then
       gStr = gStr & "○"
    Else
       gStr = gStr & "×"
    End If
    gStr = gStr & "仮登録内容印刷："
    If eKanban_CvtKari(eKanban_Proc) = 0 Then
       gStr = gStr & "○"
    Else
       gStr = gStr & "×"
    End If
    gStr = gStr & "重複印刷："
    If eKanban_CvtDbl(eKanban_Proc) = 0 Then
       gStr = gStr & "○"
    Else
       gStr = gStr & "×"
    End If
    gStr = gStr & "その他印刷："
    If eKanban_CvtEtc(eKanban_Proc) = 0 Then
       gStr = gStr & "○"
    Else
       gStr = gStr & "×"
    End If
    Call eKanban_LogPut(gStr)
    
    
' * 変換エラーリスト
    gSL_Select = "Select TH.*, "
    gSL_Select = gSL_Select & " ISNULL(HN.表示品番,'') 表示品番,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(TR.略称,'') = '' THEN TH.納入先名 Else TR.略称 End 納入先名称 "
    gSL_Select = gSL_Select & " From eかんばん変換テーブル TH "
    gSL_Select = gSL_Select & " Left Outer Join 品番マスタ HN "
    gSL_Select = gSL_Select & "     ON   HN.品番 = TH.品番"
    gSL_Select = gSL_Select & " Left Outer Join 取引先マスタ TR "
    gSL_Select = gSL_Select & "     ON   TR.取引先CD = TH.納入先"
    gSL_Select = gSL_Select & "     And  TR.取引先区分 = 0"
    gSL_Select = gSL_Select & " Where TH.処理日 = '" & eKanban_ProcYMD & "'"
    gSL_Select = gSL_Select & " And   TH.処理回数 = " & eKanban_ProcCNT & ""
    gSL_Select = gSL_Select & " And   TH.処理区分 = " & eKanban_Proc & ""
    
    gStr = ""
    ErrStr = " ERR収容数 + ERR背番号 "
    If eKanban_Proc = 1 Then
       ErrStr = ErrStr & " + ERR受注 "
    End If
    KariStr = " 品番仮 + 品取仮 + 取引先仮 + 受入仮 "
    
    gCnt1 = eKanban_CvtErr(eKanban_Proc) + eKanban_CvtDbl(eKanban_Proc) + _
            eKanban_CvtKari(eKanban_Proc) + eKanban_CvtEtc(eKanban_Proc)
    
    If gCnt1 <> 4 And gCnt1 <> 0 Then
        If eKanban_CvtErr(eKanban_Proc) = 0 Then
           gStr = " AND ( ( " & ErrStr & " ) <> 0 "
        End If
        If eKanban_CvtDbl(eKanban_Proc) = 0 Then
           If gStr = "" Then
              gStr = " AND (  ERR重複  <> 0 "
           Else
              gStr = gStr & " Or  ERR重複 <> 0 "
           End If
        End If
        If eKanban_CvtKari(eKanban_Proc) = 0 Then
           If gStr = "" Then
              gStr = " AND ( ( " & KariStr & " ) <> 0 "
           Else
              gStr = gStr & " Or ( " & KariStr & " )  <> 0 "
           End If
        End If
        If eKanban_CvtEtc(eKanban_Proc) = 0 Then
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
Public Sub eKanban_LogPut(LogMsg As String, Optional PutType As String = "", Optional SpCount As Long = 10)
    
    Call eKanban_LogProc(CisFun, LogMsg, PutType, SpCount)
    
    If eKanban_FunFlg Then
       Call eKanban_LogProc(eKanban_Fun, LogMsg, PutType, SpCount)
    End If
    
End Sub
Public Sub eKanban_LogProc(LogFun As CISFunction, LogMsg As String, Optional PutType As String = "", Optional SpCount As Long = 10)
    Dim wMsg        As String
    
    
    wMsg = Space(10) & "【 新eかんばん" & eKanban_PrefixNm(eKanban_Proc) & "データ変換処理"
    If eKanban_Rireki = 0 Then
        If eKanban_ProcKB = 0 Then
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

Attribute VB_Name = "DBSaveRst"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   データベース 退避・復元 処理
'*
'*   モジュールID  : DBSaveRst.BAS
'*
'*
'*      作成日  : 2004/03/11   By CIS
'*      変更日  : 2004/09/21   By CIS - [DataBase]-BackDel未登録時は登録
'*      変更日  : 2005.04.01   By CIS データ退避時にタイムアウトを1800に
'*---------------------------------------------------------------------------------
Public db_Proc      As Integer
Public db_ProcType  As String
Public db_BackDir   As String
Public db_LogDelDay As Long
Public db_ExecMsg   As String
Public db_FileCtl   As New CISFunction
Public db_RstDBName As String

Public db_LogFNo    As Integer
Public db_LogFName  As String
Public db_LogData   As String
Public db_BackPath  As String
Public db_SaveFile  As String
Public db_FileCopy  As Boolean

Dim ErStr           As String
Dim ExeCom          As String
Dim ExeStr()        As String

Dim DateStart       As Date
Dim DateEnd         As Date

Declare Function ShellExecute Lib "SHELL32" Alias "ShellExecuteA" (ByVal hWnd&, ByVal lpOperation$, ByVal lpFile$, ByVal lpParameters$, ByVal lpDirectory$, ByVal nShowCmd&) As Long

Private Sub Main()
    
'   #-------------------#
'   # 初 期 内 容 取 得  #
'   #-------------------#
    If Not IniGet Then End
    
    ExeCom = Command
    
    ExeStr = Split(ExeCom, ",")
    
    Call db_Init
    
    db_ProcType = ""
    
    If UBound(ExeStr) < 0 Then
       db_Proc = 0
    Else
       If UBound(ExeStr) >= 0 Then
          If UCase(ExeStr(0)) = "BACK" Then
             db_Proc = 0
             If UBound(ExeStr) > 0 Then
                If UCase(ExeStr(1)) = "AUTO" Then
                   db_ProcType = "A"
                End If
             End If
          End If
          If UCase(ExeStr(0)) = "RST" Then
             db_Proc = 1
          End If
       End If
    End If
    
    If db_ProcType = "" Then
       CDB0010.Show
    Else
       Call db_BackUp
    End If

End Sub
'*-----------------------------*
'*  退避・復元処理 初期処理    *
'*-----------------------------*
Public Sub db_Init()
' データベース情報
    With CisFun
        .INI_Section = "DataBase"
        .INI_FileID = gIniExe & gIniName
        '
        .INI_Default = gIniExe & "DBBack\"
        .INI_Key = "BackDir"
        .GetIni
        db_BackDir = RTrim(.INI_String)
        '
'*--------------------------------------------------<< 2004/09/21 >> Update Start
'       .INI_Key = "BackDel"
'       .INI_Default = "10"
'       .GetIni
'       db_LogDelDay = CisFun.Val2(.INI_String)
'       >> 未登録時は登録（Default値=10）
        .INI_Key = "BackDel"
        .INI_Default = "Error"
        .GetIni
        If .INI_Default = .INI_String Then
            .INI_String = "10"
            .PutIni
        End If
        db_LogDelDay = CisFun.Val2(.INI_String)
'*--------------------------------------------------<< 2004/09/21 >> Update End
        '
        .INI_Key = "BackDeviceDir"
        .INI_Default = ""
        .GetIni
        db_BackPath = RTrim(.INI_String)
        If Trim(db_BackPath) <> "" Then
           If Right(db_BackPath, 1) <> "\" Then db_BackPath = db_BackPath & "\"
        End If
    End With
'+---------------------+
'+  ﾌｧｲﾙ処理初期設定   +
'+---------------------+
    With db_FileCtl
        ' 変換ログ出力ファイル設定
        .Log_PutFlg = True
        .Log_FileInitial = "DBBack"
        .Log_FilePath = gIniExe & "Log\"
        .Log_DataHD = ""
        .Log_DataHDTime = あり
        .Log_DelDay = db_LogDelDay
        
        .FE_Path = db_BackDir
        If Not .DirCheck(NoMsg_Crt) Then
            db_BackDir = gIniExe
        End If
    End With
    
'+---------------------+
'+  処理履歴ﾌｧｲﾙ名ｾｯﾄ  +
'+---------------------+
    db_LogFName = "DBBackLog.txt"
End Sub
'*-----------------------------*
'*  データベース退避           *
'*-----------------------------*
Public Function db_BackUp() As Boolean
    
    Dim ErFlg           As Boolean
    Dim wPath           As String
    Dim wFName          As String
    
    db_BackUp = False
    ErStr = ""
    
    On Error GoTo db_BackUp_ER:
    If db_ProcType <> "A" Then
       db_ExecMsg = "データベース退避処理中 しばらくお待ち下さい..."
       CDB0020.Show
       DoEvents
       For gLong = 1 To 3000000
       Next gLong
    End If
    
    DateStart = Now()
    
' データベース退避
    Call db_FileCtl.LogPut("データベース　退避処理【開始】")
    gLong = 0
    Do
        With CisDB
            gLong = gLong + 1
            If gLong > 5 Then
                ErStr = "退避エラー 規定数オーバー"
                GoTo db_BackUp_ER
            End If
            .ErrDisp = False
            .SQL = "backup database " & .DBName & " to Save" & .DBName & " with init"

'----------------------------< 2005.04.01 ADD START >
            .TimeOut = 1800
'----------------------------< 2005.04.01 ADD  END  >

            .DBExec
            If .ErrCode = 0 Then Exit Do
            If .ErrCode = 3206 Then
                If Not db_crtDevice("Save" & .DBName) Then GoTo db_BackUp_Ed
            Else
                ErStr = .ErrCode & ":" & .ErrMsg
                GoTo db_BackUp_ER
            End If
        End With
        Call db_FileCtl.LogPut("データベース　退避処理【終了】")
    Loop
    
    DateEnd = Now()
    
    If db_SaveDBCopyDel Then
       gLong = DateDiff("s", DateStart, DateEnd)
       db_LogData = Format(Date, "YYYYMMDD")        ' 日付
       db_LogData = db_LogData & "," & DateStart    ' 開始日時
       db_LogData = db_LogData & "," & DateEnd      ' 終了日時
       db_LogData = db_LogData & "," & gLong        ' 処理時間
       db_LogData = db_LogData & "," & db_ProcType  ' 処理ﾀｲﾌﾟ
       db_LogData = db_LogData & "," & db_SaveFile  ' 退避ﾌｧｲﾙ名(ﾌﾙﾊﾟｽ)
       db_LogData = db_LogData & ","                ' 復元日時

       Call db_Rireki("Append")
       Call db_Rireki("Put", db_LogData)
       Call db_Rireki("Close")

       db_BackUp = True

    End If
    
    On Error GoTo 0
    
    If db_ProcType <> "A" Then
       CDB0020!ExecMsg = "データベース退避処理終了"
       DoEvents
       For gLong = 1 To 50000000
       Next gLong
       
       Unload CDB0020
       Set CDB0020 = Nothing
    End If
    
    Exit Function
db_BackUp_ER:
    
    gStr = "Code = " & Err & ",Msg=" & Error
    If ErStr = "" Then
       ErStr = gStr
    End If
    Call db_FileCtl.LogPut("データベース退避　処理異常終了")
    Call db_FileCtl.LogPut(ErStr)
    Call db_MsgDsip("BK", ErStr)
db_BackUp_Ed:
    If Not CDB0020 Is Nothing Then
       Unload CDB0020
       Set CDB0020 = Nothing
    End If

End Function
'+-----------------------------------------------------+
'+      DB処理エラーメッセージ表示
'+-----------------------------------------------------+
Public Sub db_MsgDsip(MsgType As String, ErMsg As String)
    
    If db_ProcType = "A" Then Exit Sub

    With CisFun
         .MB_Lines = 5
        If db_Proc = 0 Then
           .MB_MSG(1) = "         データベース退避　異常終了          "
           .MB_Title = "退避異常"
        Else
           .MB_MSG(1) = "         データベース復元　異常終了          "
           .MB_Title = "復元異常"
        End If
        .MB_MSG(4) = "   " & ErMsg
        .MB_Button = Error
        .MBOX
    End With
End Sub
'+-----------------------------------------------------+
'+      ﾊﾞｯｸｱｯﾌﾟﾃﾞﾊﾞｲｽ作成
'+-----------------------------------------------------+
Public Function db_crtDevice(DevName As String, Optional GetBackPath As Boolean = False) As Boolean
    Dim wBackPath       As String
    
    On Error GoTo db_crtDevice_er:
    db_crtDevice = False
    
'
    If Trim(db_BackPath) = "" Then
        Call db_FileCtl.LogPut("退避デバイスパス取得")
        
        gStr = ""
        With CisDB
             CisDB.SQL = "SELECT filename FROM master.dbo.sysdatabases Where name = '" & .DBName & "'"
             If CisDB.DBRead(1) Then
                If Not IsNull(.Recordset(0)) Then gStr = RTrim(.Recordset(0))
             End If
             On Error Resume Next
             .Recordset.Close
             On Error GoTo db_crtDevice_er:
        End With
        If Trim(gStr) = "" Then
           Call db_FileCtl.LogPut("データベース退避　処理異常終了   (ﾃﾞｰﾀｰﾍﾞｰｽ格納先取得ｴﾗｰ)")
           Call db_MsgDsip("BK", "ﾃﾞｰﾀｰﾍﾞｰｽ格納先取得ｴﾗｰ")
           Exit Function
        End If
        With db_FileCtl
             .FE_FullPath = gStr
             .GetPathName
             gStr = .FE_Path
             If Right(gStr, 1) <> "\" Then gStr = gStr & "\"
             wBackPath = gStr & "Back\"
             .FE_Path = wBackPath
             If Not .DirCheck(NoMsg_Crt) Then
                 db_BackPath = gStr
             Else
                 db_BackPath = wBackPath
             End If
        End With
        If Right(db_BackPath, 1) <> "\" Then
           db_BackPath = db_BackPath & "\"
        End If
        Call db_FileCtl.LogPut("退避デバイスパス=" & db_BackPath)
        With CisFun
             .INI_Section = "DataBase"
             .INI_FileID = gIniExe & gIniName
             .INI_Key = "BackDeviceDir"
             .INI_String = db_BackPath
             .PutIni
        End With
    End If
    
    If GetBackPath Then
       db_crtDevice = True
       Exit Function
    End If
    
    If Trim(DevName) = "" Then GoTo db_crtDevice_er
    
    
    Call db_FileCtl.LogPut("退避デバイス作成処理【開始】")
    
    Call db_FileCtl.LogPut("退避デバイス名　　 = " & DevName)
    Call db_FileCtl.LogPut("退避デバイス格納先 = " & db_BackPath)
    With CisDB
         .ErrDisp = False
         .SQL = "sp_addumpdevice "
         .SQL = .SQL & " 'disk','" & DevName & "','" & db_BackPath & DevName & ".BAK'"
         If Not .DBExec Then
            Call db_FileCtl.LogPut("退避デバイス作成処理　処理異常終了" & .ErrCode & ":" & .ErrMsg)
            Call db_MsgDsip("BK", "退避デバイス作成ｴﾗｰ" & vbCr & .ErrCode & ":" & .ErrMsg)
            Exit Function
         End If
    End With
    Call db_FileCtl.LogPut("退避デバイス作成処理【終了】")
    
    db_crtDevice = True
    Exit Function
db_crtDevice_er:
    Call db_FileCtl.LogPut("退避デバイス作成 異常処理")
    Call db_MsgDsip("BK", "退避デバイス作成　エラー" & vbCr & Err & ":" & Error)
End Function
'+-----------------------------------------------------+
'+      DB退避ﾌｧｲﾙ複写 & 不要ﾃﾞｰﾀ削除
'+-----------------------------------------------------+
Public Function db_SaveDBCopyDel() As Boolean
    Dim wDelDate    As String
    Dim wDelDate2   As String
    Dim wStrFName   As String
    Dim wStrFrom    As String
    Dim wStrTo      As String
    
    db_SaveDBCopyDel = False
    
    If db_BackDir = "" Then
       With CisFun
            .INI_Section = "DataBase"
            .INI_FileID = gIniExe & gIniName
            .INI_Default = ""
            .INI_Key = "BackDeviceDir"
            .GetIni
            db_BackDir = RTrim(.INI_String)
       End With
       If Trim(db_BackDir) = "" Then Exit Function
    End If
    
    If Trim(db_BackPath) = "" Then
        Call db_crtDevice("", True)
        With CisFun
            .INI_Section = "DataBase"
            .INI_FileID = gIniExe & gIniName
            '
            .INI_Key = "BackDeviceDir"
            .INI_Default = ""
            .GetIni
            db_BackPath = RTrim(.INI_String)
            If Trim(db_BackPath) <> "" Then
               If Right(db_BackPath, 1) <> "\" Then db_BackPath = db_BackPath & "\"
            End If
        End With
    End If
    ' 退避ﾃﾞｰﾀ削除
    wDelDate = Format(Now(), "yyyymmdd")
    With CisFun
        .Date_RsFormat = fYMD
        .Date_Add wDelDate, D, db_LogDelDay * -1
        wDelDate2 = .Date_Result
    End With
    wStrFName = Dir(db_BackDir & CisDB.DBName & "*.Bak")      ' 最初のファイル名を返します。
                
    Do While wStrFName <> ""                ' ループを開始します。
       If Mid(wStrFName, Len(CisDB.DBName) + 2, 8) < wDelDate2 Then
          Call db_FileCtl.LogPut("データベース退避ファイル　削除 ファイル名＝" & db_BackDir & wStrFName)
          Kill db_BackDir & wStrFName
       End If

       wStrFName = Dir                     ' 次のファイル名を返します。
    Loop
    ' 退避ﾃﾞｰﾀ複写
    ' パス存在チェック
    With db_FileCtl
        .FE_Path = db_BackDir
        .DirCheck NoMsg_Crt
    End With

    db_SaveFile = db_BackDir & CisDB.DBName & "." & Format(Now(), "yyyymmddhhnnss") & ".Bak"
    
    Call db_FileCtl.LogPut("退避データベース　複写処理【開始】複写名＝" & db_SaveFile)
    If UCase(CisDB.Server) = UCase(CisFun.WSID) Then
       gStr = db_BackPath & "Save" & CisDB.DBName & ".BAK"
    Else
       gStr = "\\" & CisDB.Server & "\" & Replace(db_BackPath, ":", "$") & "Save" & CisDB.DBName & ".BAK"
    End If
    ' XP 対応
    With CisFun
         .INI_Section = "DataBase"
         .INI_FileID = gIniExe & gIniName
         .INI_Default = ""
         .INI_Key = "DeviceRepFrom"
         .GetIni
         wStrFrom = RTrim(.INI_String)
         .INI_Key = "DeviceRepTo"
         .GetIni
         wStrTo = RTrim(.INI_String)
         If Trim(wStrFrom) = "" Then
            .INI_Key = "DeviceRepFrom"
            .INI_Default = ""
            .PutIni
            .INI_Key = "DeviceRepTo"
            .INI_Default = ""
            .PutIni
         Else
            gStr = Replace(gStr, wStrFrom, wStrTo)
         End If
    End With
    
    Call db_FileCtl.LogPut("   Copy＝" & gStr & "," & db_SaveFile)
    With db_FileCtl
         If CisDB.Server = CisFun.WSID Then
            .FE_FullPath = db_BackPath & "Save" & CisDB.DBName & ".BAK"
         Else
            .FE_FullPath = gStr         ' 複写元ﾌｧｲﾙ（ﾌﾙﾊﾟｽ)
         End If
         .FE_CpySak = db_SaveFile       ' 複写先ﾌｧｲﾙ（ﾌﾙﾊﾟｽ)
         If Not .FCpy Then
            Call db_FileCtl.LogPut("退避データベース　複写処理エラー")
            Call db_MsgDsip("BK", "退避データベース 複写エラー")
            Exit Function
         End If
    End With

    db_SaveDBCopyDel = True
    
End Function
'+-----------------------------------------------------+
'+      DB処理履歴ﾌｧｲﾙ操作
'+-----------------------------------------------------+
Public Function db_Rireki(ProcType As String, Optional LogData As String = "") As Boolean
    db_Rireki = False
    On Error GoTo db_Rireki_er:
    
    If UCase(ProcType) = "APPEND" Then
       db_LogFNo = FreeFile
       Open db_BackDir & db_LogFName For Append As #db_LogFNo
    End If
    If UCase(ProcType) = "OUTPUT" Then
       db_LogFNo = FreeFile
       Open db_BackDir & db_LogFName For Output As #db_LogFNo
    End If
    If UCase(ProcType) = "CLOSE" Then
       Close db_LogFNo
       db_LogFNo = 0
    End If
    If UCase(ProcType) = "PUT" Then
       Print #db_LogFNo, LogData
    End If
    If UCase(ProcType) = "INPUT" Then
       db_LogFNo = FreeFile
       Open db_BackDir & db_LogFName For Input As #db_LogFNo
    End If
    If UCase(ProcType) = "READ" Then
       If EOF(db_LogFNo) Then Exit Function
       Line Input #db_LogFNo, LogData
    End If
    
    db_Rireki = True
    On Error GoTo 0
    Exit Function
db_Rireki_er:
    
End Function
'*-----------------------------*
'*  データベース復元           *
'*-----------------------------*
Public Function db_ReStore() As Boolean
    Dim wStrFrom    As String
    Dim wStrTo      As String
    
    db_ReStore = False
    ErStr = ""
    
    On Error GoTo db_ReStore_ER:
    
    If db_ProcType <> "A" Then
       db_ExecMsg = "データベース復元処理中 しばらくお待ち下さい..."
       CDB0020.Show
       DoEvents
       For gLong = 1 To 3000000
       Next gLong
    End If

    ' 復元ﾃﾞｰﾀ複写
    ' パス存在チェック
    With db_FileCtl
        .FE_Path = db_BackPath
        .DirCheck NoMsg_Crt
    End With

    If db_FileCopy Then
        Call db_FileCtl.LogPut("復元データベース　複写処理【開始】複写名＝" & db_SaveFile)
        If CisDB.Server = CisFun.WSID Then
           gStr = db_BackPath & "Save" & CisDB.DBName & ".BAK"
        Else
           gStr = "\\" & CisDB.Server & "\" & Replace(db_BackPath, ":", "$") & "Save" & db_RstDBName & ".BAK"
        End If
        ' XP 対応
        With CisFun
             .INI_Section = "DataBase"
             .INI_FileID = gIniExe & gIniName
             .INI_Default = ""
             .INI_Key = "DeviceRepFrom"
             .GetIni
             wStrFrom = RTrim(.INI_String)
             .INI_Key = "DeviceRepTo"
             .GetIni
             wStrTo = RTrim(.INI_String)
             If Trim(wStrFrom) = "" Then
                .INI_Key = "DeviceRepFrom"
                .INI_Default = ""
                .INI_String = ""
                .PutIni
                .INI_Key = "DeviceRepTo"
                .INI_Default = ""
                .PutIni
             Else
                gStr = Replace(gStr, wStrFrom, wStrTo)
             End If
        End With
        
        Call db_FileCtl.LogPut("   Copy＝" & gStr & "," & db_SaveFile)
        With db_FileCtl
             If UCase(CisDB.Server) = UCase(CisFun.WSID) Then
                .FE_CpySak = db_BackPath & "Save" & db_RstDBName & ".BAK"
             Else
                .FE_CpySak = gStr              ' 複写先ﾌｧｲﾙ（ﾌﾙﾊﾟｽ)
             End If
             .FE_FullPath = db_SaveFile        ' 複写元ﾌｧｲﾙ（ﾌﾙﾊﾟｽ)
             If Not .FCpy Then
                Call db_FileCtl.LogPut("復元データベース　複写処理エラー")
                Call db_MsgDsip("BK", "復元データベース 複写エラー")
                Exit Function
             End If
        End With
    End If
' データベース復元
    Call db_FileCtl.LogPut("データベース　復元処理【開始】")
    With CisDB
        .ErrDisp = False
        .SQL = "RESTORE DATABASE " & db_RstDBName
        .SQL = .SQL & " FROM " & "Save" & db_RstDBName

'----------------------------< 2005.04.01 ADD START >
            .TimeOut = 1800
'----------------------------< 2005.04.01 ADD  END  >

        .DBExec
        If .ErrCode <> 0 Then
            ErStr = .ErrCode & ":" & .ErrMsg
            GoTo db_ReStore_ER
        End If
    End With
    
    Call db_FileCtl.LogPut("データベース　復元処理【終了】")

    If db_ProcType <> "A" Then
       CDB0020!ExecMsg = "データベース復元処理終了"
       DoEvents
       For gLong = 1 To 50000000
       Next gLong
       
       Unload CDB0020
       Set CDB0020 = Nothing
    End If
        
    db_ReStore = True

    Exit Function
db_ReStore_ER:
    gStr = "Code = " & Err & ",Msg=" & Error
    If ErStr = "" Then
       ErStr = gStr
    End If
    Call db_FileCtl.LogPut("データベース復元　処理異常終了")
    Call db_FileCtl.LogPut(ErStr)
    Call db_MsgDsip("BK", ErStr)
db_ReStore_Ed:
    If Not CDB0020 Is Nothing Then
       Unload CDB0020
       Set CDB0020 = Nothing
    End If
End Function


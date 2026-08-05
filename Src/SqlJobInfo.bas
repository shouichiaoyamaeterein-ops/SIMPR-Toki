Attribute VB_Name = "SqlJobInfo"
Option Explicit
'******************************************************************************************
'*
'*          SQL-Server ﾊﾞｯｸｱｯﾌﾟｽｹｼﾞｭｰﾙ情報取得
'*
'*              ( 内容 )    SQL-Serverのﾊﾞｯｸｱｯﾌﾟｽｹｼﾞｭｰﾙ情報を取得し
'*                            変数へ格納する。
'*
'******************************************************************************************

Type sql_Job_Type
    Id                As String
    JobName           As String
    DataBase          As String
    BackDeve          As String
    BackPath          As String
    EndProcYMD        As String
    EndProcTime       As String
    freq_type         As Long
    freq_interval     As Long
    active_start_time As String
    実行              As String
    間隔              As String
    
    next_run_date     As Long
    NextProcYMD       As String
    NextProcTime      As String
End Type

Public sql_Job()        As sql_Job_Type
Dim sql_Job_WK()        As sql_Job_Type

Public sql_job_max      As Long
Public sql_AllDBFlg     As Boolean
Dim CommandStr          As String
Dim CommandStrWK        As String
Dim CommandStrWK2       As String
Dim DBName              As String

Dim lCnt1               As Long
Dim lCnt2               As Long

Public Function sql_job_info_Get(Optional maxCnt As Long = 7) As Boolean
    Dim wStr        As String
    Dim wWeek       As String
    
    sql_job_max = maxCnt
    
    ReDim sql_Job(sql_job_max)
    
    With CisDB
        ' JOB基本情報を求める
        .SQL = "Select job_id,name From msdb.dbo.sysjobs"
        .SQL = .SQL & " Where enabled = 1 Order by date_created"
        If Not .DBRead Then
           sql_job_info_Get = True
           Exit Function
        End If
        lCnt1 = 0
        ReDim sql_Job_WK(.RecordCount)
        Do Until .Recordset.EOF
           If Not IsNull(.Recordset.Fields.Item(0).Value) Then sql_Job_WK(lCnt1).Id = .Recordset.Fields.Item(0).Value
           If Not IsNull(.Recordset.Fields.Item(1).Value) Then sql_Job_WK(lCnt1).JobName = .Recordset.Fields.Item(1).Value
        
           .Recordset.MoveNext
           lCnt1 = lCnt1 + 1
        Loop
        .Recordset.Close
        
        ' ﾊﾞｯｸｱｯﾌﾟ情報を求める
        For lCnt1 = 0 To UBound(sql_Job_WK) - 1
            If Trim(sql_Job_WK(lCnt1).Id) = "" Then Exit For
            .SQL = "Select subsystem,command,last_run_date,last_run_time"
            .SQL = .SQL & " From msdb.dbo.sysjobsteps "
            .SQL = .SQL & " Where job_id = '" & sql_Job_WK(lCnt1).Id & "'"
            .SQL = .SQL & "   And subsystem = 'TSQL'            "
            .SQL = .SQL & " Order by step_id"
            If .DBRead Then
                Do Until .Recordset.EOF
                   gStr = ""
                   If Not IsNull(.Recordset.Fields.Item(0).Value) Then gStr = .Recordset.Fields.Item(0).Value
                   CommandStr = ""
                   If Not IsNull(.Recordset.Fields.Item(1).Value) Then CommandStr = .Recordset.Fields.Item(1).Value
                   Call job_info_edit
                   
                   If Not IsNull(.Recordset.Fields.Item(2).Value) Then sql_Job_WK(lCnt1).EndProcYMD = .Recordset.Fields.Item(2).Value
                   If Not IsNull(.Recordset.Fields.Item(3).Value) Then sql_Job_WK(lCnt1).EndProcTime = .Recordset.Fields.Item(3).Value
            
                   .Recordset.MoveNext
                Loop
            End If
            .Recordset.Close
        Next lCnt1
        ' ﾃﾞﾊﾞｲｽ情報を求める
        For lCnt1 = 0 To UBound(sql_Job_WK) - 1
            If Trim(sql_Job_WK(lCnt1).BackDeve) <> "" Then
               With CisDB
                    gStr = ""
                    .SQL = "SELECT phyname FROM master.dbo.sysdevices Where name = '" & Trim(sql_Job_WK(lCnt1).BackDeve) & "'"
                    If .DBRead(1) Then
                       If Not IsNull(.Recordset.Fields.Item(0).Value) Then sql_Job_WK(lCnt1).BackPath = .Recordset.Fields.Item(0).Value
                    End If
                    .Recordset.Close
                End With
            End If
        Next lCnt1
        ' ｽｹｼﾞｭｰﾙ情報を求める
        For lCnt1 = 0 To UBound(sql_Job_WK) - 1
            If Trim(sql_Job_WK(lCnt1).Id) = "" Then Exit For
            .SQL = "Select freq_type,freq_interval,active_start_time,next_run_date"
            .SQL = .SQL & " From msdb.dbo.sysjobschedules "
            .SQL = .SQL & " Where job_id = '" & sql_Job_WK(lCnt1).Id & "'"
            .SQL = .SQL & "   And enabled = 1"
            If .DBRead Then
                Do Until .Recordset.EOF
                   gStr = ""
                   If Not IsNull(.Recordset.Fields.Item(0).Value) Then sql_Job_WK(lCnt1).freq_type = .Recordset.Fields.Item(0).Value
                   If Not IsNull(.Recordset.Fields.Item(1).Value) Then sql_Job_WK(lCnt1).freq_interval = .Recordset.Fields.Item(1).Value
                   If Not IsNull(.Recordset.Fields.Item(2).Value) Then sql_Job_WK(lCnt1).active_start_time = Mid(Format(.Recordset.Fields.Item(2).Value, "000000"), 1, 4)
                   If Not IsNull(.Recordset.Fields.Item(3).Value) Then sql_Job_WK(lCnt1).next_run_date = .Recordset.Fields.Item(3).Value
            
                   
                   Select Case sql_Job_WK(lCnt1).freq_type
                          Case 1: sql_Job_WK(lCnt1).実行 = "1回"
                          Case 4: sql_Job_WK(lCnt1).実行 = "毎日"
                          Case 8: sql_Job_WK(lCnt1).実行 = "毎週"
                          Case 16: sql_Job_WK(lCnt1).実行 = "毎月"
                          Case 32: sql_Job_WK(lCnt1).実行 = "月単位"
                          Case 64: sql_Job_WK(lCnt1).実行 = "Agent"
                   End Select
                   
                   If sql_Job_WK(lCnt1).freq_type = 4 Then
                      sql_Job_WK(lCnt1).間隔 = sql_Job_WK(lCnt1).freq_interval & "日毎"
                   End If
                   
                   If sql_Job_WK(lCnt1).freq_type = 8 Then
                      wWeek = ""
                      wStr = Convert10to2(sql_Job_WK(lCnt1).freq_interval)
                      If Mid(wStr, 7, 1) = "1" Then wWeek = wWeek & "･日"
                      If Mid(wStr, 6, 1) = "1" Then wWeek = wWeek & "･月"
                      If Mid(wStr, 5, 1) = "1" Then wWeek = wWeek & "･火"
                      If Mid(wStr, 4, 1) = "1" Then wWeek = wWeek & "･水"
                      If Mid(wStr, 3, 1) = "1" Then wWeek = wWeek & "･木"
                      If Mid(wStr, 2, 1) = "1" Then wWeek = wWeek & "･金"
                      If Mid(wStr, 1, 1) = "1" Then wWeek = wWeek & "･土"
                      If Mid(wWeek, 1, 1) = "･" Then wWeek = Mid(wWeek, 2)
                      If Len(wWeek) = 1 Then wWeek = wWeek & "曜日"
                      sql_Job_WK(lCnt1).間隔 = wWeek
                   End If
                   .Recordset.MoveNext
                Loop
            End If
            .Recordset.Close
        Next lCnt1
        
    End With
    
'    For lCnt1 = 0 To UBound(sql_Job_WK) - 1
'        Debug.Print sql_Job_WK(lCnt1).DataBase & "," & sql_Job_WK(lCnt1).JobName & "," & sql_Job_WK(lCnt1).BackPath & "," & sql_Job_WK(lCnt1).BackDeve & "," & _
'                    sql_Job_WK(lCnt1).EndProcYMD & "," & sql_Job_WK(lCnt1).EndProcTime
'    Next lCnt1
    
'Select job_id,name,* From msdb.dbo.sysjobs Where enabled = 1 Order by date_created

'Select * From msdb.dbo.sysjobschedules    Where job_id = '23653C47-C9B8-492B-8517-F41FBE8AC473'

'--Select * From msdb.dbo.sysjobhistory    Where job_id = '98196479-EED2-44D4-8F72-0B0C19957C4D' order by instance_id desc

'Select * From msdb.dbo.sysjobsteps  Where job_id = '23653C47-C9B8-492B-8517-F41FBE8AC473
'
    lCnt2 = 0
    For lCnt1 = 0 To UBound(sql_Job_WK) - 1
        If sql_AllDBFlg Or Trim(sql_Job_WK(lCnt1).DataBase) = Trim(CisDB.DBName) Then
           sql_Job(lCnt2).Id = sql_Job_WK(lCnt1).Id
           sql_Job(lCnt2).JobName = sql_Job_WK(lCnt1).JobName
           sql_Job(lCnt2).DataBase = sql_Job_WK(lCnt1).DataBase
           sql_Job(lCnt2).BackDeve = sql_Job_WK(lCnt1).BackDeve
           sql_Job(lCnt2).BackPath = sql_Job_WK(lCnt1).BackPath
           sql_Job(lCnt2).EndProcYMD = sql_Job_WK(lCnt1).EndProcYMD
           sql_Job(lCnt2).EndProcTime = sql_Job_WK(lCnt1).EndProcTime
           sql_Job(lCnt2).freq_type = sql_Job_WK(lCnt1).freq_type
           sql_Job(lCnt2).freq_interval = sql_Job_WK(lCnt1).freq_interval
           sql_Job(lCnt2).active_start_time = sql_Job_WK(lCnt1).active_start_time
           sql_Job(lCnt2).実行 = sql_Job_WK(lCnt1).実行
           sql_Job(lCnt2).間隔 = sql_Job_WK(lCnt1).間隔
           sql_Job(lCnt2).next_run_date = sql_Job_WK(lCnt1).next_run_date
           sql_Job(lCnt2).NextProcYMD = sql_Job_WK(lCnt1).NextProcYMD
           sql_Job(lCnt2).NextProcTime = sql_Job_WK(lCnt1).NextProcTime
           
           lCnt2 = lCnt2 + 1
        End If
        If lCnt2 >= sql_job_max Then Exit For
    
    Next lCnt1

End Function
'+-----------------------------------+
'+     JOB情報編集
'+-----------------------------------+
Private Sub job_info_edit()
    If InStr(1, CommandStr, "BACKUP DATABASE") <> 0 Then
       lCnt2 = InStr(1, CommandStr, " TO ")
       CommandStrWK = ""
       CommandStrWK2 = ""
       If lCnt2 <> 0 Then
          CommandStrWK = Mid(CommandStr, 1, lCnt2)
          CommandStrWK2 = Mid(CommandStr, lCnt2 + 4)
       End If
       CommandStrWK = Replace(CommandStrWK, "BACKUP DATABASE", "")
       CommandStrWK = Replace(LTrim(CommandStrWK), "[", "")
       lCnt2 = InStr(1, CommandStrWK, "]")
       If lCnt2 <> 0 Then
          sql_Job_WK(lCnt1).DataBase = Mid(CommandStrWK, 1, lCnt2 - 1)
          CommandStrWK = Replace(CommandStrWK, sql_Job_WK(lCnt1).DataBase & "] TO", "")
       Else
          sql_Job_WK(lCnt1).DataBase = Trim(CommandStrWK)
       End If
       lCnt2 = InStr(1, CommandStrWK2, "[")
       If lCnt2 <> 0 Then
          CommandStrWK2 = Mid(CommandStrWK2, lCnt2 + 1)
          lCnt2 = InStr(1, CommandStrWK2, "]")
          If lCnt2 <> 0 Then
             sql_Job_WK(lCnt1).BackDeve = Mid(CommandStrWK2, 1, lCnt2 - 1)
          End If
       End If
       lCnt2 = InStr(1, CommandStrWK2, "DISK = N'")
       If lCnt2 <> 0 Then
          CommandStrWK2 = Replace(LTrim(CommandStrWK2), "DISK = N'", "")
          lCnt2 = InStr(1, CommandStrWK2, "'")
          If lCnt2 <> 0 Then
             sql_Job_WK(lCnt1).BackPath = Mid(CommandStrWK2, 1, lCnt2 - 1)
          End If
       End If
       
    'BACKUP DATABASE [SIMPRA] TO [SIMPRA] WITH  INIT ,  NOUNLOAD ,  NAME = N'SIMPRA バックアップ',  NOSKIP ,  STATS = 10,  NOFORMAT
    'BACKUP DATABASE [SIMKincDB] TO  DISK = N'E:\F\USERDB\BACK\SIMKincDB.Bak' WITH  INIT ,  NOUNLOAD ,  NAME = N'SIMKincDB バックアップ',  NOSKIP ,  STATS = 10,  NOFORMAT
    End If

End Sub
'■Convert10to2
'■機能：10進数を2進数に変換する。
Private Function Convert10to2(Value As Long, Optional rtLeng As Long = 7) As String

    Dim lngBit As Long
    Dim strData As String

    Do Until (Value < 2 ^ lngBit)
        If (Value And 2 ^ lngBit) <> 0 Then
            strData = "1" & strData
        Else
            strData = "0" & strData
        End If

        lngBit = lngBit + 1
    Loop
    
    If Len(strData) < rtLeng Then
       strData = String(rtLeng - Len(strData), "0") & strData
    End If
    
    Convert10to2 = strData

End Function


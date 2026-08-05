Attribute VB_Name = "Eかんばん変換履歴テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   eかんばん変換履歴テーブル
'*
'*   モジュールID  : Eかんばん変換履歴テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*                                      /項目追加含む/
'*   作成日  :  2008/07/16  By Cis(FMV-W5240U)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_eかんばん変換履歴テーブル
        履歴管理番号          As Long
        受注出荷区分          As Byte
        開始日時              As Date
        終了日時              As Date
        処理日付              As String
        処理回数              As Long
        退避連番              As Long
        入力件数              As Long
        変換件数              As Long
        変換エラー件数        As Long
        空                    As Long
        重複                  As Long
        ERR収容数             As Long
        ERR背番号             As Long
        ERR受注               As Long
        仮品番                As Long
        仮品取                As Long
        仮取引先              As Long
        仮受入                As Long
        処理区分              As Byte
        起動                  As Byte
        入力ファイル数        As Long
        ERR                   As String
        退避先                As String
        退避先UNC             As String
'        作成日                As Date
'        作成端末              As String
        件数          As Long
'   項目追加(自動)
        処理退避先     As String
    End Type

    Public EHR As T_eかんばん変換履歴テーブル

    Public EHR_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromEHR(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearEHR
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With EHR
            Select Case adoCol.Name
                Case "履歴管理番号":          If Not IsNull(adoCol.Value) Then .履歴管理番号 = adoCol.Value
                Case "受注出荷区分":          If Not IsNull(adoCol.Value) Then .受注出荷区分 = adoCol.Value
                Case "開始日時":              If Not IsNull(adoCol.Value) Then .開始日時 = adoCol.Value
                Case "終了日時":              If Not IsNull(adoCol.Value) Then .終了日時 = adoCol.Value
                Case "処理日付":              If Not IsNull(adoCol.Value) Then .処理日付 = adoCol.Value
                Case "処理回数":              If Not IsNull(adoCol.Value) Then .処理回数 = adoCol.Value
                Case "退避連番":              If Not IsNull(adoCol.Value) Then .退避連番 = adoCol.Value
                Case "入力件数":              If Not IsNull(adoCol.Value) Then .入力件数 = adoCol.Value
                Case "変換件数":              If Not IsNull(adoCol.Value) Then .変換件数 = adoCol.Value
                Case "変換エラー件数":        If Not IsNull(adoCol.Value) Then .変換エラー件数 = adoCol.Value
                Case "空":                    If Not IsNull(adoCol.Value) Then .空 = adoCol.Value
                Case "重複":                  If Not IsNull(adoCol.Value) Then .重複 = adoCol.Value
                Case "ERR収容数":             If Not IsNull(adoCol.Value) Then .ERR収容数 = adoCol.Value
                Case "ERR背番号":             If Not IsNull(adoCol.Value) Then .ERR背番号 = adoCol.Value
                Case "ERR受注":               If Not IsNull(adoCol.Value) Then .ERR受注 = adoCol.Value
                Case "仮品番":                If Not IsNull(adoCol.Value) Then .仮品番 = adoCol.Value
                Case "仮品取":                If Not IsNull(adoCol.Value) Then .仮品取 = adoCol.Value
                Case "仮取引先":              If Not IsNull(adoCol.Value) Then .仮取引先 = adoCol.Value
                Case "仮受入":                If Not IsNull(adoCol.Value) Then .仮受入 = adoCol.Value
                Case "処理区分":              If Not IsNull(adoCol.Value) Then .処理区分 = adoCol.Value
                Case "起動":                  If Not IsNull(adoCol.Value) Then .起動 = adoCol.Value
                Case "入力ファイル数":        If Not IsNull(adoCol.Value) Then .入力ファイル数 = adoCol.Value
                Case "ERR":                   If Not IsNull(adoCol.Value) Then .ERR = adoCol.Value
                Case "退避先":                If Not IsNull(adoCol.Value) Then .退避先 = adoCol.Value
                Case "退避先UNC":             If Not IsNull(adoCol.Value) Then .退避先UNC = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "処理退避先":      If Not IsNull(adoCol.Value) Then .処理退避先 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function EHRRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    EHRRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From eかんばん変換履歴テーブル"
        RTxtSql = RTxtSql & " Where 履歴管理番号 = " & EHR.履歴管理番号 & ""
        RTxtSql = RTxtSql & " And   受注出荷区分 = " & EHR.受注出荷区分 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           EHRRead = True
           Call ItemsFromEHR(DBNo)
           EHR_RDSTS = True
        Else
           EHR_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call EHRClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub EHRReadNext(Optional DBNo As Integer = 0)
    If Not EHR_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           EHR_RDSTS = False
        Else
           Call ItemsFromEHR(DBNo)
           EHR_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub EHRClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        EHR_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( EHR )
'*----------------------------------------------*'
Public Sub ItemsClearEHR()
    With EHR
        .履歴管理番号 = 0
        .受注出荷区分 = 0
        .開始日時 = 0
        .終了日時 = 0
        .処理日付 = ""
        .処理回数 = 0
        .退避連番 = 0
        .入力件数 = 0
        .変換件数 = 0
        .変換エラー件数 = 0
        .空 = 0
        .重複 = 0
        .ERR収容数 = 0
        .ERR背番号 = 0
        .ERR受注 = 0
        .仮品番 = 0
        .仮品取 = 0
        .仮取引先 = 0
        .仮受入 = 0
        .処理区分 = 0
        .起動 = 0
        .入力ファイル数 = 0
        .ERR = ""
        .退避先 = ""
        .退避先UNC = ""
'       .作成日 = 0
'       .作成端末 = ""
        .件数 = 0
'   項目追加(自動)
        .処理退避先 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function EHRInsert(Optional DBNo As Integer) As Boolean
    EHRInsert = False
    With EHR
        gSL_Select = "Insert Into eかんばん変換履歴テーブル  ( "
        gSL_Select = gSL_Select & "履歴管理番号,受注出荷区分,開始日時,"
        gSL_Select = gSL_Select & "終了日時,処理日付,処理回数,"
        gSL_Select = gSL_Select & "退避連番,入力件数,変換件数,"
        gSL_Select = gSL_Select & "変換エラー件数,空,重複,"
        gSL_Select = gSL_Select & "ERR収容数,ERR背番号,ERR受注,"
        gSL_Select = gSL_Select & "仮品番,仮品取,仮取引先,"
        gSL_Select = gSL_Select & "仮受入,処理区分,起動,"
        gSL_Select = gSL_Select & "入力ファイル数,ERR,退避先,"
        gSL_Select = gSL_Select & "退避先UNC,"
        gSL_Select = gSL_Select & "作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .履歴管理番号 & ","
        gSL_Select = gSL_Select & "" & .受注出荷区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.開始日時) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.終了日時) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.処理日付) & "',"
        gSL_Select = gSL_Select & "" & .処理回数 & ","
        gSL_Select = gSL_Select & "" & .退避連番 & ","
        gSL_Select = gSL_Select & "" & .入力件数 & ","
        gSL_Select = gSL_Select & "" & .変換件数 & ","
        gSL_Select = gSL_Select & "" & .変換エラー件数 & ","
        gSL_Select = gSL_Select & "" & .空 & ","
        gSL_Select = gSL_Select & "" & .重複 & ","
        gSL_Select = gSL_Select & "" & .ERR収容数 & ","
        gSL_Select = gSL_Select & "" & .ERR背番号 & ","
        gSL_Select = gSL_Select & "" & .ERR受注 & ","
        gSL_Select = gSL_Select & "" & .仮品番 & ","
        gSL_Select = gSL_Select & "" & .仮品取 & ","
        gSL_Select = gSL_Select & "" & .仮取引先 & ","
        gSL_Select = gSL_Select & "" & .仮受入 & ","
        gSL_Select = gSL_Select & "" & .処理区分 & ","
        gSL_Select = gSL_Select & "" & .起動 & ","
        gSL_Select = gSL_Select & "" & .入力ファイル数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.ERR) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.退避先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.退避先UNC) & "',"
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    EHRInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function EHRUpdate(Optional DBNo As Integer) As Boolean
    EHRUpdate = False
    With EHR
        gSL_Select = "Update eかんばん変換履歴テーブル  Set "
        gSL_Select = gSL_Select & "履歴管理番号 = " & .履歴管理番号 & ","
        gSL_Select = gSL_Select & "受注出荷区分 = " & .受注出荷区分 & ","
        gSL_Select = gSL_Select & "開始日時 = '" & RTrim(.開始日時) & "',"
        gSL_Select = gSL_Select & "終了日時 = '" & RTrim(.終了日時) & "',"
        gSL_Select = gSL_Select & "処理日付 = '" & RTrim(.処理日付) & "',"
        gSL_Select = gSL_Select & "処理回数 = " & .処理回数 & ","
        gSL_Select = gSL_Select & "退避連番 = " & .退避連番 & ","
        gSL_Select = gSL_Select & "入力件数 = " & .入力件数 & ","
        gSL_Select = gSL_Select & "変換件数 = " & .変換件数 & ","
        gSL_Select = gSL_Select & "変換エラー件数 = " & .変換エラー件数 & ","
        gSL_Select = gSL_Select & "空 = " & .空 & ","
        gSL_Select = gSL_Select & "重複 = " & .重複 & ","
        gSL_Select = gSL_Select & "ERR収容数 = " & .ERR収容数 & ","
        gSL_Select = gSL_Select & "ERR背番号 = " & .ERR背番号 & ","
        gSL_Select = gSL_Select & "ERR受注 = " & .ERR受注 & ","
        gSL_Select = gSL_Select & "仮品番 = " & .仮品番 & ","
        gSL_Select = gSL_Select & "仮品取 = " & .仮品取 & ","
        gSL_Select = gSL_Select & "仮取引先 = " & .仮取引先 & ","
        gSL_Select = gSL_Select & "仮受入 = " & .仮受入 & ","
        gSL_Select = gSL_Select & "処理区分 = " & .処理区分 & ","
        gSL_Select = gSL_Select & "起動 = " & .起動 & ","
        gSL_Select = gSL_Select & "入力ファイル数 = " & .入力ファイル数 & ","
        gSL_Select = gSL_Select & "ERR = '" & RTrim(.ERR) & "',"
        gSL_Select = gSL_Select & "退避先 = '" & RTrim(.退避先) & "',"
        gSL_Select = gSL_Select & "退避先UNC = '" & RTrim(.退避先UNC) & "'"
        gSL_Select = gSL_Select & " Where 履歴管理番号 = " & .履歴管理番号 & ""
        gSL_Select = gSL_Select & "   And 受注出荷区分 = " & .受注出荷区分 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    EHRUpdate = True
End Function
'*------------------------------------------------------------------*'
'*
'*                   履歴管理No・処理回数を求める
'*
'*------------------------------------------------------------------*'
Public Function Start_EHRInsert(eKanbanProc As Byte, eKanbanProcKB As Byte, eKanbanFSu As Long) As Boolean
    
    Dim EHR_Kaisu   As Long
    Dim EHR_RNo     As Long
    Dim EHR_TRNo    As Long
    
    On Error GoTo Start_EHRInsert_Er:
    Start_EHRInsert = False
'当日内処理回数を求める
    gSL_Select = "SELECT MAX(処理回数) 処理回数 FROM eかんばん変換履歴テーブル"
    gSL_Select = gSL_Select & " WHERE 受注出荷区分 IN (" & eKanbanProc & "," & eKanbanProc + 2 & ")"
    gSL_Select = gSL_Select & "   AND 処理日付 = '" & Format(Now(), "yyyymmdd") & "'"
    With CisDB
        .SQL = gSL_Select
        .ConnectNo = 0
        If .DBRead(1, 0) Then
            If Not IsNull(.Recordset("処理回数")) Then
                EHR_Kaisu = .Recordset("処理回数")
            End If
        End If
        .Recordset.Close
    End With
    
'履歴管理番号
    gSL_Select = "SELECT MAX(履歴管理番号) 履歴管理番号 FROM eかんばん変換履歴テーブル"
    With CisDB
        .SQL = gSL_Select
        .ConnectNo = 0
        If .DBRead(1, 0) Then
            If Not IsNull(.Recordset("履歴管理番号")) Then
                EHR_RNo = .Recordset("履歴管理番号")
            End If
        End If
        .Recordset.Close
    End With
    gSL_Select = "SELECT MAX(履歴管理番号) 履歴管理番号 FROM TOPPS変換履歴テーブル"
    With CisDB
        .SQL = gSL_Select
        .ConnectNo = 0
        If .DBRead(1, 0) Then
            If Not IsNull(.Recordset("履歴管理番号")) Then
                EHR_TRNo = .Recordset("履歴管理番号")
            End If
        End If
        .Recordset.Close
    End With
    If EHR_TRNo > EHR_RNo Then
        EHR_RNo = EHR_TRNo
    End If

    eKanban_ProcCNT = EHR_Kaisu + 1
' 履歴処理開始
    If eKanban_Rireki = 0 Then
        Call ItemsClearEHR
        With EHR
            .履歴管理番号 = EHR_RNo + 1
            .受注出荷区分 = eKanbanProc
            .開始日時 = Format(Now(), "yyyy/mm/dd hh:nn:ss")
            .終了日時 = Format(Now(), "yyyy/mm/dd hh:nn:ss")
            .処理日付 = Format(Now(), "yyyymmdd")
            .処理回数 = eKanban_ProcCNT
            .起動 = eKanbanProcKB
            .入力ファイル数 = eKanbanFSu
            .ERR = "99"
            .処理区分 = 1
        End With
    Else
        gSL_Select = "Select * From eかんばん変換履歴テーブル"
        gSL_Select = gSL_Select & " Where 履歴管理番号 = " & eKanban_Rireki
        
        If Not EHRRead(gSL_Select, 1) Then
           With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "       eかんばん変換履歴テーブル読み込みエラー         "
                .MB_MSG(4) = "       履歴管理番号 = " & eKanban_Rireki
                .MB_Title = "再処理ｴﾗｰ"
                .MB_Button = Error
                Call .MBOX
                GoTo Start_EHRInsert_Er
            End With
        End If
        With EHR
            eKanban_DirMaxCnt = .退避連番
            .履歴管理番号 = EHR_RNo + 1
            .受注出荷区分 = eKanbanProc + 2
            .開始日時 = Format(Now(), "yyyy/mm/dd hh:nn:ss")
            .終了日時 = Format(Now(), "yyyy/mm/dd hh:nn:ss")
            .処理日付 = Format(Now(), "yyyymmdd")
            .処理回数 = eKanban_ProcCNT
            .起動 = eKanbanProcKB
            .ERR = "99"
            .処理区分 = 1
            .入力件数 = 0
            .変換件数 = 0
            .変換エラー件数 = 0
            .空 = 0
            .重複 = 0
            .ERR収容数 = 0
            .ERR背番号 = 0
            .ERR受注 = 0
            .仮品番 = 0
            .仮品取 = 0
            .仮取引先 = 0
            .仮受入 = 0
            .入力ファイル数 = 0
            If InStr(1, .退避先UNC, CisFun.WSID) <> 0 Then
               .処理退避先 = .退避先
            Else
               .処理退避先 = .退避先UNC
            End If
        End With
    
    End If
    
    If Not EHRInsert Then Exit Function
    
    Start_EHRInsert = True
    Exit Function
Start_EHRInsert_Er:

End Function
'*------------------------------------------------------------------*'
'*
'*                      処　理　経　過　更　新
'*
'*------------------------------------------------------------------*'
Public Function Middle_EHRUpDate(Optional ProcCode As String = "", Optional ErrMsg As String) As Boolean
    On Error GoTo Middle_EHRUpDate_Er:
    
    If EHR.履歴管理番号 = 0 Then
       Middle_EHRUpDate = True
       Exit Function
    End If
    
    Middle_EHRUpDate = False
    
    With EHR
        .終了日時 = Format(Now(), "yyyy/mm/dd hh:nn:ss")
        If .ERR < "04" Then
           .処理区分 = 0
        Else
           .処理区分 = 1
        End If
    End With
    
    If Not EHRUpdate Then
       gStr = CisDB.ErrCode & ":" & CisDB.ErrMsg
       Call eKanban_LogPut(gStr)
       Exit Function
    End If
    Middle_EHRUpDate = True
    Exit Function
Middle_EHRUpDate_Er:

End Function
'-------------------------------------------------------+'
'  エラー内容
'-------------------------------------------------------+'
' 01:対象データなし
' 04:重複起動ｴﾗｰ
' 05:環境設定ｴﾗｰ
' 10:eかんばん端末接続ｴﾗｰ
' 11:eかんばん退避先取得ｴﾗｰ
' 12:退避先作成ｴﾗｰ(日付)
' 13:退避先作成ｴﾗｰ(回数)
' 20:eかんばんTEXT移動ｴﾗｰ
' 21:eかんばんﾌｧｲﾙ移動ｴﾗｰ
' 31:eかんばんﾊﾞｯｸｱｯﾌﾟ削除ｴﾗｰ
' 32:eかんばん変換履歴削除ｴﾗｰ
' 33:eかんばん変換削除ｴﾗｰ
' 41:eかんばんﾃｷｽﾄｲﾝﾎﾟｰﾄｴﾗｰ
' 42:eかんばん変換削除ｴﾗｰ
' 43:eかんばんｲﾝﾎﾟｰﾄ削除ｴﾗｰ
' 44:eかんばんｲﾝﾎﾟｰﾄ処理ｴﾗｰ
' 45:eかんばんTEXT変換ｴﾗｰ
' 46:ｲﾝﾎﾟｰﾄ対象ﾌｧｲﾙなし"
' 51:eかんばん更新ｴﾗｰ
' 52:eかんばん受注累積ｴﾗｰ
' 53:eかんばん分納更新ｴﾗｰ
' 61:出荷指示書発行ｴﾗｰ
' 62:箱種一覧発行ｴﾗｰ
' 63:変換リスト発行ｴﾗｰ
' 99:その他エラー
'


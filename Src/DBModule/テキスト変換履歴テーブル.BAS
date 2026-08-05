Attribute VB_Name = "テキスト変換履歴テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   テキスト変換履歴テーブル
'*
'*   モジュールID  : テキスト変換履歴テーブル.Bas
'*
'*   作成日  :  2009/09/15  By Cis(FMVW5220)
'*   作成日  :  2009/09/22  By Cis(FMVW5220)
'*   変更日  :  2023/08/23  By CIS  東プレ:受注(Prefix 4)を追加
'*   変更日  :  2025/12/17  By CIS  リバース:受注(Prefix 5)を追加
'*---------------------------------------------------------------------------------
    Type T_テキスト変換履歴テーブル
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
        処理名称              As String
'        作成日                As Date
'        作成端末              As String
        件数          As Long
'   項目追加(自動)
        処理退避先     As String
    End Type

    Public TRI As T_テキスト変換履歴テーブル

    Public TRI_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTRI(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTRI
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With TRI
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
                Case "処理名称":              If Not IsNull(adoCol.Value) Then .処理名称 = adoCol.Value
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
Public Function TRIRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TRIRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From テキスト変換履歴テーブル"
        RTxtSql = RTxtSql & " Where 履歴管理番号 = " & TRI.履歴管理番号 & ""
        RTxtSql = RTxtSql & " And   受注出荷区分 = " & TRI.受注出荷区分 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TRIRead = True
           Call ItemsFromTRI(DBNo)
           TRI_RDSTS = True
        Else
           TRI_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call TRIClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TRIReadNext(Optional DBNo As Integer = 0)
    If Not TRI_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TRI_RDSTS = False
        Else
           Call ItemsFromTRI(DBNo)
           TRI_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TRIClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TRI_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TRI )
'*----------------------------------------------*'
Public Sub ItemsClearTRI()
    With TRI
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
        .処理名称 = ""
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
Public Function TRIInsert(Optional DBNo As Integer) As Boolean
    TRIInsert = False
    With TRI
        gSL_Select = "Insert Into テキスト変換履歴テーブル  ( "
        gSL_Select = gSL_Select & "履歴管理番号,受注出荷区分,開始日時,"
        gSL_Select = gSL_Select & "終了日時,処理日付,処理回数,"
        gSL_Select = gSL_Select & "退避連番,入力件数,変換件数,"
        gSL_Select = gSL_Select & "変換エラー件数,空,重複,"
        gSL_Select = gSL_Select & "ERR収容数,ERR背番号,ERR受注,"
        gSL_Select = gSL_Select & "仮品番,仮品取,仮取引先,"
        gSL_Select = gSL_Select & "仮受入,処理区分,起動,"
        gSL_Select = gSL_Select & "入力ファイル数,ERR,退避先,"
        gSL_Select = gSL_Select & "退避先UNC,処理名称,"
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
        gSL_Select = gSL_Select & "'" & RTrim(.処理名称) & "',"
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TRIInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function TRIUpdate(Optional DBNo As Integer) As Boolean
    TRIUpdate = False
    With TRI
        gSL_Select = "Update テキスト変換履歴テーブル  Set "
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
        gSL_Select = gSL_Select & "退避先UNC = '" & RTrim(.退避先UNC) & "',"
        gSL_Select = gSL_Select & "処理名称 = '" & RTrim(.処理名称) & "'"
        gSL_Select = gSL_Select & " Where 履歴管理番号 = " & .履歴管理番号 & ""
        gSL_Select = gSL_Select & "   And 受注出荷区分 = " & .受注出荷区分 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TRIUpdate = True
End Function
'*------------------------------------------------------------------*'
'*
'*                   履歴管理No・処理回数を求める
'*
'*------------------------------------------------------------------*'
Public Function Start_TRIInsert(TextProc As Byte, TextProcKB As Byte, TextFSu As Long) As Boolean
    
    Dim TRI_Kaisu   As Long
    Dim TRI_RNo     As Long
    Dim TRI_TRNo    As Long
    
    On Error GoTo Start_TRIInsert_Er:
    Start_TRIInsert = False
'当日内処理回数を求める
    gSL_Select = "SELECT MAX(処理回数) 処理回数 FROM テキスト変換履歴テーブル"
    '''gSL_Select = gSL_Select & " WHERE 受注出荷区分 IN (" & TextProc & "," & TextProc + 2 & ")"
' ------ 2023/08/23 MNT Start ------
'    If TextProc = 0 Or TextProc = 1 Or TextProc = 2 Or TextProc = 3 Then
'       gSL_Select = gSL_Select & " WHERE 受注出荷区分 IN ( 0,1,2,3 )"
' ------ 2025/12/17 UPD Start ------
'    If TextProc = 0 Or TextProc = 1 Or TextProc = 2 Or TextProc = 3 Or TextProc = 4 Then
'       gSL_Select = gSL_Select & " WHERE 受注出荷区分 IN ( 0,1,2,3,4 )"
    If TextProc = 0 Or TextProc = 1 Or TextProc = 2 Or TextProc = 3 Or TextProc = 4 Or TextProc = 5 Then
       gSL_Select = gSL_Select & " WHERE 受注出荷区分 IN ( 0,1,2,3,4,5 )"
' ------ 2025/12/17 UPD End   ------
' ------ 2023/08/23 MNT End   ------
    Else
       gSL_Select = gSL_Select & " WHERE 受注出荷区分 IN (" & TextProc & "," & TextProc + 2 & ")"
    End If
    gSL_Select = gSL_Select & "   AND 処理日付 = '" & Format(Now(), "yyyymmdd") & "'"
    With CisDB
        .SQL = gSL_Select
        .ConnectNo = 0
        If .DBRead(1, 0) Then
            If Not IsNull(.Recordset("処理回数")) Then
                TRI_Kaisu = .Recordset("処理回数")
            End If
        End If
        .Recordset.Close
    End With
    
'履歴管理番号
    gSL_Select = "SELECT MAX(履歴管理番号) 履歴管理番号 FROM テキスト変換履歴テーブル"
    With CisDB
        .SQL = gSL_Select
        .ConnectNo = 0
        If .DBRead(1, 0) Then
            If Not IsNull(.Recordset("履歴管理番号")) Then
                TRI_RNo = .Recordset("履歴管理番号")
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
                TRI_TRNo = .Recordset("履歴管理番号")
            End If
        End If
        .Recordset.Close
    End With
    If TRI_TRNo > TRI_RNo Then
        TRI_RNo = TRI_TRNo
    End If

    Text_ProcCNT = TRI_Kaisu + 1
' 履歴処理開始
    If Text_Rireki = 0 Then
        Call ItemsClearTRI
        With TRI
            .履歴管理番号 = TRI_RNo + 1
' ------ 2023/08/23 MNT Start ------
'            If TextProc <> 0 And TextProc <> 1 And TextProc <> 3 Then
' ------ 2025/12/17 UPD Start ------
'            If TextProc <> 0 And TextProc <> 1 And TextProc <> 3 And TextProc <> 4 Then
            If TextProc <> 0 And TextProc <> 1 And TextProc <> 3 And TextProc <> 4 And TextProc <> 5 Then
' ------ 2025/12/17 UPD End ------
' ------ 2023/08/23 MNT End   ------
                If TextProc = 2 Then
                   .受注出荷区分 = 1
                Else
                   .受注出荷区分 = TextProc
                End If
            Else
               .受注出荷区分 = 0
            End If
            .開始日時 = Format(Now(), "yyyy/mm/dd hh:nn:ss")
            .終了日時 = Format(Now(), "yyyy/mm/dd hh:nn:ss")
            .処理日付 = Format(Now(), "yyyymmdd")
            .処理回数 = Text_ProcCNT
            .起動 = Text_ProcKB
            .入力ファイル数 = TextFSu
            .ERR = "99"
            .処理区分 = 1
            Select Case Text_Proc
                   Case 0: .処理名称 = "愛知機械受注データ変換"
                   Case 1: .処理名称 = "ﾃｸﾉｴｲﾄ受注データ変換"
                   Case 2: .処理名称 = "愛知機械出荷データ変換"
                   Case 3: .処理名称 = "ﾖﾛｽﾞ愛知受注データ変換"
' ------ 2023/08/23 ADD Start ------
' ------ 2025/12/17 UPD Start ------
'                   Case 3: .処理名称 = "東プレ受注データ変換"
                   Case 4: .処理名称 = "東プレ受注データ変換"
' ------ 2023/08/23 ADD End   ------
                   Case 5: .処理名称 = "リバース受注データ変換"
' ------ 2025/12/17 UPD End   ------
                   Case Else: .処理名称 = ""
            End Select
            
        End With
    Else
        gSL_Select = "Select * From テキスト変換履歴テーブル"
        gSL_Select = gSL_Select & " Where 履歴管理番号 = " & Text_Rireki
        
        If Not TRIRead(gSL_Select, 1) Then
           With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "       テキスト変換履歴テーブル読み込みエラー         "
                .MB_MSG(4) = "       履歴管理番号 = " & Text_Rireki
                .MB_Title = "再処理ｴﾗｰ"
                .MB_Button = Error
                Call .MBOX
                GoTo Start_TRIInsert_Er
            End With
        End If
        With TRI
            Text_DirMaxCnt = .退避連番
            .履歴管理番号 = TRI_RNo + 1
' ------ 2023/08/23 MNT Start ------
'            If TextProc = 0 Or TextProc = 1 Or TextProc = 3 Then
' ------ 2025/12/17 UPD Start ------
'            If TextProc = 0 Or TextProc = 1 Or TextProc = 3 Or TextProc = 4 Then
            If TextProc = 0 Or TextProc = 1 Or TextProc = 3 Or TextProc = 4 Or TextProc = 5 Then
' ------ 2025/12/17 UPD End ------
' ------ 2023/08/23 MNT End   ------
               .受注出荷区分 = 2
            Else
               .受注出荷区分 = 3
            End If
            .開始日時 = Format(Now(), "yyyy/mm/dd hh:nn:ss")
            .終了日時 = Format(Now(), "yyyy/mm/dd hh:nn:ss")
            .処理日付 = Format(Now(), "yyyymmdd")
            .処理回数 = Text_ProcCNT
            .起動 = TextProcKB
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
    
    If Not TRIInsert Then Exit Function
    
    Start_TRIInsert = True
    Exit Function
Start_TRIInsert_Er:

End Function
'*------------------------------------------------------------------*'
'*
'*                      処　理　経　過　更　新
'*
'*------------------------------------------------------------------*'
Public Function Middle_TRIUpDate(Optional ProcCode As String = "", Optional ErrMsg As String) As Boolean
    On Error GoTo Middle_TRIUpDate_Er:
    
    If TRI.履歴管理番号 = 0 Then
       Middle_TRIUpDate = True
       Exit Function
    End If
    
    Middle_TRIUpDate = False
    
    With TRI
        .終了日時 = Format(Now(), "yyyy/mm/dd hh:nn:ss")
        If .ERR < "04" Then
           .処理区分 = 0
        Else
           .処理区分 = 1
        End If
    End With
    
    If Not TRIUpdate Then
       gStr = CisDB.ErrCode & ":" & CisDB.ErrMsg
       Call Text_LogPut(gStr)
       Exit Function
    End If
    Middle_TRIUpDate = True
    Exit Function
Middle_TRIUpDate_Er:

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



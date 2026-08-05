Attribute VB_Name = "TEXT変換履歴テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TEXT変換履歴テーブル
'*
'*   モジュールID  : TEXT変換履歴テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*
'*   作成日  :  2003/12/13  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_TEXT変換履歴テーブル
        履歴管理番号  As Long
        FILECD        As String
        HTYPE         As String
        FCODE         As String
        開始日時      As Date
        終了日時      As Date
        件数          As Long
        変換件数      As Long
        HEAD件数      As Integer
        TAIL件数      As Integer
        文言件数      As Integer
        出力件数      As Long
        処理コード    As Integer
        受信ｴﾗｰ       As String
'        作成日        As Date
        作成者        As String
'        作成端末      As String
    End Type

    Public THR As T_TEXT変換履歴テーブル

    Public THR_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTHR(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTHR	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With THR
            Select Case adoCol.Name
                Case "履歴管理番号":  If Not IsNull(adoCol.Value) Then .履歴管理番号 = adoCol.Value
                Case "FILECD":        If Not IsNull(adoCol.Value) Then .FILECD       = adoCol.Value
                Case "HTYPE":         If Not IsNull(adoCol.Value) Then .HTYPE        = adoCol.Value
                Case "FCODE":         If Not IsNull(adoCol.Value) Then .FCODE        = adoCol.Value
                Case "開始日時":      If Not IsNull(adoCol.Value) Then .開始日時     = adoCol.Value
                Case "終了日時":      If Not IsNull(adoCol.Value) Then .終了日時     = adoCol.Value
                Case "件数":          If Not IsNull(adoCol.Value) Then .件数         = adoCol.Value
                Case "変換件数":      If Not IsNull(adoCol.Value) Then .変換件数     = adoCol.Value
                Case "HEAD件数":      If Not IsNull(adoCol.Value) Then .HEAD件数     = adoCol.Value
                Case "TAIL件数":      If Not IsNull(adoCol.Value) Then .TAIL件数     = adoCol.Value
                Case "文言件数":      If Not IsNull(adoCol.Value) Then .文言件数     = adoCol.Value
                Case "出力件数":      If Not IsNull(adoCol.Value) Then .出力件数     = adoCol.Value
                Case "処理コード":    If Not IsNull(adoCol.Value) Then .処理コード   = adoCol.Value
                Case "受信ｴﾗｰ":       If Not IsNull(adoCol.Value) Then .受信ｴﾗｰ      = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function THRRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    THRRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From TEXT変換履歴テーブル"
        RTxtSql = RTxtSql &" Where 履歴管理番号 = " & THR.履歴管理番号 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           THRRead = True
           Call ItemsFromTHR(DBNo)
	   THR_RDSTS = True
        else
	   THR_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call THRClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub THRReadNext(Optional DBNo As Integer = 0)
    If Not THR_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           THR_RDSTS = False
        Else
           Call ItemsFromTHR(DBNo)
           THR_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub THRClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        THR_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( THR )
'*----------------------------------------------*'
Public Sub ItemsClearTHR()
    With THR
	.履歴管理番号 = 0
	.FILECD = ""
	.HTYPE = ""
	.FCODE = ""
	.開始日時 = 0
	.終了日時 = 0
	.件数 = 0
	.変換件数 = 0
	.HEAD件数 = 0
	.TAIL件数 = 0
	.文言件数 = 0
	.出力件数 = 0
	.処理コード = 0
	.受信ｴﾗｰ = ""
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Sub THRInsert(Optional DBNo As Integer)
    With THR
        gSL_Select = "Insert Into TEXT変換履歴テーブル  ( "
        gSL_Select = gSL_Select & "履歴管理番号,FILECD,HTYPE,"
        gSL_Select = gSL_Select & "FCODE,開始日時,終了日時,"
        gSL_Select = gSL_Select & "件数,変換件数,HEAD件数,"
        gSL_Select = gSL_Select & "TAIL件数,文言件数,出力件数,"
        gSL_Select = gSL_Select & "処理コード,受信ｴﾗｰ,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .履歴管理番号 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.FILECD) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.HTYPE) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.FCODE) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.開始日時) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.終了日時) & "',"		
        gSL_Select = gSL_Select & "" & .件数 & ","		
        gSL_Select = gSL_Select & "" & .変換件数 & ","		
        gSL_Select = gSL_Select & "" & .HEAD件数 & ","		
        gSL_Select = gSL_Select & "" & .TAIL件数 & ","		
        gSL_Select = gSL_Select & "" & .文言件数 & ","		
        gSL_Select = gSL_Select & "" & .出力件数 & ","		
        gSL_Select = gSL_Select & "" & .処理コード & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.受信ｴﾗｰ) & "',"		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "',"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        .DBExec
    End With
End Sub
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Sub THRUpdate(Optional DBNo As Integer)
    With THR
        gSL_Select = "Update TEXT変換履歴テーブル  Set "
        gSL_Select = gSL_Select & "履歴管理番号 = " & .履歴管理番号 & ","
        gSL_Select = gSL_Select & "FILECD = '" & RTrim(.FILECD) & "',"
        gSL_Select = gSL_Select & "HTYPE = '" & RTrim(.HTYPE) & "',"
        gSL_Select = gSL_Select & "FCODE = '" & RTrim(.FCODE) & "',"
        gSL_Select = gSL_Select & "開始日時 = '" & RTrim(.開始日時) & "',"
        gSL_Select = gSL_Select & "終了日時 = '" & RTrim(.終了日時) & "',"
        gSL_Select = gSL_Select & "件数 = " & .件数 & ","
        gSL_Select = gSL_Select & "変換件数 = " & .変換件数 & ","
        gSL_Select = gSL_Select & "HEAD件数 = " & .HEAD件数 & ","
        gSL_Select = gSL_Select & "TAIL件数 = " & .TAIL件数 & ","
        gSL_Select = gSL_Select & "文言件数 = " & .文言件数 & ","
        gSL_Select = gSL_Select & "出力件数 = " & .出力件数 & ","
        gSL_Select = gSL_Select & "処理コード = " & .処理コード & ","
        gSL_Select = gSL_Select & "受信ｴﾗｰ = '" & RTrim(.受信ｴﾗｰ) & "'"
        gSL_Select = gSL_Select & " Where 履歴管理番号 = " & .履歴管理番号 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        .DBExec
    End With

End Sub

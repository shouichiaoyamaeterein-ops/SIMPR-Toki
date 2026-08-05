Attribute VB_Name = "得意先スクラップ単価マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   得意先スクラップ単価マスタ
'*
'*   モジュールID  : 得意先スクラップ単価マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2008/05/15  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_得意先スクラップ単価マスタ
        得意先        As String
        履歴NO        As Long
        適用開始日付  As String
        適用終了日付  As String
        材質          As String
        対象年        As String
        対象期        As Byte
        単価          As Currency
'        作成日        As Date
        作成者        As String
'        作成端末      As String
'        更新日        As Date
        更新者        As String
'        更新端末      As String
        件数          As Long
'   項目追加(自動)
	材質名         As String
	前回単価       As Currency
    End Type

    Public TSM As T_得意先スクラップ単価マスタ

    Public TSM_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTSM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTSM	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With TSM
            Select Case adoCol.Name
                Case "得意先":        If Not IsNull(adoCol.Value) Then .得意先       = adoCol.Value
                Case "履歴NO":        If Not IsNull(adoCol.Value) Then .履歴NO       = adoCol.Value
                Case "適用開始日付":  If Not IsNull(adoCol.Value) Then .適用開始日付 = adoCol.Value
                Case "適用終了日付":  If Not IsNull(adoCol.Value) Then .適用終了日付 = adoCol.Value
                Case "材質":          If Not IsNull(adoCol.Value) Then .材質         = adoCol.Value
                Case "対象年":        If Not IsNull(adoCol.Value) Then .対象年       = adoCol.Value
                Case "対象期":        If Not IsNull(adoCol.Value) Then .対象期       = adoCol.Value
                Case "単価":          If Not IsNull(adoCol.Value) Then .単価         = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者       = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "材質名":	If Not IsNull(adoCol.Value) Then .材質名	= adoCol.Value
		Case "前回単価":	If Not IsNull(adoCol.Value) Then .前回単価	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TSMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TSMRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 得意先スクラップ単価マスタ"
        RTxtSql = RTxtSql &" Where 得意先 = '" & RTrim(TSM.得意先) & "'"
        RTxtSql = RTxtSql &" And   履歴NO = " & TSM.履歴NO & ""
        RTxtSql = RTxtSql &" And   適用開始日付 = '" & RTrim(TSM.適用開始日付) & "'"
        RTxtSql = RTxtSql &" And   適用終了日付 = '" & RTrim(TSM.適用終了日付) & "'"
        RTxtSql = RTxtSql &" And   材質 = '" & RTrim(TSM.材質) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TSMRead = True
           Call ItemsFromTSM(DBNo)
	   TSM_RDSTS = True
        else
	   TSM_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call TSMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TSMReadNext(Optional DBNo As Integer = 0)
    If Not TSM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TSM_RDSTS = False
        Else
           Call ItemsFromTSM(DBNo)
           TSM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TSMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TSM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TSM )
'*----------------------------------------------*'
Public Sub ItemsClearTSM()
    With TSM
	.得意先 = ""
	.履歴NO = 0
	.適用開始日付 = ""
	.適用終了日付 = ""
	.材質 = ""
	.対象年 = ""
	.対象期 = 0
	.単価 = 0
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
'	.更新日 = 0
	.更新者 = ""
'	.更新端末 = ""
        .件数 = 0
'   項目追加(自動)
	.材質名 = ""
	.前回単価 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function TSMInsert(Optional DBNo As Integer) As Boolean
    TSMInsert = False
    With TSM
        gSL_Select = "Insert Into 得意先スクラップ単価マスタ  ( "
        gSL_Select = gSL_Select & "得意先,履歴NO,適用開始日付,"
        gSL_Select = gSL_Select & "適用終了日付,材質,対象年,"
        gSL_Select = gSL_Select & "対象期,単価,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.得意先) & "',"		
        gSL_Select = gSL_Select & "" & .履歴NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.適用開始日付) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.適用終了日付) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.対象年) & "',"		
        gSL_Select = gSL_Select & "" & .対象期 & ","		
        gSL_Select = gSL_Select & "" & .単価 & ","		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "',"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TSMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function TSMUpdate(Optional DBNo As Integer) As Boolean
    TSMUpdate = False	
    With TSM
        gSL_Select = "Update 得意先スクラップ単価マスタ  Set "
        gSL_Select = gSL_Select & "得意先 = '" & RTrim(.得意先) & "',"
        gSL_Select = gSL_Select & "履歴NO = " & .履歴NO & ","
        gSL_Select = gSL_Select & "適用開始日付 = '" & RTrim(.適用開始日付) & "',"
        gSL_Select = gSL_Select & "適用終了日付 = '" & RTrim(.適用終了日付) & "',"
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "対象年 = '" & RTrim(.対象年) & "',"
        gSL_Select = gSL_Select & "対象期 = " & .対象期 & ","
        gSL_Select = gSL_Select & "単価 = " & .単価 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 得意先 = '" & RTrim(.得意先) & "'"
        gSL_Select = gSL_Select & "   And 履歴NO = " & .履歴NO & ""
        gSL_Select = gSL_Select & "   And 適用開始日付 = '" & RTrim(.適用開始日付) & "'"
        gSL_Select = gSL_Select & "   And 適用終了日付 = '" & RTrim(.適用終了日付) & "'"
        gSL_Select = gSL_Select & "   And 材質 = '" & RTrim(.材質) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TSMUpdate = True
End Function

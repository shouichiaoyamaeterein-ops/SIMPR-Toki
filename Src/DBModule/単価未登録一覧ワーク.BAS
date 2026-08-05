Attribute VB_Name = "単価未登録一覧ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   単価未登録一覧ワーク
'*
'*   モジュールID  : 単価未登録一覧ワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2005/06/14  By Cis
'*   作成日  :  2008/01/21  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_単価未登録一覧ワーク
        処理端末  As String
        取引先    As String
        契約先    As String
        品番      As String
        日付      As String
        数量      As Long
'        作成日    As Date
        件数          As Long
'   項目追加(自動)
	表示品番       As String
	日付から       As String
	日付まで       As String
	取引先名       As String
    End Type

    Public TMW As T_単価未登録一覧ワーク

    Public TMW_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTMW(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTMW	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With TMW
            Select Case adoCol.Name
                Case "処理端末":  If Not IsNull(adoCol.Value) Then .処理端末 = adoCol.Value
                Case "取引先":    If Not IsNull(adoCol.Value) Then .取引先   = adoCol.Value
                Case "契約先":    If Not IsNull(adoCol.Value) Then .契約先   = adoCol.Value
                Case "品番":      If Not IsNull(adoCol.Value) Then .品番     = adoCol.Value
                Case "日付":      If Not IsNull(adoCol.Value) Then .日付     = adoCol.Value
                Case "数量":      If Not IsNull(adoCol.Value) Then .数量     = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "日付から":	If Not IsNull(adoCol.Value) Then .日付から	= adoCol.Value
		Case "日付まで":	If Not IsNull(adoCol.Value) Then .日付まで	= adoCol.Value
		Case "取引先名":	If Not IsNull(adoCol.Value) Then .取引先名	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TMWRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TMWRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 単価未登録一覧ワーク"
        RTxtSql = RTxtSql &" Where 品番 = '" & RTrim(TMW.品番) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TMWRead = True
           Call ItemsFromTMW(DBNo)
	   TMW_RDSTS = True
        else
	   TMW_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call TMWClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TMWReadNext(Optional DBNo As Integer = 0)
    If Not TMW_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TMW_RDSTS = False
        Else
           Call ItemsFromTMW(DBNo)
           TMW_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TMWClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TMW_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TMW )
'*----------------------------------------------*'
Public Sub ItemsClearTMW()
    With TMW
	.処理端末 = ""
	.取引先 = ""
	.契約先 = ""
	.品番 = ""
	.日付 = ""
	.数量 = 0
'	.作成日 = 0
        .件数 = 0
'   項目追加(自動)
	.表示品番 = ""
	.日付から = ""
	.日付まで = ""
	.取引先名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function TMWInsert(Optional DBNo As Integer) As Boolean
    TMWInsert = False
    With TMW
        gSL_Select = "Insert Into 単価未登録一覧ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,取引先,契約先,"
        gSL_Select = gSL_Select & "品番,日付,数量,"
        gSL_Select = gSL_Select & "作成日"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.日付) & "',"		
        gSL_Select = gSL_Select & "" & .数量 & ","		
        gSL_Select = gSL_Select & "GETDATE()"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TMWInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function TMWUpdate(Optional DBNo As Integer) As Boolean
    TMWUpdate = False	
    With TMW
        gSL_Select = "Update 単価未登録一覧ワーク  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "日付 = '" & RTrim(.日付) & "',"
        gSL_Select = gSL_Select & "数量 = " & .数量 & ""
        gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(.品番) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TMWUpdate = True
End Function

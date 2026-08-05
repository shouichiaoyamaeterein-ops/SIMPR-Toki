Attribute VB_Name = "受注単価マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   受注単価マスタ
'*
'*   モジュールID  : 受注単価マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2005/06/14  By Cis
'*   作成日  :  2008/02/22  By Cis
'*   作成日  :  2008/02/25  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_受注単価マスタ
        品番          As String
        取引先        As String
        履歴NO        As Long
        仮単価区分    As String
        単価          As Currency
        適用開始日付  As String
        適用終了日付  As String
        部品単価      As Currency
        梱包単価      As Currency
'        作成日        As Date
        作成者        As String
'        作成端末      As String
'        更新日        As Date
        更新者        As String
'        更新端末      As String
        件数          As Long
'   項目追加(自動)
	表示品番       As String
	前回履歴NO     As Long
	前回仮単価区分 As String
	前回単価       As Currency
	前回部品単価   As Currency
	前回梱包単価   As Currency
	背番号         As String
	取引先名       As String
    End Type

    Public JTM As T_受注単価マスタ

    Public JTM_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromJTM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearJTM	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With JTM
            Select Case adoCol.Name
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "取引先":        If Not IsNull(adoCol.Value) Then .取引先       = adoCol.Value
                Case "履歴NO":        If Not IsNull(adoCol.Value) Then .履歴NO       = adoCol.Value
                Case "仮単価区分":    If Not IsNull(adoCol.Value) Then .仮単価区分   = adoCol.Value
                Case "単価":          If Not IsNull(adoCol.Value) Then .単価         = adoCol.Value
                Case "適用開始日付":  If Not IsNull(adoCol.Value) Then .適用開始日付 = adoCol.Value
                Case "適用終了日付":  If Not IsNull(adoCol.Value) Then .適用終了日付 = adoCol.Value
                Case "部品単価":      If Not IsNull(adoCol.Value) Then .部品単価     = adoCol.Value
                Case "梱包単価":      If Not IsNull(adoCol.Value) Then .梱包単価     = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者       = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "前回履歴NO":	If Not IsNull(adoCol.Value) Then .前回履歴NO	= adoCol.Value
		Case "前回仮単価区分":	If Not IsNull(adoCol.Value) Then .前回仮単価区分	= adoCol.Value
		Case "前回単価":	If Not IsNull(adoCol.Value) Then .前回単価	= adoCol.Value
		Case "前回部品単価":	If Not IsNull(adoCol.Value) Then .前回部品単価	= adoCol.Value
		Case "前回梱包単価":	If Not IsNull(adoCol.Value) Then .前回梱包単価	= adoCol.Value
		Case "背番号":	If Not IsNull(adoCol.Value) Then .背番号	= adoCol.Value
		Case "取引先名":	If Not IsNull(adoCol.Value) Then .取引先名	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function JTMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    JTMRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 受注単価マスタ"
        RTxtSql = RTxtSql &" Where 品番 = '" & RTrim(JTM.品番) & "'"
        RTxtSql = RTxtSql &" And   取引先 = '" & RTrim(JTM.取引先) & "'"
        RTxtSql = RTxtSql &" And   履歴NO = " & JTM.履歴NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           JTMRead = True
           Call ItemsFromJTM(DBNo)
	   JTM_RDSTS = True
        else
	   JTM_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call JTMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub JTMReadNext(Optional DBNo As Integer = 0)
    If Not JTM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           JTM_RDSTS = False
        Else
           Call ItemsFromJTM(DBNo)
           JTM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub JTMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        JTM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( JTM )
'*----------------------------------------------*'
Public Sub ItemsClearJTM()
    With JTM
	.品番 = ""
	.取引先 = ""
	.履歴NO = 0
	.仮単価区分 = ""
	.単価 = 0
	.適用開始日付 = ""
	.適用終了日付 = ""
	.部品単価 = 0
	.梱包単価 = 0
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
'	.更新日 = 0
	.更新者 = ""
'	.更新端末 = ""
        .件数 = 0
'   項目追加(自動)
	.表示品番 = ""
	.前回履歴NO = 0
	.前回仮単価区分 = ""
	.前回単価 = 0
	.前回部品単価 = 0
	.前回梱包単価 = 0
	.背番号 = ""
	.取引先名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function JTMInsert(Optional DBNo As Integer) As Boolean
    JTMInsert = False
    With JTM
        gSL_Select = "Insert Into 受注単価マスタ  ( "
        gSL_Select = gSL_Select & "品番,取引先,履歴NO,"
        gSL_Select = gSL_Select & "仮単価区分,単価,適用開始日付,"
        gSL_Select = gSL_Select & "適用終了日付,部品単価,梱包単価,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "" & .履歴NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.仮単価区分) & "',"		
        gSL_Select = gSL_Select & "" & .単価 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.適用開始日付) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.適用終了日付) & "',"		
        gSL_Select = gSL_Select & "" & .部品単価 & ","		
        gSL_Select = gSL_Select & "" & .梱包単価 & ","		
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
    JTMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function JTMUpdate(Optional DBNo As Integer) As Boolean
    JTMUpdate = False	
    With JTM
        gSL_Select = "Update 受注単価マスタ  Set "
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "履歴NO = " & .履歴NO & ","
        gSL_Select = gSL_Select & "仮単価区分 = '" & RTrim(.仮単価区分) & "',"
        gSL_Select = gSL_Select & "単価 = " & .単価 & ","
        gSL_Select = gSL_Select & "適用開始日付 = '" & RTrim(.適用開始日付) & "',"
        gSL_Select = gSL_Select & "適用終了日付 = '" & RTrim(.適用終了日付) & "',"
        gSL_Select = gSL_Select & "部品単価 = " & .部品単価 & ","
        gSL_Select = gSL_Select & "梱包単価 = " & .梱包単価 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 履歴NO = " & .履歴NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    JTMUpdate = True
End Function

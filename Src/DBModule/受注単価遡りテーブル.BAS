Attribute VB_Name = "受注単価遡りテーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   受注単価遡りテーブル
'*
'*   モジュールID  : 受注単価遡りテーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2007/06/15  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_受注単価遡りテーブル
        品番          As String
        取引先        As String
        管理番号      As String
        履歴NO        As Long
        仮単価区分    As String
        単価          As Currency
        適用開始日付  As String
        適用終了日付  As String
        部品単価      As Currency
        梱包単価      As Currency
        実績年月      As String
        実績処理日    As String
'        作成日        As Date
        作成者        As String
'        作成端末      As String
'        更新日        As Date
        更新者        As String
'        更新端末      As String
        件数          As Long
    End Type

    Public JTS As T_受注単価遡りテーブル

    Public JTS_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromJTS(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearJTS	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With JTS
            Select Case adoCol.Name
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "取引先":        If Not IsNull(adoCol.Value) Then .取引先       = adoCol.Value
                Case "管理番号":      If Not IsNull(adoCol.Value) Then .管理番号     = adoCol.Value
                Case "履歴NO":        If Not IsNull(adoCol.Value) Then .履歴NO       = adoCol.Value
                Case "仮単価区分":    If Not IsNull(adoCol.Value) Then .仮単価区分   = adoCol.Value
                Case "単価":          If Not IsNull(adoCol.Value) Then .単価         = adoCol.Value
                Case "適用開始日付":  If Not IsNull(adoCol.Value) Then .適用開始日付 = adoCol.Value
                Case "適用終了日付":  If Not IsNull(adoCol.Value) Then .適用終了日付 = adoCol.Value
                Case "部品単価":      If Not IsNull(adoCol.Value) Then .部品単価     = adoCol.Value
                Case "梱包単価":      If Not IsNull(adoCol.Value) Then .梱包単価     = adoCol.Value
                Case "実績年月":      If Not IsNull(adoCol.Value) Then .実績年月     = adoCol.Value
                Case "実績処理日":    If Not IsNull(adoCol.Value) Then .実績処理日   = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者       = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function JTSRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    JTSRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 受注単価遡りテーブル"
        RTxtSql = RTxtSql &" Where 品番 = '" & RTrim(JTS.品番) & "'"
        RTxtSql = RTxtSql &" And   取引先 = '" & RTrim(JTS.取引先) & "'"
        RTxtSql = RTxtSql &" And   管理番号 = '" & RTrim(JTS.管理番号) & "'"
        RTxtSql = RTxtSql &" And   履歴NO = " & JTS.履歴NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           JTSRead = True
           Call ItemsFromJTS(DBNo)
	   JTS_RDSTS = True
        else
	   JTS_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call JTSClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub JTSReadNext(Optional DBNo As Integer = 0)
    If Not JTS_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           JTS_RDSTS = False
        Else
           Call ItemsFromJTS(DBNo)
           JTS_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub JTSClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        JTS_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( JTS )
'*----------------------------------------------*'
Public Sub ItemsClearJTS()
    With JTS
	.品番 = ""
	.取引先 = ""
	.管理番号 = ""
	.履歴NO = 0
	.仮単価区分 = ""
	.単価 = 0
	.適用開始日付 = ""
	.適用終了日付 = ""
	.部品単価 = 0
	.梱包単価 = 0
	.実績年月 = ""
	.実績処理日 = ""
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
'	.更新日 = 0
	.更新者 = ""
'	.更新端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function JTSInsert(Optional DBNo As Integer) As Boolean
    JTSInsert = False
    With JTS
        gSL_Select = "Insert Into 受注単価遡りテーブル  ( "
        gSL_Select = gSL_Select & "品番,取引先,管理番号,"
        gSL_Select = gSL_Select & "履歴NO,仮単価区分,単価,"
        gSL_Select = gSL_Select & "適用開始日付,適用終了日付,部品単価,"
        gSL_Select = gSL_Select & "梱包単価,実績年月,実績処理日,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.管理番号) & "',"		
        gSL_Select = gSL_Select & "" & .履歴NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.仮単価区分) & "',"		
        gSL_Select = gSL_Select & "" & .単価 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.適用開始日付) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.適用終了日付) & "',"		
        gSL_Select = gSL_Select & "" & .部品単価 & ","		
        gSL_Select = gSL_Select & "" & .梱包単価 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.実績年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.実績処理日) & "',"		
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
    JTSInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function JTSUpdate(Optional DBNo As Integer) As Boolean
    JTSUpdate = False	
    With JTS
        gSL_Select = "Update 受注単価遡りテーブル  Set "
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "管理番号 = '" & RTrim(.管理番号) & "',"
        gSL_Select = gSL_Select & "履歴NO = " & .履歴NO & ","
        gSL_Select = gSL_Select & "仮単価区分 = '" & RTrim(.仮単価区分) & "',"
        gSL_Select = gSL_Select & "単価 = " & .単価 & ","
        gSL_Select = gSL_Select & "適用開始日付 = '" & RTrim(.適用開始日付) & "',"
        gSL_Select = gSL_Select & "適用終了日付 = '" & RTrim(.適用終了日付) & "',"
        gSL_Select = gSL_Select & "部品単価 = " & .部品単価 & ","
        gSL_Select = gSL_Select & "梱包単価 = " & .梱包単価 & ","
        gSL_Select = gSL_Select & "実績年月 = '" & RTrim(.実績年月) & "',"
        gSL_Select = gSL_Select & "実績処理日 = '" & RTrim(.実績処理日) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 管理番号 = '" & RTrim(.管理番号) & "'"
        gSL_Select = gSL_Select & "   And 履歴NO = " & .履歴NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    JTSUpdate = True
End Function

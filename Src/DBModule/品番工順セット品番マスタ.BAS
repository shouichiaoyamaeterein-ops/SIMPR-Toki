Attribute VB_Name = "品番工順セット品番マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   品番工順セット品番マスタ
'*
'*   モジュールID  : 品番工順セット品番マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2009/07/27  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_品番工順セット品番マスタ
        品番              As String
        取引先            As String
        受入              As String
        品目              As Byte
        工順              As Byte
        セット品番        As String
        取数              As Long
'        作成日            As Date
        作成者            As String
'        作成端末          As String
'        更新日            As Date
        更新者            As String
'        更新端末          As String
        件数          As Long
    End Type

    Public HSE As T_品番工順セット品番マスタ

    Public HSE_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHSE(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHSE	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With HSE
            Select Case adoCol.Name
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番             = adoCol.Value
                Case "取引先":            If Not IsNull(adoCol.Value) Then .取引先           = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入             = adoCol.Value
                Case "品目":              If Not IsNull(adoCol.Value) Then .品目             = adoCol.Value
                Case "工順":              If Not IsNull(adoCol.Value) Then .工順             = adoCol.Value
                Case "セット品番":        If Not IsNull(adoCol.Value) Then .セット品番       = adoCol.Value
                Case "取数":              If Not IsNull(adoCol.Value) Then .取数             = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
                Case "更新者":            If Not IsNull(adoCol.Value) Then .更新者           = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function HSERead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HSERead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 品番工順セット品番マスタ"
        RTxtSql = RTxtSql &" Where 品番 = '" & RTrim(HSE.品番) & "'"
        RTxtSql = RTxtSql &" And   取引先 = '" & RTrim(HSE.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(HSE.受入) & "'"
        RTxtSql = RTxtSql &" And   品目 = " & HSE.品目 & ""
        RTxtSql = RTxtSql &" And   工順 = " & HSE.工順 & ""
        RTxtSql = RTxtSql &" And   セット品番 = '" & RTrim(HSE.セット品番) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HSERead = True
           Call ItemsFromHSE(DBNo)
	   HSE_RDSTS = True
        else
	   HSE_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call HSEClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HSEReadNext(Optional DBNo As Integer = 0)
    If Not HSE_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HSE_RDSTS = False
        Else
           Call ItemsFromHSE(DBNo)
           HSE_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HSEClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HSE_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HSE )
'*----------------------------------------------*'
Public Sub ItemsClearHSE()
    With HSE
	.品番 = ""
	.取引先 = ""
	.受入 = ""
	.品目 = 0
	.工順 = 0
	.セット品番 = ""
	.取数 = 0
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
Public Function HSEInsert(Optional DBNo As Integer) As Boolean
    HSEInsert = False
    With HSE
        gSL_Select = "Insert Into 品番工順セット品番マスタ  ( "
        gSL_Select = gSL_Select & "品番,取引先,受入,"
        gSL_Select = gSL_Select & "品目,工順,セット品番,"
        gSL_Select = gSL_Select & "取数,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "" & .品目 & ","		
        gSL_Select = gSL_Select & "" & .工順 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.セット品番) & "',"		
        gSL_Select = gSL_Select & "" & .取数 & ","		
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
    HSEInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function HSEUpdate(Optional DBNo As Integer) As Boolean
    HSEUpdate = False	
    With HSE
        gSL_Select = "Update 品番工順セット品番マスタ  Set "
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "品目 = " & .品目 & ","
        gSL_Select = gSL_Select & "工順 = " & .工順 & ","
        gSL_Select = gSL_Select & "セット品番 = '" & RTrim(.セット品番) & "',"
        gSL_Select = gSL_Select & "取数 = " & .取数 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 品目 = " & .品目 & ""
        gSL_Select = gSL_Select & "   And 工順 = " & .工順 & ""
        gSL_Select = gSL_Select & "   And セット品番 = '" & RTrim(.セット品番) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HSEUpdate = True
End Function

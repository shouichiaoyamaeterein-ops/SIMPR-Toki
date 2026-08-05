Attribute VB_Name = "取引先工程マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   取引先工程マスタ
'*
'*   モジュールID  : 取引先工程マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2009/07/15  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_取引先工程マスタ
        取引先CD          As String
        機械              As String
        機械NO            As String
        稼働時間          As Currency
        台数              As Byte
        代替機械          As String
        代替機械NO        As String
'        作成日            As Date
        作成者            As String
'        作成端末          As String
'        更新日            As Date
        更新者            As String
'        更新端末          As String
        件数          As Long
'   項目追加(自動)
	機械名         As String
	機械NO名       As String
	代替機械名     As String
	代替機械NO名   As String
    End Type

    Public TKT As T_取引先工程マスタ

    Public TKT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTKT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTKT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With TKT
            Select Case adoCol.Name
                Case "取引先CD":          If Not IsNull(adoCol.Value) Then .取引先CD         = adoCol.Value
                Case "機械":              If Not IsNull(adoCol.Value) Then .機械             = adoCol.Value
                Case "機械NO":            If Not IsNull(adoCol.Value) Then .機械NO           = adoCol.Value
                Case "稼働時間":          If Not IsNull(adoCol.Value) Then .稼働時間         = adoCol.Value
                Case "台数":              If Not IsNull(adoCol.Value) Then .台数             = adoCol.Value
                Case "代替機械":          If Not IsNull(adoCol.Value) Then .代替機械         = adoCol.Value
                Case "代替機械NO":        If Not IsNull(adoCol.Value) Then .代替機械NO       = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
                Case "更新者":            If Not IsNull(adoCol.Value) Then .更新者           = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "機械名":	If Not IsNull(adoCol.Value) Then .機械名	= adoCol.Value
		Case "機械NO名":	If Not IsNull(adoCol.Value) Then .機械NO名	= adoCol.Value
		Case "代替機械名":	If Not IsNull(adoCol.Value) Then .代替機械名	= adoCol.Value
		Case "代替機械NO名":	If Not IsNull(adoCol.Value) Then .代替機械NO名	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TKTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TKTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 取引先工程マスタ"
        RTxtSql = RTxtSql &" Where 取引先CD = '" & RTrim(TKT.取引先CD) & "'"
        RTxtSql = RTxtSql &" And   機械 = '" & RTrim(TKT.機械) & "'"
        RTxtSql = RTxtSql &" And   機械NO = '" & RTrim(TKT.機械NO) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TKTRead = True
           Call ItemsFromTKT(DBNo)
	   TKT_RDSTS = True
        else
	   TKT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call TKTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TKTReadNext(Optional DBNo As Integer = 0)
    If Not TKT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TKT_RDSTS = False
        Else
           Call ItemsFromTKT(DBNo)
           TKT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TKTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TKT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TKT )
'*----------------------------------------------*'
Public Sub ItemsClearTKT()
    With TKT
	.取引先CD = ""
	.機械 = ""
	.機械NO = ""
	.稼働時間 = 0
	.台数 = 0
	.代替機械 = ""
	.代替機械NO = ""
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
'	.更新日 = 0
	.更新者 = ""
'	.更新端末 = ""
        .件数 = 0
'   項目追加(自動)
	.機械名 = ""
	.機械NO名 = ""
	.代替機械名 = ""
	.代替機械NO名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function TKTInsert(Optional DBNo As Integer) As Boolean
    TKTInsert = False
    With TKT
        gSL_Select = "Insert Into 取引先工程マスタ  ( "
        gSL_Select = gSL_Select & "取引先CD,機械,機械NO,"
        gSL_Select = gSL_Select & "稼働時間,台数,代替機械,"
        gSL_Select = gSL_Select & "代替機械NO,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.取引先CD) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.機械) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.機械NO) & "',"		
        gSL_Select = gSL_Select & "" & .稼働時間 & ","		
        gSL_Select = gSL_Select & "" & .台数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.代替機械) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.代替機械NO) & "',"		
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
    TKTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function TKTUpdate(Optional DBNo As Integer) As Boolean
    TKTUpdate = False	
    With TKT
        gSL_Select = "Update 取引先工程マスタ  Set "
        gSL_Select = gSL_Select & "取引先CD = '" & RTrim(.取引先CD) & "',"
        gSL_Select = gSL_Select & "機械 = '" & RTrim(.機械) & "',"
        gSL_Select = gSL_Select & "機械NO = '" & RTrim(.機械NO) & "',"
        gSL_Select = gSL_Select & "稼働時間 = " & .稼働時間 & ","
        gSL_Select = gSL_Select & "台数 = " & .台数 & ","
        gSL_Select = gSL_Select & "代替機械 = '" & RTrim(.代替機械) & "',"
        gSL_Select = gSL_Select & "代替機械NO = '" & RTrim(.代替機械NO) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 取引先CD = '" & RTrim(.取引先CD) & "'"
        gSL_Select = gSL_Select & "   And 機械 = '" & RTrim(.機械) & "'"
        gSL_Select = gSL_Select & "   And 機械NO = '" & RTrim(.機械NO) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TKTUpdate = True
End Function

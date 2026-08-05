Attribute VB_Name = "品番置換マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   品番置換マスタ
'*
'*   モジュールID  : 品番置換マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*
'*   作成日  :  2004/01/06  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_品番置換マスタ
        品番              As String
        得意先品番        As String
        納入先            As String
        受入              As String
'        作成日            As Date
        作成者            As String
'        作成端末          As String
'        更新日            As Date
        更新者            As String
'        更新端末          As String
        件数          As Long
    End Type

    Public HRM As T_品番置換マスタ

    Public HRM_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHRM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHRM	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With HRM
            Select Case adoCol.Name
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番             = adoCol.Value
                Case "得意先品番":        If Not IsNull(adoCol.Value) Then .得意先品番       = adoCol.Value
                Case "納入先":            If Not IsNull(adoCol.Value) Then .納入先           = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入             = adoCol.Value
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
Public Function HRMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HRMRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 品番置換マスタ"
        RTxtSql = RTxtSql &" Where 品番 = '" & RTrim(HRM.品番) & "'"
        RTxtSql = RTxtSql &" And   得意先品番 = '" & RTrim(HRM.得意先品番) & "'"
        RTxtSql = RTxtSql &" And   納入先 = '" & RTrim(HRM.納入先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(HRM.受入) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HRMRead = True
           Call ItemsFromHRM(DBNo)
	   HRM_RDSTS = True
        else
	   HRM_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call HRMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HRMReadNext(Optional DBNo As Integer = 0)
    If Not HRM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HRM_RDSTS = False
        Else
           Call ItemsFromHRM(DBNo)
           HRM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HRMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HRM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HRM )
'*----------------------------------------------*'
Public Sub ItemsClearHRM()
    With HRM
	.品番 = ""
	.得意先品番 = ""
	.納入先 = ""
	.受入 = ""
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
Public Function HRMInsert(Optional DBNo As Integer) As Boolean
    HRMInsert = True
    With HRM
        gSL_Select = "Insert Into 品番置換マスタ  ( "
        gSL_Select = gSL_Select & "品番,得意先品番,納入先,"
        gSL_Select = gSL_Select & "受入,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.得意先品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
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
    HRMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function HRMUpdate(Optional DBNo As Integer) As Boolean
    HRMUpdate = False	
    With HRM
        gSL_Select = "Update 品番置換マスタ  Set "
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "得意先品番 = '" & RTrim(.得意先品番) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 得意先品番 = '" & RTrim(.得意先品番) & "'"
        gSL_Select = gSL_Select & "   And 納入先 = '" & RTrim(.納入先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HRMUpdate = True
End Function

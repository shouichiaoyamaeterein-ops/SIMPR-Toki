Attribute VB_Name = "取引先受入マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   取引先受入マスタ
'*
'*   モジュールID  : 取引先受入マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2005/07/15  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_取引先受入マスタ
        取引先CD          As String
        取引先区分        As Byte
        受入              As String
        受入名称          As String
        サイクル          As String
        空明細            As String
        曜日発注          As Byte
        曜日1             As Byte
        曜日2             As Byte
        曜日3             As Byte
        曜日4             As Byte
        曜日5             As Byte
        件数          As Long
    End Type

    Public TUM As T_取引先受入マスタ

    Public TUM_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTUM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTUM	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With TUM
            Select Case adoCol.Name
                Case "取引先CD":          If Not IsNull(adoCol.Value) Then .取引先CD         = adoCol.Value
                Case "取引先区分":        If Not IsNull(adoCol.Value) Then .取引先区分       = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入             = adoCol.Value
                Case "受入名称":          If Not IsNull(adoCol.Value) Then .受入名称         = adoCol.Value
                Case "サイクル":          If Not IsNull(adoCol.Value) Then .サイクル         = adoCol.Value
                Case "空明細":            If Not IsNull(adoCol.Value) Then .空明細           = adoCol.Value
                Case "曜日発注":          If Not IsNull(adoCol.Value) Then .曜日発注         = adoCol.Value
                Case "曜日1":             If Not IsNull(adoCol.Value) Then .曜日1            = adoCol.Value
                Case "曜日2":             If Not IsNull(adoCol.Value) Then .曜日2            = adoCol.Value
                Case "曜日3":             If Not IsNull(adoCol.Value) Then .曜日3            = adoCol.Value
                Case "曜日4":             If Not IsNull(adoCol.Value) Then .曜日4            = adoCol.Value
                Case "曜日5":             If Not IsNull(adoCol.Value) Then .曜日5            = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TUMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TUMRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 取引先受入マスタ"
        RTxtSql = RTxtSql &" Where 取引先CD = '" & RTrim(TUM.取引先CD) & "'"
        RTxtSql = RTxtSql &" And   取引先区分 = " & TUM.取引先区分 & ""
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(TUM.受入) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TUMRead = True
           Call ItemsFromTUM(DBNo)
	   TUM_RDSTS = True
        else
	   TUM_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call TUMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TUMReadNext(Optional DBNo As Integer = 0)
    If Not TUM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TUM_RDSTS = False
        Else
           Call ItemsFromTUM(DBNo)
           TUM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TUMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TUM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TUM )
'*----------------------------------------------*'
Public Sub ItemsClearTUM()
    With TUM
	.取引先CD = ""
	.取引先区分 = 0
	.受入 = ""
	.受入名称 = ""
	.サイクル = ""
	.空明細 = ""
	.曜日発注 = 0
	.曜日1 = 0
	.曜日2 = 0
	.曜日3 = 0
	.曜日4 = 0
	.曜日5 = 0
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function TUMInsert(Optional DBNo As Integer) As Boolean
    TUMInsert = True
    With TUM
        gSL_Select = "Insert Into 取引先受入マスタ  ( "
        gSL_Select = gSL_Select & "取引先CD,取引先区分,受入,"
        gSL_Select = gSL_Select & "受入名称,サイクル,空明細,"
        gSL_Select = gSL_Select & "曜日発注,曜日1,曜日2,"
        gSL_Select = gSL_Select & "曜日3,曜日4,曜日5"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.取引先CD) & "',"		
        gSL_Select = gSL_Select & "" & .取引先区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入名称) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.サイクル) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.空明細) & "',"		
        gSL_Select = gSL_Select & "" & .曜日発注 & ","		
        gSL_Select = gSL_Select & "" & .曜日1 & ","		
        gSL_Select = gSL_Select & "" & .曜日2 & ","		
        gSL_Select = gSL_Select & "" & .曜日3 & ","		
        gSL_Select = gSL_Select & "" & .曜日4 & ","		
        gSL_Select = gSL_Select & "" & .曜日5 & ""		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TUMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function TUMUpdate(Optional DBNo As Integer) As Boolean
    TUMUpdate = False	
    With TUM
        gSL_Select = "Update 取引先受入マスタ  Set "
        gSL_Select = gSL_Select & "取引先CD = '" & RTrim(.取引先CD) & "',"
        gSL_Select = gSL_Select & "取引先区分 = " & .取引先区分 & ","
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "受入名称 = '" & RTrim(.受入名称) & "',"
        gSL_Select = gSL_Select & "サイクル = '" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "空明細 = '" & RTrim(.空明細) & "',"
        gSL_Select = gSL_Select & "曜日発注 = " & .曜日発注 & ","
        gSL_Select = gSL_Select & "曜日1 = " & .曜日1 & ","
        gSL_Select = gSL_Select & "曜日2 = " & .曜日2 & ","
        gSL_Select = gSL_Select & "曜日3 = " & .曜日3 & ","
        gSL_Select = gSL_Select & "曜日4 = " & .曜日4 & ","
        gSL_Select = gSL_Select & "曜日5 = " & .曜日5 & ""
        gSL_Select = gSL_Select & " Where 取引先CD = '" & RTrim(.取引先CD) & "'"
        gSL_Select = gSL_Select & "   And 取引先区分 = " & .取引先区分 & ""
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TUMUpdate = True
End Function

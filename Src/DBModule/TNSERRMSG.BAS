Attribute VB_Name = "TNSERRMSG"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TNSERRMSG
'*
'*   モジュールID  : TNSERRMSG.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2005/09/05  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_TNSERRMSG
        ERRCODE   As String
        ERRLEBEL  As String
        ERRMSG1   As String
        ERRMSG2   As String
        件数          As Long
    End Type

    Public TET As T_TNSERRMSG

    Public TET_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTET(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTET	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With TET
            Select Case adoCol.Name
                Case "ERRCODE":   If Not IsNull(adoCol.Value) Then .ERRCODE  = adoCol.Value
                Case "ERRLEBEL":  If Not IsNull(adoCol.Value) Then .ERRLEBEL = adoCol.Value
                Case "ERRMSG1":   If Not IsNull(adoCol.Value) Then .ERRMSG1  = adoCol.Value
                Case "ERRMSG2":   If Not IsNull(adoCol.Value) Then .ERRMSG2  = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TETRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TETRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From TNSERRMSG"
        RTxtSql = RTxtSql &" Where ERRCODE = '" & RTrim(TET.ERRCODE) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TETRead = True
           Call ItemsFromTET(DBNo)
	   TET_RDSTS = True
        else
	   TET_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call TETClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TETReadNext(Optional DBNo As Integer = 0)
    If Not TET_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TET_RDSTS = False
        Else
           Call ItemsFromTET(DBNo)
           TET_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TETClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TET_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TET )
'*----------------------------------------------*'
Public Sub ItemsClearTET()
    With TET
	.ERRCODE = ""
	.ERRLEBEL = ""
	.ERRMSG1 = ""
	.ERRMSG2 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function TETInsert(Optional DBNo As Integer) As Boolean
    TETInsert = True
    With TET
        gSL_Select = "Insert Into TNSERRMSG  ( "
        gSL_Select = gSL_Select & "ERRCODE,ERRLEBEL,ERRMSG1,"
        gSL_Select = gSL_Select & "ERRMSG2"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.ERRCODE) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ERRLEBEL) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ERRMSG1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ERRMSG2) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TETInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function TETUpdate(Optional DBNo As Integer) As Boolean
    TETUpdate = False	
    With TET
        gSL_Select = "Update TNSERRMSG  Set "
        gSL_Select = gSL_Select & "ERRCODE = '" & RTrim(.ERRCODE) & "',"
        gSL_Select = gSL_Select & "ERRLEBEL = '" & RTrim(.ERRLEBEL) & "',"
        gSL_Select = gSL_Select & "ERRMSG1 = '" & RTrim(.ERRMSG1) & "',"
        gSL_Select = gSL_Select & "ERRMSG2 = '" & RTrim(.ERRMSG2) & "'"
        gSL_Select = gSL_Select & " Where ERRCODE = '" & RTrim(.ERRCODE) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TETUpdate = True
End Function

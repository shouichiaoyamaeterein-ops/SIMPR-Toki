Attribute VB_Name = "TEXT変換システムワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TEXT変換システムワーク
'*
'*   モジュールID  : TEXT変換システムワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2005/12/06  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_TEXT変換システムワーク
        cname         As String
        tname         As String
        prec          As Integer
        scale         As Long
        件数          As Long
    End Type

    Public TXH As T_TEXT変換システムワーク

    Public TXH_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTXH(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTXH	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With TXH
            Select Case adoCol.Name
                Case "cname":         If Not IsNull(adoCol.Value) Then .cname        = adoCol.Value
                Case "tname":         If Not IsNull(adoCol.Value) Then .tname        = adoCol.Value
                Case "prec":          If Not IsNull(adoCol.Value) Then .prec         = adoCol.Value
                Case "scale":         If Not IsNull(adoCol.Value) Then .scale        = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TXHRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TXHRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From TEXT変換システムワーク"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TXHRead = True
           Call ItemsFromTXH(DBNo)
	   TXH_RDSTS = True
        else
	   TXH_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call TXHClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TXHReadNext(Optional DBNo As Integer = 0)
    If Not TXH_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TXH_RDSTS = False
        Else
           Call ItemsFromTXH(DBNo)
           TXH_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TXHClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TXH_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TXH )
'*----------------------------------------------*'
Public Sub ItemsClearTXH()
    With TXH
	.cname = ""
	.tname = ""
	.prec = 0
	.scale = 0
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function TXHInsert(Optional DBNo As Integer) As Boolean
    TXHInsert = True
    With TXH
        gSL_Select = "Insert Into TEXT変換システムワーク  ( "
        gSL_Select = gSL_Select & "cname,tname,prec,"
        gSL_Select = gSL_Select & "scale"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.cname) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.tname) & "',"		
        gSL_Select = gSL_Select & "" & .prec & ","		
        gSL_Select = gSL_Select & "" & .scale & ""		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TXHInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function TXHUpdate(Optional DBNo As Integer) As Boolean
    TXHUpdate = False	
    With TXH
        gSL_Select = "Update TEXT変換システムワーク  Set "
        gSL_Select = gSL_Select & "cname = '" & RTrim(.cname) & "',"
        gSL_Select = gSL_Select & "tname = '" & RTrim(.tname) & "',"
        gSL_Select = gSL_Select & "prec = " & .prec & ","
        gSL_Select = gSL_Select & "scale = " & .scale & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TXHUpdate = True
End Function

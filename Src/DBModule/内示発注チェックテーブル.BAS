Attribute VB_Name = "内示発注チェックテーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   内示発注チェックテーブル
'*
'*   モジュールID  : 内示発注チェックテーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/03/23  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_内示発注チェックテーブル
        発注年月  As String
        品番      As String
        件数          As Long
    End Type

    Public NHC As T_内示発注チェックテーブル

    Public NHC_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromNHC(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearNHC	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With NHC
            Select Case adoCol.Name
                Case "発注年月":  If Not IsNull(adoCol.Value) Then .発注年月 = adoCol.Value
                Case "品番":      If Not IsNull(adoCol.Value) Then .品番     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function NHCRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    NHCRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 内示発注チェックテーブル"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           NHCRead = True
           Call ItemsFromNHC(DBNo)
	   NHC_RDSTS = True
        else
	   NHC_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call NHCClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub NHCReadNext(Optional DBNo As Integer = 0)
    If Not NHC_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           NHC_RDSTS = False
        Else
           Call ItemsFromNHC(DBNo)
           NHC_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub NHCClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        NHC_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( NHC )
'*----------------------------------------------*'
Public Sub ItemsClearNHC()
    With NHC
	.発注年月 = ""
	.品番 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function NHCInsert(Optional DBNo As Integer) As Boolean
    NHCInsert = True
    With NHC
        gSL_Select = "Insert Into 内示発注チェックテーブル  ( "
        gSL_Select = gSL_Select & "発注年月,品番"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.発注年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    NHCInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function NHCUpdate(Optional DBNo As Integer) As Boolean
    NHCUpdate = False	
    With NHC
        gSL_Select = "Update 内示発注チェックテーブル  Set "
        gSL_Select = gSL_Select & "発注年月 = '" & RTrim(.発注年月) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    NHCUpdate = True
End Function

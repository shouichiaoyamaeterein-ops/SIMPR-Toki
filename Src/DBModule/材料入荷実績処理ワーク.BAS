Attribute VB_Name = "材料入荷実績処理ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料入荷実績処理ワーク
'*
'*   モジュールID  : 材料入荷実績処理ワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/02/24  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料入荷実績処理ワーク
        処理端末          As String
        発注管理NO        As Long
        処理区分          As Byte
'        作成日            As Date
        件数          As Long
    End Type

    Public ZJW As T_材料入荷実績処理ワーク

    Public ZJW_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZJW(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZJW	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZJW
            Select Case adoCol.Name
                Case "処理端末":          If Not IsNull(adoCol.Value) Then .処理端末         = adoCol.Value
                Case "発注管理NO":        If Not IsNull(adoCol.Value) Then .発注管理NO       = adoCol.Value
                Case "処理区分":          If Not IsNull(adoCol.Value) Then .処理区分         = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZJWRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZJWRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 材料入荷実績処理ワーク"
        RTxtSql = RTxtSql &" Where 処理端末 = '" & RTrim(ZJW.処理端末) & "'"
        RTxtSql = RTxtSql &" And   発注管理NO = " & ZJW.発注管理NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZJWRead = True
           Call ItemsFromZJW(DBNo)
	   ZJW_RDSTS = True
        else
	   ZJW_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZJWClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZJWReadNext(Optional DBNo As Integer = 0)
    If Not ZJW_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZJW_RDSTS = False
        Else
           Call ItemsFromZJW(DBNo)
           ZJW_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZJWClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZJW_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZJW )
'*----------------------------------------------*'
Public Sub ItemsClearZJW()
    With ZJW
	.処理端末 = ""
	.発注管理NO = 0
	.処理区分 = 0
'	.作成日 = 0
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZJWInsert(Optional DBNo As Integer) As Boolean
    ZJWInsert = True
    With ZJW
        gSL_Select = "Insert Into 材料入荷実績処理ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,発注管理NO,処理区分,"
        gSL_Select = gSL_Select & "作成日"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"		
        gSL_Select = gSL_Select & "" & .発注管理NO & ","		
        gSL_Select = gSL_Select & "" & .処理区分 & ","		
        gSL_Select = gSL_Select & "GETDATE()"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZJWInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZJWUpdate(Optional DBNo As Integer) As Boolean
    ZJWUpdate = False	
    With ZJW
        gSL_Select = "Update 材料入荷実績処理ワーク  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "発注管理NO = " & .発注管理NO & ","
        gSL_Select = gSL_Select & "処理区分 = " & .処理区分 & ""
        gSL_Select = gSL_Select & " Where 処理端末 = '" & RTrim(.処理端末) & "'"
        gSL_Select = gSL_Select & "   And 発注管理NO = " & .発注管理NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZJWUpdate = True
End Function

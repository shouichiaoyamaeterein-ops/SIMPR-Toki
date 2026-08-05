Attribute VB_Name = "S内示問合せ日量ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   S内示問合せ日量ワーク
'*
'*   モジュールID  : S内示問合せ日量ワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/05/20  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_S内示問合せ日量ワーク
        処理端末  As String
        順位      As String
        レベル    As Byte
        品番      As String
        手配先    As String
        受入      As String
        背番号    As String
        納期      As String
        数量      As Currency
        件数          As Long
    End Type

    Public SNN As T_S内示問合せ日量ワーク

    Public SNN_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromSNN(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearSNN	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With SNN
            Select Case adoCol.Name
                Case "処理端末":  If Not IsNull(adoCol.Value) Then .処理端末 = adoCol.Value
                Case "順位":      If Not IsNull(adoCol.Value) Then .順位     = adoCol.Value
                Case "レベル":    If Not IsNull(adoCol.Value) Then .レベル   = adoCol.Value
                Case "品番":      If Not IsNull(adoCol.Value) Then .品番     = adoCol.Value
                Case "手配先":    If Not IsNull(adoCol.Value) Then .手配先   = adoCol.Value
                Case "受入":      If Not IsNull(adoCol.Value) Then .受入     = adoCol.Value
                Case "背番号":    If Not IsNull(adoCol.Value) Then .背番号   = adoCol.Value
                Case "納期":      If Not IsNull(adoCol.Value) Then .納期     = adoCol.Value
                Case "数量":      If Not IsNull(adoCol.Value) Then .数量     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function SNNRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    SNNRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From S内示問合せ日量ワーク"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           SNNRead = True
           Call ItemsFromSNN(DBNo)
	   SNN_RDSTS = True
        else
	   SNN_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call SNNClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub SNNReadNext(Optional DBNo As Integer = 0)
    If Not SNN_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           SNN_RDSTS = False
        Else
           Call ItemsFromSNN(DBNo)
           SNN_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub SNNClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        SNN_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( SNN )
'*----------------------------------------------*'
Public Sub ItemsClearSNN()
    With SNN
	.処理端末 = ""
	.順位 = ""
	.レベル = 0
	.品番 = ""
	.手配先 = ""
	.受入 = ""
	.背番号 = ""
	.納期 = ""
	.数量 = 0
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function SNNInsert(Optional DBNo As Integer) As Boolean
    SNNInsert = True
    With SNN
        gSL_Select = "Insert Into S内示問合せ日量ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,順位,レベル,"
        gSL_Select = gSL_Select & "品番,手配先,受入,"
        gSL_Select = gSL_Select & "背番号,納期,数量"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.順位) & "',"		
        gSL_Select = gSL_Select & "" & .レベル & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納期) & "',"		
        gSL_Select = gSL_Select & "" & .数量 & ""		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SNNInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function SNNUpdate(Optional DBNo As Integer) As Boolean
    SNNUpdate = False	
    With SNN
        gSL_Select = "Update S内示問合せ日量ワーク  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "順位 = '" & RTrim(.順位) & "',"
        gSL_Select = gSL_Select & "レベル = " & .レベル & ","
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "手配先 = '" & RTrim(.手配先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "納期 = '" & RTrim(.納期) & "',"
        gSL_Select = gSL_Select & "数量 = " & .数量 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SNNUpdate = True
End Function

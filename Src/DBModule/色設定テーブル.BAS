Attribute VB_Name = "色設定テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   色設定テーブル
'*
'*   モジュールID  : 色設定テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2008/06/09  By Cis(FMV-W5210)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_色設定テーブル
        受入          As String
        背景色        As Byte
        枠色          As Byte
'        作成日        As Date
        作成者        As String
'        作成端末      As String
'        更新日        As Date
        更新者        As String
'        更新端末      As String
        件数          As Long
    End Type

    Public IST As T_色設定テーブル

    Public IST_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromIST(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearIST	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With IST
            Select Case adoCol.Name
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "背景色":        If Not IsNull(adoCol.Value) Then .背景色       = adoCol.Value
                Case "枠色":          If Not IsNull(adoCol.Value) Then .枠色         = adoCol.Value
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
Public Function ISTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ISTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 色設定テーブル"
        RTxtSql = RTxtSql &" Where 受入 = '" & RTrim(IST.受入) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ISTRead = True
           Call ItemsFromIST(DBNo)
	   IST_RDSTS = True
        else
	   IST_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ISTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ISTReadNext(Optional DBNo As Integer = 0)
    If Not IST_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           IST_RDSTS = False
        Else
           Call ItemsFromIST(DBNo)
           IST_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ISTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        IST_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( IST )
'*----------------------------------------------*'
Public Sub ItemsClearIST()
    With IST
	.受入 = ""
	.背景色 = 0
	.枠色 = 0
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
Public Function ISTInsert(Optional DBNo As Integer) As Boolean
    ISTInsert = False
    With IST
        gSL_Select = "Insert Into 色設定テーブル  ( "
        gSL_Select = gSL_Select & "受入,背景色,枠色,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "" & .背景色 & ","		
        gSL_Select = gSL_Select & "" & .枠色 & ","		
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
    ISTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ISTUpdate(Optional DBNo As Integer) As Boolean
    ISTUpdate = False	
    With IST
        gSL_Select = "Update 色設定テーブル  Set "
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "背景色 = " & .背景色 & ","
        gSL_Select = gSL_Select & "枠色 = " & .枠色 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 受入 = '" & RTrim(.受入) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ISTUpdate = True
End Function

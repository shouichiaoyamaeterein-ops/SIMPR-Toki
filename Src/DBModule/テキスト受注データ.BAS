Attribute VB_Name = "テキスト受注データ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   テキスト受注データ
'*
'*   モジュールID  : テキスト受注データ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2009/09/15  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_テキスト受注データ
        REC1  As String
        REC2  As String
        REC3  As String
        REC4  As String
        件数          As Long
    End Type

    Public TJY As T_テキスト受注データ

    Public TJY_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTJY(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTJY	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With TJY
            Select Case adoCol.Name
                Case "REC1":  If Not IsNull(adoCol.Value) Then .REC1 = adoCol.Value
                Case "REC2":  If Not IsNull(adoCol.Value) Then .REC2 = adoCol.Value
                Case "REC3":  If Not IsNull(adoCol.Value) Then .REC3 = adoCol.Value
                Case "REC4":  If Not IsNull(adoCol.Value) Then .REC4 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TJYRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TJYRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From テキスト受注データ"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TJYRead = True
           Call ItemsFromTJY(DBNo)
	   TJY_RDSTS = True
        else
	   TJY_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call TJYClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TJYReadNext(Optional DBNo As Integer = 0)
    If Not TJY_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TJY_RDSTS = False
        Else
           Call ItemsFromTJY(DBNo)
           TJY_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TJYClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TJY_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TJY )
'*----------------------------------------------*'
Public Sub ItemsClearTJY()
    With TJY
	.REC1 = ""
	.REC2 = ""
	.REC3 = ""
	.REC4 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function TJYInsert(Optional DBNo As Integer) As Boolean
    TJYInsert = False
    With TJY
        gSL_Select = "Insert Into テキスト受注データ  ( "
        gSL_Select = gSL_Select & "REC1,REC2,REC3,"
        gSL_Select = gSL_Select & "REC4"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.REC1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.REC2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.REC3) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.REC4) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TJYInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function TJYUpdate(Optional DBNo As Integer) As Boolean
    TJYUpdate = False	
    With TJY
        gSL_Select = "Update テキスト受注データ  Set "
        gSL_Select = gSL_Select & "REC1 = '" & RTrim(.REC1) & "',"
        gSL_Select = gSL_Select & "REC2 = '" & RTrim(.REC2) & "',"
        gSL_Select = gSL_Select & "REC3 = '" & RTrim(.REC3) & "',"
        gSL_Select = gSL_Select & "REC4 = '" & RTrim(.REC4) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TJYUpdate = True
End Function

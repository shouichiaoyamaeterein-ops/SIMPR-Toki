Attribute VB_Name = "工事手配書ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   工事手配書ワーク
'*
'*   モジュールID  : 工事手配書ワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2010/01/17  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_工事手配書ワーク
        端末番号      As String
        手配先        As String
        手配先名      As String
        工場名        As String
        品番          As String
        表示品番      As String
        品名          As String
        受注実績見込  As Long
        当月所要量    As Long
        件数          As Long
    End Type

    Public KJT As T_工事手配書ワーク

    Public KJT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKJT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKJT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KJT
            Select Case adoCol.Name
                Case "端末番号":      If Not IsNull(adoCol.Value) Then .端末番号     = adoCol.Value
                Case "手配先":        If Not IsNull(adoCol.Value) Then .手配先       = adoCol.Value
                Case "手配先名":      If Not IsNull(adoCol.Value) Then .手配先名     = adoCol.Value
                Case "工場名":        If Not IsNull(adoCol.Value) Then .工場名       = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "表示品番":      If Not IsNull(adoCol.Value) Then .表示品番     = adoCol.Value
                Case "品名":          If Not IsNull(adoCol.Value) Then .品名         = adoCol.Value
                Case "受注実績見込":  If Not IsNull(adoCol.Value) Then .受注実績見込 = adoCol.Value
                Case "当月所要量":    If Not IsNull(adoCol.Value) Then .当月所要量   = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KJTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KJTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 工事手配書ワーク"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KJTRead = True
           Call ItemsFromKJT(DBNo)
	   KJT_RDSTS = True
        else
	   KJT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KJTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KJTReadNext(Optional DBNo As Integer = 0)
    If Not KJT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KJT_RDSTS = False
        Else
           Call ItemsFromKJT(DBNo)
           KJT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KJTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KJT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KJT )
'*----------------------------------------------*'
Public Sub ItemsClearKJT()
    With KJT
	.端末番号 = ""
	.手配先 = ""
	.手配先名 = ""
	.工場名 = ""
	.品番 = ""
	.表示品番 = ""
	.品名 = ""
	.受注実績見込 = 0
	.当月所要量 = 0
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KJTInsert(Optional DBNo As Integer) As Boolean
    KJTInsert = False
    With KJT
        gSL_Select = "Insert Into 工事手配書ワーク  ( "
        gSL_Select = gSL_Select & "端末番号,手配先,手配先名,"
        gSL_Select = gSL_Select & "工場名,品番,表示品番,"
        gSL_Select = gSL_Select & "品名,受注実績見込,当月所要量"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.端末番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配先名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.工場名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.表示品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品名) & "',"		
        gSL_Select = gSL_Select & "" & .受注実績見込 & ","		
        gSL_Select = gSL_Select & "" & .当月所要量 & ""		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KJTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KJTUpdate(Optional DBNo As Integer) As Boolean
    KJTUpdate = False	
    With KJT
        gSL_Select = "Update 工事手配書ワーク  Set "
        gSL_Select = gSL_Select & "端末番号 = '" & RTrim(.端末番号) & "',"
        gSL_Select = gSL_Select & "手配先 = '" & RTrim(.手配先) & "',"
        gSL_Select = gSL_Select & "手配先名 = '" & RTrim(.手配先名) & "',"
        gSL_Select = gSL_Select & "工場名 = '" & RTrim(.工場名) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "表示品番 = '" & RTrim(.表示品番) & "',"
        gSL_Select = gSL_Select & "品名 = '" & RTrim(.品名) & "',"
        gSL_Select = gSL_Select & "受注実績見込 = " & .受注実績見込 & ","
        gSL_Select = gSL_Select & "当月所要量 = " & .当月所要量 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KJTUpdate = True
End Function

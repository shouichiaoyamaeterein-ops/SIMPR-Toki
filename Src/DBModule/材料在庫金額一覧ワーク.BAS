Attribute VB_Name = "材料在庫金額一覧ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料在庫金額一覧ワーク
'*
'*   モジュールID  : 材料在庫金額一覧ワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2010/01/06  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料在庫金額一覧ワーク
        処理端末              As String
        材料管理番号          As Long
        材質                  As String
        寸法                  As String
        手配先                As String
        手配先名              As String
        仕入単価              As Currency
        前月末在庫数          As Long
        前月末在庫金額        As Long
        当月入庫数            As Long
        当月入庫金額          As Long
        当月出庫数            As Long
        当月出庫金額          As Long
        調整数                As Long
        調整金額              As Long
        現在在庫数            As Long
        現在在庫金額          As Long
        最終入庫日            As String
        件数          As Long
    End Type

    Public ZKW As T_材料在庫金額一覧ワーク

    Public ZKW_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZKW(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZKW	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZKW
            Select Case adoCol.Name
                Case "処理端末":              If Not IsNull(adoCol.Value) Then .処理端末             = adoCol.Value
                Case "材料管理番号":          If Not IsNull(adoCol.Value) Then .材料管理番号         = adoCol.Value
                Case "材質":                  If Not IsNull(adoCol.Value) Then .材質                 = adoCol.Value
                Case "寸法":                  If Not IsNull(adoCol.Value) Then .寸法                 = adoCol.Value
                Case "手配先":                If Not IsNull(adoCol.Value) Then .手配先               = adoCol.Value
                Case "手配先名":              If Not IsNull(adoCol.Value) Then .手配先名             = adoCol.Value
                Case "仕入単価":              If Not IsNull(adoCol.Value) Then .仕入単価             = adoCol.Value
                Case "前月末在庫数":          If Not IsNull(adoCol.Value) Then .前月末在庫数         = adoCol.Value
                Case "前月末在庫金額":        If Not IsNull(adoCol.Value) Then .前月末在庫金額       = adoCol.Value
                Case "当月入庫数":            If Not IsNull(adoCol.Value) Then .当月入庫数           = adoCol.Value
                Case "当月入庫金額":          If Not IsNull(adoCol.Value) Then .当月入庫金額         = adoCol.Value
                Case "当月出庫数":            If Not IsNull(adoCol.Value) Then .当月出庫数           = adoCol.Value
                Case "当月出庫金額":          If Not IsNull(adoCol.Value) Then .当月出庫金額         = adoCol.Value
                Case "調整数":                If Not IsNull(adoCol.Value) Then .調整数               = adoCol.Value
                Case "調整金額":              If Not IsNull(adoCol.Value) Then .調整金額             = adoCol.Value
                Case "現在在庫数":            If Not IsNull(adoCol.Value) Then .現在在庫数           = adoCol.Value
                Case "現在在庫金額":          If Not IsNull(adoCol.Value) Then .現在在庫金額         = adoCol.Value
                Case "最終入庫日":            If Not IsNull(adoCol.Value) Then .最終入庫日           = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZKWRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZKWRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 材料在庫金額一覧ワーク"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZKWRead = True
           Call ItemsFromZKW(DBNo)
	   ZKW_RDSTS = True
        else
	   ZKW_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZKWClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZKWReadNext(Optional DBNo As Integer = 0)
    If Not ZKW_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZKW_RDSTS = False
        Else
           Call ItemsFromZKW(DBNo)
           ZKW_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZKWClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZKW_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZKW )
'*----------------------------------------------*'
Public Sub ItemsClearZKW()
    With ZKW
	.処理端末 = ""
	.材料管理番号 = 0
	.材質 = ""
	.寸法 = ""
	.手配先 = ""
	.手配先名 = ""
	.仕入単価 = 0
	.前月末在庫数 = 0
	.前月末在庫金額 = 0
	.当月入庫数 = 0
	.当月入庫金額 = 0
	.当月出庫数 = 0
	.当月出庫金額 = 0
	.調整数 = 0
	.調整金額 = 0
	.現在在庫数 = 0
	.現在在庫金額 = 0
	.最終入庫日 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZKWInsert(Optional DBNo As Integer) As Boolean
    ZKWInsert = False
    With ZKW
        gSL_Select = "Insert Into 材料在庫金額一覧ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,材料管理番号,材質,"
        gSL_Select = gSL_Select & "寸法,手配先,手配先名,"
        gSL_Select = gSL_Select & "仕入単価,前月末在庫数,前月末在庫金額,"
        gSL_Select = gSL_Select & "当月入庫数,当月入庫金額,当月出庫数,"
        gSL_Select = gSL_Select & "当月出庫金額,調整数,調整金額,"
        gSL_Select = gSL_Select & "現在在庫数,現在在庫金額,最終入庫日"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"		
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.寸法) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配先名) & "',"		
        gSL_Select = gSL_Select & "" & .仕入単価 & ","		
        gSL_Select = gSL_Select & "" & .前月末在庫数 & ","		
        gSL_Select = gSL_Select & "" & .前月末在庫金額 & ","		
        gSL_Select = gSL_Select & "" & .当月入庫数 & ","		
        gSL_Select = gSL_Select & "" & .当月入庫金額 & ","		
        gSL_Select = gSL_Select & "" & .当月出庫数 & ","		
        gSL_Select = gSL_Select & "" & .当月出庫金額 & ","		
        gSL_Select = gSL_Select & "" & .調整数 & ","		
        gSL_Select = gSL_Select & "" & .調整金額 & ","		
        gSL_Select = gSL_Select & "" & .現在在庫数 & ","		
        gSL_Select = gSL_Select & "" & .現在在庫金額 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.最終入庫日) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZKWInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZKWUpdate(Optional DBNo As Integer) As Boolean
    ZKWUpdate = False	
    With ZKW
        gSL_Select = "Update 材料在庫金額一覧ワーク  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "寸法 = '" & RTrim(.寸法) & "',"
        gSL_Select = gSL_Select & "手配先 = '" & RTrim(.手配先) & "',"
        gSL_Select = gSL_Select & "手配先名 = '" & RTrim(.手配先名) & "',"
        gSL_Select = gSL_Select & "仕入単価 = " & .仕入単価 & ","
        gSL_Select = gSL_Select & "前月末在庫数 = " & .前月末在庫数 & ","
        gSL_Select = gSL_Select & "前月末在庫金額 = " & .前月末在庫金額 & ","
        gSL_Select = gSL_Select & "当月入庫数 = " & .当月入庫数 & ","
        gSL_Select = gSL_Select & "当月入庫金額 = " & .当月入庫金額 & ","
        gSL_Select = gSL_Select & "当月出庫数 = " & .当月出庫数 & ","
        gSL_Select = gSL_Select & "当月出庫金額 = " & .当月出庫金額 & ","
        gSL_Select = gSL_Select & "調整数 = " & .調整数 & ","
        gSL_Select = gSL_Select & "調整金額 = " & .調整金額 & ","
        gSL_Select = gSL_Select & "現在在庫数 = " & .現在在庫数 & ","
        gSL_Select = gSL_Select & "現在在庫金額 = " & .現在在庫金額 & ","
        gSL_Select = gSL_Select & "最終入庫日 = '" & RTrim(.最終入庫日) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZKWUpdate = True
End Function

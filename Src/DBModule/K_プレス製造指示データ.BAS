Attribute VB_Name = "K_プレス製造指示データ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   K_プレス製造指示データ
'*
'*   モジュールID  : K_プレス製造指示データ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2010/04/07  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_K_プレス製造指示データ
        優先度                As String
        プレス部品番号        As String
        材料サイズ            As String
        打切M                 As String
        加工数1               As String
        加工数2               As String
        材料在庫量            As String
        材料使用量            As String
        共用材                As String
        原単位                As String
        BL種類                As String
        親部品番号            As String
        内示当月              As String
        内示次月              As String
        プレス工程数          As String
        プレス種              As String
        空き                  As String
        件数          As Long
    End Type

    Public KKP As T_K_プレス製造指示データ

    Public KKP_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKKP(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKKP	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KKP
            Select Case adoCol.Name
                Case "優先度":                If Not IsNull(adoCol.Value) Then .優先度               = adoCol.Value
                Case "プレス部品番号":        If Not IsNull(adoCol.Value) Then .プレス部品番号       = adoCol.Value
                Case "材料サイズ":            If Not IsNull(adoCol.Value) Then .材料サイズ           = adoCol.Value
                Case "打切M":                 If Not IsNull(adoCol.Value) Then .打切M                = adoCol.Value
                Case "加工数1":               If Not IsNull(adoCol.Value) Then .加工数1              = adoCol.Value
                Case "加工数2":               If Not IsNull(adoCol.Value) Then .加工数2              = adoCol.Value
                Case "材料在庫量":            If Not IsNull(adoCol.Value) Then .材料在庫量           = adoCol.Value
                Case "材料使用量":            If Not IsNull(adoCol.Value) Then .材料使用量           = adoCol.Value
                Case "共用材":                If Not IsNull(adoCol.Value) Then .共用材               = adoCol.Value
                Case "原単位":                If Not IsNull(adoCol.Value) Then .原単位               = adoCol.Value
                Case "BL種類":                If Not IsNull(adoCol.Value) Then .BL種類               = adoCol.Value
                Case "親部品番号":            If Not IsNull(adoCol.Value) Then .親部品番号           = adoCol.Value
                Case "内示当月":              If Not IsNull(adoCol.Value) Then .内示当月             = adoCol.Value
                Case "内示次月":              If Not IsNull(adoCol.Value) Then .内示次月             = adoCol.Value
                Case "プレス工程数":          If Not IsNull(adoCol.Value) Then .プレス工程数         = adoCol.Value
                Case "プレス種":              If Not IsNull(adoCol.Value) Then .プレス種             = adoCol.Value
                Case "空き":                  If Not IsNull(adoCol.Value) Then .空き                 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KKPRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KKPRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From K_プレス製造指示データ"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KKPRead = True
           Call ItemsFromKKP(DBNo)
	   KKP_RDSTS = True
        else
	   KKP_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KKPClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KKPReadNext(Optional DBNo As Integer = 0)
    If Not KKP_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KKP_RDSTS = False
        Else
           Call ItemsFromKKP(DBNo)
           KKP_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KKPClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KKP_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KKP )
'*----------------------------------------------*'
Public Sub ItemsClearKKP()
    With KKP
	.優先度 = ""
	.プレス部品番号 = ""
	.材料サイズ = ""
	.打切M = ""
	.加工数1 = ""
	.加工数2 = ""
	.材料在庫量 = ""
	.材料使用量 = ""
	.共用材 = ""
	.原単位 = ""
	.BL種類 = ""
	.親部品番号 = ""
	.内示当月 = ""
	.内示次月 = ""
	.プレス工程数 = ""
	.プレス種 = ""
	.空き = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KKPInsert(Optional DBNo As Integer) As Boolean
    KKPInsert = False
    With KKP
        gSL_Select = "Insert Into K_プレス製造指示データ  ( "
        gSL_Select = gSL_Select & "優先度,プレス部品番号,材料サイズ,"
        gSL_Select = gSL_Select & "打切M,加工数1,加工数2,"
        gSL_Select = gSL_Select & "材料在庫量,材料使用量,共用材,"
        gSL_Select = gSL_Select & "原単位,BL種類,親部品番号,"
        gSL_Select = gSL_Select & "内示当月,内示次月,プレス工程数,"
        gSL_Select = gSL_Select & "プレス種,空き"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.優先度) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.プレス部品番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.材料サイズ) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.打切M) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.加工数1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.加工数2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.材料在庫量) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.材料使用量) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.共用材) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.原単位) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.BL種類) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.親部品番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.内示当月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.内示次月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.プレス工程数) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.プレス種) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.空き) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KKPInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KKPUpdate(Optional DBNo As Integer) As Boolean
    KKPUpdate = False	
    With KKP
        gSL_Select = "Update K_プレス製造指示データ  Set "
        gSL_Select = gSL_Select & "優先度 = '" & RTrim(.優先度) & "',"
        gSL_Select = gSL_Select & "プレス部品番号 = '" & RTrim(.プレス部品番号) & "',"
        gSL_Select = gSL_Select & "材料サイズ = '" & RTrim(.材料サイズ) & "',"
        gSL_Select = gSL_Select & "打切M = '" & RTrim(.打切M) & "',"
        gSL_Select = gSL_Select & "加工数1 = '" & RTrim(.加工数1) & "',"
        gSL_Select = gSL_Select & "加工数2 = '" & RTrim(.加工数2) & "',"
        gSL_Select = gSL_Select & "材料在庫量 = '" & RTrim(.材料在庫量) & "',"
        gSL_Select = gSL_Select & "材料使用量 = '" & RTrim(.材料使用量) & "',"
        gSL_Select = gSL_Select & "共用材 = '" & RTrim(.共用材) & "',"
        gSL_Select = gSL_Select & "原単位 = '" & RTrim(.原単位) & "',"
        gSL_Select = gSL_Select & "BL種類 = '" & RTrim(.BL種類) & "',"
        gSL_Select = gSL_Select & "親部品番号 = '" & RTrim(.親部品番号) & "',"
        gSL_Select = gSL_Select & "内示当月 = '" & RTrim(.内示当月) & "',"
        gSL_Select = gSL_Select & "内示次月 = '" & RTrim(.内示次月) & "',"
        gSL_Select = gSL_Select & "プレス工程数 = '" & RTrim(.プレス工程数) & "',"
        gSL_Select = gSL_Select & "プレス種 = '" & RTrim(.プレス種) & "',"
        gSL_Select = gSL_Select & "空き = '" & RTrim(.空き) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KKPUpdate = True
End Function

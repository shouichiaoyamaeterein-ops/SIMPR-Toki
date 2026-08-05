Attribute VB_Name = "K_在庫マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   K_在庫マスタ
'*
'*   モジュールID  : K_在庫マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2010/04/22  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_K_在庫マスタ
        品目コード            As String
        在庫場所              As String
        前月末在庫数          As String
        入庫数                As String
        入庫返品数            As String
        入庫在庫調整数        As String
        出庫数                As String
        出庫返品数            As String
        廃却数                As String
        出庫在庫調整数        As String
        現在在庫数            As String
        在庫金額              As String
        発注残                As String
        引当残                As String
        有効残                As String
        最終入庫日            As String
        前月末実質在庫        As String
        実質在庫数            As String
        空き                  As String
        件数          As Long
    End Type

    Public KKZ As T_K_在庫マスタ

    Public KKZ_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKKZ(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKKZ	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KKZ
            Select Case adoCol.Name
                Case "品目コード":            If Not IsNull(adoCol.Value) Then .品目コード           = adoCol.Value
                Case "在庫場所":              If Not IsNull(adoCol.Value) Then .在庫場所             = adoCol.Value
                Case "前月末在庫数":          If Not IsNull(adoCol.Value) Then .前月末在庫数         = adoCol.Value
                Case "入庫数":                If Not IsNull(adoCol.Value) Then .入庫数               = adoCol.Value
                Case "入庫返品数":            If Not IsNull(adoCol.Value) Then .入庫返品数           = adoCol.Value
                Case "入庫在庫調整数":        If Not IsNull(adoCol.Value) Then .入庫在庫調整数       = adoCol.Value
                Case "出庫数":                If Not IsNull(adoCol.Value) Then .出庫数               = adoCol.Value
                Case "出庫返品数":            If Not IsNull(adoCol.Value) Then .出庫返品数           = adoCol.Value
                Case "廃却数":                If Not IsNull(adoCol.Value) Then .廃却数               = adoCol.Value
                Case "出庫在庫調整数":        If Not IsNull(adoCol.Value) Then .出庫在庫調整数       = adoCol.Value
                Case "現在在庫数":            If Not IsNull(adoCol.Value) Then .現在在庫数           = adoCol.Value
                Case "在庫金額":              If Not IsNull(adoCol.Value) Then .在庫金額             = adoCol.Value
                Case "発注残":                If Not IsNull(adoCol.Value) Then .発注残               = adoCol.Value
                Case "引当残":                If Not IsNull(adoCol.Value) Then .引当残               = adoCol.Value
                Case "有効残":                If Not IsNull(adoCol.Value) Then .有効残               = adoCol.Value
                Case "最終入庫日":            If Not IsNull(adoCol.Value) Then .最終入庫日           = adoCol.Value
                Case "前月末実質在庫":        If Not IsNull(adoCol.Value) Then .前月末実質在庫       = adoCol.Value
                Case "実質在庫数":            If Not IsNull(adoCol.Value) Then .実質在庫数           = adoCol.Value
                Case "空き":                  If Not IsNull(adoCol.Value) Then .空き                 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KKZRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KKZRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From K_在庫マスタ"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KKZRead = True
           Call ItemsFromKKZ(DBNo)
	   KKZ_RDSTS = True
        else
	   KKZ_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KKZClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KKZReadNext(Optional DBNo As Integer = 0)
    If Not KKZ_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KKZ_RDSTS = False
        Else
           Call ItemsFromKKZ(DBNo)
           KKZ_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KKZClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KKZ_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KKZ )
'*----------------------------------------------*'
Public Sub ItemsClearKKZ()
    With KKZ
	.品目コード = ""
	.在庫場所 = ""
	.前月末在庫数 = ""
	.入庫数 = ""
	.入庫返品数 = ""
	.入庫在庫調整数 = ""
	.出庫数 = ""
	.出庫返品数 = ""
	.廃却数 = ""
	.出庫在庫調整数 = ""
	.現在在庫数 = ""
	.在庫金額 = ""
	.発注残 = ""
	.引当残 = ""
	.有効残 = ""
	.最終入庫日 = ""
	.前月末実質在庫 = ""
	.実質在庫数 = ""
	.空き = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KKZInsert(Optional DBNo As Integer) As Boolean
    KKZInsert = False
    With KKZ
        gSL_Select = "Insert Into K_在庫マスタ  ( "
        gSL_Select = gSL_Select & "品目コード,在庫場所,前月末在庫数,"
        gSL_Select = gSL_Select & "入庫数,入庫返品数,入庫在庫調整数,"
        gSL_Select = gSL_Select & "出庫数,出庫返品数,廃却数,"
        gSL_Select = gSL_Select & "出庫在庫調整数,現在在庫数,在庫金額,"
        gSL_Select = gSL_Select & "発注残,引当残,有効残,"
        gSL_Select = gSL_Select & "最終入庫日,前月末実質在庫,実質在庫数,"
        gSL_Select = gSL_Select & "空き"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.品目コード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.在庫場所) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.前月末在庫数) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.入庫数) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.入庫返品数) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.入庫在庫調整数) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.出庫数) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.出庫返品数) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.廃却数) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.出庫在庫調整数) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.現在在庫数) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.在庫金額) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注残) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.引当残) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.有効残) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.最終入庫日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.前月末実質在庫) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.実質在庫数) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.空き) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KKZInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KKZUpdate(Optional DBNo As Integer) As Boolean
    KKZUpdate = False	
    With KKZ
        gSL_Select = "Update K_在庫マスタ  Set "
        gSL_Select = gSL_Select & "品目コード = '" & RTrim(.品目コード) & "',"
        gSL_Select = gSL_Select & "在庫場所 = '" & RTrim(.在庫場所) & "',"
        gSL_Select = gSL_Select & "前月末在庫数 = '" & RTrim(.前月末在庫数) & "',"
        gSL_Select = gSL_Select & "入庫数 = '" & RTrim(.入庫数) & "',"
        gSL_Select = gSL_Select & "入庫返品数 = '" & RTrim(.入庫返品数) & "',"
        gSL_Select = gSL_Select & "入庫在庫調整数 = '" & RTrim(.入庫在庫調整数) & "',"
        gSL_Select = gSL_Select & "出庫数 = '" & RTrim(.出庫数) & "',"
        gSL_Select = gSL_Select & "出庫返品数 = '" & RTrim(.出庫返品数) & "',"
        gSL_Select = gSL_Select & "廃却数 = '" & RTrim(.廃却数) & "',"
        gSL_Select = gSL_Select & "出庫在庫調整数 = '" & RTrim(.出庫在庫調整数) & "',"
        gSL_Select = gSL_Select & "現在在庫数 = '" & RTrim(.現在在庫数) & "',"
        gSL_Select = gSL_Select & "在庫金額 = '" & RTrim(.在庫金額) & "',"
        gSL_Select = gSL_Select & "発注残 = '" & RTrim(.発注残) & "',"
        gSL_Select = gSL_Select & "引当残 = '" & RTrim(.引当残) & "',"
        gSL_Select = gSL_Select & "有効残 = '" & RTrim(.有効残) & "',"
        gSL_Select = gSL_Select & "最終入庫日 = '" & RTrim(.最終入庫日) & "',"
        gSL_Select = gSL_Select & "前月末実質在庫 = '" & RTrim(.前月末実質在庫) & "',"
        gSL_Select = gSL_Select & "実質在庫数 = '" & RTrim(.実質在庫数) & "',"
        gSL_Select = gSL_Select & "空き = '" & RTrim(.空き) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KKZUpdate = True
End Function

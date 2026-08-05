Attribute VB_Name = "K_構成マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   K_構成マスタ
'*
'*   モジュールID  : K_構成マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2009/08/26  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_K_構成マスタ
        上位品目コード        As String
        下位品目コード        As String
        原単位                As String
        不良加算率            As String
        支給区分              As String
        直送区分              As String
        上位品目ダミー        As String
        使用開始日            As String
        使用終了日            As String
        下位品目区分          As String
        空き                  As String
        件数          As Long
    End Type

    Public KKK As T_K_構成マスタ

    Public KKK_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKKK(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKKK	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KKK
            Select Case adoCol.Name
                Case "上位品目コード":        If Not IsNull(adoCol.Value) Then .上位品目コード       = adoCol.Value
                Case "下位品目コード":        If Not IsNull(adoCol.Value) Then .下位品目コード       = adoCol.Value
                Case "原単位":                If Not IsNull(adoCol.Value) Then .原単位               = adoCol.Value
                Case "不良加算率":            If Not IsNull(adoCol.Value) Then .不良加算率           = adoCol.Value
                Case "支給区分":              If Not IsNull(adoCol.Value) Then .支給区分             = adoCol.Value
                Case "直送区分":              If Not IsNull(adoCol.Value) Then .直送区分             = adoCol.Value
                Case "上位品目ダミー":        If Not IsNull(adoCol.Value) Then .上位品目ダミー       = adoCol.Value
                Case "使用開始日":            If Not IsNull(adoCol.Value) Then .使用開始日           = adoCol.Value
                Case "使用終了日":            If Not IsNull(adoCol.Value) Then .使用終了日           = adoCol.Value
                Case "下位品目区分":          If Not IsNull(adoCol.Value) Then .下位品目区分         = adoCol.Value
                Case "空き":                  If Not IsNull(adoCol.Value) Then .空き                 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KKKRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KKKRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From K_構成マスタ"
        RTxtSql = RTxtSql &" Where 上位品目コード = '" & RTrim(KKK.上位品目コード) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KKKRead = True
           Call ItemsFromKKK(DBNo)
	   KKK_RDSTS = True
        else
	   KKK_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KKKClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KKKReadNext(Optional DBNo As Integer = 0)
    If Not KKK_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KKK_RDSTS = False
        Else
           Call ItemsFromKKK(DBNo)
           KKK_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KKKClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KKK_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KKK )
'*----------------------------------------------*'
Public Sub ItemsClearKKK()
    With KKK
	.上位品目コード = ""
	.下位品目コード = ""
	.原単位 = ""
	.不良加算率 = ""
	.支給区分 = ""
	.直送区分 = ""
	.上位品目ダミー = ""
	.使用開始日 = ""
	.使用終了日 = ""
	.下位品目区分 = ""
	.空き = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KKKInsert(Optional DBNo As Integer) As Boolean
    KKKInsert = False
    With KKK
        gSL_Select = "Insert Into K_構成マスタ  ( "
        gSL_Select = gSL_Select & "上位品目コード,下位品目コード,原単位,"
        gSL_Select = gSL_Select & "不良加算率,支給区分,直送区分,"
        gSL_Select = gSL_Select & "上位品目ダミー,使用開始日,使用終了日,"
        gSL_Select = gSL_Select & "下位品目区分,空き"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.上位品目コード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.下位品目コード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.原単位) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.不良加算率) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.直送区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.上位品目ダミー) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.使用開始日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.使用終了日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.下位品目区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.空き) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KKKInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KKKUpdate(Optional DBNo As Integer) As Boolean
    KKKUpdate = False	
    With KKK
        gSL_Select = "Update K_構成マスタ  Set "
        gSL_Select = gSL_Select & "上位品目コード = '" & RTrim(.上位品目コード) & "',"
        gSL_Select = gSL_Select & "下位品目コード = '" & RTrim(.下位品目コード) & "',"
        gSL_Select = gSL_Select & "原単位 = '" & RTrim(.原単位) & "',"
        gSL_Select = gSL_Select & "不良加算率 = '" & RTrim(.不良加算率) & "',"
        gSL_Select = gSL_Select & "支給区分 = '" & RTrim(.支給区分) & "',"
        gSL_Select = gSL_Select & "直送区分 = '" & RTrim(.直送区分) & "',"
        gSL_Select = gSL_Select & "上位品目ダミー = '" & RTrim(.上位品目ダミー) & "',"
        gSL_Select = gSL_Select & "使用開始日 = '" & RTrim(.使用開始日) & "',"
        gSL_Select = gSL_Select & "使用終了日 = '" & RTrim(.使用終了日) & "',"
        gSL_Select = gSL_Select & "下位品目区分 = '" & RTrim(.下位品目区分) & "',"
        gSL_Select = gSL_Select & "空き = '" & RTrim(.空き) & "'"
        gSL_Select = gSL_Select & " Where 上位品目コード = '" & RTrim(.上位品目コード) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KKKUpdate = True
End Function

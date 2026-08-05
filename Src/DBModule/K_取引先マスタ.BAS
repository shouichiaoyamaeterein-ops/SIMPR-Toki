Attribute VB_Name = "K_取引先マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   K_取引先マスタ
'*
'*   モジュールID  : K_取引先マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2009/08/26  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_K_取引先マスタ
        取引先コード          As String
        取引先区分            As String
        取引先名              As String
        取引先名略称          As String
        取引先名カナ          As String
        住所1                 As String
        住所2                 As String
        郵便番号              As String
        電話番号              As String
        注文書形式区分        As String
        締日1                 As String
        締日2                 As String
        締日3                 As String
        集金日                As String
        与信限度額            As String
        請求先コード          As String
        件数          As Long
    End Type

    Public KKT As T_K_取引先マスタ

    Public KKT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKKT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKKT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KKT
            Select Case adoCol.Name
                Case "取引先コード":          If Not IsNull(adoCol.Value) Then .取引先コード         = adoCol.Value
                Case "取引先区分":            If Not IsNull(adoCol.Value) Then .取引先区分           = adoCol.Value
                Case "取引先名":              If Not IsNull(adoCol.Value) Then .取引先名             = adoCol.Value
                Case "取引先名略称":          If Not IsNull(adoCol.Value) Then .取引先名略称         = adoCol.Value
                Case "取引先名カナ":          If Not IsNull(adoCol.Value) Then .取引先名カナ         = adoCol.Value
                Case "住所1":                 If Not IsNull(adoCol.Value) Then .住所1                = adoCol.Value
                Case "住所2":                 If Not IsNull(adoCol.Value) Then .住所2                = adoCol.Value
                Case "郵便番号":              If Not IsNull(adoCol.Value) Then .郵便番号             = adoCol.Value
                Case "電話番号":              If Not IsNull(adoCol.Value) Then .電話番号             = adoCol.Value
                Case "注文書形式区分":        If Not IsNull(adoCol.Value) Then .注文書形式区分       = adoCol.Value
                Case "締日1":                 If Not IsNull(adoCol.Value) Then .締日1                = adoCol.Value
                Case "締日2":                 If Not IsNull(adoCol.Value) Then .締日2                = adoCol.Value
                Case "締日3":                 If Not IsNull(adoCol.Value) Then .締日3                = adoCol.Value
                Case "集金日":                If Not IsNull(adoCol.Value) Then .集金日               = adoCol.Value
                Case "与信限度額":            If Not IsNull(adoCol.Value) Then .与信限度額           = adoCol.Value
                Case "請求先コード":          If Not IsNull(adoCol.Value) Then .請求先コード         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KKTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KKTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From K_取引先マスタ"
        RTxtSql = RTxtSql &" Where 取引先コード = '" & RTrim(KKT.取引先コード) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KKTRead = True
           Call ItemsFromKKT(DBNo)
	   KKT_RDSTS = True
        else
	   KKT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KKTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KKTReadNext(Optional DBNo As Integer = 0)
    If Not KKT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KKT_RDSTS = False
        Else
           Call ItemsFromKKT(DBNo)
           KKT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KKTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KKT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KKT )
'*----------------------------------------------*'
Public Sub ItemsClearKKT()
    With KKT
	.取引先コード = ""
	.取引先区分 = ""
	.取引先名 = ""
	.取引先名略称 = ""
	.取引先名カナ = ""
	.住所1 = ""
	.住所2 = ""
	.郵便番号 = ""
	.電話番号 = ""
	.注文書形式区分 = ""
	.締日1 = ""
	.締日2 = ""
	.締日3 = ""
	.集金日 = ""
	.与信限度額 = ""
	.請求先コード = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KKTInsert(Optional DBNo As Integer) As Boolean
    KKTInsert = False
    With KKT
        gSL_Select = "Insert Into K_取引先マスタ  ( "
        gSL_Select = gSL_Select & "取引先コード,取引先区分,取引先名,"
        gSL_Select = gSL_Select & "取引先名略称,取引先名カナ,住所1,"
        gSL_Select = gSL_Select & "住所2,郵便番号,電話番号,"
        gSL_Select = gSL_Select & "注文書形式区分,締日1,締日2,"
        gSL_Select = gSL_Select & "締日3,集金日,与信限度額,"
        gSL_Select = gSL_Select & "請求先コード"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.取引先コード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先名略称) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先名カナ) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.住所1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.住所2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.郵便番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.電話番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.注文書形式区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.締日1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.締日2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.締日3) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.集金日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.与信限度額) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.請求先コード) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KKTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KKTUpdate(Optional DBNo As Integer) As Boolean
    KKTUpdate = False	
    With KKT
        gSL_Select = "Update K_取引先マスタ  Set "
        gSL_Select = gSL_Select & "取引先コード = '" & RTrim(.取引先コード) & "',"
        gSL_Select = gSL_Select & "取引先区分 = '" & RTrim(.取引先区分) & "',"
        gSL_Select = gSL_Select & "取引先名 = '" & RTrim(.取引先名) & "',"
        gSL_Select = gSL_Select & "取引先名略称 = '" & RTrim(.取引先名略称) & "',"
        gSL_Select = gSL_Select & "取引先名カナ = '" & RTrim(.取引先名カナ) & "',"
        gSL_Select = gSL_Select & "住所1 = '" & RTrim(.住所1) & "',"
        gSL_Select = gSL_Select & "住所2 = '" & RTrim(.住所2) & "',"
        gSL_Select = gSL_Select & "郵便番号 = '" & RTrim(.郵便番号) & "',"
        gSL_Select = gSL_Select & "電話番号 = '" & RTrim(.電話番号) & "',"
        gSL_Select = gSL_Select & "注文書形式区分 = '" & RTrim(.注文書形式区分) & "',"
        gSL_Select = gSL_Select & "締日1 = '" & RTrim(.締日1) & "',"
        gSL_Select = gSL_Select & "締日2 = '" & RTrim(.締日2) & "',"
        gSL_Select = gSL_Select & "締日3 = '" & RTrim(.締日3) & "',"
        gSL_Select = gSL_Select & "集金日 = '" & RTrim(.集金日) & "',"
        gSL_Select = gSL_Select & "与信限度額 = '" & RTrim(.与信限度額) & "',"
        gSL_Select = gSL_Select & "請求先コード = '" & RTrim(.請求先コード) & "'"
        gSL_Select = gSL_Select & " Where 取引先コード = '" & RTrim(.取引先コード) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KKTUpdate = True
End Function

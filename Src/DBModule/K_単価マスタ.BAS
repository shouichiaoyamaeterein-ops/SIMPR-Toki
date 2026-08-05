Attribute VB_Name = "K_単価マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   K_単価マスタ
'*
'*   モジュールID  : K_単価マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2009/11/13  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_K_単価マスタ
        品目コード    As String
        単価区分      As String
        取引先コード  As String
        新単価        As String
        旧単価        As String
        変更日        As String
        空き          As String
        件数          As Long
    End Type

    Public KTA As T_K_単価マスタ

    Public KTA_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKTA(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKTA	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KTA
            Select Case adoCol.Name
                Case "品目コード":    If Not IsNull(adoCol.Value) Then .品目コード   = adoCol.Value
                Case "単価区分":      If Not IsNull(adoCol.Value) Then .単価区分     = adoCol.Value
                Case "取引先コード":  If Not IsNull(adoCol.Value) Then .取引先コード = adoCol.Value
                Case "新単価":        If Not IsNull(adoCol.Value) Then .新単価       = adoCol.Value
                Case "旧単価":        If Not IsNull(adoCol.Value) Then .旧単価       = adoCol.Value
                Case "変更日":        If Not IsNull(adoCol.Value) Then .変更日       = adoCol.Value
                Case "空き":          If Not IsNull(adoCol.Value) Then .空き         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KTARead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KTARead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From K_単価マスタ"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KTARead = True
           Call ItemsFromKTA(DBNo)
	   KTA_RDSTS = True
        else
	   KTA_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KTAClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KTAReadNext(Optional DBNo As Integer = 0)
    If Not KTA_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KTA_RDSTS = False
        Else
           Call ItemsFromKTA(DBNo)
           KTA_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KTAClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KTA_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KTA )
'*----------------------------------------------*'
Public Sub ItemsClearKTA()
    With KTA
	.品目コード = ""
	.単価区分 = ""
	.取引先コード = ""
	.新単価 = ""
	.旧単価 = ""
	.変更日 = ""
	.空き = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KTAInsert(Optional DBNo As Integer) As Boolean
    KTAInsert = False
    With KTA
        gSL_Select = "Insert Into K_単価マスタ  ( "
        gSL_Select = gSL_Select & "品目コード,単価区分,取引先コード,"
        gSL_Select = gSL_Select & "新単価,旧単価,変更日,"
        gSL_Select = gSL_Select & "空き"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.品目コード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.単価区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先コード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.新単価) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.旧単価) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.変更日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.空き) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KTAInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KTAUpdate(Optional DBNo As Integer) As Boolean
    KTAUpdate = False	
    With KTA
        gSL_Select = "Update K_単価マスタ  Set "
        gSL_Select = gSL_Select & "品目コード = '" & RTrim(.品目コード) & "',"
        gSL_Select = gSL_Select & "単価区分 = '" & RTrim(.単価区分) & "',"
        gSL_Select = gSL_Select & "取引先コード = '" & RTrim(.取引先コード) & "',"
        gSL_Select = gSL_Select & "新単価 = '" & RTrim(.新単価) & "',"
        gSL_Select = gSL_Select & "旧単価 = '" & RTrim(.旧単価) & "',"
        gSL_Select = gSL_Select & "変更日 = '" & RTrim(.変更日) & "',"
        gSL_Select = gSL_Select & "空き = '" & RTrim(.空き) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KTAUpdate = True
End Function

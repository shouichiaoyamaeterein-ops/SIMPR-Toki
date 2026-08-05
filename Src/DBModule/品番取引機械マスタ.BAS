Attribute VB_Name = "品番取引機械マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   品番取引機械マスタ
'*
'*   モジュールID  : 品番取引機械マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/12/13  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_品番取引機械マスタ
        品番      As String
        取引先    As String
        受入      As String
        機械      As String
        基本      As String
        機械備考  As String
        削除区分  As String
'        作成日    As Date
        作成者    As String
'        作成端末  As String
'        更新日    As Date
        更新者    As String
'        更新端末  As String
        件数          As Long
    End Type

    Public HKM As T_品番取引機械マスタ

    Public HKM_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHKM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHKM	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With HKM
            Select Case adoCol.Name
                Case "品番":      If Not IsNull(adoCol.Value) Then .品番     = adoCol.Value
                Case "取引先":    If Not IsNull(adoCol.Value) Then .取引先   = adoCol.Value
                Case "受入":      If Not IsNull(adoCol.Value) Then .受入     = adoCol.Value
                Case "機械":      If Not IsNull(adoCol.Value) Then .機械     = adoCol.Value
                Case "基本":      If Not IsNull(adoCol.Value) Then .基本     = adoCol.Value
                Case "機械備考":  If Not IsNull(adoCol.Value) Then .機械備考 = adoCol.Value
                Case "削除区分":  If Not IsNull(adoCol.Value) Then .削除区分 = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
                Case "作成者":    If Not IsNull(adoCol.Value) Then .作成者   = adoCol.Value
'                Case "作成端末":  If Not IsNull(adoCol.Value) Then .作成端末 = adoCol.Value
'                Case "更新日":    If Not IsNull(adoCol.Value) Then .更新日   = adoCol.Value
                Case "更新者":    If Not IsNull(adoCol.Value) Then .更新者   = adoCol.Value
'                Case "更新端末":  If Not IsNull(adoCol.Value) Then .更新端末 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function HKMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HKMRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 品番取引機械マスタ"
        RTxtSql = RTxtSql &" Where 品番 = '" & RTrim(HKM.品番) & "'"
        RTxtSql = RTxtSql &" And   取引先 = '" & RTrim(HKM.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(HKM.受入) & "'"
        RTxtSql = RTxtSql &" And   機械 = '" & RTrim(HKM.機械) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HKMRead = True
           Call ItemsFromHKM(DBNo)
	   HKM_RDSTS = True
        else
	   HKM_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call HKMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HKMReadNext(Optional DBNo As Integer = 0)
    If Not HKM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HKM_RDSTS = False
        Else
           Call ItemsFromHKM(DBNo)
           HKM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HKMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HKM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HKM )
'*----------------------------------------------*'
Public Sub ItemsClearHKM()
    With HKM
	.品番 = ""
	.取引先 = ""
	.受入 = ""
	.機械 = ""
	.基本 = ""
	.機械備考 = ""
	.削除区分 = ""
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
Public Function HKMInsert(Optional DBNo As Integer) As Boolean
    HKMInsert = True
    With HKM
        gSL_Select = "Insert Into 品番取引機械マスタ  ( "
        gSL_Select = gSL_Select & "品番,取引先,受入,"
        gSL_Select = gSL_Select & "機械,基本,機械備考,"
        gSL_Select = gSL_Select & "削除区分,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.機械) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.基本) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.機械備考) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.削除区分) & "',"		
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
    HKMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function HKMUpdate(Optional DBNo As Integer) As Boolean
    HKMUpdate = False	
    With HKM
        gSL_Select = "Update 品番取引機械マスタ  Set "
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "機械 = '" & RTrim(.機械) & "',"
        gSL_Select = gSL_Select & "基本 = '" & RTrim(.基本) & "',"
        gSL_Select = gSL_Select & "機械備考 = '" & RTrim(.機械備考) & "',"
        gSL_Select = gSL_Select & "削除区分 = '" & RTrim(.削除区分) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 機械 = '" & RTrim(.機械) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HKMUpdate = True
End Function

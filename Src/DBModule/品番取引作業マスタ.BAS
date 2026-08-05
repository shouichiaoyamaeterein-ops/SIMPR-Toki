Attribute VB_Name = "品番取引作業マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   品番取引作業マスタ
'*
'*   モジュールID  : 品番取引作業マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/03/02  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_品番取引作業マスタ
        品番      As String
        取引先    As String
        受入      As String
        作業順    As Byte
        作業      As String
        備考      As String
        作業区分  As String
        削除区分  As String
'        作成日    As Date
        作成者    As String
'        作成端末  As String
'        更新日    As Date
        更新者    As String
'        更新端末  As String
        件数          As Long
'   項目追加(自動)
	作業名        As String
    End Type

    Public HSM As T_品番取引作業マスタ

    Public HSM_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHSM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHSM	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With HSM
            Select Case adoCol.Name
                Case "品番":      If Not IsNull(adoCol.Value) Then .品番     = adoCol.Value
                Case "取引先":    If Not IsNull(adoCol.Value) Then .取引先   = adoCol.Value
                Case "受入":      If Not IsNull(adoCol.Value) Then .受入     = adoCol.Value
                Case "作業順":    If Not IsNull(adoCol.Value) Then .作業順   = adoCol.Value
                Case "作業":      If Not IsNull(adoCol.Value) Then .作業     = adoCol.Value
                Case "備考":      If Not IsNull(adoCol.Value) Then .備考     = adoCol.Value
                Case "作業区分":  If Not IsNull(adoCol.Value) Then .作業区分 = adoCol.Value
                Case "削除区分":  If Not IsNull(adoCol.Value) Then .削除区分 = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
                Case "作成者":    If Not IsNull(adoCol.Value) Then .作成者   = adoCol.Value
'                Case "作成端末":  If Not IsNull(adoCol.Value) Then .作成端末 = adoCol.Value
'                Case "更新日":    If Not IsNull(adoCol.Value) Then .更新日   = adoCol.Value
                Case "更新者":    If Not IsNull(adoCol.Value) Then .更新者   = adoCol.Value
'                Case "更新端末":  If Not IsNull(adoCol.Value) Then .更新端末 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "作業名":	If Not IsNull(adoCol.Value) Then .作業名	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function HSMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HSMRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 品番取引作業マスタ"
        RTxtSql = RTxtSql &" Where 品番 = '" & RTrim(HSM.品番) & "'"
        RTxtSql = RTxtSql &" And   取引先 = '" & RTrim(HSM.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(HSM.受入) & "'"
        RTxtSql = RTxtSql &" And   作業順 = " & HSM.作業順 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HSMRead = True
           Call ItemsFromHSM(DBNo)
	   HSM_RDSTS = True
        else
	   HSM_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call HSMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HSMReadNext(Optional DBNo As Integer = 0)
    If Not HSM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HSM_RDSTS = False
        Else
           Call ItemsFromHSM(DBNo)
           HSM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HSMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HSM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HSM )
'*----------------------------------------------*'
Public Sub ItemsClearHSM()
    With HSM
	.品番 = ""
	.取引先 = ""
	.受入 = ""
	.作業順 = 0
	.作業 = ""
	.備考 = ""
	.作業区分 = ""
	.削除区分 = ""
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
'	.更新日 = 0
	.更新者 = ""
'	.更新端末 = ""
        .件数 = 0
'   項目追加(自動)
	.作業名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function HSMInsert(Optional DBNo As Integer) As Boolean
    HSMInsert = True
    With HSM
        gSL_Select = "Insert Into 品番取引作業マスタ  ( "
        gSL_Select = gSL_Select & "品番,取引先,受入,"
        gSL_Select = gSL_Select & "作業順,作業,備考,"
        gSL_Select = gSL_Select & "作業区分,削除区分,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "" & .作業順 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.作業) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.備考) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.作業区分) & "',"		
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
    HSMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function HSMUpdate(Optional DBNo As Integer) As Boolean
    HSMUpdate = False	
    With HSM
        gSL_Select = "Update 品番取引作業マスタ  Set "
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "作業順 = " & .作業順 & ","
        gSL_Select = gSL_Select & "作業 = '" & RTrim(.作業) & "',"
        gSL_Select = gSL_Select & "備考 = '" & RTrim(.備考) & "',"
        gSL_Select = gSL_Select & "作業区分 = '" & RTrim(.作業区分) & "',"
        gSL_Select = gSL_Select & "削除区分 = '" & RTrim(.削除区分) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 作業順 = " & .作業順 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HSMUpdate = True
End Function

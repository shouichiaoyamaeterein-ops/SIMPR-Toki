Attribute VB_Name = "在庫マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   在庫マスタ
'*
'*   モジュールID  : 在庫マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/02/12  By Cis
'*   作成日  :  2010/01/07  By Cis(HP-DX7500)
'*   作成日  :  2010/01/08  By Cis(HP-DX7500)
'*   作成日  :  2010/01/21  By Cis(HP-DX7500)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_在庫マスタ
        年月度    As String
        品番      As String
        取引先    As String
        受入      As String
        工順      As Byte
        在庫数    As Currency
        実棚日    As String
        変更事由  As String
        作成区分  As Byte
'        作成日    As Date
        作成者    As String
'        作成端末  As String
        件数          As Long
'   項目追加(自動)
	在庫場所       As String
	品目区分       As String
	品名           As String
	仕様           As String
	単価           As Currency
	金額           As Currency
	最終出庫日     As String
	自ガス         As Byte
    End Type

    Public ZKM As T_在庫マスタ

    Public ZKM_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZKM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZKM	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZKM
            Select Case adoCol.Name
                Case "年月度":    If Not IsNull(adoCol.Value) Then .年月度   = adoCol.Value
                Case "品番":      If Not IsNull(adoCol.Value) Then .品番     = adoCol.Value
                Case "取引先":    If Not IsNull(adoCol.Value) Then .取引先   = adoCol.Value
                Case "受入":      If Not IsNull(adoCol.Value) Then .受入     = adoCol.Value
                Case "工順":      If Not IsNull(adoCol.Value) Then .工順     = adoCol.Value
                Case "在庫数":    If Not IsNull(adoCol.Value) Then .在庫数   = adoCol.Value
                Case "実棚日":    If Not IsNull(adoCol.Value) Then .実棚日   = adoCol.Value
                Case "変更事由":  If Not IsNull(adoCol.Value) Then .変更事由 = adoCol.Value
                Case "作成区分":  If Not IsNull(adoCol.Value) Then .作成区分 = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
                Case "作成者":    If Not IsNull(adoCol.Value) Then .作成者   = adoCol.Value
'                Case "作成端末":  If Not IsNull(adoCol.Value) Then .作成端末 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "在庫場所":	If Not IsNull(adoCol.Value) Then .在庫場所	= adoCol.Value
		Case "品目区分":	If Not IsNull(adoCol.Value) Then .品目区分	= adoCol.Value
		Case "品名":	If Not IsNull(adoCol.Value) Then .品名	= adoCol.Value
		Case "仕様":	If Not IsNull(adoCol.Value) Then .仕様	= adoCol.Value
		Case "単価":	If Not IsNull(adoCol.Value) Then .単価	= adoCol.Value
		Case "金額":	If Not IsNull(adoCol.Value) Then .金額	= adoCol.Value
		Case "最終出庫日":	If Not IsNull(adoCol.Value) Then .最終出庫日	= adoCol.Value
		Case "自ガス":	If Not IsNull(adoCol.Value) Then .自ガス	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZKMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZKMRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 在庫マスタ"
        RTxtSql = RTxtSql &" Where 年月度 = '" & RTrim(ZKM.年月度) & "'"
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(ZKM.品番) & "'"
        RTxtSql = RTxtSql &" And   取引先 = '" & RTrim(ZKM.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(ZKM.受入) & "'"
        RTxtSql = RTxtSql &" And   工順 = " & ZKM.工順 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZKMRead = True
           Call ItemsFromZKM(DBNo)
	   ZKM_RDSTS = True
        else
	   ZKM_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZKMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZKMReadNext(Optional DBNo As Integer = 0)
    If Not ZKM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZKM_RDSTS = False
        Else
           Call ItemsFromZKM(DBNo)
           ZKM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZKMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZKM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZKM )
'*----------------------------------------------*'
Public Sub ItemsClearZKM()
    With ZKM
	.年月度 = ""
	.品番 = ""
	.取引先 = ""
	.受入 = ""
	.工順 = 0
	.在庫数 = 0
	.実棚日 = ""
	.変更事由 = ""
	.作成区分 = 0
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
        .件数 = 0
'   項目追加(自動)
	.在庫場所 = ""
	.品目区分 = ""
	.品名 = ""
	.仕様 = ""
	.単価 = 0
	.金額 = 0
	.最終出庫日 = ""
	.自ガス = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZKMInsert(Optional DBNo As Integer) As Boolean
    ZKMInsert = False
    With ZKM
        gSL_Select = "Insert Into 在庫マスタ  ( "
        gSL_Select = gSL_Select & "年月度,品番,取引先,"
        gSL_Select = gSL_Select & "受入,工順,在庫数,"
        gSL_Select = gSL_Select & "実棚日,変更事由,作成区分,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.年月度) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "" & .工順 & ","		
        gSL_Select = gSL_Select & "" & .在庫数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.実棚日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.変更事由) & "',"		
        gSL_Select = gSL_Select & "" & .作成区分 & ","		
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
    ZKMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZKMUpdate(Optional DBNo As Integer) As Boolean
    ZKMUpdate = False	
    With ZKM
        gSL_Select = "Update 在庫マスタ  Set "
        gSL_Select = gSL_Select & "年月度 = '" & RTrim(.年月度) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "工順 = " & .工順 & ","
        gSL_Select = gSL_Select & "在庫数 = " & .在庫数 & ","
        gSL_Select = gSL_Select & "実棚日 = '" & RTrim(.実棚日) & "',"
        gSL_Select = gSL_Select & "変更事由 = '" & RTrim(.変更事由) & "',"
        gSL_Select = gSL_Select & "作成区分 = " & .作成区分 & ""
        gSL_Select = gSL_Select & " Where 年月度 = '" & RTrim(.年月度) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 工順 = " & .工順 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZKMUpdate = True
End Function

Attribute VB_Name = "在庫履歴マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   在庫履歴マスタ
'*
'*   モジュールID  : 在庫履歴マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/02/12  By Cis
'*   作成日  :  2010/01/07  By Cis(HP-DX7500)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_在庫履歴マスタ
        年月度            As String
        品番              As String
        取引先            As String
        受入              As String
        工順              As Byte
        入力日付          As Date
        入力担当者        As String
        在庫数            As Currency
        実棚日            As String
        変更事由          As String
        作成区分          As Byte
        履歴区分          As Byte
'        作成日            As Date
        作成者            As String
'        作成端末          As String
        件数          As Long
    End Type

    Public ZKR As T_在庫履歴マスタ

    Public ZKR_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZKR(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZKR	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZKR
            Select Case adoCol.Name
                Case "年月度":            If Not IsNull(adoCol.Value) Then .年月度           = adoCol.Value
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番             = adoCol.Value
                Case "取引先":            If Not IsNull(adoCol.Value) Then .取引先           = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入             = adoCol.Value
                Case "工順":              If Not IsNull(adoCol.Value) Then .工順             = adoCol.Value
                Case "入力日付":          If Not IsNull(adoCol.Value) Then .入力日付         = adoCol.Value
                Case "入力担当者":        If Not IsNull(adoCol.Value) Then .入力担当者       = adoCol.Value
                Case "在庫数":            If Not IsNull(adoCol.Value) Then .在庫数           = adoCol.Value
                Case "実棚日":            If Not IsNull(adoCol.Value) Then .実棚日           = adoCol.Value
                Case "変更事由":          If Not IsNull(adoCol.Value) Then .変更事由         = adoCol.Value
                Case "作成区分":          If Not IsNull(adoCol.Value) Then .作成区分         = adoCol.Value
                Case "履歴区分":          If Not IsNull(adoCol.Value) Then .履歴区分         = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZKRRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZKRRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 在庫履歴マスタ"
        RTxtSql = RTxtSql &" Where 年月度 = '" & RTrim(ZKR.年月度) & "'"
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(ZKR.品番) & "'"
        RTxtSql = RTxtSql &" And   取引先 = '" & RTrim(ZKR.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(ZKR.受入) & "'"
        RTxtSql = RTxtSql &" And   工順 = " & ZKR.工順 & ""
        RTxtSql = RTxtSql &" And   入力日付 = '" & RTrim(ZKR.入力日付) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZKRRead = True
           Call ItemsFromZKR(DBNo)
	   ZKR_RDSTS = True
        else
	   ZKR_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZKRClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZKRReadNext(Optional DBNo As Integer = 0)
    If Not ZKR_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZKR_RDSTS = False
        Else
           Call ItemsFromZKR(DBNo)
           ZKR_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZKRClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZKR_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZKR )
'*----------------------------------------------*'
Public Sub ItemsClearZKR()
    With ZKR
	.年月度 = ""
	.品番 = ""
	.取引先 = ""
	.受入 = ""
	.工順 = 0
	.入力日付 = 0
	.入力担当者 = ""
	.在庫数 = 0
	.実棚日 = ""
	.変更事由 = ""
	.作成区分 = 0
	.履歴区分 = 0
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZKRInsert(Optional DBNo As Integer) As Boolean
    ZKRInsert = False
    With ZKR
        gSL_Select = "Insert Into 在庫履歴マスタ  ( "
        gSL_Select = gSL_Select & "年月度,品番,取引先,"
        gSL_Select = gSL_Select & "受入,工順,入力日付,"
        gSL_Select = gSL_Select & "入力担当者,在庫数,実棚日,"
        gSL_Select = gSL_Select & "変更事由,作成区分,履歴区分,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.年月度) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "" & .工順 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.入力日付) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.入力担当者) & "',"		
        gSL_Select = gSL_Select & "" & .在庫数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.実棚日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.変更事由) & "',"		
        gSL_Select = gSL_Select & "" & .作成区分 & ","		
        gSL_Select = gSL_Select & "" & .履歴区分 & ","		
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
    ZKRInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZKRUpdate(Optional DBNo As Integer) As Boolean
    ZKRUpdate = False	
    With ZKR
        gSL_Select = "Update 在庫履歴マスタ  Set "
        gSL_Select = gSL_Select & "年月度 = '" & RTrim(.年月度) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "工順 = " & .工順 & ","
        gSL_Select = gSL_Select & "入力日付 = '" & RTrim(.入力日付) & "',"
        gSL_Select = gSL_Select & "入力担当者 = '" & RTrim(.入力担当者) & "',"
        gSL_Select = gSL_Select & "在庫数 = " & .在庫数 & ","
        gSL_Select = gSL_Select & "実棚日 = '" & RTrim(.実棚日) & "',"
        gSL_Select = gSL_Select & "変更事由 = '" & RTrim(.変更事由) & "',"
        gSL_Select = gSL_Select & "作成区分 = " & .作成区分 & ","
        gSL_Select = gSL_Select & "履歴区分 = " & .履歴区分 & ""
        gSL_Select = gSL_Select & " Where 年月度 = '" & RTrim(.年月度) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 工順 = " & .工順 & ""
        gSL_Select = gSL_Select & "   And 入力日付 = '" & RTrim(.入力日付) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZKRUpdate = True
End Function

Attribute VB_Name = "受注SIM伝票分納ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   受注SIM伝票分納ワーク
'*
'*   モジュールID  : 受注SIM伝票分納ワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2009/12/22  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_受注SIM伝票分納ワーク
        端末番号          As String
        受注区分          As Byte
        受注日            As String
        受注便            As String
        納入日            As String
        納入便            As String
        納入先            As String
        出荷予定日        As String
        出荷予定便        As String
        受入              As String
        機番              As String
        納入番号          As String
        納品書頁          As Byte
        納品書行          As Byte
        品番              As String
        背番号            As String
        枚数              As Long
        収容数            As Long
        納入数            As Long
        端数              As Long
        件数          As Long
'   項目追加(自動)
	納入先名       As String
	表示品番       As String
    End Type

    Public JSB As T_受注SIM伝票分納ワーク

    Public JSB_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromJSB(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearJSB	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With JSB
            Select Case adoCol.Name
                Case "端末番号":          If Not IsNull(adoCol.Value) Then .端末番号         = adoCol.Value
                Case "受注区分":          If Not IsNull(adoCol.Value) Then .受注区分         = adoCol.Value
                Case "受注日":            If Not IsNull(adoCol.Value) Then .受注日           = adoCol.Value
                Case "受注便":            If Not IsNull(adoCol.Value) Then .受注便           = adoCol.Value
                Case "納入日":            If Not IsNull(adoCol.Value) Then .納入日           = adoCol.Value
                Case "納入便":            If Not IsNull(adoCol.Value) Then .納入便           = adoCol.Value
                Case "納入先":            If Not IsNull(adoCol.Value) Then .納入先           = adoCol.Value
                Case "出荷予定日":        If Not IsNull(adoCol.Value) Then .出荷予定日       = adoCol.Value
                Case "出荷予定便":        If Not IsNull(adoCol.Value) Then .出荷予定便       = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入             = adoCol.Value
                Case "機番":              If Not IsNull(adoCol.Value) Then .機番             = adoCol.Value
                Case "納入番号":          If Not IsNull(adoCol.Value) Then .納入番号         = adoCol.Value
                Case "納品書頁":          If Not IsNull(adoCol.Value) Then .納品書頁         = adoCol.Value
                Case "納品書行":          If Not IsNull(adoCol.Value) Then .納品書行         = adoCol.Value
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番             = adoCol.Value
                Case "背番号":            If Not IsNull(adoCol.Value) Then .背番号           = adoCol.Value
                Case "枚数":              If Not IsNull(adoCol.Value) Then .枚数             = adoCol.Value
                Case "収容数":            If Not IsNull(adoCol.Value) Then .収容数           = adoCol.Value
                Case "納入数":            If Not IsNull(adoCol.Value) Then .納入数           = adoCol.Value
                Case "端数":              If Not IsNull(adoCol.Value) Then .端数             = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "納入先名":	If Not IsNull(adoCol.Value) Then .納入先名	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function JSBRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    JSBRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 受注SIM伝票分納ワーク"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           JSBRead = True
           Call ItemsFromJSB(DBNo)
	   JSB_RDSTS = True
        else
	   JSB_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call JSBClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub JSBReadNext(Optional DBNo As Integer = 0)
    If Not JSB_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           JSB_RDSTS = False
        Else
           Call ItemsFromJSB(DBNo)
           JSB_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub JSBClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        JSB_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( JSB )
'*----------------------------------------------*'
Public Sub ItemsClearJSB()
    With JSB
	.端末番号 = ""
	.受注区分 = 0
	.受注日 = ""
	.受注便 = ""
	.納入日 = ""
	.納入便 = ""
	.納入先 = ""
	.出荷予定日 = ""
	.出荷予定便 = ""
	.受入 = ""
	.機番 = ""
	.納入番号 = ""
	.納品書頁 = 0
	.納品書行 = 0
	.品番 = ""
	.背番号 = ""
	.枚数 = 0
	.収容数 = 0
	.納入数 = 0
	.端数 = 0
        .件数 = 0
'   項目追加(自動)
	.納入先名 = ""
	.表示品番 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function JSBInsert(Optional DBNo As Integer) As Boolean
    JSBInsert = False
    With JSB
        gSL_Select = "Insert Into 受注SIM伝票分納ワーク  ( "
        gSL_Select = gSL_Select & "端末番号,受注区分,受注日,"
        gSL_Select = gSL_Select & "受注便,納入日,納入便,"
        gSL_Select = gSL_Select & "納入先,出荷予定日,出荷予定便,"
        gSL_Select = gSL_Select & "受入,機番,納入番号,"
        gSL_Select = gSL_Select & "納品書頁,納品書行,品番,"
        gSL_Select = gSL_Select & "背番号,枚数,収容数,"
        gSL_Select = gSL_Select & "納入数,端数"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.端末番号) & "',"		
        gSL_Select = gSL_Select & "" & .受注区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.受注日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受注便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.出荷予定日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.出荷予定便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.機番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"		
        gSL_Select = gSL_Select & "" & .納品書頁 & ","		
        gSL_Select = gSL_Select & "" & .納品書行 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "" & .枚数 & ","		
        gSL_Select = gSL_Select & "" & .収容数 & ","		
        gSL_Select = gSL_Select & "" & .納入数 & ","		
        gSL_Select = gSL_Select & "" & .端数 & ""		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    JSBInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function JSBUpdate(Optional DBNo As Integer) As Boolean
    JSBUpdate = False	
    With JSB
        gSL_Select = "Update 受注SIM伝票分納ワーク  Set "
        gSL_Select = gSL_Select & "端末番号 = '" & RTrim(.端末番号) & "',"
        gSL_Select = gSL_Select & "受注区分 = " & .受注区分 & ","
        gSL_Select = gSL_Select & "受注日 = '" & RTrim(.受注日) & "',"
        gSL_Select = gSL_Select & "受注便 = '" & RTrim(.受注便) & "',"
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "出荷予定日 = '" & RTrim(.出荷予定日) & "',"
        gSL_Select = gSL_Select & "出荷予定便 = '" & RTrim(.出荷予定便) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "機番 = '" & RTrim(.機番) & "',"
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "納品書頁 = " & .納品書頁 & ","
        gSL_Select = gSL_Select & "納品書行 = " & .納品書行 & ","
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "枚数 = " & .枚数 & ","
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "納入数 = " & .納入数 & ","
        gSL_Select = gSL_Select & "端数 = " & .端数 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    JSBUpdate = True
End Function

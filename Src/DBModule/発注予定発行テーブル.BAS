Attribute VB_Name = "発注予定発行テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   発注予定発行テーブル
'*
'*   モジュールID  : 発注予定発行テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2005/07/22  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_発注予定発行テーブル
        端末番号          As String
        番号              As Long
        仕入先            As String
        受入              As String
        サイクル          As String
        発行日            As String
        発行加算日        As Byte
        発注日            As String
        発注便            As String
        納入日            As String
        納入便            As String
        発注予定日        As String
        発注予定便        As String
        最終発注日        As String
        最終発注便        As String
        曜日発注          As Byte
        発注時間          As String
        昼夜区分          As Byte
        件数          As Long
'   項目追加(自動)
	仕入先名       As String
    End Type

    Public HYH As T_発注予定発行テーブル

    Public HYH_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHYH(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHYH	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With HYH
            Select Case adoCol.Name
                Case "端末番号":          If Not IsNull(adoCol.Value) Then .端末番号         = adoCol.Value
                Case "番号":              If Not IsNull(adoCol.Value) Then .番号             = adoCol.Value
                Case "仕入先":            If Not IsNull(adoCol.Value) Then .仕入先           = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入             = adoCol.Value
                Case "サイクル":          If Not IsNull(adoCol.Value) Then .サイクル         = adoCol.Value
                Case "発行日":            If Not IsNull(adoCol.Value) Then .発行日           = adoCol.Value
                Case "発行加算日":        If Not IsNull(adoCol.Value) Then .発行加算日       = adoCol.Value
                Case "発注日":            If Not IsNull(adoCol.Value) Then .発注日           = adoCol.Value
                Case "発注便":            If Not IsNull(adoCol.Value) Then .発注便           = adoCol.Value
                Case "納入日":            If Not IsNull(adoCol.Value) Then .納入日           = adoCol.Value
                Case "納入便":            If Not IsNull(adoCol.Value) Then .納入便           = adoCol.Value
                Case "発注予定日":        If Not IsNull(adoCol.Value) Then .発注予定日       = adoCol.Value
                Case "発注予定便":        If Not IsNull(adoCol.Value) Then .発注予定便       = adoCol.Value
                Case "最終発注日":        If Not IsNull(adoCol.Value) Then .最終発注日       = adoCol.Value
                Case "最終発注便":        If Not IsNull(adoCol.Value) Then .最終発注便       = adoCol.Value
                Case "曜日発注":          If Not IsNull(adoCol.Value) Then .曜日発注         = adoCol.Value
                Case "発注時間":          If Not IsNull(adoCol.Value) Then .発注時間         = adoCol.Value
                Case "昼夜区分":          If Not IsNull(adoCol.Value) Then .昼夜区分         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "仕入先名":	If Not IsNull(adoCol.Value) Then .仕入先名	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function HYHRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HYHRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 発注予定発行テーブル"
        RTxtSql = RTxtSql &" Where 端末番号 = '" & RTrim(HYH.端末番号) & "'"
        RTxtSql = RTxtSql &" And   番号 = " & HYH.番号 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HYHRead = True
           Call ItemsFromHYH(DBNo)
	   HYH_RDSTS = True
        else
	   HYH_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call HYHClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HYHReadNext(Optional DBNo As Integer = 0)
    If Not HYH_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HYH_RDSTS = False
        Else
           Call ItemsFromHYH(DBNo)
           HYH_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HYHClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HYH_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HYH )
'*----------------------------------------------*'
Public Sub ItemsClearHYH()
    With HYH
	.端末番号 = ""
	.番号 = 0
	.仕入先 = ""
	.受入 = ""
	.サイクル = ""
	.発行日 = ""
	.発行加算日 = 0
	.発注日 = ""
	.発注便 = ""
	.納入日 = ""
	.納入便 = ""
	.発注予定日 = ""
	.発注予定便 = ""
	.最終発注日 = ""
	.最終発注便 = ""
	.曜日発注 = 0
	.発注時間 = ""
	.昼夜区分 = 0
        .件数 = 0
'   項目追加(自動)
	.仕入先名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function HYHInsert(Optional DBNo As Integer) As Boolean
    HYHInsert = True
    With HYH
        gSL_Select = "Insert Into 発注予定発行テーブル  ( "
        gSL_Select = gSL_Select & "端末番号,仕入先,受入,"
        gSL_Select = gSL_Select & "サイクル,発行日,発行加算日,"
        gSL_Select = gSL_Select & "発注日,発注便,納入日,"
        gSL_Select = gSL_Select & "納入便,発注予定日,発注予定便,"
        gSL_Select = gSL_Select & "最終発注日,最終発注便,曜日発注,"
        gSL_Select = gSL_Select & "発注時間,昼夜区分"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.端末番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.サイクル) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発行日) & "',"		
        gSL_Select = gSL_Select & "" & .発行加算日 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.発注日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注予定日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注予定便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.最終発注日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.最終発注便) & "',"		
        gSL_Select = gSL_Select & "" & .曜日発注 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.発注時間) & "',"		
        gSL_Select = gSL_Select & "" & .昼夜区分 & ""		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HYHInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function HYHUpdate(Optional DBNo As Integer) As Boolean
    HYHUpdate = False	
    With HYH
        gSL_Select = "Update 発注予定発行テーブル  Set "
        gSL_Select = gSL_Select & "端末番号 = '" & RTrim(.端末番号) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "サイクル = '" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "発行日 = '" & RTrim(.発行日) & "',"
        gSL_Select = gSL_Select & "発行加算日 = " & .発行加算日 & ","
        gSL_Select = gSL_Select & "発注日 = '" & RTrim(.発注日) & "',"
        gSL_Select = gSL_Select & "発注便 = '" & RTrim(.発注便) & "',"
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "発注予定日 = '" & RTrim(.発注予定日) & "',"
        gSL_Select = gSL_Select & "発注予定便 = '" & RTrim(.発注予定便) & "',"
        gSL_Select = gSL_Select & "最終発注日 = '" & RTrim(.最終発注日) & "',"
        gSL_Select = gSL_Select & "最終発注便 = '" & RTrim(.最終発注便) & "',"
        gSL_Select = gSL_Select & "曜日発注 = " & .曜日発注 & ","
        gSL_Select = gSL_Select & "発注時間 = '" & RTrim(.発注時間) & "',"
        gSL_Select = gSL_Select & "昼夜区分 = " & .昼夜区分 & ""
        gSL_Select = gSL_Select & " Where 端末番号 = '" & RTrim(.端末番号) & "'"
        gSL_Select = gSL_Select & "   And 番号 = " & .番号 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HYHUpdate = True
End Function

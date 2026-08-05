Attribute VB_Name = "売上予想未登録テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   売上予想未登録テーブル
'*
'*   モジュールID  : 売上予想未登録テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/11/04  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_売上予想未登録テーブル
        内示年月      As String
        取引先        As String
        受入          As String
        品番          As String
        管理番号      As String
        号補区分      As String
        工区          As String
        内示数        As Long
        翌月内示数    As Long
        翌々月内示数  As Long
        作成時間      As String
'        作成端末      As String
'        作成日        As String
        件数          As Long
'   項目追加(自動)
	表示品番      As String
	納入先名      As String
	契約先名      As String
	契約先        As String
    End Type

    Public UYH As T_売上予想未登録テーブル

    Public UYH_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromUYH(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearUYH	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With UYH
            Select Case adoCol.Name
                Case "内示年月":      If Not IsNull(adoCol.Value) Then .内示年月     = adoCol.Value
                Case "取引先":        If Not IsNull(adoCol.Value) Then .取引先       = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "管理番号":      If Not IsNull(adoCol.Value) Then .管理番号     = adoCol.Value
                Case "号補区分":      If Not IsNull(adoCol.Value) Then .号補区分     = adoCol.Value
                Case "工区":          If Not IsNull(adoCol.Value) Then .工区         = adoCol.Value
                Case "内示数":        If Not IsNull(adoCol.Value) Then .内示数       = adoCol.Value
                Case "翌月内示数":    If Not IsNull(adoCol.Value) Then .翌月内示数   = adoCol.Value
                Case "翌々月内示数":  If Not IsNull(adoCol.Value) Then .翌々月内示数 = adoCol.Value
                Case "作成時間":      If Not IsNull(adoCol.Value) Then .作成時間     = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "納入先名":	If Not IsNull(adoCol.Value) Then .納入先名	= adoCol.Value
		Case "契約先名":	If Not IsNull(adoCol.Value) Then .契約先名	= adoCol.Value
		Case "契約先":	If Not IsNull(adoCol.Value) Then .契約先	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function UYHRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    UYHRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 売上予想未登録テーブル"
        RTxtSql = RTxtSql &" Where 内示年月 = '" & RTrim(UYH.内示年月) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           UYHRead = True
           Call ItemsFromUYH(DBNo)
	   UYH_RDSTS = True
        else
	   UYH_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call UYHClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub UYHReadNext(Optional DBNo As Integer = 0)
    If Not UYH_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           UYH_RDSTS = False
        Else
           Call ItemsFromUYH(DBNo)
           UYH_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub UYHClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        UYH_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( UYH )
'*----------------------------------------------*'
Public Sub ItemsClearUYH()
    With UYH
	.内示年月 = ""
	.取引先 = ""
	.受入 = ""
	.品番 = ""
	.管理番号 = ""
	.号補区分 = ""
	.工区 = ""
	.内示数 = 0
	.翌月内示数 = 0
	.翌々月内示数 = 0
	.作成時間 = ""
'	.作成端末 = ""
'	.作成日 = ""
        .件数 = 0
'   項目追加(自動)
	.表示品番 = ""
	.納入先名 = ""
	.契約先名 = ""
	.契約先 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function UYHInsert(Optional DBNo As Integer) As Boolean
    UYHInsert = True
    With UYH
        gSL_Select = "Insert Into 売上予想未登録テーブル  ( "
        gSL_Select = gSL_Select & "内示年月,取引先,受入,"
        gSL_Select = gSL_Select & "品番,管理番号,号補区分,"
        gSL_Select = gSL_Select & "工区,内示数,翌月内示数,"
        gSL_Select = gSL_Select & "翌々月内示数,作成時間,"
        gSL_Select = gSL_Select & "作成端末,作成日"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.内示年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.管理番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.号補区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.工区) & "',"		
        gSL_Select = gSL_Select & "" & .内示数 & ","		
        gSL_Select = gSL_Select & "" & .翌月内示数 & ","		
        gSL_Select = gSL_Select & "" & .翌々月内示数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.作成時間) & "',"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,15),"		
        gSL_Select = gSL_Select & "GETDATE()"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    UYHInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function UYHUpdate(Optional DBNo As Integer) As Boolean
    UYHUpdate = False	
    With UYH
        gSL_Select = "Update 売上予想未登録テーブル  Set "
        gSL_Select = gSL_Select & "内示年月 = '" & RTrim(.内示年月) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "管理番号 = '" & RTrim(.管理番号) & "',"
        gSL_Select = gSL_Select & "号補区分 = '" & RTrim(.号補区分) & "',"
        gSL_Select = gSL_Select & "工区 = '" & RTrim(.工区) & "',"
        gSL_Select = gSL_Select & "内示数 = " & .内示数 & ","
        gSL_Select = gSL_Select & "翌月内示数 = " & .翌月内示数 & ","
        gSL_Select = gSL_Select & "翌々月内示数 = " & .翌々月内示数 & ","
        gSL_Select = gSL_Select & "作成時間 = '" & RTrim(.作成時間) & "'"
        gSL_Select = gSL_Select & " Where 内示年月 = '" & RTrim(.内示年月) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    UYHUpdate = True
End Function

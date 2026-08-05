Attribute VB_Name = "作業日報テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   作業日報テーブル
'*
'*   モジュールID  : 作業日報テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2009/12/03  By Cis(FMVW5220)
'*   作成日  :  2009/12/07  By Cis(FMVW5220)
'*   作成日  :  2009/12/09  By Cis(FMVW5220)
'*   作成日  :  2009/12/24  By Cis(FMVW5220)
'*   作成日  :  2010/01/05  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_作業日報テーブル
        作業日報管理NO        As Long
        作業日報管理行        As Long
        作業日                As String
        部門                  As String
        受入                  As String
        作業者                As String
        機械NO                As String
        品番                  As String
        工順                  As Byte
        加工数                As Long
        開始                  As String
        終了                  As String
        不良数                As Long
        段取                  As Byte
'        作成日                As Date
        作成者                As String
'        作成端末              As String
'        更新日                As Date
        更新者                As String
'        更新端末              As String
        件数          As Long
'   項目追加(自動)
	機械NO名       As String
	部門名         As String
	作業者名       As String
	品名           As String
    End Type

    Public SGT As T_作業日報テーブル

    Public SGT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromSGT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearSGT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With SGT
            Select Case adoCol.Name
                Case "作業日報管理NO":        If Not IsNull(adoCol.Value) Then .作業日報管理NO       = adoCol.Value
                Case "作業日報管理行":        If Not IsNull(adoCol.Value) Then .作業日報管理行       = adoCol.Value
                Case "作業日":                If Not IsNull(adoCol.Value) Then .作業日               = adoCol.Value
                Case "部門":                  If Not IsNull(adoCol.Value) Then .部門                 = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入                 = adoCol.Value
                Case "作業者":                If Not IsNull(adoCol.Value) Then .作業者               = adoCol.Value
                Case "機械NO":                If Not IsNull(adoCol.Value) Then .機械NO               = adoCol.Value
                Case "品番":                  If Not IsNull(adoCol.Value) Then .品番                 = adoCol.Value
                Case "工順":                  If Not IsNull(adoCol.Value) Then .工順                 = adoCol.Value
                Case "加工数":                If Not IsNull(adoCol.Value) Then .加工数               = adoCol.Value
                Case "開始":                  If Not IsNull(adoCol.Value) Then .開始                 = adoCol.Value
                Case "終了":                  If Not IsNull(adoCol.Value) Then .終了                 = adoCol.Value
                Case "不良数":                If Not IsNull(adoCol.Value) Then .不良数               = adoCol.Value
                Case "段取":                  If Not IsNull(adoCol.Value) Then .段取                 = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
                Case "作成者":                If Not IsNull(adoCol.Value) Then .作成者               = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
'                Case "更新日":                If Not IsNull(adoCol.Value) Then .更新日               = adoCol.Value
                Case "更新者":                If Not IsNull(adoCol.Value) Then .更新者               = adoCol.Value
'                Case "更新端末":              If Not IsNull(adoCol.Value) Then .更新端末             = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "機械NO名":	If Not IsNull(adoCol.Value) Then .機械NO名	= adoCol.Value
		Case "部門名":	If Not IsNull(adoCol.Value) Then .部門名	= adoCol.Value
		Case "作業者名":	If Not IsNull(adoCol.Value) Then .作業者名	= adoCol.Value
		Case "品名":	If Not IsNull(adoCol.Value) Then .品名	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function SGTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    SGTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 作業日報テーブル"
        RTxtSql = RTxtSql &" Where 作業日報管理NO = " & SGT.作業日報管理NO & ""
        RTxtSql = RTxtSql &" And   作業日報管理行 = " & SGT.作業日報管理行 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           SGTRead = True
           Call ItemsFromSGT(DBNo)
	   SGT_RDSTS = True
        else
	   SGT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call SGTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub SGTReadNext(Optional DBNo As Integer = 0)
    If Not SGT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           SGT_RDSTS = False
        Else
           Call ItemsFromSGT(DBNo)
           SGT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub SGTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        SGT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( SGT )
'*----------------------------------------------*'
Public Sub ItemsClearSGT()
    With SGT
	.作業日報管理NO = 0
	.作業日報管理行 = 0
	.作業日 = ""
	.部門 = ""
	.受入 = ""
	.作業者 = ""
	.機械NO = ""
	.品番 = ""
	.工順 = 0
	.加工数 = 0
	.開始 = ""
	.終了 = ""
	.不良数 = 0
	.段取 = 0
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
'	.更新日 = 0
	.更新者 = ""
'	.更新端末 = ""
        .件数 = 0
'   項目追加(自動)
	.機械NO名 = ""
	.部門名 = ""
	.作業者名 = ""
	.品名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function SGTInsert(Optional DBNo As Integer) As Boolean
    SGTInsert = False
    With SGT
        gSL_Select = "Insert Into 作業日報テーブル  ( "
        gSL_Select = gSL_Select & "作業日報管理NO,作業日報管理行,作業日,"
        gSL_Select = gSL_Select & "部門,受入,作業者,"
        gSL_Select = gSL_Select & "機械NO,品番,工順,"
        gSL_Select = gSL_Select & "加工数,開始,終了,"
        gSL_Select = gSL_Select & "不良数,段取,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .作業日報管理NO & ","		
        gSL_Select = gSL_Select & "" & .作業日報管理行 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.作業日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.部門) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.作業者) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.機械NO) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "" & .工順 & ","		
        gSL_Select = gSL_Select & "" & .加工数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.開始) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.終了) & "',"		
        gSL_Select = gSL_Select & "" & .不良数 & ","		
        gSL_Select = gSL_Select & "" & .段取 & ","		
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
    SGTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function SGTUpdate(Optional DBNo As Integer) As Boolean
    SGTUpdate = False	
    With SGT
        gSL_Select = "Update 作業日報テーブル  Set "
        gSL_Select = gSL_Select & "作業日報管理NO = " & .作業日報管理NO & ","
        gSL_Select = gSL_Select & "作業日報管理行 = " & .作業日報管理行 & ","
        gSL_Select = gSL_Select & "作業日 = '" & RTrim(.作業日) & "',"
        gSL_Select = gSL_Select & "部門 = '" & RTrim(.部門) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "作業者 = '" & RTrim(.作業者) & "',"
        gSL_Select = gSL_Select & "機械NO = '" & RTrim(.機械NO) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "工順 = " & .工順 & ","
        gSL_Select = gSL_Select & "加工数 = " & .加工数 & ","
        gSL_Select = gSL_Select & "開始 = '" & RTrim(.開始) & "',"
        gSL_Select = gSL_Select & "終了 = '" & RTrim(.終了) & "',"
        gSL_Select = gSL_Select & "不良数 = " & .不良数 & ","
        gSL_Select = gSL_Select & "段取 = " & .段取 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 作業日報管理NO = " & .作業日報管理NO & ""
        gSL_Select = gSL_Select & "   And 作業日報管理行 = " & .作業日報管理行 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SGTUpdate = True
End Function

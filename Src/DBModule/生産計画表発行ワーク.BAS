Attribute VB_Name = "生産計画表発行ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   生産計画表発行ワーク
'*
'*   モジュールID  : 生産計画表発行ワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/02/26  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_生産計画表発行ワーク
        処理端末      As String
        区分          As String
        品番          As String
        工区          As String
        前月計画数    As Long
        当月計画数    As Long
        前月内示数    As Long
        当月内示数    As Long
        翌月内示数    As Long
        翌々月内示数  As Long
        変更区分      As Byte
        件数          As Long
'   項目追加(自動)
	車型          As String
	収容数        As Long
	社内背番      As String
	品番編集      As String
    End Type

    Public SKH As T_生産計画表発行ワーク

    Public SKH_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromSKH(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearSKH	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With SKH
            Select Case adoCol.Name
                Case "処理端末":      If Not IsNull(adoCol.Value) Then .処理端末     = adoCol.Value
                Case "区分":          If Not IsNull(adoCol.Value) Then .区分         = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "工区":          If Not IsNull(adoCol.Value) Then .工区         = adoCol.Value
                Case "前月計画数":    If Not IsNull(adoCol.Value) Then .前月計画数   = adoCol.Value
                Case "当月計画数":    If Not IsNull(adoCol.Value) Then .当月計画数   = adoCol.Value
                Case "前月内示数":    If Not IsNull(adoCol.Value) Then .前月内示数   = adoCol.Value
                Case "当月内示数":    If Not IsNull(adoCol.Value) Then .当月内示数   = adoCol.Value
                Case "翌月内示数":    If Not IsNull(adoCol.Value) Then .翌月内示数   = adoCol.Value
                Case "翌々月内示数":  If Not IsNull(adoCol.Value) Then .翌々月内示数 = adoCol.Value
                Case "変更区分":      If Not IsNull(adoCol.Value) Then .変更区分     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "車型":	If Not IsNull(adoCol.Value) Then .車型	= adoCol.Value
		Case "収容数":	If Not IsNull(adoCol.Value) Then .収容数	= adoCol.Value
		Case "社内背番":	If Not IsNull(adoCol.Value) Then .社内背番	= adoCol.Value
		Case "品番編集":	If Not IsNull(adoCol.Value) Then .品番編集	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function SKHRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    SKHRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 生産計画表発行ワーク"
        RTxtSql = RTxtSql &" Where 処理端末 = '" & RTrim(SKH.処理端末) & "'"
        RTxtSql = RTxtSql &" And   区分 = '" & RTrim(SKH.区分) & "'"
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(SKH.品番) & "'"
        RTxtSql = RTxtSql &" And   工区 = '" & RTrim(SKH.工区) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           SKHRead = True
           Call ItemsFromSKH(DBNo)
	   SKH_RDSTS = True
        else
	   SKH_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call SKHClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub SKHReadNext(Optional DBNo As Integer = 0)
    If Not SKH_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           SKH_RDSTS = False
        Else
           Call ItemsFromSKH(DBNo)
           SKH_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub SKHClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        SKH_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( SKH )
'*----------------------------------------------*'
Public Sub ItemsClearSKH()
    With SKH
	.処理端末 = ""
	.区分 = ""
	.品番 = ""
	.工区 = ""
	.前月計画数 = 0
	.当月計画数 = 0
	.前月内示数 = 0
	.当月内示数 = 0
	.翌月内示数 = 0
	.翌々月内示数 = 0
	.変更区分 = 0
        .件数 = 0
'   項目追加(自動)
	.車型 = ""
	.収容数 = 0
	.社内背番 = ""
	.品番編集 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function SKHInsert(Optional DBNo As Integer) As Boolean
    SKHInsert = True
    With SKH
        gSL_Select = "Insert Into 生産計画表発行ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,区分,品番,"
        gSL_Select = gSL_Select & "工区,前月計画数,当月計画数,"
        gSL_Select = gSL_Select & "前月内示数,当月内示数,翌月内示数,"
        gSL_Select = gSL_Select & "翌々月内示数,変更区分"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.工区) & "',"		
        gSL_Select = gSL_Select & "" & .前月計画数 & ","		
        gSL_Select = gSL_Select & "" & .当月計画数 & ","		
        gSL_Select = gSL_Select & "" & .前月内示数 & ","		
        gSL_Select = gSL_Select & "" & .当月内示数 & ","		
        gSL_Select = gSL_Select & "" & .翌月内示数 & ","		
        gSL_Select = gSL_Select & "" & .翌々月内示数 & ","		
        gSL_Select = gSL_Select & "" & .変更区分 & ""		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SKHInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function SKHUpdate(Optional DBNo As Integer) As Boolean
    SKHUpdate = False	
    With SKH
        gSL_Select = "Update 生産計画表発行ワーク  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "区分 = '" & RTrim(.区分) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "工区 = '" & RTrim(.工区) & "',"
        gSL_Select = gSL_Select & "前月計画数 = " & .前月計画数 & ","
        gSL_Select = gSL_Select & "当月計画数 = " & .当月計画数 & ","
        gSL_Select = gSL_Select & "前月内示数 = " & .前月内示数 & ","
        gSL_Select = gSL_Select & "当月内示数 = " & .当月内示数 & ","
        gSL_Select = gSL_Select & "翌月内示数 = " & .翌月内示数 & ","
        gSL_Select = gSL_Select & "翌々月内示数 = " & .翌々月内示数 & ","
        gSL_Select = gSL_Select & "変更区分 = " & .変更区分 & ""
        gSL_Select = gSL_Select & " Where 処理端末 = '" & RTrim(.処理端末) & "'"
        gSL_Select = gSL_Select & "   And 区分 = '" & RTrim(.区分) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 工区 = '" & RTrim(.工区) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SKHUpdate = True
End Function

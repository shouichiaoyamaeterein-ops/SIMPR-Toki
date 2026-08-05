Attribute VB_Name = "有償売上書ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   有償売上書ワーク
'*
'*   モジュールID  : 有償売上書ワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2008/07/05  By Cis(FMV-E630)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_有償売上書ワーク
        処理端末      As String
        納入先        As String
        納入先工場CD  As String
        納入先CD      As String
        納入先名      As String
        受入          As String
        自社工区名    As String
        品番          As String
        背番号        As String
        手配単位      As String
        材質          As String
        表示寸法      As String
        単価          As Currency
        単重          As Currency
        納入日        As String
        納入数        As Currency
        金額          As Currency
        部品材料区分  As String
        件数          As Long
'   項目追加(自動)
	明細件数       As Integer
    End Type

    Public YUW As T_有償売上書ワーク

    Public YUW_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromYUW(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearYUW	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With YUW
            Select Case adoCol.Name
                Case "処理端末":      If Not IsNull(adoCol.Value) Then .処理端末     = adoCol.Value
                Case "納入先":        If Not IsNull(adoCol.Value) Then .納入先       = adoCol.Value
                Case "納入先工場CD":  If Not IsNull(adoCol.Value) Then .納入先工場CD = adoCol.Value
                Case "納入先CD":      If Not IsNull(adoCol.Value) Then .納入先CD     = adoCol.Value
                Case "納入先名":      If Not IsNull(adoCol.Value) Then .納入先名     = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "自社工区名":    If Not IsNull(adoCol.Value) Then .自社工区名   = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "背番号":        If Not IsNull(adoCol.Value) Then .背番号       = adoCol.Value
                Case "手配単位":      If Not IsNull(adoCol.Value) Then .手配単位     = adoCol.Value
                Case "材質":          If Not IsNull(adoCol.Value) Then .材質         = adoCol.Value
                Case "表示寸法":      If Not IsNull(adoCol.Value) Then .表示寸法     = adoCol.Value
                Case "単価":          If Not IsNull(adoCol.Value) Then .単価         = adoCol.Value
                Case "単重":          If Not IsNull(adoCol.Value) Then .単重         = adoCol.Value
                Case "納入日":        If Not IsNull(adoCol.Value) Then .納入日       = adoCol.Value
                Case "納入数":        If Not IsNull(adoCol.Value) Then .納入数       = adoCol.Value
                Case "金額":          If Not IsNull(adoCol.Value) Then .金額         = adoCol.Value
                Case "部品材料区分":  If Not IsNull(adoCol.Value) Then .部品材料区分 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "明細件数":	If Not IsNull(adoCol.Value) Then .明細件数	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function YUWRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    YUWRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 有償売上書ワーク"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           YUWRead = True
           Call ItemsFromYUW(DBNo)
	   YUW_RDSTS = True
        else
	   YUW_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call YUWClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub YUWReadNext(Optional DBNo As Integer = 0)
    If Not YUW_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           YUW_RDSTS = False
        Else
           Call ItemsFromYUW(DBNo)
           YUW_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub YUWClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        YUW_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( YUW )
'*----------------------------------------------*'
Public Sub ItemsClearYUW()
    With YUW
	.処理端末 = ""
	.納入先 = ""
	.納入先工場CD = ""
	.納入先CD = ""
	.納入先名 = ""
	.受入 = ""
	.自社工区名 = ""
	.品番 = ""
	.背番号 = ""
	.手配単位 = ""
	.材質 = ""
	.表示寸法 = ""
	.単価 = 0
	.単重 = 0
	.納入日 = ""
	.納入数 = 0
	.金額 = 0
	.部品材料区分 = ""
        .件数 = 0
'   項目追加(自動)
	.明細件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function YUWInsert(Optional DBNo As Integer) As Boolean
    YUWInsert = False
    With YUW
        gSL_Select = "Insert Into 有償売上書ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,納入先,納入先工場CD,"
        gSL_Select = gSL_Select & "納入先CD,納入先名,受入,"
        gSL_Select = gSL_Select & "自社工区名,品番,背番号,"
        gSL_Select = gSL_Select & "手配単位,材質,表示寸法,"
        gSL_Select = gSL_Select & "単価,単重,納入日,"
        gSL_Select = gSL_Select & "納入数,金額,部品材料区分"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先工場CD) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先CD) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.自社工区名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配単位) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.表示寸法) & "',"		
        gSL_Select = gSL_Select & "" & .単価 & ","		
        gSL_Select = gSL_Select & "" & .単重 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"		
        gSL_Select = gSL_Select & "" & .納入数 & ","		
        gSL_Select = gSL_Select & "" & .金額 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.部品材料区分) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    YUWInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function YUWUpdate(Optional DBNo As Integer) As Boolean
    YUWUpdate = False	
    With YUW
        gSL_Select = "Update 有償売上書ワーク  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "納入先工場CD = '" & RTrim(.納入先工場CD) & "',"
        gSL_Select = gSL_Select & "納入先CD = '" & RTrim(.納入先CD) & "',"
        gSL_Select = gSL_Select & "納入先名 = '" & RTrim(.納入先名) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "自社工区名 = '" & RTrim(.自社工区名) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "手配単位 = '" & RTrim(.手配単位) & "',"
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "表示寸法 = '" & RTrim(.表示寸法) & "',"
        gSL_Select = gSL_Select & "単価 = " & .単価 & ","
        gSL_Select = gSL_Select & "単重 = " & .単重 & ","
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入数 = " & .納入数 & ","
        gSL_Select = gSL_Select & "金額 = " & .金額 & ","
        gSL_Select = gSL_Select & "部品材料区分 = '" & RTrim(.部品材料区分) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    YUWUpdate = True
End Function

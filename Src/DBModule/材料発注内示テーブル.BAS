Attribute VB_Name = "材料発注内示テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料発注内示テーブル
'*
'*   モジュールID  : 材料発注内示テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/04/07  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料発注内示テーブル
        内示年月      As String
        材料管理番号  As Long
        手配先        As String
        受入          As String
        直送先        As String
        納期          As String
        仕掛日        As String
        手配日        As String
        材質          As String
        寸法区分      As Byte
        板厚          As Currency
        幅            As Currency
        長さ          As Currency
        数量          As Currency
        処理区分      As Byte
        削除区分      As Byte
        変更回数      As Byte
        計画年月      As String
        作成者        As String
'        作成日        As Date
'        作成端末      As String
        更新者        As String
'        更新日        As Date
'        更新端末      As String
        件数          As Long
'   項目追加(自動)
	当月内示数    As Currency
	翌月内示数    As Currency
	翌々月内示数  As Currency
	手配先名      As String
	材種コード    As String
	背番号        As String
	サイクル      As String
	収容数        As Currency
	材料名        As String
	手配単位名    As String
	表示寸法      As String
	発注比率      As Byte
	KB            As String
	DD            As Byte
	増減率        As Integer
	今回計画数    As Currency
	前回計画数    As Currency
	計画差        As Currency
    End Type

    Public ZNJ As T_材料発注内示テーブル

    Public ZNJ_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZNJ(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZNJ	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZNJ
            Select Case adoCol.Name
                Case "内示年月":      If Not IsNull(adoCol.Value) Then .内示年月     = adoCol.Value
                Case "材料管理番号":  If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "手配先":        If Not IsNull(adoCol.Value) Then .手配先       = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "直送先":        If Not IsNull(adoCol.Value) Then .直送先       = adoCol.Value
                Case "納期":          If Not IsNull(adoCol.Value) Then .納期         = adoCol.Value
                Case "仕掛日":        If Not IsNull(adoCol.Value) Then .仕掛日       = adoCol.Value
                Case "手配日":        If Not IsNull(adoCol.Value) Then .手配日       = adoCol.Value
                Case "材質":          If Not IsNull(adoCol.Value) Then .材質         = adoCol.Value
                Case "寸法区分":      If Not IsNull(adoCol.Value) Then .寸法区分     = adoCol.Value
                Case "板厚":          If Not IsNull(adoCol.Value) Then .板厚         = adoCol.Value
                Case "幅":            If Not IsNull(adoCol.Value) Then .幅           = adoCol.Value
                Case "長さ":          If Not IsNull(adoCol.Value) Then .長さ         = adoCol.Value
                Case "数量":          If Not IsNull(adoCol.Value) Then .数量         = adoCol.Value
                Case "処理区分":      If Not IsNull(adoCol.Value) Then .処理区分     = adoCol.Value
                Case "削除区分":      If Not IsNull(adoCol.Value) Then .削除区分     = adoCol.Value
                Case "変更回数":      If Not IsNull(adoCol.Value) Then .変更回数     = adoCol.Value
                Case "計画年月":      If Not IsNull(adoCol.Value) Then .計画年月     = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者       = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "当月内示数":	If Not IsNull(adoCol.Value) Then .当月内示数	= adoCol.Value
		Case "翌月内示数":	If Not IsNull(adoCol.Value) Then .翌月内示数	= adoCol.Value
		Case "翌々月内示数":	If Not IsNull(adoCol.Value) Then .翌々月内示数	= adoCol.Value
		Case "手配先名":	If Not IsNull(adoCol.Value) Then .手配先名	= adoCol.Value
		Case "材種コード":	If Not IsNull(adoCol.Value) Then .材種コード	= adoCol.Value
		Case "背番号":	If Not IsNull(adoCol.Value) Then .背番号	= adoCol.Value
		Case "サイクル":	If Not IsNull(adoCol.Value) Then .サイクル	= adoCol.Value
		Case "収容数":	If Not IsNull(adoCol.Value) Then .収容数	= adoCol.Value
		Case "材料名":	If Not IsNull(adoCol.Value) Then .材料名	= adoCol.Value
		Case "手配単位名":	If Not IsNull(adoCol.Value) Then .手配単位名	= adoCol.Value
		Case "表示寸法":	If Not IsNull(adoCol.Value) Then .表示寸法	= adoCol.Value
		Case "発注比率":	If Not IsNull(adoCol.Value) Then .発注比率	= adoCol.Value
		Case "KB":	If Not IsNull(adoCol.Value) Then .KB	= adoCol.Value
		Case "DD":	If Not IsNull(adoCol.Value) Then .DD	= adoCol.Value
		Case "増減率":	If Not IsNull(adoCol.Value) Then .増減率	= adoCol.Value
		Case "今回計画数":	If Not IsNull(adoCol.Value) Then .今回計画数	= adoCol.Value
		Case "前回計画数":	If Not IsNull(adoCol.Value) Then .前回計画数	= adoCol.Value
		Case "計画差":	If Not IsNull(adoCol.Value) Then .計画差	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZNJRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZNJRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 材料発注内示テーブル"
        RTxtSql = RTxtSql &" Where 内示年月 = '" & RTrim(ZNJ.内示年月) & "'"
        RTxtSql = RTxtSql &" And   材料管理番号 = " & ZNJ.材料管理番号 & ""
        RTxtSql = RTxtSql &" And   手配先 = '" & RTrim(ZNJ.手配先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(ZNJ.受入) & "'"
        RTxtSql = RTxtSql &" And   直送先 = '" & RTrim(ZNJ.直送先) & "'"
        RTxtSql = RTxtSql &" And   納期 = '" & RTrim(ZNJ.納期) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZNJRead = True
           Call ItemsFromZNJ(DBNo)
	   ZNJ_RDSTS = True
        else
	   ZNJ_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZNJClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZNJReadNext(Optional DBNo As Integer = 0)
    If Not ZNJ_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZNJ_RDSTS = False
        Else
           Call ItemsFromZNJ(DBNo)
           ZNJ_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZNJClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZNJ_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZNJ )
'*----------------------------------------------*'
Public Sub ItemsClearZNJ()
    With ZNJ
	.内示年月 = ""
	.材料管理番号 = 0
	.手配先 = ""
	.受入 = ""
	.直送先 = ""
	.納期 = ""
	.仕掛日 = ""
	.手配日 = ""
	.材質 = ""
	.寸法区分 = 0
	.板厚 = 0
	.幅 = 0
	.長さ = 0
	.数量 = 0
	.処理区分 = 0
	.削除区分 = 0
	.変更回数 = 0
	.計画年月 = ""
	.作成者 = ""
'	.作成日 = 0
'	.作成端末 = ""
	.更新者 = ""
'	.更新日 = 0
'	.更新端末 = ""
        .件数 = 0
'   項目追加(自動)
	.当月内示数 = 0
	.翌月内示数 = 0
	.翌々月内示数 = 0
	.手配先名 = ""
	.材種コード = ""
	.背番号 = ""
	.サイクル = ""
	.収容数 = 0
	.材料名 = ""
	.手配単位名 = ""
	.表示寸法 = ""
	.発注比率 = 0
	.KB = ""
	.DD = 0
	.増減率 = 0
	.今回計画数 = 0
	.前回計画数 = 0
	.計画差 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZNJInsert(Optional DBNo As Integer) As Boolean
    ZNJInsert = True
    With ZNJ
        gSL_Select = "Insert Into 材料発注内示テーブル  ( "
        gSL_Select = gSL_Select & "内示年月,材料管理番号,手配先,"
        gSL_Select = gSL_Select & "受入,直送先,納期,"
        gSL_Select = gSL_Select & "仕掛日,手配日,材質,"
        gSL_Select = gSL_Select & "寸法区分,板厚,幅,"
        gSL_Select = gSL_Select & "長さ,数量,処理区分,"
        gSL_Select = gSL_Select & "削除区分,変更回数,計画年月,"
        gSL_Select = gSL_Select & "作成者,作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.内示年月) & "',"		
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.手配先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.直送先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納期) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕掛日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"		
        gSL_Select = gSL_Select & "" & .寸法区分 & ","		
        gSL_Select = gSL_Select & "" & .板厚 & ","		
        gSL_Select = gSL_Select & "" & .幅 & ","		
        gSL_Select = gSL_Select & "" & .長さ & ","		
        gSL_Select = gSL_Select & "" & .数量 & ","		
        gSL_Select = gSL_Select & "" & .処理区分 & ","		
        gSL_Select = gSL_Select & "" & .削除区分 & ","		
        gSL_Select = gSL_Select & "" & .変更回数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.計画年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "',"		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZNJInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZNJUpdate(Optional DBNo As Integer) As Boolean
    ZNJUpdate = False	
    With ZNJ
        gSL_Select = "Update 材料発注内示テーブル  Set "
        gSL_Select = gSL_Select & "内示年月 = '" & RTrim(.内示年月) & "',"
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "手配先 = '" & RTrim(.手配先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "直送先 = '" & RTrim(.直送先) & "',"
        gSL_Select = gSL_Select & "納期 = '" & RTrim(.納期) & "',"
        gSL_Select = gSL_Select & "仕掛日 = '" & RTrim(.仕掛日) & "',"
        gSL_Select = gSL_Select & "手配日 = '" & RTrim(.手配日) & "',"
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "寸法区分 = " & .寸法区分 & ","
        gSL_Select = gSL_Select & "板厚 = " & .板厚 & ","
        gSL_Select = gSL_Select & "幅 = " & .幅 & ","
        gSL_Select = gSL_Select & "長さ = " & .長さ & ","
        gSL_Select = gSL_Select & "数量 = " & .数量 & ","
        gSL_Select = gSL_Select & "処理区分 = " & .処理区分 & ","
        gSL_Select = gSL_Select & "削除区分 = " & .削除区分 & ","
        gSL_Select = gSL_Select & "変更回数 = " & .変更回数 & ","
        gSL_Select = gSL_Select & "計画年月 = '" & RTrim(.計画年月) & "',"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 内示年月 = '" & RTrim(.内示年月) & "'"
        gSL_Select = gSL_Select & "   And 材料管理番号 = " & .材料管理番号 & ""
        gSL_Select = gSL_Select & "   And 手配先 = '" & RTrim(.手配先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 直送先 = '" & RTrim(.直送先) & "'"
        gSL_Select = gSL_Select & "   And 納期 = '" & RTrim(.納期) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZNJUpdate = True
End Function

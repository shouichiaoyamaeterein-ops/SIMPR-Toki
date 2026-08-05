Attribute VB_Name = "材料発注打切テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料発注打切テーブル
'*
'*   モジュールID  : 材料発注打切テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/02/19  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料発注打切テーブル
        打切種別      As Byte
        材質          As String
        板厚          As Currency
        幅            As Currency
        長さ          As Currency
        仕入先        As String
        材料管理番号  As Long
        起算日        As String
        起算便        As String
        設定数        As Currency
        警告数        As Currency
        発注数        As Currency
        入荷数        As Currency
        設定残数      As Currency
        完了日        As String
        完了便        As String
        打切日        As String
        警告日        As String
        切替仕入先    As String
        切替受入      As String
'        作成日        As Date
'        作成端末      As String
        作成者        As String
'        更新日        As Date
'        更新端末      As String
        更新者        As String
        件数          As Long
'   項目追加(自動)
	仕入先名      As String
    End Type

    Public ZUT As T_材料発注打切テーブル

    Public ZUT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZUT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZUT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZUT
            Select Case adoCol.Name
                Case "打切種別":      If Not IsNull(adoCol.Value) Then .打切種別     = adoCol.Value
                Case "材質":          If Not IsNull(adoCol.Value) Then .材質         = adoCol.Value
                Case "板厚":          If Not IsNull(adoCol.Value) Then .板厚         = adoCol.Value
                Case "幅":            If Not IsNull(adoCol.Value) Then .幅           = adoCol.Value
                Case "長さ":          If Not IsNull(adoCol.Value) Then .長さ         = adoCol.Value
                Case "仕入先":        If Not IsNull(adoCol.Value) Then .仕入先       = adoCol.Value
                Case "材料管理番号":  If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "起算日":        If Not IsNull(adoCol.Value) Then .起算日       = adoCol.Value
                Case "起算便":        If Not IsNull(adoCol.Value) Then .起算便       = adoCol.Value
                Case "設定数":        If Not IsNull(adoCol.Value) Then .設定数       = adoCol.Value
                Case "警告数":        If Not IsNull(adoCol.Value) Then .警告数       = adoCol.Value
                Case "発注数":        If Not IsNull(adoCol.Value) Then .発注数       = adoCol.Value
                Case "入荷数":        If Not IsNull(adoCol.Value) Then .入荷数       = adoCol.Value
                Case "設定残数":      If Not IsNull(adoCol.Value) Then .設定残数     = adoCol.Value
                Case "完了日":        If Not IsNull(adoCol.Value) Then .完了日       = adoCol.Value
                Case "完了便":        If Not IsNull(adoCol.Value) Then .完了便       = adoCol.Value
                Case "打切日":        If Not IsNull(adoCol.Value) Then .打切日       = adoCol.Value
                Case "警告日":        If Not IsNull(adoCol.Value) Then .警告日       = adoCol.Value
                Case "切替仕入先":    If Not IsNull(adoCol.Value) Then .切替仕入先   = adoCol.Value
                Case "切替受入":      If Not IsNull(adoCol.Value) Then .切替受入     = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者       = adoCol.Value
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
Public Function ZUTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZUTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 材料発注打切テーブル"
        RTxtSql = RTxtSql &" Where 打切種別 = " & ZUT.打切種別 & ""
        RTxtSql = RTxtSql &" And   材質 = '" & RTrim(ZUT.材質) & "'"
        RTxtSql = RTxtSql &" And   板厚 = " & ZUT.板厚 & ""
        RTxtSql = RTxtSql &" And   幅 = " & ZUT.幅 & ""
        RTxtSql = RTxtSql &" And   長さ = " & ZUT.長さ & ""
        RTxtSql = RTxtSql &" And   仕入先 = '" & RTrim(ZUT.仕入先) & "'"
        RTxtSql = RTxtSql &" And   材料管理番号 = " & ZUT.材料管理番号 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZUTRead = True
           Call ItemsFromZUT(DBNo)
	   ZUT_RDSTS = True
        else
	   ZUT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZUTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZUTReadNext(Optional DBNo As Integer = 0)
    If Not ZUT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZUT_RDSTS = False
        Else
           Call ItemsFromZUT(DBNo)
           ZUT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZUTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZUT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZUT )
'*----------------------------------------------*'
Public Sub ItemsClearZUT()
    With ZUT
	.打切種別 = 0
	.材質 = ""
	.板厚 = 0
	.幅 = 0
	.長さ = 0
	.仕入先 = ""
	.材料管理番号 = 0
	.起算日 = ""
	.起算便 = ""
	.設定数 = 0
	.警告数 = 0
	.発注数 = 0
	.入荷数 = 0
	.設定残数 = 0
	.完了日 = ""
	.完了便 = ""
	.打切日 = ""
	.警告日 = ""
	.切替仕入先 = ""
	.切替受入 = ""
'	.作成日 = 0
'	.作成端末 = ""
	.作成者 = ""
'	.更新日 = 0
'	.更新端末 = ""
	.更新者 = ""
        .件数 = 0
'   項目追加(自動)
	.仕入先名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZUTInsert(Optional DBNo As Integer) As Boolean
    ZUTInsert = True
    With ZUT
        gSL_Select = "Insert Into 材料発注打切テーブル  ( "
        gSL_Select = gSL_Select & "打切種別,材質,板厚,"
        gSL_Select = gSL_Select & "幅,長さ,仕入先,"
        gSL_Select = gSL_Select & "材料管理番号,起算日,起算便,"
        gSL_Select = gSL_Select & "設定数,警告数,発注数,"
        gSL_Select = gSL_Select & "入荷数,設定残数,完了日,"
        gSL_Select = gSL_Select & "完了便,打切日,警告日,"
        gSL_Select = gSL_Select & "切替仕入先,切替受入,"
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .打切種別 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"		
        gSL_Select = gSL_Select & "" & .板厚 & ","		
        gSL_Select = gSL_Select & "" & .幅 & ","		
        gSL_Select = gSL_Select & "" & .長さ & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"		
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.起算日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.起算便) & "',"		
        gSL_Select = gSL_Select & "" & .設定数 & ","		
        gSL_Select = gSL_Select & "" & .警告数 & ","		
        gSL_Select = gSL_Select & "" & .発注数 & ","		
        gSL_Select = gSL_Select & "" & .入荷数 & ","		
        gSL_Select = gSL_Select & "" & .設定残数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.完了日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.完了便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.打切日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.警告日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.切替仕入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.切替受入) & "',"		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20),"		
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZUTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZUTUpdate(Optional DBNo As Integer) As Boolean
    ZUTUpdate = False	
    With ZUT
        gSL_Select = "Update 材料発注打切テーブル  Set "
        gSL_Select = gSL_Select & "打切種別 = " & .打切種別 & ","
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "板厚 = " & .板厚 & ","
        gSL_Select = gSL_Select & "幅 = " & .幅 & ","
        gSL_Select = gSL_Select & "長さ = " & .長さ & ","
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "起算日 = '" & RTrim(.起算日) & "',"
        gSL_Select = gSL_Select & "起算便 = '" & RTrim(.起算便) & "',"
        gSL_Select = gSL_Select & "設定数 = " & .設定数 & ","
        gSL_Select = gSL_Select & "警告数 = " & .警告数 & ","
        gSL_Select = gSL_Select & "発注数 = " & .発注数 & ","
        gSL_Select = gSL_Select & "入荷数 = " & .入荷数 & ","
        gSL_Select = gSL_Select & "設定残数 = " & .設定残数 & ","
        gSL_Select = gSL_Select & "完了日 = '" & RTrim(.完了日) & "',"
        gSL_Select = gSL_Select & "完了便 = '" & RTrim(.完了便) & "',"
        gSL_Select = gSL_Select & "打切日 = '" & RTrim(.打切日) & "',"
        gSL_Select = gSL_Select & "警告日 = '" & RTrim(.警告日) & "',"
        gSL_Select = gSL_Select & "切替仕入先 = '" & RTrim(.切替仕入先) & "',"
        gSL_Select = gSL_Select & "切替受入 = '" & RTrim(.切替受入) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "'"
        gSL_Select = gSL_Select & " Where 打切種別 = " & .打切種別 & ""
        gSL_Select = gSL_Select & "   And 材質 = '" & RTrim(.材質) & "'"
        gSL_Select = gSL_Select & "   And 板厚 = " & .板厚 & ""
        gSL_Select = gSL_Select & "   And 幅 = " & .幅 & ""
        gSL_Select = gSL_Select & "   And 長さ = " & .長さ & ""
        gSL_Select = gSL_Select & "   And 仕入先 = '" & RTrim(.仕入先) & "'"
        gSL_Select = gSL_Select & "   And 材料管理番号 = " & .材料管理番号 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZUTUpdate = True
End Function

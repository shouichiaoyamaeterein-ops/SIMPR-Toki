Attribute VB_Name = "材料棚卸プレシート発行テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料棚卸プレシート発行テーブル
'*
'*   モジュールID  : 材料棚卸プレシート発行テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2010/01/12  By Cis(HP-DX7500)
'*   作成日  :  2010/01/13  By Cis(HP-DX7500)
'*   作成日  :  2010/01/14  By Cis(HP-DX7500)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料棚卸プレシート発行テーブル
        発行年月          As String
        伝票NO            As Long
        伝票行NO          As Long
        在庫場所1         As String
        在庫場所2         As String
        棚番              As String
        材料管理番号      As Long
        在庫数            As Long
        棚卸入力済フラグ  As Byte
'        作成日            As Date
        作成者            As String
'        作成端末          As String
        件数          As Long
'   項目追加(自動)
	材質           As String
	手配単位       As String
	材料名         As String
	板厚           As Currency
	幅             As Currency
	長さ           As Currency
	在庫場所名1    As String
	在庫場所名2    As String
	表示寸法       As String
	理論在庫数     As Currency
	最新単価       As Currency
	プレシートNO   As String
	置場           As String
	取引先         As String
	寸法区分       As Byte
	SIZE           As String
    End Type

    Public ZTP As T_材料棚卸プレシート発行テーブル

    Public ZTP_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZTP(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZTP	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZTP
            Select Case adoCol.Name
                Case "発行年月":          If Not IsNull(adoCol.Value) Then .発行年月         = adoCol.Value
                Case "伝票NO":            If Not IsNull(adoCol.Value) Then .伝票NO           = adoCol.Value
                Case "伝票行NO":          If Not IsNull(adoCol.Value) Then .伝票行NO         = adoCol.Value
                Case "在庫場所1":         If Not IsNull(adoCol.Value) Then .在庫場所1        = adoCol.Value
                Case "在庫場所2":         If Not IsNull(adoCol.Value) Then .在庫場所2        = adoCol.Value
                Case "棚番":              If Not IsNull(adoCol.Value) Then .棚番             = adoCol.Value
                Case "材料管理番号":      If Not IsNull(adoCol.Value) Then .材料管理番号     = adoCol.Value
                Case "在庫数":            If Not IsNull(adoCol.Value) Then .在庫数           = adoCol.Value
                Case "棚卸入力済フラグ":  If Not IsNull(adoCol.Value) Then .棚卸入力済フラグ = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "材質":	If Not IsNull(adoCol.Value) Then .材質	= adoCol.Value
		Case "手配単位":	If Not IsNull(adoCol.Value) Then .手配単位	= adoCol.Value
		Case "材料名":	If Not IsNull(adoCol.Value) Then .材料名	= adoCol.Value
		Case "板厚":	If Not IsNull(adoCol.Value) Then .板厚	= adoCol.Value
		Case "幅":	If Not IsNull(adoCol.Value) Then .幅	= adoCol.Value
		Case "長さ":	If Not IsNull(adoCol.Value) Then .長さ	= adoCol.Value
		Case "在庫場所名1":	If Not IsNull(adoCol.Value) Then .在庫場所名1	= adoCol.Value
		Case "在庫場所名2":	If Not IsNull(adoCol.Value) Then .在庫場所名2	= adoCol.Value
		Case "表示寸法":	If Not IsNull(adoCol.Value) Then .表示寸法	= adoCol.Value
		Case "理論在庫数":	If Not IsNull(adoCol.Value) Then .理論在庫数	= adoCol.Value
		Case "最新単価":	If Not IsNull(adoCol.Value) Then .最新単価	= adoCol.Value
		Case "プレシートNO":	If Not IsNull(adoCol.Value) Then .プレシートNO	= adoCol.Value
		Case "置場":	If Not IsNull(adoCol.Value) Then .置場	= adoCol.Value
		Case "取引先":	If Not IsNull(adoCol.Value) Then .取引先	= adoCol.Value
		Case "寸法区分":	If Not IsNull(adoCol.Value) Then .寸法区分	= adoCol.Value
		Case "SIZE":	If Not IsNull(adoCol.Value) Then .SIZE	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZTPRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZTPRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 材料棚卸プレシート発行テーブル"
        RTxtSql = RTxtSql &" Where 発行年月 = '" & RTrim(ZTP.発行年月) & "'"
        RTxtSql = RTxtSql &" And   伝票NO = " & ZTP.伝票NO & ""
        RTxtSql = RTxtSql &" And   伝票行NO = " & ZTP.伝票行NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZTPRead = True
           Call ItemsFromZTP(DBNo)
	   ZTP_RDSTS = True
        else
	   ZTP_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZTPClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZTPReadNext(Optional DBNo As Integer = 0)
    If Not ZTP_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZTP_RDSTS = False
        Else
           Call ItemsFromZTP(DBNo)
           ZTP_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZTPClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZTP_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZTP )
'*----------------------------------------------*'
Public Sub ItemsClearZTP()
    With ZTP
	.発行年月 = ""
	.伝票NO = 0
	.伝票行NO = 0
	.在庫場所1 = ""
	.在庫場所2 = ""
	.棚番 = ""
	.材料管理番号 = 0
	.在庫数 = 0
	.棚卸入力済フラグ = 0
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
        .件数 = 0
'   項目追加(自動)
	.材質 = ""
	.手配単位 = ""
	.材料名 = ""
	.板厚 = 0
	.幅 = 0
	.長さ = 0
	.在庫場所名1 = ""
	.在庫場所名2 = ""
	.表示寸法 = ""
	.理論在庫数 = 0
	.最新単価 = 0
	.プレシートNO = ""
	.置場 = ""
	.取引先 = ""
	.寸法区分 = 0
	.SIZE = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZTPInsert(Optional DBNo As Integer) As Boolean
    ZTPInsert = False
    With ZTP
        gSL_Select = "Insert Into 材料棚卸プレシート発行テーブル  ( "
        gSL_Select = gSL_Select & "発行年月,伝票NO,伝票行NO,"
        gSL_Select = gSL_Select & "在庫場所1,在庫場所2,棚番,"
        gSL_Select = gSL_Select & "材料管理番号,在庫数,棚卸入力済フラグ,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.発行年月) & "',"		
        gSL_Select = gSL_Select & "" & .伝票NO & ","		
        gSL_Select = gSL_Select & "" & .伝票行NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.在庫場所1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.在庫場所2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.棚番) & "',"		
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","		
        gSL_Select = gSL_Select & "" & .在庫数 & ","		
        gSL_Select = gSL_Select & "" & .棚卸入力済フラグ & ","		
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
    ZTPInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZTPUpdate(Optional DBNo As Integer) As Boolean
    ZTPUpdate = False	
    With ZTP
        gSL_Select = "Update 材料棚卸プレシート発行テーブル  Set "
        gSL_Select = gSL_Select & "発行年月 = '" & RTrim(.発行年月) & "',"
        gSL_Select = gSL_Select & "伝票NO = " & .伝票NO & ","
        gSL_Select = gSL_Select & "伝票行NO = " & .伝票行NO & ","
        gSL_Select = gSL_Select & "在庫場所1 = '" & RTrim(.在庫場所1) & "',"
        gSL_Select = gSL_Select & "在庫場所2 = '" & RTrim(.在庫場所2) & "',"
        gSL_Select = gSL_Select & "棚番 = '" & RTrim(.棚番) & "',"
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "在庫数 = " & .在庫数 & ","
        gSL_Select = gSL_Select & "棚卸入力済フラグ = " & .棚卸入力済フラグ & ""
        gSL_Select = gSL_Select & " Where 発行年月 = '" & RTrim(.発行年月) & "'"
        gSL_Select = gSL_Select & "   And 伝票NO = " & .伝票NO & ""
        gSL_Select = gSL_Select & "   And 伝票行NO = " & .伝票行NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZTPUpdate = True
End Function

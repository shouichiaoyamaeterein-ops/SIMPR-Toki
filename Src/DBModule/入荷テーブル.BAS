Attribute VB_Name = "入荷テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   入荷テーブル
'*
'*   モジュールID  : 入荷テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2005/06/14  By Cis
'*   作成日  :  2006/02/10  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_入荷テーブル
        発注管理NO            As Long
        納入番号              As String
        納品書頁              As Byte
        納品書行              As Byte
        分納区分              As Byte
        仕入先                As String
        受入                  As String
        品番                  As String
        背番号                As String
        入荷日                As String
        入荷便                As String
        収容数                As Long
        入荷枚数              As Long
        入荷端数              As Long
        入荷数                As Long
        分割数                As Long
        挽回日                As String
        挽回便                As String
        分割納品書NO          As String
        検収区分              As String
        検収年月              As String
        検収処理日            As String
        支給検収区分          As String
        支給検収年月          As String
        支給検収処理日        As String
        TNS作成区分           As Byte
        TNS作成管理NO         As Long
        TNS作成日             As String
        TNS作成処理日         As Date
        TNS連番               As Long
        TNS作成区分2          As Byte
        TNS作成管理NO2        As Long
        TNS作成日2            As String
        TNS作成処理日2        As Date
        TNS連番2              As Long
'        作成日                As Date
'        作成端末              As String
        作成者                As String
'        更新日                As Date
'        更新端末              As String
        更新者                As String
        件数          As Long
'   項目追加(自動)
	仕入先名       As String
	取引先         As String
	工場CD         As String
	工場名         As String
	表示品番       As String
	契約先         As String
	カード         As String
	納入日         As String
	納入便         As String
	支給先         As String
	支給先受入     As String
	支給元         As String
	契約先名       As String
	自社工区名     As String
	車型           As String
	予測金額       As String
	実績金額       As String
	進捗           As Currency
	実績金額0      As Currency
	実績金額1      As Currency
	実績金額2      As Currency
	実績金額3      As Currency
	実績金額4      As Currency
	実績金額5      As Currency
	日当り数       As Currency
    End Type

    Public NYT As T_入荷テーブル

    Public NYT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromNYT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearNYT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With NYT
            Select Case adoCol.Name
                Case "発注管理NO":            If Not IsNull(adoCol.Value) Then .発注管理NO           = adoCol.Value
                Case "納入番号":              If Not IsNull(adoCol.Value) Then .納入番号             = adoCol.Value
                Case "納品書頁":              If Not IsNull(adoCol.Value) Then .納品書頁             = adoCol.Value
                Case "納品書行":              If Not IsNull(adoCol.Value) Then .納品書行             = adoCol.Value
                Case "分納区分":              If Not IsNull(adoCol.Value) Then .分納区分             = adoCol.Value
                Case "仕入先":                If Not IsNull(adoCol.Value) Then .仕入先               = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入                 = adoCol.Value
                Case "品番":                  If Not IsNull(adoCol.Value) Then .品番                 = adoCol.Value
                Case "背番号":                If Not IsNull(adoCol.Value) Then .背番号               = adoCol.Value
                Case "入荷日":                If Not IsNull(adoCol.Value) Then .入荷日               = adoCol.Value
                Case "入荷便":                If Not IsNull(adoCol.Value) Then .入荷便               = adoCol.Value
                Case "収容数":                If Not IsNull(adoCol.Value) Then .収容数               = adoCol.Value
                Case "入荷枚数":              If Not IsNull(adoCol.Value) Then .入荷枚数             = adoCol.Value
                Case "入荷端数":              If Not IsNull(adoCol.Value) Then .入荷端数             = adoCol.Value
                Case "入荷数":                If Not IsNull(adoCol.Value) Then .入荷数               = adoCol.Value
                Case "分割数":                If Not IsNull(adoCol.Value) Then .分割数               = adoCol.Value
                Case "挽回日":                If Not IsNull(adoCol.Value) Then .挽回日               = adoCol.Value
                Case "挽回便":                If Not IsNull(adoCol.Value) Then .挽回便               = adoCol.Value
                Case "分割納品書NO":          If Not IsNull(adoCol.Value) Then .分割納品書NO         = adoCol.Value
                Case "検収区分":              If Not IsNull(adoCol.Value) Then .検収区分             = adoCol.Value
                Case "検収年月":              If Not IsNull(adoCol.Value) Then .検収年月             = adoCol.Value
                Case "検収処理日":            If Not IsNull(adoCol.Value) Then .検収処理日           = adoCol.Value
                Case "支給検収区分":          If Not IsNull(adoCol.Value) Then .支給検収区分         = adoCol.Value
                Case "支給検収年月":          If Not IsNull(adoCol.Value) Then .支給検収年月         = adoCol.Value
                Case "支給検収処理日":        If Not IsNull(adoCol.Value) Then .支給検収処理日       = adoCol.Value
                Case "TNS作成区分":           If Not IsNull(adoCol.Value) Then .TNS作成区分          = adoCol.Value
                Case "TNS作成管理NO":         If Not IsNull(adoCol.Value) Then .TNS作成管理NO        = adoCol.Value
                Case "TNS作成日":             If Not IsNull(adoCol.Value) Then .TNS作成日            = adoCol.Value
                Case "TNS作成処理日":         If Not IsNull(adoCol.Value) Then .TNS作成処理日        = adoCol.Value
                Case "TNS連番":               If Not IsNull(adoCol.Value) Then .TNS連番              = adoCol.Value
                Case "TNS作成区分2":          If Not IsNull(adoCol.Value) Then .TNS作成区分2         = adoCol.Value
                Case "TNS作成管理NO2":        If Not IsNull(adoCol.Value) Then .TNS作成管理NO2       = adoCol.Value
                Case "TNS作成日2":            If Not IsNull(adoCol.Value) Then .TNS作成日2           = adoCol.Value
                Case "TNS作成処理日2":        If Not IsNull(adoCol.Value) Then .TNS作成処理日2       = adoCol.Value
                Case "TNS連番2":              If Not IsNull(adoCol.Value) Then .TNS連番2             = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "作成者":                If Not IsNull(adoCol.Value) Then .作成者               = adoCol.Value
'                Case "更新日":                If Not IsNull(adoCol.Value) Then .更新日               = adoCol.Value
'                Case "更新端末":              If Not IsNull(adoCol.Value) Then .更新端末             = adoCol.Value
                Case "更新者":                If Not IsNull(adoCol.Value) Then .更新者               = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "仕入先名":	If Not IsNull(adoCol.Value) Then .仕入先名	= adoCol.Value
		Case "取引先":	If Not IsNull(adoCol.Value) Then .取引先	= adoCol.Value
		Case "工場CD":	If Not IsNull(adoCol.Value) Then .工場CD	= adoCol.Value
		Case "工場名":	If Not IsNull(adoCol.Value) Then .工場名	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "契約先":	If Not IsNull(adoCol.Value) Then .契約先	= adoCol.Value
		Case "カード":	If Not IsNull(adoCol.Value) Then .カード	= adoCol.Value
		Case "納入日":	If Not IsNull(adoCol.Value) Then .納入日	= adoCol.Value
		Case "納入便":	If Not IsNull(adoCol.Value) Then .納入便	= adoCol.Value
		Case "支給先":	If Not IsNull(adoCol.Value) Then .支給先	= adoCol.Value
		Case "支給先受入":	If Not IsNull(adoCol.Value) Then .支給先受入	= adoCol.Value
		Case "支給元":	If Not IsNull(adoCol.Value) Then .支給元	= adoCol.Value
		Case "契約先名":	If Not IsNull(adoCol.Value) Then .契約先名	= adoCol.Value
		Case "自社工区名":	If Not IsNull(adoCol.Value) Then .自社工区名	= adoCol.Value
		Case "車型":	If Not IsNull(adoCol.Value) Then .車型	= adoCol.Value
		Case "予測金額":	If Not IsNull(adoCol.Value) Then .予測金額	= adoCol.Value
		Case "実績金額":	If Not IsNull(adoCol.Value) Then .実績金額	= adoCol.Value
		Case "進捗":	If Not IsNull(adoCol.Value) Then .進捗	= adoCol.Value
		Case "実績金額0":	If Not IsNull(adoCol.Value) Then .実績金額0	= adoCol.Value
		Case "実績金額1":	If Not IsNull(adoCol.Value) Then .実績金額1	= adoCol.Value
		Case "実績金額2":	If Not IsNull(adoCol.Value) Then .実績金額2	= adoCol.Value
		Case "実績金額3":	If Not IsNull(adoCol.Value) Then .実績金額3	= adoCol.Value
		Case "実績金額4":	If Not IsNull(adoCol.Value) Then .実績金額4	= adoCol.Value
		Case "実績金額5":	If Not IsNull(adoCol.Value) Then .実績金額5	= adoCol.Value
		Case "日当り数":	If Not IsNull(adoCol.Value) Then .日当り数	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function NYTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    NYTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 入荷テーブル"
        RTxtSql = RTxtSql &" Where 発注管理NO = " & NYT.発注管理NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           NYTRead = True
           Call ItemsFromNYT(DBNo)
	   NYT_RDSTS = True
        else
	   NYT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call NYTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub NYTReadNext(Optional DBNo As Integer = 0)
    If Not NYT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           NYT_RDSTS = False
        Else
           Call ItemsFromNYT(DBNo)
           NYT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub NYTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        NYT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( NYT )
'*----------------------------------------------*'
Public Sub ItemsClearNYT()
    With NYT
	.発注管理NO = 0
	.納入番号 = ""
	.納品書頁 = 0
	.納品書行 = 0
	.分納区分 = 0
	.仕入先 = ""
	.受入 = ""
	.品番 = ""
	.背番号 = ""
	.入荷日 = ""
	.入荷便 = ""
	.収容数 = 0
	.入荷枚数 = 0
	.入荷端数 = 0
	.入荷数 = 0
	.分割数 = 0
	.挽回日 = ""
	.挽回便 = ""
	.分割納品書NO = ""
	.検収区分 = ""
	.検収年月 = ""
	.検収処理日 = ""
	.支給検収区分 = ""
	.支給検収年月 = ""
	.支給検収処理日 = ""
	.TNS作成区分 = 0
	.TNS作成管理NO = 0
	.TNS作成日 = ""
	.TNS作成処理日 = 0
	.TNS連番 = 0
	.TNS作成区分2 = 0
	.TNS作成管理NO2 = 0
	.TNS作成日2 = ""
	.TNS作成処理日2 = 0
	.TNS連番2 = 0
'	.作成日 = 0
'	.作成端末 = ""
	.作成者 = ""
'	.更新日 = 0
'	.更新端末 = ""
	.更新者 = ""
        .件数 = 0
'   項目追加(自動)
	.仕入先名 = ""
	.取引先 = ""
	.工場CD = ""
	.工場名 = ""
	.表示品番 = ""
	.契約先 = ""
	.カード = ""
	.納入日 = ""
	.納入便 = ""
	.支給先 = ""
	.支給先受入 = ""
	.支給元 = ""
	.契約先名 = ""
	.自社工区名 = ""
	.車型 = ""
	.予測金額 = ""
	.実績金額 = ""
	.進捗 = 0
	.実績金額0 = 0
	.実績金額1 = 0
	.実績金額2 = 0
	.実績金額3 = 0
	.実績金額4 = 0
	.実績金額5 = 0
	.日当り数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function NYTInsert(Optional DBNo As Integer) As Boolean
    NYTInsert = False
    With NYT
        gSL_Select = "Insert Into 入荷テーブル  ( "
        gSL_Select = gSL_Select & "発注管理NO,納入番号,納品書頁,"
        gSL_Select = gSL_Select & "納品書行,分納区分,仕入先,"
        gSL_Select = gSL_Select & "受入,品番,背番号,"
        gSL_Select = gSL_Select & "入荷日,入荷便,収容数,"
        gSL_Select = gSL_Select & "入荷枚数,入荷端数,入荷数,"
        gSL_Select = gSL_Select & "分割数,挽回日,挽回便,"
        gSL_Select = gSL_Select & "分割納品書NO,検収区分,検収年月,"
        gSL_Select = gSL_Select & "検収処理日,支給検収区分,支給検収年月,"
        gSL_Select = gSL_Select & "支給検収処理日,TNS作成区分,TNS作成管理NO,"
        gSL_Select = gSL_Select & "TNS作成日,TNS作成処理日,TNS連番,"
        gSL_Select = gSL_Select & "TNS作成区分2,TNS作成管理NO2,TNS作成日2,"
        gSL_Select = gSL_Select & "TNS作成処理日2,TNS連番2,"
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .発注管理NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"		
        gSL_Select = gSL_Select & "" & .納品書頁 & ","		
        gSL_Select = gSL_Select & "" & .納品書行 & ","		
        gSL_Select = gSL_Select & "" & .分納区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.入荷日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.入荷便) & "',"		
        gSL_Select = gSL_Select & "" & .収容数 & ","		
        gSL_Select = gSL_Select & "" & .入荷枚数 & ","		
        gSL_Select = gSL_Select & "" & .入荷端数 & ","		
        gSL_Select = gSL_Select & "" & .入荷数 & ","		
        gSL_Select = gSL_Select & "" & .分割数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.挽回日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.挽回便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.分割納品書NO) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.検収区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.検収年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.検収処理日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給検収区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給検収年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給検収処理日) & "',"		
        gSL_Select = gSL_Select & "" & .TNS作成区分 & ","		
        gSL_Select = gSL_Select & "" & .TNS作成管理NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.TNS作成日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.TNS作成処理日) & "',"		
        gSL_Select = gSL_Select & "" & .TNS連番 & ","		
        gSL_Select = gSL_Select & "" & .TNS作成区分2 & ","		
        gSL_Select = gSL_Select & "" & .TNS作成管理NO2 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.TNS作成日2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.TNS作成処理日2) & "',"		
        gSL_Select = gSL_Select & "" & .TNS連番2 & ","		
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
    NYTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function NYTUpdate(Optional DBNo As Integer) As Boolean
    NYTUpdate = False	
    With NYT
        gSL_Select = "Update 入荷テーブル  Set "
        gSL_Select = gSL_Select & "発注管理NO = " & .発注管理NO & ","
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "納品書頁 = " & .納品書頁 & ","
        gSL_Select = gSL_Select & "納品書行 = " & .納品書行 & ","
        gSL_Select = gSL_Select & "分納区分 = " & .分納区分 & ","
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "入荷日 = '" & RTrim(.入荷日) & "',"
        gSL_Select = gSL_Select & "入荷便 = '" & RTrim(.入荷便) & "',"
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "入荷枚数 = " & .入荷枚数 & ","
        gSL_Select = gSL_Select & "入荷端数 = " & .入荷端数 & ","
        gSL_Select = gSL_Select & "入荷数 = " & .入荷数 & ","
        gSL_Select = gSL_Select & "分割数 = " & .分割数 & ","
        gSL_Select = gSL_Select & "挽回日 = '" & RTrim(.挽回日) & "',"
        gSL_Select = gSL_Select & "挽回便 = '" & RTrim(.挽回便) & "',"
        gSL_Select = gSL_Select & "分割納品書NO = '" & RTrim(.分割納品書NO) & "',"
        gSL_Select = gSL_Select & "検収区分 = '" & RTrim(.検収区分) & "',"
        gSL_Select = gSL_Select & "検収年月 = '" & RTrim(.検収年月) & "',"
        gSL_Select = gSL_Select & "検収処理日 = '" & RTrim(.検収処理日) & "',"
        gSL_Select = gSL_Select & "支給検収区分 = '" & RTrim(.支給検収区分) & "',"
        gSL_Select = gSL_Select & "支給検収年月 = '" & RTrim(.支給検収年月) & "',"
        gSL_Select = gSL_Select & "支給検収処理日 = '" & RTrim(.支給検収処理日) & "',"
        gSL_Select = gSL_Select & "TNS作成区分 = " & .TNS作成区分 & ","
        gSL_Select = gSL_Select & "TNS作成管理NO = " & .TNS作成管理NO & ","
        gSL_Select = gSL_Select & "TNS作成日 = '" & RTrim(.TNS作成日) & "',"
        gSL_Select = gSL_Select & "TNS作成処理日 = '" & RTrim(.TNS作成処理日) & "',"
        gSL_Select = gSL_Select & "TNS連番 = " & .TNS連番 & ","
        gSL_Select = gSL_Select & "TNS作成区分2 = " & .TNS作成区分2 & ","
        gSL_Select = gSL_Select & "TNS作成管理NO2 = " & .TNS作成管理NO2 & ","
        gSL_Select = gSL_Select & "TNS作成日2 = '" & RTrim(.TNS作成日2) & "',"
        gSL_Select = gSL_Select & "TNS作成処理日2 = '" & RTrim(.TNS作成処理日2) & "',"
        gSL_Select = gSL_Select & "TNS連番2 = " & .TNS連番2 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "'"
        gSL_Select = gSL_Select & " Where 発注管理NO = " & .発注管理NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    NYTUpdate = True
End Function

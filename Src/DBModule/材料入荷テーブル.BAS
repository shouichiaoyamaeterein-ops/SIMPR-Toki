Attribute VB_Name = "材料入荷テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料入荷テーブル
'*
'*   モジュールID  : 材料入荷テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/02/26  By Cis
'*   作成日  :  2009/01/20  By Cis(FMV-E630)
'*   作成日  :  2009/02/12  By Cis(HP-DC7900)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料入荷テーブル
        発注管理NO            As Long
        納入番号              As String
        納品書頁              As Byte
        納品書行              As Byte
        分納区分              As Byte
        仕入先                As String
        受入                  As String
        材質                  As String
        板厚                  As Currency
        幅                    As Currency
        長さ                  As Currency
        材料管理番号          As Long
        背番号                As String
        入荷日                As String
        入荷便                As String
        収容数                As Currency
        入荷枚数              As Long
        入荷端数              As Currency
        入荷数                As Currency
        分割数                As Currency
        挽回日                As String
        挽回便                As String
        入荷備考              As String
        分割納品書NO          As String
        検収区分              As String
        検収年月              As String
        検収処理日            As String
        支給検収区分          As String
        支給検収年月          As String
        支給検収処理日        As String
        支給納入番号          As String
        支給実績管理NO        As Long
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
	契約先         As String
	表示寸法       As String
	単位           As String
	契約先名       As String
	自社工区名     As String
	車型           As String
	予測金額       As Currency
	実績金額       As Currency
	進捗           As Currency
	実績金額0      As Currency
	実績金額1      As Currency
	実績金額2      As Currency
	実績金額3      As Currency
	実績金額4      As Currency
	実績金額5      As Currency
	日当り数       As Currency
    End Type

    Public ZYT As T_材料入荷テーブル

    Public ZYT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZYT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZYT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZYT
            Select Case adoCol.Name
                Case "発注管理NO":            If Not IsNull(adoCol.Value) Then .発注管理NO           = adoCol.Value
                Case "納入番号":              If Not IsNull(adoCol.Value) Then .納入番号             = adoCol.Value
                Case "納品書頁":              If Not IsNull(adoCol.Value) Then .納品書頁             = adoCol.Value
                Case "納品書行":              If Not IsNull(adoCol.Value) Then .納品書行             = adoCol.Value
                Case "分納区分":              If Not IsNull(adoCol.Value) Then .分納区分             = adoCol.Value
                Case "仕入先":                If Not IsNull(adoCol.Value) Then .仕入先               = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入                 = adoCol.Value
                Case "材質":                  If Not IsNull(adoCol.Value) Then .材質                 = adoCol.Value
                Case "板厚":                  If Not IsNull(adoCol.Value) Then .板厚                 = adoCol.Value
                Case "幅":                    If Not IsNull(adoCol.Value) Then .幅                   = adoCol.Value
                Case "長さ":                  If Not IsNull(adoCol.Value) Then .長さ                 = adoCol.Value
                Case "材料管理番号":          If Not IsNull(adoCol.Value) Then .材料管理番号         = adoCol.Value
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
                Case "入荷備考":              If Not IsNull(adoCol.Value) Then .入荷備考             = adoCol.Value
                Case "分割納品書NO":          If Not IsNull(adoCol.Value) Then .分割納品書NO         = adoCol.Value
                Case "検収区分":              If Not IsNull(adoCol.Value) Then .検収区分             = adoCol.Value
                Case "検収年月":              If Not IsNull(adoCol.Value) Then .検収年月             = adoCol.Value
                Case "検収処理日":            If Not IsNull(adoCol.Value) Then .検収処理日           = adoCol.Value
                Case "支給検収区分":          If Not IsNull(adoCol.Value) Then .支給検収区分         = adoCol.Value
                Case "支給検収年月":          If Not IsNull(adoCol.Value) Then .支給検収年月         = adoCol.Value
                Case "支給検収処理日":        If Not IsNull(adoCol.Value) Then .支給検収処理日       = adoCol.Value
                Case "支給納入番号":          If Not IsNull(adoCol.Value) Then .支給納入番号         = adoCol.Value
                Case "支給実績管理NO":        If Not IsNull(adoCol.Value) Then .支給実績管理NO       = adoCol.Value
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
		Case "契約先":	If Not IsNull(adoCol.Value) Then .契約先	= adoCol.Value
		Case "表示寸法":	If Not IsNull(adoCol.Value) Then .表示寸法	= adoCol.Value
		Case "単位":	If Not IsNull(adoCol.Value) Then .単位	= adoCol.Value
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
Public Function ZYTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZYTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 材料入荷テーブル"
        RTxtSql = RTxtSql &" Where 発注管理NO = " & ZYT.発注管理NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZYTRead = True
           Call ItemsFromZYT(DBNo)
	   ZYT_RDSTS = True
        else
	   ZYT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZYTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZYTReadNext(Optional DBNo As Integer = 0)
    If Not ZYT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZYT_RDSTS = False
        Else
           Call ItemsFromZYT(DBNo)
           ZYT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZYTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZYT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZYT )
'*----------------------------------------------*'
Public Sub ItemsClearZYT()
    With ZYT
	.発注管理NO = 0
	.納入番号 = ""
	.納品書頁 = 0
	.納品書行 = 0
	.分納区分 = 0
	.仕入先 = ""
	.受入 = ""
	.材質 = ""
	.板厚 = 0
	.幅 = 0
	.長さ = 0
	.材料管理番号 = 0
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
	.入荷備考 = ""
	.分割納品書NO = ""
	.検収区分 = ""
	.検収年月 = ""
	.検収処理日 = ""
	.支給検収区分 = ""
	.支給検収年月 = ""
	.支給検収処理日 = ""
	.支給納入番号 = ""
	.支給実績管理NO = 0
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
	.契約先 = ""
	.表示寸法 = ""
	.単位 = ""
	.契約先名 = ""
	.自社工区名 = ""
	.車型 = ""
	.予測金額 = 0
	.実績金額 = 0
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
Public Function ZYTInsert(Optional DBNo As Integer) As Boolean
    ZYTInsert = False
    With ZYT
        gSL_Select = "Insert Into 材料入荷テーブル  ( "
        gSL_Select = gSL_Select & "発注管理NO,納入番号,納品書頁,"
        gSL_Select = gSL_Select & "納品書行,分納区分,仕入先,"
        gSL_Select = gSL_Select & "受入,材質,板厚,"
        gSL_Select = gSL_Select & "幅,長さ,材料管理番号,"
        gSL_Select = gSL_Select & "背番号,入荷日,入荷便,"
        gSL_Select = gSL_Select & "収容数,入荷枚数,入荷端数,"
        gSL_Select = gSL_Select & "入荷数,分割数,挽回日,"
        gSL_Select = gSL_Select & "挽回便,入荷備考,分割納品書NO,"
        gSL_Select = gSL_Select & "検収区分,検収年月,検収処理日,"
        gSL_Select = gSL_Select & "支給検収区分,支給検収年月,支給検収処理日,"
        gSL_Select = gSL_Select & "支給納入番号,支給実績管理NO,"
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .発注管理NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"		
        gSL_Select = gSL_Select & "" & .納品書頁 & ","		
        gSL_Select = gSL_Select & "" & .納品書行 & ","		
        gSL_Select = gSL_Select & "" & .分納区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"		
        gSL_Select = gSL_Select & "" & .板厚 & ","		
        gSL_Select = gSL_Select & "" & .幅 & ","		
        gSL_Select = gSL_Select & "" & .長さ & ","		
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","		
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
        gSL_Select = gSL_Select & "'" & RTrim(.入荷備考) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.分割納品書NO) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.検収区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.検収年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.検収処理日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給検収区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給検収年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給検収処理日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給納入番号) & "',"		
        gSL_Select = gSL_Select & "" & .支給実績管理NO & ","		
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
    ZYTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZYTUpdate(Optional DBNo As Integer) As Boolean
    ZYTUpdate = False	
    With ZYT
        gSL_Select = "Update 材料入荷テーブル  Set "
        gSL_Select = gSL_Select & "発注管理NO = " & .発注管理NO & ","
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "納品書頁 = " & .納品書頁 & ","
        gSL_Select = gSL_Select & "納品書行 = " & .納品書行 & ","
        gSL_Select = gSL_Select & "分納区分 = " & .分納区分 & ","
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "板厚 = " & .板厚 & ","
        gSL_Select = gSL_Select & "幅 = " & .幅 & ","
        gSL_Select = gSL_Select & "長さ = " & .長さ & ","
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
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
        gSL_Select = gSL_Select & "入荷備考 = '" & RTrim(.入荷備考) & "',"
        gSL_Select = gSL_Select & "分割納品書NO = '" & RTrim(.分割納品書NO) & "',"
        gSL_Select = gSL_Select & "検収区分 = '" & RTrim(.検収区分) & "',"
        gSL_Select = gSL_Select & "検収年月 = '" & RTrim(.検収年月) & "',"
        gSL_Select = gSL_Select & "検収処理日 = '" & RTrim(.検収処理日) & "',"
        gSL_Select = gSL_Select & "支給検収区分 = '" & RTrim(.支給検収区分) & "',"
        gSL_Select = gSL_Select & "支給検収年月 = '" & RTrim(.支給検収年月) & "',"
        gSL_Select = gSL_Select & "支給検収処理日 = '" & RTrim(.支給検収処理日) & "',"
        gSL_Select = gSL_Select & "支給納入番号 = '" & RTrim(.支給納入番号) & "',"
        gSL_Select = gSL_Select & "支給実績管理NO = " & .支給実績管理NO & ","
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
    ZYTUpdate = True
End Function

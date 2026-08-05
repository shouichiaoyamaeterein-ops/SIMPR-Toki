Attribute VB_Name = "材料発注テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料発注テーブル
'*
'*   モジュールID  : 材料発注テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/02/26  By Cis
'*   作成日  :  2008/07/03  By Cis(FMV-E630)
'*   作成日  :  2009/01/20  By Cis(FMV-E630)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料発注テーブル
        発注管理NO            As Long
        納品書管理NO          As Long
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
        納入日                As String
        納入便                As String
        収容数                As Currency
        枚数                  As Long
        端数                  As Currency
        発注数                As Currency
        発注区分              As String
        手配区分              As String
        SK区分                As Byte
        契約先                As String
        納入先                As String
        納入先受入            As String
        サイクル              As String
        担当者                As String
        発注日                As String
        発注便                As String
        端数区分              As Byte
        打切区分              As Byte
        打切種別              As Byte
        打切残数              As Currency
        打切日                As String
        内示超過区分          As Byte
        納品書発行日          As String
        支給書発行日          As String
        受領書発行日          As String
        明細書発行日          As String
        明細書NO              As Long
        初回発注管理NO        As Long
        作成種別              As String
'        作成日                As Date
'        作成端末              As String
        作成者                As String
'        更新日                As Date
'        更新端末              As String
        更新者                As String
        件数          As Long
'   項目追加(自動)
	仕入先名       As String
	発注区分名     As String
	手配区分名     As String
	仕入先CD       As String
	仕入先工場CD   As String
	仕入先工場名   As String
	表示寸法       As String
	材種コード     As String
	手配単位名     As String
	材料名         As String
	収容器         As String
	納入先名       As String
	納入先工場名   As String
	材種名         As String
	納入先CD       As String
	納入先工場CD   As String
	置場           As String
	実数区分       As Byte
	入荷日         As String
	入荷便         As String
	挽回日         As String
	挽回便         As String
	分割数         As Currency
	入荷枚数       As Long
	入荷端数       As Currency
	入荷数         As Currency
	入荷備考       As String
	自社工区名     As String
	品番           As String
	単価           As String
	単重           As String
	金額           As Currency
	支給実績管理NO As Long
    End Type

    Public ZCT As T_材料発注テーブル

    Public ZCT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZCT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZCT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZCT
            Select Case adoCol.Name
                Case "発注管理NO":            If Not IsNull(adoCol.Value) Then .発注管理NO           = adoCol.Value
                Case "納品書管理NO":          If Not IsNull(adoCol.Value) Then .納品書管理NO         = adoCol.Value
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
                Case "納入日":                If Not IsNull(adoCol.Value) Then .納入日               = adoCol.Value
                Case "納入便":                If Not IsNull(adoCol.Value) Then .納入便               = adoCol.Value
                Case "収容数":                If Not IsNull(adoCol.Value) Then .収容数               = adoCol.Value
                Case "枚数":                  If Not IsNull(adoCol.Value) Then .枚数                 = adoCol.Value
                Case "端数":                  If Not IsNull(adoCol.Value) Then .端数                 = adoCol.Value
                Case "発注数":                If Not IsNull(adoCol.Value) Then .発注数               = adoCol.Value
                Case "発注区分":              If Not IsNull(adoCol.Value) Then .発注区分             = adoCol.Value
                Case "手配区分":              If Not IsNull(adoCol.Value) Then .手配区分             = adoCol.Value
                Case "SK区分":                If Not IsNull(adoCol.Value) Then .SK区分               = adoCol.Value
                Case "契約先":                If Not IsNull(adoCol.Value) Then .契約先               = adoCol.Value
                Case "納入先":                If Not IsNull(adoCol.Value) Then .納入先               = adoCol.Value
                Case "納入先受入":            If Not IsNull(adoCol.Value) Then .納入先受入           = adoCol.Value
                Case "サイクル":              If Not IsNull(adoCol.Value) Then .サイクル             = adoCol.Value
                Case "担当者":                If Not IsNull(adoCol.Value) Then .担当者               = adoCol.Value
                Case "発注日":                If Not IsNull(adoCol.Value) Then .発注日               = adoCol.Value
                Case "発注便":                If Not IsNull(adoCol.Value) Then .発注便               = adoCol.Value
                Case "端数区分":              If Not IsNull(adoCol.Value) Then .端数区分             = adoCol.Value
                Case "打切区分":              If Not IsNull(adoCol.Value) Then .打切区分             = adoCol.Value
                Case "打切種別":              If Not IsNull(adoCol.Value) Then .打切種別             = adoCol.Value
                Case "打切残数":              If Not IsNull(adoCol.Value) Then .打切残数             = adoCol.Value
                Case "打切日":                If Not IsNull(adoCol.Value) Then .打切日               = adoCol.Value
                Case "内示超過区分":          If Not IsNull(adoCol.Value) Then .内示超過区分         = adoCol.Value
                Case "納品書発行日":          If Not IsNull(adoCol.Value) Then .納品書発行日         = adoCol.Value
                Case "支給書発行日":          If Not IsNull(adoCol.Value) Then .支給書発行日         = adoCol.Value
                Case "受領書発行日":          If Not IsNull(adoCol.Value) Then .受領書発行日         = adoCol.Value
                Case "明細書発行日":          If Not IsNull(adoCol.Value) Then .明細書発行日         = adoCol.Value
                Case "明細書NO":              If Not IsNull(adoCol.Value) Then .明細書NO             = adoCol.Value
                Case "初回発注管理NO":        If Not IsNull(adoCol.Value) Then .初回発注管理NO       = adoCol.Value
                Case "作成種別":              If Not IsNull(adoCol.Value) Then .作成種別             = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "作成者":                If Not IsNull(adoCol.Value) Then .作成者               = adoCol.Value
'                Case "更新日":                If Not IsNull(adoCol.Value) Then .更新日               = adoCol.Value
'                Case "更新端末":              If Not IsNull(adoCol.Value) Then .更新端末             = adoCol.Value
                Case "更新者":                If Not IsNull(adoCol.Value) Then .更新者               = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "仕入先名":	If Not IsNull(adoCol.Value) Then .仕入先名	= adoCol.Value
		Case "発注区分名":	If Not IsNull(adoCol.Value) Then .発注区分名	= adoCol.Value
		Case "手配区分名":	If Not IsNull(adoCol.Value) Then .手配区分名	= adoCol.Value
		Case "仕入先CD":	If Not IsNull(adoCol.Value) Then .仕入先CD	= adoCol.Value
		Case "仕入先工場CD":	If Not IsNull(adoCol.Value) Then .仕入先工場CD	= adoCol.Value
		Case "仕入先工場名":	If Not IsNull(adoCol.Value) Then .仕入先工場名	= adoCol.Value
		Case "表示寸法":	If Not IsNull(adoCol.Value) Then .表示寸法	= adoCol.Value
		Case "材種コード":	If Not IsNull(adoCol.Value) Then .材種コード	= adoCol.Value
		Case "手配単位名":	If Not IsNull(adoCol.Value) Then .手配単位名	= adoCol.Value
		Case "材料名":	If Not IsNull(adoCol.Value) Then .材料名	= adoCol.Value
		Case "収容器":	If Not IsNull(adoCol.Value) Then .収容器	= adoCol.Value
		Case "納入先名":	If Not IsNull(adoCol.Value) Then .納入先名	= adoCol.Value
		Case "納入先工場名":	If Not IsNull(adoCol.Value) Then .納入先工場名	= adoCol.Value
		Case "材種名":	If Not IsNull(adoCol.Value) Then .材種名	= adoCol.Value
		Case "納入先CD":	If Not IsNull(adoCol.Value) Then .納入先CD	= adoCol.Value
		Case "納入先工場CD":	If Not IsNull(adoCol.Value) Then .納入先工場CD	= adoCol.Value
		Case "置場":	If Not IsNull(adoCol.Value) Then .置場	= adoCol.Value
		Case "実数区分":	If Not IsNull(adoCol.Value) Then .実数区分	= adoCol.Value
		Case "入荷日":	If Not IsNull(adoCol.Value) Then .入荷日	= adoCol.Value
		Case "入荷便":	If Not IsNull(adoCol.Value) Then .入荷便	= adoCol.Value
		Case "挽回日":	If Not IsNull(adoCol.Value) Then .挽回日	= adoCol.Value
		Case "挽回便":	If Not IsNull(adoCol.Value) Then .挽回便	= adoCol.Value
		Case "分割数":	If Not IsNull(adoCol.Value) Then .分割数	= adoCol.Value
		Case "入荷枚数":	If Not IsNull(adoCol.Value) Then .入荷枚数	= adoCol.Value
		Case "入荷端数":	If Not IsNull(adoCol.Value) Then .入荷端数	= adoCol.Value
		Case "入荷数":	If Not IsNull(adoCol.Value) Then .入荷数	= adoCol.Value
		Case "入荷備考":	If Not IsNull(adoCol.Value) Then .入荷備考	= adoCol.Value
		Case "自社工区名":	If Not IsNull(adoCol.Value) Then .自社工区名	= adoCol.Value
		Case "品番":	If Not IsNull(adoCol.Value) Then .品番	= adoCol.Value
		Case "単価":	If Not IsNull(adoCol.Value) Then .単価	= adoCol.Value
		Case "単重":	If Not IsNull(adoCol.Value) Then .単重	= adoCol.Value
		Case "金額":	If Not IsNull(adoCol.Value) Then .金額	= adoCol.Value
		Case "支給実績管理NO":	If Not IsNull(adoCol.Value) Then .支給実績管理NO	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZCTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZCTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 材料発注テーブル"
        RTxtSql = RTxtSql &" Where 発注管理NO = " & ZCT.発注管理NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZCTRead = True
           Call ItemsFromZCT(DBNo)
	   ZCT_RDSTS = True
        else
	   ZCT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZCTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZCTReadNext(Optional DBNo As Integer = 0)
    If Not ZCT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZCT_RDSTS = False
        Else
           Call ItemsFromZCT(DBNo)
           ZCT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZCTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZCT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZCT )
'*----------------------------------------------*'
Public Sub ItemsClearZCT()
    With ZCT
	.発注管理NO = 0
	.納品書管理NO = 0
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
	.納入日 = ""
	.納入便 = ""
	.収容数 = 0
	.枚数 = 0
	.端数 = 0
	.発注数 = 0
	.発注区分 = ""
	.手配区分 = ""
	.SK区分 = 0
	.契約先 = ""
	.納入先 = ""
	.納入先受入 = ""
	.サイクル = ""
	.担当者 = ""
	.発注日 = ""
	.発注便 = ""
	.端数区分 = 0
	.打切区分 = 0
	.打切種別 = 0
	.打切残数 = 0
	.打切日 = ""
	.内示超過区分 = 0
	.納品書発行日 = ""
	.支給書発行日 = ""
	.受領書発行日 = ""
	.明細書発行日 = ""
	.明細書NO = 0
	.初回発注管理NO = 0
	.作成種別 = ""
'	.作成日 = 0
'	.作成端末 = ""
	.作成者 = ""
'	.更新日 = 0
'	.更新端末 = ""
	.更新者 = ""
        .件数 = 0
'   項目追加(自動)
	.仕入先名 = ""
	.発注区分名 = ""
	.手配区分名 = ""
	.仕入先CD = ""
	.仕入先工場CD = ""
	.仕入先工場名 = ""
	.表示寸法 = ""
	.材種コード = ""
	.手配単位名 = ""
	.材料名 = ""
	.収容器 = ""
	.納入先名 = ""
	.納入先工場名 = ""
	.材種名 = ""
	.納入先CD = ""
	.納入先工場CD = ""
	.置場 = ""
	.実数区分 = 0
	.入荷日 = ""
	.入荷便 = ""
	.挽回日 = ""
	.挽回便 = ""
	.分割数 = 0
	.入荷枚数 = 0
	.入荷端数 = 0
	.入荷数 = 0
	.入荷備考 = ""
	.自社工区名 = ""
	.品番 = ""
	.単価 = ""
	.単重 = ""
	.金額 = 0
	.支給実績管理NO = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZCTInsert(Optional DBNo As Integer) As Boolean
    ZCTInsert = False
    With ZCT
        gSL_Select = "Insert Into 材料発注テーブル  ( "
        gSL_Select = gSL_Select & "発注管理NO,納品書管理NO,納入番号,"
        gSL_Select = gSL_Select & "納品書頁,納品書行,分納区分,"
        gSL_Select = gSL_Select & "仕入先,受入,材質,"
        gSL_Select = gSL_Select & "板厚,幅,長さ,"
        gSL_Select = gSL_Select & "材料管理番号,背番号,納入日,"
        gSL_Select = gSL_Select & "納入便,収容数,枚数,"
        gSL_Select = gSL_Select & "端数,発注数,発注区分,"
        gSL_Select = gSL_Select & "手配区分,SK区分,契約先,"
        gSL_Select = gSL_Select & "納入先,納入先受入,サイクル,"
        gSL_Select = gSL_Select & "担当者,発注日,発注便,"
        gSL_Select = gSL_Select & "端数区分,打切区分,打切種別,"
        gSL_Select = gSL_Select & "打切残数,打切日,内示超過区分,"
        gSL_Select = gSL_Select & "納品書発行日,支給書発行日,受領書発行日,"
        gSL_Select = gSL_Select & "明細書発行日,明細書NO,初回発注管理NO,"
        gSL_Select = gSL_Select & "作成種別,"
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .発注管理NO & ","		
        gSL_Select = gSL_Select & "" & .納品書管理NO & ","		
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
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"		
        gSL_Select = gSL_Select & "" & .収容数 & ","		
        gSL_Select = gSL_Select & "" & .枚数 & ","		
        gSL_Select = gSL_Select & "" & .端数 & ","		
        gSL_Select = gSL_Select & "" & .発注数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.発注区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配区分) & "',"		
        gSL_Select = gSL_Select & "" & .SK区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.サイクル) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.担当者) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注便) & "',"		
        gSL_Select = gSL_Select & "" & .端数区分 & ","		
        gSL_Select = gSL_Select & "" & .打切区分 & ","		
        gSL_Select = gSL_Select & "" & .打切種別 & ","		
        gSL_Select = gSL_Select & "" & .打切残数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.打切日) & "',"		
        gSL_Select = gSL_Select & "" & .内示超過区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.納品書発行日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給書発行日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受領書発行日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.明細書発行日) & "',"		
        gSL_Select = gSL_Select & "" & .明細書NO & ","		
        gSL_Select = gSL_Select & "" & .初回発注管理NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.作成種別) & "',"		
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
    ZCTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZCTUpdate(Optional DBNo As Integer) As Boolean
    ZCTUpdate = False	
    With ZCT
        gSL_Select = "Update 材料発注テーブル  Set "
        gSL_Select = gSL_Select & "発注管理NO = " & .発注管理NO & ","
        gSL_Select = gSL_Select & "納品書管理NO = " & .納品書管理NO & ","
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
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "枚数 = " & .枚数 & ","
        gSL_Select = gSL_Select & "端数 = " & .端数 & ","
        gSL_Select = gSL_Select & "発注数 = " & .発注数 & ","
        gSL_Select = gSL_Select & "発注区分 = '" & RTrim(.発注区分) & "',"
        gSL_Select = gSL_Select & "手配区分 = '" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "SK区分 = " & .SK区分 & ","
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "納入先受入 = '" & RTrim(.納入先受入) & "',"
        gSL_Select = gSL_Select & "サイクル = '" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "担当者 = '" & RTrim(.担当者) & "',"
        gSL_Select = gSL_Select & "発注日 = '" & RTrim(.発注日) & "',"
        gSL_Select = gSL_Select & "発注便 = '" & RTrim(.発注便) & "',"
        gSL_Select = gSL_Select & "端数区分 = " & .端数区分 & ","
        gSL_Select = gSL_Select & "打切区分 = " & .打切区分 & ","
        gSL_Select = gSL_Select & "打切種別 = " & .打切種別 & ","
        gSL_Select = gSL_Select & "打切残数 = " & .打切残数 & ","
        gSL_Select = gSL_Select & "打切日 = '" & RTrim(.打切日) & "',"
        gSL_Select = gSL_Select & "内示超過区分 = " & .内示超過区分 & ","
        gSL_Select = gSL_Select & "納品書発行日 = '" & RTrim(.納品書発行日) & "',"
        gSL_Select = gSL_Select & "支給書発行日 = '" & RTrim(.支給書発行日) & "',"
        gSL_Select = gSL_Select & "受領書発行日 = '" & RTrim(.受領書発行日) & "',"
        gSL_Select = gSL_Select & "明細書発行日 = '" & RTrim(.明細書発行日) & "',"
        gSL_Select = gSL_Select & "明細書NO = " & .明細書NO & ","
        gSL_Select = gSL_Select & "初回発注管理NO = " & .初回発注管理NO & ","
        gSL_Select = gSL_Select & "作成種別 = '" & RTrim(.作成種別) & "',"
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
    ZCTUpdate = True
End Function

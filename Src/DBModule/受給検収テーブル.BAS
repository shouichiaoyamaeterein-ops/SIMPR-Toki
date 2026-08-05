Attribute VB_Name = "受給検収テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   受給検収テーブル
'*
'*   モジュールID  : 受給検収テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2007/12/19  By Cis
'*   作成日  :  2007/12/25  By Cis
'*   作成日  :  2008/01/18  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_受給検収テーブル
        受給管理NO            As Long
        受給管理行NO          As Byte
        検収日                As String
        仕入先                As String
        受入                  As String
        契約先                As String
        代行契約先            As String
        支給元                As String
        納入日                As String
        納入便                As String
        納入番号              As String
        納品書頁              As Byte
        納品書行              As Byte
        品番                  As String
        背番号                As String
        赤黒区分              As Byte
        枚数                  As Long
        収容数                As Long
        検収数                As Long
        支給元体系            As String
        支給元1               As String
        支給元2               As String
        出荷場                As String
        契約先仕入先          As String
        契約先納入先          As String
        契約先受入            As String
        受注者                As String
        受注者工場            As String
        発注者                As String
        発注者工場            As String
        発行日                As String
        発注区分              As String
        カード                As String
        機番                  As String
        総頁                  As Byte
        伝票種類              As Byte
        作成種別              As String
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
        実績年月              As String
        実績処理日            As String
        実績結果              As String
        工区                  As String
'        作成日                As Date
'        作成端末              As String
        作成者                As String
'        更新日                As Date
'        更新端末              As String
        更新者                As String
        件数          As Long
'   項目追加(自動)
	表示品番       As String
	発注区分名     As String
	契約先名       As String
	仕入先名       As String
	赤黒           As String
    End Type

    Public JYK As T_受給検収テーブル

    Public JYK_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromJYK(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearJYK	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With JYK
            Select Case adoCol.Name
                Case "受給管理NO":            If Not IsNull(adoCol.Value) Then .受給管理NO           = adoCol.Value
                Case "受給管理行NO":          If Not IsNull(adoCol.Value) Then .受給管理行NO         = adoCol.Value
                Case "検収日":                If Not IsNull(adoCol.Value) Then .検収日               = adoCol.Value
                Case "仕入先":                If Not IsNull(adoCol.Value) Then .仕入先               = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入                 = adoCol.Value
                Case "契約先":                If Not IsNull(adoCol.Value) Then .契約先               = adoCol.Value
                Case "代行契約先":            If Not IsNull(adoCol.Value) Then .代行契約先           = adoCol.Value
                Case "支給元":                If Not IsNull(adoCol.Value) Then .支給元               = adoCol.Value
                Case "納入日":                If Not IsNull(adoCol.Value) Then .納入日               = adoCol.Value
                Case "納入便":                If Not IsNull(adoCol.Value) Then .納入便               = adoCol.Value
                Case "納入番号":              If Not IsNull(adoCol.Value) Then .納入番号             = adoCol.Value
                Case "納品書頁":              If Not IsNull(adoCol.Value) Then .納品書頁             = adoCol.Value
                Case "納品書行":              If Not IsNull(adoCol.Value) Then .納品書行             = adoCol.Value
                Case "品番":                  If Not IsNull(adoCol.Value) Then .品番                 = adoCol.Value
                Case "背番号":                If Not IsNull(adoCol.Value) Then .背番号               = adoCol.Value
                Case "赤黒区分":              If Not IsNull(adoCol.Value) Then .赤黒区分             = adoCol.Value
                Case "枚数":                  If Not IsNull(adoCol.Value) Then .枚数                 = adoCol.Value
                Case "収容数":                If Not IsNull(adoCol.Value) Then .収容数               = adoCol.Value
                Case "検収数":                If Not IsNull(adoCol.Value) Then .検収数               = adoCol.Value
                Case "支給元体系":            If Not IsNull(adoCol.Value) Then .支給元体系           = adoCol.Value
                Case "支給元1":               If Not IsNull(adoCol.Value) Then .支給元1              = adoCol.Value
                Case "支給元2":               If Not IsNull(adoCol.Value) Then .支給元2              = adoCol.Value
                Case "出荷場":                If Not IsNull(adoCol.Value) Then .出荷場               = adoCol.Value
                Case "契約先仕入先":          If Not IsNull(adoCol.Value) Then .契約先仕入先         = adoCol.Value
                Case "契約先納入先":          If Not IsNull(adoCol.Value) Then .契約先納入先         = adoCol.Value
                Case "契約先受入":            If Not IsNull(adoCol.Value) Then .契約先受入           = adoCol.Value
                Case "受注者":                If Not IsNull(adoCol.Value) Then .受注者               = adoCol.Value
                Case "受注者工場":            If Not IsNull(adoCol.Value) Then .受注者工場           = adoCol.Value
                Case "発注者":                If Not IsNull(adoCol.Value) Then .発注者               = adoCol.Value
                Case "発注者工場":            If Not IsNull(adoCol.Value) Then .発注者工場           = adoCol.Value
                Case "発行日":                If Not IsNull(adoCol.Value) Then .発行日               = adoCol.Value
                Case "発注区分":              If Not IsNull(adoCol.Value) Then .発注区分             = adoCol.Value
                Case "カード":                If Not IsNull(adoCol.Value) Then .カード               = adoCol.Value
                Case "機番":                  If Not IsNull(adoCol.Value) Then .機番                 = adoCol.Value
                Case "総頁":                  If Not IsNull(adoCol.Value) Then .総頁                 = adoCol.Value
                Case "伝票種類":              If Not IsNull(adoCol.Value) Then .伝票種類             = adoCol.Value
                Case "作成種別":              If Not IsNull(adoCol.Value) Then .作成種別             = adoCol.Value
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
                Case "実績年月":              If Not IsNull(adoCol.Value) Then .実績年月             = adoCol.Value
                Case "実績処理日":            If Not IsNull(adoCol.Value) Then .実績処理日           = adoCol.Value
                Case "実績結果":              If Not IsNull(adoCol.Value) Then .実績結果             = adoCol.Value
                Case "工区":                  If Not IsNull(adoCol.Value) Then .工区                 = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "作成者":                If Not IsNull(adoCol.Value) Then .作成者               = adoCol.Value
'                Case "更新日":                If Not IsNull(adoCol.Value) Then .更新日               = adoCol.Value
'                Case "更新端末":              If Not IsNull(adoCol.Value) Then .更新端末             = adoCol.Value
                Case "更新者":                If Not IsNull(adoCol.Value) Then .更新者               = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "発注区分名":	If Not IsNull(adoCol.Value) Then .発注区分名	= adoCol.Value
		Case "契約先名":	If Not IsNull(adoCol.Value) Then .契約先名	= adoCol.Value
		Case "仕入先名":	If Not IsNull(adoCol.Value) Then .仕入先名	= adoCol.Value
		Case "赤黒":	If Not IsNull(adoCol.Value) Then .赤黒	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function JYKRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    JYKRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 受給検収テーブル"
        RTxtSql = RTxtSql &" Where 受給管理NO = " & JYK.受給管理NO & ""
        RTxtSql = RTxtSql &" And   受給管理行NO = " & JYK.受給管理行NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           JYKRead = True
           Call ItemsFromJYK(DBNo)
	   JYK_RDSTS = True
        else
	   JYK_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call JYKClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub JYKReadNext(Optional DBNo As Integer = 0)
    If Not JYK_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           JYK_RDSTS = False
        Else
           Call ItemsFromJYK(DBNo)
           JYK_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub JYKClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        JYK_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( JYK )
'*----------------------------------------------*'
Public Sub ItemsClearJYK()
    With JYK
	.受給管理NO = 0
	.受給管理行NO = 0
	.検収日 = ""
	.仕入先 = ""
	.受入 = ""
	.契約先 = ""
	.代行契約先 = ""
	.支給元 = ""
	.納入日 = ""
	.納入便 = ""
	.納入番号 = ""
	.納品書頁 = 0
	.納品書行 = 0
	.品番 = ""
	.背番号 = ""
	.赤黒区分 = 0
	.枚数 = 0
	.収容数 = 0
	.検収数 = 0
	.支給元体系 = ""
	.支給元1 = ""
	.支給元2 = ""
	.出荷場 = ""
	.契約先仕入先 = ""
	.契約先納入先 = ""
	.契約先受入 = ""
	.受注者 = ""
	.受注者工場 = ""
	.発注者 = ""
	.発注者工場 = ""
	.発行日 = ""
	.発注区分 = ""
	.カード = ""
	.機番 = ""
	.総頁 = 0
	.伝票種類 = 0
	.作成種別 = ""
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
	.実績年月 = ""
	.実績処理日 = ""
	.実績結果 = ""
	.工区 = ""
'	.作成日 = 0
'	.作成端末 = ""
	.作成者 = ""
'	.更新日 = 0
'	.更新端末 = ""
	.更新者 = ""
        .件数 = 0
'   項目追加(自動)
	.表示品番 = ""
	.発注区分名 = ""
	.契約先名 = ""
	.仕入先名 = ""
	.赤黒 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function JYKInsert(Optional DBNo As Integer) As Boolean
    JYKInsert = False
    With JYK
        gSL_Select = "Insert Into 受給検収テーブル  ( "
        gSL_Select = gSL_Select & "受給管理NO,受給管理行NO,検収日,"
        gSL_Select = gSL_Select & "仕入先,受入,契約先,"
        gSL_Select = gSL_Select & "代行契約先,支給元,納入日,"
        gSL_Select = gSL_Select & "納入便,納入番号,納品書頁,"
        gSL_Select = gSL_Select & "納品書行,品番,背番号,"
        gSL_Select = gSL_Select & "赤黒区分,枚数,収容数,"
        gSL_Select = gSL_Select & "検収数,支給元体系,支給元1,"
        gSL_Select = gSL_Select & "支給元2,出荷場,契約先仕入先,"
        gSL_Select = gSL_Select & "契約先納入先,契約先受入,受注者,"
        gSL_Select = gSL_Select & "受注者工場,発注者,発注者工場,"
        gSL_Select = gSL_Select & "発行日,発注区分,カード,"
        gSL_Select = gSL_Select & "機番,総頁,伝票種類,"
        gSL_Select = gSL_Select & "作成種別,TNS作成区分,TNS作成管理NO,"
        gSL_Select = gSL_Select & "TNS作成日,TNS作成処理日,TNS連番,"
        gSL_Select = gSL_Select & "TNS作成区分2,TNS作成管理NO2,TNS作成日2,"
        gSL_Select = gSL_Select & "TNS作成処理日2,TNS連番2,実績年月,"
        gSL_Select = gSL_Select & "実績処理日,実績結果,工区,"
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .受給管理NO & ","		
        gSL_Select = gSL_Select & "" & .受給管理行NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.検収日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.代行契約先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給元) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"		
        gSL_Select = gSL_Select & "" & .納品書頁 & ","		
        gSL_Select = gSL_Select & "" & .納品書行 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "" & .赤黒区分 & ","		
        gSL_Select = gSL_Select & "" & .枚数 & ","		
        gSL_Select = gSL_Select & "" & .収容数 & ","		
        gSL_Select = gSL_Select & "" & .検収数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.支給元体系) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給元1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給元2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.出荷場) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先仕入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先納入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受注者) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受注者工場) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注者) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注者工場) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発行日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.カード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.機番) & "',"		
        gSL_Select = gSL_Select & "" & .総頁 & ","		
        gSL_Select = gSL_Select & "" & .伝票種類 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.作成種別) & "',"		
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
        gSL_Select = gSL_Select & "'" & RTrim(.実績年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.実績処理日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.実績結果) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.工区) & "',"		
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
    JYKInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function JYKUpdate(Optional DBNo As Integer) As Boolean
    JYKUpdate = False	
    With JYK
        gSL_Select = "Update 受給検収テーブル  Set "
        gSL_Select = gSL_Select & "受給管理NO = " & .受給管理NO & ","
        gSL_Select = gSL_Select & "受給管理行NO = " & .受給管理行NO & ","
        gSL_Select = gSL_Select & "検収日 = '" & RTrim(.検収日) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "代行契約先 = '" & RTrim(.代行契約先) & "',"
        gSL_Select = gSL_Select & "支給元 = '" & RTrim(.支給元) & "',"
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "納品書頁 = " & .納品書頁 & ","
        gSL_Select = gSL_Select & "納品書行 = " & .納品書行 & ","
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "赤黒区分 = " & .赤黒区分 & ","
        gSL_Select = gSL_Select & "枚数 = " & .枚数 & ","
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "検収数 = " & .検収数 & ","
        gSL_Select = gSL_Select & "支給元体系 = '" & RTrim(.支給元体系) & "',"
        gSL_Select = gSL_Select & "支給元1 = '" & RTrim(.支給元1) & "',"
        gSL_Select = gSL_Select & "支給元2 = '" & RTrim(.支給元2) & "',"
        gSL_Select = gSL_Select & "出荷場 = '" & RTrim(.出荷場) & "',"
        gSL_Select = gSL_Select & "契約先仕入先 = '" & RTrim(.契約先仕入先) & "',"
        gSL_Select = gSL_Select & "契約先納入先 = '" & RTrim(.契約先納入先) & "',"
        gSL_Select = gSL_Select & "契約先受入 = '" & RTrim(.契約先受入) & "',"
        gSL_Select = gSL_Select & "受注者 = '" & RTrim(.受注者) & "',"
        gSL_Select = gSL_Select & "受注者工場 = '" & RTrim(.受注者工場) & "',"
        gSL_Select = gSL_Select & "発注者 = '" & RTrim(.発注者) & "',"
        gSL_Select = gSL_Select & "発注者工場 = '" & RTrim(.発注者工場) & "',"
        gSL_Select = gSL_Select & "発行日 = '" & RTrim(.発行日) & "',"
        gSL_Select = gSL_Select & "発注区分 = '" & RTrim(.発注区分) & "',"
        gSL_Select = gSL_Select & "カード = '" & RTrim(.カード) & "',"
        gSL_Select = gSL_Select & "機番 = '" & RTrim(.機番) & "',"
        gSL_Select = gSL_Select & "総頁 = " & .総頁 & ","
        gSL_Select = gSL_Select & "伝票種類 = " & .伝票種類 & ","
        gSL_Select = gSL_Select & "作成種別 = '" & RTrim(.作成種別) & "',"
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
        gSL_Select = gSL_Select & "実績年月 = '" & RTrim(.実績年月) & "',"
        gSL_Select = gSL_Select & "実績処理日 = '" & RTrim(.実績処理日) & "',"
        gSL_Select = gSL_Select & "実績結果 = '" & RTrim(.実績結果) & "',"
        gSL_Select = gSL_Select & "工区 = '" & RTrim(.工区) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "'"
        gSL_Select = gSL_Select & " Where 受給管理NO = " & .受給管理NO & ""
        gSL_Select = gSL_Select & "   And 受給管理行NO = " & .受給管理行NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    JYKUpdate = True
End Function

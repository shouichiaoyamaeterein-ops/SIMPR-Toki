Attribute VB_Name = "前工程入荷ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   前工程入荷ワーク
'*
'*   モジュールID  : 前工程入荷ワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2007/07/26  By Cis
'*   作成日  :  2007/07/27  By Cis
'*   作成日  :  2007/12/19  By Cis
'*   作成日  :  2007/12/20  By Cis
'*   作成日  :  2007/12/21  By Cis
'*   作成日  :  2007/12/25  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_前工程入荷ワーク
        処理端末          As String
        納入番号          As String
        納品書頁          As Byte
        納品書行          As Byte
        納入日            As String
        納入便            As String
        受注者            As String
        受注者工場        As String
        発注区分          As String
        カード            As String
        機番              As String
        支給元体系        As String
        支給元1           As String
        支給元2           As String
        総頁              As Byte
        出荷場            As String
        発注者            As String
        発注者工場        As String
        受入              As String
        発行日            As String
        品番              As String
        背番号            As String
        収容数            As Long
        枚数              As Long
        納入数            As Long
        検収日            As String
        仕入先            As String
        仕入先受入        As String
        契約先            As String
        代行契約先        As String
        処理区分          As Byte
        ERR伝票区分       As Byte
        ERRCD             As Byte
'        作成日            As Date
'        作成端末          As String
        作成者            As String
        件数          As Long
'   項目追加(自動)
	表示品番       As String
	エラー内容     As String
    End Type

    Public MKW As T_前工程入荷ワーク

    Public MKW_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromMKW(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearMKW	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With MKW
            Select Case adoCol.Name
                Case "処理端末":          If Not IsNull(adoCol.Value) Then .処理端末         = adoCol.Value
                Case "納入番号":          If Not IsNull(adoCol.Value) Then .納入番号         = adoCol.Value
                Case "納品書頁":          If Not IsNull(adoCol.Value) Then .納品書頁         = adoCol.Value
                Case "納品書行":          If Not IsNull(adoCol.Value) Then .納品書行         = adoCol.Value
                Case "納入日":            If Not IsNull(adoCol.Value) Then .納入日           = adoCol.Value
                Case "納入便":            If Not IsNull(adoCol.Value) Then .納入便           = adoCol.Value
                Case "受注者":            If Not IsNull(adoCol.Value) Then .受注者           = adoCol.Value
                Case "受注者工場":        If Not IsNull(adoCol.Value) Then .受注者工場       = adoCol.Value
                Case "発注区分":          If Not IsNull(adoCol.Value) Then .発注区分         = adoCol.Value
                Case "カード":            If Not IsNull(adoCol.Value) Then .カード           = adoCol.Value
                Case "機番":              If Not IsNull(adoCol.Value) Then .機番             = adoCol.Value
                Case "支給元体系":        If Not IsNull(adoCol.Value) Then .支給元体系       = adoCol.Value
                Case "支給元1":           If Not IsNull(adoCol.Value) Then .支給元1          = adoCol.Value
                Case "支給元2":           If Not IsNull(adoCol.Value) Then .支給元2          = adoCol.Value
                Case "総頁":              If Not IsNull(adoCol.Value) Then .総頁             = adoCol.Value
                Case "出荷場":            If Not IsNull(adoCol.Value) Then .出荷場           = adoCol.Value
                Case "発注者":            If Not IsNull(adoCol.Value) Then .発注者           = adoCol.Value
                Case "発注者工場":        If Not IsNull(adoCol.Value) Then .発注者工場       = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入             = adoCol.Value
                Case "発行日":            If Not IsNull(adoCol.Value) Then .発行日           = adoCol.Value
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番             = adoCol.Value
                Case "背番号":            If Not IsNull(adoCol.Value) Then .背番号           = adoCol.Value
                Case "収容数":            If Not IsNull(adoCol.Value) Then .収容数           = adoCol.Value
                Case "枚数":              If Not IsNull(adoCol.Value) Then .枚数             = adoCol.Value
                Case "納入数":            If Not IsNull(adoCol.Value) Then .納入数           = adoCol.Value
                Case "検収日":            If Not IsNull(adoCol.Value) Then .検収日           = adoCol.Value
                Case "仕入先":            If Not IsNull(adoCol.Value) Then .仕入先           = adoCol.Value
                Case "仕入先受入":        If Not IsNull(adoCol.Value) Then .仕入先受入       = adoCol.Value
                Case "契約先":            If Not IsNull(adoCol.Value) Then .契約先           = adoCol.Value
                Case "代行契約先":        If Not IsNull(adoCol.Value) Then .代行契約先       = adoCol.Value
                Case "処理区分":          If Not IsNull(adoCol.Value) Then .処理区分         = adoCol.Value
                Case "ERR伝票区分":       If Not IsNull(adoCol.Value) Then .ERR伝票区分      = adoCol.Value
                Case "ERRCD":             If Not IsNull(adoCol.Value) Then .ERRCD            = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者           = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "エラー内容":	If Not IsNull(adoCol.Value) Then .エラー内容	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function MKWRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    MKWRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 前工程入荷ワーク"
        RTxtSql = RTxtSql &" Where 処理端末 = '" & RTrim(MKW.処理端末) & "'"
        RTxtSql = RTxtSql &" And   納入番号 = '" & RTrim(MKW.納入番号) & "'"
        RTxtSql = RTxtSql &" And   納品書頁 = " & MKW.納品書頁 & ""
        RTxtSql = RTxtSql &" And   納品書行 = " & MKW.納品書行 & ""
        RTxtSql = RTxtSql &" And   納入日 = '" & RTrim(MKW.納入日) & "'"
        RTxtSql = RTxtSql &" And   納入便 = '" & RTrim(MKW.納入便) & "'"
        RTxtSql = RTxtSql &" And   受注者 = '" & RTrim(MKW.受注者) & "'"
        RTxtSql = RTxtSql &" And   受注者工場 = '" & RTrim(MKW.受注者工場) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           MKWRead = True
           Call ItemsFromMKW(DBNo)
	   MKW_RDSTS = True
        else
	   MKW_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call MKWClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub MKWReadNext(Optional DBNo As Integer = 0)
    If Not MKW_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           MKW_RDSTS = False
        Else
           Call ItemsFromMKW(DBNo)
           MKW_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub MKWClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        MKW_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( MKW )
'*----------------------------------------------*'
Public Sub ItemsClearMKW()
    With MKW
	.処理端末 = ""
	.納入番号 = ""
	.納品書頁 = 0
	.納品書行 = 0
	.納入日 = ""
	.納入便 = ""
	.受注者 = ""
	.受注者工場 = ""
	.発注区分 = ""
	.カード = ""
	.機番 = ""
	.支給元体系 = ""
	.支給元1 = ""
	.支給元2 = ""
	.総頁 = 0
	.出荷場 = ""
	.発注者 = ""
	.発注者工場 = ""
	.受入 = ""
	.発行日 = ""
	.品番 = ""
	.背番号 = ""
	.収容数 = 0
	.枚数 = 0
	.納入数 = 0
	.検収日 = ""
	.仕入先 = ""
	.仕入先受入 = ""
	.契約先 = ""
	.代行契約先 = ""
	.処理区分 = 0
	.ERR伝票区分 = 0
	.ERRCD = 0
'	.作成日 = 0
'	.作成端末 = ""
	.作成者 = ""
        .件数 = 0
'   項目追加(自動)
	.表示品番 = ""
	.エラー内容 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function MKWInsert(Optional DBNo As Integer) As Boolean
    MKWInsert = False
    With MKW
        gSL_Select = "Insert Into 前工程入荷ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,納入番号,納品書頁,"
        gSL_Select = gSL_Select & "納品書行,納入日,納入便,"
        gSL_Select = gSL_Select & "受注者,受注者工場,発注区分,"
        gSL_Select = gSL_Select & "カード,機番,支給元体系,"
        gSL_Select = gSL_Select & "支給元1,支給元2,総頁,"
        gSL_Select = gSL_Select & "出荷場,発注者,発注者工場,"
        gSL_Select = gSL_Select & "受入,発行日,品番,"
        gSL_Select = gSL_Select & "背番号,収容数,枚数,"
        gSL_Select = gSL_Select & "納入数,検収日,仕入先,"
        gSL_Select = gSL_Select & "仕入先受入,契約先,代行契約先,"
        gSL_Select = gSL_Select & "処理区分,ERR伝票区分,ERRCD,"
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"		
        gSL_Select = gSL_Select & "" & .納品書頁 & ","		
        gSL_Select = gSL_Select & "" & .納品書行 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受注者) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受注者工場) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.カード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.機番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給元体系) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給元1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給元2) & "',"		
        gSL_Select = gSL_Select & "" & .総頁 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.出荷場) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注者) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注者工場) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発行日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "" & .収容数 & ","		
        gSL_Select = gSL_Select & "" & .枚数 & ","		
        gSL_Select = gSL_Select & "" & .納入数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.検収日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.代行契約先) & "',"		
        gSL_Select = gSL_Select & "" & .処理区分 & ","		
        gSL_Select = gSL_Select & "" & .ERR伝票区分 & ","		
        gSL_Select = gSL_Select & "" & .ERRCD & ","		
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
    MKWInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function MKWUpdate(Optional DBNo As Integer) As Boolean
    MKWUpdate = False	
    With MKW
        gSL_Select = "Update 前工程入荷ワーク  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "納品書頁 = " & .納品書頁 & ","
        gSL_Select = gSL_Select & "納品書行 = " & .納品書行 & ","
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "受注者 = '" & RTrim(.受注者) & "',"
        gSL_Select = gSL_Select & "受注者工場 = '" & RTrim(.受注者工場) & "',"
        gSL_Select = gSL_Select & "発注区分 = '" & RTrim(.発注区分) & "',"
        gSL_Select = gSL_Select & "カード = '" & RTrim(.カード) & "',"
        gSL_Select = gSL_Select & "機番 = '" & RTrim(.機番) & "',"
        gSL_Select = gSL_Select & "支給元体系 = '" & RTrim(.支給元体系) & "',"
        gSL_Select = gSL_Select & "支給元1 = '" & RTrim(.支給元1) & "',"
        gSL_Select = gSL_Select & "支給元2 = '" & RTrim(.支給元2) & "',"
        gSL_Select = gSL_Select & "総頁 = " & .総頁 & ","
        gSL_Select = gSL_Select & "出荷場 = '" & RTrim(.出荷場) & "',"
        gSL_Select = gSL_Select & "発注者 = '" & RTrim(.発注者) & "',"
        gSL_Select = gSL_Select & "発注者工場 = '" & RTrim(.発注者工場) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "発行日 = '" & RTrim(.発行日) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "枚数 = " & .枚数 & ","
        gSL_Select = gSL_Select & "納入数 = " & .納入数 & ","
        gSL_Select = gSL_Select & "検収日 = '" & RTrim(.検収日) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "仕入先受入 = '" & RTrim(.仕入先受入) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "代行契約先 = '" & RTrim(.代行契約先) & "',"
        gSL_Select = gSL_Select & "処理区分 = " & .処理区分 & ","
        gSL_Select = gSL_Select & "ERR伝票区分 = " & .ERR伝票区分 & ","
        gSL_Select = gSL_Select & "ERRCD = " & .ERRCD & ""
        gSL_Select = gSL_Select & " Where 処理端末 = '" & RTrim(.処理端末) & "'"
        gSL_Select = gSL_Select & "   And 納入番号 = '" & RTrim(.納入番号) & "'"
        gSL_Select = gSL_Select & "   And 納品書頁 = " & .納品書頁 & ""
        gSL_Select = gSL_Select & "   And 納品書行 = " & .納品書行 & ""
        gSL_Select = gSL_Select & "   And 納入日 = '" & RTrim(.納入日) & "'"
        gSL_Select = gSL_Select & "   And 納入便 = '" & RTrim(.納入便) & "'"
        gSL_Select = gSL_Select & "   And 受注者 = '" & RTrim(.受注者) & "'"
        gSL_Select = gSL_Select & "   And 受注者工場 = '" & RTrim(.受注者工場) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    MKWUpdate = True
End Function

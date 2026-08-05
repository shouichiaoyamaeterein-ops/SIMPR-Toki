Attribute VB_Name = "前工程発行JAMA入荷問合ビュー"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   前工程発行JAMA入荷問合ビュー
'*
'*   モジュールID  : 前工程発行JAMA入荷問合ビュー.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2007/08/06  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_前工程発行JAMA入荷問合ビュー
        納入日便              As String
        納入日                As String
        納入便                As String
        仕入先                As String
        品番                  As String
        背番号                As String
        発注区分              As String
        表示品番              As String
        表示背番号            As String
        品名                  As String
        支給元1               As String
        編集支給元            As String
        支給元体系            As String
        支給元2               As String
        編集仕入先            As String
        受入                  As String
        収容数                As Long
        枚数                  As Long
        納入数                As Long
        納入番号              As String
        表示頁                As String
        発注区分名            As String
        編集納入先            As String
        出荷場                As String
        カード                As String
        編集発行日            As String
        編集TNS送信日         As String
        TNS作成日             As String
        編集検収年月          As String
        入荷管理NO            As Long
        件数          As Long
    End Type

    Public MJV As T_前工程発行JAMA入荷問合ビュー

    Public MJV_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromMJV(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearMJV	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With MJV
            Select Case adoCol.Name
                Case "納入日便":              If Not IsNull(adoCol.Value) Then .納入日便             = adoCol.Value
                Case "納入日":                If Not IsNull(adoCol.Value) Then .納入日               = adoCol.Value
                Case "納入便":                If Not IsNull(adoCol.Value) Then .納入便               = adoCol.Value
                Case "仕入先":                If Not IsNull(adoCol.Value) Then .仕入先               = adoCol.Value
                Case "品番":                  If Not IsNull(adoCol.Value) Then .品番                 = adoCol.Value
                Case "背番号":                If Not IsNull(adoCol.Value) Then .背番号               = adoCol.Value
                Case "発注区分":              If Not IsNull(adoCol.Value) Then .発注区分             = adoCol.Value
                Case "表示品番":              If Not IsNull(adoCol.Value) Then .表示品番             = adoCol.Value
                Case "表示背番号":            If Not IsNull(adoCol.Value) Then .表示背番号           = adoCol.Value
                Case "品名":                  If Not IsNull(adoCol.Value) Then .品名                 = adoCol.Value
                Case "支給元1":               If Not IsNull(adoCol.Value) Then .支給元1              = adoCol.Value
                Case "編集支給元":            If Not IsNull(adoCol.Value) Then .編集支給元           = adoCol.Value
                Case "支給元体系":            If Not IsNull(adoCol.Value) Then .支給元体系           = adoCol.Value
                Case "支給元2":               If Not IsNull(adoCol.Value) Then .支給元2              = adoCol.Value
                Case "編集仕入先":            If Not IsNull(adoCol.Value) Then .編集仕入先           = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入                 = adoCol.Value
                Case "収容数":                If Not IsNull(adoCol.Value) Then .収容数               = adoCol.Value
                Case "枚数":                  If Not IsNull(adoCol.Value) Then .枚数                 = adoCol.Value
                Case "納入数":                If Not IsNull(adoCol.Value) Then .納入数               = adoCol.Value
                Case "納入番号":              If Not IsNull(adoCol.Value) Then .納入番号             = adoCol.Value
                Case "表示頁":                If Not IsNull(adoCol.Value) Then .表示頁               = adoCol.Value
                Case "発注区分名":            If Not IsNull(adoCol.Value) Then .発注区分名           = adoCol.Value
                Case "編集納入先":            If Not IsNull(adoCol.Value) Then .編集納入先           = adoCol.Value
                Case "出荷場":                If Not IsNull(adoCol.Value) Then .出荷場               = adoCol.Value
                Case "カード":                If Not IsNull(adoCol.Value) Then .カード               = adoCol.Value
                Case "編集発行日":            If Not IsNull(adoCol.Value) Then .編集発行日           = adoCol.Value
                Case "編集TNS送信日":         If Not IsNull(adoCol.Value) Then .編集TNS送信日        = adoCol.Value
                Case "TNS作成日":             If Not IsNull(adoCol.Value) Then .TNS作成日            = adoCol.Value
                Case "編集検収年月":          If Not IsNull(adoCol.Value) Then .編集検収年月         = adoCol.Value
                Case "入荷管理NO":            If Not IsNull(adoCol.Value) Then .入荷管理NO           = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function MJVRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    MJVRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 前工程発行JAMA入荷問合ビュー"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           MJVRead = True
           Call ItemsFromMJV(DBNo)
	   MJV_RDSTS = True
        else
	   MJV_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call MJVClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub MJVReadNext(Optional DBNo As Integer = 0)
    If Not MJV_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           MJV_RDSTS = False
        Else
           Call ItemsFromMJV(DBNo)
           MJV_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub MJVClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        MJV_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( MJV )
'*----------------------------------------------*'
Public Sub ItemsClearMJV()
    With MJV
	.納入日便 = ""
	.納入日 = ""
	.納入便 = ""
	.仕入先 = ""
	.品番 = ""
	.背番号 = ""
	.発注区分 = ""
	.表示品番 = ""
	.表示背番号 = ""
	.品名 = ""
	.支給元1 = ""
	.編集支給元 = ""
	.支給元体系 = ""
	.支給元2 = ""
	.編集仕入先 = ""
	.受入 = ""
	.収容数 = 0
	.枚数 = 0
	.納入数 = 0
	.納入番号 = ""
	.表示頁 = ""
	.発注区分名 = ""
	.編集納入先 = ""
	.出荷場 = ""
	.カード = ""
	.編集発行日 = ""
	.編集TNS送信日 = ""
	.TNS作成日 = ""
	.編集検収年月 = ""
	.入荷管理NO = 0
        .件数 = 0
    End With
End Sub

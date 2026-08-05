Attribute VB_Name = "発注入荷状況問合ビュー"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   発注入荷状況問合ビュー
'*
'*   モジュールID  : 発注入荷状況問合ビュー.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/03/01  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_発注入荷状況問合ビュー
        納入日便      As String
        表示品番      As String
        表示背番号    As String
        品名          As String
        編集仕入先    As String
        受入          As String
        収容数        As Long
        発注枚数      As Long
        発注数        As Long
        入荷枚数      As Long
        入荷数        As Long
        入荷日便      As String
        表示納入番号  As String
        分納数        As Long
        入荷予定日便  As String
        車型          As String
        手配区分名    As String
        発注区分名    As String
        編集納入先    As String
        納入先受入    As String
        SK区分名      As String
        置場          As String
        サイクル      As String
        内示超過      As Long
        編集契約先    As String
        納入日        As String
        納入便        As String
        仕入先        As String
        品番          As String
        背番号        As String
        手配区分      As String
        発注区分      As String
        分納区分      As Byte
        納入番号      As String
        納品書頁      As Byte
        納品書行      As Byte
        入荷日        As String
        入荷便        As String
        入荷予定日    As String
        入荷予定便    As String
        発注管理NO    As Long
        TNS作成日     As String
        元TNS作成日   As String
        金額          As Currency
        件数          As Long
    End Type

    Public HNV As T_発注入荷状況問合ビュー

    Public HNV_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHNV(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHNV	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With HNV
            Select Case adoCol.Name
                Case "納入日便":      If Not IsNull(adoCol.Value) Then .納入日便     = adoCol.Value
                Case "表示品番":      If Not IsNull(adoCol.Value) Then .表示品番     = adoCol.Value
                Case "表示背番号":    If Not IsNull(adoCol.Value) Then .表示背番号   = adoCol.Value
                Case "品名":          If Not IsNull(adoCol.Value) Then .品名         = adoCol.Value
                Case "編集仕入先":    If Not IsNull(adoCol.Value) Then .編集仕入先   = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "収容数":        If Not IsNull(adoCol.Value) Then .収容数       = adoCol.Value
                Case "発注枚数":      If Not IsNull(adoCol.Value) Then .発注枚数     = adoCol.Value
                Case "発注数":        If Not IsNull(adoCol.Value) Then .発注数       = adoCol.Value
                Case "入荷枚数":      If Not IsNull(adoCol.Value) Then .入荷枚数     = adoCol.Value
                Case "入荷数":        If Not IsNull(adoCol.Value) Then .入荷数       = adoCol.Value
                Case "入荷日便":      If Not IsNull(adoCol.Value) Then .入荷日便     = adoCol.Value
                Case "表示納入番号":  If Not IsNull(adoCol.Value) Then .表示納入番号 = adoCol.Value
                Case "分納数":        If Not IsNull(adoCol.Value) Then .分納数       = adoCol.Value
                Case "入荷予定日便":  If Not IsNull(adoCol.Value) Then .入荷予定日便 = adoCol.Value
                Case "車型":          If Not IsNull(adoCol.Value) Then .車型         = adoCol.Value
                Case "手配区分名":    If Not IsNull(adoCol.Value) Then .手配区分名   = adoCol.Value
                Case "発注区分名":    If Not IsNull(adoCol.Value) Then .発注区分名   = adoCol.Value
                Case "編集納入先":    If Not IsNull(adoCol.Value) Then .編集納入先   = adoCol.Value
                Case "納入先受入":    If Not IsNull(adoCol.Value) Then .納入先受入   = adoCol.Value
                Case "SK区分名":      If Not IsNull(adoCol.Value) Then .SK区分名     = adoCol.Value
                Case "置場":          If Not IsNull(adoCol.Value) Then .置場         = adoCol.Value
                Case "サイクル":      If Not IsNull(adoCol.Value) Then .サイクル     = adoCol.Value
                Case "内示超過":      If Not IsNull(adoCol.Value) Then .内示超過     = adoCol.Value
                Case "編集契約先":    If Not IsNull(adoCol.Value) Then .編集契約先   = adoCol.Value
                Case "納入日":        If Not IsNull(adoCol.Value) Then .納入日       = adoCol.Value
                Case "納入便":        If Not IsNull(adoCol.Value) Then .納入便       = adoCol.Value
                Case "仕入先":        If Not IsNull(adoCol.Value) Then .仕入先       = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "背番号":        If Not IsNull(adoCol.Value) Then .背番号       = adoCol.Value
                Case "手配区分":      If Not IsNull(adoCol.Value) Then .手配区分     = adoCol.Value
                Case "発注区分":      If Not IsNull(adoCol.Value) Then .発注区分     = adoCol.Value
                Case "分納区分":      If Not IsNull(adoCol.Value) Then .分納区分     = adoCol.Value
                Case "納入番号":      If Not IsNull(adoCol.Value) Then .納入番号     = adoCol.Value
                Case "納品書頁":      If Not IsNull(adoCol.Value) Then .納品書頁     = adoCol.Value
                Case "納品書行":      If Not IsNull(adoCol.Value) Then .納品書行     = adoCol.Value
                Case "入荷日":        If Not IsNull(adoCol.Value) Then .入荷日       = adoCol.Value
                Case "入荷便":        If Not IsNull(adoCol.Value) Then .入荷便       = adoCol.Value
                Case "入荷予定日":    If Not IsNull(adoCol.Value) Then .入荷予定日   = adoCol.Value
                Case "入荷予定便":    If Not IsNull(adoCol.Value) Then .入荷予定便   = adoCol.Value
                Case "発注管理NO":    If Not IsNull(adoCol.Value) Then .発注管理NO   = adoCol.Value
                Case "TNS作成日":     If Not IsNull(adoCol.Value) Then .TNS作成日    = adoCol.Value
                Case "元TNS作成日":   If Not IsNull(adoCol.Value) Then .元TNS作成日  = adoCol.Value
                Case "金額":          If Not IsNull(adoCol.Value) Then .金額         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function HNVRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HNVRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 発注入荷状況問合ビュー"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HNVRead = True
           Call ItemsFromHNV(DBNo)
	   HNV_RDSTS = True
        else
	   HNV_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call HNVClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HNVReadNext(Optional DBNo As Integer = 0)
    If Not HNV_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HNV_RDSTS = False
        Else
           Call ItemsFromHNV(DBNo)
           HNV_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HNVClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HNV_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HNV )
'*----------------------------------------------*'
Public Sub ItemsClearHNV()
    With HNV
	.納入日便 = ""
	.表示品番 = ""
	.表示背番号 = ""
	.品名 = ""
	.編集仕入先 = ""
	.受入 = ""
	.収容数 = 0
	.発注枚数 = 0
	.発注数 = 0
	.入荷枚数 = 0
	.入荷数 = 0
	.入荷日便 = ""
	.表示納入番号 = ""
	.分納数 = 0
	.入荷予定日便 = ""
	.車型 = ""
	.手配区分名 = ""
	.発注区分名 = ""
	.編集納入先 = ""
	.納入先受入 = ""
	.SK区分名 = ""
	.置場 = ""
	.サイクル = ""
	.内示超過 = 0
	.編集契約先 = ""
	.納入日 = ""
	.納入便 = ""
	.仕入先 = ""
	.品番 = ""
	.背番号 = ""
	.手配区分 = ""
	.発注区分 = ""
	.分納区分 = 0
	.納入番号 = ""
	.納品書頁 = 0
	.納品書行 = 0
	.入荷日 = ""
	.入荷便 = ""
	.入荷予定日 = ""
	.入荷予定便 = ""
	.発注管理NO = 0
	.TNS作成日 = ""
	.元TNS作成日 = ""
	.金額 = 0
        .件数 = 0
    End With
End Sub

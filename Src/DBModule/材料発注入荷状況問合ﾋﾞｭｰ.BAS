Attribute VB_Name = "材料発注入荷状況問合ﾋﾞｭｰ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料発注入荷状況問合ﾋﾞｭｰ
'*
'*   モジュールID  : 材料発注入荷状況問合ﾋﾞｭｰ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/03/01  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料発注入荷状況問合ﾋﾞｭｰ
        納入日便      As String
        表示寸法      As String
        表示背番号    As String
        材料名        As String
        材種コード    As String
        編集仕入先    As String
        受入          As String
        収容数        As Currency
        発注枚数      As Long
        発注数        As Currency
        入荷枚数      As Long
        入荷数        As Currency
        入荷日便      As String
        表示納入番号  As String
        分納数        As Currency
        入荷予定日便  As String
        入荷備考      As String
        手配区分名    As String
        発注区分名    As String
        材種名        As String
        形状名        As String
        手配単位名    As String
        編集納入先    As String
        編集納受      As String
        編集メーカ    As String
        編集メ受      As String
        SK区分名      As String
        置場          As String
        サイクル      As String
        編集契約先    As String
        納入日        As String
        納入便        As String
        仕入先        As String
        材質          As String
        板厚          As Currency
        幅            As Currency
        長さ          As Currency
        寸法区分      As Byte
        材料管理番号  As Long
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
        件数          As Long
    End Type

    Public ZHV As T_材料発注入荷状況問合ﾋﾞｭｰ

    Public ZHV_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZHV(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZHV	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZHV
            Select Case adoCol.Name
                Case "納入日便":      If Not IsNull(adoCol.Value) Then .納入日便     = adoCol.Value
                Case "表示寸法":      If Not IsNull(adoCol.Value) Then .表示寸法     = adoCol.Value
                Case "表示背番号":    If Not IsNull(adoCol.Value) Then .表示背番号   = adoCol.Value
                Case "材料名":        If Not IsNull(adoCol.Value) Then .材料名       = adoCol.Value
                Case "材種コード":    If Not IsNull(adoCol.Value) Then .材種コード   = adoCol.Value
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
                Case "入荷備考":      If Not IsNull(adoCol.Value) Then .入荷備考     = adoCol.Value
                Case "手配区分名":    If Not IsNull(adoCol.Value) Then .手配区分名   = adoCol.Value
                Case "発注区分名":    If Not IsNull(adoCol.Value) Then .発注区分名   = adoCol.Value
                Case "材種名":        If Not IsNull(adoCol.Value) Then .材種名       = adoCol.Value
                Case "形状名":        If Not IsNull(adoCol.Value) Then .形状名       = adoCol.Value
                Case "手配単位名":    If Not IsNull(adoCol.Value) Then .手配単位名   = adoCol.Value
                Case "編集納入先":    If Not IsNull(adoCol.Value) Then .編集納入先   = adoCol.Value
                Case "編集納受":      If Not IsNull(adoCol.Value) Then .編集納受     = adoCol.Value
                Case "編集メーカ":    If Not IsNull(adoCol.Value) Then .編集メーカ   = adoCol.Value
                Case "編集メ受":      If Not IsNull(adoCol.Value) Then .編集メ受     = adoCol.Value
                Case "SK区分名":      If Not IsNull(adoCol.Value) Then .SK区分名     = adoCol.Value
                Case "置場":          If Not IsNull(adoCol.Value) Then .置場         = adoCol.Value
                Case "サイクル":      If Not IsNull(adoCol.Value) Then .サイクル     = adoCol.Value
                Case "編集契約先":    If Not IsNull(adoCol.Value) Then .編集契約先   = adoCol.Value
                Case "納入日":        If Not IsNull(adoCol.Value) Then .納入日       = adoCol.Value
                Case "納入便":        If Not IsNull(adoCol.Value) Then .納入便       = adoCol.Value
                Case "仕入先":        If Not IsNull(adoCol.Value) Then .仕入先       = adoCol.Value
                Case "材質":          If Not IsNull(adoCol.Value) Then .材質         = adoCol.Value
                Case "板厚":          If Not IsNull(adoCol.Value) Then .板厚         = adoCol.Value
                Case "幅":            If Not IsNull(adoCol.Value) Then .幅           = adoCol.Value
                Case "長さ":          If Not IsNull(adoCol.Value) Then .長さ         = adoCol.Value
                Case "寸法区分":      If Not IsNull(adoCol.Value) Then .寸法区分     = adoCol.Value
                Case "材料管理番号":  If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
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
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZHVRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZHVRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 材料発注入荷状況問合ﾋﾞｭｰ"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZHVRead = True
           Call ItemsFromZHV(DBNo)
	   ZHV_RDSTS = True
        else
	   ZHV_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZHVClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZHVReadNext(Optional DBNo As Integer = 0)
    If Not ZHV_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZHV_RDSTS = False
        Else
           Call ItemsFromZHV(DBNo)
           ZHV_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZHVClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZHV_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZHV )
'*----------------------------------------------*'
Public Sub ItemsClearZHV()
    With ZHV
	.納入日便 = ""
	.表示寸法 = ""
	.表示背番号 = ""
	.材料名 = ""
	.材種コード = ""
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
	.入荷備考 = ""
	.手配区分名 = ""
	.発注区分名 = ""
	.材種名 = ""
	.形状名 = ""
	.手配単位名 = ""
	.編集納入先 = ""
	.編集納受 = ""
	.編集メーカ = ""
	.編集メ受 = ""
	.SK区分名 = ""
	.置場 = ""
	.サイクル = ""
	.編集契約先 = ""
	.納入日 = ""
	.納入便 = ""
	.仕入先 = ""
	.材質 = ""
	.板厚 = 0
	.幅 = 0
	.長さ = 0
	.寸法区分 = 0
	.材料管理番号 = 0
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
        .件数 = 0
    End With
End Sub

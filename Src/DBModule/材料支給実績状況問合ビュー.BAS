Attribute VB_Name = "材料支給実績状況問合ビュー"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料支給実績状況問合ビュー
'*
'*   モジュールID  : 材料支給実績状況問合ビュー.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2008/07/11  By Cis(FMV-E620DE)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料支給実績状況問合ビュー
        実績日                As String
        実績便                As String
        実績日便              As String
        納品書頁              As Byte
        納品書行              As Byte
        納入番号              As String
        表示納入番号          As String
        材料管理番号          As Long
        材質                  As String
        背番号                As String
        表示背番号            As String
        収容数                As Currency
        実績枚数              As Currency
        実績数                As Currency
        単位名                As String
        表示寸法              As String
        実績先                As String
        略称                  As String
        編集支給先            As String
        受入                  As String
        材種コード            As String
        データ区分名          As String
        材種名                As String
        形状名                As String
        置場                  As String
        サイクル              As String
        有無償区分            As Byte
        表示有無償区分        As String
        実績換算値            As Currency
        材料名                As String
        単価                  As Currency
        金額                  As Long
        件数          As Long
    End Type

    Public ZSV As T_材料支給実績状況問合ビュー

    Public ZSV_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZSV(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZSV	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZSV
            Select Case adoCol.Name
                Case "実績日":                If Not IsNull(adoCol.Value) Then .実績日               = adoCol.Value
                Case "実績便":                If Not IsNull(adoCol.Value) Then .実績便               = adoCol.Value
                Case "実績日便":              If Not IsNull(adoCol.Value) Then .実績日便             = adoCol.Value
                Case "納品書頁":              If Not IsNull(adoCol.Value) Then .納品書頁             = adoCol.Value
                Case "納品書行":              If Not IsNull(adoCol.Value) Then .納品書行             = adoCol.Value
                Case "納入番号":              If Not IsNull(adoCol.Value) Then .納入番号             = adoCol.Value
                Case "表示納入番号":          If Not IsNull(adoCol.Value) Then .表示納入番号         = adoCol.Value
                Case "材料管理番号":          If Not IsNull(adoCol.Value) Then .材料管理番号         = adoCol.Value
                Case "材質":                  If Not IsNull(adoCol.Value) Then .材質                 = adoCol.Value
                Case "背番号":                If Not IsNull(adoCol.Value) Then .背番号               = adoCol.Value
                Case "表示背番号":            If Not IsNull(adoCol.Value) Then .表示背番号           = adoCol.Value
                Case "収容数":                If Not IsNull(adoCol.Value) Then .収容数               = adoCol.Value
                Case "実績枚数":              If Not IsNull(adoCol.Value) Then .実績枚数             = adoCol.Value
                Case "実績数":                If Not IsNull(adoCol.Value) Then .実績数               = adoCol.Value
                Case "単位名":                If Not IsNull(adoCol.Value) Then .単位名               = adoCol.Value
                Case "表示寸法":              If Not IsNull(adoCol.Value) Then .表示寸法             = adoCol.Value
                Case "実績先":                If Not IsNull(adoCol.Value) Then .実績先               = adoCol.Value
                Case "略称":                  If Not IsNull(adoCol.Value) Then .略称                 = adoCol.Value
                Case "編集支給先":            If Not IsNull(adoCol.Value) Then .編集支給先           = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入                 = adoCol.Value
                Case "材種コード":            If Not IsNull(adoCol.Value) Then .材種コード           = adoCol.Value
                Case "データ区分名":          If Not IsNull(adoCol.Value) Then .データ区分名         = adoCol.Value
                Case "材種名":                If Not IsNull(adoCol.Value) Then .材種名               = adoCol.Value
                Case "形状名":                If Not IsNull(adoCol.Value) Then .形状名               = adoCol.Value
                Case "置場":                  If Not IsNull(adoCol.Value) Then .置場                 = adoCol.Value
                Case "サイクル":              If Not IsNull(adoCol.Value) Then .サイクル             = adoCol.Value
                Case "有無償区分":            If Not IsNull(adoCol.Value) Then .有無償区分           = adoCol.Value
                Case "表示有無償区分":        If Not IsNull(adoCol.Value) Then .表示有無償区分       = adoCol.Value
                Case "実績換算値":            If Not IsNull(adoCol.Value) Then .実績換算値           = adoCol.Value
                Case "材料名":                If Not IsNull(adoCol.Value) Then .材料名               = adoCol.Value
                Case "単価":                  If Not IsNull(adoCol.Value) Then .単価                 = adoCol.Value
                Case "金額":                  If Not IsNull(adoCol.Value) Then .金額                 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZSVRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZSVRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 材料支給実績状況問合ビュー"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZSVRead = True
           Call ItemsFromZSV(DBNo)
	   ZSV_RDSTS = True
        else
	   ZSV_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZSVClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZSVReadNext(Optional DBNo As Integer = 0)
    If Not ZSV_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZSV_RDSTS = False
        Else
           Call ItemsFromZSV(DBNo)
           ZSV_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZSVClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZSV_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZSV )
'*----------------------------------------------*'
Public Sub ItemsClearZSV()
    With ZSV
	.実績日 = ""
	.実績便 = ""
	.実績日便 = ""
	.納品書頁 = 0
	.納品書行 = 0
	.納入番号 = ""
	.表示納入番号 = ""
	.材料管理番号 = 0
	.材質 = ""
	.背番号 = ""
	.表示背番号 = ""
	.収容数 = 0
	.実績枚数 = 0
	.実績数 = 0
	.単位名 = ""
	.表示寸法 = ""
	.実績先 = ""
	.略称 = ""
	.編集支給先 = ""
	.受入 = ""
	.材種コード = ""
	.データ区分名 = ""
	.材種名 = ""
	.形状名 = ""
	.置場 = ""
	.サイクル = ""
	.有無償区分 = 0
	.表示有無償区分 = ""
	.実績換算値 = 0
	.材料名 = ""
	.単価 = 0
	.金額 = 0
        .件数 = 0
    End With
End Sub

Attribute VB_Name = "TNS支給入荷ﾋﾞｭｰ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TNS支給入荷ﾋﾞｭｰ
'*
'*   モジュールID  : TNS支給入荷ﾋﾞｭｰ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2007/03/09  By Cis
'*   作成日  :  2007/03/12  By Cis
'*   作成日  :  2007/08/03  By Cis
'*   作成日  :  2008/01/17  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_TNS支給入荷ﾋﾞｭｰ
        データ種別            As Long
        発注管理NO            As Long
        納入番号              As String
        入荷日                As String
        納入日                As String
        納入便                As String
        品番                  As String
        入荷数                As Long
        契約先                As String
        代行契約先            As String
        自社仕入先            As String
        自社受入              As String
        TNS作成区分           As Byte
        TNS作成管理NO         As Long
        TNS作成区分2          As Byte
        TNS作成管理NO2        As Long
        支給元2               As String
        支給先CD              As String
        支給先工場CD          As String
        支給先受入            As String
        支給元                As String
        仕入先                As String
        件数          As Long
'   項目追加(自動)
	カード         As String
	支給先         As String
	支給先工場     As String
    End Type

    Public TSV As T_TNS支給入荷ﾋﾞｭｰ

    Public TSV_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTSV(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTSV	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With TSV
            Select Case adoCol.Name
                Case "データ種別":            If Not IsNull(adoCol.Value) Then .データ種別           = adoCol.Value
                Case "発注管理NO":            If Not IsNull(adoCol.Value) Then .発注管理NO           = adoCol.Value
                Case "納入番号":              If Not IsNull(adoCol.Value) Then .納入番号             = adoCol.Value
                Case "入荷日":                If Not IsNull(adoCol.Value) Then .入荷日               = adoCol.Value
                Case "納入日":                If Not IsNull(adoCol.Value) Then .納入日               = adoCol.Value
                Case "納入便":                If Not IsNull(adoCol.Value) Then .納入便               = adoCol.Value
                Case "品番":                  If Not IsNull(adoCol.Value) Then .品番                 = adoCol.Value
                Case "入荷数":                If Not IsNull(adoCol.Value) Then .入荷数               = adoCol.Value
                Case "契約先":                If Not IsNull(adoCol.Value) Then .契約先               = adoCol.Value
                Case "代行契約先":            If Not IsNull(adoCol.Value) Then .代行契約先           = adoCol.Value
                Case "自社仕入先":            If Not IsNull(adoCol.Value) Then .自社仕入先           = adoCol.Value
                Case "自社受入":              If Not IsNull(adoCol.Value) Then .自社受入             = adoCol.Value
                Case "TNS作成区分":           If Not IsNull(adoCol.Value) Then .TNS作成区分          = adoCol.Value
                Case "TNS作成管理NO":         If Not IsNull(adoCol.Value) Then .TNS作成管理NO        = adoCol.Value
                Case "TNS作成区分2":          If Not IsNull(adoCol.Value) Then .TNS作成区分2         = adoCol.Value
                Case "TNS作成管理NO2":        If Not IsNull(adoCol.Value) Then .TNS作成管理NO2       = adoCol.Value
                Case "支給元2":               If Not IsNull(adoCol.Value) Then .支給元2              = adoCol.Value
                Case "支給先CD":              If Not IsNull(adoCol.Value) Then .支給先CD             = adoCol.Value
                Case "支給先工場CD":          If Not IsNull(adoCol.Value) Then .支給先工場CD         = adoCol.Value
                Case "支給先受入":            If Not IsNull(adoCol.Value) Then .支給先受入           = adoCol.Value
                Case "支給元":                If Not IsNull(adoCol.Value) Then .支給元               = adoCol.Value
                Case "仕入先":                If Not IsNull(adoCol.Value) Then .仕入先               = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "カード":	If Not IsNull(adoCol.Value) Then .カード	= adoCol.Value
		Case "支給先":	If Not IsNull(adoCol.Value) Then .支給先	= adoCol.Value
		Case "支給先工場":	If Not IsNull(adoCol.Value) Then .支給先工場	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TSVRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TSVRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From TNS支給入荷ﾋﾞｭｰ"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TSVRead = True
           Call ItemsFromTSV(DBNo)
	   TSV_RDSTS = True
        else
	   TSV_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call TSVClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TSVReadNext(Optional DBNo As Integer = 0)
    If Not TSV_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TSV_RDSTS = False
        Else
           Call ItemsFromTSV(DBNo)
           TSV_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TSVClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TSV_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TSV )
'*----------------------------------------------*'
Public Sub ItemsClearTSV()
    With TSV
	.データ種別 = 0
	.発注管理NO = 0
	.納入番号 = ""
	.入荷日 = ""
	.納入日 = ""
	.納入便 = ""
	.品番 = ""
	.入荷数 = 0
	.契約先 = ""
	.代行契約先 = ""
	.自社仕入先 = ""
	.自社受入 = ""
	.TNS作成区分 = 0
	.TNS作成管理NO = 0
	.TNS作成区分2 = 0
	.TNS作成管理NO2 = 0
	.支給元2 = ""
	.支給先CD = ""
	.支給先工場CD = ""
	.支給先受入 = ""
	.支給元 = ""
	.仕入先 = ""
        .件数 = 0
'   項目追加(自動)
	.カード = ""
	.支給先 = ""
	.支給先工場 = ""
    End With
End Sub

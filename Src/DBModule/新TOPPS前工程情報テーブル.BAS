Attribute VB_Name = "新TOPPS前工程情報テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   新TOPPS前工程情報テーブル
'*
'*   モジュールID  : 新TOPPS前工程情報テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2008/07/15  By Cis(FMV-E630)
'*   作成日  :  2008/07/17  By Cis(FMV-E630)
'*   作成日  :  2008/07/22  By Cis(FMV-E630)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_新TOPPS前工程情報テーブル
        納入先            As String
        納入先コード      As String
        納入先工区        As String
        納入先事業体      As String
        納入先事業体工区  As String
        受入              As String
        自工程            As String
        契約先            As String
        仕入先事業体      As String
        仕入先            As String
        仕入先コード      As String
        仕入先工区        As String
        出荷場            As String
        前工程            As String
'        作成日            As Date
'        作成端末          As String
'        更新日            As Date
'        更新端末          As String
        件数          As Long
'   項目追加(自動)
	納入先略称     As String
	契約先略称     As String
	被自社         As String
	被自社工区     As String
	対象件数       As Integer
	品番           As String
	表示品番       As String
	収容数1        As Integer
	収容器1        As String
	変更日         As String
	変更便         As String
	収容数2        As Integer
	収容器2        As String
	置場           As String
	所番地変更日   As String
	所番地変更便   As String
	変更後所番地   As String
	工場集荷名称   As String
	ポイント       As String
	工場集荷順路   As Integer
	工場集荷コード As String
	背番号         As String
    End Type

    Public NTM As T_新TOPPS前工程情報テーブル

    Public NTM_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromNTM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearNTM	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With NTM
            Select Case adoCol.Name
                Case "納入先":            If Not IsNull(adoCol.Value) Then .納入先           = adoCol.Value
                Case "納入先コード":      If Not IsNull(adoCol.Value) Then .納入先コード     = adoCol.Value
                Case "納入先工区":        If Not IsNull(adoCol.Value) Then .納入先工区       = adoCol.Value
                Case "納入先事業体":      If Not IsNull(adoCol.Value) Then .納入先事業体     = adoCol.Value
                Case "納入先事業体工区":  If Not IsNull(adoCol.Value) Then .納入先事業体工区 = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入             = adoCol.Value
                Case "自工程":            If Not IsNull(adoCol.Value) Then .自工程           = adoCol.Value
                Case "契約先":            If Not IsNull(adoCol.Value) Then .契約先           = adoCol.Value
                Case "仕入先事業体":      If Not IsNull(adoCol.Value) Then .仕入先事業体     = adoCol.Value
                Case "仕入先":            If Not IsNull(adoCol.Value) Then .仕入先           = adoCol.Value
                Case "仕入先コード":      If Not IsNull(adoCol.Value) Then .仕入先コード     = adoCol.Value
                Case "仕入先工区":        If Not IsNull(adoCol.Value) Then .仕入先工区       = adoCol.Value
                Case "出荷場":            If Not IsNull(adoCol.Value) Then .出荷場           = adoCol.Value
                Case "前工程":            If Not IsNull(adoCol.Value) Then .前工程           = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "納入先略称":	If Not IsNull(adoCol.Value) Then .納入先略称	= adoCol.Value
		Case "契約先略称":	If Not IsNull(adoCol.Value) Then .契約先略称	= adoCol.Value
		Case "被自社":	If Not IsNull(adoCol.Value) Then .被自社	= adoCol.Value
		Case "被自社工区":	If Not IsNull(adoCol.Value) Then .被自社工区	= adoCol.Value
		Case "対象件数":	If Not IsNull(adoCol.Value) Then .対象件数	= adoCol.Value
		Case "品番":	If Not IsNull(adoCol.Value) Then .品番	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "収容数1":	If Not IsNull(adoCol.Value) Then .収容数1	= adoCol.Value
		Case "収容器1":	If Not IsNull(adoCol.Value) Then .収容器1	= adoCol.Value
		Case "変更日":	If Not IsNull(adoCol.Value) Then .変更日	= adoCol.Value
		Case "変更便":	If Not IsNull(adoCol.Value) Then .変更便	= adoCol.Value
		Case "収容数2":	If Not IsNull(adoCol.Value) Then .収容数2	= adoCol.Value
		Case "収容器2":	If Not IsNull(adoCol.Value) Then .収容器2	= adoCol.Value
		Case "置場":	If Not IsNull(adoCol.Value) Then .置場	= adoCol.Value
		Case "所番地変更日":	If Not IsNull(adoCol.Value) Then .所番地変更日	= adoCol.Value
		Case "所番地変更便":	If Not IsNull(adoCol.Value) Then .所番地変更便	= adoCol.Value
		Case "変更後所番地":	If Not IsNull(adoCol.Value) Then .変更後所番地	= adoCol.Value
		Case "工場集荷名称":	If Not IsNull(adoCol.Value) Then .工場集荷名称	= adoCol.Value
		Case "ポイント":	If Not IsNull(adoCol.Value) Then .ポイント	= adoCol.Value
		Case "工場集荷順路":	If Not IsNull(adoCol.Value) Then .工場集荷順路	= adoCol.Value
		Case "工場集荷コード":	If Not IsNull(adoCol.Value) Then .工場集荷コード	= adoCol.Value
		Case "背番号":	If Not IsNull(adoCol.Value) Then .背番号	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function NTMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    NTMRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 新TOPPS前工程情報テーブル"
        RTxtSql = RTxtSql &" Where 納入先 = '" & RTrim(NTM.納入先) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           NTMRead = True
           Call ItemsFromNTM(DBNo)
	   NTM_RDSTS = True
        else
	   NTM_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call NTMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub NTMReadNext(Optional DBNo As Integer = 0)
    If Not NTM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           NTM_RDSTS = False
        Else
           Call ItemsFromNTM(DBNo)
           NTM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub NTMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        NTM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( NTM )
'*----------------------------------------------*'
Public Sub ItemsClearNTM()
    With NTM
	.納入先 = ""
	.納入先コード = ""
	.納入先工区 = ""
	.納入先事業体 = ""
	.納入先事業体工区 = ""
	.受入 = ""
	.自工程 = ""
	.契約先 = ""
	.仕入先事業体 = ""
	.仕入先 = ""
	.仕入先コード = ""
	.仕入先工区 = ""
	.出荷場 = ""
	.前工程 = ""
'	.作成日 = 0
'	.作成端末 = ""
'	.更新日 = 0
'	.更新端末 = ""
        .件数 = 0
'   項目追加(自動)
	.納入先略称 = ""
	.契約先略称 = ""
	.被自社 = ""
	.被自社工区 = ""
	.対象件数 = 0
	.品番 = ""
	.表示品番 = ""
	.収容数1 = 0
	.収容器1 = ""
	.変更日 = ""
	.変更便 = ""
	.収容数2 = 0
	.収容器2 = ""
	.置場 = ""
	.所番地変更日 = ""
	.所番地変更便 = ""
	.変更後所番地 = ""
	.工場集荷名称 = ""
	.ポイント = ""
	.工場集荷順路 = 0
	.工場集荷コード = ""
	.背番号 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function NTMInsert(Optional DBNo As Integer) As Boolean
    NTMInsert = False
    With NTM
        gSL_Select = "Insert Into 新TOPPS前工程情報テーブル  ( "
        gSL_Select = gSL_Select & "納入先,納入先コード,納入先工区,"
        gSL_Select = gSL_Select & "納入先事業体,納入先事業体工区,受入,"
        gSL_Select = gSL_Select & "自工程,契約先,仕入先事業体,"
        gSL_Select = gSL_Select & "仕入先,仕入先コード,仕入先工区,"
        gSL_Select = gSL_Select & "出荷場,前工程,"
        gSL_Select = gSL_Select & "作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先コード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先工区) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先事業体) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先事業体工区) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.自工程) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先事業体) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先コード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先工区) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.出荷場) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.前工程) & "',"		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    NTMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function NTMUpdate(Optional DBNo As Integer) As Boolean
    NTMUpdate = False	
    With NTM
        gSL_Select = "Update 新TOPPS前工程情報テーブル  Set "
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "納入先コード = '" & RTrim(.納入先コード) & "',"
        gSL_Select = gSL_Select & "納入先工区 = '" & RTrim(.納入先工区) & "',"
        gSL_Select = gSL_Select & "納入先事業体 = '" & RTrim(.納入先事業体) & "',"
        gSL_Select = gSL_Select & "納入先事業体工区 = '" & RTrim(.納入先事業体工区) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "自工程 = '" & RTrim(.自工程) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "仕入先事業体 = '" & RTrim(.仕入先事業体) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "仕入先コード = '" & RTrim(.仕入先コード) & "',"
        gSL_Select = gSL_Select & "仕入先工区 = '" & RTrim(.仕入先工区) & "',"
        gSL_Select = gSL_Select & "出荷場 = '" & RTrim(.出荷場) & "',"
        gSL_Select = gSL_Select & "前工程 = '" & RTrim(.前工程) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 納入先 = '" & RTrim(.納入先) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    NTMUpdate = True
End Function

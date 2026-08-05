Attribute VB_Name = "出荷画面入力ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   出荷画面入力ワーク
'*
'*   モジュールID  : 出荷画面入力ワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2005/06/03  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_出荷画面入力ワーク
        処理端末              As String
        処理区分              As Byte
        受注管理NO            As Long
        受注管理行NO          As Byte
        受注区分              As Byte
        受注日                As String
        受注便                As String
        納入日                As String
        納入便                As String
        納入先                As String
        出荷予定日            As String
        出荷予定便            As String
        受入                  As String
        機番                  As String
        納入番号              As String
        納品書頁              As Byte
        納品書行              As Byte
        分納区分              As Byte
        分納回数              As Byte
        品番                  As String
        背番号                As String
        枚数                  As Long
        収容数                As Long
        納入数                As Long
        端数                  As Long
        出荷指示書NO          As Long
        出荷指示書行NO        As Byte
        契約先                As String
        SK区分                As Byte
        作成種別              As String
        伝票種類              As String
        納品書管理NO          As Long
        作成者                As String
        件数          As Long
'   項目追加(自動)
	納入先CD      As String
	納入先工場CD  As String
	納入先名      As String
	納入先工場名  As String
	受注区分名    As String
	表示品番      As String
    End Type

    Public SGW As T_出荷画面入力ワーク

    Public SGW_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromSGW(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearSGW	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With SGW
            Select Case adoCol.Name
                Case "処理端末":              If Not IsNull(adoCol.Value) Then .処理端末             = adoCol.Value
                Case "処理区分":              If Not IsNull(adoCol.Value) Then .処理区分             = adoCol.Value
                Case "受注管理NO":            If Not IsNull(adoCol.Value) Then .受注管理NO           = adoCol.Value
                Case "受注管理行NO":          If Not IsNull(adoCol.Value) Then .受注管理行NO         = adoCol.Value
                Case "受注区分":              If Not IsNull(adoCol.Value) Then .受注区分             = adoCol.Value
                Case "受注日":                If Not IsNull(adoCol.Value) Then .受注日               = adoCol.Value
                Case "受注便":                If Not IsNull(adoCol.Value) Then .受注便               = adoCol.Value
                Case "納入日":                If Not IsNull(adoCol.Value) Then .納入日               = adoCol.Value
                Case "納入便":                If Not IsNull(adoCol.Value) Then .納入便               = adoCol.Value
                Case "納入先":                If Not IsNull(adoCol.Value) Then .納入先               = adoCol.Value
                Case "出荷予定日":            If Not IsNull(adoCol.Value) Then .出荷予定日           = adoCol.Value
                Case "出荷予定便":            If Not IsNull(adoCol.Value) Then .出荷予定便           = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入                 = adoCol.Value
                Case "機番":                  If Not IsNull(adoCol.Value) Then .機番                 = adoCol.Value
                Case "納入番号":              If Not IsNull(adoCol.Value) Then .納入番号             = adoCol.Value
                Case "納品書頁":              If Not IsNull(adoCol.Value) Then .納品書頁             = adoCol.Value
                Case "納品書行":              If Not IsNull(adoCol.Value) Then .納品書行             = adoCol.Value
                Case "分納区分":              If Not IsNull(adoCol.Value) Then .分納区分             = adoCol.Value
                Case "分納回数":              If Not IsNull(adoCol.Value) Then .分納回数             = adoCol.Value
                Case "品番":                  If Not IsNull(adoCol.Value) Then .品番                 = adoCol.Value
                Case "背番号":                If Not IsNull(adoCol.Value) Then .背番号               = adoCol.Value
                Case "枚数":                  If Not IsNull(adoCol.Value) Then .枚数                 = adoCol.Value
                Case "収容数":                If Not IsNull(adoCol.Value) Then .収容数               = adoCol.Value
                Case "納入数":                If Not IsNull(adoCol.Value) Then .納入数               = adoCol.Value
                Case "端数":                  If Not IsNull(adoCol.Value) Then .端数                 = adoCol.Value
                Case "出荷指示書NO":          If Not IsNull(adoCol.Value) Then .出荷指示書NO         = adoCol.Value
                Case "出荷指示書行NO":        If Not IsNull(adoCol.Value) Then .出荷指示書行NO       = adoCol.Value
                Case "契約先":                If Not IsNull(adoCol.Value) Then .契約先               = adoCol.Value
                Case "SK区分":                If Not IsNull(adoCol.Value) Then .SK区分               = adoCol.Value
                Case "作成種別":              If Not IsNull(adoCol.Value) Then .作成種別             = adoCol.Value
                Case "伝票種類":              If Not IsNull(adoCol.Value) Then .伝票種類             = adoCol.Value
                Case "納品書管理NO":          If Not IsNull(adoCol.Value) Then .納品書管理NO         = adoCol.Value
                Case "作成者":                If Not IsNull(adoCol.Value) Then .作成者               = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "納入先CD":	If Not IsNull(adoCol.Value) Then .納入先CD	= adoCol.Value
		Case "納入先工場CD":	If Not IsNull(adoCol.Value) Then .納入先工場CD	= adoCol.Value
		Case "納入先名":	If Not IsNull(adoCol.Value) Then .納入先名	= adoCol.Value
		Case "納入先工場名":	If Not IsNull(adoCol.Value) Then .納入先工場名	= adoCol.Value
		Case "受注区分名":	If Not IsNull(adoCol.Value) Then .受注区分名	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function SGWRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    SGWRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 出荷画面入力ワーク"
        RTxtSql = RTxtSql &" Where 処理端末 = '" & RTrim(SGW.処理端末) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           SGWRead = True
           Call ItemsFromSGW(DBNo)
	   SGW_RDSTS = True
        else
	   SGW_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call SGWClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub SGWReadNext(Optional DBNo As Integer = 0)
    If Not SGW_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           SGW_RDSTS = False
        Else
           Call ItemsFromSGW(DBNo)
           SGW_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub SGWClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        SGW_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( SGW )
'*----------------------------------------------*'
Public Sub ItemsClearSGW()
    With SGW
	.処理端末 = ""
	.処理区分 = 0
	.受注管理NO = 0
	.受注管理行NO = 0
	.受注区分 = 0
	.受注日 = ""
	.受注便 = ""
	.納入日 = ""
	.納入便 = ""
	.納入先 = ""
	.出荷予定日 = ""
	.出荷予定便 = ""
	.受入 = ""
	.機番 = ""
	.納入番号 = ""
	.納品書頁 = 0
	.納品書行 = 0
	.分納区分 = 0
	.分納回数 = 0
	.品番 = ""
	.背番号 = ""
	.枚数 = 0
	.収容数 = 0
	.納入数 = 0
	.端数 = 0
	.出荷指示書NO = 0
	.出荷指示書行NO = 0
	.契約先 = ""
	.SK区分 = 0
	.作成種別 = ""
	.伝票種類 = ""
	.納品書管理NO = 0
	.作成者 = ""
        .件数 = 0
'   項目追加(自動)
	.納入先CD = ""
	.納入先工場CD = ""
	.納入先名 = ""
	.納入先工場名 = ""
	.受注区分名 = ""
	.表示品番 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function SGWInsert(Optional DBNo As Integer) As Boolean
    SGWInsert = True
    With SGW
        gSL_Select = "Insert Into 出荷画面入力ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,処理区分,受注管理NO,"
        gSL_Select = gSL_Select & "受注管理行NO,受注区分,受注日,"
        gSL_Select = gSL_Select & "受注便,納入日,納入便,"
        gSL_Select = gSL_Select & "納入先,出荷予定日,出荷予定便,"
        gSL_Select = gSL_Select & "受入,機番,納入番号,"
        gSL_Select = gSL_Select & "納品書頁,納品書行,分納区分,"
        gSL_Select = gSL_Select & "分納回数,品番,背番号,"
        gSL_Select = gSL_Select & "枚数,収容数,納入数,"
        gSL_Select = gSL_Select & "端数,出荷指示書NO,出荷指示書行NO,"
        gSL_Select = gSL_Select & "契約先,SK区分,作成種別,"
        gSL_Select = gSL_Select & "伝票種類,納品書管理NO,"
        gSL_Select = gSL_Select & "作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"		
        gSL_Select = gSL_Select & "" & .処理区分 & ","		
        gSL_Select = gSL_Select & "" & .受注管理NO & ","		
        gSL_Select = gSL_Select & "" & .受注管理行NO & ","		
        gSL_Select = gSL_Select & "" & .受注区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.受注日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受注便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.出荷予定日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.出荷予定便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.機番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"		
        gSL_Select = gSL_Select & "" & .納品書頁 & ","		
        gSL_Select = gSL_Select & "" & .納品書行 & ","		
        gSL_Select = gSL_Select & "" & .分納区分 & ","		
        gSL_Select = gSL_Select & "" & .分納回数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "" & .枚数 & ","		
        gSL_Select = gSL_Select & "" & .収容数 & ","		
        gSL_Select = gSL_Select & "" & .納入数 & ","		
        gSL_Select = gSL_Select & "" & .端数 & ","		
        gSL_Select = gSL_Select & "" & .出荷指示書NO & ","		
        gSL_Select = gSL_Select & "" & .出荷指示書行NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"		
        gSL_Select = gSL_Select & "" & .SK区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.作成種別) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.伝票種類) & "',"		
        gSL_Select = gSL_Select & "" & .納品書管理NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SGWInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function SGWUpdate(Optional DBNo As Integer) As Boolean
    SGWUpdate = False	
    With SGW
        gSL_Select = "Update 出荷画面入力ワーク  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "処理区分 = " & .処理区分 & ","
        gSL_Select = gSL_Select & "受注管理NO = " & .受注管理NO & ","
        gSL_Select = gSL_Select & "受注管理行NO = " & .受注管理行NO & ","
        gSL_Select = gSL_Select & "受注区分 = " & .受注区分 & ","
        gSL_Select = gSL_Select & "受注日 = '" & RTrim(.受注日) & "',"
        gSL_Select = gSL_Select & "受注便 = '" & RTrim(.受注便) & "',"
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "出荷予定日 = '" & RTrim(.出荷予定日) & "',"
        gSL_Select = gSL_Select & "出荷予定便 = '" & RTrim(.出荷予定便) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "機番 = '" & RTrim(.機番) & "',"
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "納品書頁 = " & .納品書頁 & ","
        gSL_Select = gSL_Select & "納品書行 = " & .納品書行 & ","
        gSL_Select = gSL_Select & "分納区分 = " & .分納区分 & ","
        gSL_Select = gSL_Select & "分納回数 = " & .分納回数 & ","
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "枚数 = " & .枚数 & ","
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "納入数 = " & .納入数 & ","
        gSL_Select = gSL_Select & "端数 = " & .端数 & ","
        gSL_Select = gSL_Select & "出荷指示書NO = " & .出荷指示書NO & ","
        gSL_Select = gSL_Select & "出荷指示書行NO = " & .出荷指示書行NO & ","
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "SK区分 = " & .SK区分 & ","
        gSL_Select = gSL_Select & "作成種別 = '" & RTrim(.作成種別) & "',"
        gSL_Select = gSL_Select & "伝票種類 = '" & RTrim(.伝票種類) & "',"
        gSL_Select = gSL_Select & "納品書管理NO = " & .納品書管理NO & ""
        gSL_Select = gSL_Select & " Where 処理端末 = '" & RTrim(.処理端末) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SGWUpdate = True
End Function

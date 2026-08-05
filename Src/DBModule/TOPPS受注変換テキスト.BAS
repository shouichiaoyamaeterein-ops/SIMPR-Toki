Attribute VB_Name = "TOPPS受注変換テキスト"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TOPPS受注変換テキスト
'*
'*   モジュールID  : TOPPS受注変換テキスト.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/06/02  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_TOPPS受注変換テキスト
        端末番号          As String
        受注管理NO        As Long
        受注管理行NO      As Byte
        受注区分          As Byte
        受注日            As String
        受注便            As String
        納入日            As String
        納入便            As String
        納入先            As String
        受入              As String
        出荷予定日        As String
        出荷予定便        As String
        納入番号          As String
        納品書頁          As Byte
        納品書行          As Byte
        分納区分          As Byte
        分納回数          As Byte
        品番              As String
        背番号            As String
        枚数              As Long
        収容数            As Long
        納入数            As Long
        端数              As Long
        出荷指示書NO      As Long
        出荷指示書行NO    As Byte
        納品書管理NO      As Long
        元出荷指示書NO    As Long
        元出荷指示書行NO  As Byte
        ERR1              As String
        ERR2              As String
        ERR3              As String
        ERR4              As String
        ERR5              As String
        ERR6              As String
        ERR7              As String
        履歴管理番号      As Long
        受注出荷区分      As Byte
        ファイルID        As String
'        作成日            As Date
        件数          As Long
'   項目追加(自動)
	納入先名      As String
	表示品番      As String
    End Type

    Public PHW As T_TOPPS受注変換テキスト

    Public PHW_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromPHW(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearPHW	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With PHW
            Select Case adoCol.Name
                Case "端末番号":          If Not IsNull(adoCol.Value) Then .端末番号         = adoCol.Value
                Case "受注管理NO":        If Not IsNull(adoCol.Value) Then .受注管理NO       = adoCol.Value
                Case "受注管理行NO":      If Not IsNull(adoCol.Value) Then .受注管理行NO     = adoCol.Value
                Case "受注区分":          If Not IsNull(adoCol.Value) Then .受注区分         = adoCol.Value
                Case "受注日":            If Not IsNull(adoCol.Value) Then .受注日           = adoCol.Value
                Case "受注便":            If Not IsNull(adoCol.Value) Then .受注便           = adoCol.Value
                Case "納入日":            If Not IsNull(adoCol.Value) Then .納入日           = adoCol.Value
                Case "納入便":            If Not IsNull(adoCol.Value) Then .納入便           = adoCol.Value
                Case "納入先":            If Not IsNull(adoCol.Value) Then .納入先           = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入             = adoCol.Value
                Case "出荷予定日":        If Not IsNull(adoCol.Value) Then .出荷予定日       = adoCol.Value
                Case "出荷予定便":        If Not IsNull(adoCol.Value) Then .出荷予定便       = adoCol.Value
                Case "納入番号":          If Not IsNull(adoCol.Value) Then .納入番号         = adoCol.Value
                Case "納品書頁":          If Not IsNull(adoCol.Value) Then .納品書頁         = adoCol.Value
                Case "納品書行":          If Not IsNull(adoCol.Value) Then .納品書行         = adoCol.Value
                Case "分納区分":          If Not IsNull(adoCol.Value) Then .分納区分         = adoCol.Value
                Case "分納回数":          If Not IsNull(adoCol.Value) Then .分納回数         = adoCol.Value
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番             = adoCol.Value
                Case "背番号":            If Not IsNull(adoCol.Value) Then .背番号           = adoCol.Value
                Case "枚数":              If Not IsNull(adoCol.Value) Then .枚数             = adoCol.Value
                Case "収容数":            If Not IsNull(adoCol.Value) Then .収容数           = adoCol.Value
                Case "納入数":            If Not IsNull(adoCol.Value) Then .納入数           = adoCol.Value
                Case "端数":              If Not IsNull(adoCol.Value) Then .端数             = adoCol.Value
                Case "出荷指示書NO":      If Not IsNull(adoCol.Value) Then .出荷指示書NO     = adoCol.Value
                Case "出荷指示書行NO":    If Not IsNull(adoCol.Value) Then .出荷指示書行NO   = adoCol.Value
                Case "納品書管理NO":      If Not IsNull(adoCol.Value) Then .納品書管理NO     = adoCol.Value
                Case "元出荷指示書NO":    If Not IsNull(adoCol.Value) Then .元出荷指示書NO   = adoCol.Value
                Case "元出荷指示書行NO":  If Not IsNull(adoCol.Value) Then .元出荷指示書行NO = adoCol.Value
                Case "ERR1":              If Not IsNull(adoCol.Value) Then .ERR1             = adoCol.Value
                Case "ERR2":              If Not IsNull(adoCol.Value) Then .ERR2             = adoCol.Value
                Case "ERR3":              If Not IsNull(adoCol.Value) Then .ERR3             = adoCol.Value
                Case "ERR4":              If Not IsNull(adoCol.Value) Then .ERR4             = adoCol.Value
                Case "ERR5":              If Not IsNull(adoCol.Value) Then .ERR5             = adoCol.Value
                Case "ERR6":              If Not IsNull(adoCol.Value) Then .ERR6             = adoCol.Value
                Case "ERR7":              If Not IsNull(adoCol.Value) Then .ERR7             = adoCol.Value
                Case "履歴管理番号":      If Not IsNull(adoCol.Value) Then .履歴管理番号     = adoCol.Value
                Case "受注出荷区分":      If Not IsNull(adoCol.Value) Then .受注出荷区分     = adoCol.Value
                Case "ファイルID":        If Not IsNull(adoCol.Value) Then .ファイルID       = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "納入先名":	If Not IsNull(adoCol.Value) Then .納入先名	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function PHWRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    PHWRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From TOPPS受注変換テキスト"
        RTxtSql = RTxtSql &" Where 端末番号 = '" & RTrim(PHW.端末番号) & "'"
        RTxtSql = RTxtSql &" And   受注管理NO = " & PHW.受注管理NO & ""
        RTxtSql = RTxtSql &" And   受注管理行NO = " & PHW.受注管理行NO & ""
        RTxtSql = RTxtSql &" And   納入番号 = '" & RTrim(PHW.納入番号) & "'"
        RTxtSql = RTxtSql &" And   納品書行 = " & PHW.納品書行 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           PHWRead = True
           Call ItemsFromPHW(DBNo)
	   PHW_RDSTS = True
        else
	   PHW_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call PHWClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub PHWReadNext(Optional DBNo As Integer = 0)
    If Not PHW_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           PHW_RDSTS = False
        Else
           Call ItemsFromPHW(DBNo)
           PHW_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub PHWClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        PHW_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( PHW )
'*----------------------------------------------*'
Public Sub ItemsClearPHW()
    With PHW
	.端末番号 = ""
	.受注管理NO = 0
	.受注管理行NO = 0
	.受注区分 = 0
	.受注日 = ""
	.受注便 = ""
	.納入日 = ""
	.納入便 = ""
	.納入先 = ""
	.受入 = ""
	.出荷予定日 = ""
	.出荷予定便 = ""
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
	.納品書管理NO = 0
	.元出荷指示書NO = 0
	.元出荷指示書行NO = 0
	.ERR1 = ""
	.ERR2 = ""
	.ERR3 = ""
	.ERR4 = ""
	.ERR5 = ""
	.ERR6 = ""
	.ERR7 = ""
	.履歴管理番号 = 0
	.受注出荷区分 = 0
	.ファイルID = ""
'	.作成日 = 0
        .件数 = 0
'   項目追加(自動)
	.納入先名 = ""
	.表示品番 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function PHWInsert(Optional DBNo As Integer) As Boolean
    PHWInsert = True
    With PHW
        gSL_Select = "Insert Into TOPPS受注変換テキスト  ( "
        gSL_Select = gSL_Select & "端末番号,受注管理NO,受注管理行NO,"
        gSL_Select = gSL_Select & "受注区分,受注日,受注便,"
        gSL_Select = gSL_Select & "納入日,納入便,納入先,"
        gSL_Select = gSL_Select & "受入,出荷予定日,出荷予定便,"
        gSL_Select = gSL_Select & "納入番号,納品書頁,納品書行,"
        gSL_Select = gSL_Select & "分納区分,分納回数,品番,"
        gSL_Select = gSL_Select & "背番号,枚数,収容数,"
        gSL_Select = gSL_Select & "納入数,端数,出荷指示書NO,"
        gSL_Select = gSL_Select & "出荷指示書行NO,納品書管理NO,元出荷指示書NO,"
        gSL_Select = gSL_Select & "元出荷指示書行NO,ERR1,ERR2,"
        gSL_Select = gSL_Select & "ERR3,ERR4,ERR5,"
        gSL_Select = gSL_Select & "ERR6,ERR7,履歴管理番号,"
        gSL_Select = gSL_Select & "受注出荷区分,ファイルID,"
        gSL_Select = gSL_Select & "作成日"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.端末番号) & "',"		
        gSL_Select = gSL_Select & "" & .受注管理NO & ","		
        gSL_Select = gSL_Select & "" & .受注管理行NO & ","		
        gSL_Select = gSL_Select & "" & .受注区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.受注日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受注便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.出荷予定日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.出荷予定便) & "',"		
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
        gSL_Select = gSL_Select & "" & .納品書管理NO & ","		
        gSL_Select = gSL_Select & "" & .元出荷指示書NO & ","		
        gSL_Select = gSL_Select & "" & .元出荷指示書行NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.ERR1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ERR2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ERR3) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ERR4) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ERR5) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ERR6) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ERR7) & "',"		
        gSL_Select = gSL_Select & "" & .履歴管理番号 & ","		
        gSL_Select = gSL_Select & "" & .受注出荷区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.ファイルID) & "',"		
        gSL_Select = gSL_Select & "GETDATE()"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    PHWInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function PHWUpdate(Optional DBNo As Integer) As Boolean
    PHWUpdate = False	
    With PHW
        gSL_Select = "Update TOPPS受注変換テキスト  Set "
        gSL_Select = gSL_Select & "端末番号 = '" & RTrim(.端末番号) & "',"
        gSL_Select = gSL_Select & "受注管理NO = " & .受注管理NO & ","
        gSL_Select = gSL_Select & "受注管理行NO = " & .受注管理行NO & ","
        gSL_Select = gSL_Select & "受注区分 = " & .受注区分 & ","
        gSL_Select = gSL_Select & "受注日 = '" & RTrim(.受注日) & "',"
        gSL_Select = gSL_Select & "受注便 = '" & RTrim(.受注便) & "',"
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "出荷予定日 = '" & RTrim(.出荷予定日) & "',"
        gSL_Select = gSL_Select & "出荷予定便 = '" & RTrim(.出荷予定便) & "',"
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
        gSL_Select = gSL_Select & "納品書管理NO = " & .納品書管理NO & ","
        gSL_Select = gSL_Select & "元出荷指示書NO = " & .元出荷指示書NO & ","
        gSL_Select = gSL_Select & "元出荷指示書行NO = " & .元出荷指示書行NO & ","
        gSL_Select = gSL_Select & "ERR1 = '" & RTrim(.ERR1) & "',"
        gSL_Select = gSL_Select & "ERR2 = '" & RTrim(.ERR2) & "',"
        gSL_Select = gSL_Select & "ERR3 = '" & RTrim(.ERR3) & "',"
        gSL_Select = gSL_Select & "ERR4 = '" & RTrim(.ERR4) & "',"
        gSL_Select = gSL_Select & "ERR5 = '" & RTrim(.ERR5) & "',"
        gSL_Select = gSL_Select & "ERR6 = '" & RTrim(.ERR6) & "',"
        gSL_Select = gSL_Select & "ERR7 = '" & RTrim(.ERR7) & "',"
        gSL_Select = gSL_Select & "履歴管理番号 = " & .履歴管理番号 & ","
        gSL_Select = gSL_Select & "受注出荷区分 = " & .受注出荷区分 & ","
        gSL_Select = gSL_Select & "ファイルID = '" & RTrim(.ファイルID) & "'"
        gSL_Select = gSL_Select & " Where 端末番号 = '" & RTrim(.端末番号) & "'"
        gSL_Select = gSL_Select & "   And 受注管理NO = " & .受注管理NO & ""
        gSL_Select = gSL_Select & "   And 受注管理行NO = " & .受注管理行NO & ""
        gSL_Select = gSL_Select & "   And 納入番号 = '" & RTrim(.納入番号) & "'"
        gSL_Select = gSL_Select & "   And 納品書行 = " & .納品書行 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    PHWUpdate = True
End Function

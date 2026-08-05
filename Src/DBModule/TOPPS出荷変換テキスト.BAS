Attribute VB_Name = "TOPPS出荷変換テキスト"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TOPPS出荷変換テキスト
'*
'*   モジュールID  : TOPPS出荷変換テキスト.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/04/28  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_TOPPS出荷変換テキスト
        端末番号                  As String
        仕入先                    As String
        工区                      As String
        出荷場                    As String
        納入先                    As String
        受入                      As String
        便区分                    As Byte
        納入日                    As String
        納入便                    As String
        挽回日                    As String
        挽回便                    As String
        納入番号                  As String
        輸送中FLG                 As Byte
        処理日                    As String
        処理時間                  As String
        行番号                    As Integer
        背番号                    As String
        品番                      As String
        収容数                    As Long
        打切端数                  As Long
        出荷個数                  As Long
        打切区分                  As Byte
        端数打出FLG               As Byte
        前工程情報任意項目        As String
        ERR                       As String
        履歴管理番号              As Long
        受注出荷区分              As Byte
        ファイルID                As String
        受注管理NO                As Long
        受注管理行NO              As Byte
        出荷指示書NO              As Long
        出荷指示書行NO            As Byte
'        作成日                    As Date
        件数          As Long
'   項目追加(自動)
	表示品番      As String
	納入先名      As String
	枚数          As Long
	端数          As Long
    End Type

    Public PSW As T_TOPPS出荷変換テキスト

    Public PSW_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromPSW(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearPSW	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With PSW
            Select Case adoCol.Name
                Case "端末番号":                  If Not IsNull(adoCol.Value) Then .端末番号                 = adoCol.Value
                Case "仕入先":                    If Not IsNull(adoCol.Value) Then .仕入先                   = adoCol.Value
                Case "工区":                      If Not IsNull(adoCol.Value) Then .工区                     = adoCol.Value
                Case "出荷場":                    If Not IsNull(adoCol.Value) Then .出荷場                   = adoCol.Value
                Case "納入先":                    If Not IsNull(adoCol.Value) Then .納入先                   = adoCol.Value
                Case "受入":                      If Not IsNull(adoCol.Value) Then .受入                     = adoCol.Value
                Case "便区分":                    If Not IsNull(adoCol.Value) Then .便区分                   = adoCol.Value
                Case "納入日":                    If Not IsNull(adoCol.Value) Then .納入日                   = adoCol.Value
                Case "納入便":                    If Not IsNull(adoCol.Value) Then .納入便                   = adoCol.Value
                Case "挽回日":                    If Not IsNull(adoCol.Value) Then .挽回日                   = adoCol.Value
                Case "挽回便":                    If Not IsNull(adoCol.Value) Then .挽回便                   = adoCol.Value
                Case "納入番号":                  If Not IsNull(adoCol.Value) Then .納入番号                 = adoCol.Value
                Case "輸送中FLG":                 If Not IsNull(adoCol.Value) Then .輸送中FLG                = adoCol.Value
                Case "処理日":                    If Not IsNull(adoCol.Value) Then .処理日                   = adoCol.Value
                Case "処理時間":                  If Not IsNull(adoCol.Value) Then .処理時間                 = adoCol.Value
                Case "行番号":                    If Not IsNull(adoCol.Value) Then .行番号                   = adoCol.Value
                Case "背番号":                    If Not IsNull(adoCol.Value) Then .背番号                   = adoCol.Value
                Case "品番":                      If Not IsNull(adoCol.Value) Then .品番                     = adoCol.Value
                Case "収容数":                    If Not IsNull(adoCol.Value) Then .収容数                   = adoCol.Value
                Case "打切端数":                  If Not IsNull(adoCol.Value) Then .打切端数                 = adoCol.Value
                Case "出荷個数":                  If Not IsNull(adoCol.Value) Then .出荷個数                 = adoCol.Value
                Case "打切区分":                  If Not IsNull(adoCol.Value) Then .打切区分                 = adoCol.Value
                Case "端数打出FLG":               If Not IsNull(adoCol.Value) Then .端数打出FLG              = adoCol.Value
                Case "前工程情報任意項目":        If Not IsNull(adoCol.Value) Then .前工程情報任意項目       = adoCol.Value
                Case "ERR":                       If Not IsNull(adoCol.Value) Then .ERR                      = adoCol.Value
                Case "履歴管理番号":              If Not IsNull(adoCol.Value) Then .履歴管理番号             = adoCol.Value
                Case "受注出荷区分":              If Not IsNull(adoCol.Value) Then .受注出荷区分             = adoCol.Value
                Case "ファイルID":                If Not IsNull(adoCol.Value) Then .ファイルID               = adoCol.Value
                Case "受注管理NO":                If Not IsNull(adoCol.Value) Then .受注管理NO               = adoCol.Value
                Case "受注管理行NO":              If Not IsNull(adoCol.Value) Then .受注管理行NO             = adoCol.Value
                Case "出荷指示書NO":              If Not IsNull(adoCol.Value) Then .出荷指示書NO             = adoCol.Value
                Case "出荷指示書行NO":            If Not IsNull(adoCol.Value) Then .出荷指示書行NO           = adoCol.Value
'                Case "作成日":                    If Not IsNull(adoCol.Value) Then .作成日                   = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "納入先名":	If Not IsNull(adoCol.Value) Then .納入先名	= adoCol.Value
		Case "枚数":	If Not IsNull(adoCol.Value) Then .枚数	= adoCol.Value
		Case "端数":	If Not IsNull(adoCol.Value) Then .端数	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function PSWRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    PSWRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From TOPPS出荷変換テキスト"
        RTxtSql = RTxtSql &" Where 端末番号 = '" & RTrim(PSW.端末番号) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           PSWRead = True
           Call ItemsFromPSW(DBNo)
	   PSW_RDSTS = True
        else
	   PSW_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call PSWClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub PSWReadNext(Optional DBNo As Integer = 0)
    If Not PSW_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           PSW_RDSTS = False
        Else
           Call ItemsFromPSW(DBNo)
           PSW_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub PSWClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        PSW_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( PSW )
'*----------------------------------------------*'
Public Sub ItemsClearPSW()
    With PSW
	.端末番号 = ""
	.仕入先 = ""
	.工区 = ""
	.出荷場 = ""
	.納入先 = ""
	.受入 = ""
	.便区分 = 0
	.納入日 = ""
	.納入便 = ""
	.挽回日 = ""
	.挽回便 = ""
	.納入番号 = ""
	.輸送中FLG = 0
	.処理日 = ""
	.処理時間 = ""
	.行番号 = 0
	.背番号 = ""
	.品番 = ""
	.収容数 = 0
	.打切端数 = 0
	.出荷個数 = 0
	.打切区分 = 0
	.端数打出FLG = 0
	.前工程情報任意項目 = ""
	.ERR = ""
	.履歴管理番号 = 0
	.受注出荷区分 = 0
	.ファイルID = ""
	.受注管理NO = 0
	.受注管理行NO = 0
	.出荷指示書NO = 0
	.出荷指示書行NO = 0
'	.作成日 = 0
        .件数 = 0
'   項目追加(自動)
	.表示品番 = ""
	.納入先名 = ""
	.枚数 = 0
	.端数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function PSWInsert(Optional DBNo As Integer) As Boolean
    PSWInsert = True
    With PSW
        gSL_Select = "Insert Into TOPPS出荷変換テキスト  ( "
        gSL_Select = gSL_Select & "端末番号,仕入先,工区,"
        gSL_Select = gSL_Select & "出荷場,納入先,受入,"
        gSL_Select = gSL_Select & "便区分,納入日,納入便,"
        gSL_Select = gSL_Select & "挽回日,挽回便,納入番号,"
        gSL_Select = gSL_Select & "輸送中FLG,処理日,処理時間,"
        gSL_Select = gSL_Select & "行番号,背番号,品番,"
        gSL_Select = gSL_Select & "収容数,打切端数,出荷個数,"
        gSL_Select = gSL_Select & "打切区分,端数打出FLG,前工程情報任意項目,"
        gSL_Select = gSL_Select & "ERR,履歴管理番号,受注出荷区分,"
        gSL_Select = gSL_Select & "ファイルID,受注管理NO,受注管理行NO,"
        gSL_Select = gSL_Select & "出荷指示書NO,出荷指示書行NO,"
        gSL_Select = gSL_Select & "作成日"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.端末番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.工区) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.出荷場) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "" & .便区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.挽回日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.挽回便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"		
        gSL_Select = gSL_Select & "" & .輸送中FLG & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.処理日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.処理時間) & "',"		
        gSL_Select = gSL_Select & "" & .行番号 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "" & .収容数 & ","		
        gSL_Select = gSL_Select & "" & .打切端数 & ","		
        gSL_Select = gSL_Select & "" & .出荷個数 & ","		
        gSL_Select = gSL_Select & "" & .打切区分 & ","		
        gSL_Select = gSL_Select & "" & .端数打出FLG & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.前工程情報任意項目) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ERR) & "',"		
        gSL_Select = gSL_Select & "" & .履歴管理番号 & ","		
        gSL_Select = gSL_Select & "" & .受注出荷区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.ファイルID) & "',"		
        gSL_Select = gSL_Select & "" & .受注管理NO & ","		
        gSL_Select = gSL_Select & "" & .受注管理行NO & ","		
        gSL_Select = gSL_Select & "" & .出荷指示書NO & ","		
        gSL_Select = gSL_Select & "" & .出荷指示書行NO & ","		
        gSL_Select = gSL_Select & "GETDATE()"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    PSWInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function PSWUpdate(Optional DBNo As Integer) As Boolean
    PSWUpdate = False	
    With PSW
        gSL_Select = "Update TOPPS出荷変換テキスト  Set "
        gSL_Select = gSL_Select & "端末番号 = '" & RTrim(.端末番号) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "工区 = '" & RTrim(.工区) & "',"
        gSL_Select = gSL_Select & "出荷場 = '" & RTrim(.出荷場) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "便区分 = " & .便区分 & ","
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "挽回日 = '" & RTrim(.挽回日) & "',"
        gSL_Select = gSL_Select & "挽回便 = '" & RTrim(.挽回便) & "',"
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "輸送中FLG = " & .輸送中FLG & ","
        gSL_Select = gSL_Select & "処理日 = '" & RTrim(.処理日) & "',"
        gSL_Select = gSL_Select & "処理時間 = '" & RTrim(.処理時間) & "',"
        gSL_Select = gSL_Select & "行番号 = " & .行番号 & ","
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "打切端数 = " & .打切端数 & ","
        gSL_Select = gSL_Select & "出荷個数 = " & .出荷個数 & ","
        gSL_Select = gSL_Select & "打切区分 = " & .打切区分 & ","
        gSL_Select = gSL_Select & "端数打出FLG = " & .端数打出FLG & ","
        gSL_Select = gSL_Select & "前工程情報任意項目 = '" & RTrim(.前工程情報任意項目) & "',"
        gSL_Select = gSL_Select & "ERR = '" & RTrim(.ERR) & "',"
        gSL_Select = gSL_Select & "履歴管理番号 = " & .履歴管理番号 & ","
        gSL_Select = gSL_Select & "受注出荷区分 = " & .受注出荷区分 & ","
        gSL_Select = gSL_Select & "ファイルID = '" & RTrim(.ファイルID) & "',"
        gSL_Select = gSL_Select & "受注管理NO = " & .受注管理NO & ","
        gSL_Select = gSL_Select & "受注管理行NO = " & .受注管理行NO & ","
        gSL_Select = gSL_Select & "出荷指示書NO = " & .出荷指示書NO & ","
        gSL_Select = gSL_Select & "出荷指示書行NO = " & .出荷指示書行NO & ""
        gSL_Select = gSL_Select & " Where 端末番号 = '" & RTrim(.端末番号) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    PSWUpdate = True
End Function

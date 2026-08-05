Attribute VB_Name = "S内示問合せワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   S内示問合せワーク
'*
'*   モジュールID  : S内示問合せワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/05/20  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_S内示問合せワーク
        処理端末  As String
        年月      As String
        順位      As String
        レベル    As Byte
        品番      As String
        手配先    As String
        受入      As String
        直送先    As String
        区分      As Byte
        当月数    As Long
        翌月数    As Long
        翌々月数  As Long
        日量数01  As Currency
        日量数02  As Currency
        日量数03  As Currency
        日量数04  As Currency
        日量数05  As Currency
        日量数06  As Currency
        日量数07  As Currency
        日量数08  As Currency
        日量数09  As Currency
        日量数10  As Currency
        日量数11  As Currency
        日量数12  As Currency
        日量数13  As Currency
        日量数14  As Currency
        日量数15  As Currency
        日量数16  As Currency
        日量数17  As Currency
        日量数18  As Currency
        日量数19  As Currency
        日量数20  As Currency
        日量数21  As Currency
        日量数22  As Currency
        日量数23  As Currency
        日量数24  As Currency
        日量数25  As Currency
        日量数26  As Currency
        日量数27  As Currency
        日量数28  As Currency
        日量数29  As Currency
        日量数30  As Currency
        日量数31  As Currency
        件数          As Long
'   項目追加(自動)
	手配先名      As String
	収容数        As Long
	SK区分        As Byte
	表示品番      As String
	ロット        As Long
	背番号        As String
    End Type

    Public SNW As T_S内示問合せワーク

    Public SNW_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromSNW(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearSNW	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With SNW
            Select Case adoCol.Name
                Case "処理端末":  If Not IsNull(adoCol.Value) Then .処理端末 = adoCol.Value
                Case "年月":      If Not IsNull(adoCol.Value) Then .年月     = adoCol.Value
                Case "順位":      If Not IsNull(adoCol.Value) Then .順位     = adoCol.Value
                Case "レベル":    If Not IsNull(adoCol.Value) Then .レベル   = adoCol.Value
                Case "品番":      If Not IsNull(adoCol.Value) Then .品番     = adoCol.Value
                Case "手配先":    If Not IsNull(adoCol.Value) Then .手配先   = adoCol.Value
                Case "受入":      If Not IsNull(adoCol.Value) Then .受入     = adoCol.Value
                Case "直送先":    If Not IsNull(adoCol.Value) Then .直送先   = adoCol.Value
                Case "区分":      If Not IsNull(adoCol.Value) Then .区分     = adoCol.Value
                Case "当月数":    If Not IsNull(adoCol.Value) Then .当月数   = adoCol.Value
                Case "翌月数":    If Not IsNull(adoCol.Value) Then .翌月数   = adoCol.Value
                Case "翌々月数":  If Not IsNull(adoCol.Value) Then .翌々月数 = adoCol.Value
                Case "日量数01":  If Not IsNull(adoCol.Value) Then .日量数01 = adoCol.Value
                Case "日量数02":  If Not IsNull(adoCol.Value) Then .日量数02 = adoCol.Value
                Case "日量数03":  If Not IsNull(adoCol.Value) Then .日量数03 = adoCol.Value
                Case "日量数04":  If Not IsNull(adoCol.Value) Then .日量数04 = adoCol.Value
                Case "日量数05":  If Not IsNull(adoCol.Value) Then .日量数05 = adoCol.Value
                Case "日量数06":  If Not IsNull(adoCol.Value) Then .日量数06 = adoCol.Value
                Case "日量数07":  If Not IsNull(adoCol.Value) Then .日量数07 = adoCol.Value
                Case "日量数08":  If Not IsNull(adoCol.Value) Then .日量数08 = adoCol.Value
                Case "日量数09":  If Not IsNull(adoCol.Value) Then .日量数09 = adoCol.Value
                Case "日量数10":  If Not IsNull(adoCol.Value) Then .日量数10 = adoCol.Value
                Case "日量数11":  If Not IsNull(adoCol.Value) Then .日量数11 = adoCol.Value
                Case "日量数12":  If Not IsNull(adoCol.Value) Then .日量数12 = adoCol.Value
                Case "日量数13":  If Not IsNull(adoCol.Value) Then .日量数13 = adoCol.Value
                Case "日量数14":  If Not IsNull(adoCol.Value) Then .日量数14 = adoCol.Value
                Case "日量数15":  If Not IsNull(adoCol.Value) Then .日量数15 = adoCol.Value
                Case "日量数16":  If Not IsNull(adoCol.Value) Then .日量数16 = adoCol.Value
                Case "日量数17":  If Not IsNull(adoCol.Value) Then .日量数17 = adoCol.Value
                Case "日量数18":  If Not IsNull(adoCol.Value) Then .日量数18 = adoCol.Value
                Case "日量数19":  If Not IsNull(adoCol.Value) Then .日量数19 = adoCol.Value
                Case "日量数20":  If Not IsNull(adoCol.Value) Then .日量数20 = adoCol.Value
                Case "日量数21":  If Not IsNull(adoCol.Value) Then .日量数21 = adoCol.Value
                Case "日量数22":  If Not IsNull(adoCol.Value) Then .日量数22 = adoCol.Value
                Case "日量数23":  If Not IsNull(adoCol.Value) Then .日量数23 = adoCol.Value
                Case "日量数24":  If Not IsNull(adoCol.Value) Then .日量数24 = adoCol.Value
                Case "日量数25":  If Not IsNull(adoCol.Value) Then .日量数25 = adoCol.Value
                Case "日量数26":  If Not IsNull(adoCol.Value) Then .日量数26 = adoCol.Value
                Case "日量数27":  If Not IsNull(adoCol.Value) Then .日量数27 = adoCol.Value
                Case "日量数28":  If Not IsNull(adoCol.Value) Then .日量数28 = adoCol.Value
                Case "日量数29":  If Not IsNull(adoCol.Value) Then .日量数29 = adoCol.Value
                Case "日量数30":  If Not IsNull(adoCol.Value) Then .日量数30 = adoCol.Value
                Case "日量数31":  If Not IsNull(adoCol.Value) Then .日量数31 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "手配先名":	If Not IsNull(adoCol.Value) Then .手配先名	= adoCol.Value
		Case "収容数":	If Not IsNull(adoCol.Value) Then .収容数	= adoCol.Value
		Case "SK区分":	If Not IsNull(adoCol.Value) Then .SK区分	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "ロット":	If Not IsNull(adoCol.Value) Then .ロット	= adoCol.Value
		Case "背番号":	If Not IsNull(adoCol.Value) Then .背番号	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function SNWRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    SNWRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From S内示問合せワーク"
        RTxtSql = RTxtSql &" Where 順位 = '" & RTrim(SNW.順位) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           SNWRead = True
           Call ItemsFromSNW(DBNo)
	   SNW_RDSTS = True
        else
	   SNW_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call SNWClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub SNWReadNext(Optional DBNo As Integer = 0)
    If Not SNW_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           SNW_RDSTS = False
        Else
           Call ItemsFromSNW(DBNo)
           SNW_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub SNWClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        SNW_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( SNW )
'*----------------------------------------------*'
Public Sub ItemsClearSNW()
    With SNW
	.処理端末 = ""
	.年月 = ""
	.順位 = ""
	.レベル = 0
	.品番 = ""
	.手配先 = ""
	.受入 = ""
	.直送先 = ""
	.区分 = 0
	.当月数 = 0
	.翌月数 = 0
	.翌々月数 = 0
	.日量数01 = 0
	.日量数02 = 0
	.日量数03 = 0
	.日量数04 = 0
	.日量数05 = 0
	.日量数06 = 0
	.日量数07 = 0
	.日量数08 = 0
	.日量数09 = 0
	.日量数10 = 0
	.日量数11 = 0
	.日量数12 = 0
	.日量数13 = 0
	.日量数14 = 0
	.日量数15 = 0
	.日量数16 = 0
	.日量数17 = 0
	.日量数18 = 0
	.日量数19 = 0
	.日量数20 = 0
	.日量数21 = 0
	.日量数22 = 0
	.日量数23 = 0
	.日量数24 = 0
	.日量数25 = 0
	.日量数26 = 0
	.日量数27 = 0
	.日量数28 = 0
	.日量数29 = 0
	.日量数30 = 0
	.日量数31 = 0
        .件数 = 0
'   項目追加(自動)
	.手配先名 = ""
	.収容数 = 0
	.SK区分 = 0
	.表示品番 = ""
	.ロット = 0
	.背番号 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function SNWInsert(Optional DBNo As Integer) As Boolean
    SNWInsert = True
    With SNW
        gSL_Select = "Insert Into S内示問合せワーク  ( "
        gSL_Select = gSL_Select & "処理端末,年月,順位,"
        gSL_Select = gSL_Select & "レベル,品番,手配先,"
        gSL_Select = gSL_Select & "受入,直送先,区分,"
        gSL_Select = gSL_Select & "当月数,翌月数,翌々月数,"
        gSL_Select = gSL_Select & "日量数01,日量数02,日量数03,"
        gSL_Select = gSL_Select & "日量数04,日量数05,日量数06,"
        gSL_Select = gSL_Select & "日量数07,日量数08,日量数09,"
        gSL_Select = gSL_Select & "日量数10,日量数11,日量数12,"
        gSL_Select = gSL_Select & "日量数13,日量数14,日量数15,"
        gSL_Select = gSL_Select & "日量数16,日量数17,日量数18,"
        gSL_Select = gSL_Select & "日量数19,日量数20,日量数21,"
        gSL_Select = gSL_Select & "日量数22,日量数23,日量数24,"
        gSL_Select = gSL_Select & "日量数25,日量数26,日量数27,"
        gSL_Select = gSL_Select & "日量数28,日量数29,日量数30,"
        gSL_Select = gSL_Select & "日量数31"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.順位) & "',"		
        gSL_Select = gSL_Select & "" & .レベル & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.直送先) & "',"		
        gSL_Select = gSL_Select & "" & .区分 & ","		
        gSL_Select = gSL_Select & "" & .当月数 & ","		
        gSL_Select = gSL_Select & "" & .翌月数 & ","		
        gSL_Select = gSL_Select & "" & .翌々月数 & ","		
        gSL_Select = gSL_Select & "" & .日量数01 & ","		
        gSL_Select = gSL_Select & "" & .日量数02 & ","		
        gSL_Select = gSL_Select & "" & .日量数03 & ","		
        gSL_Select = gSL_Select & "" & .日量数04 & ","		
        gSL_Select = gSL_Select & "" & .日量数05 & ","		
        gSL_Select = gSL_Select & "" & .日量数06 & ","		
        gSL_Select = gSL_Select & "" & .日量数07 & ","		
        gSL_Select = gSL_Select & "" & .日量数08 & ","		
        gSL_Select = gSL_Select & "" & .日量数09 & ","		
        gSL_Select = gSL_Select & "" & .日量数10 & ","		
        gSL_Select = gSL_Select & "" & .日量数11 & ","		
        gSL_Select = gSL_Select & "" & .日量数12 & ","		
        gSL_Select = gSL_Select & "" & .日量数13 & ","		
        gSL_Select = gSL_Select & "" & .日量数14 & ","		
        gSL_Select = gSL_Select & "" & .日量数15 & ","		
        gSL_Select = gSL_Select & "" & .日量数16 & ","		
        gSL_Select = gSL_Select & "" & .日量数17 & ","		
        gSL_Select = gSL_Select & "" & .日量数18 & ","		
        gSL_Select = gSL_Select & "" & .日量数19 & ","		
        gSL_Select = gSL_Select & "" & .日量数20 & ","		
        gSL_Select = gSL_Select & "" & .日量数21 & ","		
        gSL_Select = gSL_Select & "" & .日量数22 & ","		
        gSL_Select = gSL_Select & "" & .日量数23 & ","		
        gSL_Select = gSL_Select & "" & .日量数24 & ","		
        gSL_Select = gSL_Select & "" & .日量数25 & ","		
        gSL_Select = gSL_Select & "" & .日量数26 & ","		
        gSL_Select = gSL_Select & "" & .日量数27 & ","		
        gSL_Select = gSL_Select & "" & .日量数28 & ","		
        gSL_Select = gSL_Select & "" & .日量数29 & ","		
        gSL_Select = gSL_Select & "" & .日量数30 & ","		
        gSL_Select = gSL_Select & "" & .日量数31 & ""		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SNWInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function SNWUpdate(Optional DBNo As Integer) As Boolean
    SNWUpdate = False	
    With SNW
        gSL_Select = "Update S内示問合せワーク  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "年月 = '" & RTrim(.年月) & "',"
        gSL_Select = gSL_Select & "順位 = '" & RTrim(.順位) & "',"
        gSL_Select = gSL_Select & "レベル = " & .レベル & ","
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "手配先 = '" & RTrim(.手配先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "直送先 = '" & RTrim(.直送先) & "',"
        gSL_Select = gSL_Select & "区分 = " & .区分 & ","
        gSL_Select = gSL_Select & "当月数 = " & .当月数 & ","
        gSL_Select = gSL_Select & "翌月数 = " & .翌月数 & ","
        gSL_Select = gSL_Select & "翌々月数 = " & .翌々月数 & ","
        gSL_Select = gSL_Select & "日量数01 = " & .日量数01 & ","
        gSL_Select = gSL_Select & "日量数02 = " & .日量数02 & ","
        gSL_Select = gSL_Select & "日量数03 = " & .日量数03 & ","
        gSL_Select = gSL_Select & "日量数04 = " & .日量数04 & ","
        gSL_Select = gSL_Select & "日量数05 = " & .日量数05 & ","
        gSL_Select = gSL_Select & "日量数06 = " & .日量数06 & ","
        gSL_Select = gSL_Select & "日量数07 = " & .日量数07 & ","
        gSL_Select = gSL_Select & "日量数08 = " & .日量数08 & ","
        gSL_Select = gSL_Select & "日量数09 = " & .日量数09 & ","
        gSL_Select = gSL_Select & "日量数10 = " & .日量数10 & ","
        gSL_Select = gSL_Select & "日量数11 = " & .日量数11 & ","
        gSL_Select = gSL_Select & "日量数12 = " & .日量数12 & ","
        gSL_Select = gSL_Select & "日量数13 = " & .日量数13 & ","
        gSL_Select = gSL_Select & "日量数14 = " & .日量数14 & ","
        gSL_Select = gSL_Select & "日量数15 = " & .日量数15 & ","
        gSL_Select = gSL_Select & "日量数16 = " & .日量数16 & ","
        gSL_Select = gSL_Select & "日量数17 = " & .日量数17 & ","
        gSL_Select = gSL_Select & "日量数18 = " & .日量数18 & ","
        gSL_Select = gSL_Select & "日量数19 = " & .日量数19 & ","
        gSL_Select = gSL_Select & "日量数20 = " & .日量数20 & ","
        gSL_Select = gSL_Select & "日量数21 = " & .日量数21 & ","
        gSL_Select = gSL_Select & "日量数22 = " & .日量数22 & ","
        gSL_Select = gSL_Select & "日量数23 = " & .日量数23 & ","
        gSL_Select = gSL_Select & "日量数24 = " & .日量数24 & ","
        gSL_Select = gSL_Select & "日量数25 = " & .日量数25 & ","
        gSL_Select = gSL_Select & "日量数26 = " & .日量数26 & ","
        gSL_Select = gSL_Select & "日量数27 = " & .日量数27 & ","
        gSL_Select = gSL_Select & "日量数28 = " & .日量数28 & ","
        gSL_Select = gSL_Select & "日量数29 = " & .日量数29 & ","
        gSL_Select = gSL_Select & "日量数30 = " & .日量数30 & ","
        gSL_Select = gSL_Select & "日量数31 = " & .日量数31 & ""
        gSL_Select = gSL_Select & " Where 順位 = '" & RTrim(.順位) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SNWUpdate = True
End Function

Attribute VB_Name = "S内示テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   S内示テーブル
'*
'*   モジュールID  : S内示テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/05/10  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_S内示テーブル
        内示年月      As String
        品番          As String
        納入先        As String
        受入          As String
        内示数        As Long
        翌月内示数    As Long
        翌々月内示数  As Long
        繰越在庫数    As Long
        収容数        As Long
        背番号        As String
        サイクル      As String
        内示数量01    As Long
        内示数量02    As Long
        内示数量03    As Long
        内示数量04    As Long
        内示数量05    As Long
        内示数量06    As Long
        内示数量07    As Long
        内示数量08    As Long
        内示数量09    As Long
        内示数量10    As Long
        内示数量11    As Long
        内示数量12    As Long
        内示数量13    As Long
        内示数量14    As Long
        内示数量15    As Long
        内示数量16    As Long
        内示数量17    As Long
        内示数量18    As Long
        内示数量19    As Long
        内示数量20    As Long
        内示数量21    As Long
        内示数量22    As Long
        内示数量23    As Long
        内示数量24    As Long
        内示数量25    As Long
        内示数量26    As Long
        内示数量27    As Long
        内示数量28    As Long
        内示数量29    As Long
        内示数量30    As Long
        内示数量31    As Long
        展開区分      As Byte
        作成者        As String
'        作成日        As Date
'        作成端末      As String
        件数          As Long
    End Type

    Public SNJ As T_S内示テーブル

    Public SNJ_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromSNJ(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearSNJ	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With SNJ
            Select Case adoCol.Name
                Case "内示年月":      If Not IsNull(adoCol.Value) Then .内示年月     = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "納入先":        If Not IsNull(adoCol.Value) Then .納入先       = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "内示数":        If Not IsNull(adoCol.Value) Then .内示数       = adoCol.Value
                Case "翌月内示数":    If Not IsNull(adoCol.Value) Then .翌月内示数   = adoCol.Value
                Case "翌々月内示数":  If Not IsNull(adoCol.Value) Then .翌々月内示数 = adoCol.Value
                Case "繰越在庫数":    If Not IsNull(adoCol.Value) Then .繰越在庫数   = adoCol.Value
                Case "収容数":        If Not IsNull(adoCol.Value) Then .収容数       = adoCol.Value
                Case "背番号":        If Not IsNull(adoCol.Value) Then .背番号       = adoCol.Value
                Case "サイクル":      If Not IsNull(adoCol.Value) Then .サイクル     = adoCol.Value
                Case "内示数量01":    If Not IsNull(adoCol.Value) Then .内示数量01   = adoCol.Value
                Case "内示数量02":    If Not IsNull(adoCol.Value) Then .内示数量02   = adoCol.Value
                Case "内示数量03":    If Not IsNull(adoCol.Value) Then .内示数量03   = adoCol.Value
                Case "内示数量04":    If Not IsNull(adoCol.Value) Then .内示数量04   = adoCol.Value
                Case "内示数量05":    If Not IsNull(adoCol.Value) Then .内示数量05   = adoCol.Value
                Case "内示数量06":    If Not IsNull(adoCol.Value) Then .内示数量06   = adoCol.Value
                Case "内示数量07":    If Not IsNull(adoCol.Value) Then .内示数量07   = adoCol.Value
                Case "内示数量08":    If Not IsNull(adoCol.Value) Then .内示数量08   = adoCol.Value
                Case "内示数量09":    If Not IsNull(adoCol.Value) Then .内示数量09   = adoCol.Value
                Case "内示数量10":    If Not IsNull(adoCol.Value) Then .内示数量10   = adoCol.Value
                Case "内示数量11":    If Not IsNull(adoCol.Value) Then .内示数量11   = adoCol.Value
                Case "内示数量12":    If Not IsNull(adoCol.Value) Then .内示数量12   = adoCol.Value
                Case "内示数量13":    If Not IsNull(adoCol.Value) Then .内示数量13   = adoCol.Value
                Case "内示数量14":    If Not IsNull(adoCol.Value) Then .内示数量14   = adoCol.Value
                Case "内示数量15":    If Not IsNull(adoCol.Value) Then .内示数量15   = adoCol.Value
                Case "内示数量16":    If Not IsNull(adoCol.Value) Then .内示数量16   = adoCol.Value
                Case "内示数量17":    If Not IsNull(adoCol.Value) Then .内示数量17   = adoCol.Value
                Case "内示数量18":    If Not IsNull(adoCol.Value) Then .内示数量18   = adoCol.Value
                Case "内示数量19":    If Not IsNull(adoCol.Value) Then .内示数量19   = adoCol.Value
                Case "内示数量20":    If Not IsNull(adoCol.Value) Then .内示数量20   = adoCol.Value
                Case "内示数量21":    If Not IsNull(adoCol.Value) Then .内示数量21   = adoCol.Value
                Case "内示数量22":    If Not IsNull(adoCol.Value) Then .内示数量22   = adoCol.Value
                Case "内示数量23":    If Not IsNull(adoCol.Value) Then .内示数量23   = adoCol.Value
                Case "内示数量24":    If Not IsNull(adoCol.Value) Then .内示数量24   = adoCol.Value
                Case "内示数量25":    If Not IsNull(adoCol.Value) Then .内示数量25   = adoCol.Value
                Case "内示数量26":    If Not IsNull(adoCol.Value) Then .内示数量26   = adoCol.Value
                Case "内示数量27":    If Not IsNull(adoCol.Value) Then .内示数量27   = adoCol.Value
                Case "内示数量28":    If Not IsNull(adoCol.Value) Then .内示数量28   = adoCol.Value
                Case "内示数量29":    If Not IsNull(adoCol.Value) Then .内示数量29   = adoCol.Value
                Case "内示数量30":    If Not IsNull(adoCol.Value) Then .内示数量30   = adoCol.Value
                Case "内示数量31":    If Not IsNull(adoCol.Value) Then .内示数量31   = adoCol.Value
                Case "展開区分":      If Not IsNull(adoCol.Value) Then .展開区分     = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function SNJRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    SNJRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From S内示テーブル"
        RTxtSql = RTxtSql &" Where 内示年月 = '" & RTrim(SNJ.内示年月) & "'"
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(SNJ.品番) & "'"
        RTxtSql = RTxtSql &" And   納入先 = '" & RTrim(SNJ.納入先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(SNJ.受入) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           SNJRead = True
           Call ItemsFromSNJ(DBNo)
	   SNJ_RDSTS = True
        else
	   SNJ_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call SNJClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub SNJReadNext(Optional DBNo As Integer = 0)
    If Not SNJ_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           SNJ_RDSTS = False
        Else
           Call ItemsFromSNJ(DBNo)
           SNJ_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub SNJClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        SNJ_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( SNJ )
'*----------------------------------------------*'
Public Sub ItemsClearSNJ()
    With SNJ
	.内示年月 = ""
	.品番 = ""
	.納入先 = ""
	.受入 = ""
	.内示数 = 0
	.翌月内示数 = 0
	.翌々月内示数 = 0
	.繰越在庫数 = 0
	.収容数 = 0
	.背番号 = ""
	.サイクル = ""
	.内示数量01 = 0
	.内示数量02 = 0
	.内示数量03 = 0
	.内示数量04 = 0
	.内示数量05 = 0
	.内示数量06 = 0
	.内示数量07 = 0
	.内示数量08 = 0
	.内示数量09 = 0
	.内示数量10 = 0
	.内示数量11 = 0
	.内示数量12 = 0
	.内示数量13 = 0
	.内示数量14 = 0
	.内示数量15 = 0
	.内示数量16 = 0
	.内示数量17 = 0
	.内示数量18 = 0
	.内示数量19 = 0
	.内示数量20 = 0
	.内示数量21 = 0
	.内示数量22 = 0
	.内示数量23 = 0
	.内示数量24 = 0
	.内示数量25 = 0
	.内示数量26 = 0
	.内示数量27 = 0
	.内示数量28 = 0
	.内示数量29 = 0
	.内示数量30 = 0
	.内示数量31 = 0
	.展開区分 = 0
	.作成者 = ""
'	.作成日 = 0
'	.作成端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function SNJInsert(Optional DBNo As Integer) As Boolean
    SNJInsert = True
    With SNJ
        gSL_Select = "Insert Into S内示テーブル  ( "
        gSL_Select = gSL_Select & "内示年月,品番,納入先,"
        gSL_Select = gSL_Select & "受入,内示数,翌月内示数,"
        gSL_Select = gSL_Select & "翌々月内示数,繰越在庫数,収容数,"
        gSL_Select = gSL_Select & "背番号,サイクル,内示数量01,"
        gSL_Select = gSL_Select & "内示数量02,内示数量03,内示数量04,"
        gSL_Select = gSL_Select & "内示数量05,内示数量06,内示数量07,"
        gSL_Select = gSL_Select & "内示数量08,内示数量09,内示数量10,"
        gSL_Select = gSL_Select & "内示数量11,内示数量12,内示数量13,"
        gSL_Select = gSL_Select & "内示数量14,内示数量15,内示数量16,"
        gSL_Select = gSL_Select & "内示数量17,内示数量18,内示数量19,"
        gSL_Select = gSL_Select & "内示数量20,内示数量21,内示数量22,"
        gSL_Select = gSL_Select & "内示数量23,内示数量24,内示数量25,"
        gSL_Select = gSL_Select & "内示数量26,内示数量27,内示数量28,"
        gSL_Select = gSL_Select & "内示数量29,内示数量30,内示数量31,"
        gSL_Select = gSL_Select & "展開区分,"
        gSL_Select = gSL_Select & "作成者,作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.内示年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "" & .内示数 & ","		
        gSL_Select = gSL_Select & "" & .翌月内示数 & ","		
        gSL_Select = gSL_Select & "" & .翌々月内示数 & ","		
        gSL_Select = gSL_Select & "" & .繰越在庫数 & ","		
        gSL_Select = gSL_Select & "" & .収容数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.サイクル) & "',"		
        gSL_Select = gSL_Select & "" & .内示数量01 & ","		
        gSL_Select = gSL_Select & "" & .内示数量02 & ","		
        gSL_Select = gSL_Select & "" & .内示数量03 & ","		
        gSL_Select = gSL_Select & "" & .内示数量04 & ","		
        gSL_Select = gSL_Select & "" & .内示数量05 & ","		
        gSL_Select = gSL_Select & "" & .内示数量06 & ","		
        gSL_Select = gSL_Select & "" & .内示数量07 & ","		
        gSL_Select = gSL_Select & "" & .内示数量08 & ","		
        gSL_Select = gSL_Select & "" & .内示数量09 & ","		
        gSL_Select = gSL_Select & "" & .内示数量10 & ","		
        gSL_Select = gSL_Select & "" & .内示数量11 & ","		
        gSL_Select = gSL_Select & "" & .内示数量12 & ","		
        gSL_Select = gSL_Select & "" & .内示数量13 & ","		
        gSL_Select = gSL_Select & "" & .内示数量14 & ","		
        gSL_Select = gSL_Select & "" & .内示数量15 & ","		
        gSL_Select = gSL_Select & "" & .内示数量16 & ","		
        gSL_Select = gSL_Select & "" & .内示数量17 & ","		
        gSL_Select = gSL_Select & "" & .内示数量18 & ","		
        gSL_Select = gSL_Select & "" & .内示数量19 & ","		
        gSL_Select = gSL_Select & "" & .内示数量20 & ","		
        gSL_Select = gSL_Select & "" & .内示数量21 & ","		
        gSL_Select = gSL_Select & "" & .内示数量22 & ","		
        gSL_Select = gSL_Select & "" & .内示数量23 & ","		
        gSL_Select = gSL_Select & "" & .内示数量24 & ","		
        gSL_Select = gSL_Select & "" & .内示数量25 & ","		
        gSL_Select = gSL_Select & "" & .内示数量26 & ","		
        gSL_Select = gSL_Select & "" & .内示数量27 & ","		
        gSL_Select = gSL_Select & "" & .内示数量28 & ","		
        gSL_Select = gSL_Select & "" & .内示数量29 & ","		
        gSL_Select = gSL_Select & "" & .内示数量30 & ","		
        gSL_Select = gSL_Select & "" & .内示数量31 & ","		
        gSL_Select = gSL_Select & "" & .展開区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "',"		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SNJInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function SNJUpdate(Optional DBNo As Integer) As Boolean
    SNJUpdate = False	
    With SNJ
        gSL_Select = "Update S内示テーブル  Set "
        gSL_Select = gSL_Select & "内示年月 = '" & RTrim(.内示年月) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "内示数 = " & .内示数 & ","
        gSL_Select = gSL_Select & "翌月内示数 = " & .翌月内示数 & ","
        gSL_Select = gSL_Select & "翌々月内示数 = " & .翌々月内示数 & ","
        gSL_Select = gSL_Select & "繰越在庫数 = " & .繰越在庫数 & ","
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "サイクル = '" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "内示数量01 = " & .内示数量01 & ","
        gSL_Select = gSL_Select & "内示数量02 = " & .内示数量02 & ","
        gSL_Select = gSL_Select & "内示数量03 = " & .内示数量03 & ","
        gSL_Select = gSL_Select & "内示数量04 = " & .内示数量04 & ","
        gSL_Select = gSL_Select & "内示数量05 = " & .内示数量05 & ","
        gSL_Select = gSL_Select & "内示数量06 = " & .内示数量06 & ","
        gSL_Select = gSL_Select & "内示数量07 = " & .内示数量07 & ","
        gSL_Select = gSL_Select & "内示数量08 = " & .内示数量08 & ","
        gSL_Select = gSL_Select & "内示数量09 = " & .内示数量09 & ","
        gSL_Select = gSL_Select & "内示数量10 = " & .内示数量10 & ","
        gSL_Select = gSL_Select & "内示数量11 = " & .内示数量11 & ","
        gSL_Select = gSL_Select & "内示数量12 = " & .内示数量12 & ","
        gSL_Select = gSL_Select & "内示数量13 = " & .内示数量13 & ","
        gSL_Select = gSL_Select & "内示数量14 = " & .内示数量14 & ","
        gSL_Select = gSL_Select & "内示数量15 = " & .内示数量15 & ","
        gSL_Select = gSL_Select & "内示数量16 = " & .内示数量16 & ","
        gSL_Select = gSL_Select & "内示数量17 = " & .内示数量17 & ","
        gSL_Select = gSL_Select & "内示数量18 = " & .内示数量18 & ","
        gSL_Select = gSL_Select & "内示数量19 = " & .内示数量19 & ","
        gSL_Select = gSL_Select & "内示数量20 = " & .内示数量20 & ","
        gSL_Select = gSL_Select & "内示数量21 = " & .内示数量21 & ","
        gSL_Select = gSL_Select & "内示数量22 = " & .内示数量22 & ","
        gSL_Select = gSL_Select & "内示数量23 = " & .内示数量23 & ","
        gSL_Select = gSL_Select & "内示数量24 = " & .内示数量24 & ","
        gSL_Select = gSL_Select & "内示数量25 = " & .内示数量25 & ","
        gSL_Select = gSL_Select & "内示数量26 = " & .内示数量26 & ","
        gSL_Select = gSL_Select & "内示数量27 = " & .内示数量27 & ","
        gSL_Select = gSL_Select & "内示数量28 = " & .内示数量28 & ","
        gSL_Select = gSL_Select & "内示数量29 = " & .内示数量29 & ","
        gSL_Select = gSL_Select & "内示数量30 = " & .内示数量30 & ","
        gSL_Select = gSL_Select & "内示数量31 = " & .内示数量31 & ","
        gSL_Select = gSL_Select & "展開区分 = " & .展開区分 & ""
        gSL_Select = gSL_Select & " Where 内示年月 = '" & RTrim(.内示年月) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 納入先 = '" & RTrim(.納入先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SNJUpdate = True
End Function

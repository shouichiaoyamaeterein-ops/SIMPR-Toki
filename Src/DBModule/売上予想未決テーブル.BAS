Attribute VB_Name = "売上予想未決テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   売上予想未決テーブル
'*
'*   モジュールID  : 売上予想未決テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/11/04  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_売上予想未決テーブル
        内示年月  As String
        取引先    As String
        受入      As String
        品番      As String
        車型      As String
        管理番号  As String
        号補区分  As String
        ライン    As String
        工区      As String
        数量1     As Long
        数量2     As Long
        数量3     As Long
        要因      As String
        作成時間  As String
'        作成端末  As String
'        作成日    As String
        件数          As Long
'   項目追加(自動)
	表示品番      As String
	納入先名      As String
	契約先名      As String
	契約先        As String
    End Type

    Public UYM As T_売上予想未決テーブル

    Public UYM_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromUYM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearUYM	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With UYM
            Select Case adoCol.Name
                Case "内示年月":  If Not IsNull(adoCol.Value) Then .内示年月 = adoCol.Value
                Case "取引先":    If Not IsNull(adoCol.Value) Then .取引先   = adoCol.Value
                Case "受入":      If Not IsNull(adoCol.Value) Then .受入     = adoCol.Value
                Case "品番":      If Not IsNull(adoCol.Value) Then .品番     = adoCol.Value
                Case "車型":      If Not IsNull(adoCol.Value) Then .車型     = adoCol.Value
                Case "管理番号":  If Not IsNull(adoCol.Value) Then .管理番号 = adoCol.Value
                Case "号補区分":  If Not IsNull(adoCol.Value) Then .号補区分 = adoCol.Value
                Case "ライン":    If Not IsNull(adoCol.Value) Then .ライン   = adoCol.Value
                Case "工区":      If Not IsNull(adoCol.Value) Then .工区     = adoCol.Value
                Case "数量1":     If Not IsNull(adoCol.Value) Then .数量1    = adoCol.Value
                Case "数量2":     If Not IsNull(adoCol.Value) Then .数量2    = adoCol.Value
                Case "数量3":     If Not IsNull(adoCol.Value) Then .数量3    = adoCol.Value
                Case "要因":      If Not IsNull(adoCol.Value) Then .要因     = adoCol.Value
                Case "作成時間":  If Not IsNull(adoCol.Value) Then .作成時間 = adoCol.Value
'                Case "作成端末":  If Not IsNull(adoCol.Value) Then .作成端末 = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "納入先名":	If Not IsNull(adoCol.Value) Then .納入先名	= adoCol.Value
		Case "契約先名":	If Not IsNull(adoCol.Value) Then .契約先名	= adoCol.Value
		Case "契約先":	If Not IsNull(adoCol.Value) Then .契約先	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function UYMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    UYMRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 売上予想未決テーブル"
        RTxtSql = RTxtSql &" Where 内示年月 = '" & RTrim(UYM.内示年月) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           UYMRead = True
           Call ItemsFromUYM(DBNo)
	   UYM_RDSTS = True
        else
	   UYM_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call UYMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub UYMReadNext(Optional DBNo As Integer = 0)
    If Not UYM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           UYM_RDSTS = False
        Else
           Call ItemsFromUYM(DBNo)
           UYM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub UYMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        UYM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( UYM )
'*----------------------------------------------*'
Public Sub ItemsClearUYM()
    With UYM
	.内示年月 = ""
	.取引先 = ""
	.受入 = ""
	.品番 = ""
	.車型 = ""
	.管理番号 = ""
	.号補区分 = ""
	.ライン = ""
	.工区 = ""
	.数量1 = 0
	.数量2 = 0
	.数量3 = 0
	.要因 = ""
	.作成時間 = ""
'	.作成端末 = ""
'	.作成日 = ""
        .件数 = 0
'   項目追加(自動)
	.表示品番 = ""
	.納入先名 = ""
	.契約先名 = ""
	.契約先 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function UYMInsert(Optional DBNo As Integer) As Boolean
    UYMInsert = True
    With UYM
        gSL_Select = "Insert Into 売上予想未決テーブル  ( "
        gSL_Select = gSL_Select & "内示年月,取引先,受入,"
        gSL_Select = gSL_Select & "品番,車型,管理番号,"
        gSL_Select = gSL_Select & "号補区分,ライン,工区,"
        gSL_Select = gSL_Select & "数量1,数量2,数量3,"
        gSL_Select = gSL_Select & "要因,作成時間,"
        gSL_Select = gSL_Select & "作成端末,作成日"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.内示年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.車型) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.管理番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.号補区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ライン) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.工区) & "',"		
        gSL_Select = gSL_Select & "" & .数量1 & ","		
        gSL_Select = gSL_Select & "" & .数量2 & ","		
        gSL_Select = gSL_Select & "" & .数量3 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.要因) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.作成時間) & "',"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,15),"		
        gSL_Select = gSL_Select & "GETDATE()"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    UYMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function UYMUpdate(Optional DBNo As Integer) As Boolean
    UYMUpdate = False	
    With UYM
        gSL_Select = "Update 売上予想未決テーブル  Set "
        gSL_Select = gSL_Select & "内示年月 = '" & RTrim(.内示年月) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "車型 = '" & RTrim(.車型) & "',"
        gSL_Select = gSL_Select & "管理番号 = '" & RTrim(.管理番号) & "',"
        gSL_Select = gSL_Select & "号補区分 = '" & RTrim(.号補区分) & "',"
        gSL_Select = gSL_Select & "ライン = '" & RTrim(.ライン) & "',"
        gSL_Select = gSL_Select & "工区 = '" & RTrim(.工区) & "',"
        gSL_Select = gSL_Select & "数量1 = " & .数量1 & ","
        gSL_Select = gSL_Select & "数量2 = " & .数量2 & ","
        gSL_Select = gSL_Select & "数量3 = " & .数量3 & ","
        gSL_Select = gSL_Select & "要因 = '" & RTrim(.要因) & "',"
        gSL_Select = gSL_Select & "作成時間 = '" & RTrim(.作成時間) & "'"
        gSL_Select = gSL_Select & " Where 内示年月 = '" & RTrim(.内示年月) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    UYMUpdate = True
End Function

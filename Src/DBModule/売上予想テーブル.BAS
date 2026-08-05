Attribute VB_Name = "売上予想テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   売上予想テーブル
'*
'*   モジュールID  : 売上予想テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/11/04  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_売上予想テーブル
        内示年月  As String
        契約先    As String
        品番      As String
        工区      As String
        車型      As String
        号補区分  As String
        ライン    As String
        連番      As Byte
        単価1     As Currency
        仮区分1   As String
        数量1     As Long
        金額1     As Currency
        前月比1   As Currency
        単価2     As Currency
        仮区分2   As String
        数量2     As Long
        金額2     As Currency
        前月比2   As Currency
        単価3     As Currency
        仮区分3   As String
        数量3     As Long
        金額3     As Currency
        前月比3   As Currency
        集約品番  As String
        品番編集  As String
'        作成日    As Date
'        作成端末  As String
'        更新日    As Date
'        更新端末  As String
        件数          As Long
'   項目追加(自動)
	契約先名      As String
    End Type

    Public UYT As T_売上予想テーブル

    Public UYT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromUYT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearUYT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With UYT
            Select Case adoCol.Name
                Case "内示年月":  If Not IsNull(adoCol.Value) Then .内示年月 = adoCol.Value
                Case "契約先":    If Not IsNull(adoCol.Value) Then .契約先   = adoCol.Value
                Case "品番":      If Not IsNull(adoCol.Value) Then .品番     = adoCol.Value
                Case "工区":      If Not IsNull(adoCol.Value) Then .工区     = adoCol.Value
                Case "車型":      If Not IsNull(adoCol.Value) Then .車型     = adoCol.Value
                Case "号補区分":  If Not IsNull(adoCol.Value) Then .号補区分 = adoCol.Value
                Case "ライン":    If Not IsNull(adoCol.Value) Then .ライン   = adoCol.Value
                Case "連番":      If Not IsNull(adoCol.Value) Then .連番     = adoCol.Value
                Case "単価1":     If Not IsNull(adoCol.Value) Then .単価1    = adoCol.Value
                Case "仮区分1":   If Not IsNull(adoCol.Value) Then .仮区分1  = adoCol.Value
                Case "数量1":     If Not IsNull(adoCol.Value) Then .数量1    = adoCol.Value
                Case "金額1":     If Not IsNull(adoCol.Value) Then .金額1    = adoCol.Value
                Case "前月比1":   If Not IsNull(adoCol.Value) Then .前月比1  = adoCol.Value
                Case "単価2":     If Not IsNull(adoCol.Value) Then .単価2    = adoCol.Value
                Case "仮区分2":   If Not IsNull(adoCol.Value) Then .仮区分2  = adoCol.Value
                Case "数量2":     If Not IsNull(adoCol.Value) Then .数量2    = adoCol.Value
                Case "金額2":     If Not IsNull(adoCol.Value) Then .金額2    = adoCol.Value
                Case "前月比2":   If Not IsNull(adoCol.Value) Then .前月比2  = adoCol.Value
                Case "単価3":     If Not IsNull(adoCol.Value) Then .単価3    = adoCol.Value
                Case "仮区分3":   If Not IsNull(adoCol.Value) Then .仮区分3  = adoCol.Value
                Case "数量3":     If Not IsNull(adoCol.Value) Then .数量3    = adoCol.Value
                Case "金額3":     If Not IsNull(adoCol.Value) Then .金額3    = adoCol.Value
                Case "前月比3":   If Not IsNull(adoCol.Value) Then .前月比3  = adoCol.Value
                Case "集約品番":  If Not IsNull(adoCol.Value) Then .集約品番 = adoCol.Value
                Case "品番編集":  If Not IsNull(adoCol.Value) Then .品番編集 = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
'                Case "作成端末":  If Not IsNull(adoCol.Value) Then .作成端末 = adoCol.Value
'                Case "更新日":    If Not IsNull(adoCol.Value) Then .更新日   = adoCol.Value
'                Case "更新端末":  If Not IsNull(adoCol.Value) Then .更新端末 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "契約先名":	If Not IsNull(adoCol.Value) Then .契約先名	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function UYTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    UYTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 売上予想テーブル"
        RTxtSql = RTxtSql &" Where 内示年月 = '" & RTrim(UYT.内示年月) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           UYTRead = True
           Call ItemsFromUYT(DBNo)
	   UYT_RDSTS = True
        else
	   UYT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call UYTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub UYTReadNext(Optional DBNo As Integer = 0)
    If Not UYT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           UYT_RDSTS = False
        Else
           Call ItemsFromUYT(DBNo)
           UYT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub UYTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        UYT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( UYT )
'*----------------------------------------------*'
Public Sub ItemsClearUYT()
    With UYT
	.内示年月 = ""
	.契約先 = ""
	.品番 = ""
	.工区 = ""
	.車型 = ""
	.号補区分 = ""
	.ライン = ""
	.連番 = 0
	.単価1 = 0
	.仮区分1 = ""
	.数量1 = 0
	.金額1 = 0
	.前月比1 = 0
	.単価2 = 0
	.仮区分2 = ""
	.数量2 = 0
	.金額2 = 0
	.前月比2 = 0
	.単価3 = 0
	.仮区分3 = ""
	.数量3 = 0
	.金額3 = 0
	.前月比3 = 0
	.集約品番 = ""
	.品番編集 = ""
'	.作成日 = 0
'	.作成端末 = ""
'	.更新日 = 0
'	.更新端末 = ""
        .件数 = 0
'   項目追加(自動)
	.契約先名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function UYTInsert(Optional DBNo As Integer) As Boolean
    UYTInsert = True
    With UYT
        gSL_Select = "Insert Into 売上予想テーブル  ( "
        gSL_Select = gSL_Select & "内示年月,契約先,品番,"
        gSL_Select = gSL_Select & "工区,車型,号補区分,"
        gSL_Select = gSL_Select & "ライン,連番,単価1,"
        gSL_Select = gSL_Select & "仮区分1,数量1,金額1,"
        gSL_Select = gSL_Select & "前月比1,単価2,仮区分2,"
        gSL_Select = gSL_Select & "数量2,金額2,前月比2,"
        gSL_Select = gSL_Select & "単価3,仮区分3,数量3,"
        gSL_Select = gSL_Select & "金額3,前月比3,集約品番,"
        gSL_Select = gSL_Select & "品番編集,"
        gSL_Select = gSL_Select & "作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.内示年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.工区) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.車型) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.号補区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ライン) & "',"		
        gSL_Select = gSL_Select & "" & .連番 & ","		
        gSL_Select = gSL_Select & "" & .単価1 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.仮区分1) & "',"		
        gSL_Select = gSL_Select & "" & .数量1 & ","		
        gSL_Select = gSL_Select & "" & .金額1 & ","		
        gSL_Select = gSL_Select & "" & .前月比1 & ","		
        gSL_Select = gSL_Select & "" & .単価2 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.仮区分2) & "',"		
        gSL_Select = gSL_Select & "" & .数量2 & ","		
        gSL_Select = gSL_Select & "" & .金額2 & ","		
        gSL_Select = gSL_Select & "" & .前月比2 & ","		
        gSL_Select = gSL_Select & "" & .単価3 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.仮区分3) & "',"		
        gSL_Select = gSL_Select & "" & .数量3 & ","		
        gSL_Select = gSL_Select & "" & .金額3 & ","		
        gSL_Select = gSL_Select & "" & .前月比3 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.集約品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番編集) & "',"		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,15)"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    UYTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function UYTUpdate(Optional DBNo As Integer) As Boolean
    UYTUpdate = False	
    With UYT
        gSL_Select = "Update 売上予想テーブル  Set "
        gSL_Select = gSL_Select & "内示年月 = '" & RTrim(.内示年月) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "工区 = '" & RTrim(.工区) & "',"
        gSL_Select = gSL_Select & "車型 = '" & RTrim(.車型) & "',"
        gSL_Select = gSL_Select & "号補区分 = '" & RTrim(.号補区分) & "',"
        gSL_Select = gSL_Select & "ライン = '" & RTrim(.ライン) & "',"
        gSL_Select = gSL_Select & "連番 = " & .連番 & ","
        gSL_Select = gSL_Select & "単価1 = " & .単価1 & ","
        gSL_Select = gSL_Select & "仮区分1 = '" & RTrim(.仮区分1) & "',"
        gSL_Select = gSL_Select & "数量1 = " & .数量1 & ","
        gSL_Select = gSL_Select & "金額1 = " & .金額1 & ","
        gSL_Select = gSL_Select & "前月比1 = " & .前月比1 & ","
        gSL_Select = gSL_Select & "単価2 = " & .単価2 & ","
        gSL_Select = gSL_Select & "仮区分2 = '" & RTrim(.仮区分2) & "',"
        gSL_Select = gSL_Select & "数量2 = " & .数量2 & ","
        gSL_Select = gSL_Select & "金額2 = " & .金額2 & ","
        gSL_Select = gSL_Select & "前月比2 = " & .前月比2 & ","
        gSL_Select = gSL_Select & "単価3 = " & .単価3 & ","
        gSL_Select = gSL_Select & "仮区分3 = '" & RTrim(.仮区分3) & "',"
        gSL_Select = gSL_Select & "数量3 = " & .数量3 & ","
        gSL_Select = gSL_Select & "金額3 = " & .金額3 & ","
        gSL_Select = gSL_Select & "前月比3 = " & .前月比3 & ","
        gSL_Select = gSL_Select & "集約品番 = '" & RTrim(.集約品番) & "',"
        gSL_Select = gSL_Select & "品番編集 = '" & RTrim(.品番編集) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,15)"
        gSL_Select = gSL_Select & " Where 内示年月 = '" & RTrim(.内示年月) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    UYTUpdate = True
End Function

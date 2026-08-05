Attribute VB_Name = "S発注内示テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   S発注内示テーブル
'*
'*   モジュールID  : S発注内示テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/05/10  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_S発注内示テーブル
        内示年月          As String
        品番              As String
        手配先            As String
        受入              As String
        直送先            As String
        納期              As String
        仕掛日            As String
        手配日            As String
        数量              As Long
        処理区分          As Byte
        削除区分          As Byte
        変更回数          As Byte
        ローレベル        As Byte
        計画年月          As String
        作成者            As String
'        作成日            As Date
'        作成端末          As String
        更新者            As String
'        更新日            As Date
'        更新端末          As String
        件数          As Long
    End Type

    Public SHN As T_S発注内示テーブル

    Public SHN_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromSHN(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearSHN	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With SHN
            Select Case adoCol.Name
                Case "内示年月":          If Not IsNull(adoCol.Value) Then .内示年月         = adoCol.Value
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番             = adoCol.Value
                Case "手配先":            If Not IsNull(adoCol.Value) Then .手配先           = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入             = adoCol.Value
                Case "直送先":            If Not IsNull(adoCol.Value) Then .直送先           = adoCol.Value
                Case "納期":              If Not IsNull(adoCol.Value) Then .納期             = adoCol.Value
                Case "仕掛日":            If Not IsNull(adoCol.Value) Then .仕掛日           = adoCol.Value
                Case "手配日":            If Not IsNull(adoCol.Value) Then .手配日           = adoCol.Value
                Case "数量":              If Not IsNull(adoCol.Value) Then .数量             = adoCol.Value
                Case "処理区分":          If Not IsNull(adoCol.Value) Then .処理区分         = adoCol.Value
                Case "削除区分":          If Not IsNull(adoCol.Value) Then .削除区分         = adoCol.Value
                Case "変更回数":          If Not IsNull(adoCol.Value) Then .変更回数         = adoCol.Value
                Case "ローレベル":        If Not IsNull(adoCol.Value) Then .ローレベル       = adoCol.Value
                Case "計画年月":          If Not IsNull(adoCol.Value) Then .計画年月         = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者           = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
                Case "更新者":            If Not IsNull(adoCol.Value) Then .更新者           = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function SHNRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    SHNRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From S発注内示テーブル"
        RTxtSql = RTxtSql &" Where 内示年月 = '" & RTrim(SHN.内示年月) & "'"
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(SHN.品番) & "'"
        RTxtSql = RTxtSql &" And   手配先 = '" & RTrim(SHN.手配先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(SHN.受入) & "'"
        RTxtSql = RTxtSql &" And   直送先 = '" & RTrim(SHN.直送先) & "'"
        RTxtSql = RTxtSql &" And   納期 = '" & RTrim(SHN.納期) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           SHNRead = True
           Call ItemsFromSHN(DBNo)
	   SHN_RDSTS = True
        else
	   SHN_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call SHNClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub SHNReadNext(Optional DBNo As Integer = 0)
    If Not SHN_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           SHN_RDSTS = False
        Else
           Call ItemsFromSHN(DBNo)
           SHN_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub SHNClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        SHN_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( SHN )
'*----------------------------------------------*'
Public Sub ItemsClearSHN()
    With SHN
	.内示年月 = ""
	.品番 = ""
	.手配先 = ""
	.受入 = ""
	.直送先 = ""
	.納期 = ""
	.仕掛日 = ""
	.手配日 = ""
	.数量 = 0
	.処理区分 = 0
	.削除区分 = 0
	.変更回数 = 0
	.ローレベル = 0
	.計画年月 = ""
	.作成者 = ""
'	.作成日 = 0
'	.作成端末 = ""
	.更新者 = ""
'	.更新日 = 0
'	.更新端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function SHNInsert(Optional DBNo As Integer) As Boolean
    SHNInsert = True
    With SHN
        gSL_Select = "Insert Into S発注内示テーブル  ( "
        gSL_Select = gSL_Select & "内示年月,品番,手配先,"
        gSL_Select = gSL_Select & "受入,直送先,納期,"
        gSL_Select = gSL_Select & "仕掛日,手配日,数量,"
        gSL_Select = gSL_Select & "処理区分,削除区分,変更回数,"
        gSL_Select = gSL_Select & "ローレベル,計画年月,"
        gSL_Select = gSL_Select & "作成者,作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.内示年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.直送先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納期) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕掛日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配日) & "',"		
        gSL_Select = gSL_Select & "" & .数量 & ","		
        gSL_Select = gSL_Select & "" & .処理区分 & ","		
        gSL_Select = gSL_Select & "" & .削除区分 & ","		
        gSL_Select = gSL_Select & "" & .変更回数 & ","		
        gSL_Select = gSL_Select & "" & .ローレベル & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.計画年月) & "',"		
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
    SHNInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function SHNUpdate(Optional DBNo As Integer) As Boolean
    SHNUpdate = False	
    With SHN
        gSL_Select = "Update S発注内示テーブル  Set "
        gSL_Select = gSL_Select & "内示年月 = '" & RTrim(.内示年月) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "手配先 = '" & RTrim(.手配先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "直送先 = '" & RTrim(.直送先) & "',"
        gSL_Select = gSL_Select & "納期 = '" & RTrim(.納期) & "',"
        gSL_Select = gSL_Select & "仕掛日 = '" & RTrim(.仕掛日) & "',"
        gSL_Select = gSL_Select & "手配日 = '" & RTrim(.手配日) & "',"
        gSL_Select = gSL_Select & "数量 = " & .数量 & ","
        gSL_Select = gSL_Select & "処理区分 = " & .処理区分 & ","
        gSL_Select = gSL_Select & "削除区分 = " & .削除区分 & ","
        gSL_Select = gSL_Select & "変更回数 = " & .変更回数 & ","
        gSL_Select = gSL_Select & "ローレベル = " & .ローレベル & ","
        gSL_Select = gSL_Select & "計画年月 = '" & RTrim(.計画年月) & "',"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 内示年月 = '" & RTrim(.内示年月) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 手配先 = '" & RTrim(.手配先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 直送先 = '" & RTrim(.直送先) & "'"
        gSL_Select = gSL_Select & "   And 納期 = '" & RTrim(.納期) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SHNUpdate = True
End Function

Attribute VB_Name = "S発注内示構成テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   S発注内示構成テーブル
'*
'*   モジュールID  : S発注内示構成テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/05/10  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_S発注内示構成テーブル
        処理端末  As String
        順位      As String
        レベル    As Byte
        親品番    As String
        品番      As String
        手配先    As String
        受入      As String
        直送先    As String
        親納期    As String
        納期      As String
        仕掛日    As String
        手配日    As String
        数量      As Currency
        構成数    As Currency
        手配区分  As String
        展開区分  As Byte
'        作成日    As Date
        作成者    As String
'        作成端末  As String
        件数          As Long
    End Type

    Public SHK As T_S発注内示構成テーブル

    Public SHK_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromSHK(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearSHK	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With SHK
            Select Case adoCol.Name
                Case "処理端末":  If Not IsNull(adoCol.Value) Then .処理端末 = adoCol.Value
                Case "順位":      If Not IsNull(adoCol.Value) Then .順位     = adoCol.Value
                Case "レベル":    If Not IsNull(adoCol.Value) Then .レベル   = adoCol.Value
                Case "親品番":    If Not IsNull(adoCol.Value) Then .親品番   = adoCol.Value
                Case "品番":      If Not IsNull(adoCol.Value) Then .品番     = adoCol.Value
                Case "手配先":    If Not IsNull(adoCol.Value) Then .手配先   = adoCol.Value
                Case "受入":      If Not IsNull(adoCol.Value) Then .受入     = adoCol.Value
                Case "直送先":    If Not IsNull(adoCol.Value) Then .直送先   = adoCol.Value
                Case "親納期":    If Not IsNull(adoCol.Value) Then .親納期   = adoCol.Value
                Case "納期":      If Not IsNull(adoCol.Value) Then .納期     = adoCol.Value
                Case "仕掛日":    If Not IsNull(adoCol.Value) Then .仕掛日   = adoCol.Value
                Case "手配日":    If Not IsNull(adoCol.Value) Then .手配日   = adoCol.Value
                Case "数量":      If Not IsNull(adoCol.Value) Then .数量     = adoCol.Value
                Case "構成数":    If Not IsNull(adoCol.Value) Then .構成数   = adoCol.Value
                Case "手配区分":  If Not IsNull(adoCol.Value) Then .手配区分 = adoCol.Value
                Case "展開区分":  If Not IsNull(adoCol.Value) Then .展開区分 = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
                Case "作成者":    If Not IsNull(adoCol.Value) Then .作成者   = adoCol.Value
'                Case "作成端末":  If Not IsNull(adoCol.Value) Then .作成端末 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function SHKRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    SHKRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From S発注内示構成テーブル"
        RTxtSql = RTxtSql &" Where 処理端末 = '" & RTrim(SHK.処理端末) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           SHKRead = True
           Call ItemsFromSHK(DBNo)
	   SHK_RDSTS = True
        else
	   SHK_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call SHKClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub SHKReadNext(Optional DBNo As Integer = 0)
    If Not SHK_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           SHK_RDSTS = False
        Else
           Call ItemsFromSHK(DBNo)
           SHK_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub SHKClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        SHK_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( SHK )
'*----------------------------------------------*'
Public Sub ItemsClearSHK()
    With SHK
	.処理端末 = ""
	.順位 = ""
	.レベル = 0
	.親品番 = ""
	.品番 = ""
	.手配先 = ""
	.受入 = ""
	.直送先 = ""
	.親納期 = ""
	.納期 = ""
	.仕掛日 = ""
	.手配日 = ""
	.数量 = 0
	.構成数 = 0
	.手配区分 = ""
	.展開区分 = 0
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function SHKInsert(Optional DBNo As Integer) As Boolean
    SHKInsert = True
    With SHK
        gSL_Select = "Insert Into S発注内示構成テーブル  ( "
        gSL_Select = gSL_Select & "処理端末,順位,レベル,"
        gSL_Select = gSL_Select & "親品番,品番,手配先,"
        gSL_Select = gSL_Select & "受入,直送先,親納期,"
        gSL_Select = gSL_Select & "納期,仕掛日,手配日,"
        gSL_Select = gSL_Select & "数量,構成数,手配区分,"
        gSL_Select = gSL_Select & "展開区分,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.順位) & "',"		
        gSL_Select = gSL_Select & "" & .レベル & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.親品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.直送先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.親納期) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納期) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕掛日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配日) & "',"		
        gSL_Select = gSL_Select & "" & .数量 & ","		
        gSL_Select = gSL_Select & "" & .構成数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.手配区分) & "',"		
        gSL_Select = gSL_Select & "" & .展開区分 & ","		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "',"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SHKInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function SHKUpdate(Optional DBNo As Integer) As Boolean
    SHKUpdate = False	
    With SHK
        gSL_Select = "Update S発注内示構成テーブル  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "順位 = '" & RTrim(.順位) & "',"
        gSL_Select = gSL_Select & "レベル = " & .レベル & ","
        gSL_Select = gSL_Select & "親品番 = '" & RTrim(.親品番) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "手配先 = '" & RTrim(.手配先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "直送先 = '" & RTrim(.直送先) & "',"
        gSL_Select = gSL_Select & "親納期 = '" & RTrim(.親納期) & "',"
        gSL_Select = gSL_Select & "納期 = '" & RTrim(.納期) & "',"
        gSL_Select = gSL_Select & "仕掛日 = '" & RTrim(.仕掛日) & "',"
        gSL_Select = gSL_Select & "手配日 = '" & RTrim(.手配日) & "',"
        gSL_Select = gSL_Select & "数量 = " & .数量 & ","
        gSL_Select = gSL_Select & "構成数 = " & .構成数 & ","
        gSL_Select = gSL_Select & "手配区分 = '" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "展開区分 = " & .展開区分 & ""
        gSL_Select = gSL_Select & " Where 処理端末 = '" & RTrim(.処理端末) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SHKUpdate = True
End Function

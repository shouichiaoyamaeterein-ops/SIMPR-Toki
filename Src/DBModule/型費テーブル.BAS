Attribute VB_Name = "型費テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   型費テーブル
'*
'*   モジュールID  : 型費テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/12/07  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_型費テーブル
        取引先        As String
        品番          As String
        連番          As Byte
        対象開始年月  As String
        適用開始年月  As String
        適用終了年月  As String
        月発生金額    As Long
        投資金額      As Long
        回数          As Byte
        調整金額      As Long
        切替年月      As String
        金額          As Long
        照合年月      As String
        照合処理日    As String
        照合結果      As String
        実績年月      As String
        実績処理日    As String
        実績結果      As String
        予想年月      As String
        予想処理日    As String
'        作成日        As Date
        作成者        As String
'        作成端末      As String
'        更新日        As Date
        更新者        As String
'        更新端末      As String
        件数          As Long
    End Type

    Public KTT As T_型費テーブル

    Public KTT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKTT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKTT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KTT
            Select Case adoCol.Name
                Case "取引先":        If Not IsNull(adoCol.Value) Then .取引先       = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "連番":          If Not IsNull(adoCol.Value) Then .連番         = adoCol.Value
                Case "対象開始年月":  If Not IsNull(adoCol.Value) Then .対象開始年月 = adoCol.Value
                Case "適用開始年月":  If Not IsNull(adoCol.Value) Then .適用開始年月 = adoCol.Value
                Case "適用終了年月":  If Not IsNull(adoCol.Value) Then .適用終了年月 = adoCol.Value
                Case "月発生金額":    If Not IsNull(adoCol.Value) Then .月発生金額   = adoCol.Value
                Case "投資金額":      If Not IsNull(adoCol.Value) Then .投資金額     = adoCol.Value
                Case "回数":          If Not IsNull(adoCol.Value) Then .回数         = adoCol.Value
                Case "調整金額":      If Not IsNull(adoCol.Value) Then .調整金額     = adoCol.Value
                Case "切替年月":      If Not IsNull(adoCol.Value) Then .切替年月     = adoCol.Value
                Case "金額":          If Not IsNull(adoCol.Value) Then .金額         = adoCol.Value
                Case "照合年月":      If Not IsNull(adoCol.Value) Then .照合年月     = adoCol.Value
                Case "照合処理日":    If Not IsNull(adoCol.Value) Then .照合処理日   = adoCol.Value
                Case "照合結果":      If Not IsNull(adoCol.Value) Then .照合結果     = adoCol.Value
                Case "実績年月":      If Not IsNull(adoCol.Value) Then .実績年月     = adoCol.Value
                Case "実績処理日":    If Not IsNull(adoCol.Value) Then .実績処理日   = adoCol.Value
                Case "実績結果":      If Not IsNull(adoCol.Value) Then .実績結果     = adoCol.Value
                Case "予想年月":      If Not IsNull(adoCol.Value) Then .予想年月     = adoCol.Value
                Case "予想処理日":    If Not IsNull(adoCol.Value) Then .予想処理日   = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者       = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KTTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KTTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 型費テーブル"
        RTxtSql = RTxtSql &" Where 取引先 = '" & RTrim(KTT.取引先) & "'"
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(KTT.品番) & "'"
        RTxtSql = RTxtSql &" And   連番 = " & KTT.連番 & ""
        RTxtSql = RTxtSql &" And   対象開始年月 = '" & RTrim(KTT.対象開始年月) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KTTRead = True
           Call ItemsFromKTT(DBNo)
	   KTT_RDSTS = True
        else
	   KTT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KTTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KTTReadNext(Optional DBNo As Integer = 0)
    If Not KTT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KTT_RDSTS = False
        Else
           Call ItemsFromKTT(DBNo)
           KTT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KTTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KTT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KTT )
'*----------------------------------------------*'
Public Sub ItemsClearKTT()
    With KTT
	.取引先 = ""
	.品番 = ""
	.連番 = 0
	.対象開始年月 = ""
	.適用開始年月 = ""
	.適用終了年月 = ""
	.月発生金額 = 0
	.投資金額 = 0
	.回数 = 0
	.調整金額 = 0
	.切替年月 = ""
	.金額 = 0
	.照合年月 = ""
	.照合処理日 = ""
	.照合結果 = ""
	.実績年月 = ""
	.実績処理日 = ""
	.実績結果 = ""
	.予想年月 = ""
	.予想処理日 = ""
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
'	.更新日 = 0
	.更新者 = ""
'	.更新端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KTTInsert(Optional DBNo As Integer) As Boolean
    KTTInsert = True
    With KTT
        gSL_Select = "Insert Into 型費テーブル  ( "
        gSL_Select = gSL_Select & "取引先,品番,連番,"
        gSL_Select = gSL_Select & "対象開始年月,適用開始年月,適用終了年月,"
        gSL_Select = gSL_Select & "月発生金額,投資金額,回数,"
        gSL_Select = gSL_Select & "調整金額,切替年月,金額,"
        gSL_Select = gSL_Select & "照合年月,照合処理日,照合結果,"
        gSL_Select = gSL_Select & "実績年月,実績処理日,実績結果,"
        gSL_Select = gSL_Select & "予想年月,予想処理日,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "" & .連番 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.対象開始年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.適用開始年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.適用終了年月) & "',"		
        gSL_Select = gSL_Select & "" & .月発生金額 & ","		
        gSL_Select = gSL_Select & "" & .投資金額 & ","		
        gSL_Select = gSL_Select & "" & .回数 & ","		
        gSL_Select = gSL_Select & "" & .調整金額 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.切替年月) & "',"		
        gSL_Select = gSL_Select & "" & .金額 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.照合年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合処理日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合結果) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.実績年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.実績処理日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.実績結果) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.予想年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.予想処理日) & "',"		
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
    KTTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KTTUpdate(Optional DBNo As Integer) As Boolean
    KTTUpdate = False	
    With KTT
        gSL_Select = "Update 型費テーブル  Set "
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "連番 = " & .連番 & ","
        gSL_Select = gSL_Select & "対象開始年月 = '" & RTrim(.対象開始年月) & "',"
        gSL_Select = gSL_Select & "適用開始年月 = '" & RTrim(.適用開始年月) & "',"
        gSL_Select = gSL_Select & "適用終了年月 = '" & RTrim(.適用終了年月) & "',"
        gSL_Select = gSL_Select & "月発生金額 = " & .月発生金額 & ","
        gSL_Select = gSL_Select & "投資金額 = " & .投資金額 & ","
        gSL_Select = gSL_Select & "回数 = " & .回数 & ","
        gSL_Select = gSL_Select & "調整金額 = " & .調整金額 & ","
        gSL_Select = gSL_Select & "切替年月 = '" & RTrim(.切替年月) & "',"
        gSL_Select = gSL_Select & "金額 = " & .金額 & ","
        gSL_Select = gSL_Select & "照合年月 = '" & RTrim(.照合年月) & "',"
        gSL_Select = gSL_Select & "照合処理日 = '" & RTrim(.照合処理日) & "',"
        gSL_Select = gSL_Select & "照合結果 = '" & RTrim(.照合結果) & "',"
        gSL_Select = gSL_Select & "実績年月 = '" & RTrim(.実績年月) & "',"
        gSL_Select = gSL_Select & "実績処理日 = '" & RTrim(.実績処理日) & "',"
        gSL_Select = gSL_Select & "実績結果 = '" & RTrim(.実績結果) & "',"
        gSL_Select = gSL_Select & "予想年月 = '" & RTrim(.予想年月) & "',"
        gSL_Select = gSL_Select & "予想処理日 = '" & RTrim(.予想処理日) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 連番 = " & .連番 & ""
        gSL_Select = gSL_Select & "   And 対象開始年月 = '" & RTrim(.対象開始年月) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KTTUpdate = True
End Function

Attribute VB_Name = "買掛型費テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   買掛型費テーブル
'*
'*   モジュールID  : 買掛型費テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/12/07  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_買掛型費テーブル
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
'   項目追加(自動)
	取引先名      As String
    End Type

    Public KTK As T_買掛型費テーブル

    Public KTK_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKTK(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKTK	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KTK
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
'   項目追加(自動)
		Case "取引先名":	If Not IsNull(adoCol.Value) Then .取引先名	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KTKRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KTKRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 買掛型費テーブル"
        RTxtSql = RTxtSql &" Where 取引先 = '" & RTrim(KTK.取引先) & "'"
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(KTK.品番) & "'"
        RTxtSql = RTxtSql &" And   連番 = " & KTK.連番 & ""
        RTxtSql = RTxtSql &" And   対象開始年月 = '" & RTrim(KTK.対象開始年月) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KTKRead = True
           Call ItemsFromKTK(DBNo)
	   KTK_RDSTS = True
        else
	   KTK_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KTKClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KTKReadNext(Optional DBNo As Integer = 0)
    If Not KTK_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KTK_RDSTS = False
        Else
           Call ItemsFromKTK(DBNo)
           KTK_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KTKClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KTK_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KTK )
'*----------------------------------------------*'
Public Sub ItemsClearKTK()
    With KTK
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
'   項目追加(自動)
	.取引先名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KTKInsert(Optional DBNo As Integer) As Boolean
    KTKInsert = True
    With KTK
        gSL_Select = "Insert Into 買掛型費テーブル  ( "
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
    KTKInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KTKUpdate(Optional DBNo As Integer) As Boolean
    KTKUpdate = False	
    With KTK
        gSL_Select = "Update 買掛型費テーブル  Set "
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
    KTKUpdate = True
End Function

Attribute VB_Name = "検収テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   検収テーブル
'*
'*   モジュールID  : 検収テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2005/06/14  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_検収テーブル
        検収管理NO    As Long
        検収管理行NO  As Byte
        指示書NO      As Long
        指示書行NO    As Byte
        データ区分    As Byte
        検収日        As String
        納入先        As String
        受入          As String
        機番          As String
        納入番号      As String
        納品書頁      As Byte
        納品書行      As Byte
        品番          As String
        背番号        As String
        赤黒区分      As Byte
        枚数          As Long
        収容数        As Long
        検収数        As Long
        作成種別      As String
        照合年月      As String
        照合処理日    As String
        照合結果      As String
        実績年月      As String
        実績処理日    As String
        実績結果      As String
'        作成日        As Date
'        作成端末      As String
        作成者        As String
'        更新日        As Date
'        更新端末      As String
        更新者        As String
        件数          As Long
'   項目追加(自動)
	納入先名      As String
	契約先名      As String
	表示品番      As String
	契約先        As String
    End Type

    Public KNT As T_検収テーブル

    Public KNT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKNT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKNT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KNT
            Select Case adoCol.Name
                Case "検収管理NO":    If Not IsNull(adoCol.Value) Then .検収管理NO   = adoCol.Value
                Case "検収管理行NO":  If Not IsNull(adoCol.Value) Then .検収管理行NO = adoCol.Value
                Case "指示書NO":      If Not IsNull(adoCol.Value) Then .指示書NO     = adoCol.Value
                Case "指示書行NO":    If Not IsNull(adoCol.Value) Then .指示書行NO   = adoCol.Value
                Case "データ区分":    If Not IsNull(adoCol.Value) Then .データ区分   = adoCol.Value
                Case "検収日":        If Not IsNull(adoCol.Value) Then .検収日       = adoCol.Value
                Case "納入先":        If Not IsNull(adoCol.Value) Then .納入先       = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "機番":          If Not IsNull(adoCol.Value) Then .機番         = adoCol.Value
                Case "納入番号":      If Not IsNull(adoCol.Value) Then .納入番号     = adoCol.Value
                Case "納品書頁":      If Not IsNull(adoCol.Value) Then .納品書頁     = adoCol.Value
                Case "納品書行":      If Not IsNull(adoCol.Value) Then .納品書行     = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "背番号":        If Not IsNull(adoCol.Value) Then .背番号       = adoCol.Value
                Case "赤黒区分":      If Not IsNull(adoCol.Value) Then .赤黒区分     = adoCol.Value
                Case "枚数":          If Not IsNull(adoCol.Value) Then .枚数         = adoCol.Value
                Case "収容数":        If Not IsNull(adoCol.Value) Then .収容数       = adoCol.Value
                Case "検収数":        If Not IsNull(adoCol.Value) Then .検収数       = adoCol.Value
                Case "作成種別":      If Not IsNull(adoCol.Value) Then .作成種別     = adoCol.Value
                Case "照合年月":      If Not IsNull(adoCol.Value) Then .照合年月     = adoCol.Value
                Case "照合処理日":    If Not IsNull(adoCol.Value) Then .照合処理日   = adoCol.Value
                Case "照合結果":      If Not IsNull(adoCol.Value) Then .照合結果     = adoCol.Value
                Case "実績年月":      If Not IsNull(adoCol.Value) Then .実績年月     = adoCol.Value
                Case "実績処理日":    If Not IsNull(adoCol.Value) Then .実績処理日   = adoCol.Value
                Case "実績結果":      If Not IsNull(adoCol.Value) Then .実績結果     = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者       = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "納入先名":	If Not IsNull(adoCol.Value) Then .納入先名	= adoCol.Value
		Case "契約先名":	If Not IsNull(adoCol.Value) Then .契約先名	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "契約先":	If Not IsNull(adoCol.Value) Then .契約先	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KNTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KNTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 検収テーブル"
        RTxtSql = RTxtSql &" Where 検収管理NO = " & KNT.検収管理NO & ""
        RTxtSql = RTxtSql &" And   検収管理行NO = " & KNT.検収管理行NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KNTRead = True
           Call ItemsFromKNT(DBNo)
	   KNT_RDSTS = True
        else
	   KNT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KNTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KNTReadNext(Optional DBNo As Integer = 0)
    If Not KNT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KNT_RDSTS = False
        Else
           Call ItemsFromKNT(DBNo)
           KNT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KNTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KNT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KNT )
'*----------------------------------------------*'
Public Sub ItemsClearKNT()
    With KNT
	.検収管理NO = 0
	.検収管理行NO = 0
	.指示書NO = 0
	.指示書行NO = 0
	.データ区分 = 0
	.検収日 = ""
	.納入先 = ""
	.受入 = ""
	.機番 = ""
	.納入番号 = ""
	.納品書頁 = 0
	.納品書行 = 0
	.品番 = ""
	.背番号 = ""
	.赤黒区分 = 0
	.枚数 = 0
	.収容数 = 0
	.検収数 = 0
	.作成種別 = ""
	.照合年月 = ""
	.照合処理日 = ""
	.照合結果 = ""
	.実績年月 = ""
	.実績処理日 = ""
	.実績結果 = ""
'	.作成日 = Nothing
'	.作成端末 = ""
	.作成者 = ""
'	.更新日 = Nothing
'	.更新端末 = ""
	.更新者 = ""
        .件数 = 0
'   項目追加(自動)
	.納入先名 = ""
	.契約先名 = ""
	.表示品番 = ""
	.契約先 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KNTInsert(Optional DBNo As Integer) As Boolean
    KNTInsert = True
    With KNT
        gSL_Select = "Insert Into 検収テーブル  ( "
        gSL_Select = gSL_Select & "検収管理NO,検収管理行NO,指示書NO,"
        gSL_Select = gSL_Select & "指示書行NO,データ区分,検収日,"
        gSL_Select = gSL_Select & "納入先,受入,機番,"
        gSL_Select = gSL_Select & "納入番号,納品書頁,納品書行,"
        gSL_Select = gSL_Select & "品番,背番号,赤黒区分,"
        gSL_Select = gSL_Select & "枚数,収容数,検収数,"
        gSL_Select = gSL_Select & "作成種別,照合年月,照合処理日,"
        gSL_Select = gSL_Select & "照合結果,実績年月,実績処理日,"
        gSL_Select = gSL_Select & "実績結果,"
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .検収管理NO & ","		
        gSL_Select = gSL_Select & "" & .検収管理行NO & ","		
        gSL_Select = gSL_Select & "" & .指示書NO & ","		
        gSL_Select = gSL_Select & "" & .指示書行NO & ","		
        gSL_Select = gSL_Select & "" & .データ区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.検収日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.機番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"		
        gSL_Select = gSL_Select & "" & .納品書頁 & ","		
        gSL_Select = gSL_Select & "" & .納品書行 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "" & .赤黒区分 & ","		
        gSL_Select = gSL_Select & "" & .枚数 & ","		
        gSL_Select = gSL_Select & "" & .収容数 & ","		
        gSL_Select = gSL_Select & "" & .検収数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.作成種別) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合処理日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合結果) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.実績年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.実績処理日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.実績結果) & "',"		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20),"		
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KNTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KNTUpdate(Optional DBNo As Integer) As Boolean
    KNTUpdate = False	
    With KNT
        gSL_Select = "Update 検収テーブル  Set "
        gSL_Select = gSL_Select & "検収管理NO = " & .検収管理NO & ","
        gSL_Select = gSL_Select & "検収管理行NO = " & .検収管理行NO & ","
        gSL_Select = gSL_Select & "指示書NO = " & .指示書NO & ","
        gSL_Select = gSL_Select & "指示書行NO = " & .指示書行NO & ","
        gSL_Select = gSL_Select & "データ区分 = " & .データ区分 & ","
        gSL_Select = gSL_Select & "検収日 = '" & RTrim(.検収日) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "機番 = '" & RTrim(.機番) & "',"
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "納品書頁 = " & .納品書頁 & ","
        gSL_Select = gSL_Select & "納品書行 = " & .納品書行 & ","
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "赤黒区分 = " & .赤黒区分 & ","
        gSL_Select = gSL_Select & "枚数 = " & .枚数 & ","
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "検収数 = " & .検収数 & ","
        gSL_Select = gSL_Select & "作成種別 = '" & RTrim(.作成種別) & "',"
        gSL_Select = gSL_Select & "照合年月 = '" & RTrim(.照合年月) & "',"
        gSL_Select = gSL_Select & "照合処理日 = '" & RTrim(.照合処理日) & "',"
        gSL_Select = gSL_Select & "照合結果 = '" & RTrim(.照合結果) & "',"
        gSL_Select = gSL_Select & "実績年月 = '" & RTrim(.実績年月) & "',"
        gSL_Select = gSL_Select & "実績処理日 = '" & RTrim(.実績処理日) & "',"
        gSL_Select = gSL_Select & "実績結果 = '" & RTrim(.実績結果) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "'"
        gSL_Select = gSL_Select & " Where 検収管理NO = " & .検収管理NO & ""
        gSL_Select = gSL_Select & "   And 検収管理行NO = " & .検収管理行NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KNTUpdate = True
End Function

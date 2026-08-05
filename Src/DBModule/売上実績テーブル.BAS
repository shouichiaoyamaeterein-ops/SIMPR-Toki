Attribute VB_Name = "売上実績テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   売上実績テーブル
'*
'*   モジュールID  : 売上実績テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/11/26  By Cis
'*   作成日  :  2008/06/06  By Cis(FMVD5330)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_売上実績テーブル
        検収管理NO            As Long
        検収管理行NO          As Byte
        実績年月              As String
        品番                  As String
        背番号                As String
        納入先                As String
        契約先                As String
        受入                  As String
        車型                  As String
        実績区分              As Byte
        実績処理日            As String
        正式実績年月          As String
        正式実績処理日        As String
        納入番号              As String
        検収年月日            As String
        赤黒区分              As Byte
        単価                  As Currency
        仮単価区分            As String
        納入数                As Long
        適用開始日付          As String
        金額                  As Long
        計算方法              As Byte
        計算単位              As Byte
        締日                  As Byte
        前回実績年月          As String
        明細件数              As Integer
        入力区分              As String
'        作成日                As Date
'        作成端末              As String
        作成者                As String
        件数          As Long
'   項目追加(自動)
	契約先名       As String
	表示品番       As String
	納入先名       As String
	照合対象       As String
	自社工区       As String
	自社工区名     As String
    End Type

    Public UJT As T_売上実績テーブル

    Public UJT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromUJT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearUJT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With UJT
            Select Case adoCol.Name
                Case "検収管理NO":            If Not IsNull(adoCol.Value) Then .検収管理NO           = adoCol.Value
                Case "検収管理行NO":          If Not IsNull(adoCol.Value) Then .検収管理行NO         = adoCol.Value
                Case "実績年月":              If Not IsNull(adoCol.Value) Then .実績年月             = adoCol.Value
                Case "品番":                  If Not IsNull(adoCol.Value) Then .品番                 = adoCol.Value
                Case "背番号":                If Not IsNull(adoCol.Value) Then .背番号               = adoCol.Value
                Case "納入先":                If Not IsNull(adoCol.Value) Then .納入先               = adoCol.Value
                Case "契約先":                If Not IsNull(adoCol.Value) Then .契約先               = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入                 = adoCol.Value
                Case "車型":                  If Not IsNull(adoCol.Value) Then .車型                 = adoCol.Value
                Case "実績区分":              If Not IsNull(adoCol.Value) Then .実績区分             = adoCol.Value
                Case "実績処理日":            If Not IsNull(adoCol.Value) Then .実績処理日           = adoCol.Value
                Case "正式実績年月":          If Not IsNull(adoCol.Value) Then .正式実績年月         = adoCol.Value
                Case "正式実績処理日":        If Not IsNull(adoCol.Value) Then .正式実績処理日       = adoCol.Value
                Case "納入番号":              If Not IsNull(adoCol.Value) Then .納入番号             = adoCol.Value
                Case "検収年月日":            If Not IsNull(adoCol.Value) Then .検収年月日           = adoCol.Value
                Case "赤黒区分":              If Not IsNull(adoCol.Value) Then .赤黒区分             = adoCol.Value
                Case "単価":                  If Not IsNull(adoCol.Value) Then .単価                 = adoCol.Value
                Case "仮単価区分":            If Not IsNull(adoCol.Value) Then .仮単価区分           = adoCol.Value
                Case "納入数":                If Not IsNull(adoCol.Value) Then .納入数               = adoCol.Value
                Case "適用開始日付":          If Not IsNull(adoCol.Value) Then .適用開始日付         = adoCol.Value
                Case "金額":                  If Not IsNull(adoCol.Value) Then .金額                 = adoCol.Value
                Case "計算方法":              If Not IsNull(adoCol.Value) Then .計算方法             = adoCol.Value
                Case "計算単位":              If Not IsNull(adoCol.Value) Then .計算単位             = adoCol.Value
                Case "締日":                  If Not IsNull(adoCol.Value) Then .締日                 = adoCol.Value
                Case "前回実績年月":          If Not IsNull(adoCol.Value) Then .前回実績年月         = adoCol.Value
                Case "明細件数":              If Not IsNull(adoCol.Value) Then .明細件数             = adoCol.Value
                Case "入力区分":              If Not IsNull(adoCol.Value) Then .入力区分             = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "作成者":                If Not IsNull(adoCol.Value) Then .作成者               = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "契約先名":	If Not IsNull(adoCol.Value) Then .契約先名	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "納入先名":	If Not IsNull(adoCol.Value) Then .納入先名	= adoCol.Value
		Case "照合対象":	If Not IsNull(adoCol.Value) Then .照合対象	= adoCol.Value
		Case "自社工区":	If Not IsNull(adoCol.Value) Then .自社工区	= adoCol.Value
		Case "自社工区名":	If Not IsNull(adoCol.Value) Then .自社工区名	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function UJTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    UJTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 売上実績テーブル"
        RTxtSql = RTxtSql &" Where 検収管理NO = " & UJT.検収管理NO & ""
        RTxtSql = RTxtSql &" And   検収管理行NO = " & UJT.検収管理行NO & ""
        RTxtSql = RTxtSql &" And   実績年月 = '" & RTrim(UJT.実績年月) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           UJTRead = True
           Call ItemsFromUJT(DBNo)
	   UJT_RDSTS = True
        else
	   UJT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call UJTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub UJTReadNext(Optional DBNo As Integer = 0)
    If Not UJT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           UJT_RDSTS = False
        Else
           Call ItemsFromUJT(DBNo)
           UJT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub UJTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        UJT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( UJT )
'*----------------------------------------------*'
Public Sub ItemsClearUJT()
    With UJT
	.検収管理NO = 0
	.検収管理行NO = 0
	.実績年月 = ""
	.品番 = ""
	.背番号 = ""
	.納入先 = ""
	.契約先 = ""
	.受入 = ""
	.車型 = ""
	.実績区分 = 0
	.実績処理日 = ""
	.正式実績年月 = ""
	.正式実績処理日 = ""
	.納入番号 = ""
	.検収年月日 = ""
	.赤黒区分 = 0
	.単価 = 0
	.仮単価区分 = ""
	.納入数 = 0
	.適用開始日付 = ""
	.金額 = 0
	.計算方法 = 0
	.計算単位 = 0
	.締日 = 0
	.前回実績年月 = ""
	.明細件数 = 0
	.入力区分 = ""
'	.作成日 = 0
'	.作成端末 = ""
	.作成者 = ""
        .件数 = 0
'   項目追加(自動)
	.契約先名 = ""
	.表示品番 = ""
	.納入先名 = ""
	.照合対象 = ""
	.自社工区 = ""
	.自社工区名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function UJTInsert(Optional DBNo As Integer) As Boolean
    UJTInsert = False
    With UJT
        gSL_Select = "Insert Into 売上実績テーブル  ( "
        gSL_Select = gSL_Select & "検収管理NO,検収管理行NO,実績年月,"
        gSL_Select = gSL_Select & "品番,背番号,納入先,"
        gSL_Select = gSL_Select & "契約先,受入,車型,"
        gSL_Select = gSL_Select & "実績区分,実績処理日,正式実績年月,"
        gSL_Select = gSL_Select & "正式実績処理日,納入番号,検収年月日,"
        gSL_Select = gSL_Select & "赤黒区分,単価,仮単価区分,"
        gSL_Select = gSL_Select & "納入数,適用開始日付,金額,"
        gSL_Select = gSL_Select & "計算方法,計算単位,締日,"
        gSL_Select = gSL_Select & "前回実績年月,明細件数,入力区分,"
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .検収管理NO & ","		
        gSL_Select = gSL_Select & "" & .検収管理行NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.実績年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.車型) & "',"		
        gSL_Select = gSL_Select & "" & .実績区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.実績処理日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.正式実績年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.正式実績処理日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.検収年月日) & "',"		
        gSL_Select = gSL_Select & "" & .赤黒区分 & ","		
        gSL_Select = gSL_Select & "" & .単価 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.仮単価区分) & "',"		
        gSL_Select = gSL_Select & "" & .納入数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.適用開始日付) & "',"		
        gSL_Select = gSL_Select & "" & .金額 & ","		
        gSL_Select = gSL_Select & "" & .計算方法 & ","		
        gSL_Select = gSL_Select & "" & .計算単位 & ","		
        gSL_Select = gSL_Select & "" & .締日 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.前回実績年月) & "',"		
        gSL_Select = gSL_Select & "" & .明細件数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.入力区分) & "',"		
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
    UJTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function UJTUpdate(Optional DBNo As Integer) As Boolean
    UJTUpdate = False	
    With UJT
        gSL_Select = "Update 売上実績テーブル  Set "
        gSL_Select = gSL_Select & "検収管理NO = " & .検収管理NO & ","
        gSL_Select = gSL_Select & "検収管理行NO = " & .検収管理行NO & ","
        gSL_Select = gSL_Select & "実績年月 = '" & RTrim(.実績年月) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "車型 = '" & RTrim(.車型) & "',"
        gSL_Select = gSL_Select & "実績区分 = " & .実績区分 & ","
        gSL_Select = gSL_Select & "実績処理日 = '" & RTrim(.実績処理日) & "',"
        gSL_Select = gSL_Select & "正式実績年月 = '" & RTrim(.正式実績年月) & "',"
        gSL_Select = gSL_Select & "正式実績処理日 = '" & RTrim(.正式実績処理日) & "',"
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "検収年月日 = '" & RTrim(.検収年月日) & "',"
        gSL_Select = gSL_Select & "赤黒区分 = " & .赤黒区分 & ","
        gSL_Select = gSL_Select & "単価 = " & .単価 & ","
        gSL_Select = gSL_Select & "仮単価区分 = '" & RTrim(.仮単価区分) & "',"
        gSL_Select = gSL_Select & "納入数 = " & .納入数 & ","
        gSL_Select = gSL_Select & "適用開始日付 = '" & RTrim(.適用開始日付) & "',"
        gSL_Select = gSL_Select & "金額 = " & .金額 & ","
        gSL_Select = gSL_Select & "計算方法 = " & .計算方法 & ","
        gSL_Select = gSL_Select & "計算単位 = " & .計算単位 & ","
        gSL_Select = gSL_Select & "締日 = " & .締日 & ","
        gSL_Select = gSL_Select & "前回実績年月 = '" & RTrim(.前回実績年月) & "',"
        gSL_Select = gSL_Select & "明細件数 = " & .明細件数 & ","
        gSL_Select = gSL_Select & "入力区分 = '" & RTrim(.入力区分) & "'"
        gSL_Select = gSL_Select & " Where 検収管理NO = " & .検収管理NO & ""
        gSL_Select = gSL_Select & "   And 検収管理行NO = " & .検収管理行NO & ""
        gSL_Select = gSL_Select & "   And 実績年月 = '" & RTrim(.実績年月) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    UJTUpdate = True
End Function

Attribute VB_Name = "売上差異明細テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   売上差異明細テーブル
'*
'*   モジュールID  : 売上差異明細テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/11/17  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_売上差異明細テーブル
        照合年月      As String
        照合結果      As String
        契約先        As String
        照合日付      As String
        照合品番      As String
        照合納番      As String
        照合管理      As String
        財務管理番号  As Long
        受入          As String
        繰越          As String
        売上年月      As String
        対象年月      As String
        管理番号      As String
        品番          As String
        納入日        As String
        納入番号      As String
        FCD           As String
        繰越区分      As String
        識別子        As String
        支給先        As String
        仮単価        As Byte
        自社単価      As Currency
        自社数量      As Long
        自社金額      As Long
        契約先単価    As Currency
        契約先数量    As Long
        契約先金額    As Long
        最大日付      As String
        削除区分      As String
'        作成日        As Date
'        作成端末      As String
'        更新日        As Date
'        更新端末      As String
        件数          As Long
'   項目追加(自動)
	契約先名      As String
	照合結果名    As String
	表示品番      As String
	識別子名      As String
    End Type

    Public UST As T_売上差異明細テーブル

    Public UST_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromUST(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearUST	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With UST
            Select Case adoCol.Name
                Case "照合年月":      If Not IsNull(adoCol.Value) Then .照合年月     = adoCol.Value
                Case "照合結果":      If Not IsNull(adoCol.Value) Then .照合結果     = adoCol.Value
                Case "契約先":        If Not IsNull(adoCol.Value) Then .契約先       = adoCol.Value
                Case "照合日付":      If Not IsNull(adoCol.Value) Then .照合日付     = adoCol.Value
                Case "照合品番":      If Not IsNull(adoCol.Value) Then .照合品番     = adoCol.Value
                Case "照合納番":      If Not IsNull(adoCol.Value) Then .照合納番     = adoCol.Value
                Case "照合管理":      If Not IsNull(adoCol.Value) Then .照合管理     = adoCol.Value
                Case "財務管理番号":  If Not IsNull(adoCol.Value) Then .財務管理番号 = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "繰越":          If Not IsNull(adoCol.Value) Then .繰越         = adoCol.Value
                Case "売上年月":      If Not IsNull(adoCol.Value) Then .売上年月     = adoCol.Value
                Case "対象年月":      If Not IsNull(adoCol.Value) Then .対象年月     = adoCol.Value
                Case "管理番号":      If Not IsNull(adoCol.Value) Then .管理番号     = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "納入日":        If Not IsNull(adoCol.Value) Then .納入日       = adoCol.Value
                Case "納入番号":      If Not IsNull(adoCol.Value) Then .納入番号     = adoCol.Value
                Case "FCD":           If Not IsNull(adoCol.Value) Then .FCD          = adoCol.Value
                Case "繰越区分":      If Not IsNull(adoCol.Value) Then .繰越区分     = adoCol.Value
                Case "識別子":        If Not IsNull(adoCol.Value) Then .識別子       = adoCol.Value
                Case "支給先":        If Not IsNull(adoCol.Value) Then .支給先       = adoCol.Value
                Case "仮単価":        If Not IsNull(adoCol.Value) Then .仮単価       = adoCol.Value
                Case "自社単価":      If Not IsNull(adoCol.Value) Then .自社単価     = adoCol.Value
                Case "自社数量":      If Not IsNull(adoCol.Value) Then .自社数量     = adoCol.Value
                Case "自社金額":      If Not IsNull(adoCol.Value) Then .自社金額     = adoCol.Value
                Case "契約先単価":    If Not IsNull(adoCol.Value) Then .契約先単価   = adoCol.Value
                Case "契約先数量":    If Not IsNull(adoCol.Value) Then .契約先数量   = adoCol.Value
                Case "契約先金額":    If Not IsNull(adoCol.Value) Then .契約先金額   = adoCol.Value
                Case "最大日付":      If Not IsNull(adoCol.Value) Then .最大日付     = adoCol.Value
                Case "削除区分":      If Not IsNull(adoCol.Value) Then .削除区分     = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "契約先名":	If Not IsNull(adoCol.Value) Then .契約先名	= adoCol.Value
		Case "照合結果名":	If Not IsNull(adoCol.Value) Then .照合結果名	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "識別子名":	If Not IsNull(adoCol.Value) Then .識別子名	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function USTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    USTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 売上差異明細テーブル"
        RTxtSql = RTxtSql &" Where 照合年月 = '" & RTrim(UST.照合年月) & "'"
        RTxtSql = RTxtSql &" And   照合結果 = '" & RTrim(UST.照合結果) & "'"
        RTxtSql = RTxtSql &" And   契約先 = '" & RTrim(UST.契約先) & "'"
        RTxtSql = RTxtSql &" And   照合日付 = '" & RTrim(UST.照合日付) & "'"
        RTxtSql = RTxtSql &" And   照合品番 = '" & RTrim(UST.照合品番) & "'"
        RTxtSql = RTxtSql &" And   照合納番 = '" & RTrim(UST.照合納番) & "'"
        RTxtSql = RTxtSql &" And   照合管理 = '" & RTrim(UST.照合管理) & "'"
        RTxtSql = RTxtSql &" And   財務管理番号 = " & UST.財務管理番号 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           USTRead = True
           Call ItemsFromUST(DBNo)
	   UST_RDSTS = True
        else
	   UST_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call USTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub USTReadNext(Optional DBNo As Integer = 0)
    If Not UST_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           UST_RDSTS = False
        Else
           Call ItemsFromUST(DBNo)
           UST_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub USTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        UST_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( UST )
'*----------------------------------------------*'
Public Sub ItemsClearUST()
    With UST
	.照合年月 = ""
	.照合結果 = ""
	.契約先 = ""
	.照合日付 = ""
	.照合品番 = ""
	.照合納番 = ""
	.照合管理 = ""
	.財務管理番号 = 0
	.受入 = ""
	.繰越 = ""
	.売上年月 = ""
	.対象年月 = ""
	.管理番号 = ""
	.品番 = ""
	.納入日 = ""
	.納入番号 = ""
	.FCD = ""
	.繰越区分 = ""
	.識別子 = ""
	.支給先 = ""
	.仮単価 = 0
	.自社単価 = 0
	.自社数量 = 0
	.自社金額 = 0
	.契約先単価 = 0
	.契約先数量 = 0
	.契約先金額 = 0
	.最大日付 = ""
	.削除区分 = ""
'	.作成日 = 0
'	.作成端末 = ""
'	.更新日 = 0
'	.更新端末 = ""
        .件数 = 0
'   項目追加(自動)
	.契約先名 = ""
	.照合結果名 = ""
	.表示品番 = ""
	.識別子名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function USTInsert(Optional DBNo As Integer) As Boolean
    USTInsert = True
    With UST
        gSL_Select = "Insert Into 売上差異明細テーブル  ( "
        gSL_Select = gSL_Select & "照合年月,照合結果,契約先,"
        gSL_Select = gSL_Select & "照合日付,照合品番,照合納番,"
        gSL_Select = gSL_Select & "照合管理,財務管理番号,受入,"
        gSL_Select = gSL_Select & "繰越,売上年月,対象年月,"
        gSL_Select = gSL_Select & "管理番号,品番,納入日,"
        gSL_Select = gSL_Select & "納入番号,FCD,繰越区分,"
        gSL_Select = gSL_Select & "識別子,支給先,仮単価,"
        gSL_Select = gSL_Select & "自社単価,自社数量,自社金額,"
        gSL_Select = gSL_Select & "契約先単価,契約先数量,契約先金額,"
        gSL_Select = gSL_Select & "最大日付,削除区分,"
        gSL_Select = gSL_Select & "作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.照合年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合結果) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合日付) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合納番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合管理) & "',"		
        gSL_Select = gSL_Select & "" & .財務管理番号 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.繰越) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.売上年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.対象年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.管理番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.FCD) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.繰越区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.識別子) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給先) & "',"		
        gSL_Select = gSL_Select & "" & .仮単価 & ","		
        gSL_Select = gSL_Select & "" & .自社単価 & ","		
        gSL_Select = gSL_Select & "" & .自社数量 & ","		
        gSL_Select = gSL_Select & "" & .自社金額 & ","		
        gSL_Select = gSL_Select & "" & .契約先単価 & ","		
        gSL_Select = gSL_Select & "" & .契約先数量 & ","		
        gSL_Select = gSL_Select & "" & .契約先金額 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.最大日付) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.削除区分) & "',"		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    USTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function USTUpdate(Optional DBNo As Integer) As Boolean
    USTUpdate = False	
    With UST
        gSL_Select = "Update 売上差異明細テーブル  Set "
        gSL_Select = gSL_Select & "照合年月 = '" & RTrim(.照合年月) & "',"
        gSL_Select = gSL_Select & "照合結果 = '" & RTrim(.照合結果) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "照合日付 = '" & RTrim(.照合日付) & "',"
        gSL_Select = gSL_Select & "照合品番 = '" & RTrim(.照合品番) & "',"
        gSL_Select = gSL_Select & "照合納番 = '" & RTrim(.照合納番) & "',"
        gSL_Select = gSL_Select & "照合管理 = '" & RTrim(.照合管理) & "',"
        gSL_Select = gSL_Select & "財務管理番号 = " & .財務管理番号 & ","
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "繰越 = '" & RTrim(.繰越) & "',"
        gSL_Select = gSL_Select & "売上年月 = '" & RTrim(.売上年月) & "',"
        gSL_Select = gSL_Select & "対象年月 = '" & RTrim(.対象年月) & "',"
        gSL_Select = gSL_Select & "管理番号 = '" & RTrim(.管理番号) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "FCD = '" & RTrim(.FCD) & "',"
        gSL_Select = gSL_Select & "繰越区分 = '" & RTrim(.繰越区分) & "',"
        gSL_Select = gSL_Select & "識別子 = '" & RTrim(.識別子) & "',"
        gSL_Select = gSL_Select & "支給先 = '" & RTrim(.支給先) & "',"
        gSL_Select = gSL_Select & "仮単価 = " & .仮単価 & ","
        gSL_Select = gSL_Select & "自社単価 = " & .自社単価 & ","
        gSL_Select = gSL_Select & "自社数量 = " & .自社数量 & ","
        gSL_Select = gSL_Select & "自社金額 = " & .自社金額 & ","
        gSL_Select = gSL_Select & "契約先単価 = " & .契約先単価 & ","
        gSL_Select = gSL_Select & "契約先数量 = " & .契約先数量 & ","
        gSL_Select = gSL_Select & "契約先金額 = " & .契約先金額 & ","
        gSL_Select = gSL_Select & "最大日付 = '" & RTrim(.最大日付) & "',"
        gSL_Select = gSL_Select & "削除区分 = '" & RTrim(.削除区分) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 照合年月 = '" & RTrim(.照合年月) & "'"
        gSL_Select = gSL_Select & "   And 照合結果 = '" & RTrim(.照合結果) & "'"
        gSL_Select = gSL_Select & "   And 契約先 = '" & RTrim(.契約先) & "'"
        gSL_Select = gSL_Select & "   And 照合日付 = '" & RTrim(.照合日付) & "'"
        gSL_Select = gSL_Select & "   And 照合品番 = '" & RTrim(.照合品番) & "'"
        gSL_Select = gSL_Select & "   And 照合納番 = '" & RTrim(.照合納番) & "'"
        gSL_Select = gSL_Select & "   And 照合管理 = '" & RTrim(.照合管理) & "'"
        gSL_Select = gSL_Select & "   And 財務管理番号 = " & .財務管理番号 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    USTUpdate = True
End Function

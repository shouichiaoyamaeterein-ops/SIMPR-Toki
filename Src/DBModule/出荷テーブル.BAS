Attribute VB_Name = "出荷テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   出荷テーブル
'*
'*   モジュールID  : 出荷テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/03/25  By Cis
'*   作成日  :  2008/05/16  By Cis
'*   作成日  :  2008/05/19  By Cis
'*   作成日  :  2008/05/22  By Cis
'*   作成日  :  2008/05/23  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_出荷テーブル
        出荷指示書NO          As Long
        出荷指示書行NO        As Byte
        出荷日                As String
        出荷便                As String
        出荷枚数              As Long
        出荷端数              As Long
        出荷数                As Long
        分納数                As Long
        照合年月              As String
        照合処理日            As String
        照合結果              As String
        実績年月              As String
        実績処理日            As String
        実績結果              As String
'        作成日                As Date
'        作成端末              As String
        作成者                As String
'        更新日                As Date
'        更新端末              As String
        更新者                As String
        件数          As Long
'   項目追加(自動)
	契約先名       As String
	車型           As String
	品番           As String
	予測金額       As Currency
	実績金額       As Currency
	自社工区名     As String
	進捗           As Currency
	実績金額0      As Currency
	実績金額1      As Currency
	実績金額2      As Currency
	実績金額3      As Currency
	実績金額4      As Currency
	実績金額5      As Currency
	日当り数       As Currency
	自社工区       As String
	契約先         As String
	編集契約先     As String
    End Type

    Public SKT As T_出荷テーブル

    Public SKT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromSKT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearSKT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With SKT
            Select Case adoCol.Name
                Case "出荷指示書NO":          If Not IsNull(adoCol.Value) Then .出荷指示書NO         = adoCol.Value
                Case "出荷指示書行NO":        If Not IsNull(adoCol.Value) Then .出荷指示書行NO       = adoCol.Value
                Case "出荷日":                If Not IsNull(adoCol.Value) Then .出荷日               = adoCol.Value
                Case "出荷便":                If Not IsNull(adoCol.Value) Then .出荷便               = adoCol.Value
                Case "出荷枚数":              If Not IsNull(adoCol.Value) Then .出荷枚数             = adoCol.Value
                Case "出荷端数":              If Not IsNull(adoCol.Value) Then .出荷端数             = adoCol.Value
                Case "出荷数":                If Not IsNull(adoCol.Value) Then .出荷数               = adoCol.Value
                Case "分納数":                If Not IsNull(adoCol.Value) Then .分納数               = adoCol.Value
                Case "照合年月":              If Not IsNull(adoCol.Value) Then .照合年月             = adoCol.Value
                Case "照合処理日":            If Not IsNull(adoCol.Value) Then .照合処理日           = adoCol.Value
                Case "照合結果":              If Not IsNull(adoCol.Value) Then .照合結果             = adoCol.Value
                Case "実績年月":              If Not IsNull(adoCol.Value) Then .実績年月             = adoCol.Value
                Case "実績処理日":            If Not IsNull(adoCol.Value) Then .実績処理日           = adoCol.Value
                Case "実績結果":              If Not IsNull(adoCol.Value) Then .実績結果             = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "作成者":                If Not IsNull(adoCol.Value) Then .作成者               = adoCol.Value
'                Case "更新日":                If Not IsNull(adoCol.Value) Then .更新日               = adoCol.Value
'                Case "更新端末":              If Not IsNull(adoCol.Value) Then .更新端末             = adoCol.Value
                Case "更新者":                If Not IsNull(adoCol.Value) Then .更新者               = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "契約先名":	If Not IsNull(adoCol.Value) Then .契約先名	= adoCol.Value
		Case "車型":	If Not IsNull(adoCol.Value) Then .車型	= adoCol.Value
		Case "品番":	If Not IsNull(adoCol.Value) Then .品番	= adoCol.Value
		Case "予測金額":	If Not IsNull(adoCol.Value) Then .予測金額	= adoCol.Value
		Case "実績金額":	If Not IsNull(adoCol.Value) Then .実績金額	= adoCol.Value
		Case "自社工区名":	If Not IsNull(adoCol.Value) Then .自社工区名	= adoCol.Value
		Case "進捗":	If Not IsNull(adoCol.Value) Then .進捗	= adoCol.Value
		Case "実績金額0":	If Not IsNull(adoCol.Value) Then .実績金額0	= adoCol.Value
		Case "実績金額1":	If Not IsNull(adoCol.Value) Then .実績金額1	= adoCol.Value
		Case "実績金額2":	If Not IsNull(adoCol.Value) Then .実績金額2	= adoCol.Value
		Case "実績金額3":	If Not IsNull(adoCol.Value) Then .実績金額3	= adoCol.Value
		Case "実績金額4":	If Not IsNull(adoCol.Value) Then .実績金額4	= adoCol.Value
		Case "実績金額5":	If Not IsNull(adoCol.Value) Then .実績金額5	= adoCol.Value
		Case "日当り数":	If Not IsNull(adoCol.Value) Then .日当り数	= adoCol.Value
		Case "自社工区":	If Not IsNull(adoCol.Value) Then .自社工区	= adoCol.Value
		Case "契約先":	If Not IsNull(adoCol.Value) Then .契約先	= adoCol.Value
		Case "編集契約先":	If Not IsNull(adoCol.Value) Then .編集契約先	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function SKTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    SKTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 出荷テーブル"
        RTxtSql = RTxtSql &" Where 出荷指示書NO = " & SKT.出荷指示書NO & ""
        RTxtSql = RTxtSql &" And   出荷指示書行NO = " & SKT.出荷指示書行NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           SKTRead = True
           Call ItemsFromSKT(DBNo)
	   SKT_RDSTS = True
        else
	   SKT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call SKTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub SKTReadNext(Optional DBNo As Integer = 0)
    If Not SKT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           SKT_RDSTS = False
        Else
           Call ItemsFromSKT(DBNo)
           SKT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub SKTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        SKT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( SKT )
'*----------------------------------------------*'
Public Sub ItemsClearSKT()
    With SKT
	.出荷指示書NO = 0
	.出荷指示書行NO = 0
	.出荷日 = ""
	.出荷便 = ""
	.出荷枚数 = 0
	.出荷端数 = 0
	.出荷数 = 0
	.分納数 = 0
	.照合年月 = ""
	.照合処理日 = ""
	.照合結果 = ""
	.実績年月 = ""
	.実績処理日 = ""
	.実績結果 = ""
'	.作成日 = 0
'	.作成端末 = ""
	.作成者 = ""
'	.更新日 = 0
'	.更新端末 = ""
	.更新者 = ""
        .件数 = 0
'   項目追加(自動)
	.契約先名 = ""
	.車型 = ""
	.品番 = ""
	.予測金額 = 0
	.実績金額 = 0
	.自社工区名 = ""
	.進捗 = 0
	.実績金額0 = 0
	.実績金額1 = 0
	.実績金額2 = 0
	.実績金額3 = 0
	.実績金額4 = 0
	.実績金額5 = 0
	.日当り数 = 0
	.自社工区 = ""
	.契約先 = ""
	.編集契約先 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function SKTInsert(Optional DBNo As Integer) As Boolean
    SKTInsert = False
    With SKT
        gSL_Select = "Insert Into 出荷テーブル  ( "
        gSL_Select = gSL_Select & "出荷指示書NO,出荷指示書行NO,出荷日,"
        gSL_Select = gSL_Select & "出荷便,出荷枚数,出荷端数,"
        gSL_Select = gSL_Select & "出荷数,分納数,照合年月,"
        gSL_Select = gSL_Select & "照合処理日,照合結果,実績年月,"
        gSL_Select = gSL_Select & "実績処理日,実績結果,"
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .出荷指示書NO & ","		
        gSL_Select = gSL_Select & "" & .出荷指示書行NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.出荷日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.出荷便) & "',"		
        gSL_Select = gSL_Select & "" & .出荷枚数 & ","		
        gSL_Select = gSL_Select & "" & .出荷端数 & ","		
        gSL_Select = gSL_Select & "" & .出荷数 & ","		
        gSL_Select = gSL_Select & "" & .分納数 & ","		
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
    SKTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function SKTUpdate(Optional DBNo As Integer) As Boolean
    SKTUpdate = False	
    With SKT
        gSL_Select = "Update 出荷テーブル  Set "
        gSL_Select = gSL_Select & "出荷指示書NO = " & .出荷指示書NO & ","
        gSL_Select = gSL_Select & "出荷指示書行NO = " & .出荷指示書行NO & ","
        gSL_Select = gSL_Select & "出荷日 = '" & RTrim(.出荷日) & "',"
        gSL_Select = gSL_Select & "出荷便 = '" & RTrim(.出荷便) & "',"
        gSL_Select = gSL_Select & "出荷枚数 = " & .出荷枚数 & ","
        gSL_Select = gSL_Select & "出荷端数 = " & .出荷端数 & ","
        gSL_Select = gSL_Select & "出荷数 = " & .出荷数 & ","
        gSL_Select = gSL_Select & "分納数 = " & .分納数 & ","
        gSL_Select = gSL_Select & "照合年月 = '" & RTrim(.照合年月) & "',"
        gSL_Select = gSL_Select & "照合処理日 = '" & RTrim(.照合処理日) & "',"
        gSL_Select = gSL_Select & "照合結果 = '" & RTrim(.照合結果) & "',"
        gSL_Select = gSL_Select & "実績年月 = '" & RTrim(.実績年月) & "',"
        gSL_Select = gSL_Select & "実績処理日 = '" & RTrim(.実績処理日) & "',"
        gSL_Select = gSL_Select & "実績結果 = '" & RTrim(.実績結果) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "'"
        gSL_Select = gSL_Select & " Where 出荷指示書NO = " & .出荷指示書NO & ""
        gSL_Select = gSL_Select & "   And 出荷指示書行NO = " & .出荷指示書行NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SKTUpdate = True
End Function

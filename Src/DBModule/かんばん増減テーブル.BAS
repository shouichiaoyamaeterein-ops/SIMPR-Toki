Attribute VB_Name = "かんばん増減テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   かんばん増減テーブル
'*
'*   モジュールID  : かんばん増減テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2005/05/27  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_かんばん増減テーブル
        取引先        As String
        受入          As String
        品番          As String
        開始日        As String
        終了日        As String
        設定内容区分  As Byte
        増減区分      As Byte
        設定枚数      As Long
        便当枚数      As Long
        累積枚数      As Long
        完了日        As String
'        作成日        As Date
        作成者        As String
'        作成端末      As String
'        更新日        As Date
        更新者        As String
'        更新端末      As String
        件数          As Long
'   項目追加(自動)
	取引先名      As String
	表示品番      As String
	背番号        As String
	収容数        As Long
	当月内示数    As Long
	翌月内示数    As Long
	前月内示数    As Long
	変更日        As String
	収容数1       As Long
	収容数2       As Long
	回転枚数      As Long
	有効枚数      As Long
    End Type

    Public KZT As T_かんばん増減テーブル

    Public KZT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKZT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKZT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KZT
            Select Case adoCol.Name
                Case "取引先":        If Not IsNull(adoCol.Value) Then .取引先       = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "開始日":        If Not IsNull(adoCol.Value) Then .開始日       = adoCol.Value
                Case "終了日":        If Not IsNull(adoCol.Value) Then .終了日       = adoCol.Value
                Case "設定内容区分":  If Not IsNull(adoCol.Value) Then .設定内容区分 = adoCol.Value
                Case "増減区分":      If Not IsNull(adoCol.Value) Then .増減区分     = adoCol.Value
                Case "設定枚数":      If Not IsNull(adoCol.Value) Then .設定枚数     = adoCol.Value
                Case "便当枚数":      If Not IsNull(adoCol.Value) Then .便当枚数     = adoCol.Value
                Case "累積枚数":      If Not IsNull(adoCol.Value) Then .累積枚数     = adoCol.Value
                Case "完了日":        If Not IsNull(adoCol.Value) Then .完了日       = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者       = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "取引先名":	If Not IsNull(adoCol.Value) Then .取引先名	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "背番号":	If Not IsNull(adoCol.Value) Then .背番号	= adoCol.Value
		Case "収容数":	If Not IsNull(adoCol.Value) Then .収容数	= adoCol.Value
		Case "当月内示数":	If Not IsNull(adoCol.Value) Then .当月内示数	= adoCol.Value
		Case "翌月内示数":	If Not IsNull(adoCol.Value) Then .翌月内示数	= adoCol.Value
		Case "前月内示数":	If Not IsNull(adoCol.Value) Then .前月内示数	= adoCol.Value
		Case "変更日":	If Not IsNull(adoCol.Value) Then .変更日	= adoCol.Value
		Case "収容数1":	If Not IsNull(adoCol.Value) Then .収容数1	= adoCol.Value
		Case "収容数2":	If Not IsNull(adoCol.Value) Then .収容数2	= adoCol.Value
		Case "回転枚数":	If Not IsNull(adoCol.Value) Then .回転枚数	= adoCol.Value
		Case "有効枚数":	If Not IsNull(adoCol.Value) Then .有効枚数	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KZTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KZTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From かんばん増減テーブル"
        RTxtSql = RTxtSql &" Where 取引先 = '" & RTrim(KZT.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(KZT.受入) & "'"
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(KZT.品番) & "'"
        RTxtSql = RTxtSql &" And   開始日 = '" & RTrim(KZT.開始日) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KZTRead = True
           Call ItemsFromKZT(DBNo)
	   KZT_RDSTS = True
        else
	   KZT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KZTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KZTReadNext(Optional DBNo As Integer = 0)
    If Not KZT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KZT_RDSTS = False
        Else
           Call ItemsFromKZT(DBNo)
           KZT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KZTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KZT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KZT )
'*----------------------------------------------*'
Public Sub ItemsClearKZT()
    With KZT
	.取引先 = ""
	.受入 = ""
	.品番 = ""
	.開始日 = ""
	.終了日 = ""
	.設定内容区分 = 0
	.増減区分 = 0
	.設定枚数 = 0
	.便当枚数 = 0
	.累積枚数 = 0
	.完了日 = ""
'	.作成日 = Nothing
	.作成者 = ""
'	.作成端末 = ""
'	.更新日 = Nothing
	.更新者 = ""
'	.更新端末 = ""
        .件数 = 0
'   項目追加(自動)
	.取引先名 = ""
	.表示品番 = ""
	.背番号 = ""
	.収容数 = 0
	.当月内示数 = 0
	.翌月内示数 = 0
	.前月内示数 = 0
	.変更日 = ""
	.収容数1 = 0
	.収容数2 = 0
	.回転枚数 = 0
	.有効枚数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KZTInsert(Optional DBNo As Integer) As Boolean
    KZTInsert = True
    With KZT
        gSL_Select = "Insert Into かんばん増減テーブル  ( "
        gSL_Select = gSL_Select & "取引先,受入,品番,"
        gSL_Select = gSL_Select & "開始日,終了日,設定内容区分,"
        gSL_Select = gSL_Select & "増減区分,設定枚数,便当枚数,"
        gSL_Select = gSL_Select & "累積枚数,完了日,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.開始日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.終了日) & "',"		
        gSL_Select = gSL_Select & "" & .設定内容区分 & ","		
        gSL_Select = gSL_Select & "" & .増減区分 & ","		
        gSL_Select = gSL_Select & "" & .設定枚数 & ","		
        gSL_Select = gSL_Select & "" & .便当枚数 & ","		
        gSL_Select = gSL_Select & "" & .累積枚数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.完了日) & "',"		
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
    KZTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KZTUpdate(Optional DBNo As Integer) As Boolean
    KZTUpdate = False	
    With KZT
        gSL_Select = "Update かんばん増減テーブル  Set "
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "開始日 = '" & RTrim(.開始日) & "',"
        gSL_Select = gSL_Select & "終了日 = '" & RTrim(.終了日) & "',"
        gSL_Select = gSL_Select & "設定内容区分 = " & .設定内容区分 & ","
        gSL_Select = gSL_Select & "増減区分 = " & .増減区分 & ","
        gSL_Select = gSL_Select & "設定枚数 = " & .設定枚数 & ","
        gSL_Select = gSL_Select & "便当枚数 = " & .便当枚数 & ","
        gSL_Select = gSL_Select & "累積枚数 = " & .累積枚数 & ","
        gSL_Select = gSL_Select & "完了日 = '" & RTrim(.完了日) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 開始日 = '" & RTrim(.開始日) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KZTUpdate = True
End Function

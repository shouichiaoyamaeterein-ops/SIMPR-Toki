Attribute VB_Name = "空明細書テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   空明細書テーブル
'*
'*   モジュールID  : 空明細書テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/05/31  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_空明細書テーブル
        明細書NO      As Long
        発行区分      As String
        仕入先        As String
        受入          As String
        納入日        As String
        納入便        As String
        発注日        As String
        発注便        As String
        発行時間      As String
        発注区分      As String
        手配区分      As String
        SK区分        As Byte
        仕入先名      As String
        明細書発行日  As String
        作成種別      As String
'        作成日        As Date
'        作成端末      As String
        作成者        As String
        件数          As Long
    End Type

    Public KMT As T_空明細書テーブル

    Public KMT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKMT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKMT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KMT
            Select Case adoCol.Name
                Case "明細書NO":      If Not IsNull(adoCol.Value) Then .明細書NO     = adoCol.Value
                Case "発行区分":      If Not IsNull(adoCol.Value) Then .発行区分     = adoCol.Value
                Case "仕入先":        If Not IsNull(adoCol.Value) Then .仕入先       = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "納入日":        If Not IsNull(adoCol.Value) Then .納入日       = adoCol.Value
                Case "納入便":        If Not IsNull(adoCol.Value) Then .納入便       = adoCol.Value
                Case "発注日":        If Not IsNull(adoCol.Value) Then .発注日       = adoCol.Value
                Case "発注便":        If Not IsNull(adoCol.Value) Then .発注便       = adoCol.Value
                Case "発行時間":      If Not IsNull(adoCol.Value) Then .発行時間     = adoCol.Value
                Case "発注区分":      If Not IsNull(adoCol.Value) Then .発注区分     = adoCol.Value
                Case "手配区分":      If Not IsNull(adoCol.Value) Then .手配区分     = adoCol.Value
                Case "SK区分":        If Not IsNull(adoCol.Value) Then .SK区分       = adoCol.Value
                Case "仕入先名":      If Not IsNull(adoCol.Value) Then .仕入先名     = adoCol.Value
                Case "明細書発行日":  If Not IsNull(adoCol.Value) Then .明細書発行日 = adoCol.Value
                Case "作成種別":      If Not IsNull(adoCol.Value) Then .作成種別     = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KMTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KMTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 空明細書テーブル"
        RTxtSql = RTxtSql &" Where 明細書NO = " & KMT.明細書NO & ""
        RTxtSql = RTxtSql &" And   発行区分 = '" & RTrim(KMT.発行区分) & "'"
        RTxtSql = RTxtSql &" And   仕入先 = '" & RTrim(KMT.仕入先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(KMT.受入) & "'"
        RTxtSql = RTxtSql &" And   納入日 = '" & RTrim(KMT.納入日) & "'"
        RTxtSql = RTxtSql &" And   納入便 = '" & RTrim(KMT.納入便) & "'"
        RTxtSql = RTxtSql &" And   発注日 = '" & RTrim(KMT.発注日) & "'"
        RTxtSql = RTxtSql &" And   発注便 = '" & RTrim(KMT.発注便) & "'"
        RTxtSql = RTxtSql &" And   発行時間 = '" & RTrim(KMT.発行時間) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KMTRead = True
           Call ItemsFromKMT(DBNo)
	   KMT_RDSTS = True
        else
	   KMT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KMTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KMTReadNext(Optional DBNo As Integer = 0)
    If Not KMT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KMT_RDSTS = False
        Else
           Call ItemsFromKMT(DBNo)
           KMT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KMTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KMT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KMT )
'*----------------------------------------------*'
Public Sub ItemsClearKMT()
    With KMT
	.明細書NO = 0
	.発行区分 = ""
	.仕入先 = ""
	.受入 = ""
	.納入日 = ""
	.納入便 = ""
	.発注日 = ""
	.発注便 = ""
	.発行時間 = ""
	.発注区分 = ""
	.手配区分 = ""
	.SK区分 = 0
	.仕入先名 = ""
	.明細書発行日 = ""
	.作成種別 = ""
'	.作成日 = 0
'	.作成端末 = ""
	.作成者 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KMTInsert(Optional DBNo As Integer) As Boolean
    KMTInsert = True
    With KMT
        gSL_Select = "Insert Into 空明細書テーブル  ( "
        gSL_Select = gSL_Select & "明細書NO,発行区分,仕入先,"
        gSL_Select = gSL_Select & "受入,納入日,納入便,"
        gSL_Select = gSL_Select & "発注日,発注便,発行時間,"
        gSL_Select = gSL_Select & "発注区分,手配区分,SK区分,"
        gSL_Select = gSL_Select & "仕入先名,明細書発行日,作成種別,"
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .明細書NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.発行区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発行時間) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配区分) & "',"		
        gSL_Select = gSL_Select & "" & .SK区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.明細書発行日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.作成種別) & "',"		
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
    KMTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KMTUpdate(Optional DBNo As Integer) As Boolean
    KMTUpdate = False	
    With KMT
        gSL_Select = "Update 空明細書テーブル  Set "
        gSL_Select = gSL_Select & "明細書NO = " & .明細書NO & ","
        gSL_Select = gSL_Select & "発行区分 = '" & RTrim(.発行区分) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "発注日 = '" & RTrim(.発注日) & "',"
        gSL_Select = gSL_Select & "発注便 = '" & RTrim(.発注便) & "',"
        gSL_Select = gSL_Select & "発行時間 = '" & RTrim(.発行時間) & "',"
        gSL_Select = gSL_Select & "発注区分 = '" & RTrim(.発注区分) & "',"
        gSL_Select = gSL_Select & "手配区分 = '" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "SK区分 = " & .SK区分 & ","
        gSL_Select = gSL_Select & "仕入先名 = '" & RTrim(.仕入先名) & "',"
        gSL_Select = gSL_Select & "明細書発行日 = '" & RTrim(.明細書発行日) & "',"
        gSL_Select = gSL_Select & "作成種別 = '" & RTrim(.作成種別) & "'"
        gSL_Select = gSL_Select & " Where 明細書NO = " & .明細書NO & ""
        gSL_Select = gSL_Select & "   And 発行区分 = '" & RTrim(.発行区分) & "'"
        gSL_Select = gSL_Select & "   And 仕入先 = '" & RTrim(.仕入先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 納入日 = '" & RTrim(.納入日) & "'"
        gSL_Select = gSL_Select & "   And 納入便 = '" & RTrim(.納入便) & "'"
        gSL_Select = gSL_Select & "   And 発注日 = '" & RTrim(.発注日) & "'"
        gSL_Select = gSL_Select & "   And 発注便 = '" & RTrim(.発注便) & "'"
        gSL_Select = gSL_Select & "   And 発行時間 = '" & RTrim(.発行時間) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KMTUpdate = True
End Function

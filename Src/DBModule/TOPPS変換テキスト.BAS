Attribute VB_Name = "TOPPS変換テキスト"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TOPPS変換テキスト
'*
'*   モジュールID  : TOPPS変換テキスト.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/12/02  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_TOPPS変換テキスト
        履歴管理番号  As Long
        受注出荷区分  As Byte
        ファイルID    As String
        SEQ           As Long
        納入番号      As String
        REC           As String
        納入先        As String
        納入先工場    As String
        受入          As String
        納入日        As String
        納入便        As String
'        作成日        As Date
'        作成端末      As String
        件数          As Long
    End Type

    Public PHT As T_TOPPS変換テキスト

    Public PHT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromPHT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearPHT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With PHT
            Select Case adoCol.Name
                Case "履歴管理番号":  If Not IsNull(adoCol.Value) Then .履歴管理番号 = adoCol.Value
                Case "受注出荷区分":  If Not IsNull(adoCol.Value) Then .受注出荷区分 = adoCol.Value
                Case "ファイルID":    If Not IsNull(adoCol.Value) Then .ファイルID   = adoCol.Value
                Case "SEQ":           If Not IsNull(adoCol.Value) Then .SEQ          = adoCol.Value
                Case "納入番号":      If Not IsNull(adoCol.Value) Then .納入番号     = adoCol.Value
                Case "REC":           If Not IsNull(adoCol.Value) Then .REC          = adoCol.Value
                Case "納入先":        If Not IsNull(adoCol.Value) Then .納入先       = adoCol.Value
                Case "納入先工場":    If Not IsNull(adoCol.Value) Then .納入先工場   = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "納入日":        If Not IsNull(adoCol.Value) Then .納入日       = adoCol.Value
                Case "納入便":        If Not IsNull(adoCol.Value) Then .納入便       = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function PHTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    PHTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From TOPPS変換テキスト"
        RTxtSql = RTxtSql &" Where 履歴管理番号 = " & PHT.履歴管理番号 & ""
        RTxtSql = RTxtSql &" And   受注出荷区分 = " & PHT.受注出荷区分 & ""
        RTxtSql = RTxtSql &" And   ファイルID = '" & RTrim(PHT.ファイルID) & "'"
        RTxtSql = RTxtSql &" And   SEQ = " & PHT.SEQ & ""
        RTxtSql = RTxtSql &" And   納入番号 = '" & RTrim(PHT.納入番号) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           PHTRead = True
           Call ItemsFromPHT(DBNo)
	   PHT_RDSTS = True
        else
	   PHT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call PHTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub PHTReadNext(Optional DBNo As Integer = 0)
    If Not PHT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           PHT_RDSTS = False
        Else
           Call ItemsFromPHT(DBNo)
           PHT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub PHTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        PHT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( PHT )
'*----------------------------------------------*'
Public Sub ItemsClearPHT()
    With PHT
	.履歴管理番号 = 0
	.受注出荷区分 = 0
	.ファイルID = ""
	.SEQ = 0
	.納入番号 = ""
	.REC = ""
	.納入先 = ""
	.納入先工場 = ""
	.受入 = ""
	.納入日 = ""
	.納入便 = ""
'	.作成日 = 0
'	.作成端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function PHTInsert(Optional DBNo As Integer) As Boolean
    PHTInsert = True
    With PHT
        gSL_Select = "Insert Into TOPPS変換テキスト  ( "
        gSL_Select = gSL_Select & "履歴管理番号,受注出荷区分,ファイルID,"
        gSL_Select = gSL_Select & "SEQ,納入番号,REC,"
        gSL_Select = gSL_Select & "納入先,納入先工場,受入,"
        gSL_Select = gSL_Select & "納入日,納入便,"
        gSL_Select = gSL_Select & "作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .履歴管理番号 & ","		
        gSL_Select = gSL_Select & "" & .受注出荷区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.ファイルID) & "',"		
        gSL_Select = gSL_Select & "" & .SEQ & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.REC) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先工場) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    PHTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function PHTUpdate(Optional DBNo As Integer) As Boolean
    PHTUpdate = False	
    With PHT
        gSL_Select = "Update TOPPS変換テキスト  Set "
        gSL_Select = gSL_Select & "履歴管理番号 = " & .履歴管理番号 & ","
        gSL_Select = gSL_Select & "受注出荷区分 = " & .受注出荷区分 & ","
        gSL_Select = gSL_Select & "ファイルID = '" & RTrim(.ファイルID) & "',"
        gSL_Select = gSL_Select & "SEQ = " & .SEQ & ","
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "REC = '" & RTrim(.REC) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "納入先工場 = '" & RTrim(.納入先工場) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "'"
        gSL_Select = gSL_Select & " Where 履歴管理番号 = " & .履歴管理番号 & ""
        gSL_Select = gSL_Select & "   And 受注出荷区分 = " & .受注出荷区分 & ""
        gSL_Select = gSL_Select & "   And ファイルID = '" & RTrim(.ファイルID) & "'"
        gSL_Select = gSL_Select & "   And SEQ = " & .SEQ & ""
        gSL_Select = gSL_Select & "   And 納入番号 = '" & RTrim(.納入番号) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    PHTUpdate = True
End Function

Attribute VB_Name = "帳票出力テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   帳票出力テーブル
'*
'*   モジュールID  : 帳票出力テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/01/21  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_帳票出力テーブル
        端末番号  As String
        帳票区分  As Integer
        管理文字  As String
        検索情報  As String
        管理NO    As Long
        管理行NO  As Integer
        補助番号  As Long
        出力先    As Integer
'        作成日    As Date
        件数          As Long
    End Type

    Public CHT As T_帳票出力テーブル

    Public CHT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromCHT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearCHT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With CHT
            Select Case adoCol.Name
                Case "端末番号":  If Not IsNull(adoCol.Value) Then .端末番号 = adoCol.Value
                Case "帳票区分":  If Not IsNull(adoCol.Value) Then .帳票区分 = adoCol.Value
                Case "管理文字":  If Not IsNull(adoCol.Value) Then .管理文字 = adoCol.Value
                Case "検索情報":  If Not IsNull(adoCol.Value) Then .検索情報 = adoCol.Value
                Case "管理NO":    If Not IsNull(adoCol.Value) Then .管理NO   = adoCol.Value
                Case "管理行NO":  If Not IsNull(adoCol.Value) Then .管理行NO = adoCol.Value
                Case "補助番号":  If Not IsNull(adoCol.Value) Then .補助番号 = adoCol.Value
                Case "出力先":    If Not IsNull(adoCol.Value) Then .出力先   = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function CHTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    CHTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 帳票出力テーブル"
        RTxtSql = RTxtSql &" Where 端末番号 = '" & RTrim(CHT.端末番号) & "'"
        RTxtSql = RTxtSql &" And   帳票区分 = " & CHT.帳票区分 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           CHTRead = True
           Call ItemsFromCHT(DBNo)
	   CHT_RDSTS = True
        else
	   CHT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call CHTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub CHTReadNext(Optional DBNo As Integer = 0)
    If Not CHT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           CHT_RDSTS = False
        Else
           Call ItemsFromCHT(DBNo)
           CHT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub CHTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        CHT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( CHT )
'*----------------------------------------------*'
Public Sub ItemsClearCHT()
    With CHT
	.端末番号 = ""
	.帳票区分 = 0
	.管理文字 = ""
	.検索情報 = ""
	.管理NO = 0
	.管理行NO = 0
	.補助番号 = 0
	.出力先 = 0
'	.作成日 = 0
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function CHTInsert(Optional DBNo As Integer) As Boolean
    CHTInsert = True
    With CHT
        gSL_Select = "Insert Into 帳票出力テーブル  ( "
        gSL_Select = gSL_Select & "端末番号,帳票区分,管理文字,"
        gSL_Select = gSL_Select & "検索情報,管理NO,管理行NO,"
        gSL_Select = gSL_Select & "補助番号,出力先,"
        gSL_Select = gSL_Select & "作成日"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.端末番号) & "',"		
        gSL_Select = gSL_Select & "" & .帳票区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.管理文字) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.検索情報) & "',"		
        gSL_Select = gSL_Select & "" & .管理NO & ","		
        gSL_Select = gSL_Select & "" & .管理行NO & ","		
        gSL_Select = gSL_Select & "" & .補助番号 & ","		
        gSL_Select = gSL_Select & "" & .出力先 & ","		
        gSL_Select = gSL_Select & "GETDATE()"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    CHTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function CHTUpdate(Optional DBNo As Integer) As Boolean
    CHTUpdate = False	
    With CHT
        gSL_Select = "Update 帳票出力テーブル  Set "
        gSL_Select = gSL_Select & "端末番号 = '" & RTrim(.端末番号) & "',"
        gSL_Select = gSL_Select & "帳票区分 = " & .帳票区分 & ","
        gSL_Select = gSL_Select & "管理文字 = '" & RTrim(.管理文字) & "',"
        gSL_Select = gSL_Select & "検索情報 = '" & RTrim(.検索情報) & "',"
        gSL_Select = gSL_Select & "管理NO = " & .管理NO & ","
        gSL_Select = gSL_Select & "管理行NO = " & .管理行NO & ","
        gSL_Select = gSL_Select & "補助番号 = " & .補助番号 & ","
        gSL_Select = gSL_Select & "出力先 = " & .出力先 & ""
        gSL_Select = gSL_Select & " Where 端末番号 = '" & RTrim(.端末番号) & "'"
        gSL_Select = gSL_Select & "   And 帳票区分 = " & .帳票区分 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    CHTUpdate = True
End Function

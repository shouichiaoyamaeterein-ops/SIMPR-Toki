Attribute VB_Name = "帳票可変項目テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   帳票可変項目テーブル
'*
'*   モジュールID  : 帳票可変項目テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*
'*   作成日  :  2003/12/08  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_帳票可変項目テーブル
        帳票番号  As Integer
        連番      As Integer
        帳票名    As String
        可変項目  As String
'        作成日    As Date
        作成者    As String
'        作成端末  As String
'        更新日    As Date
        更新者    As String
'        更新端末  As String
        件数          As Long
    End Type

    Public CKT As T_帳票可変項目テーブル

    Public CKT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromCKT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearCKT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With CKT
            Select Case adoCol.Name
                Case "帳票番号":  If Not IsNull(adoCol.Value) Then .帳票番号 = adoCol.Value
                Case "連番":      If Not IsNull(adoCol.Value) Then .連番     = adoCol.Value
                Case "帳票名":    If Not IsNull(adoCol.Value) Then .帳票名   = adoCol.Value
                Case "可変項目":  If Not IsNull(adoCol.Value) Then .可変項目 = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
                Case "作成者":    If Not IsNull(adoCol.Value) Then .作成者   = adoCol.Value
'                Case "作成端末":  If Not IsNull(adoCol.Value) Then .作成端末 = adoCol.Value
'                Case "更新日":    If Not IsNull(adoCol.Value) Then .更新日   = adoCol.Value
                Case "更新者":    If Not IsNull(adoCol.Value) Then .更新者   = adoCol.Value
'                Case "更新端末":  If Not IsNull(adoCol.Value) Then .更新端末 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function CKTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    CKTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 帳票可変項目テーブル"
        RTxtSql = RTxtSql &" Where 帳票番号 = " & CKT.帳票番号 & ""
        RTxtSql = RTxtSql &" And   連番 = " & CKT.連番 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           CKTRead = True
           Call ItemsFromCKT(DBNo)
	   CKT_RDSTS = True
        else
	   CKT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call CKTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub CKTReadNext(Optional DBNo As Integer = 0)
    If Not CKT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           CKT_RDSTS = False
        Else
           Call ItemsFromCKT(DBNo)
           CKT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub CKTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        CKT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( CKT )
'*----------------------------------------------*'
Public Sub ItemsClearCKT()
    With CKT
	.帳票番号 = 0
	.連番 = 0
	.帳票名 = ""
	.可変項目 = ""
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
Public Sub CKTInsert(Optional DBNo As Integer)
    With CKT
        gSL_Select = "Insert Into 帳票可変項目テーブル  ( "
        gSL_Select = gSL_Select & "帳票番号,連番,帳票名,"
        gSL_Select = gSL_Select & "可変項目,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .帳票番号 & ","		
        gSL_Select = gSL_Select & "" & .連番 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.帳票名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.可変項目) & "',"		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "',"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        .DBExec
    End With
End Sub
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Sub CKTUpdate(Optional DBNo As Integer)
    With CKT
        gSL_Select = "Update 帳票可変項目テーブル  Set "
        gSL_Select = gSL_Select & "帳票番号 = " & .帳票番号 & ","
        gSL_Select = gSL_Select & "連番 = " & .連番 & ","
        gSL_Select = gSL_Select & "帳票名 = '" & RTrim(.帳票名) & "',"
        gSL_Select = gSL_Select & "可変項目 = '" & RTrim(.可変項目) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 帳票番号 = " & .帳票番号 & ""
        gSL_Select = gSL_Select & "   And 連番 = " & .連番 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        .DBExec
    End With

End Sub

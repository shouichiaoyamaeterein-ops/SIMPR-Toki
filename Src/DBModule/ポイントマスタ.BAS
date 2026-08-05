Attribute VB_Name = "ポイントマスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   ポイントマスタ
'*
'*   モジュールID  : ポイントマスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2005/11/08  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_ポイントマスタ
        ポイント  As String
        イメージ  As String
'        作成日    As Date
        作成者    As String
'        作成端末  As String
        件数          As Long
    End Type

    Public POM As T_ポイントマスタ

    Public POM_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromPOM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearPOM	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With POM
            Select Case adoCol.Name
                Case "ポイント":  If Not IsNull(adoCol.Value) Then .ポイント = adoCol.Value
                Case "イメージ":  If Not IsNull(adoCol.Value) Then .イメージ = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
                Case "作成者":    If Not IsNull(adoCol.Value) Then .作成者   = adoCol.Value
'                Case "作成端末":  If Not IsNull(adoCol.Value) Then .作成端末 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function POMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    POMRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From ポイントマスタ"
        RTxtSql = RTxtSql &" Where ポイント = '" & RTrim(POM.ポイント) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           POMRead = True
           Call ItemsFromPOM(DBNo)
	   POM_RDSTS = True
        else
	   POM_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call POMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub POMReadNext(Optional DBNo As Integer = 0)
    If Not POM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           POM_RDSTS = False
        Else
           Call ItemsFromPOM(DBNo)
           POM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub POMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        POM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( POM )
'*----------------------------------------------*'
Public Sub ItemsClearPOM()
    With POM
	.ポイント = ""
	.イメージ = ""
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function POMInsert(Optional DBNo As Integer) As Boolean
    POMInsert = True
    With POM
        gSL_Select = "Insert Into ポイントマスタ  ( "
        gSL_Select = gSL_Select & "ポイント,イメージ,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.ポイント) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.イメージ) & "',"		
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
    POMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function POMUpdate(Optional DBNo As Integer) As Boolean
    POMUpdate = False	
    With POM
        gSL_Select = "Update ポイントマスタ  Set "
        gSL_Select = gSL_Select & "ポイント = '" & RTrim(.ポイント) & "',"
        gSL_Select = gSL_Select & "イメージ = '" & RTrim(.イメージ) & "'"
        gSL_Select = gSL_Select & " Where ポイント = '" & RTrim(.ポイント) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    POMUpdate = True
End Function

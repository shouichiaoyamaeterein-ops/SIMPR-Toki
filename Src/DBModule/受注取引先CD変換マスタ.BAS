Attribute VB_Name = "受注取引先CD変換マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   受注取引先CD変換マスタ
'*
'*   モジュールID  : 受注取引先CD変換マスタ.Bas
'*
'*
'*   作成日  :  2009/01/06  By Cis(FMV-E630)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_受注取引先CD変換マスタ
        変換元CD          As String
        変換元受入        As String
        変換先CD          As String
        変換先受入        As String
        契約先            As String
'        作成日            As Date
'        作成端末          As String
        件数          As Long
    End Type

    Public JTH As T_受注取引先CD変換マスタ

    Public JTH_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromJTH(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearJTH
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With JTH
            Select Case adoCol.Name
                Case "変換元CD":          If Not IsNull(adoCol.Value) Then .変換元CD = adoCol.Value
                Case "変換元受入":        If Not IsNull(adoCol.Value) Then .変換元受入 = adoCol.Value
                Case "変換先CD":          If Not IsNull(adoCol.Value) Then .変換先CD = adoCol.Value
                Case "変換先受入":        If Not IsNull(adoCol.Value) Then .変換先受入 = adoCol.Value
                Case "契約先":             If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function JTHRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    JTHRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 受注取引先CD変換マスタ"
        RTxtSql = RTxtSql & " Where 変換元CD = '" & RTrim(JTH.変換元CD) & "'"
        RTxtSql = RTxtSql & " And   変換元受入 = '" & RTrim(JTH.変換元受入) & "'"
        RTxtSql = RTxtSql & " And   契約先 = '" & RTrim(JTH.契約先) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           JTHRead = True
           Call ItemsFromJTH(DBNo)
           JTH_RDSTS = True
        Else
           JTH_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call JTHClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub JTHReadNext(Optional DBNo As Integer = 0)
    If Not JTH_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           JTH_RDSTS = False
        Else
           Call ItemsFromJTH(DBNo)
           JTH_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub JTHClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        JTH_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( JTH )
'*----------------------------------------------*'
Public Sub ItemsClearJTH()
    With JTH
        .変換元CD = ""
        .変換元受入 = ""
        .変換先CD = ""
        .変換先受入 = ""
        .契約先 = ""
'       .作成日 = 0
'       .作成端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function JTHInsert(Optional DBNo As Integer) As Boolean
    JTHInsert = False
    With JTH
        gSL_Select = "Insert Into 受注取引先CD変換マスタ  ( "
        gSL_Select = gSL_Select & "変換元CD,変換元受入,変換先CD,"
        gSL_Select = gSL_Select & "変換先受入,"
        gSL_Select = gSL_Select & "作成日,作成端末,契約先"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.変換元CD) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.変換元受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.変換先CD) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.変換先受入) & "',"
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)" & ","
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "'"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    JTHInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function JTHUpdate(Optional DBNo As Integer) As Boolean
    JTHUpdate = False
    With JTH
        gSL_Select = "Update 受注取引先CD変換マスタ  Set "
        gSL_Select = gSL_Select & "変換元CD = '" & RTrim(.変換元CD) & "',"
        gSL_Select = gSL_Select & "変換元受入 = '" & RTrim(.変換元受入) & "',"
        gSL_Select = gSL_Select & "変換先CD = '" & RTrim(.変換先CD) & "',"
        gSL_Select = gSL_Select & "変換先受入 = '" & RTrim(.変換先受入) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "'"
        gSL_Select = gSL_Select & " Where 変換元CD = '" & RTrim(.変換元CD) & "'"
        gSL_Select = gSL_Select & "   And 変換元受入 = '" & RTrim(.変換元受入) & "'"
        gSL_Select = gSL_Select & "   And 契約先 = '" & RTrim(.契約先) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    JTHUpdate = True
End Function

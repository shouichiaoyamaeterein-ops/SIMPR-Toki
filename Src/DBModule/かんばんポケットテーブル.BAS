Attribute VB_Name = "かんばんポケットテーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   かんばんポケットテーブル
'*
'*   モジュールID  : かんばんポケットテーブル.Bas
'*
'*
'*
'*   作成日  :  2003/12/05  By Cis
'*   変更日  :  2003/12/05  項目追加
'*---------------------------------------------------------------------------------
    Type T_かんばんポケットテーブル
        グループ          As Byte
        ポケットNO        As Byte
        取引先            As String
        受入              As String
'        作成日            As Date
        作成者            As String
'        作成端末          As String
'        更新日            As Date
        更新者            As String
'        更新端末          As String
        件数          As Long
    
        取引先名            As String
    End Type

    Public KPT As T_かんばんポケットテーブル

    Public KPT_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKPT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKPT
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With KPT
            Select Case adoCol.Name
                Case "グループ":          If Not IsNull(adoCol.Value) Then .グループ = adoCol.Value
                Case "ポケットNO":        If Not IsNull(adoCol.Value) Then .ポケットNO = adoCol.Value
                Case "取引先":            If Not IsNull(adoCol.Value) Then .取引先 = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
                Case "更新者":            If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
            
                Case "取引先名":          If Not IsNull(adoCol.Value) Then .取引先名 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KPTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KPTRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From かんばんポケットテーブル"
        RTxtSql = RTxtSql & " Where グループ = " & KPT.グループ & ""
        RTxtSql = RTxtSql & " And   ポケットNO = " & KPT.ポケットNO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KPTRead = True
           Call ItemsFromKPT(DBNo)
           KPT_RDSTS = True
        Else
           KPT_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call KPTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KPTReadNext(Optional DBNo As Integer = 0)
    If Not KPT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KPT_RDSTS = False
        Else
           Call ItemsFromKPT(DBNo)
           KPT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KPTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KPT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KPT )
'*----------------------------------------------*'
Public Sub ItemsClearKPT()
    With KPT
        .グループ = 0
        .ポケットNO = 0
        .取引先 = ""
        .受入 = ""
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
'       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
        .件数 = 0
    
        .取引先名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Sub KPTInsert(Optional DBNo As Integer)
    With KPT
        gSL_Select = "Insert Into かんばんポケットテーブル  ( "
        gSL_Select = gSL_Select & "グループ,ポケットNO,取引先,"
        gSL_Select = gSL_Select & "受入,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .グループ & ","
        gSL_Select = gSL_Select & "" & .ポケットNO & ","
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
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
Public Sub KPTUpdate(Optional DBNo As Integer)
    With KPT
        gSL_Select = "Update かんばんポケットテーブル  Set "
        gSL_Select = gSL_Select & "グループ = " & .グループ & ","
        gSL_Select = gSL_Select & "ポケットNO = " & .ポケットNO & ","
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where グループ = " & .グループ & ""
        gSL_Select = gSL_Select & "   And ポケットNO = " & .ポケットNO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        .DBExec
    End With

End Sub

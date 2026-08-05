Attribute VB_Name = "構成メンテワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   構成メンテワーク
'*
'*   モジュールID  : 構成メンテワーク.Bas
'*
'*
'*   作成日  :  2004/01/08  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_構成メンテワーク
        処理端末          As String
        順位              As String
        レベル            As Byte
        親品番            As String
        子品番            As String
        構成数            As Currency
        展開区分          As Byte
        LT               As Byte
        適用開始日        As String
        適用終了日        As String
        件数          As Long
        
        品名              As String
        車型              As String
        社内背番          As String
        品番編集          As String
    End Type

    Public KOW As T_構成メンテワーク

    Public KOW_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKOW(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKOW
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With KOW
            Select Case adoCol.Name
                Case "処理端末":          If Not IsNull(adoCol.Value) Then .処理端末 = adoCol.Value
                Case "順位":              If Not IsNull(adoCol.Value) Then .順位 = adoCol.Value
                Case "レベル":            If Not IsNull(adoCol.Value) Then .レベル = adoCol.Value
                Case "親品番":            If Not IsNull(adoCol.Value) Then .親品番 = adoCol.Value
                Case "子品番":            If Not IsNull(adoCol.Value) Then .子品番 = adoCol.Value
                Case "構成数":            If Not IsNull(adoCol.Value) Then .構成数 = adoCol.Value
                Case "展開区分":          If Not IsNull(adoCol.Value) Then .展開区分 = adoCol.Value
                Case "LT":               If Not IsNull(adoCol.Value) Then .LT = adoCol.Value
                Case "適用開始日":        If Not IsNull(adoCol.Value) Then .適用開始日 = adoCol.Value
                Case "適用終了日":        If Not IsNull(adoCol.Value) Then .適用終了日 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
            
                Case "品名":        If Not IsNull(adoCol.Value) Then .品名 = adoCol.Value
                Case "車型":        If Not IsNull(adoCol.Value) Then .車型 = adoCol.Value
                Case "社内背番":        If Not IsNull(adoCol.Value) Then .社内背番 = adoCol.Value
                Case "品番編集":    If Not IsNull(adoCol.Value) Then .品番編集 = adoCol.Value
            
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KOWRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KOWRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 構成メンテワーク"
        RTxtSql = RTxtSql & " Where 処理端末 = '" & RTrim(KOW.処理端末) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KOWRead = True
           Call ItemsFromKOW(DBNo)
           KOW_RDSTS = True
        Else
           KOW_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call KOWClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KOWReadNext(Optional DBNo As Integer = 0)
    If Not KOW_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KOW_RDSTS = False
        Else
           Call ItemsFromKOW(DBNo)
           KOW_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KOWClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KOW_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KOW )
'*----------------------------------------------*'
Public Sub ItemsClearKOW()
    With KOW
        .処理端末 = ""
        .順位 = ""
        .レベル = 0
        .親品番 = ""
        .子品番 = ""
        .構成数 = 0
        .展開区分 = 0
        .LT = 0
        .適用開始日 = ""
        .適用終了日 = ""
        .件数 = 0
    
        .品名 = ""
        .車型 = ""
        .社内背番 = ""
        .品番編集 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KOWInsert(Optional DBNo As Integer) As Boolean
    KOWInsert = True
    With KOW
        gSL_Select = "Insert Into 構成メンテワーク  ( "
        gSL_Select = gSL_Select & "処理端末,順位,レベル,"
        gSL_Select = gSL_Select & "親品番,子品番,構成数,"
        gSL_Select = gSL_Select & "展開区分,適用開始日,適用終了日"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.順位) & "',"
        gSL_Select = gSL_Select & "" & .レベル & ","
        gSL_Select = gSL_Select & "'" & RTrim(.親品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.子品番) & "',"
        gSL_Select = gSL_Select & "" & .構成数 & ","
        gSL_Select = gSL_Select & "" & .展開区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.適用開始日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.適用終了日) & "'"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KOWInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KOWUpdate(Optional DBNo As Integer) As Boolean
    KOWUpdate = False
    With KOW
        gSL_Select = "Update 構成メンテワーク  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "順位 = '" & RTrim(.順位) & "',"
        gSL_Select = gSL_Select & "レベル = " & .レベル & ","
        gSL_Select = gSL_Select & "親品番 = '" & RTrim(.親品番) & "',"
        gSL_Select = gSL_Select & "子品番 = '" & RTrim(.子品番) & "',"
        gSL_Select = gSL_Select & "構成数 = " & .構成数 & ","
        gSL_Select = gSL_Select & "展開区分 = " & .展開区分 & ","
        gSL_Select = gSL_Select & "適用開始日 = '" & RTrim(.適用開始日) & "',"
        gSL_Select = gSL_Select & "適用終了日 = '" & RTrim(.適用終了日) & "'"
        gSL_Select = gSL_Select & " Where 処理端末 = '" & RTrim(.処理端末) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KOWUpdate = True
End Function

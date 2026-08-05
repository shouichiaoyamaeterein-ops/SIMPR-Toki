Attribute VB_Name = "構成マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   構成マスタ
'*
'*   モジュールID  : 構成マスタ.Bas
'*
'*
'*   作成日  :  2004/01/08  By Cis
'*   変更日  :  2004/01/17          変更前親品番・子品番を項目追加(更新処理時に使用)
'*   変更日  :  2004/01/19          構成マスタ更新２(親品番・置換え)
'*---------------------------------------------------------------------------------
    Type T_構成マスタ
        親品番            As String
        子品番            As String
        構成数            As Currency
        展開区分          As Byte
        LT                As Byte
        適用開始日        As String
        適用終了日        As String
        レベル            As Byte
'        作成日            As Date
        作成者            As String
'        作成端末          As String
'        更新日            As Date
        更新者            As String
'        更新端末          As String
        件数          As Long
    
        変更前親品番        As String
        変更前子品番        As String
    End Type

    Public KOM As T_構成マスタ

    Public KOM_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKOM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKOM
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With KOM
            Select Case adoCol.Name
                Case "親品番":            If Not IsNull(adoCol.Value) Then .親品番 = adoCol.Value
                Case "子品番":            If Not IsNull(adoCol.Value) Then .子品番 = adoCol.Value
                Case "構成数":            If Not IsNull(adoCol.Value) Then .構成数 = adoCol.Value
                Case "展開区分":          If Not IsNull(adoCol.Value) Then .展開区分 = adoCol.Value
                Case "LT":          If Not IsNull(adoCol.Value) Then .LT = adoCol.Value
                Case "適用開始日":        If Not IsNull(adoCol.Value) Then .適用開始日 = adoCol.Value
                Case "適用終了日":        If Not IsNull(adoCol.Value) Then .適用終了日 = adoCol.Value
                Case "レベル":            If Not IsNull(adoCol.Value) Then .レベル = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
                Case "更新者":            If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
            
                Case "変更前親品番":      If Not IsNull(adoCol.Value) Then .変更前親品番 = adoCol.Value
                Case "変更前子品番":      If Not IsNull(adoCol.Value) Then .変更前子品番 = adoCol.Value
            
            
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KOMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KOMRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 構成マスタ"
        RTxtSql = RTxtSql & " Where 親品番 = '" & RTrim(KOM.親品番) & "'"
        RTxtSql = RTxtSql & " And   子品番 = '" & RTrim(KOM.子品番) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KOMRead = True
           Call ItemsFromKOM(DBNo)
           KOM_RDSTS = True
        Else
           KOM_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call KOMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KOMReadNext(Optional DBNo As Integer = 0)
    If Not KOM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KOM_RDSTS = False
        Else
           Call ItemsFromKOM(DBNo)
           KOM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KOMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KOM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KOM )
'*----------------------------------------------*'
Public Sub ItemsClearKOM()
    With KOM
        .親品番 = ""
        .子品番 = ""
        .構成数 = 0
        .展開区分 = 0
        .LT = 0
        .適用開始日 = ""
        .適用終了日 = ""
        .レベル = 0
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
'       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
        .件数 = 0
    
        .変更前親品番 = ""
        .変更前子品番 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KOMInsert(Optional DBNo As Integer) As Boolean
    KOMInsert = True
    With KOM
        gSL_Select = "Insert Into 構成マスタ  ( "
        gSL_Select = gSL_Select & "親品番,子品番,構成数,"
        gSL_Select = gSL_Select & "展開区分,LT,適用開始日,適用終了日,"
        gSL_Select = gSL_Select & "レベル,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.親品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.子品番) & "',"
        gSL_Select = gSL_Select & "" & .構成数 & ","
        gSL_Select = gSL_Select & "" & .展開区分 & ","
        gSL_Select = gSL_Select & "" & .LT & ","
        gSL_Select = gSL_Select & "'" & RTrim(.適用開始日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.適用終了日) & "',"
        gSL_Select = gSL_Select & "0,"
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
    KOMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KOMUpdate(Optional DBNo As Integer) As Boolean
    KOMUpdate = False
    With KOM
        gSL_Select = "Update 構成マスタ  Set "
        gSL_Select = gSL_Select & "親品番 = '" & RTrim(.親品番) & "',"
        gSL_Select = gSL_Select & "子品番 = '" & RTrim(.子品番) & "',"
        gSL_Select = gSL_Select & "構成数 = " & .構成数 & ","
        gSL_Select = gSL_Select & "展開区分 = " & .展開区分 & ","
        gSL_Select = gSL_Select & "LT = " & .LT & ","
        gSL_Select = gSL_Select & "適用開始日 = '" & RTrim(.適用開始日) & "',"
        gSL_Select = gSL_Select & "適用終了日 = '" & RTrim(.適用終了日) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 親品番 = '" & RTrim(.変更前親品番) & "'"
        gSL_Select = gSL_Select & "   And 子品番 = '" & RTrim(.変更前子品番) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KOMUpdate = True
End Function
'*--------------------*'
'*      更新処理2      *
'*--------------------*'
Public Function KOMUpdate2(Optional DBNo As Integer) As Boolean
    KOMUpdate2 = False
    With KOM
        gSL_Select = "Update 構成マスタ  Set "
        gSL_Select = gSL_Select & "親品番 = '" & RTrim(.親品番) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 親品番 = '" & RTrim(.変更前親品番) & "'"
        gSL_Select = gSL_Select & "   And 子品番 = '" & RTrim(.変更前子品番) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KOMUpdate2 = True
End Function


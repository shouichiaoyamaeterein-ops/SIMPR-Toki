Attribute VB_Name = "発注サイクル設定マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   発注サイクル設定マスタ
'*
'*   モジュールID  : 発注サイクル設定マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*
'*   作成日  :  2005/02/28  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_発注サイクル設定マスタ
        回数              As Byte
        昼夜区分          As Byte
        回収時間          As String
        発注時間          As String
        引取時間          As String
        備考              As String
        最大回数          As Byte
        最終発注日        As String
        最終回            As Byte
'        作成日            As Date
        作成者            As String
'        作成端末          As String
'        更新日            As Date
        更新者            As String
'        更新端末          As String
        件数          As Long
    End Type

    Public HCS As T_発注サイクル設定マスタ

    Public HCS_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHCS(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHCS
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With HCS
            Select Case adoCol.Name
                Case "回数":              If Not IsNull(adoCol.Value) Then .回数 = adoCol.Value
                Case "昼夜区分":          If Not IsNull(adoCol.Value) Then .昼夜区分 = adoCol.Value
                Case "回収時間":          If Not IsNull(adoCol.Value) Then .回収時間 = adoCol.Value
                Case "発注時間":          If Not IsNull(adoCol.Value) Then .発注時間 = adoCol.Value
                Case "引取時間":          If Not IsNull(adoCol.Value) Then .引取時間 = adoCol.Value
                Case "備考":              If Not IsNull(adoCol.Value) Then .備考 = adoCol.Value
                Case "最大回数":          If Not IsNull(adoCol.Value) Then .最大回数 = adoCol.Value
                Case "最終発注日":        If Not IsNull(adoCol.Value) Then .最終発注日 = adoCol.Value
                Case "最終回":            If Not IsNull(adoCol.Value) Then .最終回 = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
                Case "更新者":            If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function HCSRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HCSRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 発注サイクル設定マスタ"
        RTxtSql = RTxtSql & " Where 回数 = " & HCS.回数 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HCSRead = True
           Call ItemsFromHCS(DBNo)
           HCS_RDSTS = True
        Else
           HCS_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call HCSClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HCSReadNext(Optional DBNo As Integer = 0)
    If Not HCS_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HCS_RDSTS = False
        Else
           Call ItemsFromHCS(DBNo)
           HCS_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HCSClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HCS_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HCS )
'*----------------------------------------------*'
Public Sub ItemsClearHCS()
    With HCS
        .回数 = 0
        .昼夜区分 = 0
        .回収時間 = ""
        .発注時間 = ""
        .引取時間 = ""
        .備考 = ""
        .最大回数 = 0
        .最終発注日 = ""
        .最終回 = 0
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
'       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function HCSInsert(Optional DBNo As Integer) As Boolean
    HCSInsert = True
    With HCS
        gSL_Select = "Insert Into 発注サイクル設定マスタ  ( "
        gSL_Select = gSL_Select & "回数,昼夜区分,回収時間,"
        gSL_Select = gSL_Select & "発注時間,引取時間,備考,"
        gSL_Select = gSL_Select & "最大回数,最終発注日,最終回,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .回数 & ","
        gSL_Select = gSL_Select & "" & .昼夜区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.回収時間) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発注時間) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.引取時間) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.備考) & "',"
        gSL_Select = gSL_Select & "" & .最大回数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.最終発注日) & "',"
        gSL_Select = gSL_Select & "" & .最終回 & ","
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
    HCSInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function HCSUpdate(Optional DBNo As Integer) As Boolean
    HCSUpdate = False
    With HCS
        gSL_Select = "Update 発注サイクル設定マスタ  Set "
        gSL_Select = gSL_Select & "回数 = " & .回数 & ","
        gSL_Select = gSL_Select & "昼夜区分 = " & .昼夜区分 & ","
        gSL_Select = gSL_Select & "回収時間 = '" & RTrim(.回収時間) & "',"
        gSL_Select = gSL_Select & "発注時間 = '" & RTrim(.発注時間) & "',"
        gSL_Select = gSL_Select & "引取時間 = '" & RTrim(.引取時間) & "',"
        gSL_Select = gSL_Select & "備考 = '" & RTrim(.備考) & "',"
        gSL_Select = gSL_Select & "最大回数 = " & .最大回数 & ","
        gSL_Select = gSL_Select & "最終発注日 = '" & RTrim(.最終発注日) & "',"
        gSL_Select = gSL_Select & "最終回 = " & .最終回 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 回数 = " & .回数 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HCSUpdate = True
End Function

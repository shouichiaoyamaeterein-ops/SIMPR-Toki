Attribute VB_Name = "TEXT設定テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TEXT設定テーブル
'*
'*   モジュールID  : TEXT設定テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*
'*   作成日  :  2003/12/13  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_TEXT設定テーブル
        レイアウト番号    As Byte
        区分              As Byte
        連番              As Byte
        レコード番号      As Long
        レコード定義件数  As Byte
        開始レコード      As Byte
        最終レコード      As Byte
        項目              As String
        開始位置          As Long
        長さ              As Long
        データ項目        As String
        データタイプ      As Byte
        出力長さ          As Long
        少数桁数          As Byte
'        作成日            As Date
        作成者            As String
'        作成端末          As String
'        更新日            As Date
        更新者            As String
'        更新端末          As String
        件数          As Long
    End Type

    Public TXS As T_TEXT設定テーブル

    Public TXS_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTXS(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTXS
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With TXS
            Select Case adoCol.Name
                Case "レイアウト番号":    If Not IsNull(adoCol.Value) Then .レイアウト番号 = adoCol.Value
                Case "区分":              If Not IsNull(adoCol.Value) Then .区分 = adoCol.Value
                Case "連番":              If Not IsNull(adoCol.Value) Then .連番 = adoCol.Value
                Case "レコード番号":      If Not IsNull(adoCol.Value) Then .レコード番号 = adoCol.Value
                Case "レコード定義件数":  If Not IsNull(adoCol.Value) Then .レコード定義件数 = adoCol.Value
                Case "開始レコード":      If Not IsNull(adoCol.Value) Then .開始レコード = adoCol.Value
                Case "最終レコード":      If Not IsNull(adoCol.Value) Then .最終レコード = adoCol.Value
                Case "項目":              If Not IsNull(adoCol.Value) Then .項目 = adoCol.Value
                Case "開始位置":          If Not IsNull(adoCol.Value) Then .開始位置 = adoCol.Value
                Case "長さ":              If Not IsNull(adoCol.Value) Then .長さ = adoCol.Value
                Case "データ項目":        If Not IsNull(adoCol.Value) Then .データ項目 = adoCol.Value
                Case "データタイプ":      If Not IsNull(adoCol.Value) Then .データタイプ = adoCol.Value
                Case "出力長さ":          If Not IsNull(adoCol.Value) Then .出力長さ = adoCol.Value
                Case "少数桁数":          If Not IsNull(adoCol.Value) Then .少数桁数 = adoCol.Value
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
Public Function TXSRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TXSRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From TEXT設定テーブル"
        RTxtSql = RTxtSql & " Where レイアウト番号 = " & TXS.レイアウト番号 & ""
        RTxtSql = RTxtSql & " And   区分 = " & TXS.区分 & ""
        RTxtSql = RTxtSql & " And   連番 = " & TXS.連番 & ""
        RTxtSql = RTxtSql & " And   レコード番号 = " & TXS.レコード番号 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TXSRead = True
           Call ItemsFromTXS(DBNo)
           TXS_RDSTS = True
        Else
           TXS_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call TXSClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TXSReadNext(Optional DBNo As Integer = 0)
    If Not TXS_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TXS_RDSTS = False
        Else
           Call ItemsFromTXS(DBNo)
           TXS_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TXSClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TXS_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TXS )
'*----------------------------------------------*'
Public Sub ItemsClearTXS()
    With TXS
        .レイアウト番号 = 0
        .区分 = 0
        .連番 = 0
        .レコード番号 = 0
        .レコード定義件数 = 0
        .開始レコード = 0
        .最終レコード = 0
        .項目 = ""
        .開始位置 = 0
        .長さ = 0
        .データ項目 = ""
        .データタイプ = 0
        .出力長さ = 0
        .少数桁数 = 0
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
Public Sub TXSInsert(Optional DBNo As Integer)
    With TXS
        gSL_Select = "Insert Into TEXT設定テーブル  ( "
        gSL_Select = gSL_Select & "レイアウト番号,区分,連番,"
        gSL_Select = gSL_Select & "レコード番号,レコード定義件数,開始レコード,"
        gSL_Select = gSL_Select & "最終レコード,項目,開始位置,"
        gSL_Select = gSL_Select & "長さ,データ項目,データタイプ,"
        gSL_Select = gSL_Select & "出力長さ,少数桁数,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .レイアウト番号 & ","
        gSL_Select = gSL_Select & "" & .区分 & ","
        gSL_Select = gSL_Select & "" & .連番 & ","
        gSL_Select = gSL_Select & "" & .レコード番号 & ","
        gSL_Select = gSL_Select & "" & .レコード定義件数 & ","
        gSL_Select = gSL_Select & "" & .開始レコード & ","
        gSL_Select = gSL_Select & "" & .最終レコード & ","
        gSL_Select = gSL_Select & "'" & RTrim(.項目) & "',"
        gSL_Select = gSL_Select & "" & .開始位置 & ","
        gSL_Select = gSL_Select & "" & .長さ & ","
        gSL_Select = gSL_Select & "'" & RTrim(.データ項目) & "',"
        gSL_Select = gSL_Select & "" & .データタイプ & ","
        gSL_Select = gSL_Select & "" & .出力長さ & ","
        gSL_Select = gSL_Select & "" & .少数桁数 & ","
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
Public Sub TXSUpdate(Optional DBNo As Integer)
    With TXS
        gSL_Select = "Update TEXT設定テーブル  Set "
        gSL_Select = gSL_Select & "レイアウト番号 = " & .レイアウト番号 & ","
        gSL_Select = gSL_Select & "区分 = " & .区分 & ","
        gSL_Select = gSL_Select & "連番 = " & .連番 & ","
        gSL_Select = gSL_Select & "レコード番号 = " & .レコード番号 & ","
        gSL_Select = gSL_Select & "レコード定義件数 = " & .レコード定義件数 & ","
        gSL_Select = gSL_Select & "開始レコード = " & .開始レコード & ","
        gSL_Select = gSL_Select & "最終レコード = " & .最終レコード & ","
        gSL_Select = gSL_Select & "項目 = '" & RTrim(.項目) & "',"
        gSL_Select = gSL_Select & "開始位置 = " & .開始位置 & ","
        gSL_Select = gSL_Select & "長さ = " & .長さ & ","
        gSL_Select = gSL_Select & "データ項目 = '" & RTrim(.データ項目) & "',"
        gSL_Select = gSL_Select & "データタイプ = " & .データタイプ & ","
        gSL_Select = gSL_Select & "出力長さ = " & .出力長さ & ","
        gSL_Select = gSL_Select & "少数桁数 = " & .少数桁数 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where レイアウト番号 = " & .レイアウト番号 & ""
        gSL_Select = gSL_Select & "   And 区分 = " & .区分 & ""
        gSL_Select = gSL_Select & "   And 連番 = " & .連番 & ""
        gSL_Select = gSL_Select & "   And レコード番号 = " & .レコード番号 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        .DBExec
    End With

End Sub

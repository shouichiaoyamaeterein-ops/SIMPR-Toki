Attribute VB_Name = "発注サイクルマスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   発注サイクルマスタ
'*
'*   モジュールID  : 発注サイクルマスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*                                      /項目追加含む/
'*   作成日  :  2005/03/03  By Cis
'*   作成日  :  2005/11/17  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_発注サイクルマスタ
        仕入先        As String
        サイクル      As String
        受入          As String
        振出便        As String
        行識別        As Byte
        昼夜区分      As Byte
        振出時間      As String
        発行時間      As String
        回収時間      As String
        入荷時間      As String
        振出開始      As String
        振出終了      As String
        振出加算日    As Integer
        発注加算日    As Integer
        発注便        As String
        発行加算日    As Integer
        シーケンス便  As String
        最終発注日    As String
        最終発注便    As String
        発注処理回数  As Byte
        最終便        As Byte
'        作成日        As Date
        作成者        As String
'        作成端末      As String
'        更新日        As Date
        更新者        As String
'        更新端末      As String
        件数          As Long
'   項目追加(自動)
        仕入先名       As String
        取引先         As String
        工場CD         As String
        編集サイクル   As String
        曜日発注       As Byte
    End Type

    Public HCM As T_発注サイクルマスタ

    Public HCM_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHCM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHCM
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With HCM
            Select Case adoCol.Name
                Case "仕入先":        If Not IsNull(adoCol.Value) Then .仕入先 = adoCol.Value
                Case "サイクル":      If Not IsNull(adoCol.Value) Then .サイクル = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "振出便":        If Not IsNull(adoCol.Value) Then .振出便 = adoCol.Value
                Case "行識別":        If Not IsNull(adoCol.Value) Then .行識別 = adoCol.Value
                Case "昼夜区分":      If Not IsNull(adoCol.Value) Then .昼夜区分 = adoCol.Value
                Case "振出時間":      If Not IsNull(adoCol.Value) Then .振出時間 = adoCol.Value
                Case "発行時間":      If Not IsNull(adoCol.Value) Then .発行時間 = adoCol.Value
                Case "回収時間":      If Not IsNull(adoCol.Value) Then .回収時間 = adoCol.Value
                Case "入荷時間":      If Not IsNull(adoCol.Value) Then .入荷時間 = adoCol.Value
                Case "振出開始":      If Not IsNull(adoCol.Value) Then .振出開始 = adoCol.Value
                Case "振出終了":      If Not IsNull(adoCol.Value) Then .振出終了 = adoCol.Value
                Case "振出加算日":    If Not IsNull(adoCol.Value) Then .振出加算日 = adoCol.Value
                Case "発注加算日":    If Not IsNull(adoCol.Value) Then .発注加算日 = adoCol.Value
                Case "発注便":        If Not IsNull(adoCol.Value) Then .発注便 = adoCol.Value
                Case "発行加算日":    If Not IsNull(adoCol.Value) Then .発行加算日 = adoCol.Value
                Case "シーケンス便":  If Not IsNull(adoCol.Value) Then .シーケンス便 = adoCol.Value
                Case "最終発注日":    If Not IsNull(adoCol.Value) Then .最終発注日 = adoCol.Value
                Case "最終発注便":    If Not IsNull(adoCol.Value) Then .最終発注便 = adoCol.Value
                Case "発注処理回数":  If Not IsNull(adoCol.Value) Then .発注処理回数 = adoCol.Value
                Case "最終便":        If Not IsNull(adoCol.Value) Then .最終便 = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "仕入先名":        If Not IsNull(adoCol.Value) Then .仕入先名 = adoCol.Value
                Case "取引先":  If Not IsNull(adoCol.Value) Then .取引先 = adoCol.Value
                Case "工場CD":  If Not IsNull(adoCol.Value) Then .工場CD = adoCol.Value
                Case "編集サイクル":    If Not IsNull(adoCol.Value) Then .編集サイクル = adoCol.Value
                Case "曜日発注":        If Not IsNull(adoCol.Value) Then .曜日発注 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function HCMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HCMRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 発注サイクルマスタ"
        RTxtSql = RTxtSql & " Where 仕入先 = '" & RTrim(HCM.仕入先) & "'"
        RTxtSql = RTxtSql & " And   サイクル = '" & RTrim(HCM.サイクル) & "'"
        RTxtSql = RTxtSql & " And   受入 = '" & RTrim(HCM.受入) & "'"
        RTxtSql = RTxtSql & " And   振出便 = '" & RTrim(HCM.振出便) & "'"
        RTxtSql = RTxtSql & " And   行識別 = " & HCM.行識別 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HCMRead = True
           Call ItemsFromHCM(DBNo)
           HCM_RDSTS = True
        Else
           HCM_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call HCMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HCMReadNext(Optional DBNo As Integer = 0)
    If Not HCM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HCM_RDSTS = False
        Else
           Call ItemsFromHCM(DBNo)
           HCM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HCMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HCM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HCM )
'*----------------------------------------------*'
Public Sub ItemsClearHCM()
    With HCM
        .仕入先 = ""
        .サイクル = ""
        .受入 = ""
        .振出便 = ""
        .行識別 = 0
        .昼夜区分 = 0
        .振出時間 = ""
        .発行時間 = ""
        .回収時間 = ""
        .入荷時間 = ""
        .振出開始 = ""
        .振出終了 = ""
        .振出加算日 = 0
        .発注加算日 = 0
        .発注便 = ""
        .発行加算日 = 0
        .シーケンス便 = ""
        .最終発注日 = ""
        .最終発注便 = ""
        .発注処理回数 = 0
        .最終便 = 0
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
'       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
        .件数 = 0
'   項目追加(自動)
        .仕入先名 = ""
        .取引先 = ""
        .工場CD = ""
        .編集サイクル = ""
        .曜日発注 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function HCMInsert(Optional DBNo As Integer) As Boolean
    HCMInsert = True
    With HCM
        gSL_Select = "Insert Into 発注サイクルマスタ  ( "
        gSL_Select = gSL_Select & "仕入先,サイクル,受入,"
        gSL_Select = gSL_Select & "振出便,行識別,昼夜区分,"
        gSL_Select = gSL_Select & "振出時間,発行時間,回収時間,"
        gSL_Select = gSL_Select & "入荷時間,振出開始,振出終了,"
        gSL_Select = gSL_Select & "振出加算日,発注加算日,発注便,"
        gSL_Select = gSL_Select & "発行加算日,シーケンス便,最終発注日,"
        gSL_Select = gSL_Select & "最終発注便,発注処理回数,最終便,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.振出便) & "',"
        gSL_Select = gSL_Select & "" & .行識別 & ","
        gSL_Select = gSL_Select & "" & .昼夜区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.振出時間) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発行時間) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.回収時間) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.入荷時間) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.振出開始) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.振出終了) & "',"
        gSL_Select = gSL_Select & "" & .振出加算日 & ","
        gSL_Select = gSL_Select & "" & .発注加算日 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.発注便) & "',"
        gSL_Select = gSL_Select & "" & .発行加算日 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.シーケンス便) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.最終発注日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.最終発注便) & "',"
        gSL_Select = gSL_Select & "" & .発注処理回数 & ","
        gSL_Select = gSL_Select & "" & .最終便 & ","
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
    HCMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function HCMUpdate(Optional DBNo As Integer) As Boolean
    HCMUpdate = False
    With HCM
        gSL_Select = "Update 発注サイクルマスタ  Set "
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "サイクル = '" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "振出便 = '" & RTrim(.振出便) & "',"
        gSL_Select = gSL_Select & "行識別 = " & .行識別 & ","
        gSL_Select = gSL_Select & "昼夜区分 = " & .昼夜区分 & ","
        gSL_Select = gSL_Select & "振出時間 = '" & RTrim(.振出時間) & "',"
        gSL_Select = gSL_Select & "発行時間 = '" & RTrim(.発行時間) & "',"
        gSL_Select = gSL_Select & "回収時間 = '" & RTrim(.回収時間) & "',"
        gSL_Select = gSL_Select & "入荷時間 = '" & RTrim(.入荷時間) & "',"
        gSL_Select = gSL_Select & "振出開始 = '" & RTrim(.振出開始) & "',"
        gSL_Select = gSL_Select & "振出終了 = '" & RTrim(.振出終了) & "',"
        gSL_Select = gSL_Select & "振出加算日 = " & .振出加算日 & ","
        gSL_Select = gSL_Select & "発注加算日 = " & .発注加算日 & ","
        gSL_Select = gSL_Select & "発注便 = '" & RTrim(.発注便) & "',"
        gSL_Select = gSL_Select & "発行加算日 = " & .発行加算日 & ","
        gSL_Select = gSL_Select & "シーケンス便 = '" & RTrim(.シーケンス便) & "',"
        gSL_Select = gSL_Select & "最終発注日 = '" & RTrim(.最終発注日) & "',"
        gSL_Select = gSL_Select & "最終発注便 = '" & RTrim(.最終発注便) & "',"
        gSL_Select = gSL_Select & "発注処理回数 = " & .発注処理回数 & ","
        gSL_Select = gSL_Select & "最終便 = " & .最終便 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 仕入先 = '" & RTrim(.仕入先) & "'"
        gSL_Select = gSL_Select & "   And サイクル = '" & RTrim(.サイクル) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 振出便 = '" & RTrim(.振出便) & "'"
        gSL_Select = gSL_Select & "   And 行識別 = " & .行識別 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HCMUpdate = True
End Function

Attribute VB_Name = "発注画面入力ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   発注画面入力ワーク
'*
'*   モジュールID  : 発注画面入力ワーク.Bas
'*
'*
'*
'*   作成日  :  2004/01/22  By Cis
'*   変更日  :  2004/01/22  処理端末=host_name()
'*---------------------------------------------------------------------------------
    Type T_発注画面入力ワーク
        処理端末  As String
        発注区分  As String
        仕入先    As String
        受入      As String
        納入日    As String
        納入便    As String
        発注日    As String
        発注便    As String
        品番      As String
        背番号    As String
        手配区分  As String
        SK区分    As Byte
        サイクル  As String
        契約先    As String
        収容数    As Long
        枚数      As Long
        端数      As Long
        発注数    As Long
        端数区分  As Byte
        打切区分  As Byte
        打切種別  As Byte
        打切残数  As Long
        打切日    As String
        作成種別  As String
        処理区分  As Byte
'        作成日    As Date
        作成者    As String
        件数          As Long
    End Type

    Public HGW As T_発注画面入力ワーク

    Public HGW_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHGW(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHGW
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With HGW
            Select Case adoCol.Name
                Case "処理端末":  If Not IsNull(adoCol.Value) Then .処理端末 = adoCol.Value
                Case "発注区分":  If Not IsNull(adoCol.Value) Then .発注区分 = adoCol.Value
                Case "仕入先":    If Not IsNull(adoCol.Value) Then .仕入先 = adoCol.Value
                Case "受入":      If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "納入日":    If Not IsNull(adoCol.Value) Then .納入日 = adoCol.Value
                Case "納入便":    If Not IsNull(adoCol.Value) Then .納入便 = adoCol.Value
                Case "発注日":    If Not IsNull(adoCol.Value) Then .発注日 = adoCol.Value
                Case "発注便":    If Not IsNull(adoCol.Value) Then .発注便 = adoCol.Value
                Case "品番":      If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "背番号":    If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "手配区分":  If Not IsNull(adoCol.Value) Then .手配区分 = adoCol.Value
                Case "SK区分":    If Not IsNull(adoCol.Value) Then .SK区分 = adoCol.Value
                Case "サイクル":  If Not IsNull(adoCol.Value) Then .サイクル = adoCol.Value
                Case "契約先":    If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
                Case "収容数":    If Not IsNull(adoCol.Value) Then .収容数 = adoCol.Value
                Case "枚数":      If Not IsNull(adoCol.Value) Then .枚数 = adoCol.Value
                Case "端数":      If Not IsNull(adoCol.Value) Then .端数 = adoCol.Value
                Case "発注数":    If Not IsNull(adoCol.Value) Then .発注数 = adoCol.Value
                Case "端数区分":  If Not IsNull(adoCol.Value) Then .端数区分 = adoCol.Value
                Case "打切区分":  If Not IsNull(adoCol.Value) Then .打切区分 = adoCol.Value
                Case "打切種別":  If Not IsNull(adoCol.Value) Then .打切種別 = adoCol.Value
                Case "打切残数":  If Not IsNull(adoCol.Value) Then .打切残数 = adoCol.Value
                Case "打切日":    If Not IsNull(adoCol.Value) Then .打切日 = adoCol.Value
                Case "作成種別":  If Not IsNull(adoCol.Value) Then .作成種別 = adoCol.Value
                Case "処理区分":  If Not IsNull(adoCol.Value) Then .処理区分 = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
                Case "作成者":    If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function HGWRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HGWRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 発注画面入力ワーク"
        RTxtSql = RTxtSql & " Where 処理端末 = '" & RTrim(HGW.処理端末) & "'"
        RTxtSql = RTxtSql & " And   発注区分 = '" & RTrim(HGW.発注区分) & "'"
        RTxtSql = RTxtSql & " And   仕入先 = '" & RTrim(HGW.仕入先) & "'"
        RTxtSql = RTxtSql & " And   受入 = '" & RTrim(HGW.受入) & "'"
        RTxtSql = RTxtSql & " And   納入日 = '" & RTrim(HGW.納入日) & "'"
        RTxtSql = RTxtSql & " And   納入便 = '" & RTrim(HGW.納入便) & "'"
        RTxtSql = RTxtSql & " And   発注日 = '" & RTrim(HGW.発注日) & "'"
        RTxtSql = RTxtSql & " And   発注便 = '" & RTrim(HGW.発注便) & "'"
        RTxtSql = RTxtSql & " And   品番 = '" & RTrim(HGW.品番) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HGWRead = True
           Call ItemsFromHGW(DBNo)
           HGW_RDSTS = True
        Else
           HGW_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call HGWClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HGWReadNext(Optional DBNo As Integer = 0)
    If Not HGW_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HGW_RDSTS = False
        Else
           Call ItemsFromHGW(DBNo)
           HGW_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HGWClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HGW_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HGW )
'*----------------------------------------------*'
Public Sub ItemsClearHGW()
    With HGW
        .処理端末 = ""
        .発注区分 = ""
        .仕入先 = ""
        .受入 = ""
        .納入日 = ""
        .納入便 = ""
        .発注日 = ""
        .発注便 = ""
        .品番 = ""
        .背番号 = ""
        .手配区分 = ""
        .SK区分 = 0
        .サイクル = ""
        .契約先 = ""
        .収容数 = 0
        .枚数 = 0
        .端数 = 0
        .発注数 = 0
        .端数区分 = 0
        .打切区分 = 0
        .打切種別 = 0
        .打切残数 = 0
        .打切日 = ""
        .作成種別 = ""
        .処理区分 = 0
'       .作成日 = 0
        .作成者 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function HGWInsert(Optional DBNo As Integer) As Boolean
    HGWInsert = True
    With HGW
        gSL_Select = "Insert Into 発注画面入力ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,発注区分,仕入先,"
        gSL_Select = gSL_Select & "受入,納入日,納入便,"
        gSL_Select = gSL_Select & "発注日,発注便,品番,"
        gSL_Select = gSL_Select & "背番号,手配区分,SK区分,"
        gSL_Select = gSL_Select & "サイクル,契約先,収容数,"
        gSL_Select = gSL_Select & "枚数,端数,発注数,"
        gSL_Select = gSL_Select & "端数区分,打切区分,打切種別,"
        gSL_Select = gSL_Select & "打切残数,打切日,作成種別,"
        gSL_Select = gSL_Select & "処理区分,"
        gSL_Select = gSL_Select & "作成日,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "HOST_NAME(),"
        gSL_Select = gSL_Select & "'" & RTrim(.発注区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発注日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発注便) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "" & .SK区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "" & .収容数 & ","
        gSL_Select = gSL_Select & "" & .枚数 & ","
        gSL_Select = gSL_Select & "" & .端数 & ","
        gSL_Select = gSL_Select & "" & .発注数 & ","
        gSL_Select = gSL_Select & "" & .端数区分 & ","
        gSL_Select = gSL_Select & "" & .打切区分 & ","
        gSL_Select = gSL_Select & "" & .打切種別 & ","
        gSL_Select = gSL_Select & "" & .打切残数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.打切日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.作成種別) & "',"
        gSL_Select = gSL_Select & "" & .処理区分 & ","
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "'"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HGWInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function HGWUpdate(Optional DBNo As Integer) As Boolean
    HGWUpdate = False
    With HGW
        gSL_Select = "Update 発注画面入力ワーク  Set "
        gSL_Select = gSL_Select & "発注区分 = '" & RTrim(.発注区分) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "発注日 = '" & RTrim(.発注日) & "',"
        gSL_Select = gSL_Select & "発注便 = '" & RTrim(.発注便) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "手配区分 = '" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "SK区分 = " & .SK区分 & ","
        gSL_Select = gSL_Select & "サイクル = '" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "枚数 = " & .枚数 & ","
        gSL_Select = gSL_Select & "端数 = " & .端数 & ","
        gSL_Select = gSL_Select & "発注数 = " & .発注数 & ","
        gSL_Select = gSL_Select & "端数区分 = " & .端数区分 & ","
        gSL_Select = gSL_Select & "打切区分 = " & .打切区分 & ","
        gSL_Select = gSL_Select & "打切種別 = " & .打切種別 & ","
        gSL_Select = gSL_Select & "打切残数 = " & .打切残数 & ","
        gSL_Select = gSL_Select & "打切日 = '" & RTrim(.打切日) & "',"
        gSL_Select = gSL_Select & "作成種別 = '" & RTrim(.作成種別) & "',"
        gSL_Select = gSL_Select & "処理区分 = " & .処理区分 & ""
        gSL_Select = gSL_Select & " Where 処理端末 = HOST_NAME()"
        gSL_Select = gSL_Select & "   And 発注区分 = '" & RTrim(.発注区分) & "'"
        gSL_Select = gSL_Select & "   And 仕入先 = '" & RTrim(.仕入先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 納入日 = '" & RTrim(.納入日) & "'"
        gSL_Select = gSL_Select & "   And 納入便 = '" & RTrim(.納入便) & "'"
        gSL_Select = gSL_Select & "   And 発注日 = '" & RTrim(.発注日) & "'"
        gSL_Select = gSL_Select & "   And 発注便 = '" & RTrim(.発注便) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HGWUpdate = True
End Function

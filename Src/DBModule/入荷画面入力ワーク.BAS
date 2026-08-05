Attribute VB_Name = "入荷画面入力ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   入荷画面入力ワーク
'*
'*   モジュールID  : 入荷画面入力ワーク.Bas
'*
'*
'*
'*   作成日  :  2004/02/18  By Cis
'*   変更日  :  2004/02/17  処理端末にHOST_NAME()
'*   変更日  :  2009/01/19  納入先/納入先受入 追加
'*---------------------------------------------------------------------------------
    Type T_入荷画面入力ワーク
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
        納入番号  As String
        納品書頁  As Byte
        納品書行  As Byte
        収容数    As Long
        枚数      As Long
        端数      As Long
        発注数    As Long
        作成種別  As String
        処理区分  As Byte
'        作成日    As Date
        作成者    As String
        件数          As Long
        
        表示品番        As String
        仕入先名        As String
        仕入先工場名    As String
        仕入先CD        As String
        仕入先工場CD    As String
        発注区分名      As String
        手配区分名      As String
        
        納入先      As String           '2009/01/19
        納入先受入      As String           '2009/01/19
    End Type

    Public NGW As T_入荷画面入力ワーク

    Public NGW_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromNGW(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearNGW
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With NGW
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
                Case "納入番号":  If Not IsNull(adoCol.Value) Then .納入番号 = adoCol.Value
                Case "納品書頁":  If Not IsNull(adoCol.Value) Then .納品書頁 = adoCol.Value
                Case "納品書行":  If Not IsNull(adoCol.Value) Then .納品書行 = adoCol.Value
                Case "収容数":    If Not IsNull(adoCol.Value) Then .収容数 = adoCol.Value
                Case "枚数":      If Not IsNull(adoCol.Value) Then .枚数 = adoCol.Value
                Case "端数":      If Not IsNull(adoCol.Value) Then .端数 = adoCol.Value
                Case "発注数":    If Not IsNull(adoCol.Value) Then .発注数 = adoCol.Value
                Case "作成種別":  If Not IsNull(adoCol.Value) Then .作成種別 = adoCol.Value
                Case "処理区分":  If Not IsNull(adoCol.Value) Then .処理区分 = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
                Case "作成者":    If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
            
                Case "表示品番":      If Not IsNull(adoCol.Value) Then .表示品番 = adoCol.Value
                Case "仕入先名":    If Not IsNull(adoCol.Value) Then .仕入先名 = adoCol.Value
                Case "仕入先工場名":    If Not IsNull(adoCol.Value) Then .仕入先工場名 = adoCol.Value
                Case "仕入先CD":    If Not IsNull(adoCol.Value) Then .仕入先CD = adoCol.Value
                Case "仕入先工場CD":    If Not IsNull(adoCol.Value) Then .仕入先工場CD = adoCol.Value
                Case "発注区分名":    If Not IsNull(adoCol.Value) Then .発注区分名 = adoCol.Value
                Case "手配区分名":    If Not IsNull(adoCol.Value) Then .手配区分名 = adoCol.Value
                
                Case "納入先":    If Not IsNull(adoCol.Value) Then .納入先 = adoCol.Value
                Case "納入先受入":    If Not IsNull(adoCol.Value) Then .納入先受入 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function NGWRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    NGWRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 入荷画面入力ワーク"
        RTxtSql = RTxtSql & " Where 処理端末 = HOST_NAME()"
        RTxtSql = RTxtSql & " And   発注区分 = '" & RTrim(NGW.発注区分) & "'"
        RTxtSql = RTxtSql & " And   仕入先 = '" & RTrim(NGW.仕入先) & "'"
        RTxtSql = RTxtSql & " And   受入 = '" & RTrim(NGW.受入) & "'"
        RTxtSql = RTxtSql & " And   納入日 = '" & RTrim(NGW.納入日) & "'"
        RTxtSql = RTxtSql & " And   納入便 = '" & RTrim(NGW.納入便) & "'"
        RTxtSql = RTxtSql & " And   発注日 = '" & RTrim(NGW.発注日) & "'"
        RTxtSql = RTxtSql & " And   発注便 = '" & RTrim(NGW.発注便) & "'"
        RTxtSql = RTxtSql & " And   品番 = '" & RTrim(NGW.品番) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           NGWRead = True
           Call ItemsFromNGW(DBNo)
           NGW_RDSTS = True
        Else
           NGW_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call NGWClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub NGWReadNext(Optional DBNo As Integer = 0)
    If Not NGW_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           NGW_RDSTS = False
        Else
           Call ItemsFromNGW(DBNo)
           NGW_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub NGWClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        NGW_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( NGW )
'*----------------------------------------------*'
Public Sub ItemsClearNGW()
    With NGW
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
        .納入番号 = ""
        .納品書頁 = 0
        .納品書行 = 0
        .収容数 = 0
        .枚数 = 0
        .端数 = 0
        .発注数 = 0
        .作成種別 = ""
        .処理区分 = 0
'       .作成日 = 0
        .作成者 = ""
        .件数 = 0
    
        .表示品番 = ""
        .仕入先名 = ""
        .仕入先工場名 = ""
        .仕入先CD = ""
        .仕入先工場CD = ""
        .発注区分名 = ""
        .手配区分名 = ""
        
        .納入先 = ""
        .納入先受入 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function NGWInsert(Optional DBNo As Integer) As Boolean
    NGWInsert = True
    With NGW
        gSL_Select = "Insert Into 入荷画面入力ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,発注区分,仕入先,"
        gSL_Select = gSL_Select & "受入,納入日,納入便,"
        gSL_Select = gSL_Select & "発注日,発注便,品番,"
        gSL_Select = gSL_Select & "背番号,手配区分,SK区分,"
        gSL_Select = gSL_Select & "サイクル,契約先,納入番号,"
        gSL_Select = gSL_Select & "納品書頁,納品書行,収容数,"
        gSL_Select = gSL_Select & "枚数,端数,発注数,"
        gSL_Select = gSL_Select & "作成種別,処理区分,"
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
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "" & .納品書頁 & ","
        gSL_Select = gSL_Select & "" & .納品書行 & ","
        gSL_Select = gSL_Select & "" & .収容数 & ","
        gSL_Select = gSL_Select & "" & .枚数 & ","
        gSL_Select = gSL_Select & "" & .端数 & ","
        gSL_Select = gSL_Select & "" & .発注数 & ","
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
    NGWInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function NGWUpdate(Optional DBNo As Integer) As Boolean
    NGWUpdate = False
    With NGW
        gSL_Select = "Update 入荷画面入力ワーク  Set "
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
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "納品書頁 = " & .納品書頁 & ","
        gSL_Select = gSL_Select & "納品書行 = " & .納品書行 & ","
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "枚数 = " & .枚数 & ","
        gSL_Select = gSL_Select & "端数 = " & .端数 & ","
        gSL_Select = gSL_Select & "発注数 = " & .発注数 & ","
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
    NGWUpdate = True
End Function

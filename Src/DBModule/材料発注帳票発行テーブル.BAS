Attribute VB_Name = "材料発注帳票発行テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料発注帳票発行テーブル
'*
'*   モジュールID  : 材料発注帳票発行テーブル.Bas
'*
'*
'*                                      /項目追加含む/
'*   作成日  :  2004/02/21  By Cis
'*   変更日  :  2004/02/21  行～納入数を配列に
'*---------------------------------------------------------------------------------
    Type T_材料発注帳票発行テーブル
        納品書管理NO          As Long
        納入番号              As String
        頁NO                  As Byte
        カード                As String
        発行順                As Byte
        契約先                As String
        仕入先                As String
        受入                  As String
        納入先                As String
        納入先受入            As String
        発注区分              As Byte
        納入日                As String
        納入便                As String
        手配区分              As String
        SK区分                As Byte
        箱総数                As Long
        総頁数                As Byte
        作成時間              As String
        グループNO            As String
        物理現頁数            As Byte
        物理総頁数            As Byte
        END表示区分           As String
        ALLEND表示区分        As String
        明細行数              As Byte
        行(10)                  As Byte
        材料管理番号(10)        As Long
        収容数(10)              As Currency
        箱数(10)                As Long
        端数(10)                As Currency
        納入数(10)              As Currency
        処理端末              As String
        処理区分              As Byte
        TEXTNAME              As String
        発行区分              As Byte
        発行種類              As String
        発行形態              As String
        発行日                As String
        作成日                As String
'        作成端末              As String
        件数          As Long
'   項目追加(自動)
        発注区分名    As String
        手配区分名    As String
        仕入先名      As String
        納入先名      As String
        仕入先CD      As String
        仕入先工場    As String
        仕入先工場名  As String
        納入先CD      As String
        納入先工場    As String
        納入先工場名  As String
        契約取引先    As String
        契約工場      As String
        被自社        As String
        被自社工場    As String
        被自社工場名  As String
        自社工場      As String
        機番          As String
        発行元        As String
    
        材質         As String
        表示寸法      As String
        単位名        As String
    End Type

    Public ZCH As T_材料発注帳票発行テーブル

    Public ZCH_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZCH(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZCH
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With ZCH
            Select Case adoCol.Name
                Case "納品書管理NO":          If Not IsNull(adoCol.Value) Then .納品書管理NO = adoCol.Value
                Case "納入番号":              If Not IsNull(adoCol.Value) Then .納入番号 = adoCol.Value
                Case "頁NO":                  If Not IsNull(adoCol.Value) Then .頁NO = adoCol.Value
                Case "カード":                If Not IsNull(adoCol.Value) Then .カード = adoCol.Value
                Case "発行順":                If Not IsNull(adoCol.Value) Then .発行順 = adoCol.Value
                Case "契約先":                If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
                Case "仕入先":                If Not IsNull(adoCol.Value) Then .仕入先 = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "納入先":                If Not IsNull(adoCol.Value) Then .納入先 = adoCol.Value
                Case "納入先受入":            If Not IsNull(adoCol.Value) Then .納入先受入 = adoCol.Value
                Case "発注区分":              If Not IsNull(adoCol.Value) Then .発注区分 = adoCol.Value
                Case "納入日":                If Not IsNull(adoCol.Value) Then .納入日 = adoCol.Value
                Case "納入便":                If Not IsNull(adoCol.Value) Then .納入便 = adoCol.Value
                Case "手配区分":              If Not IsNull(adoCol.Value) Then .手配区分 = adoCol.Value
                Case "SK区分":                If Not IsNull(adoCol.Value) Then .SK区分 = adoCol.Value
                Case "箱総数":                If Not IsNull(adoCol.Value) Then .箱総数 = adoCol.Value
                Case "総頁数":                If Not IsNull(adoCol.Value) Then .総頁数 = adoCol.Value
                Case "作成時間":              If Not IsNull(adoCol.Value) Then .作成時間 = adoCol.Value
                Case "グループNO":            If Not IsNull(adoCol.Value) Then .グループNO = adoCol.Value
                Case "物理現頁数":            If Not IsNull(adoCol.Value) Then .物理現頁数 = adoCol.Value
                Case "物理総頁数":            If Not IsNull(adoCol.Value) Then .物理総頁数 = adoCol.Value
                Case "END表示区分":           If Not IsNull(adoCol.Value) Then .END表示区分 = adoCol.Value
                Case "ALLEND表示区分":        If Not IsNull(adoCol.Value) Then .ALLEND表示区分 = adoCol.Value
                Case "明細行数":              If Not IsNull(adoCol.Value) Then .明細行数 = adoCol.Value
                Case "処理端末":              If Not IsNull(adoCol.Value) Then .処理端末 = adoCol.Value
                Case "処理区分":              If Not IsNull(adoCol.Value) Then .処理区分 = adoCol.Value
                Case "TEXTNAME":              If Not IsNull(adoCol.Value) Then .TEXTNAME = adoCol.Value
                Case "発行区分":              If Not IsNull(adoCol.Value) Then .発行区分 = adoCol.Value
                Case "発行種類":              If Not IsNull(adoCol.Value) Then .発行種類 = adoCol.Value
                Case "発行形態":              If Not IsNull(adoCol.Value) Then .発行形態 = adoCol.Value
                Case "発行日":                If Not IsNull(adoCol.Value) Then .発行日 = adoCol.Value
                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日 = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "発注区分名":      If Not IsNull(adoCol.Value) Then .発注区分名 = adoCol.Value
                Case "手配区分名":      If Not IsNull(adoCol.Value) Then .手配区分名 = adoCol.Value
                Case "仕入先名":        If Not IsNull(adoCol.Value) Then .仕入先名 = adoCol.Value
                Case "納入先名":        If Not IsNull(adoCol.Value) Then .納入先名 = adoCol.Value
                Case "仕入先CD":        If Not IsNull(adoCol.Value) Then .仕入先CD = adoCol.Value
                Case "仕入先工場":      If Not IsNull(adoCol.Value) Then .仕入先工場 = adoCol.Value
                Case "仕入先工場名":    If Not IsNull(adoCol.Value) Then .仕入先工場名 = adoCol.Value
                Case "納入先CD":        If Not IsNull(adoCol.Value) Then .納入先CD = adoCol.Value
                Case "納入先工場":      If Not IsNull(adoCol.Value) Then .納入先工場 = adoCol.Value
                Case "納入先工場名":    If Not IsNull(adoCol.Value) Then .納入先工場名 = adoCol.Value
                Case "契約取引先":      If Not IsNull(adoCol.Value) Then .契約取引先 = adoCol.Value
                Case "契約工場":        If Not IsNull(adoCol.Value) Then .契約工場 = adoCol.Value
                Case "被自社":  If Not IsNull(adoCol.Value) Then .被自社 = adoCol.Value
                Case "被自社工場":      If Not IsNull(adoCol.Value) Then .被自社工場 = adoCol.Value
                Case "被自社工場名":      If Not IsNull(adoCol.Value) Then .被自社工場名 = adoCol.Value
                Case "自社工場":        If Not IsNull(adoCol.Value) Then .自社工場 = adoCol.Value
                Case "機番":    If Not IsNull(adoCol.Value) Then .機番 = adoCol.Value
                Case "発行元":  If Not IsNull(adoCol.Value) Then .発行元 = adoCol.Value
                
                Case "材質":  If Not IsNull(adoCol.Value) Then .材質 = adoCol.Value
                Case "表示寸法":  If Not IsNull(adoCol.Value) Then .表示寸法 = adoCol.Value
                Case "単位名":  If Not IsNull(adoCol.Value) Then .単位名 = adoCol.Value
                Case Else:
                    If Not IsNull(adoCol.Value) Then
                        If adoCol.Name Like "行*" Then .行(Val(Mid(adoCol.Name, 2, 2))) = adoCol.Value
                        If adoCol.Name Like "材料管理番号*" Then .材料管理番号(Val(Mid(adoCol.Name, 7, 2))) = adoCol.Value
                        If adoCol.Name Like "収容数*" Then .収容数(Val(Mid(adoCol.Name, 4, 2))) = adoCol.Value
                        If adoCol.Name Like "箱数*" Then .箱数(Val(Mid(adoCol.Name, 3, 2))) = adoCol.Value
                        If adoCol.Name Like "端数*" Then .端数(Val(Mid(adoCol.Name, 3, 2))) = adoCol.Value
                        If adoCol.Name Like "納入数*" Then .納入数(Val(Mid(adoCol.Name, 4, 2))) = adoCol.Value
                    End If
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZCHRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZCHRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 材料発注帳票発行テーブル"
        RTxtSql = RTxtSql & " Where 納品書管理NO = " & ZCH.納品書管理NO & ""
        RTxtSql = RTxtSql & " And   納入番号 = '" & RTrim(ZCH.納入番号) & "'"
        RTxtSql = RTxtSql & " And   頁NO = " & ZCH.頁NO & ""
        RTxtSql = RTxtSql & " And   カード = '" & RTrim(ZCH.カード) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZCHRead = True
           Call ItemsFromZCH(DBNo)
           ZCH_RDSTS = True
        Else
           ZCH_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call ZCHClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZCHReadNext(Optional DBNo As Integer = 0)
    If Not ZCH_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZCH_RDSTS = False
        Else
           Call ItemsFromZCH(DBNo)
           ZCH_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZCHClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZCH_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZCH )
'*----------------------------------------------*'
Public Sub ItemsClearZCH()
    With ZCH
        .納品書管理NO = 0
        .納入番号 = ""
        .頁NO = 0
        .カード = ""
        .発行順 = 0
        .契約先 = ""
        .仕入先 = ""
        .受入 = ""
        .納入先 = ""
        .納入先受入 = ""
        .発注区分 = 0
        .納入日 = ""
        .納入便 = ""
        .手配区分 = ""
        .SK区分 = 0
        .箱総数 = 0
        .総頁数 = 0
        .作成時間 = ""
        .グループNO = ""
        .物理現頁数 = 0
        .物理総頁数 = 0
        .END表示区分 = ""
        .ALLEND表示区分 = ""
        .明細行数 = 0
        For gInt = 1 To 10
            .行(gInt) = 0
            .材料管理番号(gInt) = 0
            .収容数(gInt) = 0
            .箱数(gInt) = 0
            .端数(gInt) = 0
            .納入数(gInt) = 0
        Next
        .処理端末 = ""
        .処理区分 = 0
        .TEXTNAME = ""
        .発行区分 = 0
        .発行種類 = ""
        .発行形態 = ""
        .発行日 = ""
       .作成日 = ""
'       .作成端末 = ""
        .件数 = 0
'   項目追加(自動)
        .発注区分名 = ""
        .手配区分名 = ""
        .仕入先名 = ""
        .納入先名 = ""
        .仕入先CD = ""
        .仕入先工場 = ""
        .仕入先工場名 = ""
        .納入先CD = ""
        .納入先工場 = ""
        .納入先工場名 = ""
        .契約取引先 = ""
        .契約工場 = ""
        .被自社 = ""
        .被自社工場 = ""
        .被自社工場名 = ""
        .自社工場 = ""
        .機番 = ""
        .発行元 = ""
    
        .材質 = ""
        .表示寸法 = ""
        .単位名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZCHInsert(Optional DBNo As Integer) As Boolean
    ZCHInsert = True
    With ZCH
        gSL_Select = "Insert Into 材料発注帳票発行テーブル  ( "
        gSL_Select = gSL_Select & "納品書管理NO,納入番号,頁NO,"
        gSL_Select = gSL_Select & "カード,発行順,契約先,"
        gSL_Select = gSL_Select & "仕入先,受入,納入先,"
        gSL_Select = gSL_Select & "納入先受入,発注区分,納入日,"
        gSL_Select = gSL_Select & "納入便,手配区分,SK区分,"
        gSL_Select = gSL_Select & "箱総数,総頁数,作成時間,"
        gSL_Select = gSL_Select & "グループNO,物理現頁数,物理総頁数,"
        gSL_Select = gSL_Select & "END表示区分,ALLEND表示区分,明細行数,"
        For gInt = 1 To 10
            gSL_Select = gSL_Select & "行" & gInt & ","
            gSL_Select = gSL_Select & "材料管理番号" & gInt & ","
            gSL_Select = gSL_Select & "収容数" & gInt & ","
            gSL_Select = gSL_Select & "箱数" & gInt & ","
            gSL_Select = gSL_Select & "端数" & gInt & ","
            gSL_Select = gSL_Select & "納入数" & gInt & ","
        Next
        gSL_Select = gSL_Select & "処理端末,処理区分,TEXTNAME,"
        gSL_Select = gSL_Select & "発行区分,発行種類,発行形態,"
        gSL_Select = gSL_Select & "発行日,"
        gSL_Select = gSL_Select & "作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .納品書管理NO & ","
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "" & .頁NO & ","
        gSL_Select = gSL_Select & "'" & RTrim(.カード) & "',"
        gSL_Select = gSL_Select & "" & .発行順 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入先受入) & "',"
        gSL_Select = gSL_Select & "" & .発注区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "" & .SK区分 & ","
        gSL_Select = gSL_Select & "" & .箱総数 & ","
        gSL_Select = gSL_Select & "" & .総頁数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.作成時間) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.グループNO) & "',"
        gSL_Select = gSL_Select & "" & .物理現頁数 & ","
        gSL_Select = gSL_Select & "" & .物理総頁数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.END表示区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.ALLEND表示区分) & "',"
        gSL_Select = gSL_Select & "" & .明細行数 & ","
        For gInt = 1 To 10
            gSL_Select = gSL_Select & "" & .行(gInt) & ","
            gSL_Select = gSL_Select & "" & .材料管理番号(gInt) & ","
            gSL_Select = gSL_Select & "" & .収容数(gInt) & ","
            gSL_Select = gSL_Select & "" & .箱数(gInt) & ","
            gSL_Select = gSL_Select & "" & .端数(gInt) & ","
            gSL_Select = gSL_Select & "" & .納入数(gInt) & ","
        Next
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "" & .処理区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.TEXTNAME) & "',"
        gSL_Select = gSL_Select & "" & .発行区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.発行種類) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発行形態) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発行日) & "',"
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZCHInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZCHUpdate(Optional DBNo As Integer) As Boolean
    ZCHUpdate = False
    With ZCH
        gSL_Select = "Update 材料発注帳票発行テーブル  Set "
        gSL_Select = gSL_Select & "納品書管理NO = " & .納品書管理NO & ","
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "頁NO = " & .頁NO & ","
        gSL_Select = gSL_Select & "カード = '" & RTrim(.カード) & "',"
        gSL_Select = gSL_Select & "発行順 = " & .発行順 & ","
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "納入先受入 = '" & RTrim(.納入先受入) & "',"
        gSL_Select = gSL_Select & "発注区分 = " & .発注区分 & ","
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "手配区分 = '" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "SK区分 = " & .SK区分 & ","
        gSL_Select = gSL_Select & "箱総数 = " & .箱総数 & ","
        gSL_Select = gSL_Select & "総頁数 = " & .総頁数 & ","
        gSL_Select = gSL_Select & "作成時間 = '" & RTrim(.作成時間) & "',"
        gSL_Select = gSL_Select & "グループNO = '" & RTrim(.グループNO) & "',"
        gSL_Select = gSL_Select & "物理現頁数 = " & .物理現頁数 & ","
        gSL_Select = gSL_Select & "物理総頁数 = " & .物理総頁数 & ","
        gSL_Select = gSL_Select & "END表示区分 = '" & RTrim(.END表示区分) & "',"
        gSL_Select = gSL_Select & "ALLEND表示区分 = '" & RTrim(.ALLEND表示区分) & "',"
        gSL_Select = gSL_Select & "明細行数 = " & .明細行数 & ","
        For gInt = 1 To 10
            gSL_Select = gSL_Select & "行" & gInt & " = " & .行(gInt) & ","
            gSL_Select = gSL_Select & "材料管理番号" & gInt & " = " & .材料管理番号(gInt) & ","
            gSL_Select = gSL_Select & "収容数" & gInt & " = " & .収容数(gInt) & ","
            gSL_Select = gSL_Select & "箱数" & gInt & " = " & .箱数(gInt) & ","
            gSL_Select = gSL_Select & "端数" & gInt & " = " & .端数(gInt) & ","
            gSL_Select = gSL_Select & "納入数" & gInt & " = " & .納入数(gInt) & ","
        Next
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "処理区分 = " & .処理区分 & ","
        gSL_Select = gSL_Select & "TEXTNAME = '" & RTrim(.TEXTNAME) & "',"
        gSL_Select = gSL_Select & "発行区分 = " & .発行区分 & ","
        gSL_Select = gSL_Select & "発行種類 = '" & RTrim(.発行種類) & "',"
        gSL_Select = gSL_Select & "発行形態 = '" & RTrim(.発行形態) & "',"
        gSL_Select = gSL_Select & "発行日 = '" & RTrim(.発行日) & "'"
        gSL_Select = gSL_Select & " Where 納品書管理NO = " & .納品書管理NO & ""
        gSL_Select = gSL_Select & "   And 納入番号 = '" & RTrim(.納入番号) & "'"
        gSL_Select = gSL_Select & "   And 頁NO = " & .頁NO & ""
        gSL_Select = gSL_Select & "   And カード = '" & RTrim(.カード) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZCHUpdate = True
End Function

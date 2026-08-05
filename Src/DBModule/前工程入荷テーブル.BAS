Attribute VB_Name = "前工程入荷テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   前工程入荷テーブル
'*
'*   モジュールID  : 前工程入荷テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*                                      /項目追加含む/
'*   作成日  :  2007/07/26  By Cis
'*   作成日  :  2007/07/27  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_前工程入荷テーブル
        入荷管理NO            As Long
        納入番号              As String
        納品書頁              As Byte
        納品書行              As Byte
        納入日                As String
        納入便                As String
        仕入先                As String
        受注者                As String
        受注者工場            As String
        受入                  As String
        支給元体系            As String
        支給元1               As String
        支給元2               As String
        出荷場                As String
        納入先                As String
        発注者                As String
        発注者工場            As String
        発行日                As String
        発注区分              As String
        カード                As String
        機番                  As String
        総頁                  As Byte
        品番                  As String
        背番号                As String
        収容数                As Long
        枚数                  As Long
        納入数                As Long
        入荷日                As String
        入荷便                As String
        TNS作成区分           As Byte
        TNS作成管理NO         As Long
        TNS作成日             As String
        TNS作成処理日         As Date
        TNS連番               As Long
        検収区分              As String
        検収年月              As String
        検収処理日            As String
'        作成日                As Date
'        作成端末              As String
        作成者                As String
'        更新日                As Date
'        更新端末              As String
        更新者                As String
        件数          As Long
'   項目追加(自動)
        表示品番       As String
        発注区分名     As String
        契約先名       As String
    End Type

    Public MKN As T_前工程入荷テーブル

    Public MKN_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromMKN(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearMKN
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With MKN
            Select Case adoCol.Name
                Case "入荷管理NO":            If Not IsNull(adoCol.Value) Then .入荷管理NO = adoCol.Value
                Case "納入番号":              If Not IsNull(adoCol.Value) Then .納入番号 = adoCol.Value
                Case "納品書頁":              If Not IsNull(adoCol.Value) Then .納品書頁 = adoCol.Value
                Case "納品書行":              If Not IsNull(adoCol.Value) Then .納品書行 = adoCol.Value
                Case "納入日":                If Not IsNull(adoCol.Value) Then .納入日 = adoCol.Value
                Case "納入便":                If Not IsNull(adoCol.Value) Then .納入便 = adoCol.Value
                Case "仕入先":                If Not IsNull(adoCol.Value) Then .仕入先 = adoCol.Value
                Case "受注者":                If Not IsNull(adoCol.Value) Then .受注者 = adoCol.Value
                Case "受注者工場":            If Not IsNull(adoCol.Value) Then .受注者工場 = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "支給元体系":            If Not IsNull(adoCol.Value) Then .支給元体系 = adoCol.Value
                Case "支給元1":               If Not IsNull(adoCol.Value) Then .支給元1 = adoCol.Value
                Case "支給元2":               If Not IsNull(adoCol.Value) Then .支給元2 = adoCol.Value
                Case "出荷場":                If Not IsNull(adoCol.Value) Then .出荷場 = adoCol.Value
                Case "納入先":                If Not IsNull(adoCol.Value) Then .納入先 = adoCol.Value
                Case "発注者":                If Not IsNull(adoCol.Value) Then .発注者 = adoCol.Value
                Case "発注者工場":            If Not IsNull(adoCol.Value) Then .発注者工場 = adoCol.Value
                Case "発行日":                If Not IsNull(adoCol.Value) Then .発行日 = adoCol.Value
                Case "発注区分":              If Not IsNull(adoCol.Value) Then .発注区分 = adoCol.Value
                Case "カード":                If Not IsNull(adoCol.Value) Then .カード = adoCol.Value
                Case "機番":                  If Not IsNull(adoCol.Value) Then .機番 = adoCol.Value
                Case "総頁":                  If Not IsNull(adoCol.Value) Then .総頁 = adoCol.Value
                Case "品番":                  If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "背番号":                If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "収容数":                If Not IsNull(adoCol.Value) Then .収容数 = adoCol.Value
                Case "枚数":                  If Not IsNull(adoCol.Value) Then .枚数 = adoCol.Value
                Case "納入数":                If Not IsNull(adoCol.Value) Then .納入数 = adoCol.Value
                Case "入荷日":                If Not IsNull(adoCol.Value) Then .入荷日 = adoCol.Value
                Case "入荷便":                If Not IsNull(adoCol.Value) Then .入荷便 = adoCol.Value
                Case "TNS作成区分":           If Not IsNull(adoCol.Value) Then .TNS作成区分 = adoCol.Value
                Case "TNS作成管理NO":         If Not IsNull(adoCol.Value) Then .TNS作成管理NO = adoCol.Value
                Case "TNS作成日":             If Not IsNull(adoCol.Value) Then .TNS作成日 = adoCol.Value
                Case "TNS作成処理日":         If Not IsNull(adoCol.Value) Then .TNS作成処理日 = adoCol.Value
                Case "TNS連番":               If Not IsNull(adoCol.Value) Then .TNS連番 = adoCol.Value
                Case "検収区分":              If Not IsNull(adoCol.Value) Then .検収区分 = adoCol.Value
                Case "検収年月":              If Not IsNull(adoCol.Value) Then .検収年月 = adoCol.Value
                Case "検収処理日":            If Not IsNull(adoCol.Value) Then .検収処理日 = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "作成者":                If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "更新日":                If Not IsNull(adoCol.Value) Then .更新日               = adoCol.Value
'                Case "更新端末":              If Not IsNull(adoCol.Value) Then .更新端末             = adoCol.Value
                Case "更新者":                If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "表示品番":        If Not IsNull(adoCol.Value) Then .表示品番 = adoCol.Value
                Case "発注区分名":      If Not IsNull(adoCol.Value) Then .発注区分名 = adoCol.Value
                Case "契約先名":        If Not IsNull(adoCol.Value) Then .契約先名 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function MKNRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    MKNRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 前工程入荷テーブル"
        RTxtSql = RTxtSql & " Where 入荷管理NO = " & MKN.入荷管理NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           MKNRead = True
           Call ItemsFromMKN(DBNo)
           MKN_RDSTS = True
        Else
           MKN_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call MKNClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub MKNReadNext(Optional DBNo As Integer = 0)
    If Not MKN_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           MKN_RDSTS = False
        Else
           Call ItemsFromMKN(DBNo)
           MKN_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub MKNClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        MKN_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( MKN )
'*----------------------------------------------*'
Public Sub ItemsClearMKN()
    With MKN
        .入荷管理NO = 0
        .納入番号 = ""
        .納品書頁 = 0
        .納品書行 = 0
        .納入日 = ""
        .納入便 = ""
        .仕入先 = ""
        .受注者 = ""
        .受注者工場 = ""
        .受入 = ""
        .支給元体系 = ""
        .支給元1 = ""
        .支給元2 = ""
        .出荷場 = ""
        .納入先 = ""
        .発注者 = ""
        .発注者工場 = ""
        .発行日 = ""
        .発注区分 = ""
        .カード = ""
        .機番 = ""
        .総頁 = 0
        .品番 = ""
        .背番号 = ""
        .収容数 = 0
        .枚数 = 0
        .納入数 = 0
        .入荷日 = ""
        .入荷便 = ""
        .TNS作成区分 = 0
        .TNS作成管理NO = 0
        .TNS作成日 = ""
        .TNS作成処理日 = 0
        .TNS連番 = 0
        .検収区分 = ""
        .検収年月 = ""
        .検収処理日 = ""
'       .作成日 = 0
'       .作成端末 = ""
        .作成者 = ""
'       .更新日 = 0
'       .更新端末 = ""
        .更新者 = ""
        .件数 = 0
'   項目追加(自動)
        .表示品番 = ""
        .発注区分名 = ""
        .契約先名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function MKNInsert(Optional DBNo As Integer) As Boolean
    MKNInsert = False
    With MKN
        gSL_Select = "Insert Into 前工程入荷テーブル  ( "
        gSL_Select = gSL_Select & "納入番号,納品書頁,納品書行,"
        gSL_Select = gSL_Select & "納入日,納入便,仕入先,"
        gSL_Select = gSL_Select & "受注者,受注者工場,受入,"
        gSL_Select = gSL_Select & "支給元体系,支給元1,支給元2,"
        gSL_Select = gSL_Select & "出荷場,納入先,発注者,"
        gSL_Select = gSL_Select & "発注者工場,発行日,発注区分,"
        gSL_Select = gSL_Select & "カード,機番,総頁,"
        gSL_Select = gSL_Select & "品番,背番号,収容数,"
        gSL_Select = gSL_Select & "枚数,納入数,入荷日,"
        gSL_Select = gSL_Select & "入荷便,TNS作成区分,TNS作成管理NO,"
        gSL_Select = gSL_Select & "TNS作成日,TNS作成処理日,TNS連番,"
        gSL_Select = gSL_Select & "検収区分,検収年月,検収処理日,"
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "" & .納品書頁 & ","
        gSL_Select = gSL_Select & "" & .納品書行 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受注者) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受注者工場) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.支給元体系) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.支給元1) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.支給元2) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.出荷場) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発注者) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発注者工場) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発行日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発注区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.カード) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.機番) & "',"
        gSL_Select = gSL_Select & "" & .総頁 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "" & .収容数 & ","
        gSL_Select = gSL_Select & "" & .枚数 & ","
        gSL_Select = gSL_Select & "" & .納入数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.入荷日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.入荷便) & "',"
        gSL_Select = gSL_Select & "" & .TNS作成区分 & ","
        gSL_Select = gSL_Select & "" & .TNS作成管理NO & ","
        gSL_Select = gSL_Select & "'" & RTrim(.TNS作成日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.TNS作成処理日) & "',"
        gSL_Select = gSL_Select & "" & .TNS連番 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.検収区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.検収年月) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.検収処理日) & "',"
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "'"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    MKNInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function MKNUpdate(Optional DBNo As Integer) As Boolean
    MKNUpdate = False
    With MKN
        gSL_Select = "Update 前工程入荷テーブル  Set "
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "納品書頁 = " & .納品書頁 & ","
        gSL_Select = gSL_Select & "納品書行 = " & .納品書行 & ","
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "受注者 = '" & RTrim(.受注者) & "',"
        gSL_Select = gSL_Select & "受注者工場 = '" & RTrim(.受注者工場) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "支給元体系 = '" & RTrim(.支給元体系) & "',"
        gSL_Select = gSL_Select & "支給元1 = '" & RTrim(.支給元1) & "',"
        gSL_Select = gSL_Select & "支給元2 = '" & RTrim(.支給元2) & "',"
        gSL_Select = gSL_Select & "出荷場 = '" & RTrim(.出荷場) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "発注者 = '" & RTrim(.発注者) & "',"
        gSL_Select = gSL_Select & "発注者工場 = '" & RTrim(.発注者工場) & "',"
        gSL_Select = gSL_Select & "発行日 = '" & RTrim(.発行日) & "',"
        gSL_Select = gSL_Select & "発注区分 = '" & RTrim(.発注区分) & "',"
        gSL_Select = gSL_Select & "カード = '" & RTrim(.カード) & "',"
        gSL_Select = gSL_Select & "機番 = '" & RTrim(.機番) & "',"
        gSL_Select = gSL_Select & "総頁 = " & .総頁 & ","
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "枚数 = " & .枚数 & ","
        gSL_Select = gSL_Select & "納入数 = " & .納入数 & ","
        gSL_Select = gSL_Select & "入荷日 = '" & RTrim(.入荷日) & "',"
        gSL_Select = gSL_Select & "入荷便 = '" & RTrim(.入荷便) & "',"
        gSL_Select = gSL_Select & "TNS作成区分 = " & .TNS作成区分 & ","
        gSL_Select = gSL_Select & "TNS作成管理NO = " & .TNS作成管理NO & ","
        gSL_Select = gSL_Select & "TNS作成日 = '" & RTrim(.TNS作成日) & "',"
        gSL_Select = gSL_Select & "TNS作成処理日 = '" & RTrim(.TNS作成処理日) & "',"
        gSL_Select = gSL_Select & "TNS連番 = " & .TNS連番 & ","
        gSL_Select = gSL_Select & "検収区分 = '" & RTrim(.検収区分) & "',"
        gSL_Select = gSL_Select & "検収年月 = '" & RTrim(.検収年月) & "',"
        gSL_Select = gSL_Select & "検収処理日 = '" & RTrim(.検収処理日) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "'"
        gSL_Select = gSL_Select & " Where 入荷管理NO = " & .入荷管理NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    MKNUpdate = True
End Function

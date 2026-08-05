Attribute VB_Name = "受注テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   受注テーブル
'*
'*   モジュールID  : 受注テーブル.Bas
'*
'*   作成日  :  2004/04/15  By Cis
'*   作成日  :  2008/05/02  By Cis
'*   変更日  :  2009/09/21  By Cis   出荷管理部署/出荷管理部署名
'*---------------------------------------------------------------------------------
    Type T_受注テーブル
        受注管理NO        As Long
        受注管理行NO      As Byte
        出荷指示書NO      As Long
        出荷指示書行NO    As Byte
        受注区分          As Byte
        受注日            As String
        受注便            As String
        納入日            As String
        納入便            As String
        出荷予定日        As String
        出荷予定便        As String
        納入先            As String
        受入              As String
        機番              As String
        納入番号          As String
        納品書頁          As Byte
        納品書行          As Byte
        分納区分          As Byte
        分納回数          As Byte
        品番              As String
        背番号            As String
        枚数              As Long
        収容数            As Long
        納入数            As Long
        端数              As Long
        抜取枚数          As Integer
        エフ発行枚数      As Integer
        エフ発行端数      As Long
        作成種別          As String
        納品書発行区分    As Byte
        指示書発行区分    As Byte
        出荷実績区分      As Byte
        納品書管理NO      As Long
        元出荷指示書NO    As Long
        元出荷指示書行NO  As Byte
'        作成日            As Date
'        作成端末          As String
        作成者            As String
'        更新日            As Date
'        更新端末          As String
        更新者            As String
        件数          As Long
'   項目追加(自動)
        納入先名       As String
        契約先名       As String
        SK区分名       As String
        受注区分名     As String
        表示品番       As String
        SK区分         As Byte
        品名           As String
        工場名         As String
        社内背番       As String
        出荷日         As String
        出荷便         As String
        出荷枚数       As Long
        出荷端数       As Long
        出荷数         As Long
        検収日         As String
        検収管理NO     As Long
        実数区分       As Byte
        手配区分       As String
        手配区分名     As String
        取引先CD       As String
        工場CD         As String
        契約先         As String
        契約取引先     As String
        契約工場       As String
        品番編集       As String
        カード         As String
        実績年月       As String
        金額           As Currency
        検収管理行NO   As Integer
        出荷管理部署   As String
        出荷管理部署名 As String
    End Type

    Public JYT As T_受注テーブル

    Public JYT_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromJYT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearJYT
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With JYT
            Select Case adoCol.Name
                Case "受注管理NO":        If Not IsNull(adoCol.Value) Then .受注管理NO = adoCol.Value
                Case "受注管理行NO":      If Not IsNull(adoCol.Value) Then .受注管理行NO = adoCol.Value
                Case "出荷指示書NO":      If Not IsNull(adoCol.Value) Then .出荷指示書NO = adoCol.Value
                Case "出荷指示書行NO":    If Not IsNull(adoCol.Value) Then .出荷指示書行NO = adoCol.Value
                Case "受注区分":          If Not IsNull(adoCol.Value) Then .受注区分 = adoCol.Value
                Case "受注日":            If Not IsNull(adoCol.Value) Then .受注日 = adoCol.Value
                Case "受注便":            If Not IsNull(adoCol.Value) Then .受注便 = adoCol.Value
                Case "納入日":            If Not IsNull(adoCol.Value) Then .納入日 = adoCol.Value
                Case "納入便":            If Not IsNull(adoCol.Value) Then .納入便 = adoCol.Value
                Case "出荷予定日":        If Not IsNull(adoCol.Value) Then .出荷予定日 = adoCol.Value
                Case "出荷予定便":        If Not IsNull(adoCol.Value) Then .出荷予定便 = adoCol.Value
                Case "納入先":            If Not IsNull(adoCol.Value) Then .納入先 = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "機番":              If Not IsNull(adoCol.Value) Then .機番 = adoCol.Value
                Case "納入番号":          If Not IsNull(adoCol.Value) Then .納入番号 = adoCol.Value
                Case "納品書頁":          If Not IsNull(adoCol.Value) Then .納品書頁 = adoCol.Value
                Case "納品書行":          If Not IsNull(adoCol.Value) Then .納品書行 = adoCol.Value
                Case "分納区分":          If Not IsNull(adoCol.Value) Then .分納区分 = adoCol.Value
                Case "分納回数":          If Not IsNull(adoCol.Value) Then .分納回数 = adoCol.Value
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "背番号":            If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "枚数":              If Not IsNull(adoCol.Value) Then .枚数 = adoCol.Value
                Case "収容数":            If Not IsNull(adoCol.Value) Then .収容数 = adoCol.Value
                Case "納入数":            If Not IsNull(adoCol.Value) Then .納入数 = adoCol.Value
                Case "端数":              If Not IsNull(adoCol.Value) Then .端数 = adoCol.Value
                Case "抜取枚数":          If Not IsNull(adoCol.Value) Then .抜取枚数 = adoCol.Value
                Case "エフ発行枚数":      If Not IsNull(adoCol.Value) Then .エフ発行枚数 = adoCol.Value
                Case "エフ発行端数":      If Not IsNull(adoCol.Value) Then .エフ発行端数 = adoCol.Value
                Case "作成種別":          If Not IsNull(adoCol.Value) Then .作成種別 = adoCol.Value
                Case "納品書発行区分":    If Not IsNull(adoCol.Value) Then .納品書発行区分 = adoCol.Value
                Case "指示書発行区分":    If Not IsNull(adoCol.Value) Then .指示書発行区分 = adoCol.Value
                Case "出荷実績区分":      If Not IsNull(adoCol.Value) Then .出荷実績区分 = adoCol.Value
                Case "納品書管理NO":      If Not IsNull(adoCol.Value) Then .納品書管理NO = adoCol.Value
                Case "元出荷指示書NO":    If Not IsNull(adoCol.Value) Then .元出荷指示書NO = adoCol.Value
                Case "元出荷指示書行NO":  If Not IsNull(adoCol.Value) Then .元出荷指示書行NO = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "更新者":            If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "納入先名":        If Not IsNull(adoCol.Value) Then .納入先名 = adoCol.Value
                Case "契約先名":        If Not IsNull(adoCol.Value) Then .契約先名 = adoCol.Value
                Case "SK区分名":        If Not IsNull(adoCol.Value) Then .SK区分名 = adoCol.Value
                Case "受注区分名":      If Not IsNull(adoCol.Value) Then .受注区分名 = adoCol.Value
                Case "表示品番":        If Not IsNull(adoCol.Value) Then .表示品番 = adoCol.Value
                Case "SK区分":  If Not IsNull(adoCol.Value) Then .SK区分 = adoCol.Value
                Case "品名":    If Not IsNull(adoCol.Value) Then .品名 = adoCol.Value
                Case "工場名":  If Not IsNull(adoCol.Value) Then .工場名 = adoCol.Value
                Case "社内背番":        If Not IsNull(adoCol.Value) Then .社内背番 = adoCol.Value
                Case "出荷日":  If Not IsNull(adoCol.Value) Then .出荷日 = adoCol.Value
                Case "出荷便":  If Not IsNull(adoCol.Value) Then .出荷便 = adoCol.Value
                Case "出荷枚数":        If Not IsNull(adoCol.Value) Then .出荷枚数 = adoCol.Value
                Case "出荷端数":        If Not IsNull(adoCol.Value) Then .出荷端数 = adoCol.Value
                Case "出荷数":  If Not IsNull(adoCol.Value) Then .出荷数 = adoCol.Value
                Case "検収日":  If Not IsNull(adoCol.Value) Then .検収日 = adoCol.Value
                Case "検収管理NO":      If Not IsNull(adoCol.Value) Then .検収管理NO = adoCol.Value
                Case "実数区分":        If Not IsNull(adoCol.Value) Then .実数区分 = adoCol.Value
                Case "手配区分":        If Not IsNull(adoCol.Value) Then .手配区分 = adoCol.Value
                Case "手配区分名":      If Not IsNull(adoCol.Value) Then .手配区分名 = adoCol.Value
                Case "取引先CD":        If Not IsNull(adoCol.Value) Then .取引先CD = adoCol.Value
                Case "工場CD":  If Not IsNull(adoCol.Value) Then .工場CD = adoCol.Value
                Case "契約先":  If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
                Case "契約取引先":      If Not IsNull(adoCol.Value) Then .契約取引先 = adoCol.Value
                Case "契約工場":        If Not IsNull(adoCol.Value) Then .契約工場 = adoCol.Value
                Case "品番編集":        If Not IsNull(adoCol.Value) Then .品番編集 = adoCol.Value
                Case "カード":  If Not IsNull(adoCol.Value) Then .カード = adoCol.Value
                Case "実績年月":        If Not IsNull(adoCol.Value) Then .実績年月 = adoCol.Value
                Case "金額":    If Not IsNull(adoCol.Value) Then .金額 = adoCol.Value
                Case "検収管理行NO":    If Not IsNull(adoCol.Value) Then .検収管理行NO = adoCol.Value
                Case "出荷管理部署":    If Not IsNull(adoCol.Value) Then .出荷管理部署 = adoCol.Value
                Case "出荷管理部署名":  If Not IsNull(adoCol.Value) Then .出荷管理部署名 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function JYTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    JYTRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 受注テーブル"
        RTxtSql = RTxtSql & " Where 受注管理NO = " & JYT.受注管理NO & ""
        RTxtSql = RTxtSql & " And   受注管理行NO = " & JYT.受注管理行NO & ""
        RTxtSql = RTxtSql & " And   出荷指示書NO = " & JYT.出荷指示書NO & ""
        RTxtSql = RTxtSql & " And   出荷指示書行NO = " & JYT.出荷指示書行NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           JYTRead = True
           Call ItemsFromJYT(DBNo)
           JYT_RDSTS = True
        Else
           JYT_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call JYTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub JYTReadNext(Optional DBNo As Integer = 0)
    If Not JYT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           JYT_RDSTS = False
        Else
           Call ItemsFromJYT(DBNo)
           JYT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub JYTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        JYT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( JYT )
'*----------------------------------------------*'
Public Sub ItemsClearJYT()
    With JYT
        .受注管理NO = 0
        .受注管理行NO = 0
        .出荷指示書NO = 0
        .出荷指示書行NO = 0
        .受注区分 = 0
        .受注日 = ""
        .受注便 = ""
        .納入日 = ""
        .納入便 = ""
        .出荷予定日 = ""
        .出荷予定便 = ""
        .納入先 = ""
        .受入 = ""
        .機番 = ""
        .納入番号 = ""
        .納品書頁 = 0
        .納品書行 = 0
        .分納区分 = 0
        .分納回数 = 0
        .品番 = ""
        .背番号 = ""
        .枚数 = 0
        .収容数 = 0
        .納入数 = 0
        .端数 = 0
        .抜取枚数 = 0
        .エフ発行枚数 = 0
        .エフ発行端数 = 0
        .作成種別 = ""
        .納品書発行区分 = 0
        .指示書発行区分 = 0
        .出荷実績区分 = 0
        .納品書管理NO = 0
        .元出荷指示書NO = 0
        .元出荷指示書行NO = 0
'       .作成日 = 0
'       .作成端末 = ""
        .作成者 = ""
'       .更新日 = 0
'       .更新端末 = ""
        .更新者 = ""
        .件数 = 0
'   項目追加(自動)
        .納入先名 = ""
        .契約先名 = ""
        .SK区分名 = ""
        .受注区分名 = ""
        .表示品番 = ""
        .SK区分 = 0
        .品名 = ""
        .工場名 = ""
        .社内背番 = ""
        .出荷日 = ""
        .出荷便 = ""
        .出荷枚数 = 0
        .出荷端数 = 0
        .出荷数 = 0
        .検収日 = ""
        .検収管理NO = 0
        .実数区分 = 0
        .手配区分 = ""
        .手配区分名 = ""
        .取引先CD = ""
        .工場CD = ""
        .契約先 = ""
        .契約取引先 = ""
        .契約工場 = ""
        .品番編集 = ""
        .カード = ""
        .実績年月 = ""
        .金額 = 0
        .検収管理行NO = 0
    
        .出荷管理部署 = ""
        .出荷管理部署名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function JYTInsert(Optional DBNo As Integer) As Boolean
    JYTInsert = False
    With JYT
        gSL_Select = "Insert Into 受注テーブル  ( "
        gSL_Select = gSL_Select & "受注管理NO,受注管理行NO,出荷指示書NO,"
        gSL_Select = gSL_Select & "出荷指示書行NO,受注区分,受注日,"
        gSL_Select = gSL_Select & "受注便,納入日,納入便,"
        gSL_Select = gSL_Select & "出荷予定日,出荷予定便,納入先,"
        gSL_Select = gSL_Select & "受入,機番,納入番号,"
        gSL_Select = gSL_Select & "納品書頁,納品書行,分納区分,"
        gSL_Select = gSL_Select & "分納回数,品番,背番号,"
        gSL_Select = gSL_Select & "枚数,収容数,納入数,"
        gSL_Select = gSL_Select & "端数,抜取枚数,エフ発行枚数,"
        gSL_Select = gSL_Select & "エフ発行端数,作成種別,納品書発行区分,"
        gSL_Select = gSL_Select & "指示書発行区分,出荷実績区分,納品書管理NO,"
        gSL_Select = gSL_Select & "元出荷指示書NO,元出荷指示書行NO,"
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .受注管理NO & ","
        gSL_Select = gSL_Select & "" & .受注管理行NO & ","
        gSL_Select = gSL_Select & "" & .出荷指示書NO & ","
        gSL_Select = gSL_Select & "" & .出荷指示書行NO & ","
        gSL_Select = gSL_Select & "" & .受注区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.受注日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受注便) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.出荷予定日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.出荷予定便) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.機番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "" & .納品書頁 & ","
        gSL_Select = gSL_Select & "" & .納品書行 & ","
        gSL_Select = gSL_Select & "" & .分納区分 & ","
        gSL_Select = gSL_Select & "" & .分納回数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "" & .枚数 & ","
        gSL_Select = gSL_Select & "" & .収容数 & ","
        gSL_Select = gSL_Select & "" & .納入数 & ","
        gSL_Select = gSL_Select & "" & .端数 & ","
        gSL_Select = gSL_Select & "" & .抜取枚数 & ","
        gSL_Select = gSL_Select & "" & .エフ発行枚数 & ","
        gSL_Select = gSL_Select & "" & .エフ発行端数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.作成種別) & "',"
        gSL_Select = gSL_Select & "" & .納品書発行区分 & ","
        gSL_Select = gSL_Select & "" & .指示書発行区分 & ","
        gSL_Select = gSL_Select & "" & .出荷実績区分 & ","
        gSL_Select = gSL_Select & "" & .納品書管理NO & ","
        gSL_Select = gSL_Select & "" & .元出荷指示書NO & ","
        gSL_Select = gSL_Select & "" & .元出荷指示書行NO & ","
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
    JYTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function JYTUpdate(Optional DBNo As Integer) As Boolean
    JYTUpdate = False
    With JYT
        gSL_Select = "Update 受注テーブル  Set "
        gSL_Select = gSL_Select & "受注管理NO = " & .受注管理NO & ","
        gSL_Select = gSL_Select & "受注管理行NO = " & .受注管理行NO & ","
        gSL_Select = gSL_Select & "出荷指示書NO = " & .出荷指示書NO & ","
        gSL_Select = gSL_Select & "出荷指示書行NO = " & .出荷指示書行NO & ","
        gSL_Select = gSL_Select & "受注区分 = " & .受注区分 & ","
        gSL_Select = gSL_Select & "受注日 = '" & RTrim(.受注日) & "',"
        gSL_Select = gSL_Select & "受注便 = '" & RTrim(.受注便) & "',"
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "出荷予定日 = '" & RTrim(.出荷予定日) & "',"
        gSL_Select = gSL_Select & "出荷予定便 = '" & RTrim(.出荷予定便) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "機番 = '" & RTrim(.機番) & "',"
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "納品書頁 = " & .納品書頁 & ","
        gSL_Select = gSL_Select & "納品書行 = " & .納品書行 & ","
        gSL_Select = gSL_Select & "分納区分 = " & .分納区分 & ","
        gSL_Select = gSL_Select & "分納回数 = " & .分納回数 & ","
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "枚数 = " & .枚数 & ","
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "納入数 = " & .納入数 & ","
        gSL_Select = gSL_Select & "端数 = " & .端数 & ","
        gSL_Select = gSL_Select & "抜取枚数 = " & .抜取枚数 & ","
        gSL_Select = gSL_Select & "エフ発行枚数 = " & .エフ発行枚数 & ","
        gSL_Select = gSL_Select & "エフ発行端数 = " & .エフ発行端数 & ","
        gSL_Select = gSL_Select & "作成種別 = '" & RTrim(.作成種別) & "',"
        gSL_Select = gSL_Select & "納品書発行区分 = " & .納品書発行区分 & ","
        gSL_Select = gSL_Select & "指示書発行区分 = " & .指示書発行区分 & ","
        gSL_Select = gSL_Select & "出荷実績区分 = " & .出荷実績区分 & ","
        gSL_Select = gSL_Select & "納品書管理NO = " & .納品書管理NO & ","
        gSL_Select = gSL_Select & "元出荷指示書NO = " & .元出荷指示書NO & ","
        gSL_Select = gSL_Select & "元出荷指示書行NO = " & .元出荷指示書行NO & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "'"
        gSL_Select = gSL_Select & " Where 受注管理NO = " & .受注管理NO & ""
        gSL_Select = gSL_Select & "   And 受注管理行NO = " & .受注管理行NO & ""
        gSL_Select = gSL_Select & "   And 出荷指示書NO = " & .出荷指示書NO & ""
        gSL_Select = gSL_Select & "   And 出荷指示書行NO = " & .出荷指示書行NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    JYTUpdate = True
End Function

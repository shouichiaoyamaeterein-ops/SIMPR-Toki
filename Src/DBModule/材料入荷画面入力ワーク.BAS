Attribute VB_Name = "材料入荷画面入力ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料入荷画面入力ワーク
'*
'*   モジュールID  : 材料入荷画面入力ワーク.Bas
'*
'*
'*                                      /項目追加含む/
'*   作成日  :  2004/02/25  By Cis
'*   変更日  :  2004/02/25
'*   変更日  :  2008/12/16  By Cis   品番 追加
'*   変更日  :  2009/01/16  By Cis   発注管理NO 追加
'*---------------------------------------------------------------------------------
    Type T_材料入荷画面入力ワーク
        処理端末      As String
        発注区分      As String
        仕入先        As String
        受入          As String
        納入日        As String
        納入便        As String
        発注日        As String
        発注便        As String
        材料管理番号  As Long
        材質          As String
        板厚          As Currency
        幅            As Currency
        長さ          As Currency
        背番号        As String
        手配区分      As String
        SK区分        As Byte
        サイクル      As String
        契約先        As String
        納入先        As String
        納入先受入    As String
        納入番号      As String
        納品書頁      As Byte
        納品書行      As Byte
        収容数        As Currency
        枚数          As Long
        端数          As Currency
        発注数        As Currency
        作成種別      As String
        処理区分      As Byte
'        作成日        As Date
        作成者        As String
        件数          As Long
        発注管理NO    As Long
'   項目追加(自動)
        仕入先名      As String
        仕入先工場名  As String
        仕入先CD      As String
        仕入先工場CD  As String
        手配区分名    As String
        発注区分名    As String
        表示寸法      As String
        手配単位名    As String
        材種コード    As String
        材種名        As String
        品番          As String
    End Type

    Public ZNG As T_材料入荷画面入力ワーク

    Public ZNG_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZNG(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZNG
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With ZNG
            Select Case adoCol.Name
                Case "処理端末":      If Not IsNull(adoCol.Value) Then .処理端末 = adoCol.Value
                Case "発注区分":      If Not IsNull(adoCol.Value) Then .発注区分 = adoCol.Value
                Case "仕入先":        If Not IsNull(adoCol.Value) Then .仕入先 = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "納入日":        If Not IsNull(adoCol.Value) Then .納入日 = adoCol.Value
                Case "納入便":        If Not IsNull(adoCol.Value) Then .納入便 = adoCol.Value
                Case "発注日":        If Not IsNull(adoCol.Value) Then .発注日 = adoCol.Value
                Case "発注便":        If Not IsNull(adoCol.Value) Then .発注便 = adoCol.Value
                Case "材料管理番号":  If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "材質":          If Not IsNull(adoCol.Value) Then .材質 = adoCol.Value
                Case "板厚":          If Not IsNull(adoCol.Value) Then .板厚 = adoCol.Value
                Case "幅":            If Not IsNull(adoCol.Value) Then .幅 = adoCol.Value
                Case "長さ":          If Not IsNull(adoCol.Value) Then .長さ = adoCol.Value
                Case "背番号":        If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "手配区分":      If Not IsNull(adoCol.Value) Then .手配区分 = adoCol.Value
                Case "SK区分":        If Not IsNull(adoCol.Value) Then .SK区分 = adoCol.Value
                Case "サイクル":      If Not IsNull(adoCol.Value) Then .サイクル = adoCol.Value
                Case "契約先":        If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
                Case "納入先":        If Not IsNull(adoCol.Value) Then .納入先 = adoCol.Value
                Case "納入先受入":    If Not IsNull(adoCol.Value) Then .納入先受入 = adoCol.Value
                Case "納入番号":      If Not IsNull(adoCol.Value) Then .納入番号 = adoCol.Value
                Case "納品書頁":      If Not IsNull(adoCol.Value) Then .納品書頁 = adoCol.Value
                Case "納品書行":      If Not IsNull(adoCol.Value) Then .納品書行 = adoCol.Value
                Case "収容数":        If Not IsNull(adoCol.Value) Then .収容数 = adoCol.Value
                Case "枚数":          If Not IsNull(adoCol.Value) Then .枚数 = adoCol.Value
                Case "端数":          If Not IsNull(adoCol.Value) Then .端数 = adoCol.Value
                Case "発注数":        If Not IsNull(adoCol.Value) Then .発注数 = adoCol.Value
                Case "作成種別":      If Not IsNull(adoCol.Value) Then .作成種別 = adoCol.Value
                Case "処理区分":      If Not IsNull(adoCol.Value) Then .処理区分 = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
                Case "発注管理NO":    If Not IsNull(adoCol.Value) Then .発注管理NO = adoCol.Value
'   項目追加(自動)
                Case "仕入先名":        If Not IsNull(adoCol.Value) Then .仕入先名 = adoCol.Value
                Case "仕入先工場名":    If Not IsNull(adoCol.Value) Then .仕入先工場名 = adoCol.Value
                Case "仕入先CD":        If Not IsNull(adoCol.Value) Then .仕入先CD = adoCol.Value
                Case "仕入先工場CD":    If Not IsNull(adoCol.Value) Then .仕入先工場CD = adoCol.Value
                Case "手配区分名":      If Not IsNull(adoCol.Value) Then .手配区分名 = adoCol.Value
                Case "発注区分名":      If Not IsNull(adoCol.Value) Then .発注区分名 = adoCol.Value
                Case "表示寸法":        If Not IsNull(adoCol.Value) Then .表示寸法 = adoCol.Value
                Case "手配単位名":      If Not IsNull(adoCol.Value) Then .手配単位名 = adoCol.Value
                Case "材種コード":      If Not IsNull(adoCol.Value) Then .材種コード = adoCol.Value
                Case "材種名":  If Not IsNull(adoCol.Value) Then .材種名 = adoCol.Value
                Case "品番":            If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZNGRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZNGRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 材料入荷画面入力ワーク"
        RTxtSql = RTxtSql & " Where 処理端末 = '" & RTrim(ZNG.処理端末) & "'"
        RTxtSql = RTxtSql & " And   発注区分 = '" & RTrim(ZNG.発注区分) & "'"
        RTxtSql = RTxtSql & " And   仕入先 = '" & RTrim(ZNG.仕入先) & "'"
        RTxtSql = RTxtSql & " And   受入 = '" & RTrim(ZNG.受入) & "'"
        RTxtSql = RTxtSql & " And   納入日 = '" & RTrim(ZNG.納入日) & "'"
        RTxtSql = RTxtSql & " And   納入便 = '" & RTrim(ZNG.納入便) & "'"
        RTxtSql = RTxtSql & " And   発注日 = '" & RTrim(ZNG.発注日) & "'"
        RTxtSql = RTxtSql & " And   発注便 = '" & RTrim(ZNG.発注便) & "'"
        RTxtSql = RTxtSql & " And   材料管理番号 = " & ZNG.材料管理番号 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZNGRead = True
           Call ItemsFromZNG(DBNo)
           ZNG_RDSTS = True
        Else
           ZNG_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call ZNGClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZNGReadNext(Optional DBNo As Integer = 0)
    If Not ZNG_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZNG_RDSTS = False
        Else
           Call ItemsFromZNG(DBNo)
           ZNG_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZNGClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZNG_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZNG )
'*----------------------------------------------*'
Public Sub ItemsClearZNG()
    With ZNG
        .処理端末 = ""
        .発注区分 = ""
        .仕入先 = ""
        .受入 = ""
        .納入日 = ""
        .納入便 = ""
        .発注日 = ""
        .発注便 = ""
        .材料管理番号 = 0
        .材質 = ""
        .板厚 = 0
        .幅 = 0
        .長さ = 0
        .背番号 = ""
        .手配区分 = ""
        .SK区分 = 0
        .サイクル = ""
        .契約先 = ""
        .納入先 = ""
        .納入先受入 = ""
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
        .発注管理NO = 0
'   項目追加(自動)
        .仕入先名 = ""
        .仕入先工場名 = ""
        .仕入先CD = ""
        .仕入先工場CD = ""
        .手配区分名 = ""
        .発注区分名 = ""
        .表示寸法 = ""
        .手配単位名 = ""
        .材種コード = ""
        .材種名 = ""
        .品番 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZNGInsert(Optional DBNo As Integer) As Boolean
    ZNGInsert = True
    With ZNG
        gSL_Select = "Insert Into 材料入荷画面入力ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,発注区分,仕入先,"
        gSL_Select = gSL_Select & "受入,納入日,納入便,"
        gSL_Select = gSL_Select & "発注日,発注便,材料管理番号,"
        gSL_Select = gSL_Select & "材質,板厚,幅,"
        gSL_Select = gSL_Select & "長さ,背番号,手配区分,"
        gSL_Select = gSL_Select & "SK区分,サイクル,契約先,"
        gSL_Select = gSL_Select & "納入先,納入先受入,納入番号,"
        gSL_Select = gSL_Select & "納品書頁,納品書行,収容数,"
        gSL_Select = gSL_Select & "枚数,端数,発注数,"
        gSL_Select = gSL_Select & "作成種別,処理区分,"
        gSL_Select = gSL_Select & "作成日,作成者,発注管理NO"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "HOST_NAME(),"
        gSL_Select = gSL_Select & "'" & RTrim(.発注区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発注日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発注便) & "',"
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "" & .板厚 & ","
        gSL_Select = gSL_Select & "" & .幅 & ","
        gSL_Select = gSL_Select & "" & .長さ & ","
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "" & .SK区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入先受入) & "',"
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
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "',"
        gSL_Select = gSL_Select & "" & .発注管理NO & ""
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZNGInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZNGUpdate(Optional DBNo As Integer) As Boolean
    ZNGUpdate = False
    With ZNG
        gSL_Select = "Update 材料入荷画面入力ワーク  Set "
        gSL_Select = gSL_Select & "発注区分 = '" & RTrim(.発注区分) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "発注日 = '" & RTrim(.発注日) & "',"
        gSL_Select = gSL_Select & "発注便 = '" & RTrim(.発注便) & "',"
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "板厚 = " & .板厚 & ","
        gSL_Select = gSL_Select & "幅 = " & .幅 & ","
        gSL_Select = gSL_Select & "長さ = " & .長さ & ","
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "手配区分 = '" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "SK区分 = " & .SK区分 & ","
        gSL_Select = gSL_Select & "サイクル = '" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "納入先受入 = '" & RTrim(.納入先受入) & "',"
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "納品書頁 = " & .納品書頁 & ","
        gSL_Select = gSL_Select & "納品書行 = " & .納品書行 & ","
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "枚数 = " & .枚数 & ","
        gSL_Select = gSL_Select & "端数 = " & .端数 & ","
        gSL_Select = gSL_Select & "発注数 = " & .発注数 & ","
        gSL_Select = gSL_Select & "作成種別 = '" & RTrim(.作成種別) & "',"
        gSL_Select = gSL_Select & "処理区分 = " & .処理区分 & ","
        gSL_Select = gSL_Select & "枚数 = " & .発注管理NO & ""
        gSL_Select = gSL_Select & " Where 処理端末 = HOST_NAME()"
        gSL_Select = gSL_Select & "   And 発注区分 = '" & RTrim(.発注区分) & "'"
        gSL_Select = gSL_Select & "   And 仕入先 = '" & RTrim(.仕入先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 納入日 = '" & RTrim(.納入日) & "'"
        gSL_Select = gSL_Select & "   And 納入便 = '" & RTrim(.納入便) & "'"
        gSL_Select = gSL_Select & "   And 発注日 = '" & RTrim(.発注日) & "'"
        gSL_Select = gSL_Select & "   And 発注便 = '" & RTrim(.発注便) & "'"
        gSL_Select = gSL_Select & "   And 材料管理番号 = " & .材料管理番号 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZNGUpdate = True
End Function

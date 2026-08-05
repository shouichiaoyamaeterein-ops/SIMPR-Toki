Attribute VB_Name = "材料発注画面入力ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料発注画面入力ワーク
'*
'*   モジュールID  : 材料発注画面入力ワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*
'*   作成日  :  2004/02/20  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料発注画面入力ワーク
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
        収容数        As Currency
        枚数          As Long
        端数          As Currency
        発注数        As Currency
        端数区分      As Byte
        打切区分      As Byte
        打切種別      As Byte
        打切残数      As Long
        打切日        As String
        作成種別      As String
        処理区分      As Byte
'        作成日        As Date
        作成者        As String
        件数          As Long
    End Type

    Public ZHG As T_材料発注画面入力ワーク

    Public ZHG_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZHG(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZHG
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With ZHG
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
                Case "収容数":        If Not IsNull(adoCol.Value) Then .収容数 = adoCol.Value
                Case "枚数":          If Not IsNull(adoCol.Value) Then .枚数 = adoCol.Value
                Case "端数":          If Not IsNull(adoCol.Value) Then .端数 = adoCol.Value
                Case "発注数":        If Not IsNull(adoCol.Value) Then .発注数 = adoCol.Value
                Case "端数区分":      If Not IsNull(adoCol.Value) Then .端数区分 = adoCol.Value
                Case "打切区分":      If Not IsNull(adoCol.Value) Then .打切区分 = adoCol.Value
                Case "打切種別":      If Not IsNull(adoCol.Value) Then .打切種別 = adoCol.Value
                Case "打切残数":      If Not IsNull(adoCol.Value) Then .打切残数 = adoCol.Value
                Case "打切日":        If Not IsNull(adoCol.Value) Then .打切日 = adoCol.Value
                Case "作成種別":      If Not IsNull(adoCol.Value) Then .作成種別 = adoCol.Value
                Case "処理区分":      If Not IsNull(adoCol.Value) Then .処理区分 = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZHGRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZHGRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 材料発注画面入力ワーク"
        RTxtSql = RTxtSql & " Where 処理端末 = '" & RTrim(ZHG.処理端末) & "'"
        RTxtSql = RTxtSql & " And   発注区分 = '" & RTrim(ZHG.発注区分) & "'"
        RTxtSql = RTxtSql & " And   仕入先 = '" & RTrim(ZHG.仕入先) & "'"
        RTxtSql = RTxtSql & " And   受入 = '" & RTrim(ZHG.受入) & "'"
        RTxtSql = RTxtSql & " And   納入日 = '" & RTrim(ZHG.納入日) & "'"
        RTxtSql = RTxtSql & " And   納入便 = '" & RTrim(ZHG.納入便) & "'"
        RTxtSql = RTxtSql & " And   発注日 = '" & RTrim(ZHG.発注日) & "'"
        RTxtSql = RTxtSql & " And   発注便 = '" & RTrim(ZHG.発注便) & "'"
        RTxtSql = RTxtSql & " And   材料管理番号 = " & ZHG.材料管理番号 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZHGRead = True
           Call ItemsFromZHG(DBNo)
           ZHG_RDSTS = True
        Else
           ZHG_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call ZHGClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZHGReadNext(Optional DBNo As Integer = 0)
    If Not ZHG_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZHG_RDSTS = False
        Else
           Call ItemsFromZHG(DBNo)
           ZHG_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZHGClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZHG_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZHG )
'*----------------------------------------------*'
Public Sub ItemsClearZHG()
    With ZHG
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
Public Function ZHGInsert(Optional DBNo As Integer) As Boolean
    ZHGInsert = True
    With ZHG
        gSL_Select = "Insert Into 材料発注画面入力ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,発注区分,仕入先,"
        gSL_Select = gSL_Select & "受入,納入日,納入便,"
        gSL_Select = gSL_Select & "発注日,発注便,材料管理番号,"
        gSL_Select = gSL_Select & "材質,板厚,幅,"
        gSL_Select = gSL_Select & "長さ,背番号,手配区分,"
        gSL_Select = gSL_Select & "SK区分,サイクル,契約先,"
        gSL_Select = gSL_Select & "収容数,枚数,端数,"
        gSL_Select = gSL_Select & "発注数,端数区分,打切区分,"
        gSL_Select = gSL_Select & "打切種別,打切残数,打切日,"
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
    ZHGInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZHGUpdate(Optional DBNo As Integer) As Boolean
    ZHGUpdate = False
    With ZHG
        gSL_Select = "Update 材料発注画面入力ワーク  Set "
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
        gSL_Select = gSL_Select & "   And 材料管理番号 = " & .材料管理番号 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZHGUpdate = True
End Function

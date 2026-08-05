Attribute VB_Name = "構成問合せワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   構成問合せワーク
'*
'*   モジュールID  : 構成問合せワーク1150.Bas   (Based 構成問合せワーク.Bas)
'*
'*   ( 当モジュールは自動生成されたモジュールに不足分を追加したものです。)
'*
'*   作成日  :  2019/07/26  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_構成問合せワーク
        処理端末          As String
        順位              As String
        レベル            As Byte
        親品番            As String
        子品番            As String
        構成数            As Currency
        展開区分          As Byte
        LT                As Byte
        適用開始日        As String
        適用終了日        As String
        区分              As String
        表示品番          As String
        表示寸法          As String
        品名              As String
        材料管理番号      As Long
        件数              As Long
'   項目追加(手動)                          ' 2019/07/18 ADD
        検索元品番      As String           ' 2019/07/18 ADD
        号補区分        As String           ' 2019/07/26 ADD
        手配先          As String           ' 2019/07/26 ADD
        取引先          As String           ' 2019/07/26 ADD
    End Type

    Public KOL As T_構成問合せワーク

    Public KOL_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKOL(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKOL
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With KOL
            Select Case adoCol.Name
                Case "処理端末":          If Not IsNull(adoCol.Value) Then .処理端末 = adoCol.Value
                Case "順位":              If Not IsNull(adoCol.Value) Then .順位 = adoCol.Value
                Case "レベル":            If Not IsNull(adoCol.Value) Then .レベル = adoCol.Value
                Case "親品番":            If Not IsNull(adoCol.Value) Then .親品番 = adoCol.Value
                Case "子品番":            If Not IsNull(adoCol.Value) Then .子品番 = adoCol.Value
                Case "構成数":            If Not IsNull(adoCol.Value) Then .構成数 = adoCol.Value
                Case "展開区分":          If Not IsNull(adoCol.Value) Then .展開区分 = adoCol.Value
                Case "LT":                If Not IsNull(adoCol.Value) Then .LT = adoCol.Value
                Case "適用開始日":        If Not IsNull(adoCol.Value) Then .適用開始日 = adoCol.Value
                Case "適用終了日":        If Not IsNull(adoCol.Value) Then .適用終了日 = adoCol.Value
                Case "区分":              If Not IsNull(adoCol.Value) Then .区分 = adoCol.Value
                Case "表示品番":          If Not IsNull(adoCol.Value) Then .表示品番 = adoCol.Value
                Case "表示寸法":          If Not IsNull(adoCol.Value) Then .表示寸法 = adoCol.Value
                Case "品名":              If Not IsNull(adoCol.Value) Then .品名 = adoCol.Value
                Case "材料管理番号":      If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
                ''
                Case "検索元品番":        If Not IsNull(adoCol.Value) Then .検索元品番 = adoCol.Value       ' 2019/07/26 ADD
                Case "号補区分":          If Not IsNull(adoCol.Value) Then .号補区分 = adoCol.Value         ' 2019/07/26 ADD
                Case "手配先":            If Not IsNull(adoCol.Value) Then .手配先 = adoCol.Value           ' 2019/07/26 ADD
                Case "取引先":            If Not IsNull(adoCol.Value) Then .手配先 = adoCol.Value           ' 2019/07/26 ADD

            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KOLRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KOLRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 構成問合せワーク"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KOLRead = True
           Call ItemsFromKOL(DBNo)
           KOL_RDSTS = True
        Else
           KOL_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call KOLClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KOLReadNext(Optional DBNo As Integer = 0)
    If Not KOL_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KOL_RDSTS = False
        Else
           Call ItemsFromKOL(DBNo)
           KOL_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KOLClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KOL_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KOL )
'*----------------------------------------------*'
Public Sub ItemsClearKOL()
    With KOL
        .処理端末 = ""
        .順位 = ""
        .レベル = 0
        .親品番 = ""
        .子品番 = ""
        .構成数 = 0
        .展開区分 = 0
        .LT = 0
        .適用開始日 = ""
        .適用終了日 = ""
        .区分 = ""
        .表示品番 = ""
        .表示寸法 = ""
        .品名 = ""
        .材料管理番号 = 0
        .件数 = 0
        .検索元品番 = ""        ' 2019/07/18 ADD
        .号補区分 = ""          ' 2019/07/26 ADD
        .手配先 = ""            ' 2019/07/26 ADD
        .取引先 = ""            ' 2019/07/26 ADD
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KOLInsert(Optional DBNo As Integer) As Boolean
    KOLInsert = True
    With KOL
        gSL_Select = "Insert Into 構成問合せワーク  ( "
        gSL_Select = gSL_Select & "処理端末,順位,レベル,"
        gSL_Select = gSL_Select & "親品番,子品番,構成数,"
        gSL_Select = gSL_Select & "展開区分,LT,適用開始日,"
        gSL_Select = gSL_Select & "適用終了日,区分,表示品番,"
        gSL_Select = gSL_Select & "表示寸法,品名"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.順位) & "',"
        gSL_Select = gSL_Select & "" & .レベル & ","
        gSL_Select = gSL_Select & "'" & RTrim(.親品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.子品番) & "',"
        gSL_Select = gSL_Select & "" & .構成数 & ","
        gSL_Select = gSL_Select & "" & .展開区分 & ","
        gSL_Select = gSL_Select & "" & .LT & ","
        gSL_Select = gSL_Select & "'" & RTrim(.適用開始日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.適用終了日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.表示品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.表示寸法) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品名) & "'"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KOLInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KOLUpdate(Optional DBNo As Integer) As Boolean
    KOLUpdate = False
    With KOL
        gSL_Select = "Update 構成問合せワーク  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "順位 = '" & RTrim(.順位) & "',"
        gSL_Select = gSL_Select & "レベル = " & .レベル & ","
        gSL_Select = gSL_Select & "親品番 = '" & RTrim(.親品番) & "',"
        gSL_Select = gSL_Select & "子品番 = '" & RTrim(.子品番) & "',"
        gSL_Select = gSL_Select & "構成数 = " & .構成数 & ","
        gSL_Select = gSL_Select & "展開区分 = " & .展開区分 & ","
        gSL_Select = gSL_Select & "LT = " & .LT & ","
        gSL_Select = gSL_Select & "適用開始日 = '" & RTrim(.適用開始日) & "',"
        gSL_Select = gSL_Select & "適用終了日 = '" & RTrim(.適用終了日) & "',"
        gSL_Select = gSL_Select & "区分 = '" & RTrim(.区分) & "',"
        gSL_Select = gSL_Select & "表示品番 = '" & RTrim(.表示品番) & "',"
        gSL_Select = gSL_Select & "表示寸法 = '" & RTrim(.表示寸法) & "',"
        gSL_Select = gSL_Select & "品名 = '" & RTrim(.品名) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KOLUpdate = True
End Function

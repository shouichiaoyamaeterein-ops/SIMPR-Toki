Attribute VB_Name = "支給実績状況問合ビュー"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   支給実績状況問合ビュー
'*
'*   モジュールID  : 支給実績状況問合ビュー.Bas
'*
'*   作成日  :  2008/07/07  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_支給実績状況問合ビュー
        実績日          As String
        実績便          As String
        実績日便        As String
        納品書頁        As String
        納品書行        As String
        納入番号        As String
        表示納入番号    As String
        品番            As String
        表示品番        As String
        背番号          As String
        表示背番号      As String
        実績先          As String
        略称            As String
        編集支給先      As String
        受入            As String
        収容数          As Long
        実績枚数        As Long
        実績数          As Long
        車型            As String
        データ区分名    As String
        置場            As String
        サイクル        As String
        品名            As String
        実績管理NO      As Long
        有無償区分      As String
        表示有無償区分  As String
        単価            As Currency
        金額            As Currency
    End Type

    Public SJV As T_支給実績状況問合ビュー

    Public SJV_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromSJV(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearSJV
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With SJV
            Select Case adoCol.Name
                Case "実績日":          If Not IsNull(adoCol.Value) Then .実績日 = adoCol.Value
                Case "実績便":          If Not IsNull(adoCol.Value) Then .実績便 = adoCol.Value
                Case "実績日便":        If Not IsNull(adoCol.Value) Then .実績日便 = adoCol.Value
                Case "納品書頁":        If Not IsNull(adoCol.Value) Then .納品書頁 = adoCol.Value
                Case "納品書行":        If Not IsNull(adoCol.Value) Then .納品書行 = adoCol.Value
                Case "納入番号":        If Not IsNull(adoCol.Value) Then .納入番号 = adoCol.Value
                Case "表示納入番号":    If Not IsNull(adoCol.Value) Then .表示納入番号 = adoCol.Value
                Case "品番":            If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "表示品番":        If Not IsNull(adoCol.Value) Then .表示品番 = adoCol.Value
                Case "背番号":          If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "表示背番号":      If Not IsNull(adoCol.Value) Then .表示背番号 = adoCol.Value
                Case "実績先":          If Not IsNull(adoCol.Value) Then .実績先 = adoCol.Value
                Case "略称":            If Not IsNull(adoCol.Value) Then .略称 = adoCol.Value
                Case "編集支給先":      If Not IsNull(adoCol.Value) Then .編集支給先 = adoCol.Value
                Case "受入":            If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "収容数":          If Not IsNull(adoCol.Value) Then .収容数 = adoCol.Value
                Case "実績枚数":        If Not IsNull(adoCol.Value) Then .実績枚数 = adoCol.Value
                Case "実績数":          If Not IsNull(adoCol.Value) Then .実績数 = adoCol.Value
                Case "車型":            If Not IsNull(adoCol.Value) Then .車型 = adoCol.Value
                Case "データ区分名":    If Not IsNull(adoCol.Value) Then .データ区分名 = adoCol.Value
                Case "置場":            If Not IsNull(adoCol.Value) Then .置場 = adoCol.Value
                Case "サイクル":        If Not IsNull(adoCol.Value) Then .サイクル = adoCol.Value
                Case "品名":            If Not IsNull(adoCol.Value) Then .品名 = adoCol.Value
                Case "実績管理NO":      If Not IsNull(adoCol.Value) Then .実績管理NO = adoCol.Value
                Case "有無償区分":      If Not IsNull(adoCol.Value) Then .有無償区分 = adoCol.Value
                Case "表示有無償区分":  If Not IsNull(adoCol.Value) Then .表示有無償区分 = adoCol.Value
                Case "単価":            If Not IsNull(adoCol.Value) Then .単価 = adoCol.Value
                Case "金額":            If Not IsNull(adoCol.Value) Then .金額 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function SJVRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    SJVRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 支給実績状況問合ビュー"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           SJVRead = True
           Call ItemsFromSJV(DBNo)
           SJV_RDSTS = True
        Else
           SJV_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call SJVClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub SJVReadNext(Optional DBNo As Integer = 0)
    If Not SJV_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           SJV_RDSTS = False
        Else
           Call ItemsFromSJV(DBNo)
           SJV_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub SJVClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        SJV_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( SJV )
'*----------------------------------------------*'
Public Sub ItemsClearSJV()
    With SJV
        .実績日 = ""
        .実績便 = ""
        .実績日便 = ""
        .納品書頁 = ""
        .納品書行 = ""
        .納入番号 = ""
        .表示納入番号 = ""
        .品番 = ""
        .表示品番 = ""
        .背番号 = ""
        .表示背番号 = ""
        .実績先 = ""
        .略称 = ""
        .編集支給先 = ""
        .受入 = ""
        .収容数 = 0
        .実績枚数 = 0
        .実績数 = 0
        .車型 = ""
        .データ区分名 = ""
        .置場 = ""
        .サイクル = ""
        .品名 = ""
        .実績管理NO = 0
        .有無償区分 = ""
        .表示有無償区分 = ""
        .単価 = 0
        .金額 = 0
    End With
End Sub

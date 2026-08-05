Attribute VB_Name = "材料かんばん重複チェックテーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料かんばん重複チェックテーブル
'*
'*   モジュールID  : 材料かんばん重複チェックテーブル.Bas
'*
'*                                      /項目追加含む/
'*   作成日  :  2004/09/22  By Cis      ' 作成日コメントアウト
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料かんばん重複チェックテーブル
        納入日        As String
        納入便        As String
        材料管理番号  As Long
        取引先        As String
        背番号        As String
        受入          As String
        連番          As Integer
'        作成端末      As String
        作成日        As Date
        件数          As Long
'   項目追加(自動)
        表示寸法      As String
        仕入先名      As String
        材質          As String
    End Type

    Public ZJC As T_材料かんばん重複チェックテーブル

    Public ZJC_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZJC(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZJC
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With ZJC
            Select Case adoCol.Name
                Case "納入日":        If Not IsNull(adoCol.Value) Then .納入日 = adoCol.Value
                Case "納入便":        If Not IsNull(adoCol.Value) Then .納入便 = adoCol.Value
                Case "材料管理番号":  If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "取引先":        If Not IsNull(adoCol.Value) Then .取引先 = adoCol.Value
                Case "背番号":        If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "連番":          If Not IsNull(adoCol.Value) Then .連番 = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "表示寸法":        If Not IsNull(adoCol.Value) Then .表示寸法 = adoCol.Value
                Case "仕入先名":        If Not IsNull(adoCol.Value) Then .仕入先名 = adoCol.Value
                Case "材質":    If Not IsNull(adoCol.Value) Then .材質 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZJCRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZJCRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 材料かんばん重複チェックテーブル"
        RTxtSql = RTxtSql & " Where 納入日 = '" & RTrim(ZJC.納入日) & "'"
        RTxtSql = RTxtSql & " And   納入便 = '" & RTrim(ZJC.納入便) & "'"
        RTxtSql = RTxtSql & " And   材料管理番号 = " & ZJC.材料管理番号 & ""
        RTxtSql = RTxtSql & " And   取引先 = '" & RTrim(ZJC.取引先) & "'"
        RTxtSql = RTxtSql & " And   背番号 = '" & RTrim(ZJC.背番号) & "'"
        RTxtSql = RTxtSql & " And   受入 = '" & RTrim(ZJC.受入) & "'"
        RTxtSql = RTxtSql & " And   連番 = " & ZJC.連番 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZJCRead = True
           Call ItemsFromZJC(DBNo)
           ZJC_RDSTS = True
        Else
           ZJC_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call ZJCClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZJCReadNext(Optional DBNo As Integer = 0)
    If Not ZJC_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZJC_RDSTS = False
        Else
           Call ItemsFromZJC(DBNo)
           ZJC_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZJCClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZJC_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZJC )
'*----------------------------------------------*'
Public Sub ItemsClearZJC()
    With ZJC
        .納入日 = ""
        .納入便 = ""
        .材料管理番号 = 0
        .取引先 = ""
        .背番号 = ""
        .受入 = ""
        .連番 = 0
'       .作成端末 = ""
'       .作成日 = 0
        .件数 = 0
'   項目追加(自動)
        .表示寸法 = ""
        .仕入先名 = ""
        .材質 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZJCInsert(Optional DBNo As Integer) As Boolean
    ZJCInsert = True
    With ZJC
        gSL_Select = "Insert Into 材料かんばん重複チェックテーブル  ( "
        gSL_Select = gSL_Select & "納入日,納入便,材料管理番号,"
        gSL_Select = gSL_Select & "取引先,背番号,受入,"
        gSL_Select = gSL_Select & "連番,"
        gSL_Select = gSL_Select & "作成端末,作成日"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "" & .連番 & ","
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "GETDATE()"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZJCInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZJCUpdate(Optional DBNo As Integer) As Boolean
    ZJCUpdate = False
    With ZJC
        gSL_Select = "Update 材料かんばん重複チェックテーブル  Set "
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "連番 = " & .連番 & ""
        gSL_Select = gSL_Select & " Where 納入日 = '" & RTrim(.納入日) & "'"
        gSL_Select = gSL_Select & "   And 納入便 = '" & RTrim(.納入便) & "'"
        gSL_Select = gSL_Select & "   And 材料管理番号 = " & .材料管理番号 & ""
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 背番号 = '" & RTrim(.背番号) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 連番 = " & .連番 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZJCUpdate = True
End Function

Attribute VB_Name = "かんばん重複チェックテーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   かんばん重複チェックテーブル
'*
'*   モジュールID  : かんばん重複チェックテーブル.Bas
'*
'*                                      /項目追加含む/
'*   作成日  :  2005/06/15  By Cis
'*   変更日  :  2006/12/20  項目追加（版数）
'*---------------------------------------------------------------------------------
    Type T_かんばん重複チェックテーブル
        処理区分  As String
        納入日    As String
        納入便    As String
        取引先    As String
        品番      As String
        背番号    As String
        受入      As String
        連番      As Integer
        版数      As Byte
'        作成端末  As String
        作成日    As Date
        件数          As Long
'   項目追加(自動)
        仕入先名      As String
        表示品番      As String
        日付          As String
        収容区分      As Byte
    End Type

    Public KJC As T_かんばん重複チェックテーブル

    Public KJC_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKJC(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKJC
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With KJC
            Select Case adoCol.Name
                Case "処理区分":  If Not IsNull(adoCol.Value) Then .処理区分 = adoCol.Value
                Case "納入日":    If Not IsNull(adoCol.Value) Then .納入日 = adoCol.Value
                Case "納入便":    If Not IsNull(adoCol.Value) Then .納入便 = adoCol.Value
                Case "取引先":    If Not IsNull(adoCol.Value) Then .取引先 = adoCol.Value
                Case "品番":      If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "背番号":    If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "受入":      If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "連番":      If Not IsNull(adoCol.Value) Then .連番 = adoCol.Value
                Case "版数":      If Not IsNull(adoCol.Value) Then .版数 = adoCol.Value
'                Case "作成端末":  If Not IsNull(adoCol.Value) Then .作成端末 = adoCol.Value
                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "仕入先名":        If Not IsNull(adoCol.Value) Then .仕入先名 = adoCol.Value
                Case "表示品番":        If Not IsNull(adoCol.Value) Then .表示品番 = adoCol.Value
                Case "日付":    If Not IsNull(adoCol.Value) Then .日付 = adoCol.Value
                Case "収容区分":        If Not IsNull(adoCol.Value) Then .収容区分 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KJCRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KJCRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From かんばん重複チェックテーブル"
        RTxtSql = RTxtSql & " Where 処理区分 = '" & RTrim(KJC.処理区分) & "'"
        RTxtSql = RTxtSql & " And   納入日 = '" & RTrim(KJC.納入日) & "'"
        RTxtSql = RTxtSql & " And   納入便 = '" & RTrim(KJC.納入便) & "'"
        RTxtSql = RTxtSql & " And   取引先 = '" & RTrim(KJC.取引先) & "'"
        RTxtSql = RTxtSql & " And   品番 = '" & RTrim(KJC.品番) & "'"
        RTxtSql = RTxtSql & " And   背番号 = '" & RTrim(KJC.背番号) & "'"
        RTxtSql = RTxtSql & " And   受入 = '" & RTrim(KJC.受入) & "'"
        RTxtSql = RTxtSql & " And   連番 = " & KJC.連番 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KJCRead = True
           Call ItemsFromKJC(DBNo)
           KJC_RDSTS = True
        Else
           KJC_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call KJCClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KJCReadNext(Optional DBNo As Integer = 0)
    If Not KJC_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KJC_RDSTS = False
        Else
           Call ItemsFromKJC(DBNo)
           KJC_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KJCClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KJC_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KJC )
'*----------------------------------------------*'
Public Sub ItemsClearKJC()
    With KJC
        .処理区分 = ""
        .納入日 = ""
        .納入便 = ""
        .取引先 = ""
        .品番 = ""
        .背番号 = ""
        .受入 = ""
        .連番 = 0
        .版数 = 0
'       .作成端末 = ""
        .作成日 = 0
        .件数 = 0
'   項目追加(自動)
        .仕入先名 = ""
        .表示品番 = ""
        .日付 = ""
        .収容区分 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KJCInsert(Optional DBNo As Integer) As Boolean
    KJCInsert = True
    With KJC
        gSL_Select = "Insert Into かんばん重複チェックテーブル  ( "
        gSL_Select = gSL_Select & "処理区分,納入日,納入便,"
        gSL_Select = gSL_Select & "取引先,品番,背番号,"
        gSL_Select = gSL_Select & "受入,連番,版数,"
        gSL_Select = gSL_Select & "作成端末,作成日"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "" & .連番 & ","
        gSL_Select = gSL_Select & "" & .版数 & ","
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "GETDATE()"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KJCInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KJCUpdate(Optional DBNo As Integer) As Boolean
    KJCUpdate = False
    With KJC
        gSL_Select = "Update かんばん重複チェックテーブル  Set "
        gSL_Select = gSL_Select & "処理区分 = '" & RTrim(.処理区分) & "',"
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "連番 = " & .連番 & ","
        gSL_Select = gSL_Select & "版数 = " & .版数 & ""
        gSL_Select = gSL_Select & " Where 処理区分 = '" & RTrim(.処理区分) & "'"
        gSL_Select = gSL_Select & "   And 納入日 = '" & RTrim(.納入日) & "'"
        gSL_Select = gSL_Select & "   And 納入便 = '" & RTrim(.納入便) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 背番号 = '" & RTrim(.背番号) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 連番 = " & .連番 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KJCUpdate = True
End Function

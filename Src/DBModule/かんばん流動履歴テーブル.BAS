Attribute VB_Name = "かんばん流動履歴テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   かんばん流動履歴テーブル
'*
'*   モジュールID  : かんばん流動履歴テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*
'*   作成日  :  2005/06/18  By Cis
'*   変更日  :  2006/12/20  項目追加（版数）
'*---------------------------------------------------------------------------------
    Type T_かんばん流動履歴テーブル
        処理区分  As String
        流動日    As String
        納入日    As String
        納入便    As String
        取引先    As String
        品番      As String
        背番号    As String
        受入      As String
        連番      As Integer
        版数      As Byte
        収容区分  As Byte
'        作成端末  As String
'        作成日    As Date
        件数          As Long
    End Type

    Public KRR As T_かんばん流動履歴テーブル

    Public KRR_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKRR(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKRR
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With KRR
            Select Case adoCol.Name
                Case "処理区分":  If Not IsNull(adoCol.Value) Then .処理区分 = adoCol.Value
                Case "流動日":    If Not IsNull(adoCol.Value) Then .流動日 = adoCol.Value
                Case "納入日":    If Not IsNull(adoCol.Value) Then .納入日 = adoCol.Value
                Case "納入便":    If Not IsNull(adoCol.Value) Then .納入便 = adoCol.Value
                Case "取引先":    If Not IsNull(adoCol.Value) Then .取引先 = adoCol.Value
                Case "品番":      If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "背番号":    If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "受入":      If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "連番":      If Not IsNull(adoCol.Value) Then .連番 = adoCol.Value
                Case "版数":      If Not IsNull(adoCol.Value) Then .版数 = adoCol.Value
                Case "収容区分":  If Not IsNull(adoCol.Value) Then .収容区分 = adoCol.Value
'                Case "作成端末":  If Not IsNull(adoCol.Value) Then .作成端末 = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KRRRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KRRRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From かんばん流動履歴テーブル"
        RTxtSql = RTxtSql & " Where 処理区分 = '" & RTrim(KRR.処理区分) & "'"
        RTxtSql = RTxtSql & " And   流動日 = '" & RTrim(KRR.流動日) & "'"
        RTxtSql = RTxtSql & " And   納入日 = '" & RTrim(KRR.納入日) & "'"
        RTxtSql = RTxtSql & " And   納入便 = '" & RTrim(KRR.納入便) & "'"
        RTxtSql = RTxtSql & " And   取引先 = '" & RTrim(KRR.取引先) & "'"
        RTxtSql = RTxtSql & " And   品番 = '" & RTrim(KRR.品番) & "'"
        RTxtSql = RTxtSql & " And   背番号 = '" & RTrim(KRR.背番号) & "'"
        RTxtSql = RTxtSql & " And   受入 = '" & RTrim(KRR.受入) & "'"
        RTxtSql = RTxtSql & " And   連番 = " & KRR.連番 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KRRRead = True
           Call ItemsFromKRR(DBNo)
           KRR_RDSTS = True
        Else
           KRR_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call KRRClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KRRReadNext(Optional DBNo As Integer = 0)
    If Not KRR_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KRR_RDSTS = False
        Else
           Call ItemsFromKRR(DBNo)
           KRR_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KRRClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KRR_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KRR )
'*----------------------------------------------*'
Public Sub ItemsClearKRR()
    With KRR
        .処理区分 = ""
        .流動日 = ""
        .納入日 = ""
        .納入便 = ""
        .取引先 = ""
        .品番 = ""
        .背番号 = ""
        .受入 = ""
        .連番 = 0
        .版数 = 0
        .収容区分 = 0
'       .作成端末 = ""
'       .作成日 = Nothing
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KRRInsert(Optional DBNo As Integer) As Boolean
    KRRInsert = True
    With KRR
        gSL_Select = "Insert Into かんばん流動履歴テーブル  ( "
        gSL_Select = gSL_Select & "処理区分,流動日,納入日,"
        gSL_Select = gSL_Select & "納入便,取引先,品番,"
        gSL_Select = gSL_Select & "背番号,受入,連番,版数,"
        gSL_Select = gSL_Select & "収容区分,"
        gSL_Select = gSL_Select & "作成端末,作成日"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.流動日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "" & .連番 & ","
        gSL_Select = gSL_Select & "" & .版数 & ","
        gSL_Select = gSL_Select & "" & .収容区分 & ","
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "GETDATE()"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KRRInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KRRUpdate(Optional DBNo As Integer) As Boolean
    KRRUpdate = False
    With KRR
        gSL_Select = "Update かんばん流動履歴テーブル  Set "
        gSL_Select = gSL_Select & "処理区分 = '" & RTrim(.処理区分) & "',"
        gSL_Select = gSL_Select & "流動日 = '" & RTrim(.流動日) & "',"
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "連番 = " & .連番 & ","
        gSL_Select = gSL_Select & "版数 = " & .版数 & ","
        gSL_Select = gSL_Select & "収容区分 = " & .収容区分 & ""
        gSL_Select = gSL_Select & " Where 処理区分 = '" & RTrim(.処理区分) & "'"
        gSL_Select = gSL_Select & "   And 流動日 = '" & RTrim(.流動日) & "'"
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
    KRRUpdate = True
End Function

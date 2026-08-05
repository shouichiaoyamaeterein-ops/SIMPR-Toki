Attribute VB_Name = "TOPPS変換履歴テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TOPPS変換履歴テーブル
'*
'*   モジュールID  : TOPPS変換履歴テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*
'*   作成日  :  2004/04/19  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_TOPPS変換履歴テーブル
        履歴管理番号          As Long
        受注出荷区分          As Byte
        ファイルID            As String
        開始日時              As Date
        終了日時              As Date
        入力件数              As Long
        変換件数              As Long
        変換エラー件数        As Long
        処理区分              As Byte
        削除区分              As Byte
        ERR                   As String
        ERR内容               As String
'        作成日                As Date
'        作成端末              As String
        件数          As Long
    End Type

    Public PHR As T_TOPPS変換履歴テーブル

    Public PHR_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
    
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromPHR(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearPHR
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With PHR
            Select Case adoCol.Name
                Case "履歴管理番号":          If Not IsNull(adoCol.Value) Then .履歴管理番号 = adoCol.Value
                Case "受注出荷区分":          If Not IsNull(adoCol.Value) Then .受注出荷区分 = adoCol.Value
                Case "ファイルID":            If Not IsNull(adoCol.Value) Then .ファイルID = adoCol.Value
                Case "開始日時":              If Not IsNull(adoCol.Value) Then .開始日時 = adoCol.Value
                Case "終了日時":              If Not IsNull(adoCol.Value) Then .終了日時 = adoCol.Value
                Case "入力件数":              If Not IsNull(adoCol.Value) Then .入力件数 = adoCol.Value
                Case "変換件数":              If Not IsNull(adoCol.Value) Then .変換件数 = adoCol.Value
                Case "変換エラー件数":        If Not IsNull(adoCol.Value) Then .変換エラー件数 = adoCol.Value
                Case "処理区分":              If Not IsNull(adoCol.Value) Then .処理区分 = adoCol.Value
                Case "削除区分":              If Not IsNull(adoCol.Value) Then .削除区分 = adoCol.Value
                Case "ERR":                   If Not IsNull(adoCol.Value) Then .ERR = adoCol.Value
                Case "ERR内容":               If Not IsNull(adoCol.Value) Then .ERR内容 = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function PHRRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    PHRRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From TOPPS変換履歴テーブル"
        RTxtSql = RTxtSql & " Where 履歴管理番号 = " & PHR.履歴管理番号 & ""
        RTxtSql = RTxtSql & " And   受注出荷区分 = " & PHR.受注出荷区分 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           PHRRead = True
           Call ItemsFromPHR(DBNo)
           PHR_RDSTS = True
        Else
           PHR_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call PHRClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub PHRReadNext(Optional DBNo As Integer = 0)
    If Not PHR_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           PHR_RDSTS = False
        Else
           Call ItemsFromPHR(DBNo)
           PHR_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub PHRClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        PHR_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( PHR )
'*----------------------------------------------*'
Public Sub ItemsClearPHR()
    With PHR
        .履歴管理番号 = 0
        .受注出荷区分 = 0
        .ファイルID = ""
        .開始日時 = 0
        .終了日時 = 0
        .入力件数 = 0
        .変換件数 = 0
        .変換エラー件数 = 0
        .処理区分 = 0
        .削除区分 = 0
        .ERR = ""
        .ERR内容 = ""
'       .作成日 = 0
'       .作成端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function PHRInsert(Optional DBNo As Integer) As Boolean
    PHRInsert = True
    With PHR
        gSL_Select = "Insert Into TOPPS変換履歴テーブル  ( "
        gSL_Select = gSL_Select & "履歴管理番号,受注出荷区分,ファイルID,"
        gSL_Select = gSL_Select & "開始日時,終了日時,入力件数,"
        gSL_Select = gSL_Select & "変換件数,変換エラー件数,処理区分,"
        gSL_Select = gSL_Select & "削除区分,ERR,ERR内容,"
        gSL_Select = gSL_Select & "作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .履歴管理番号 & ","
        gSL_Select = gSL_Select & "" & .受注出荷区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.ファイルID) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.開始日時) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.終了日時) & "',"
        gSL_Select = gSL_Select & "" & .入力件数 & ","
        gSL_Select = gSL_Select & "" & .変換件数 & ","
        gSL_Select = gSL_Select & "" & .変換エラー件数 & ","
        gSL_Select = gSL_Select & "" & .処理区分 & ","
        gSL_Select = gSL_Select & "" & .削除区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.ERR) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.ERR内容) & "',"
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    PHRInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function PHRUpdate(Optional DBNo As Integer) As Boolean
    PHRUpdate = False
    With PHR
        gSL_Select = "Update TOPPS変換履歴テーブル  Set "
        gSL_Select = gSL_Select & "履歴管理番号 = " & .履歴管理番号 & ","
        gSL_Select = gSL_Select & "受注出荷区分 = " & .受注出荷区分 & ","
        gSL_Select = gSL_Select & "ファイルID = '" & RTrim(.ファイルID) & "',"
        gSL_Select = gSL_Select & "開始日時 = '" & RTrim(.開始日時) & "',"
        gSL_Select = gSL_Select & "終了日時 = '" & RTrim(.終了日時) & "',"
        gSL_Select = gSL_Select & "入力件数 = " & .入力件数 & ","
        gSL_Select = gSL_Select & "変換件数 = " & .変換件数 & ","
        gSL_Select = gSL_Select & "変換エラー件数 = " & .変換エラー件数 & ","
        gSL_Select = gSL_Select & "処理区分 = " & .処理区分 & ","
        gSL_Select = gSL_Select & "削除区分 = " & .削除区分 & ","
        gSL_Select = gSL_Select & "ERR = '" & RTrim(.ERR) & "',"
        gSL_Select = gSL_Select & "ERR内容 = '" & RTrim(.ERR内容) & "'"
        gSL_Select = gSL_Select & " Where 履歴管理番号 = " & .履歴管理番号 & ""
        gSL_Select = gSL_Select & "   And 受注出荷区分 = " & .受注出荷区分 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    PHRUpdate = True
End Function

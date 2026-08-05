Attribute VB_Name = "購入単価内訳マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   購入単価内訳マスタ
'*
'*   モジュールID  : 購入単価内訳マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*                                      /項目追加含む/
'*   作成日  :  2004/11/25  By Cis
'*   作成日  :  2008/01/21  By Cis  契約先追加
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_購入単価内訳マスタ
        品番              As String
        取引先            As String
        契約先            As String
        履歴NO            As Long
        連番              As Byte
        内訳コード        As String
        内訳内容          As String
        単価              As Currency
'        作成日            As Date
        作成者            As String
'        作成端末          As String
'        更新日            As Date
        更新者            As String
'        更新端末          As String
        件数          As Long
'   項目追加(自動)
        取引先名       As String
    End Type

    Public KTU As T_購入単価内訳マスタ

    Public KTU_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKTU(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKTU
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With KTU
            Select Case adoCol.Name
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "取引先":            If Not IsNull(adoCol.Value) Then .取引先 = adoCol.Value
                Case "契約先":            If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
                Case "履歴NO":            If Not IsNull(adoCol.Value) Then .履歴NO = adoCol.Value
                Case "連番":              If Not IsNull(adoCol.Value) Then .連番 = adoCol.Value
                Case "内訳コード":        If Not IsNull(adoCol.Value) Then .内訳コード = adoCol.Value
                Case "内訳内容":          If Not IsNull(adoCol.Value) Then .内訳内容 = adoCol.Value
                Case "単価":              If Not IsNull(adoCol.Value) Then .単価 = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
                Case "更新者":            If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "取引先名":        If Not IsNull(adoCol.Value) Then .取引先名 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KTURead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KTURead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 購入単価内訳マスタ"
        RTxtSql = RTxtSql & " Where 品番 = '" & RTrim(KTU.品番) & "'"
        RTxtSql = RTxtSql & " And   取引先 = '" & RTrim(KTU.取引先) & "'"
        RTxtSql = RTxtSql & " And   契約先 = '" & RTrim(KTU.契約先) & "'"
        RTxtSql = RTxtSql & " And   履歴NO = " & KTU.履歴NO & ""
        RTxtSql = RTxtSql & " And   連番 = " & KTU.連番 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KTURead = True
           Call ItemsFromKTU(DBNo)
           KTU_RDSTS = True
        Else
           KTU_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call KTUClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KTUReadNext(Optional DBNo As Integer = 0)
    If Not KTU_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KTU_RDSTS = False
        Else
           Call ItemsFromKTU(DBNo)
           KTU_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KTUClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KTU_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KTU )
'*----------------------------------------------*'
Public Sub ItemsClearKTU()
    With KTU
        .品番 = ""
        .取引先 = ""
        .契約先 = ""
        .履歴NO = 0
        .連番 = 0
        .内訳コード = ""
        .内訳内容 = ""
        .単価 = 0
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
'       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
        .件数 = 0
'   項目追加(自動)
        .取引先名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KTUInsert(Optional DBNo As Integer) As Boolean
    KTUInsert = False
    With KTU
        gSL_Select = "Insert Into 購入単価内訳マスタ  ( "
        gSL_Select = gSL_Select & "品番,取引先,契約先,"
        gSL_Select = gSL_Select & "履歴NO,連番,内訳コード,"
        gSL_Select = gSL_Select & "内訳内容,単価,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "" & .履歴NO & ","
        gSL_Select = gSL_Select & "" & .連番 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.内訳コード) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.内訳内容) & "',"
        gSL_Select = gSL_Select & "" & .単価 & ","
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "',"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KTUInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KTUUpdate(Optional DBNo As Integer) As Boolean
    KTUUpdate = False
    With KTU
        gSL_Select = "Update 購入単価内訳マスタ  Set "
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "履歴NO = " & .履歴NO & ","
        gSL_Select = gSL_Select & "連番 = " & .連番 & ","
        gSL_Select = gSL_Select & "内訳コード = '" & RTrim(.内訳コード) & "',"
        gSL_Select = gSL_Select & "内訳内容 = '" & RTrim(.内訳内容) & "',"
        gSL_Select = gSL_Select & "単価 = " & .単価 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 契約先 = '" & RTrim(.契約先) & "'"
        gSL_Select = gSL_Select & "   And 履歴NO = " & .履歴NO & ""
        gSL_Select = gSL_Select & "   And 連番 = " & .連番 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KTUUpdate = True
End Function

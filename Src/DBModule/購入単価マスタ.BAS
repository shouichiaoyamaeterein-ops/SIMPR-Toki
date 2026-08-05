Attribute VB_Name = "購入単価マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   購入単価マスタ
'*
'*   モジュールID  : 購入単価マスタ.Bas
'*
'*   作成日  :  2004/11/25  By Cis
'*   作成日  :  2008/01/21  By Cis  契約先追加
'*   変更日  :  2009/10/09  By Cis  手作業にて項目追加
'*   　　　　　　　　　　　　　　　(コメント「自動生成」は削除)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_購入単価マスタ
        品番          As String
        取引先        As String
        契約先        As String
        履歴NO        As Long
        仮単価区分    As String
        単価          As Currency
        適用開始日付  As String
        適用終了日付  As String
        有無償区分    As Byte
        部品単価      As Currency
        加工単価      As Currency
        その他単価    As Currency
'        作成日        As Date
        作成者        As String
'        作成端末      As String
'        更新日        As Date
        更新者        As String
'        更新端末      As String
        件数          As Long
'   項目追加(自動)
        取引先名       As String
'-------------------------------------------------------' 2009.10.09 add start : 手作業で追加
        背番号         As String
        表示品番       As String
        前回履歴NO     As Long
        前回仮単価区分 As String
        前回単価       As Currency
        材料管理番号   As String
        前回有無賞区分 As String
'
        発注品番       As String
        受給品番       As String
        自社品番       As String
        材質           As String
        板厚           As Currency
        幅             As Currency
        長さ           As Currency
'-------------------------------------------------------' 2009.10.09 add end
    End Type

    Public KTM As T_購入単価マスタ

    Public KTM_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKTM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKTM
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With KTM
            Select Case adoCol.Name
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "取引先":        If Not IsNull(adoCol.Value) Then .取引先 = adoCol.Value
                Case "契約先":        If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
                Case "履歴NO":        If Not IsNull(adoCol.Value) Then .履歴NO = adoCol.Value
                Case "仮単価区分":    If Not IsNull(adoCol.Value) Then .仮単価区分 = adoCol.Value
                Case "単価":          If Not IsNull(adoCol.Value) Then .単価 = adoCol.Value
                Case "適用開始日付":  If Not IsNull(adoCol.Value) Then .適用開始日付 = adoCol.Value
                Case "適用終了日付":  If Not IsNull(adoCol.Value) Then .適用終了日付 = adoCol.Value
                Case "有無償区分":    If Not IsNull(adoCol.Value) Then .有無償区分 = adoCol.Value
                Case "部品単価":      If Not IsNull(adoCol.Value) Then .部品単価 = adoCol.Value
                Case "加工単価":      If Not IsNull(adoCol.Value) Then .加工単価 = adoCol.Value
                Case "その他単価":    If Not IsNull(adoCol.Value) Then .その他単価 = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "取引先名":        If Not IsNull(adoCol.Value) Then .取引先名 = adoCol.Value
'-------------------------------------------------------------------------------------------------------2009.10.09 add start : 手作業で追加
                Case "背番号":  If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "表示品番":        If Not IsNull(adoCol.Value) Then .表示品番 = adoCol.Value
                Case "前回履歴NO":      If Not IsNull(adoCol.Value) Then .前回履歴NO = adoCol.Value
                Case "前回仮単価区分":  If Not IsNull(adoCol.Value) Then .前回仮単価区分 = adoCol.Value
                Case "前回単価":        If Not IsNull(adoCol.Value) Then .前回単価 = adoCol.Value
                Case "材料管理番号":    If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "前回有無賞区分":  If Not IsNull(adoCol.Value) Then .前回有無賞区分 = adoCol.Value
'
                Case "発注品番":        If Not IsNull(adoCol.Value) Then .発注品番 = adoCol.Value
                Case "受給品番":        If Not IsNull(adoCol.Value) Then .受給品番 = adoCol.Value
                Case "自社品番":        If Not IsNull(adoCol.Value) Then .自社品番 = adoCol.Value
                Case "材質":            If Not IsNull(adoCol.Value) Then .材質 = adoCol.Value
                Case "板厚":            If Not IsNull(adoCol.Value) Then .板厚 = adoCol.Value
                Case "幅":              If Not IsNull(adoCol.Value) Then .幅 = adoCol.Value
                Case "長さ":            If Not IsNull(adoCol.Value) Then .長さ = adoCol.Value
'-------------------------------------------------------------------------------------------------------' 2009.10.09 add end
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KTMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KTMRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 購入単価マスタ"
        RTxtSql = RTxtSql & " Where 品番 = '" & RTrim(KTM.品番) & "'"
        RTxtSql = RTxtSql & " And   取引先 = '" & RTrim(KTM.取引先) & "'"
        RTxtSql = RTxtSql & " And   契約先 = '" & RTrim(KTM.契約先) & "'"
        RTxtSql = RTxtSql & " And   履歴NO = " & KTM.履歴NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KTMRead = True
           Call ItemsFromKTM(DBNo)
           KTM_RDSTS = True
        Else
           KTM_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call KTMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KTMReadNext(Optional DBNo As Integer = 0)
    If Not KTM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KTM_RDSTS = False
        Else
           Call ItemsFromKTM(DBNo)
           KTM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KTMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KTM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KTM )
'*----------------------------------------------*'
Public Sub ItemsClearKTM()
    With KTM
        .品番 = ""
        .取引先 = ""
        .契約先 = ""
        .履歴NO = 0
        .仮単価区分 = ""
        .単価 = 0
        .適用開始日付 = ""
        .適用終了日付 = ""
        .有無償区分 = 0
        .部品単価 = 0
        .加工単価 = 0
        .その他単価 = 0
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
'       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
        .件数 = 0
'   項目追加(自動)
        .取引先名 = ""
'-------------------------------------------------------' 2009.10.09 add start : 手作業で追加
        .背番号 = ""
        .表示品番 = ""
        .前回履歴NO = 0
        .前回仮単価区分 = ""
        .前回単価 = 0
        .材料管理番号 = ""
        .前回有無賞区分 = ""
'   さらに追加
        .発注品番 = ""
        .受給品番 = ""
        .自社品番 = ""
        .材質 = ""
        .板厚 = 0
        .幅 = 0
        .長さ = 0
'-------------------------------------------------------' 2009.10.09 add end
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KTMInsert(Optional DBNo As Integer) As Boolean
    KTMInsert = False
    With KTM
        gSL_Select = "Insert Into 購入単価マスタ  ( "
        gSL_Select = gSL_Select & "品番,取引先,契約先,"
        gSL_Select = gSL_Select & "履歴NO,仮単価区分,単価,"
        gSL_Select = gSL_Select & "適用開始日付,適用終了日付,有無償区分,"
        gSL_Select = gSL_Select & "部品単価,加工単価,その他単価,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "" & .履歴NO & ","
        gSL_Select = gSL_Select & "'" & RTrim(.仮単価区分) & "',"
        gSL_Select = gSL_Select & "" & .単価 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.適用開始日付) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.適用終了日付) & "',"
        gSL_Select = gSL_Select & "" & .有無償区分 & ","
        gSL_Select = gSL_Select & "" & .部品単価 & ","
        gSL_Select = gSL_Select & "" & .加工単価 & ","
        gSL_Select = gSL_Select & "" & .その他単価 & ","
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
    KTMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KTMUpdate(Optional DBNo As Integer) As Boolean
    KTMUpdate = False
    With KTM
        gSL_Select = "Update 購入単価マスタ  Set "
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "履歴NO = " & .履歴NO & ","
        gSL_Select = gSL_Select & "仮単価区分 = '" & RTrim(.仮単価区分) & "',"
        gSL_Select = gSL_Select & "単価 = " & .単価 & ","
        gSL_Select = gSL_Select & "適用開始日付 = '" & RTrim(.適用開始日付) & "',"
        gSL_Select = gSL_Select & "適用終了日付 = '" & RTrim(.適用終了日付) & "',"
        gSL_Select = gSL_Select & "有無償区分 = " & .有無償区分 & ","
        gSL_Select = gSL_Select & "部品単価 = " & .部品単価 & ","
        gSL_Select = gSL_Select & "加工単価 = " & .加工単価 & ","
        gSL_Select = gSL_Select & "その他単価 = " & .その他単価 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 契約先 = '" & RTrim(.契約先) & "'"
        gSL_Select = gSL_Select & "   And 履歴NO = " & .履歴NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KTMUpdate = True
End Function

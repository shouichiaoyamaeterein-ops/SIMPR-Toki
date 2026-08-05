Attribute VB_Name = "材料支給単価マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料支給単価マスタ
'*
'*   モジュールID  : 材料支給単価マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*
'*   作成日  :  2008/02/29  By Cis
'*   作成日  :  2008/03/01  By Cis
'*   作成日  :  2008/03/05  By Cis
'*   作成日  :  2009/09/28  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料支給単価マスタ
        材料管理番号  As Long
        取引先        As String
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
'-----------------------------------------------' 仮追加(手動) start
        背番号         As String
        取引先名       As String
        発注品番       As String
        受給品番       As String
        自社品番       As String
        材質           As String
        板厚           As Currency
        幅             As Currency
        長さ           As Currency
'-----------------------------------------------' 仮追加(手動) end
    End Type

    Public ZSM As T_材料支給単価マスタ

    Public ZSM_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZSM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZSM
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With ZSM
            Select Case adoCol.Name
                Case "材料管理番号":  If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "取引先":        If Not IsNull(adoCol.Value) Then .取引先 = adoCol.Value
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

'-----------------------------------------------' 仮追加(手動) start
                Case "背番号":  If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "取引先名":        If Not IsNull(adoCol.Value) Then .取引先名 = adoCol.Value
                Case "発注品番":        If Not IsNull(adoCol.Value) Then .発注品番 = adoCol.Value
                Case "受給品番":        If Not IsNull(adoCol.Value) Then .受給品番 = adoCol.Value
                Case "自社品番":        If Not IsNull(adoCol.Value) Then .自社品番 = adoCol.Value
                Case "材質":    If Not IsNull(adoCol.Value) Then .材質 = adoCol.Value
                Case "板厚":    If Not IsNull(adoCol.Value) Then .板厚 = adoCol.Value
                Case "幅":      If Not IsNull(adoCol.Value) Then .幅 = adoCol.Value
                Case "長さ":    If Not IsNull(adoCol.Value) Then .長さ = adoCol.Value
'-----------------------------------------------' 仮追加(手動) end
            
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZSMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZSMRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 材料支給単価マスタ"
        RTxtSql = RTxtSql & " Where 材料管理番号 = " & ZSM.材料管理番号 & ""
        RTxtSql = RTxtSql & " And   取引先 = '" & RTrim(ZSM.取引先) & "'"
        RTxtSql = RTxtSql & " And   履歴NO = " & ZSM.履歴NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZSMRead = True
           Call ItemsFromZSM(DBNo)
           ZSM_RDSTS = True
        Else
           ZSM_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call ZSMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZSMReadNext(Optional DBNo As Integer = 0)
    If Not ZSM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZSM_RDSTS = False
        Else
           Call ItemsFromZSM(DBNo)
           ZSM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZSMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZSM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZSM )
'*----------------------------------------------*'
Public Sub ItemsClearZSM()
    With ZSM
        .材料管理番号 = 0
        .取引先 = ""
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
'-----------------------------------------------' 仮追加(手動) start
        .背番号 = ""
        .取引先名 = ""
        .発注品番 = ""
        .受給品番 = ""
        .自社品番 = ""
        .材質 = ""
        .板厚 = 0
        .幅 = 0
        .長さ = 0
'-----------------------------------------------' 仮追加(手動) end
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZSMInsert(Optional DBNo As Integer) As Boolean
    ZSMInsert = False
    With ZSM
        gSL_Select = "Insert Into 材料支給単価マスタ  ( "
        gSL_Select = gSL_Select & "材料管理番号,取引先,履歴NO,"
        gSL_Select = gSL_Select & "仮単価区分,単価,適用開始日付,"
        gSL_Select = gSL_Select & "適用終了日付,有無償区分,部品単価,"
        gSL_Select = gSL_Select & "加工単価,その他単価,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"
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
    ZSMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZSMUpdate(Optional DBNo As Integer) As Boolean
    ZSMUpdate = False
    With ZSM
        gSL_Select = "Update 材料支給単価マスタ  Set "
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
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
        gSL_Select = gSL_Select & " Where 材料管理番号 = " & .材料管理番号 & ""
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 履歴NO = " & .履歴NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZSMUpdate = True
End Function

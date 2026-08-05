Attribute VB_Name = "材料在庫履歴マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料在庫履歴マスタ
'*
'*   モジュールID  : 材料在庫履歴マスタ.Bas
'*
'*
'*   作成日  :  2004/02/21  By Cis
'*   変更日  :  2010/01/14 ZZRInsert2 追加
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料在庫履歴マスタ
        年月度        As String
        材料管理番号  As Long
        材質          As String
        寸法区分      As Byte
        板厚          As Currency
        幅            As Currency
        長さ          As Currency
        入力日付      As Date
        入力担当者    As String
        在庫数        As Currency
        実棚日        As String
        変更事由      As String
        作成区分      As Byte
        履歴区分      As Byte
'        作成日        As Date
        作成者        As String
'        作成端末      As String
        件数          As Long
    End Type

    Public ZZR As T_材料在庫履歴マスタ

    Public ZZR_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZZR(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZZR
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With ZZR
            Select Case adoCol.Name
                Case "年月度":        If Not IsNull(adoCol.Value) Then .年月度 = adoCol.Value
                Case "材料管理番号":  If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "材質":          If Not IsNull(adoCol.Value) Then .材質 = adoCol.Value
                Case "寸法区分":      If Not IsNull(adoCol.Value) Then .寸法区分 = adoCol.Value
                Case "板厚":          If Not IsNull(adoCol.Value) Then .板厚 = adoCol.Value
                Case "幅":            If Not IsNull(adoCol.Value) Then .幅 = adoCol.Value
                Case "長さ":          If Not IsNull(adoCol.Value) Then .長さ = adoCol.Value
                Case "入力日付":      If Not IsNull(adoCol.Value) Then .入力日付 = adoCol.Value
                Case "入力担当者":    If Not IsNull(adoCol.Value) Then .入力担当者 = adoCol.Value
                Case "在庫数":        If Not IsNull(adoCol.Value) Then .在庫数 = adoCol.Value
                Case "実棚日":        If Not IsNull(adoCol.Value) Then .実棚日 = adoCol.Value
                Case "変更事由":      If Not IsNull(adoCol.Value) Then .変更事由 = adoCol.Value
                Case "作成区分":      If Not IsNull(adoCol.Value) Then .作成区分 = adoCol.Value
                Case "履歴区分":      If Not IsNull(adoCol.Value) Then .履歴区分 = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZZRRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZZRRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 材料在庫履歴マスタ"
        RTxtSql = RTxtSql & " Where 年月度 = '" & RTrim(ZZR.年月度) & "'"
        RTxtSql = RTxtSql & " And   材料管理番号 = " & ZZR.材料管理番号 & ""
        RTxtSql = RTxtSql & " And   入力日付 = '" & RTrim(ZZR.入力日付) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZZRRead = True
           Call ItemsFromZZR(DBNo)
           ZZR_RDSTS = True
        Else
           ZZR_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call ZZRClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZZRReadNext(Optional DBNo As Integer = 0)
    If Not ZZR_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZZR_RDSTS = False
        Else
           Call ItemsFromZZR(DBNo)
           ZZR_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZZRClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZZR_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZZR )
'*----------------------------------------------*'
Public Sub ItemsClearZZR()
    With ZZR
        .年月度 = ""
        .材料管理番号 = 0
        .材質 = ""
        .寸法区分 = 0
        .板厚 = 0
        .幅 = 0
        .長さ = 0
        .入力日付 = 0
        .入力担当者 = ""
        .在庫数 = 0
        .実棚日 = ""
        .変更事由 = ""
        .作成区分 = 0
        .履歴区分 = 0
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZZRInsert(Optional DBNo As Integer) As Boolean
    ZZRInsert = True
    With ZZR
        gSL_Select = "Insert Into 材料在庫履歴マスタ  ( "
        gSL_Select = gSL_Select & "年月度,材料管理番号,材質,"
        gSL_Select = gSL_Select & "寸法区分,板厚,幅,"
        gSL_Select = gSL_Select & "長さ,入力日付,入力担当者,"
        gSL_Select = gSL_Select & "在庫数,実棚日,変更事由,"
        gSL_Select = gSL_Select & "作成区分,履歴区分,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.年月度) & "',"
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "" & .寸法区分 & ","
        gSL_Select = gSL_Select & "" & .板厚 & ","
        gSL_Select = gSL_Select & "" & .幅 & ","
        gSL_Select = gSL_Select & "" & .長さ & ","
        gSL_Select = gSL_Select & "'" & RTrim(.入力日付) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.入力担当者) & "',"
        gSL_Select = gSL_Select & "" & .在庫数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.実棚日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.変更事由) & "',"
        gSL_Select = gSL_Select & "" & .作成区分 & ","
        gSL_Select = gSL_Select & "" & .履歴区分 & ","
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
    ZZRInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZZRUpdate(Optional DBNo As Integer) As Boolean
    ZZRUpdate = False
    With ZZR
        gSL_Select = "Update 材料在庫履歴マスタ  Set "
        gSL_Select = gSL_Select & "年月度 = '" & RTrim(.年月度) & "',"
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "寸法区分 = " & .寸法区分 & ","
        gSL_Select = gSL_Select & "板厚 = " & .板厚 & ","
        gSL_Select = gSL_Select & "幅 = " & .幅 & ","
        gSL_Select = gSL_Select & "長さ = " & .長さ & ","
        gSL_Select = gSL_Select & "入力日付 = '" & RTrim(.入力日付) & "',"
        gSL_Select = gSL_Select & "入力担当者 = '" & RTrim(.入力担当者) & "',"
        gSL_Select = gSL_Select & "在庫数 = " & .在庫数 & ","
        gSL_Select = gSL_Select & "実棚日 = '" & RTrim(.実棚日) & "',"
        gSL_Select = gSL_Select & "変更事由 = '" & RTrim(.変更事由) & "',"
        gSL_Select = gSL_Select & "作成区分 = " & .作成区分 & ","
        gSL_Select = gSL_Select & "履歴区分 = " & .履歴区分 & ""
        gSL_Select = gSL_Select & " Where 年月度 = '" & RTrim(.年月度) & "'"
        gSL_Select = gSL_Select & "   And 材料管理番号 = " & .材料管理番号 & ""
        gSL_Select = gSL_Select & "   And 入力日付 = '" & RTrim(.入力日付) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZZRUpdate = True
End Function

'*--------------------*'
'*      追加処理2       *
'*--------------------*'
Public Function ZZRInsert2(Optional DBNo As Integer, Optional Counter As Integer) As Boolean
    ZZRInsert2 = True
    With ZZR
        gSL_Select = "Insert Into 材料在庫履歴マスタ  ( "
        gSL_Select = gSL_Select & "年月度,材料管理番号,材質,"
        gSL_Select = gSL_Select & "寸法区分,板厚,幅,"
        gSL_Select = gSL_Select & "長さ,入力日付,入力担当者,"
        gSL_Select = gSL_Select & "在庫数,実棚日,変更事由,"
        gSL_Select = gSL_Select & "作成区分,履歴区分,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.年月度) & "',"
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "" & .寸法区分 & ","
        gSL_Select = gSL_Select & "" & .板厚 & ","
        gSL_Select = gSL_Select & "" & .幅 & ","
        gSL_Select = gSL_Select & "" & .長さ & ","
        gSL_Select = gSL_Select & "'" & RTrim(Format(.入力日付, "YYYY-MM-DD HH:mm:ss")) & "." & Right("00" & CStr(Counter) & "0", 3) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.入力担当者) & "',"
        gSL_Select = gSL_Select & "" & .在庫数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.実棚日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.変更事由) & "',"
        gSL_Select = gSL_Select & "" & .作成区分 & ","
        gSL_Select = gSL_Select & "" & .履歴区分 & ","
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
    ZZRInsert2 = True
End Function

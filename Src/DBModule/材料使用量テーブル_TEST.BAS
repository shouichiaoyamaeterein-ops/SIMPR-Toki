Attribute VB_Name = "材料使用量テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料使用量テーブル
'*
'*   モジュールID  : 材料使用量テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*                                      /項目追加含む/
'*   作成日  :  2009/12/11  By Cis(FMVW5220)
'*   作成日  :  2009/12/17  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料使用量テーブル
        材料使用量管理NO  As Long
        材料使用量管理行  As Long
        作業日            As String
        部門              As String
        作業者            As String
        品番              As String
        材質              As String
        板厚              As Currency
        幅                As Currency
        長さ              As Currency
        材料管理番号      As Long
        加工数            As Long
        単重              As Currency
        使用量            As Currency
        ラベル重量1       As Currency
        ラベル本数1       As Byte
        ラベル重量2       As Currency
        ラベル本数2       As Byte
        ラベル重量3       As Currency
        ラベル本数3       As Byte
        ラベル重量4       As Currency
        ラベル本数4       As Byte
        戻し重量          As Byte
        合計重量          As Currency
'        作成日            As Date
        作成者            As String
'        作成端末          As String
'        更新日            As Date
        更新者            As String
'        更新端末          As String
        件数          As Long
'   項目追加(自動)
        品名           As String
        部門名         As String
        表示品番       As String
        表示材料       As String
'   項目追加(手動) ** 試験のため作成 **********
        表示加工数     As String
        表示単重       As String
        表示使用重量   As String
        表示ラベル1    As String
        表示ラベル2    As String
        表示ラベル3    As String
        表示ラベル4    As String
        表示戻し       As String
        表示合計重量   As String
    End Type

    Public ZSR As T_材料使用量テーブル

    Public ZSR_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZSR(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZSR
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With ZSR
            Select Case adoCol.Name
                Case "材料使用量管理NO":  If Not IsNull(adoCol.Value) Then .材料使用量管理NO = adoCol.Value
                Case "材料使用量管理行":  If Not IsNull(adoCol.Value) Then .材料使用量管理行 = adoCol.Value
                Case "作業日":            If Not IsNull(adoCol.Value) Then .作業日 = adoCol.Value
                Case "部門":              If Not IsNull(adoCol.Value) Then .部門 = adoCol.Value
                Case "作業者":            If Not IsNull(adoCol.Value) Then .作業者 = adoCol.Value
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "材質":              If Not IsNull(adoCol.Value) Then .材質 = adoCol.Value
                Case "板厚":              If Not IsNull(adoCol.Value) Then .板厚 = adoCol.Value
                Case "幅":                If Not IsNull(adoCol.Value) Then .幅 = adoCol.Value
                Case "長さ":              If Not IsNull(adoCol.Value) Then .長さ = adoCol.Value
                Case "材料管理番号":      If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "加工数":            If Not IsNull(adoCol.Value) Then .加工数 = adoCol.Value
                Case "単重":              If Not IsNull(adoCol.Value) Then .単重 = adoCol.Value
                Case "使用量":            If Not IsNull(adoCol.Value) Then .使用量 = adoCol.Value
                Case "ラベル重量1":       If Not IsNull(adoCol.Value) Then .ラベル重量1 = adoCol.Value
                Case "ラベル本数1":       If Not IsNull(adoCol.Value) Then .ラベル本数1 = adoCol.Value
                Case "ラベル重量2":       If Not IsNull(adoCol.Value) Then .ラベル重量2 = adoCol.Value
                Case "ラベル本数2":       If Not IsNull(adoCol.Value) Then .ラベル本数2 = adoCol.Value
                Case "ラベル重量3":       If Not IsNull(adoCol.Value) Then .ラベル重量3 = adoCol.Value
                Case "ラベル本数3":       If Not IsNull(adoCol.Value) Then .ラベル本数3 = adoCol.Value
                Case "ラベル重量4":       If Not IsNull(adoCol.Value) Then .ラベル重量4 = adoCol.Value
                Case "ラベル本数4":       If Not IsNull(adoCol.Value) Then .ラベル本数4 = adoCol.Value
                Case "戻し重量":          If Not IsNull(adoCol.Value) Then .戻し重量 = adoCol.Value
                Case "合計重量":          If Not IsNull(adoCol.Value) Then .合計重量 = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
                Case "更新者":            If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "品名":    If Not IsNull(adoCol.Value) Then .品名 = adoCol.Value
                Case "部門名":  If Not IsNull(adoCol.Value) Then .部門名 = adoCol.Value
                Case "表示品番":        If Not IsNull(adoCol.Value) Then .表示品番 = adoCol.Value
                Case "表示材料":        If Not IsNull(adoCol.Value) Then .表示材料 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZSRRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZSRRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 材料使用量テーブル"
        RTxtSql = RTxtSql & " Where 材料使用量管理NO = " & ZSR.材料使用量管理NO & ""
        RTxtSql = RTxtSql & " And   材料使用量管理行 = " & ZSR.材料使用量管理行 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZSRRead = True
           Call ItemsFromZSR(DBNo)
           ZSR_RDSTS = True
        Else
           ZSR_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call ZSRClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZSRReadNext(Optional DBNo As Integer = 0)
    If Not ZSR_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZSR_RDSTS = False
        Else
           Call ItemsFromZSR(DBNo)
           ZSR_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZSRClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZSR_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZSR )
'*----------------------------------------------*'
Public Sub ItemsClearZSR()
    With ZSR
        .材料使用量管理NO = 0
        .材料使用量管理行 = 0
        .作業日 = ""
        .部門 = ""
        .作業者 = ""
        .品番 = ""
        .材質 = ""
        .板厚 = 0
        .幅 = 0
        .長さ = 0
        .材料管理番号 = 0
        .加工数 = 0
        .単重 = 0
        .使用量 = 0
        .ラベル重量1 = 0
        .ラベル本数1 = 0
        .ラベル重量2 = 0
        .ラベル本数2 = 0
        .ラベル重量3 = 0
        .ラベル本数3 = 0
        .ラベル重量4 = 0
        .ラベル本数4 = 0
        .戻し重量 = 0
        .合計重量 = 0
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
'       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
        .件数 = 0
'   項目追加(自動)
        .品名 = ""
        .部門名 = ""
        .表示品番 = ""
        .表示材料 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZSRInsert(Optional DBNo As Integer) As Boolean
    ZSRInsert = False
    With ZSR
        gSL_Select = "Insert Into 材料使用量テーブル  ( "
        gSL_Select = gSL_Select & "材料使用量管理NO,材料使用量管理行,作業日,"
        gSL_Select = gSL_Select & "部門,作業者,品番,"
        gSL_Select = gSL_Select & "材質,板厚,幅,"
        gSL_Select = gSL_Select & "長さ,材料管理番号,加工数,"
        gSL_Select = gSL_Select & "単重,使用量,ラベル重量1,"
        gSL_Select = gSL_Select & "ラベル本数1,ラベル重量2,ラベル本数2,"
        gSL_Select = gSL_Select & "ラベル重量3,ラベル本数3,ラベル重量4,"
        gSL_Select = gSL_Select & "ラベル本数4,戻し重量,合計重量,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .材料使用量管理NO & ","
        gSL_Select = gSL_Select & "" & .材料使用量管理行 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.作業日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.部門) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.作業者) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "" & .板厚 & ","
        gSL_Select = gSL_Select & "" & .幅 & ","
        gSL_Select = gSL_Select & "" & .長さ & ","
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","
        gSL_Select = gSL_Select & "" & .加工数 & ","
        gSL_Select = gSL_Select & "" & .単重 & ","
        gSL_Select = gSL_Select & "" & .使用量 & ","
        gSL_Select = gSL_Select & "" & .ラベル重量1 & ","
        gSL_Select = gSL_Select & "" & .ラベル本数1 & ","
        gSL_Select = gSL_Select & "" & .ラベル重量2 & ","
        gSL_Select = gSL_Select & "" & .ラベル本数2 & ","
        gSL_Select = gSL_Select & "" & .ラベル重量3 & ","
        gSL_Select = gSL_Select & "" & .ラベル本数3 & ","
        gSL_Select = gSL_Select & "" & .ラベル重量4 & ","
        gSL_Select = gSL_Select & "" & .ラベル本数4 & ","
        gSL_Select = gSL_Select & "" & .戻し重量 & ","
        gSL_Select = gSL_Select & "" & .合計重量 & ","
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
    ZSRInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZSRUpdate(Optional DBNo As Integer) As Boolean
    ZSRUpdate = False
    With ZSR
        gSL_Select = "Update 材料使用量テーブル  Set "
        gSL_Select = gSL_Select & "材料使用量管理NO = " & .材料使用量管理NO & ","
        gSL_Select = gSL_Select & "材料使用量管理行 = " & .材料使用量管理行 & ","
        gSL_Select = gSL_Select & "作業日 = '" & RTrim(.作業日) & "',"
        gSL_Select = gSL_Select & "部門 = '" & RTrim(.部門) & "',"
        gSL_Select = gSL_Select & "作業者 = '" & RTrim(.作業者) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "板厚 = " & .板厚 & ","
        gSL_Select = gSL_Select & "幅 = " & .幅 & ","
        gSL_Select = gSL_Select & "長さ = " & .長さ & ","
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "加工数 = " & .加工数 & ","
        gSL_Select = gSL_Select & "単重 = " & .単重 & ","
        gSL_Select = gSL_Select & "使用量 = " & .使用量 & ","
        gSL_Select = gSL_Select & "ラベル重量1 = " & .ラベル重量1 & ","
        gSL_Select = gSL_Select & "ラベル本数1 = " & .ラベル本数1 & ","
        gSL_Select = gSL_Select & "ラベル重量2 = " & .ラベル重量2 & ","
        gSL_Select = gSL_Select & "ラベル本数2 = " & .ラベル本数2 & ","
        gSL_Select = gSL_Select & "ラベル重量3 = " & .ラベル重量3 & ","
        gSL_Select = gSL_Select & "ラベル本数3 = " & .ラベル本数3 & ","
        gSL_Select = gSL_Select & "ラベル重量4 = " & .ラベル重量4 & ","
        gSL_Select = gSL_Select & "ラベル本数4 = " & .ラベル本数4 & ","
        gSL_Select = gSL_Select & "戻し重量 = " & .戻し重量 & ","
        gSL_Select = gSL_Select & "合計重量 = " & .合計重量 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 材料使用量管理NO = " & .材料使用量管理NO & ""
        gSL_Select = gSL_Select & "   And 材料使用量管理行 = " & .材料使用量管理行 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZSRUpdate = True
End Function

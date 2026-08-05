Attribute VB_Name = "材料在庫マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料在庫マスタ
'*
'*   モジュールID  : 材料在庫マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*
'*   作成日  :  2004/02/21  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料在庫マスタ
        年月度        As String
        材料管理番号  As Long
        材質          As String
        寸法区分      As Byte
        板厚          As Currency
        幅            As Currency
        長さ          As Currency
        在庫数        As Currency
        実棚日        As String
        変更事由      As String
        SIZE            As String
        作成区分        As Byte
'        作成日        As Date
        作成者        As String
'        作成端末      As String
        件数          As Long
    End Type

    Public ZZM As T_材料在庫マスタ

    Public ZZM_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZZM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZZM
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With ZZM
            Select Case adoCol.Name
                Case "年月度":        If Not IsNull(adoCol.Value) Then .年月度 = adoCol.Value
                Case "材料管理番号":  If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "材質":          If Not IsNull(adoCol.Value) Then .材質 = adoCol.Value
                Case "寸法区分":      If Not IsNull(adoCol.Value) Then .寸法区分 = adoCol.Value
                Case "板厚":          If Not IsNull(adoCol.Value) Then .板厚 = adoCol.Value
                Case "幅":            If Not IsNull(adoCol.Value) Then .幅 = adoCol.Value
                Case "長さ":          If Not IsNull(adoCol.Value) Then .長さ = adoCol.Value
                Case "在庫数":        If Not IsNull(adoCol.Value) Then .在庫数 = adoCol.Value
                Case "実棚日":        If Not IsNull(adoCol.Value) Then .実棚日 = adoCol.Value
                Case "変更事由":      If Not IsNull(adoCol.Value) Then .変更事由 = adoCol.Value
                Case "SIZE":      If Not IsNull(adoCol.Value) Then .SIZE = adoCol.Value
                Case "作成区分":      If Not IsNull(adoCol.Value) Then .作成区分 = adoCol.Value
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
Public Function ZZMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZZMRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 材料在庫マスタ"
        RTxtSql = RTxtSql & " Where 年月度 = '" & RTrim(ZZM.年月度) & "'"
        RTxtSql = RTxtSql & " And   材料管理番号 = " & ZZM.材料管理番号 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZZMRead = True
           Call ItemsFromZZM(DBNo)
           ZZM_RDSTS = True
        Else
           ZZM_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call ZZMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZZMReadNext(Optional DBNo As Integer = 0)
    If Not ZZM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZZM_RDSTS = False
        Else
           Call ItemsFromZZM(DBNo)
           ZZM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZZMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZZM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZZM )
'*----------------------------------------------*'
Public Sub ItemsClearZZM()
    With ZZM
        .年月度 = ""
        .材料管理番号 = 0
        .材質 = ""
        .寸法区分 = 0
        .板厚 = 0
        .幅 = 0
        .長さ = 0
        .在庫数 = 0
        .実棚日 = ""
        .変更事由 = ""
        .SIZE = ""
        .作成区分 = 0
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZZMInsert(Optional DBNo As Integer) As Boolean
    ZZMInsert = True
    With ZZM
        gSL_Select = "Insert Into 材料在庫マスタ  ( "
        gSL_Select = gSL_Select & "年月度,材料管理番号,材質,"
        gSL_Select = gSL_Select & "寸法区分,板厚,幅,"
        gSL_Select = gSL_Select & "長さ,在庫数,実棚日,"
        gSL_Select = gSL_Select & "変更事由,SIZE,作成区分,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.年月度) & "',"
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "" & .寸法区分 & ","
        gSL_Select = gSL_Select & "" & .板厚 & ","
        gSL_Select = gSL_Select & "" & .幅 & ","
        gSL_Select = gSL_Select & "" & .長さ & ","
        gSL_Select = gSL_Select & "" & .在庫数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.実棚日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.変更事由) & "',"
        gSL_Select = gSL_Select & "'" & Format(.板厚, "0000.00")
        gSL_Select = gSL_Select & "" & Format(.幅, "0000.00")
        gSL_Select = gSL_Select & "" & Format(.長さ, "0000.00") & "',"
        gSL_Select = gSL_Select & "" & .作成区分 & ","
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
    ZZMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZZMUpdate(Optional DBNo As Integer) As Boolean
    ZZMUpdate = False
    With ZZM
        gSL_Select = "Update 材料在庫マスタ  Set "
        gSL_Select = gSL_Select & "年月度 = '" & RTrim(.年月度) & "',"
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "寸法区分 = " & .寸法区分 & ","
        gSL_Select = gSL_Select & "板厚 = " & .板厚 & ","
        gSL_Select = gSL_Select & "幅 = " & .幅 & ","
        gSL_Select = gSL_Select & "長さ = " & .長さ & ","
        gSL_Select = gSL_Select & "在庫数 = " & .在庫数 & ","
        gSL_Select = gSL_Select & "実棚日 = '" & RTrim(.実棚日) & "',"
        gSL_Select = gSL_Select & "変更事由 = '" & RTrim(.変更事由) & "'"
        gSL_Select = gSL_Select & "作成区分 = " & .作成区分 & ","
        gSL_Select = gSL_Select & " Where 年月度 = '" & RTrim(.年月度) & "'"
        gSL_Select = gSL_Select & "   And 材料管理番号 = " & .材料管理番号 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZZMUpdate = True
End Function

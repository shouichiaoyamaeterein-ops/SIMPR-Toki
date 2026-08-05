Attribute VB_Name = "品番工順マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   品番工順マスタ
'*
'*   モジュールID  : 品番工順マスタ.Bas
'*
'*
'*   作成日  :  2009/07/24  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_品番工順マスタ
        品番                      As String
        取引先                    As String
        受入                      As String
        品目                      As Byte
        工順                      As Byte
        機械                      As String
        外注加工                  As Byte
        外注加工先CD              As String
        機械NO                    As String
        ロボット区分              As Byte
        工程                      As String
        基準工数名1               As String
        基準工数1                 As Currency
        基準工数名2               As String
        基準工数2                 As Currency
        工程数                    As Currency
        棚卸プレシート発行        As Byte
        置場                      As String
        出荷管理                  As Byte
        セット品番                As Byte
        サイクル                  As String
        実工数                    As Currency
        係数1                     As Currency
        係数2                     As Currency
        LT                        As Byte
        在庫管理区分              As Byte
        在庫管理数値              As Currency
        在庫係数                  As Currency
        在庫日数                  As Currency
        在庫発注点                As Long
        最大在庫管理区分          As Byte
        最大在庫管理数値          As Currency
        最大在庫係数              As Currency
        最大在庫日数              As Currency
        最大在庫発注点            As Long
        収容数                    As Long
        日報出力                  As Byte
'        作成日                    As Date
        作成者                    As String
'        作成端末                  As String
'        更新日                    As Date
        更新者                    As String
'        更新端末                  As String
        件数          As Long
    
        機械名                    As String
        外注加工名                As String
        機械NO名                  As String
        工程名                    As String
    End Type

    Public HKT As T_品番工順マスタ

    Public HKT_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHKT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHKT
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With HKT
            Select Case adoCol.Name
                Case "品番":                      If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "取引先":                    If Not IsNull(adoCol.Value) Then .取引先 = adoCol.Value
                Case "受入":                      If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "品目":                      If Not IsNull(adoCol.Value) Then .品目 = adoCol.Value
                Case "工順":                      If Not IsNull(adoCol.Value) Then .工順 = adoCol.Value
                Case "機械":                      If Not IsNull(adoCol.Value) Then .機械 = adoCol.Value
                Case "外注加工":                  If Not IsNull(adoCol.Value) Then .外注加工 = adoCol.Value
                Case "外注加工先CD":              If Not IsNull(adoCol.Value) Then .外注加工先CD = adoCol.Value
                Case "機械NO":                    If Not IsNull(adoCol.Value) Then .機械NO = adoCol.Value
                Case "ロボット区分":              If Not IsNull(adoCol.Value) Then .ロボット区分 = adoCol.Value
                Case "工程":                      If Not IsNull(adoCol.Value) Then .工程 = adoCol.Value
                Case "基準工数名1":               If Not IsNull(adoCol.Value) Then .基準工数名1 = adoCol.Value
                Case "基準工数1":                 If Not IsNull(adoCol.Value) Then .基準工数1 = adoCol.Value
                Case "基準工数名2":               If Not IsNull(adoCol.Value) Then .基準工数名2 = adoCol.Value
                Case "基準工数2":                 If Not IsNull(adoCol.Value) Then .基準工数2 = adoCol.Value
                Case "工程数":                    If Not IsNull(adoCol.Value) Then .工程数 = adoCol.Value
                Case "棚卸プレシート発行":        If Not IsNull(adoCol.Value) Then .棚卸プレシート発行 = adoCol.Value
                Case "置場":                      If Not IsNull(adoCol.Value) Then .置場 = adoCol.Value
                Case "出荷管理":                  If Not IsNull(adoCol.Value) Then .出荷管理 = adoCol.Value
                Case "セット品番":                If Not IsNull(adoCol.Value) Then .セット品番 = adoCol.Value
                Case "サイクル":                  If Not IsNull(adoCol.Value) Then .サイクル = adoCol.Value
                Case "実工数":                    If Not IsNull(adoCol.Value) Then .実工数 = adoCol.Value
                Case "係数1":                     If Not IsNull(adoCol.Value) Then .係数1 = adoCol.Value
                Case "係数2":                     If Not IsNull(adoCol.Value) Then .係数2 = adoCol.Value
                Case "LT":                        If Not IsNull(adoCol.Value) Then .LT = adoCol.Value
                Case "在庫管理区分":              If Not IsNull(adoCol.Value) Then .在庫管理区分 = adoCol.Value
                Case "在庫管理数値":              If Not IsNull(adoCol.Value) Then .在庫管理数値 = adoCol.Value
                Case "在庫係数":                  If Not IsNull(adoCol.Value) Then .在庫係数 = adoCol.Value
                Case "在庫日数":                  If Not IsNull(adoCol.Value) Then .在庫日数 = adoCol.Value
                Case "在庫発注点":                If Not IsNull(adoCol.Value) Then .在庫発注点 = adoCol.Value
                Case "最大在庫管理区分":          If Not IsNull(adoCol.Value) Then .最大在庫管理区分 = adoCol.Value
                Case "最大在庫管理数値":          If Not IsNull(adoCol.Value) Then .最大在庫管理数値 = adoCol.Value
                Case "最大在庫係数":              If Not IsNull(adoCol.Value) Then .最大在庫係数 = adoCol.Value
                Case "最大在庫日数":              If Not IsNull(adoCol.Value) Then .最大在庫日数 = adoCol.Value
                Case "最大在庫発注点":            If Not IsNull(adoCol.Value) Then .最大在庫発注点 = adoCol.Value
                Case "収容数":                    If Not IsNull(adoCol.Value) Then .収容数 = adoCol.Value
                Case "日報出力":                  If Not IsNull(adoCol.Value) Then .日報出力 = adoCol.Value
'                Case "作成日":                    If Not IsNull(adoCol.Value) Then .作成日                   = adoCol.Value
                Case "作成者":                    If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":                  If Not IsNull(adoCol.Value) Then .作成端末                 = adoCol.Value
'                Case "更新日":                    If Not IsNull(adoCol.Value) Then .更新日                   = adoCol.Value
                Case "更新者":                    If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新端末":                  If Not IsNull(adoCol.Value) Then .更新端末                 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
            
                Case "機械名":                    If Not IsNull(adoCol.Value) Then .機械名 = adoCol.Value
                Case "外注加工名":                If Not IsNull(adoCol.Value) Then .外注加工名 = adoCol.Value
                Case "機械NO名":                  If Not IsNull(adoCol.Value) Then .機械NO名 = adoCol.Value
                Case "工程名":                    If Not IsNull(adoCol.Value) Then .工程名 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function HKTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HKTRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 品番工順マスタ"
        RTxtSql = RTxtSql & " Where 品番 = '" & RTrim(HKT.品番) & "'"
        RTxtSql = RTxtSql & " And   取引先 = '" & RTrim(HKT.取引先) & "'"
        RTxtSql = RTxtSql & " And   受入 = '" & RTrim(HKT.受入) & "'"
        RTxtSql = RTxtSql & " And   品目 = " & HKT.品目 & ""
        RTxtSql = RTxtSql & " And   工順 = " & HKT.工順 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HKTRead = True
           Call ItemsFromHKT(DBNo)
           HKT_RDSTS = True
        Else
           HKT_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call HKTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HKTReadNext(Optional DBNo As Integer = 0)
    If Not HKT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HKT_RDSTS = False
        Else
           Call ItemsFromHKT(DBNo)
           HKT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HKTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HKT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HKT )
'*----------------------------------------------*'
Public Sub ItemsClearHKT()
    With HKT
        .品番 = ""
        .取引先 = ""
        .受入 = ""
        .品目 = 0
        .工順 = 0
        .機械 = 0
        .外注加工 = 0
        .外注加工先CD = ""
        .機械NO = 0
        .ロボット区分 = 0
        .工程 = 0
        .基準工数名1 = ""
        .基準工数1 = 0
        .基準工数名2 = ""
        .基準工数2 = 0
        .工程数 = 0
        .棚卸プレシート発行 = 0
        .置場 = ""
        .出荷管理 = 0
        .セット品番 = 0
        .サイクル = ""
        .実工数 = 0
        .係数1 = 0
        .係数2 = 0
        .LT = 0
        .在庫管理区分 = 0
        .在庫管理数値 = 0
        .在庫係数 = 0
        .在庫日数 = 0
        .在庫発注点 = 0
        .最大在庫管理区分 = 0
        .最大在庫管理数値 = 0
        .最大在庫係数 = 0
        .最大在庫日数 = 0
        .最大在庫発注点 = 0
        .収容数 = 0
        .日報出力 = 0
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
'       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
        .件数 = 0
        .機械名 = ""
        .外注加工名 = ""
        .機械NO名 = ""
        .工程名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function HKTInsert(Optional DBNo As Integer) As Boolean
    HKTInsert = False
    With HKT
        gSL_Select = "Insert Into 品番工順マスタ  ( "
        gSL_Select = gSL_Select & "品番,取引先,受入,"
        gSL_Select = gSL_Select & "品目,工順,機械,"
        gSL_Select = gSL_Select & "外注加工,外注加工先CD,機械NO,"
        gSL_Select = gSL_Select & "ロボット区分,工程,基準工数名1,"
        gSL_Select = gSL_Select & "基準工数1,基準工数名2,基準工数2,工程数,"
        gSL_Select = gSL_Select & "棚卸プレシート発行,置場,出荷管理,"
        gSL_Select = gSL_Select & "セット品番,サイクル,実工数,"
        gSL_Select = gSL_Select & "係数1,係数2,LT,"
        gSL_Select = gSL_Select & "在庫管理区分,在庫管理数値,在庫係数,"
        gSL_Select = gSL_Select & "在庫日数,在庫発注点,最大在庫管理区分,"
        gSL_Select = gSL_Select & "最大在庫管理数値,最大在庫係数,最大在庫日数,"
        gSL_Select = gSL_Select & "最大在庫発注点,収容数,日報出力,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "" & .品目 & ","
        gSL_Select = gSL_Select & "" & .工順 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.機械) & "',"
        gSL_Select = gSL_Select & "" & .外注加工 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.外注加工先CD) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.機械NO) & "',"
        gSL_Select = gSL_Select & "" & .ロボット区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.工程) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.基準工数名1) & "',"
        gSL_Select = gSL_Select & "" & .基準工数1 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.基準工数名2) & "',"
        gSL_Select = gSL_Select & "" & .基準工数2 & ","
        gSL_Select = gSL_Select & "" & .工程数 & ","
        gSL_Select = gSL_Select & "" & .棚卸プレシート発行 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.置場) & "',"
        gSL_Select = gSL_Select & "" & .出荷管理 & ","
        gSL_Select = gSL_Select & "" & .セット品番 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "" & .実工数 & ","
        gSL_Select = gSL_Select & "" & .係数1 & ","
        gSL_Select = gSL_Select & "" & .係数2 & ","
        gSL_Select = gSL_Select & "" & .LT & ","
        gSL_Select = gSL_Select & "" & .在庫管理区分 & ","
        gSL_Select = gSL_Select & "" & .在庫管理数値 & ","
        gSL_Select = gSL_Select & "" & .在庫係数 & ","
        gSL_Select = gSL_Select & "" & .在庫日数 & ","
        gSL_Select = gSL_Select & "" & .在庫発注点 & ","
        gSL_Select = gSL_Select & "" & .最大在庫管理区分 & ","
        gSL_Select = gSL_Select & "" & .最大在庫管理数値 & ","
        gSL_Select = gSL_Select & "" & .最大在庫係数 & ","
        gSL_Select = gSL_Select & "" & .最大在庫日数 & ","
        gSL_Select = gSL_Select & "" & .最大在庫発注点 & ","
        gSL_Select = gSL_Select & "" & .収容数 & ","
        gSL_Select = gSL_Select & "" & .日報出力 & ","
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
    HKTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function HKTUpdate(Optional DBNo As Integer) As Boolean
    HKTUpdate = False
    With HKT
        gSL_Select = "Update 品番工順マスタ  Set "
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "品目 = " & .品目 & ","
        gSL_Select = gSL_Select & "工順 = " & .工順 & ","
        gSL_Select = gSL_Select & "機械 = '" & RTrim(.機械) & "',"
        gSL_Select = gSL_Select & "外注加工 = " & .外注加工 & ","
        gSL_Select = gSL_Select & "外注加工先CD = '" & RTrim(.外注加工先CD) & "',"
        gSL_Select = gSL_Select & "機械NO = '" & RTrim(.機械NO) & "',"
        gSL_Select = gSL_Select & "ロボット区分 = " & .ロボット区分 & ","
        gSL_Select = gSL_Select & "工程 = '" & RTrim(.工程) & "',"
        gSL_Select = gSL_Select & "基準工数名1 = '" & RTrim(.基準工数名1) & "',"
        gSL_Select = gSL_Select & "基準工数1 = " & .基準工数1 & ","
        gSL_Select = gSL_Select & "基準工数名2 = '" & RTrim(.基準工数名2) & "',"
        gSL_Select = gSL_Select & "基準工数2 = " & .基準工数2 & ","
        gSL_Select = gSL_Select & "工程数 = " & .工程数 & ","
        gSL_Select = gSL_Select & "棚卸プレシート発行 = " & .棚卸プレシート発行 & ","
        gSL_Select = gSL_Select & "置場 = '" & RTrim(.置場) & "',"
        gSL_Select = gSL_Select & "出荷管理 = " & .出荷管理 & ","
        gSL_Select = gSL_Select & "セット品番 = " & .セット品番 & ","
        gSL_Select = gSL_Select & "サイクル = '" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "実工数 = " & .実工数 & ","
        gSL_Select = gSL_Select & "係数1 = " & .係数1 & ","
        gSL_Select = gSL_Select & "係数2 = " & .係数2 & ","
        gSL_Select = gSL_Select & "LT = " & .LT & ","
        gSL_Select = gSL_Select & "在庫管理区分 = " & .在庫管理区分 & ","
        gSL_Select = gSL_Select & "在庫管理数値 = " & .在庫管理数値 & ","
        gSL_Select = gSL_Select & "在庫係数 = " & .在庫係数 & ","
        gSL_Select = gSL_Select & "在庫日数 = " & .在庫日数 & ","
        gSL_Select = gSL_Select & "在庫発注点 = " & .在庫発注点 & ","
        gSL_Select = gSL_Select & "最大在庫管理区分 = " & .最大在庫管理区分 & ","
        gSL_Select = gSL_Select & "最大在庫管理数値 = " & .最大在庫管理数値 & ","
        gSL_Select = gSL_Select & "最大在庫係数 = " & .最大在庫係数 & ","
        gSL_Select = gSL_Select & "最大在庫日数 = " & .最大在庫日数 & ","
        gSL_Select = gSL_Select & "最大在庫発注点 = " & .最大在庫発注点 & ","
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "日報出力 = " & .日報出力 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 品目 = " & .品目 & ""
        gSL_Select = gSL_Select & "   And 工順 = " & .工順 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HKTUpdate = True
End Function

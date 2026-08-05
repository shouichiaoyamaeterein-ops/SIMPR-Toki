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
'*   作成日  :  2009/07/24  By Cis(FMVW5220)
'*   作成日  :  2009/12/05  By Cis(FMVW5220)
'*   作成日  :  2009/12/10  By Cis(FMVW5220)
'*   作成日  :  2009/12/18  By Cis(FMVW5220)
'*   作成日  :  2010/01/21  By Cis(HP-DX7500)
'*   作成日  :  2010/01/25  By Cis(HP-DX7500)
'*   作成日  :  2010/01/26  By Cis(FMVW5220)
'*   作成日  :  2010/01/27  By Cis(FMVW5220)
'*   変更日  :  2010/09/16  By CIS  (使用数1～5 追加)
'*   変更日  :  2011/01/20  By CIS  (品名,車型 追加)
'*---------------------------------------------------------------------------------
    Type T_品番工順マスタ
        品番                  As String
        取引先                As String
        受入                  As String
        品目                  As Byte
        工順                  As Byte
        機械                  As String
        外注加工              As Byte
        外注加工先CD          As String
        機械NO                As String
        ロボット区分          As Byte
        工程                  As String
        工数値選択            As Byte
        基準工数名1           As String
        基準工数1             As Currency
        基準段取1             As Currency
        基準工数名2           As String
        基準工数2             As Currency
        基準段取2             As Currency
        工程数                As Currency
        棚卸プレシート発行    As Byte
        置場                  As String
        出荷管理              As Byte
        セット品番            As Byte
        サイクル              As String
        実工数                As Currency
        係数1                 As Currency
        係数2                 As Currency
        LT                    As Byte
        在庫管理区分          As Byte
        在庫管理数値          As Currency
        在庫係数              As Currency
        在庫日数              As Currency
        在庫発注点            As Long
        最大在庫管理区分      As Byte
        最大在庫管理数値      As Currency
        最大在庫係数          As Currency
        最大在庫日数          As Currency
        最大在庫発注点        As Long
        収容数                As Long
        収容器                As String
        日報出力              As Byte
        生産種別              As Byte
        かんばん区分          As Byte
        かんばん区分サイクル  As String
        かんばん区分ロット    As Long
        生産区分ロット        As Long
        備考                  As String
        優先順位              As String
        加工数                As Long
'        作成日                As Date
        作成者                As String
'        作成端末              As String
'        更新日                As Date
        更新者                As String
'        更新端末              As String
        件数          As Long

        使用数1               As Currency
        使用数2               As Currency
        使用数3               As Currency
        使用数4               As Currency
        使用数5               As Currency
'   項目追加(自動)
        機械名         As String
        外注加工名     As String
        機械NO名       As String
        工程名         As String
        納期           As String
        部門           As String
        区分           As String
        数量           As Currency
        基準時間       As Currency
        目標時間       As Currency
        表示工程数     As Currency
        取引先名       As String
        得意先名       As String
        パンチ数       As Currency
        材料名         As String
        内示当月       As Currency
        内示翌月       As Currency
        内示翌々月     As Currency
        得意先         As String
        表示品番       As String
        背番号         As String
        材料在庫量     As Currency
        部品区分       As Byte
        材料管理番号   As Long
        材質           As String
        板厚           As Currency
        幅             As Currency
        長さ           As Currency
        表示寸法       As String
    
        品名           As String
        車型           As String
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
                Case "品番":                  If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "取引先":                If Not IsNull(adoCol.Value) Then .取引先 = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "品目":                  If Not IsNull(adoCol.Value) Then .品目 = adoCol.Value
                Case "工順":                  If Not IsNull(adoCol.Value) Then .工順 = adoCol.Value
                Case "機械":                  If Not IsNull(adoCol.Value) Then .機械 = adoCol.Value
                Case "外注加工":              If Not IsNull(adoCol.Value) Then .外注加工 = adoCol.Value
                Case "外注加工先CD":          If Not IsNull(adoCol.Value) Then .外注加工先CD = adoCol.Value
                Case "機械NO":                If Not IsNull(adoCol.Value) Then .機械NO = adoCol.Value
                Case "ロボット区分":          If Not IsNull(adoCol.Value) Then .ロボット区分 = adoCol.Value
                Case "工程":                  If Not IsNull(adoCol.Value) Then .工程 = adoCol.Value
                Case "工数値選択":            If Not IsNull(adoCol.Value) Then .工数値選択 = adoCol.Value
                Case "基準工数名1":           If Not IsNull(adoCol.Value) Then .基準工数名1 = adoCol.Value
                Case "基準工数1":             If Not IsNull(adoCol.Value) Then .基準工数1 = adoCol.Value
                Case "基準段取1":             If Not IsNull(adoCol.Value) Then .基準段取1 = adoCol.Value
                Case "基準工数名2":           If Not IsNull(adoCol.Value) Then .基準工数名2 = adoCol.Value
                Case "基準工数2":             If Not IsNull(adoCol.Value) Then .基準工数2 = adoCol.Value
                Case "基準段取2":             If Not IsNull(adoCol.Value) Then .基準段取2 = adoCol.Value
                Case "工程数":                If Not IsNull(adoCol.Value) Then .工程数 = adoCol.Value
                Case "棚卸プレシート発行":    If Not IsNull(adoCol.Value) Then .棚卸プレシート発行 = adoCol.Value
                Case "置場":                  If Not IsNull(adoCol.Value) Then .置場 = adoCol.Value
                Case "出荷管理":              If Not IsNull(adoCol.Value) Then .出荷管理 = adoCol.Value
                Case "セット品番":            If Not IsNull(adoCol.Value) Then .セット品番 = adoCol.Value
                Case "サイクル":              If Not IsNull(adoCol.Value) Then .サイクル = adoCol.Value
                Case "実工数":                If Not IsNull(adoCol.Value) Then .実工数 = adoCol.Value
                Case "係数1":                 If Not IsNull(adoCol.Value) Then .係数1 = adoCol.Value
                Case "係数2":                 If Not IsNull(adoCol.Value) Then .係数2 = adoCol.Value
                Case "LT":                    If Not IsNull(adoCol.Value) Then .LT = adoCol.Value
                Case "在庫管理区分":          If Not IsNull(adoCol.Value) Then .在庫管理区分 = adoCol.Value
                Case "在庫管理数値":          If Not IsNull(adoCol.Value) Then .在庫管理数値 = adoCol.Value
                Case "在庫係数":              If Not IsNull(adoCol.Value) Then .在庫係数 = adoCol.Value
                Case "在庫日数":              If Not IsNull(adoCol.Value) Then .在庫日数 = adoCol.Value
                Case "在庫発注点":            If Not IsNull(adoCol.Value) Then .在庫発注点 = adoCol.Value
                Case "最大在庫管理区分":      If Not IsNull(adoCol.Value) Then .最大在庫管理区分 = adoCol.Value
                Case "最大在庫管理数値":      If Not IsNull(adoCol.Value) Then .最大在庫管理数値 = adoCol.Value
                Case "最大在庫係数":          If Not IsNull(adoCol.Value) Then .最大在庫係数 = adoCol.Value
                Case "最大在庫日数":          If Not IsNull(adoCol.Value) Then .最大在庫日数 = adoCol.Value
                Case "最大在庫発注点":        If Not IsNull(adoCol.Value) Then .最大在庫発注点 = adoCol.Value
                Case "収容数":                If Not IsNull(adoCol.Value) Then .収容数 = adoCol.Value
                Case "収容器":                If Not IsNull(adoCol.Value) Then .収容器 = adoCol.Value
                Case "日報出力":              If Not IsNull(adoCol.Value) Then .日報出力 = adoCol.Value
                Case "生産種別":              If Not IsNull(adoCol.Value) Then .生産種別 = adoCol.Value
                Case "かんばん区分":          If Not IsNull(adoCol.Value) Then .かんばん区分 = adoCol.Value
                Case "かんばん区分サイクル":  If Not IsNull(adoCol.Value) Then .かんばん区分サイクル = adoCol.Value
                Case "かんばん区分ロット":    If Not IsNull(adoCol.Value) Then .かんばん区分ロット = adoCol.Value
                Case "生産区分ロット":        If Not IsNull(adoCol.Value) Then .生産区分ロット = adoCol.Value
                Case "備考":                  If Not IsNull(adoCol.Value) Then .備考 = adoCol.Value
                Case "優先順位":              If Not IsNull(adoCol.Value) Then .優先順位 = adoCol.Value
                Case "加工数":                If Not IsNull(adoCol.Value) Then .加工数 = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
                Case "作成者":                If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
'                Case "更新日":                If Not IsNull(adoCol.Value) Then .更新日               = adoCol.Value
                Case "更新者":                If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新端末":              If Not IsNull(adoCol.Value) Then .更新端末             = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value

                Case "使用数1":               If Not IsNull(adoCol.Value) Then .使用数1 = adoCol.Value
                Case "使用数2":               If Not IsNull(adoCol.Value) Then .使用数2 = adoCol.Value
                Case "使用数3":               If Not IsNull(adoCol.Value) Then .使用数3 = adoCol.Value
                Case "使用数4":               If Not IsNull(adoCol.Value) Then .使用数4 = adoCol.Value
                Case "使用数5":               If Not IsNull(adoCol.Value) Then .使用数5 = adoCol.Value
'   項目追加(自動)
                Case "機械名":  If Not IsNull(adoCol.Value) Then .機械名 = adoCol.Value
                Case "外注加工名":      If Not IsNull(adoCol.Value) Then .外注加工名 = adoCol.Value
                Case "機械NO名":        If Not IsNull(adoCol.Value) Then .機械NO名 = adoCol.Value
                Case "工程名":  If Not IsNull(adoCol.Value) Then .工程名 = adoCol.Value
                Case "納期":    If Not IsNull(adoCol.Value) Then .納期 = adoCol.Value
                Case "部門":    If Not IsNull(adoCol.Value) Then .部門 = adoCol.Value
                Case "区分":    If Not IsNull(adoCol.Value) Then .区分 = adoCol.Value
                Case "数量":    If Not IsNull(adoCol.Value) Then .数量 = adoCol.Value
                Case "基準時間":        If Not IsNull(adoCol.Value) Then .基準時間 = adoCol.Value
                Case "目標時間":        If Not IsNull(adoCol.Value) Then .目標時間 = adoCol.Value
                Case "表示工程数":      If Not IsNull(adoCol.Value) Then .表示工程数 = adoCol.Value
                Case "取引先名":        If Not IsNull(adoCol.Value) Then .取引先名 = adoCol.Value
                Case "得意先名":        If Not IsNull(adoCol.Value) Then .得意先名 = adoCol.Value
                Case "パンチ数":        If Not IsNull(adoCol.Value) Then .パンチ数 = adoCol.Value
                Case "材料名":  If Not IsNull(adoCol.Value) Then .材料名 = adoCol.Value
                Case "内示当月":        If Not IsNull(adoCol.Value) Then .内示当月 = adoCol.Value
                Case "内示翌月":        If Not IsNull(adoCol.Value) Then .内示翌月 = adoCol.Value
                Case "内示翌々月":      If Not IsNull(adoCol.Value) Then .内示翌々月 = adoCol.Value
                Case "得意先":  If Not IsNull(adoCol.Value) Then .得意先 = adoCol.Value
                Case "表示品番":        If Not IsNull(adoCol.Value) Then .表示品番 = adoCol.Value
                Case "背番号":  If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "材料在庫量":      If Not IsNull(adoCol.Value) Then .材料在庫量 = adoCol.Value
                Case "部品区分":        If Not IsNull(adoCol.Value) Then .部品区分 = adoCol.Value
                Case "材料管理番号":    If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "材質":    If Not IsNull(adoCol.Value) Then .材質 = adoCol.Value
                Case "板厚":    If Not IsNull(adoCol.Value) Then .板厚 = adoCol.Value
                Case "幅":      If Not IsNull(adoCol.Value) Then .幅 = adoCol.Value
                Case "長さ":    If Not IsNull(adoCol.Value) Then .長さ = adoCol.Value
                Case "表示寸法":        If Not IsNull(adoCol.Value) Then .表示寸法 = adoCol.Value
            
                Case "品名":        If Not IsNull(adoCol.Value) Then .品名 = adoCol.Value
                Case "車型":        If Not IsNull(adoCol.Value) Then .車型 = adoCol.Value
                Case "表示品番":        If Not IsNull(adoCol.Value) Then .表示品番 = adoCol.Value
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
        .機械 = ""
        .外注加工 = 0
        .外注加工先CD = ""
        .機械NO = ""
        .ロボット区分 = 0
        .工程 = ""
        .工数値選択 = 0
        .基準工数名1 = ""
        .基準工数1 = 0
        .基準段取1 = 0
        .基準工数名2 = ""
        .基準工数2 = 0
        .基準段取2 = 0
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
        .収容器 = ""
        .日報出力 = 0
        .生産種別 = 0
        .かんばん区分 = 0
        .かんばん区分サイクル = ""
        .かんばん区分ロット = 0
        .生産区分ロット = 0
        .備考 = ""
        .優先順位 = ""
        .加工数 = 0
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
'       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
        .件数 = 0

        .使用数1 = 0
        .使用数2 = 0
        .使用数3 = 0
        .使用数4 = 0
        .使用数5 = 0
'   項目追加(自動)
        .機械名 = ""
        .外注加工名 = ""
        .機械NO名 = ""
        .工程名 = ""
        .納期 = ""
        .部門 = ""
        .区分 = ""
        .数量 = 0
        .基準時間 = 0
        .目標時間 = 0
        .表示工程数 = 0
        .取引先名 = ""
        .得意先名 = ""
        .パンチ数 = 0
        .材料名 = ""
        .内示当月 = 0
        .内示翌月 = 0
        .内示翌々月 = 0
        .得意先 = ""
        .表示品番 = ""
        .背番号 = ""
        .材料在庫量 = 0
        .部品区分 = 0
        .材料管理番号 = 0
        .材質 = ""
        .板厚 = 0
        .幅 = 0
        .長さ = 0
        .表示寸法 = ""
    
        .品名 = ""
        .車型 = ""
        .表示品番 = ""
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
        gSL_Select = gSL_Select & "ロボット区分,工程,工数値選択,"
        gSL_Select = gSL_Select & "基準工数名1,基準工数1,基準段取1,"
        gSL_Select = gSL_Select & "基準工数名2,基準工数2,基準段取2,"
        gSL_Select = gSL_Select & "工程数,棚卸プレシート発行,置場,"
        gSL_Select = gSL_Select & "出荷管理,セット品番,サイクル,"
        gSL_Select = gSL_Select & "実工数,係数1,係数2,"
        gSL_Select = gSL_Select & "LT,在庫管理区分,在庫管理数値,"
        gSL_Select = gSL_Select & "在庫係数,在庫日数,在庫発注点,"
        gSL_Select = gSL_Select & "最大在庫管理区分,最大在庫管理数値,最大在庫係数,"
        gSL_Select = gSL_Select & "最大在庫日数,最大在庫発注点,収容数,"
        gSL_Select = gSL_Select & "収容器,日報出力,生産種別,"
        gSL_Select = gSL_Select & "かんばん区分,かんばん区分サイクル,かんばん区分ロット,"
        gSL_Select = gSL_Select & "生産区分ロット,備考,優先順位,"
        gSL_Select = gSL_Select & "加工数,使用数1,使用数2,使用数3,使用数4,使用数5,"
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
        gSL_Select = gSL_Select & "" & .工数値選択 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.基準工数名1) & "',"
        gSL_Select = gSL_Select & "" & .基準工数1 & ","
        gSL_Select = gSL_Select & "" & .基準段取1 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.基準工数名2) & "',"
        gSL_Select = gSL_Select & "" & .基準工数2 & ","
        gSL_Select = gSL_Select & "" & .基準段取2 & ","
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
        gSL_Select = gSL_Select & "'" & RTrim(.収容器) & "',"
        gSL_Select = gSL_Select & "" & .日報出力 & ","
        gSL_Select = gSL_Select & "" & .生産種別 & ","
        gSL_Select = gSL_Select & "" & .かんばん区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.かんばん区分サイクル) & "',"
        gSL_Select = gSL_Select & "" & .かんばん区分ロット & ","
        gSL_Select = gSL_Select & "" & .生産区分ロット & ","
        gSL_Select = gSL_Select & "'" & RTrim(.備考) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.優先順位) & "',"
        gSL_Select = gSL_Select & "" & .加工数 & ","
        
        gSL_Select = gSL_Select & "" & .使用数1 & ","
        gSL_Select = gSL_Select & "" & .使用数2 & ","
        gSL_Select = gSL_Select & "" & .使用数3 & ","
        gSL_Select = gSL_Select & "" & .使用数4 & ","
        gSL_Select = gSL_Select & "" & .使用数5 & ","
        
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
        gSL_Select = gSL_Select & "工数値選択 = " & .工数値選択 & ","
        gSL_Select = gSL_Select & "基準工数名1 = '" & RTrim(.基準工数名1) & "',"
        gSL_Select = gSL_Select & "基準工数1 = " & .基準工数1 & ","
        gSL_Select = gSL_Select & "基準段取1 = " & .基準段取1 & ","
        gSL_Select = gSL_Select & "基準工数名2 = '" & RTrim(.基準工数名2) & "',"
        gSL_Select = gSL_Select & "基準工数2 = " & .基準工数2 & ","
        gSL_Select = gSL_Select & "基準段取2 = " & .基準段取2 & ","
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
        gSL_Select = gSL_Select & "収容器 = '" & RTrim(.収容器) & "',"
        gSL_Select = gSL_Select & "日報出力 = " & .日報出力 & ","
        gSL_Select = gSL_Select & "生産種別 = " & .生産種別 & ","
        gSL_Select = gSL_Select & "かんばん区分 = " & .かんばん区分 & ","
        gSL_Select = gSL_Select & "かんばん区分サイクル = '" & RTrim(.かんばん区分サイクル) & "',"
        gSL_Select = gSL_Select & "かんばん区分ロット = " & .かんばん区分ロット & ","
        gSL_Select = gSL_Select & "生産区分ロット = " & .生産区分ロット & ","
        gSL_Select = gSL_Select & "備考 = '" & RTrim(.備考) & "',"
        gSL_Select = gSL_Select & "優先順位 = '" & RTrim(.優先順位) & "',"
        gSL_Select = gSL_Select & "加工数 = " & .加工数 & ","
        
        gSL_Select = gSL_Select & "使用数1 = " & .使用数1 & ","
        gSL_Select = gSL_Select & "使用数2 = " & .使用数2 & ","
        gSL_Select = gSL_Select & "使用数3 = " & .使用数3 & ","
        gSL_Select = gSL_Select & "使用数4 = " & .使用数4 & ","
        gSL_Select = gSL_Select & "使用数5 = " & .使用数5 & ","
        
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

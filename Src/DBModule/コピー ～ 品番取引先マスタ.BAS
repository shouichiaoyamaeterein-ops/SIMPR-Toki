Attribute VB_Name = "品番取引先マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   品番取引先マスタ
'*
'*   モジュールID  : 品番取引先マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*                                      /項目追加含む/
'*   作成日  :  2004/12/15  By Cis
'*   作成日  :  2007/12/10  By Cis
'*   作成日  :  2007/12/14  By Cis      代行区分・代行契約先追加
'*   作成日  :  2008/02/05  By Cis
'*   作成日  :  2008/02/06  By Cis
'*   作成日  :  2008/04/04  By Cis
'*   作成日  :  2009/01/19  By Cis(FMV-E630)
'*   作成日  :  2009/07/06  By Cis(FMV-E630)
'*   作成日  :  2009/07/27  By Cis(FMVW5220)
'*   作成日  :  2009/09/21  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_品番取引先マスタ
        品番              As String
        取引先            As String
        受入              As String
        アドレス          As String
        品目              As Byte
        背番号            As String
        SK区分            As Byte
        か係数            As Currency
        手配区分          As String
        納入先            As String
        納入先受入        As String
        契約先            As String
        サイクル          As String
        収容数1           As Long
        収容器1           As String
        変更日            As String
        変更便            As String
        収容数2           As Long
        収容器2           As String
        置場              As String
        最終発行連番      As Integer
        伝票種類          As String
        納品書形態        As String
        受領書形態        As String
        受領書発行        As Byte
        明細書発行        As String
        平準化区分        As Byte
        平準化係数        As Byte
        係数1             As Currency
        係数2             As Currency
        係数3             As Currency
        LT                As Byte
        ロット1           As Long
        ロット2           As Long
        Week1             As Byte
        Week2             As Byte
        Week3             As Byte
        Week4             As Byte
        Week5             As Byte
        Day1              As Byte
        Day2              As Byte
        Day3              As Byte
        Day4              As Byte
        Day5              As Byte
        内示発注          As Byte
        発注比率          As Byte
        有効期限開始      As String
        有効期限終了      As String
        基準工数          As Currency
        PNO               As Byte
        備考              As String
        色パターン        As Byte
        納品書発行区分    As Byte
        支給書発行区分    As Byte
        受領書発行区分    As Byte
        代行区分          As Byte
        代行契約先        As String
        自社工区          As String
        号補区分          As String
        前工程区分        As Byte
        直送伝票発行区分  As Byte
        出荷場            As String
        部品区分          As Byte
        基準単価          As Currency
        指示書発行区分    As Byte
        出荷管理部署      As String
'        作成日            As Date
        作成者            As String
'        作成端末          As String
'        更新日            As Date
        更新者            As String
'        更新端末          As String
        件数          As Long
'   項目追加(自動)
        品名           As String
        車型           As String
        社内背番       As String
        表示品番       As String
        手配区分名     As String
        取引先名       As String
        納入先名       As String
        在庫管理区分   As Byte
        在庫管理数値   As Currency
        在庫管理名     As String
        機械           As String
        曜日発注       As Byte
        自社工区名     As String
        出荷管理部署名 As String
    End Type

    Public HTM As T_品番取引先マスタ

    Public HTM_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHTM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHTM
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With HTM
            Select Case adoCol.Name
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "取引先":            If Not IsNull(adoCol.Value) Then .取引先 = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "アドレス":          If Not IsNull(adoCol.Value) Then .アドレス = adoCol.Value
                Case "品目":              If Not IsNull(adoCol.Value) Then .品目 = adoCol.Value
                Case "背番号":            If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "SK区分":            If Not IsNull(adoCol.Value) Then .SK区分 = adoCol.Value
                Case "か係数":            If Not IsNull(adoCol.Value) Then .か係数 = adoCol.Value
                Case "手配区分":          If Not IsNull(adoCol.Value) Then .手配区分 = adoCol.Value
                Case "納入先":            If Not IsNull(adoCol.Value) Then .納入先 = adoCol.Value
                Case "納入先受入":        If Not IsNull(adoCol.Value) Then .納入先受入 = adoCol.Value
                Case "契約先":            If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
                Case "サイクル":          If Not IsNull(adoCol.Value) Then .サイクル = adoCol.Value
                Case "収容数1":           If Not IsNull(adoCol.Value) Then .収容数1 = adoCol.Value
                Case "収容器1":           If Not IsNull(adoCol.Value) Then .収容器1 = adoCol.Value
                Case "変更日":            If Not IsNull(adoCol.Value) Then .変更日 = adoCol.Value
                Case "変更便":            If Not IsNull(adoCol.Value) Then .変更便 = adoCol.Value
                Case "収容数2":           If Not IsNull(adoCol.Value) Then .収容数2 = adoCol.Value
                Case "収容器2":           If Not IsNull(adoCol.Value) Then .収容器2 = adoCol.Value
                Case "置場":              If Not IsNull(adoCol.Value) Then .置場 = adoCol.Value
                Case "最終発行連番":      If Not IsNull(adoCol.Value) Then .最終発行連番 = adoCol.Value
                Case "伝票種類":          If Not IsNull(adoCol.Value) Then .伝票種類 = adoCol.Value
                Case "納品書形態":        If Not IsNull(adoCol.Value) Then .納品書形態 = adoCol.Value
                Case "受領書形態":        If Not IsNull(adoCol.Value) Then .受領書形態 = adoCol.Value
                Case "受領書発行":        If Not IsNull(adoCol.Value) Then .受領書発行 = adoCol.Value
                Case "明細書発行":        If Not IsNull(adoCol.Value) Then .明細書発行 = adoCol.Value
                Case "平準化区分":        If Not IsNull(adoCol.Value) Then .平準化区分 = adoCol.Value
                Case "平準化係数":        If Not IsNull(adoCol.Value) Then .平準化係数 = adoCol.Value
                Case "係数1":             If Not IsNull(adoCol.Value) Then .係数1 = adoCol.Value
                Case "係数2":             If Not IsNull(adoCol.Value) Then .係数2 = adoCol.Value
                Case "係数3":             If Not IsNull(adoCol.Value) Then .係数3 = adoCol.Value
                Case "LT":                If Not IsNull(adoCol.Value) Then .LT = adoCol.Value
                Case "ロット1":           If Not IsNull(adoCol.Value) Then .ロット1 = adoCol.Value
                Case "ロット2":           If Not IsNull(adoCol.Value) Then .ロット2 = adoCol.Value
                Case "Week1":             If Not IsNull(adoCol.Value) Then .Week1 = adoCol.Value
                Case "Week2":             If Not IsNull(adoCol.Value) Then .Week2 = adoCol.Value
                Case "Week3":             If Not IsNull(adoCol.Value) Then .Week3 = adoCol.Value
                Case "Week4":             If Not IsNull(adoCol.Value) Then .Week4 = adoCol.Value
                Case "Week5":             If Not IsNull(adoCol.Value) Then .Week5 = adoCol.Value
                Case "Day1":              If Not IsNull(adoCol.Value) Then .Day1 = adoCol.Value
                Case "Day2":              If Not IsNull(adoCol.Value) Then .Day2 = adoCol.Value
                Case "Day3":              If Not IsNull(adoCol.Value) Then .Day3 = adoCol.Value
                Case "Day4":              If Not IsNull(adoCol.Value) Then .Day4 = adoCol.Value
                Case "Day5":              If Not IsNull(adoCol.Value) Then .Day5 = adoCol.Value
                Case "内示発注":          If Not IsNull(adoCol.Value) Then .内示発注 = adoCol.Value
                Case "発注比率":          If Not IsNull(adoCol.Value) Then .発注比率 = adoCol.Value
                Case "有効期限開始":      If Not IsNull(adoCol.Value) Then .有効期限開始 = adoCol.Value
                Case "有効期限終了":      If Not IsNull(adoCol.Value) Then .有効期限終了 = adoCol.Value
                Case "基準工数":          If Not IsNull(adoCol.Value) Then .基準工数 = adoCol.Value
                Case "PNO":               If Not IsNull(adoCol.Value) Then .PNO = adoCol.Value
                Case "備考":              If Not IsNull(adoCol.Value) Then .備考 = adoCol.Value
                Case "色パターン":        If Not IsNull(adoCol.Value) Then .色パターン = adoCol.Value
                Case "納品書発行区分":    If Not IsNull(adoCol.Value) Then .納品書発行区分 = adoCol.Value
                Case "支給書発行区分":    If Not IsNull(adoCol.Value) Then .支給書発行区分 = adoCol.Value
                Case "受領書発行区分":    If Not IsNull(adoCol.Value) Then .受領書発行区分 = adoCol.Value
                Case "代行区分":          If Not IsNull(adoCol.Value) Then .代行区分 = adoCol.Value
                Case "代行契約先":        If Not IsNull(adoCol.Value) Then .代行契約先 = adoCol.Value
                Case "自社工区":          If Not IsNull(adoCol.Value) Then .自社工区 = adoCol.Value
                Case "号補区分":          If Not IsNull(adoCol.Value) Then .号補区分 = adoCol.Value
                Case "前工程区分":        If Not IsNull(adoCol.Value) Then .前工程区分 = adoCol.Value
                Case "直送伝票発行区分":  If Not IsNull(adoCol.Value) Then .直送伝票発行区分 = adoCol.Value
                Case "出荷場":            If Not IsNull(adoCol.Value) Then .出荷場 = adoCol.Value
                Case "部品区分":          If Not IsNull(adoCol.Value) Then .部品区分 = adoCol.Value
                Case "基準単価":          If Not IsNull(adoCol.Value) Then .基準単価 = adoCol.Value
                Case "指示書発行区分":    If Not IsNull(adoCol.Value) Then .指示書発行区分 = adoCol.Value
                Case "出荷管理部署":      If Not IsNull(adoCol.Value) Then .出荷管理部署 = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
                Case "更新者":            If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "品名":    If Not IsNull(adoCol.Value) Then .品名 = adoCol.Value
                Case "車型":    If Not IsNull(adoCol.Value) Then .車型 = adoCol.Value
                Case "社内背番":        If Not IsNull(adoCol.Value) Then .社内背番 = adoCol.Value
                Case "表示品番":        If Not IsNull(adoCol.Value) Then .表示品番 = adoCol.Value
                Case "手配区分名":      If Not IsNull(adoCol.Value) Then .手配区分名 = adoCol.Value
                Case "取引先名":        If Not IsNull(adoCol.Value) Then .取引先名 = adoCol.Value
                Case "納入先名":        If Not IsNull(adoCol.Value) Then .納入先名 = adoCol.Value
                Case "在庫管理区分":    If Not IsNull(adoCol.Value) Then .在庫管理区分 = adoCol.Value
                Case "在庫管理数値":    If Not IsNull(adoCol.Value) Then .在庫管理数値 = adoCol.Value
                Case "在庫管理名":      If Not IsNull(adoCol.Value) Then .在庫管理名 = adoCol.Value
                Case "機械":    If Not IsNull(adoCol.Value) Then .機械 = adoCol.Value
                Case "曜日発注":        If Not IsNull(adoCol.Value) Then .曜日発注 = adoCol.Value
                Case "自社工区名":      If Not IsNull(adoCol.Value) Then .自社工区名 = adoCol.Value
                Case "出荷管理部署名":  If Not IsNull(adoCol.Value) Then .出荷管理部署名 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function HTMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HTMRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 品番取引先マスタ"
        RTxtSql = RTxtSql & " Where 品番 = '" & RTrim(HTM.品番) & "'"
        RTxtSql = RTxtSql & " And   取引先 = '" & RTrim(HTM.取引先) & "'"
        RTxtSql = RTxtSql & " And   受入 = '" & RTrim(HTM.受入) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HTMRead = True
           Call ItemsFromHTM(DBNo)
           HTM_RDSTS = True
        Else
           HTM_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call HTMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HTMReadNext(Optional DBNo As Integer = 0)
    If Not HTM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HTM_RDSTS = False
        Else
           Call ItemsFromHTM(DBNo)
           HTM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HTMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HTM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HTM )
'*----------------------------------------------*'
Public Sub ItemsClearHTM()
    With HTM
        .品番 = ""
        .取引先 = ""
        .受入 = ""
        .アドレス = ""
        .品目 = 0
        .背番号 = ""
        .SK区分 = 0
        .か係数 = 0
        .手配区分 = ""
        .納入先 = ""
        .納入先受入 = ""
        .契約先 = ""
        .サイクル = ""
        .収容数1 = 0
        .収容器1 = ""
        .変更日 = ""
        .変更便 = ""
        .収容数2 = 0
        .収容器2 = ""
        .置場 = ""
        .最終発行連番 = 0
        .伝票種類 = ""
        .納品書形態 = ""
        .受領書形態 = ""
        .受領書発行 = 0
        .明細書発行 = ""
        .平準化区分 = 0
        .平準化係数 = 0
        .係数1 = 0
        .係数2 = 0
        .係数3 = 0
        .LT = 0
        .ロット1 = 0
        .ロット2 = 0
        .Week1 = 0
        .Week2 = 0
        .Week3 = 0
        .Week4 = 0
        .Week5 = 0
        .Day1 = 0
        .Day2 = 0
        .Day3 = 0
        .Day4 = 0
        .Day5 = 0
        .内示発注 = 0
        .発注比率 = 0
        .有効期限開始 = ""
        .有効期限終了 = ""
        .基準工数 = 0
        .PNO = 0
        .備考 = ""
        .色パターン = 0
        .納品書発行区分 = 0
        .支給書発行区分 = 0
        .受領書発行区分 = 0
        .代行区分 = 0
        .代行契約先 = ""
        .自社工区 = ""
        .号補区分 = ""
        .前工程区分 = 0
        .直送伝票発行区分 = 0
        .出荷場 = ""
        .部品区分 = 0
        .基準単価 = 0
        .指示書発行区分 = 0
        .出荷管理部署 = ""
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
'       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
        .件数 = 0
'   項目追加(自動)
        .品名 = ""
        .車型 = ""
        .社内背番 = ""
        .表示品番 = ""
        .手配区分名 = ""
        .取引先名 = ""
        .納入先名 = ""
        .在庫管理区分 = 0
        .在庫管理数値 = 0
        .在庫管理名 = ""
        .機械 = ""
        .曜日発注 = 0
        .自社工区名 = ""
        .出荷管理部署名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function HTMInsert(Optional DBNo As Integer) As Boolean
    HTMInsert = False
    With HTM
        gSL_Select = "Insert Into 品番取引先マスタ  ( "
        gSL_Select = gSL_Select & "品番,取引先,受入,"
        gSL_Select = gSL_Select & "アドレス,品目,背番号,"
        gSL_Select = gSL_Select & "SK区分,か係数,手配区分,"
        gSL_Select = gSL_Select & "納入先,納入先受入,契約先,"
        gSL_Select = gSL_Select & "サイクル,収容数1,収容器1,"
        gSL_Select = gSL_Select & "変更日,変更便,収容数2,"
        gSL_Select = gSL_Select & "収容器2,置場,最終発行連番,"
        gSL_Select = gSL_Select & "伝票種類,納品書形態,受領書形態,"
        gSL_Select = gSL_Select & "受領書発行,明細書発行,平準化区分,"
        gSL_Select = gSL_Select & "平準化係数,係数1,係数2,"
        gSL_Select = gSL_Select & "係数3,LT,ロット1,"
        gSL_Select = gSL_Select & "ロット2,Week1,Week2,"
        gSL_Select = gSL_Select & "Week3,Week4,Week5,"
        gSL_Select = gSL_Select & "Day1,Day2,Day3,"
        gSL_Select = gSL_Select & "Day4,Day5,内示発注,"
        gSL_Select = gSL_Select & "発注比率,有効期限開始,有効期限終了,"
        gSL_Select = gSL_Select & "基準工数,PNO,備考,"
        gSL_Select = gSL_Select & "色パターン,納品書発行区分,支給書発行区分,"
        gSL_Select = gSL_Select & "受領書発行区分,代行区分,代行契約先,"
        gSL_Select = gSL_Select & "自社工区,号補区分,前工程区分,"
        gSL_Select = gSL_Select & "直送伝票発行区分,出荷場,部品区分,"
        gSL_Select = gSL_Select & "基準単価,指示書発行区分,出荷管理部署,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.アドレス) & "',"
        gSL_Select = gSL_Select & "" & .品目 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "" & .SK区分 & ","
        gSL_Select = gSL_Select & "" & .か係数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入先受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "" & .収容数1 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.収容器1) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.変更日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.変更便) & "',"
        gSL_Select = gSL_Select & "" & .収容数2 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.収容器2) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.置場) & "',"
        gSL_Select = gSL_Select & "" & .最終発行連番 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.伝票種類) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納品書形態) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受領書形態) & "',"
        gSL_Select = gSL_Select & "" & .受領書発行 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.明細書発行) & "',"
        gSL_Select = gSL_Select & "" & .平準化区分 & ","
        gSL_Select = gSL_Select & "" & .平準化係数 & ","
        gSL_Select = gSL_Select & "" & .係数1 & ","
        gSL_Select = gSL_Select & "" & .係数2 & ","
        gSL_Select = gSL_Select & "" & .係数3 & ","
        gSL_Select = gSL_Select & "" & .LT & ","
        gSL_Select = gSL_Select & "" & .ロット1 & ","
        gSL_Select = gSL_Select & "" & .ロット2 & ","
        gSL_Select = gSL_Select & "" & .Week1 & ","
        gSL_Select = gSL_Select & "" & .Week2 & ","
        gSL_Select = gSL_Select & "" & .Week3 & ","
        gSL_Select = gSL_Select & "" & .Week4 & ","
        gSL_Select = gSL_Select & "" & .Week5 & ","
        gSL_Select = gSL_Select & "" & .Day1 & ","
        gSL_Select = gSL_Select & "" & .Day2 & ","
        gSL_Select = gSL_Select & "" & .Day3 & ","
        gSL_Select = gSL_Select & "" & .Day4 & ","
        gSL_Select = gSL_Select & "" & .Day5 & ","
        gSL_Select = gSL_Select & "" & .内示発注 & ","
        gSL_Select = gSL_Select & "" & .発注比率 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.有効期限開始) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.有効期限終了) & "',"
        gSL_Select = gSL_Select & "" & .基準工数 & ","
        gSL_Select = gSL_Select & "" & .PNO & ","
        gSL_Select = gSL_Select & "'" & RTrim(.備考) & "',"
        gSL_Select = gSL_Select & "" & .色パターン & ","
        gSL_Select = gSL_Select & "" & .納品書発行区分 & ","
        gSL_Select = gSL_Select & "" & .支給書発行区分 & ","
        gSL_Select = gSL_Select & "" & .受領書発行区分 & ","
        gSL_Select = gSL_Select & "" & .代行区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.代行契約先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.自社工区) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.号補区分) & "',"
        gSL_Select = gSL_Select & "" & .前工程区分 & ","
        gSL_Select = gSL_Select & "" & .直送伝票発行区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.出荷場) & "',"
        gSL_Select = gSL_Select & "" & .部品区分 & ","
        gSL_Select = gSL_Select & "" & .基準単価 & ","
        gSL_Select = gSL_Select & "" & .指示書発行区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.出荷管理部署) & "',"
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
    HTMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function HTMUpdate(Optional DBNo As Integer) As Boolean
    HTMUpdate = False
    With HTM
        gSL_Select = "Update 品番取引先マスタ  Set "
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "アドレス = '" & RTrim(.アドレス) & "',"
        gSL_Select = gSL_Select & "品目 = " & .品目 & ","
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "SK区分 = " & .SK区分 & ","
        gSL_Select = gSL_Select & "か係数 = " & .か係数 & ","
        gSL_Select = gSL_Select & "手配区分 = '" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "納入先受入 = '" & RTrim(.納入先受入) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "サイクル = '" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "収容数1 = " & .収容数1 & ","
        gSL_Select = gSL_Select & "収容器1 = '" & RTrim(.収容器1) & "',"
        gSL_Select = gSL_Select & "変更日 = '" & RTrim(.変更日) & "',"
        gSL_Select = gSL_Select & "変更便 = '" & RTrim(.変更便) & "',"
        gSL_Select = gSL_Select & "収容数2 = " & .収容数2 & ","
        gSL_Select = gSL_Select & "収容器2 = '" & RTrim(.収容器2) & "',"
        gSL_Select = gSL_Select & "置場 = '" & RTrim(.置場) & "',"
        gSL_Select = gSL_Select & "最終発行連番 = " & .最終発行連番 & ","
        gSL_Select = gSL_Select & "伝票種類 = '" & RTrim(.伝票種類) & "',"
        gSL_Select = gSL_Select & "納品書形態 = '" & RTrim(.納品書形態) & "',"
        gSL_Select = gSL_Select & "受領書形態 = '" & RTrim(.受領書形態) & "',"
        gSL_Select = gSL_Select & "受領書発行 = " & .受領書発行 & ","
        gSL_Select = gSL_Select & "明細書発行 = '" & RTrim(.明細書発行) & "',"
        gSL_Select = gSL_Select & "平準化区分 = " & .平準化区分 & ","
        gSL_Select = gSL_Select & "平準化係数 = " & .平準化係数 & ","
        gSL_Select = gSL_Select & "係数1 = " & .係数1 & ","
        gSL_Select = gSL_Select & "係数2 = " & .係数2 & ","
        gSL_Select = gSL_Select & "係数3 = " & .係数3 & ","
        gSL_Select = gSL_Select & "LT = " & .LT & ","
        gSL_Select = gSL_Select & "ロット1 = " & .ロット1 & ","
        gSL_Select = gSL_Select & "ロット2 = " & .ロット2 & ","
        gSL_Select = gSL_Select & "Week1 = " & .Week1 & ","
        gSL_Select = gSL_Select & "Week2 = " & .Week2 & ","
        gSL_Select = gSL_Select & "Week3 = " & .Week3 & ","
        gSL_Select = gSL_Select & "Week4 = " & .Week4 & ","
        gSL_Select = gSL_Select & "Week5 = " & .Week5 & ","
        gSL_Select = gSL_Select & "Day1 = " & .Day1 & ","
        gSL_Select = gSL_Select & "Day2 = " & .Day2 & ","
        gSL_Select = gSL_Select & "Day3 = " & .Day3 & ","
        gSL_Select = gSL_Select & "Day4 = " & .Day4 & ","
        gSL_Select = gSL_Select & "Day5 = " & .Day5 & ","
        gSL_Select = gSL_Select & "内示発注 = " & .内示発注 & ","
        gSL_Select = gSL_Select & "発注比率 = " & .発注比率 & ","
        gSL_Select = gSL_Select & "有効期限開始 = '" & RTrim(.有効期限開始) & "',"
        gSL_Select = gSL_Select & "有効期限終了 = '" & RTrim(.有効期限終了) & "',"
        gSL_Select = gSL_Select & "基準工数 = " & .基準工数 & ","
        gSL_Select = gSL_Select & "PNO = " & .PNO & ","
        gSL_Select = gSL_Select & "備考 = '" & RTrim(.備考) & "',"
        gSL_Select = gSL_Select & "色パターン = " & .色パターン & ","
        gSL_Select = gSL_Select & "納品書発行区分 = " & .納品書発行区分 & ","
        gSL_Select = gSL_Select & "支給書発行区分 = " & .支給書発行区分 & ","
        gSL_Select = gSL_Select & "受領書発行区分 = " & .受領書発行区分 & ","
        gSL_Select = gSL_Select & "代行区分 = " & .代行区分 & ","
        gSL_Select = gSL_Select & "代行契約先 = '" & RTrim(.代行契約先) & "',"
        gSL_Select = gSL_Select & "自社工区 = '" & RTrim(.自社工区) & "',"
        gSL_Select = gSL_Select & "号補区分 = '" & RTrim(.号補区分) & "',"
        gSL_Select = gSL_Select & "前工程区分 = " & .前工程区分 & ","
        gSL_Select = gSL_Select & "直送伝票発行区分 = " & .直送伝票発行区分 & ","
        gSL_Select = gSL_Select & "出荷場 = '" & RTrim(.出荷場) & "',"
        gSL_Select = gSL_Select & "部品区分 = " & .部品区分 & ","
        gSL_Select = gSL_Select & "基準単価 = " & .基準単価 & ","
        gSL_Select = gSL_Select & "指示書発行区分 = " & .指示書発行区分 & ","
        gSL_Select = gSL_Select & "出荷管理部署 = '" & RTrim(.出荷管理部署) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HTMUpdate = True
End Function

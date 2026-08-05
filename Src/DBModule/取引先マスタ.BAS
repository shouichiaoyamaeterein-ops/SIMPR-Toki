Attribute VB_Name = "取引先マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   取引先マスタ
'*
'*   モジュールID  : 取引先マスタ.Bas
'*
'*   作成日  :  2004/02/04  By Cis 名称取得モジュール追加
'*   変更日  :  2004/04/05  By Cis 前倒日数 項目追加
'*   作成日  :  2004/05/11  By Cis 名称取得モジュール変更 FAX,Mail 取得追加
'*   作成日  :  2004/05/12  By Cis 開始発注日,MSGTYPE 追加
'*   作成日  :  2007/12/10  By Cis 納品書発行区分,支給書発行区分,受領書発行区分 追加
'*   変更日  :  2008/02/04  By Cis 号補区分,前工程発行区分 追加
'*   変更日  :  2008/06/11  By Cis 計算位1/2 追加
'*   変更日  :  2008/12/26  By Cis 変換元取引先CD,変換元取引先工場,変換元取引先CD追加
'*   変更日  :  2009/07/06  By Cis 新e-かんばん対応 納入先事業体/仕入先事業体/出荷場 追加
'*   変更日  :  2009/07/15  By Cis 土岐津カスタマイズ
'*---------------------------------------------------------------------------------
    Type T_取引先マスタ
        取引先CD          As String
        取引先区分        As Byte
        取引先            As String
        工場CD            As String
        社内区分          As Byte
        契約先区分        As Byte
        正式名称          As String
        略称              As String
        工場名            As String
        索引カナ          As String
        下請法対象        As Byte
        被自社            As String
        被自社工場        As String
        被自社CD          As String
        契約先            As String
        計算方法          As Byte
        計算単位          As Byte
        締日              As Byte
        平準化区分        As Byte
        平準化係数        As Byte
        伝票種類          As String
        納品書形態        As String
        受領書形態        As String
        受領書発行        As Byte
        発注内示          As Byte
        明細書発行        As String
        前倒日数          As Byte
        郵便番号          As String
        代表TEL           As String
        代表FAX           As String
        住所1             As String
        住所2             As String
        部署名1           As String
        部署名2           As String
        役職名            As String
        担当者            As String
        TEL1              As String
        TEL2              As String
        FAX               As String
        EMAIL             As String
        銀行              As String
        銀行名            As String
        支店              As String
        支店名            As String
        口座番号          As String
        預金区分          As Byte
        名義人            As String
        FAX取引先1        As String
        FAX取引先2        As String
        FAX取引先3        As String
        MAIL取引先1       As String
        MAIL取引先2       As String
        MAIL取引先3       As String
        開始発注日        As String
        MSGTYPE           As String
'        作成日            As Date
        作成者            As String
'        作成端末          As String
'        更新日            As Date
        更新者            As String
'        更新端末          As String
        件数          As Long
        
        納品書発行区分     As Byte
        支給書発行区分     As Byte
        受領書発行区分     As Byte
        
        号補区分           As String
        前工程区分         As Byte
        計算位1            As Byte
        計算位2            As Byte
        
        変換元取引先       As String
        変換元取引先工場   As String
        変換元取引先CD     As String
    
        納入先事業体      As String
        納入先事業体工区  As String
        仕入先事業体      As String
        出荷場            As String
    
        '============================= 2009/07/15 Start
        部門CD1           As String
        部門CD2           As String
        正社員人員        As Currency
        正社員稼働時間    As Currency
        正社員原価        As Currency
        パート人員        As Currency
        パート稼働時間    As Currency
        パート原価        As Currency
        その他人員        As Currency
        その他稼働時間    As Currency
        その他原価        As Currency
        日報出力          As Byte
        負荷管理          As Byte
        プレススポット区分  As Byte
        指示書発行区分    As Byte
        '============================= 2009/07/15 End
        品番チェック桁数  As Byte                       ' 2009.11.06 add
            
    End Type

    Public TRM As T_取引先マスタ

    Public TRM_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
    
    Public TRM_Name        As String
    Public TRM_RName       As String
    Public TRM_Shita       As Byte
    Public TRM_Fax         As String
    Public TRM_EMail       As String
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTRM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTRM
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With TRM
            Select Case adoCol.Name
                Case "取引先CD":          If Not IsNull(adoCol.Value) Then .取引先CD = adoCol.Value
                Case "取引先区分":        If Not IsNull(adoCol.Value) Then .取引先区分 = adoCol.Value
                Case "取引先":            If Not IsNull(adoCol.Value) Then .取引先 = adoCol.Value
                Case "工場CD":            If Not IsNull(adoCol.Value) Then .工場CD = adoCol.Value
                Case "社内区分":          If Not IsNull(adoCol.Value) Then .社内区分 = adoCol.Value
                Case "契約先区分":        If Not IsNull(adoCol.Value) Then .契約先区分 = adoCol.Value
                Case "正式名称":          If Not IsNull(adoCol.Value) Then .正式名称 = adoCol.Value
                Case "略称":              If Not IsNull(adoCol.Value) Then .略称 = adoCol.Value
                Case "工場名":            If Not IsNull(adoCol.Value) Then .工場名 = adoCol.Value
                Case "索引カナ":          If Not IsNull(adoCol.Value) Then .索引カナ = adoCol.Value
                Case "下請法対象":        If Not IsNull(adoCol.Value) Then .下請法対象 = adoCol.Value
                Case "被自社":            If Not IsNull(adoCol.Value) Then .被自社 = adoCol.Value
                Case "被自社工場":        If Not IsNull(adoCol.Value) Then .被自社工場 = adoCol.Value
                Case "被自社CD":          If Not IsNull(adoCol.Value) Then .被自社CD = adoCol.Value
                Case "契約先":            If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
                Case "計算方法":          If Not IsNull(adoCol.Value) Then .計算方法 = adoCol.Value
                Case "計算単位":          If Not IsNull(adoCol.Value) Then .計算単位 = adoCol.Value
                Case "締日":              If Not IsNull(adoCol.Value) Then .締日 = adoCol.Value
                Case "平準化区分":        If Not IsNull(adoCol.Value) Then .平準化区分 = adoCol.Value
                Case "平準化係数":        If Not IsNull(adoCol.Value) Then .平準化係数 = adoCol.Value
                Case "伝票種類":          If Not IsNull(adoCol.Value) Then .伝票種類 = adoCol.Value
                Case "納品書形態":        If Not IsNull(adoCol.Value) Then .納品書形態 = adoCol.Value
                Case "受領書形態":        If Not IsNull(adoCol.Value) Then .受領書形態 = adoCol.Value
                Case "受領書発行":        If Not IsNull(adoCol.Value) Then .受領書発行 = adoCol.Value
                Case "発注内示":          If Not IsNull(adoCol.Value) Then .発注内示 = adoCol.Value
                Case "明細書発行":        If Not IsNull(adoCol.Value) Then .明細書発行 = adoCol.Value
                Case "前倒日数":          If Not IsNull(adoCol.Value) Then .前倒日数 = adoCol.Value
                Case "郵便番号":          If Not IsNull(adoCol.Value) Then .郵便番号 = adoCol.Value
                Case "代表TEL":           If Not IsNull(adoCol.Value) Then .代表TEL = adoCol.Value
                Case "代表FAX":           If Not IsNull(adoCol.Value) Then .代表FAX = adoCol.Value
                Case "住所1":             If Not IsNull(adoCol.Value) Then .住所1 = adoCol.Value
                Case "住所2":             If Not IsNull(adoCol.Value) Then .住所2 = adoCol.Value
                Case "部署名1":           If Not IsNull(adoCol.Value) Then .部署名1 = adoCol.Value
                Case "部署名2":           If Not IsNull(adoCol.Value) Then .部署名2 = adoCol.Value
                Case "役職名":            If Not IsNull(adoCol.Value) Then .役職名 = adoCol.Value
                Case "担当者":            If Not IsNull(adoCol.Value) Then .担当者 = adoCol.Value
                Case "TEL1":              If Not IsNull(adoCol.Value) Then .TEL1 = adoCol.Value
                Case "TEL2":              If Not IsNull(adoCol.Value) Then .TEL2 = adoCol.Value
                Case "FAX":               If Not IsNull(adoCol.Value) Then .FAX = adoCol.Value
                Case "EMAIL":             If Not IsNull(adoCol.Value) Then .EMAIL = adoCol.Value
                Case "銀行":              If Not IsNull(adoCol.Value) Then .銀行 = adoCol.Value
                Case "銀行名":            If Not IsNull(adoCol.Value) Then .銀行名 = adoCol.Value
                Case "支店":              If Not IsNull(adoCol.Value) Then .支店 = adoCol.Value
                Case "支店名":            If Not IsNull(adoCol.Value) Then .支店名 = adoCol.Value
                Case "口座番号":          If Not IsNull(adoCol.Value) Then .口座番号 = adoCol.Value
                Case "預金区分":          If Not IsNull(adoCol.Value) Then .預金区分 = adoCol.Value
                Case "名義人":            If Not IsNull(adoCol.Value) Then .名義人 = adoCol.Value
                Case "FAX取引先1":        If Not IsNull(adoCol.Value) Then .FAX取引先1 = adoCol.Value
                Case "FAX取引先2":        If Not IsNull(adoCol.Value) Then .FAX取引先2 = adoCol.Value
                Case "FAX取引先3":        If Not IsNull(adoCol.Value) Then .FAX取引先3 = adoCol.Value
                Case "MAIL取引先1":       If Not IsNull(adoCol.Value) Then .MAIL取引先1 = adoCol.Value
                Case "MAIL取引先2":       If Not IsNull(adoCol.Value) Then .MAIL取引先2 = adoCol.Value
                Case "MAIL取引先3":       If Not IsNull(adoCol.Value) Then .MAIL取引先3 = adoCol.Value
                Case "開始発注日":        If Not IsNull(adoCol.Value) Then .開始発注日 = adoCol.Value
                Case "MSGTYPE":           If Not IsNull(adoCol.Value) Then .MSGTYPE = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
                Case "更新者":            If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":              If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
                
                Case "納品書発行区分":    If Not IsNull(adoCol.Value) Then .納品書発行区分 = adoCol.Value
                Case "支給書発行区分":    If Not IsNull(adoCol.Value) Then .支給書発行区分 = adoCol.Value
                Case "受領書発行区分":    If Not IsNull(adoCol.Value) Then .受領書発行区分 = adoCol.Value
                
                Case "号補区分":          If Not IsNull(adoCol.Value) Then .号補区分 = adoCol.Value
                Case "前工程区分":        If Not IsNull(adoCol.Value) Then .前工程区分 = adoCol.Value
                Case "計算位1":           If Not IsNull(adoCol.Value) Then .計算位1 = adoCol.Value
                Case "計算位2":           If Not IsNull(adoCol.Value) Then .計算位2 = adoCol.Value
                
                Case "変換元取引先":      If Not IsNull(adoCol.Value) Then .変換元取引先 = adoCol.Value
                Case "変換元取引先工場":  If Not IsNull(adoCol.Value) Then .変換元取引先工場 = adoCol.Value
                Case "変換元取引先CD":    If Not IsNull(adoCol.Value) Then .変換元取引先CD = adoCol.Value
            
                Case "納入先事業体":      If Not IsNull(adoCol.Value) Then .納入先事業体 = adoCol.Value
                Case "納入先事業体工区":  If Not IsNull(adoCol.Value) Then .納入先事業体工区 = adoCol.Value
                Case "仕入先事業体":      If Not IsNull(adoCol.Value) Then .仕入先事業体 = adoCol.Value
                Case "出荷場":            If Not IsNull(adoCol.Value) Then .出荷場 = adoCol.Value
            
                Case "部門CD1":            If Not IsNull(adoCol.Value) Then .部門CD1 = adoCol.Value
                Case "部門CD2":            If Not IsNull(adoCol.Value) Then .部門CD2 = adoCol.Value
                Case "正社員人員":         If Not IsNull(adoCol.Value) Then .正社員人員 = adoCol.Value
                Case "正社員稼働時間":     If Not IsNull(adoCol.Value) Then .正社員稼働時間 = adoCol.Value
                Case "正社員原価":         If Not IsNull(adoCol.Value) Then .正社員原価 = adoCol.Value
                Case "パート人員":         If Not IsNull(adoCol.Value) Then .パート人員 = adoCol.Value
                Case "パート稼働時間":     If Not IsNull(adoCol.Value) Then .パート稼働時間 = adoCol.Value
                Case "パート原価":         If Not IsNull(adoCol.Value) Then .パート原価 = adoCol.Value
                Case "その他人員":         If Not IsNull(adoCol.Value) Then .その他人員 = adoCol.Value
                Case "その他稼働時間":     If Not IsNull(adoCol.Value) Then .その他稼働時間 = adoCol.Value
                Case "その他原価":         If Not IsNull(adoCol.Value) Then .その他原価 = adoCol.Value
                Case "日報出力":           If Not IsNull(adoCol.Value) Then .日報出力 = adoCol.Value
                Case "負荷管理":           If Not IsNull(adoCol.Value) Then .負荷管理 = adoCol.Value
                Case "プレススポット区分": If Not IsNull(adoCol.Value) Then .プレススポット区分 = adoCol.Value
                Case "指示書発行区分":     If Not IsNull(adoCol.Value) Then .指示書発行区分 = adoCol.Value
                Case "品番チェック桁数":   If Not IsNull(adoCol.Value) Then .品番チェック桁数 = adoCol.Value    ' 2009.11.06 add
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TRMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TRMRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 取引先マスタ"
        RTxtSql = RTxtSql & " Where 取引先CD = '" & RTrim(TRM.取引先CD) & "'"
        RTxtSql = RTxtSql & " And   取引先区分 = " & TRM.取引先区分 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TRMRead = True
           Call ItemsFromTRM(DBNo)
           TRM_RDSTS = True
        Else
           TRM_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call TRMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TRMReadNext(Optional DBNo As Integer = 0)
    If Not TRM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TRM_RDSTS = False
        Else
           Call ItemsFromTRM(DBNo)
           TRM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TRMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TRM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TRM )
'*----------------------------------------------*'
Public Sub ItemsClearTRM()
    With TRM
        .取引先CD = ""
        .取引先区分 = 0
        .取引先 = ""
        .工場CD = ""
        .社内区分 = 0
        .契約先区分 = 0
        .正式名称 = ""
        .略称 = ""
        .工場名 = ""
        .索引カナ = ""
        .下請法対象 = 0
        .被自社 = ""
        .被自社工場 = ""
        .被自社CD = ""
        .契約先 = ""
        .計算方法 = 0
        .計算単位 = 0
        .締日 = 0
        .平準化区分 = 0
        .平準化係数 = 0
        .伝票種類 = ""
        .納品書形態 = ""
        .受領書形態 = ""
        .受領書発行 = 0
        .発注内示 = 0
        .明細書発行 = ""
        .前倒日数 = 0
        .郵便番号 = ""
        .代表TEL = ""
        .代表FAX = ""
        .住所1 = ""
        .住所2 = ""
        .部署名1 = ""
        .部署名2 = ""
        .役職名 = ""
        .担当者 = ""
        .TEL1 = ""
        .TEL2 = ""
        .FAX = ""
        .EMAIL = ""
        .銀行 = ""
        .銀行名 = ""
        .支店 = ""
        .支店名 = ""
        .口座番号 = ""
        .預金区分 = 0
        .名義人 = ""
        .FAX取引先1 = ""
        .FAX取引先2 = ""
        .FAX取引先3 = ""
        .MAIL取引先1 = ""
        .MAIL取引先2 = ""
        .MAIL取引先3 = ""
        .開始発注日 = ""
        .MSGTYPE = ""
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
'       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
        .納品書発行区分 = 0
        .支給書発行区分 = 0
        .受領書発行区分 = 0
        
        .号補区分 = ""
        .前工程区分 = 0
        .計算位1 = 0
        .計算位2 = 0
        
        .件数 = 0
    
        .変換元取引先 = ""
        .変換元取引先工場 = ""
        .変換元取引先CD = ""
    
        .納入先事業体 = ""
        .納入先事業体工区 = ""
        .仕入先事業体 = ""
        .出荷場 = ""
        
        
        .部門CD1 = ""
        .部門CD2 = ""
        .正社員人員 = 0
        .正社員稼働時間 = 0
        .正社員原価 = 0
        .パート人員 = 0
        .パート稼働時間 = 0
        .パート原価 = 0
        .その他人員 = 0
        .その他稼働時間 = 0
        .その他原価 = 0
        .日報出力 = 0
        .負荷管理 = 0
        .プレススポット区分 = 0
        .指示書発行区分 = 0
        .品番チェック桁数 = 0   ' 2009.11.06 add
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function TRMInsert(Optional DBNo As Integer) As Boolean
    TRMInsert = True
    With TRM
        gSL_Select = "Insert Into 取引先マスタ  ( "
        gSL_Select = gSL_Select & "取引先CD,取引先区分,取引先,"
        gSL_Select = gSL_Select & "工場CD,社内区分,契約先区分,"
        gSL_Select = gSL_Select & "正式名称,略称,工場名,"
        gSL_Select = gSL_Select & "索引カナ,下請法対象,被自社,"
        gSL_Select = gSL_Select & "被自社工場,被自社CD,契約先,"
        gSL_Select = gSL_Select & "計算方法,計算単位,締日,"
        gSL_Select = gSL_Select & "平準化区分,平準化係数,伝票種類,"
        gSL_Select = gSL_Select & "納品書形態,受領書形態,受領書発行,"
        gSL_Select = gSL_Select & "発注内示,明細書発行,前倒日数,"
        gSL_Select = gSL_Select & "郵便番号,代表TEL,代表FAX,"
        gSL_Select = gSL_Select & "住所1,住所2,部署名1,"
        gSL_Select = gSL_Select & "部署名2,役職名,担当者,"
        gSL_Select = gSL_Select & "TEL1,TEL2,FAX,"
        gSL_Select = gSL_Select & "EMAIL,銀行,銀行名,"
        gSL_Select = gSL_Select & "支店,支店名,口座番号,"
        gSL_Select = gSL_Select & "預金区分,名義人,FAX取引先1,"
        gSL_Select = gSL_Select & "FAX取引先2,FAX取引先3,MAIL取引先1,"
        gSL_Select = gSL_Select & "MAIL取引先2,MAIL取引先3,開始発注日,MSGTYPE,"
        gSL_Select = gSL_Select & "納入先事業体,納入先事業体工区,仕入先事業体,出荷場,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末,"
        gSL_Select = gSL_Select & "納品書発行区分,支給書発行区分,受領書発行区分,"
        gSL_Select = gSL_Select & "号補区分, 前工程区分, 計算位1, 計算位2 ,"
        gSL_Select = gSL_Select & "部門CD1 , 部門CD2 , 正社員人員 , 正社員稼働時間 , 正社員原価 ,"
        
        gSL_Select = gSL_Select & "パート人員 , パート稼働時間 , パート原価 , その他人員 , その他稼働時間 , その他原価 , "
        gSL_Select = gSL_Select & "日報出力 , 負荷管理 , プレススポット区分 , 指示書発行区分 "
        gSL_Select = gSL_Select & ", 品番チェック桁数 "     ' 2009.11.06 add
        
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.取引先CD) & "',"
        gSL_Select = gSL_Select & "" & .取引先区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.工場CD) & "',"
        gSL_Select = gSL_Select & "" & .社内区分 & ","
        gSL_Select = gSL_Select & "" & .契約先区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.正式名称) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.略称) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.工場名) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.索引カナ) & "',"
        gSL_Select = gSL_Select & "" & .下請法対象 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.被自社) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.被自社工場) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.被自社CD) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "" & .計算方法 & ","
        gSL_Select = gSL_Select & "" & .計算単位 & ","
        gSL_Select = gSL_Select & "" & .締日 & ","
        gSL_Select = gSL_Select & "" & .平準化区分 & ","
        gSL_Select = gSL_Select & "" & .平準化係数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.伝票種類) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納品書形態) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受領書形態) & "',"
        gSL_Select = gSL_Select & "" & .受領書発行 & ","
        gSL_Select = gSL_Select & "" & .発注内示 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.明細書発行) & "',"
        gSL_Select = gSL_Select & "" & .前倒日数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.郵便番号) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.代表TEL) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.代表FAX) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.住所1) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.住所2) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.部署名1) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.部署名2) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.役職名) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.担当者) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.TEL1) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.TEL2) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.FAX) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.EMAIL) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.銀行) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.銀行名) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.支店) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.支店名) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.口座番号) & "',"
        gSL_Select = gSL_Select & "" & .預金区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.名義人) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.FAX取引先1) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.FAX取引先2) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.FAX取引先3) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.MAIL取引先1) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.MAIL取引先2) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.MAIL取引先3) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.開始発注日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.MSGTYPE) & "',"
        
        gSL_Select = gSL_Select & "'" & RTrim(.納入先事業体) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入先事業体工区) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先事業体) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.出荷場) & "',"
        
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "',"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "" & .納品書発行区分 & ","
        gSL_Select = gSL_Select & "" & .受領書発行区分 & ","
        gSL_Select = gSL_Select & "" & .支給書発行区分 & ","
        gSL_Select = gSL_Select & "'" & .号補区分 & "',"
        gSL_Select = gSL_Select & "" & .前工程区分 & ","
        gSL_Select = gSL_Select & "" & .計算位1 & ","
        gSL_Select = gSL_Select & "" & .計算位2 & ","
        
        gSL_Select = gSL_Select & "'" & RTrim(.部門CD1) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.部門CD2) & "',"
        gSL_Select = gSL_Select & "" & .正社員人員 & ","
        gSL_Select = gSL_Select & "" & .正社員稼働時間 & ","
        gSL_Select = gSL_Select & "" & .正社員原価 & ","
        gSL_Select = gSL_Select & "" & .パート人員 & ","
        gSL_Select = gSL_Select & "" & .パート稼働時間 & ","
        gSL_Select = gSL_Select & "" & .パート原価 & ","
        gSL_Select = gSL_Select & "" & .その他人員 & ","
        gSL_Select = gSL_Select & "" & .その他稼働時間 & ","
        gSL_Select = gSL_Select & "" & .その他原価 & ","
        gSL_Select = gSL_Select & "" & .日報出力 & ","
        gSL_Select = gSL_Select & "" & .負荷管理 & ","
        gSL_Select = gSL_Select & "" & .プレススポット区分 & ","
        gSL_Select = gSL_Select & "" & .指示書発行区分 & ""
        gSL_Select = gSL_Select & "," & .品番チェック桁数 & ""      ' 2009.11.06 add
        
        
        gSL_Select = gSL_Select & " )"
    
    
    
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TRMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function TRMUpdate(Optional DBNo As Integer) As Boolean
    TRMUpdate = False
    With TRM
        gSL_Select = "Update 取引先マスタ  Set "
        gSL_Select = gSL_Select & "取引先CD = '" & RTrim(.取引先CD) & "',"
        gSL_Select = gSL_Select & "取引先区分 = " & .取引先区分 & ","
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "工場CD = '" & RTrim(.工場CD) & "',"
        gSL_Select = gSL_Select & "社内区分 = " & .社内区分 & ","
        gSL_Select = gSL_Select & "契約先区分 = " & .契約先区分 & ","
        gSL_Select = gSL_Select & "正式名称 = '" & RTrim(.正式名称) & "',"
        gSL_Select = gSL_Select & "略称 = '" & RTrim(.略称) & "',"
        gSL_Select = gSL_Select & "工場名 = '" & RTrim(.工場名) & "',"
        gSL_Select = gSL_Select & "索引カナ = '" & RTrim(.索引カナ) & "',"
        gSL_Select = gSL_Select & "下請法対象 = " & .下請法対象 & ","
        gSL_Select = gSL_Select & "被自社 = '" & RTrim(.被自社) & "',"
        gSL_Select = gSL_Select & "被自社工場 = '" & RTrim(.被自社工場) & "',"
        gSL_Select = gSL_Select & "被自社CD = '" & RTrim(.被自社CD) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "計算方法 = " & .計算方法 & ","
        gSL_Select = gSL_Select & "計算単位 = " & .計算単位 & ","
        gSL_Select = gSL_Select & "締日 = " & .締日 & ","
        gSL_Select = gSL_Select & "平準化区分 = " & .平準化区分 & ","
        gSL_Select = gSL_Select & "平準化係数 = " & .平準化係数 & ","
        gSL_Select = gSL_Select & "伝票種類 = '" & RTrim(.伝票種類) & "',"
        gSL_Select = gSL_Select & "納品書形態 = '" & RTrim(.納品書形態) & "',"
        gSL_Select = gSL_Select & "受領書形態 = '" & RTrim(.受領書形態) & "',"
        gSL_Select = gSL_Select & "受領書発行 = " & .受領書発行 & ","
        gSL_Select = gSL_Select & "発注内示 = " & .発注内示 & ","
        gSL_Select = gSL_Select & "明細書発行 = '" & RTrim(.明細書発行) & "',"
        gSL_Select = gSL_Select & "前倒日数 = " & .前倒日数 & ","
        gSL_Select = gSL_Select & "郵便番号 = '" & RTrim(.郵便番号) & "',"
        gSL_Select = gSL_Select & "代表TEL = '" & RTrim(.代表TEL) & "',"
        gSL_Select = gSL_Select & "代表FAX = '" & RTrim(.代表FAX) & "',"
        gSL_Select = gSL_Select & "住所1 = '" & RTrim(.住所1) & "',"
        gSL_Select = gSL_Select & "住所2 = '" & RTrim(.住所2) & "',"
        gSL_Select = gSL_Select & "部署名1 = '" & RTrim(.部署名1) & "',"
        gSL_Select = gSL_Select & "部署名2 = '" & RTrim(.部署名2) & "',"
        gSL_Select = gSL_Select & "役職名 = '" & RTrim(.役職名) & "',"
        gSL_Select = gSL_Select & "担当者 = '" & RTrim(.担当者) & "',"
        gSL_Select = gSL_Select & "TEL1 = '" & RTrim(.TEL1) & "',"
        gSL_Select = gSL_Select & "TEL2 = '" & RTrim(.TEL2) & "',"
        gSL_Select = gSL_Select & "FAX = '" & RTrim(.FAX) & "',"
        gSL_Select = gSL_Select & "EMAIL = '" & RTrim(.EMAIL) & "',"
        gSL_Select = gSL_Select & "銀行 = '" & RTrim(.銀行) & "',"
        gSL_Select = gSL_Select & "銀行名 = '" & RTrim(.銀行名) & "',"
        gSL_Select = gSL_Select & "支店 = '" & RTrim(.支店) & "',"
        gSL_Select = gSL_Select & "支店名 = '" & RTrim(.支店名) & "',"
        gSL_Select = gSL_Select & "口座番号 = '" & RTrim(.口座番号) & "',"
        gSL_Select = gSL_Select & "預金区分 = " & .預金区分 & ","
        gSL_Select = gSL_Select & "名義人 = '" & RTrim(.名義人) & "',"
        gSL_Select = gSL_Select & "FAX取引先1 = '" & RTrim(.FAX取引先1) & "',"
        gSL_Select = gSL_Select & "FAX取引先2 = '" & RTrim(.FAX取引先2) & "',"
        gSL_Select = gSL_Select & "FAX取引先3 = '" & RTrim(.FAX取引先3) & "',"
        gSL_Select = gSL_Select & "MAIL取引先1 = '" & RTrim(.MAIL取引先1) & "',"
        gSL_Select = gSL_Select & "MAIL取引先2 = '" & RTrim(.MAIL取引先2) & "',"
        gSL_Select = gSL_Select & "MAIL取引先3 = '" & RTrim(.MAIL取引先3) & "',"
        gSL_Select = gSL_Select & "開始発注日 = '" & RTrim(.開始発注日) & "',"
        gSL_Select = gSL_Select & "MSGTYPE = '" & RTrim(.MSGTYPE) & "',"
        
        gSL_Select = gSL_Select & "納入先事業体 = '" & RTrim(.納入先事業体) & "',"
        gSL_Select = gSL_Select & "納入先事業体工区 = '" & RTrim(.納入先事業体工区) & "',"
        gSL_Select = gSL_Select & "仕入先事業体 = '" & RTrim(.仕入先事業体) & "',"
        gSL_Select = gSL_Select & "出荷場 = '" & RTrim(.出荷場) & "',"
        
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "納品書発行区分 = " & .納品書発行区分 & ","
        gSL_Select = gSL_Select & "支給書発行区分 = " & .支給書発行区分 & ","
        gSL_Select = gSL_Select & "受領書発行区分 = " & .受領書発行区分 & ","
        gSL_Select = gSL_Select & "号補区分 = '" & .号補区分 & "',"
        gSL_Select = gSL_Select & "前工程区分 = " & .前工程区分 & ","
        gSL_Select = gSL_Select & "計算位1 = " & .計算位1 & ","
        gSL_Select = gSL_Select & "計算位2 = " & .計算位2 & ","
        
        gSL_Select = gSL_Select & "部門CD1 = '" & RTrim(.部門CD1) & "',"
        gSL_Select = gSL_Select & "部門CD2 = '" & RTrim(.部門CD2) & "',"
        gSL_Select = gSL_Select & "正社員人員 = " & .正社員人員 & ","
        gSL_Select = gSL_Select & "正社員稼働時間 = " & .正社員稼働時間 & ","
        gSL_Select = gSL_Select & "正社員原価 = " & .正社員原価 & ","
        gSL_Select = gSL_Select & "パート人員 = " & .パート人員 & ","
        gSL_Select = gSL_Select & "パート稼働時間 = " & .パート稼働時間 & ","
        gSL_Select = gSL_Select & "パート原価 = " & .パート原価 & ","
        gSL_Select = gSL_Select & "その他人員 = " & .その他人員 & ","
        gSL_Select = gSL_Select & "その他稼働時間 = " & .その他稼働時間 & ","
        gSL_Select = gSL_Select & "その他原価 = " & .その他原価 & ","
        gSL_Select = gSL_Select & "日報出力 = " & .日報出力 & ","
        gSL_Select = gSL_Select & "負荷管理 = " & .負荷管理 & ","
        gSL_Select = gSL_Select & "プレススポット区分 = " & .プレススポット区分 & ","
        gSL_Select = gSL_Select & "指示書発行区分 = " & .指示書発行区分 & ""
        gSL_Select = gSL_Select & ", 品番チェック桁数 = " & .品番チェック桁数 & ""      ' 2009.11.06 add
        
        gSL_Select = gSL_Select & " Where 取引先CD = '" & RTrim(.取引先CD) & "'"
        gSL_Select = gSL_Select & "   And 取引先区分 = " & .取引先区分 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TRMUpdate = True
End Function
'*--------------------------------------------------------------------------------
'*
'*      名称を求める
'*
'*      モジュールID    : TornmGet
'*
'*
'*          引数        : Torcd   =   区分名
'*                      : TorKB   =   0:全て,1:受注,2:受注契約先,3:受注・契約先,
'*                                           4:発注,5:発注契約先,6:社内,7:発注・社内
'*                                           8:発注･契約先,9:発注･契約先･社内
'*                      : KeiKB   =   0:契約先以外,1:契約先
'*          　　        : DBNo    =   データベースNO ( Optional )
'*
'*          戻値        : True Or False
'*                      : TRM_Name : 正式名称
'*                      : TRM_RName : 略称
'*                      : TRM_Shita : 下請法区分
'*
'*      作成日          :   2003/11/18   By Cis
'*      変更日          :
'*
'*---------------------------------------------------------------------------------
Public Function TorNmGet(Torcd As String, TorKB As Integer, Optional DBNo As Integer) As Boolean
    TorNmGet = False
    
    TRM_Name = ""
    TRM_RName = ""
    TRM_Shita = 0
    TRM_Fax = ""
    TRM_EMail = ""
    
    gSL_Select = "SELECT 正式名称,略称,下請法対象,計算単位,計算方法,締日,FAX,EMAIL,工場名,被自社,被自社工場"
    gSL_Select = gSL_Select & " FROM 取引先マスタ "
    gSL_Select = gSL_Select & " Where 取引先CD = '" & RTrim(Torcd) & "'"
    Select Case TorKB
        Case 1
            gSL_Select = gSL_Select & "   And 取引先区分 = 0 And 契約先区分 = 0"
        Case 2
            gSL_Select = gSL_Select & "   And 取引先区分 = 0 And 契約先区分 = 1"
        Case 3
            gSL_Select = gSL_Select & "   And 取引先区分 = 0 "
        Case 4
            gSL_Select = gSL_Select & "   And 取引先区分 = 1 And 契約先区分 = 0 And 社内区分 = 0"
        Case 5
            gSL_Select = gSL_Select & "   And 取引先区分 = 1 And 契約先区分 = 1 And 社内区分 = 0"
        Case 6
            gSL_Select = gSL_Select & "   And 取引先区分 = 1 And 契約先区分 = 0 And 社内区分 = 1"
        Case 7
            gSL_Select = gSL_Select & "   And 取引先区分 = 1 And 契約先区分 = 0 "
        Case 8
            gSL_Select = gSL_Select & "   And 取引先区分 = 1 And 社内区分 = 0 "
        Case 9
            gSL_Select = gSL_Select & "   And 取引先区分 = 1 "
    End Select
    If Not TRMRead(gSL_Select, 1, DBNo) Then Exit Function

    TorNmGet = True
    
    TRM_Name = RTrim(TRM.正式名称)
    TRM_RName = RTrim(TRM.略称)
    TRM_Shita = TRM.下請法対象
    TRM_Fax = TRM.FAX
    TRM_EMail = TRM.EMAIL
End Function



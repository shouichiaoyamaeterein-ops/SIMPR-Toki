Attribute VB_Name = "材料取引先マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料取引先マスタ
'*
'*   モジュールID  : 材料取引先マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/04/02  By Cis
'*   作成日  :  2007/12/10  By Cis
'*   作成日  :  2008/02/12  By Cis
'*   作成日  :  2008/03/01  By Cis
'*   作成日  :  2008/06/27  By Cis(FMV-E630)
'*   作成日  :  2009/01/15  By Cis(FMV-E630)
'*   作成日  :  2009/02/27  By Cis(HP-DC7900)
'*   作成日  :  2009/09/23  By Cis(FMVW5220)
'*   作成日  :  2010/01/15  By Cis(HP-DX7500)
'*   作成日  :  2010/01/16  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料取引先マスタ
        材料管理番号      As Long
        取引先            As String
        受入              As String
        材質              As String
        寸法区分          As Byte
        板厚              As Currency
        幅                As Currency
        長さ              As Currency
        背番号            As String
        SK区分            As Byte
        手配区分          As Byte
        メーカー          As String
        メーカー受入      As String
        契約先            As String
        サイクル          As String
        収容数１          As Currency
        収容器１          As String
        ロット１          As Currency
        変更日            As String
        変更便            As String
        収容数２          As Currency
        収容器２          As String
        ロット２          As Currency
        置場              As String
        最終発行連番      As Integer
        係数１            As Currency
        係数２            As Currency
        係数３            As Currency
        LT                As Byte
        Week１            As Byte
        Week２            As Byte
        Week３            As Byte
        Week４            As Byte
        Week５            As Byte
        Ｄａｙ１          As Byte
        Ｄａｙ２          As Byte
        Ｄａｙ３          As Byte
        Ｄａｙ４          As Byte
        Ｄａｙ５          As Byte
        発注比率          As Byte
        有効期限開始      As String
        有効期限終了      As String
        内示発注区分      As Byte
        アドレス          As String
        伝票種類          As String
        納品書形態        As String
        受領書形態        As String
        受領書発行        As Byte
        明細書発行        As String
        備考              As String
        PNO               As Byte
        納品書発行区分    As Byte
        支給書発行区分    As Byte
        受領書発行区分    As Byte
        受給品番          As String
        発注品番          As String
        自社品番          As String
        加工先            As String
        加工先背景色      As Currency
        得意先            As String
        かんばんサイズ    As Byte
        現品票サイズ      As Byte
        使用品番1         As String
        使用品番2         As String
        前工程区分        As Byte
        実績換算値        As Currency
        材質表示          As Byte
        直送伝票発行区分  As Byte
        棚卸区分          As Byte
'        作成日            As Date
        作成者            As String
'        作成端末          As String
'        更新日            As Date
        更新者            As String
'        更新端末          As String
        件数          As Long
'   項目追加(自動)
	手配先名       As String
	手配区分名     As String
	契約先名       As String
	メーカー名     As String
	伝票種類名     As String
	納品書形態名   As String
	受領書形態名   As String
	明細書発行名   As String
	品番           As String
	区分           As Byte
	表示寸法       As String
	材料名         As String
	手配単位名     As String
	表示品番       As String
	材種コード     As String
	材種名         As String
	在庫管理名     As String
	在庫管理区分   As Byte
	在庫管理数値   As Currency
	在庫場所       As String
	実棚日         As String
	材寸           As String
	仕様           As String
	単価           As Currency
	金額           As Currency
	最終出庫日     As String
	在庫数         As Currency
	取引先名       As String
    End Type

    Public ZTM As T_材料取引先マスタ

    Public ZTM_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZTM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZTM	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZTM
            Select Case adoCol.Name
                Case "材料管理番号":      If Not IsNull(adoCol.Value) Then .材料管理番号     = adoCol.Value
                Case "取引先":            If Not IsNull(adoCol.Value) Then .取引先           = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入             = adoCol.Value
                Case "材質":              If Not IsNull(adoCol.Value) Then .材質             = adoCol.Value
                Case "寸法区分":          If Not IsNull(adoCol.Value) Then .寸法区分         = adoCol.Value
                Case "板厚":              If Not IsNull(adoCol.Value) Then .板厚             = adoCol.Value
                Case "幅":                If Not IsNull(adoCol.Value) Then .幅               = adoCol.Value
                Case "長さ":              If Not IsNull(adoCol.Value) Then .長さ             = adoCol.Value
                Case "背番号":            If Not IsNull(adoCol.Value) Then .背番号           = adoCol.Value
                Case "SK区分":            If Not IsNull(adoCol.Value) Then .SK区分           = adoCol.Value
                Case "手配区分":          If Not IsNull(adoCol.Value) Then .手配区分         = adoCol.Value
                Case "メーカー":          If Not IsNull(adoCol.Value) Then .メーカー         = adoCol.Value
                Case "メーカー受入":      If Not IsNull(adoCol.Value) Then .メーカー受入     = adoCol.Value
                Case "契約先":            If Not IsNull(adoCol.Value) Then .契約先           = adoCol.Value
                Case "サイクル":          If Not IsNull(adoCol.Value) Then .サイクル         = adoCol.Value
                Case "収容数１":          If Not IsNull(adoCol.Value) Then .収容数１         = adoCol.Value
                Case "収容器１":          If Not IsNull(adoCol.Value) Then .収容器１         = adoCol.Value
                Case "ロット１":          If Not IsNull(adoCol.Value) Then .ロット１         = adoCol.Value
                Case "変更日":            If Not IsNull(adoCol.Value) Then .変更日           = adoCol.Value
                Case "変更便":            If Not IsNull(adoCol.Value) Then .変更便           = adoCol.Value
                Case "収容数２":          If Not IsNull(adoCol.Value) Then .収容数２         = adoCol.Value
                Case "収容器２":          If Not IsNull(adoCol.Value) Then .収容器２         = adoCol.Value
                Case "ロット２":          If Not IsNull(adoCol.Value) Then .ロット２         = adoCol.Value
                Case "置場":              If Not IsNull(adoCol.Value) Then .置場             = adoCol.Value
                Case "最終発行連番":      If Not IsNull(adoCol.Value) Then .最終発行連番     = adoCol.Value
                Case "係数１":            If Not IsNull(adoCol.Value) Then .係数１           = adoCol.Value
                Case "係数２":            If Not IsNull(adoCol.Value) Then .係数２           = adoCol.Value
                Case "係数３":            If Not IsNull(adoCol.Value) Then .係数３           = adoCol.Value
                Case "LT":                If Not IsNull(adoCol.Value) Then .LT               = adoCol.Value
                Case "Week１":            If Not IsNull(adoCol.Value) Then .Week１           = adoCol.Value
                Case "Week２":            If Not IsNull(adoCol.Value) Then .Week２           = adoCol.Value
                Case "Week３":            If Not IsNull(adoCol.Value) Then .Week３           = adoCol.Value
                Case "Week４":            If Not IsNull(adoCol.Value) Then .Week４           = adoCol.Value
                Case "Week５":            If Not IsNull(adoCol.Value) Then .Week５           = adoCol.Value
                Case "Ｄａｙ１":          If Not IsNull(adoCol.Value) Then .Ｄａｙ１         = adoCol.Value
                Case "Ｄａｙ２":          If Not IsNull(adoCol.Value) Then .Ｄａｙ２         = adoCol.Value
                Case "Ｄａｙ３":          If Not IsNull(adoCol.Value) Then .Ｄａｙ３         = adoCol.Value
                Case "Ｄａｙ４":          If Not IsNull(adoCol.Value) Then .Ｄａｙ４         = adoCol.Value
                Case "Ｄａｙ５":          If Not IsNull(adoCol.Value) Then .Ｄａｙ５         = adoCol.Value
                Case "発注比率":          If Not IsNull(adoCol.Value) Then .発注比率         = adoCol.Value
                Case "有効期限開始":      If Not IsNull(adoCol.Value) Then .有効期限開始     = adoCol.Value
                Case "有効期限終了":      If Not IsNull(adoCol.Value) Then .有効期限終了     = adoCol.Value
                Case "内示発注区分":      If Not IsNull(adoCol.Value) Then .内示発注区分     = adoCol.Value
                Case "アドレス":          If Not IsNull(adoCol.Value) Then .アドレス         = adoCol.Value
                Case "伝票種類":          If Not IsNull(adoCol.Value) Then .伝票種類         = adoCol.Value
                Case "納品書形態":        If Not IsNull(adoCol.Value) Then .納品書形態       = adoCol.Value
                Case "受領書形態":        If Not IsNull(adoCol.Value) Then .受領書形態       = adoCol.Value
                Case "受領書発行":        If Not IsNull(adoCol.Value) Then .受領書発行       = adoCol.Value
                Case "明細書発行":        If Not IsNull(adoCol.Value) Then .明細書発行       = adoCol.Value
                Case "備考":              If Not IsNull(adoCol.Value) Then .備考             = adoCol.Value
                Case "PNO":               If Not IsNull(adoCol.Value) Then .PNO              = adoCol.Value
                Case "納品書発行区分":    If Not IsNull(adoCol.Value) Then .納品書発行区分   = adoCol.Value
                Case "支給書発行区分":    If Not IsNull(adoCol.Value) Then .支給書発行区分   = adoCol.Value
                Case "受領書発行区分":    If Not IsNull(adoCol.Value) Then .受領書発行区分   = adoCol.Value
                Case "受給品番":          If Not IsNull(adoCol.Value) Then .受給品番         = adoCol.Value
                Case "発注品番":          If Not IsNull(adoCol.Value) Then .発注品番         = adoCol.Value
                Case "自社品番":          If Not IsNull(adoCol.Value) Then .自社品番         = adoCol.Value
                Case "加工先":            If Not IsNull(adoCol.Value) Then .加工先           = adoCol.Value
                Case "加工先背景色":      If Not IsNull(adoCol.Value) Then .加工先背景色     = adoCol.Value
                Case "得意先":            If Not IsNull(adoCol.Value) Then .得意先           = adoCol.Value
                Case "かんばんサイズ":    If Not IsNull(adoCol.Value) Then .かんばんサイズ   = adoCol.Value
                Case "現品票サイズ":      If Not IsNull(adoCol.Value) Then .現品票サイズ     = adoCol.Value
                Case "使用品番1":         If Not IsNull(adoCol.Value) Then .使用品番1        = adoCol.Value
                Case "使用品番2":         If Not IsNull(adoCol.Value) Then .使用品番2        = adoCol.Value
                Case "前工程区分":        If Not IsNull(adoCol.Value) Then .前工程区分       = adoCol.Value
                Case "実績換算値":        If Not IsNull(adoCol.Value) Then .実績換算値       = adoCol.Value
                Case "材質表示":          If Not IsNull(adoCol.Value) Then .材質表示         = adoCol.Value
                Case "直送伝票発行区分":  If Not IsNull(adoCol.Value) Then .直送伝票発行区分 = adoCol.Value
                Case "棚卸区分":          If Not IsNull(adoCol.Value) Then .棚卸区分         = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
                Case "更新者":            If Not IsNull(adoCol.Value) Then .更新者           = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "手配先名":	If Not IsNull(adoCol.Value) Then .手配先名	= adoCol.Value
		Case "手配区分名":	If Not IsNull(adoCol.Value) Then .手配区分名	= adoCol.Value
		Case "契約先名":	If Not IsNull(adoCol.Value) Then .契約先名	= adoCol.Value
		Case "メーカー名":	If Not IsNull(adoCol.Value) Then .メーカー名	= adoCol.Value
		Case "伝票種類名":	If Not IsNull(adoCol.Value) Then .伝票種類名	= adoCol.Value
		Case "納品書形態名":	If Not IsNull(adoCol.Value) Then .納品書形態名	= adoCol.Value
		Case "受領書形態名":	If Not IsNull(adoCol.Value) Then .受領書形態名	= adoCol.Value
		Case "明細書発行名":	If Not IsNull(adoCol.Value) Then .明細書発行名	= adoCol.Value
		Case "品番":	If Not IsNull(adoCol.Value) Then .品番	= adoCol.Value
		Case "区分":	If Not IsNull(adoCol.Value) Then .区分	= adoCol.Value
		Case "表示寸法":	If Not IsNull(adoCol.Value) Then .表示寸法	= adoCol.Value
		Case "材料名":	If Not IsNull(adoCol.Value) Then .材料名	= adoCol.Value
		Case "手配単位名":	If Not IsNull(adoCol.Value) Then .手配単位名	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "材種コード":	If Not IsNull(adoCol.Value) Then .材種コード	= adoCol.Value
		Case "材種名":	If Not IsNull(adoCol.Value) Then .材種名	= adoCol.Value
		Case "在庫管理名":	If Not IsNull(adoCol.Value) Then .在庫管理名	= adoCol.Value
		Case "在庫管理区分":	If Not IsNull(adoCol.Value) Then .在庫管理区分	= adoCol.Value
		Case "在庫管理数値":	If Not IsNull(adoCol.Value) Then .在庫管理数値	= adoCol.Value
		Case "在庫場所":	If Not IsNull(adoCol.Value) Then .在庫場所	= adoCol.Value
		Case "実棚日":	If Not IsNull(adoCol.Value) Then .実棚日	= adoCol.Value
		Case "材寸":	If Not IsNull(adoCol.Value) Then .材寸	= adoCol.Value
		Case "仕様":	If Not IsNull(adoCol.Value) Then .仕様	= adoCol.Value
		Case "単価":	If Not IsNull(adoCol.Value) Then .単価	= adoCol.Value
		Case "金額":	If Not IsNull(adoCol.Value) Then .金額	= adoCol.Value
		Case "最終出庫日":	If Not IsNull(adoCol.Value) Then .最終出庫日	= adoCol.Value
		Case "在庫数":	If Not IsNull(adoCol.Value) Then .在庫数	= adoCol.Value
		Case "取引先名":	If Not IsNull(adoCol.Value) Then .取引先名	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZTMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZTMRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 材料取引先マスタ"
        RTxtSql = RTxtSql &" Where 材料管理番号 = " & ZTM.材料管理番号 & ""
        RTxtSql = RTxtSql &" And   取引先 = '" & RTrim(ZTM.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(ZTM.受入) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZTMRead = True
           Call ItemsFromZTM(DBNo)
	   ZTM_RDSTS = True
        else
	   ZTM_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZTMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZTMReadNext(Optional DBNo As Integer = 0)
    If Not ZTM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZTM_RDSTS = False
        Else
           Call ItemsFromZTM(DBNo)
           ZTM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZTMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZTM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZTM )
'*----------------------------------------------*'
Public Sub ItemsClearZTM()
    With ZTM
	.材料管理番号 = 0
	.取引先 = ""
	.受入 = ""
	.材質 = ""
	.寸法区分 = 0
	.板厚 = 0
	.幅 = 0
	.長さ = 0
	.背番号 = ""
	.SK区分 = 0
	.手配区分 = 0
	.メーカー = ""
	.メーカー受入 = ""
	.契約先 = ""
	.サイクル = ""
	.収容数１ = 0
	.収容器１ = ""
	.ロット１ = 0
	.変更日 = ""
	.変更便 = ""
	.収容数２ = 0
	.収容器２ = ""
	.ロット２ = 0
	.置場 = ""
	.最終発行連番 = 0
	.係数１ = 0
	.係数２ = 0
	.係数３ = 0
	.LT = 0
	.Week１ = 0
	.Week２ = 0
	.Week３ = 0
	.Week４ = 0
	.Week５ = 0
	.Ｄａｙ１ = 0
	.Ｄａｙ２ = 0
	.Ｄａｙ３ = 0
	.Ｄａｙ４ = 0
	.Ｄａｙ５ = 0
	.発注比率 = 0
	.有効期限開始 = ""
	.有効期限終了 = ""
	.内示発注区分 = 0
	.アドレス = ""
	.伝票種類 = ""
	.納品書形態 = ""
	.受領書形態 = ""
	.受領書発行 = 0
	.明細書発行 = ""
	.備考 = ""
	.PNO = 0
	.納品書発行区分 = 0
	.支給書発行区分 = 0
	.受領書発行区分 = 0
	.受給品番 = ""
	.発注品番 = ""
	.自社品番 = ""
	.加工先 = ""
	.加工先背景色 = 0
	.得意先 = ""
	.かんばんサイズ = 0
	.現品票サイズ = 0
	.使用品番1 = ""
	.使用品番2 = ""
	.前工程区分 = 0
	.実績換算値 = 0
	.材質表示 = 0
	.直送伝票発行区分 = 0
	.棚卸区分 = 0
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
'	.更新日 = 0
	.更新者 = ""
'	.更新端末 = ""
        .件数 = 0
'   項目追加(自動)
	.手配先名 = ""
	.手配区分名 = ""
	.契約先名 = ""
	.メーカー名 = ""
	.伝票種類名 = ""
	.納品書形態名 = ""
	.受領書形態名 = ""
	.明細書発行名 = ""
	.品番 = ""
	.区分 = 0
	.表示寸法 = ""
	.材料名 = ""
	.手配単位名 = ""
	.表示品番 = ""
	.材種コード = ""
	.材種名 = ""
	.在庫管理名 = ""
	.在庫管理区分 = 0
	.在庫管理数値 = 0
	.在庫場所 = ""
	.実棚日 = ""
	.材寸 = ""
	.仕様 = ""
	.単価 = 0
	.金額 = 0
	.最終出庫日 = ""
	.在庫数 = 0
	.取引先名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZTMInsert(Optional DBNo As Integer) As Boolean
    ZTMInsert = False
    With ZTM
        gSL_Select = "Insert Into 材料取引先マスタ  ( "
        gSL_Select = gSL_Select & "材料管理番号,取引先,受入,"
        gSL_Select = gSL_Select & "材質,寸法区分,板厚,"
        gSL_Select = gSL_Select & "幅,長さ,背番号,"
        gSL_Select = gSL_Select & "SK区分,手配区分,メーカー,"
        gSL_Select = gSL_Select & "メーカー受入,契約先,サイクル,"
        gSL_Select = gSL_Select & "収容数１,収容器１,ロット１,"
        gSL_Select = gSL_Select & "変更日,変更便,収容数２,"
        gSL_Select = gSL_Select & "収容器２,ロット２,置場,"
        gSL_Select = gSL_Select & "最終発行連番,係数１,係数２,"
        gSL_Select = gSL_Select & "係数３,LT,Week１,"
        gSL_Select = gSL_Select & "Week２,Week３,Week４,"
        gSL_Select = gSL_Select & "Week５,Ｄａｙ１,Ｄａｙ２,"
        gSL_Select = gSL_Select & "Ｄａｙ３,Ｄａｙ４,Ｄａｙ５,"
        gSL_Select = gSL_Select & "発注比率,有効期限開始,有効期限終了,"
        gSL_Select = gSL_Select & "内示発注区分,アドレス,伝票種類,"
        gSL_Select = gSL_Select & "納品書形態,受領書形態,受領書発行,"
        gSL_Select = gSL_Select & "明細書発行,備考,PNO,"
        gSL_Select = gSL_Select & "納品書発行区分,支給書発行区分,受領書発行区分,"
        gSL_Select = gSL_Select & "受給品番,発注品番,自社品番,"
        gSL_Select = gSL_Select & "加工先,加工先背景色,得意先,"
        gSL_Select = gSL_Select & "かんばんサイズ,現品票サイズ,使用品番1,"
        gSL_Select = gSL_Select & "使用品番2,前工程区分,実績換算値,"
        gSL_Select = gSL_Select & "材質表示,直送伝票発行区分,棚卸区分,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"		
        gSL_Select = gSL_Select & "" & .寸法区分 & ","		
        gSL_Select = gSL_Select & "" & .板厚 & ","		
        gSL_Select = gSL_Select & "" & .幅 & ","		
        gSL_Select = gSL_Select & "" & .長さ & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "" & .SK区分 & ","		
        gSL_Select = gSL_Select & "" & .手配区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.メーカー) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.メーカー受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.サイクル) & "',"		
        gSL_Select = gSL_Select & "" & .収容数１ & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.収容器１) & "',"		
        gSL_Select = gSL_Select & "" & .ロット１ & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.変更日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.変更便) & "',"		
        gSL_Select = gSL_Select & "" & .収容数２ & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.収容器２) & "',"		
        gSL_Select = gSL_Select & "" & .ロット２ & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.置場) & "',"		
        gSL_Select = gSL_Select & "" & .最終発行連番 & ","		
        gSL_Select = gSL_Select & "" & .係数１ & ","		
        gSL_Select = gSL_Select & "" & .係数２ & ","		
        gSL_Select = gSL_Select & "" & .係数３ & ","		
        gSL_Select = gSL_Select & "" & .LT & ","		
        gSL_Select = gSL_Select & "" & .Week１ & ","		
        gSL_Select = gSL_Select & "" & .Week２ & ","		
        gSL_Select = gSL_Select & "" & .Week３ & ","		
        gSL_Select = gSL_Select & "" & .Week４ & ","		
        gSL_Select = gSL_Select & "" & .Week５ & ","		
        gSL_Select = gSL_Select & "" & .Ｄａｙ１ & ","		
        gSL_Select = gSL_Select & "" & .Ｄａｙ２ & ","		
        gSL_Select = gSL_Select & "" & .Ｄａｙ３ & ","		
        gSL_Select = gSL_Select & "" & .Ｄａｙ４ & ","		
        gSL_Select = gSL_Select & "" & .Ｄａｙ５ & ","		
        gSL_Select = gSL_Select & "" & .発注比率 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.有効期限開始) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.有効期限終了) & "',"		
        gSL_Select = gSL_Select & "" & .内示発注区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.アドレス) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.伝票種類) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納品書形態) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受領書形態) & "',"		
        gSL_Select = gSL_Select & "" & .受領書発行 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.明細書発行) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.備考) & "',"		
        gSL_Select = gSL_Select & "" & .PNO & ","		
        gSL_Select = gSL_Select & "" & .納品書発行区分 & ","		
        gSL_Select = gSL_Select & "" & .支給書発行区分 & ","		
        gSL_Select = gSL_Select & "" & .受領書発行区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.受給品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.自社品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.加工先) & "',"		
        gSL_Select = gSL_Select & "" & .加工先背景色 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.得意先) & "',"		
        gSL_Select = gSL_Select & "" & .かんばんサイズ & ","		
        gSL_Select = gSL_Select & "" & .現品票サイズ & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.使用品番1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.使用品番2) & "',"		
        gSL_Select = gSL_Select & "" & .前工程区分 & ","		
        gSL_Select = gSL_Select & "" & .実績換算値 & ","		
        gSL_Select = gSL_Select & "" & .材質表示 & ","		
        gSL_Select = gSL_Select & "" & .直送伝票発行区分 & ","		
        gSL_Select = gSL_Select & "" & .棚卸区分 & ","		
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
    ZTMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZTMUpdate(Optional DBNo As Integer) As Boolean
    ZTMUpdate = False	
    With ZTM
        gSL_Select = "Update 材料取引先マスタ  Set "
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "寸法区分 = " & .寸法区分 & ","
        gSL_Select = gSL_Select & "板厚 = " & .板厚 & ","
        gSL_Select = gSL_Select & "幅 = " & .幅 & ","
        gSL_Select = gSL_Select & "長さ = " & .長さ & ","
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "SK区分 = " & .SK区分 & ","
        gSL_Select = gSL_Select & "手配区分 = " & .手配区分 & ","
        gSL_Select = gSL_Select & "メーカー = '" & RTrim(.メーカー) & "',"
        gSL_Select = gSL_Select & "メーカー受入 = '" & RTrim(.メーカー受入) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "サイクル = '" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "収容数１ = " & .収容数１ & ","
        gSL_Select = gSL_Select & "収容器１ = '" & RTrim(.収容器１) & "',"
        gSL_Select = gSL_Select & "ロット１ = " & .ロット１ & ","
        gSL_Select = gSL_Select & "変更日 = '" & RTrim(.変更日) & "',"
        gSL_Select = gSL_Select & "変更便 = '" & RTrim(.変更便) & "',"
        gSL_Select = gSL_Select & "収容数２ = " & .収容数２ & ","
        gSL_Select = gSL_Select & "収容器２ = '" & RTrim(.収容器２) & "',"
        gSL_Select = gSL_Select & "ロット２ = " & .ロット２ & ","
        gSL_Select = gSL_Select & "置場 = '" & RTrim(.置場) & "',"
        gSL_Select = gSL_Select & "最終発行連番 = " & .最終発行連番 & ","
        gSL_Select = gSL_Select & "係数１ = " & .係数１ & ","
        gSL_Select = gSL_Select & "係数２ = " & .係数２ & ","
        gSL_Select = gSL_Select & "係数３ = " & .係数３ & ","
        gSL_Select = gSL_Select & "LT = " & .LT & ","
        gSL_Select = gSL_Select & "Week１ = " & .Week１ & ","
        gSL_Select = gSL_Select & "Week２ = " & .Week２ & ","
        gSL_Select = gSL_Select & "Week３ = " & .Week３ & ","
        gSL_Select = gSL_Select & "Week４ = " & .Week４ & ","
        gSL_Select = gSL_Select & "Week５ = " & .Week５ & ","
        gSL_Select = gSL_Select & "Ｄａｙ１ = " & .Ｄａｙ１ & ","
        gSL_Select = gSL_Select & "Ｄａｙ２ = " & .Ｄａｙ２ & ","
        gSL_Select = gSL_Select & "Ｄａｙ３ = " & .Ｄａｙ３ & ","
        gSL_Select = gSL_Select & "Ｄａｙ４ = " & .Ｄａｙ４ & ","
        gSL_Select = gSL_Select & "Ｄａｙ５ = " & .Ｄａｙ５ & ","
        gSL_Select = gSL_Select & "発注比率 = " & .発注比率 & ","
        gSL_Select = gSL_Select & "有効期限開始 = '" & RTrim(.有効期限開始) & "',"
        gSL_Select = gSL_Select & "有効期限終了 = '" & RTrim(.有効期限終了) & "',"
        gSL_Select = gSL_Select & "内示発注区分 = " & .内示発注区分 & ","
        gSL_Select = gSL_Select & "アドレス = '" & RTrim(.アドレス) & "',"
        gSL_Select = gSL_Select & "伝票種類 = '" & RTrim(.伝票種類) & "',"
        gSL_Select = gSL_Select & "納品書形態 = '" & RTrim(.納品書形態) & "',"
        gSL_Select = gSL_Select & "受領書形態 = '" & RTrim(.受領書形態) & "',"
        gSL_Select = gSL_Select & "受領書発行 = " & .受領書発行 & ","
        gSL_Select = gSL_Select & "明細書発行 = '" & RTrim(.明細書発行) & "',"
        gSL_Select = gSL_Select & "備考 = '" & RTrim(.備考) & "',"
        gSL_Select = gSL_Select & "PNO = " & .PNO & ","
        gSL_Select = gSL_Select & "納品書発行区分 = " & .納品書発行区分 & ","
        gSL_Select = gSL_Select & "支給書発行区分 = " & .支給書発行区分 & ","
        gSL_Select = gSL_Select & "受領書発行区分 = " & .受領書発行区分 & ","
        gSL_Select = gSL_Select & "受給品番 = '" & RTrim(.受給品番) & "',"
        gSL_Select = gSL_Select & "発注品番 = '" & RTrim(.発注品番) & "',"
        gSL_Select = gSL_Select & "自社品番 = '" & RTrim(.自社品番) & "',"
        gSL_Select = gSL_Select & "加工先 = '" & RTrim(.加工先) & "',"
        gSL_Select = gSL_Select & "加工先背景色 = " & .加工先背景色 & ","
        gSL_Select = gSL_Select & "得意先 = '" & RTrim(.得意先) & "',"
        gSL_Select = gSL_Select & "かんばんサイズ = " & .かんばんサイズ & ","
        gSL_Select = gSL_Select & "現品票サイズ = " & .現品票サイズ & ","
        gSL_Select = gSL_Select & "使用品番1 = '" & RTrim(.使用品番1) & "',"
        gSL_Select = gSL_Select & "使用品番2 = '" & RTrim(.使用品番2) & "',"
        gSL_Select = gSL_Select & "前工程区分 = " & .前工程区分 & ","
        gSL_Select = gSL_Select & "実績換算値 = " & .実績換算値 & ","
        gSL_Select = gSL_Select & "材質表示 = " & .材質表示 & ","
        gSL_Select = gSL_Select & "直送伝票発行区分 = " & .直送伝票発行区分 & ","
        gSL_Select = gSL_Select & "棚卸区分 = " & .棚卸区分 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 材料管理番号 = " & .材料管理番号 & ""
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZTMUpdate = True
End Function

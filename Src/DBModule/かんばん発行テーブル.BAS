Attribute VB_Name = "かんばん発行テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   かんばん発行テーブル
'*
'*   モジュールID  : かんばん発行テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/03/24  By Cis
'*   作成日  :  2009/10/20  By Cis(HP-DC7900)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_かんばん発行テーブル
        かんばん種類          As Byte
        品番                  As String
        取引先                As String
        受入                  As String
        発行NO                As Integer
        発行種類              As Byte
        材料管理番号          As Long
        背番号                As String
        版数                  As Byte
        発行区分              As Byte
        表示品番              As String
        品番14                As String
        品番30                As String
        品目                  As Byte
        品名                  As String
        受入名                As String
        収容器                As String
        収容数                As Currency
        収容区分              As Byte
        手配区分              As String
        手配区分名            As String
        アドレス              As String
        IMAGE                 As String
        取引先CD              As String
        取引先名              As String
        取引先名1             As String
        取引先名2             As String
        取引先工場名          As String
        取引先工場名1         As String
        取引先工場名2         As String
        納入先固定            As String
        納入先                As String
        納入先CD              As String
        納入先名              As String
        納入先名1             As String
        納入先名2             As String
        納入先工場名          As String
        納入先受入            As String
        メーカ固定            As String
        メーカ                As String
        メーカCD              As String
        メーカ名              As String
        メーカ名1             As String
        メーカ名2             As String
        メーカ受入            As String
        サイクル              As String
        置場                  As String
        車型                  As String
        自社CD                As String
        自社名1               As String
        自社名2               As String
        工程                  As String
        工程名                As String
        次工程                As String
        次工程名              As String
        基準工数              As Currency
        生産ロット            As Long
        基準生産日            As String
        管理番号              As String
        材質                  As String
        材質18                As String
        寸法区分              As Byte
        板厚                  As Currency
        幅                    As Currency
        長さ                  As Currency
        材料名                As String
        使用品番              As String
        単位名                As String
        手配単位名            As String
        実数区分              As Byte
        材種コード            As String
        備考                  As String
        StHNo                 As Byte
        EdHNo                 As Byte
        Stack1                As Byte
        Stack2                As Byte
        区分                  As String
        納入時間1             As String
        納入時間2             As String
        納入時間3             As String
        納入時間4             As String
        納入時間5             As String
        納入時間6             As String
        納入時間7             As String
        納入時間8             As String
        得意先背番1           As String
        得意先背番2           As String
        得意先背番3           As String
        得意先背番4           As String
        得意先工場名1         As String
        得意先工場名2         As String
        得意先工場名3         As String
        得意先工場名4         As String
        部品区分              As Byte
        ポイント              As String
        ポイントIMAGE         As String
        部署                  As String
        部署名                As String
        返却場所              As String
        返却場所名            As String
        後工程                As String
        後工程名1             As String
        後工程名2             As String
        得意先                As String
        得意先名              As String
        生産区分              As String
        かんばん枠色          As Currency
        発行済                As Byte
        発行FLG               As String
'        作成日                As Date
        作成者                As String
'        作成端末              As String
        件数          As Long
'   項目追加(自動)
	最終発行連番   As Integer
    End Type

    Public KHT As T_かんばん発行テーブル

    Public KHT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKHT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKHT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KHT
            Select Case adoCol.Name
                Case "かんばん種類":          If Not IsNull(adoCol.Value) Then .かんばん種類         = adoCol.Value
                Case "品番":                  If Not IsNull(adoCol.Value) Then .品番                 = adoCol.Value
                Case "取引先":                If Not IsNull(adoCol.Value) Then .取引先               = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入                 = adoCol.Value
                Case "発行NO":                If Not IsNull(adoCol.Value) Then .発行NO               = adoCol.Value
                Case "発行種類":              If Not IsNull(adoCol.Value) Then .発行種類             = adoCol.Value
                Case "材料管理番号":          If Not IsNull(adoCol.Value) Then .材料管理番号         = adoCol.Value
                Case "背番号":                If Not IsNull(adoCol.Value) Then .背番号               = adoCol.Value
                Case "版数":                  If Not IsNull(adoCol.Value) Then .版数                 = adoCol.Value
                Case "発行区分":              If Not IsNull(adoCol.Value) Then .発行区分             = adoCol.Value
                Case "表示品番":              If Not IsNull(adoCol.Value) Then .表示品番             = adoCol.Value
                Case "品番14":                If Not IsNull(adoCol.Value) Then .品番14               = adoCol.Value
                Case "品番30":                If Not IsNull(adoCol.Value) Then .品番30               = adoCol.Value
                Case "品目":                  If Not IsNull(adoCol.Value) Then .品目                 = adoCol.Value
                Case "品名":                  If Not IsNull(adoCol.Value) Then .品名                 = adoCol.Value
                Case "受入名":                If Not IsNull(adoCol.Value) Then .受入名               = adoCol.Value
                Case "収容器":                If Not IsNull(adoCol.Value) Then .収容器               = adoCol.Value
                Case "収容数":                If Not IsNull(adoCol.Value) Then .収容数               = adoCol.Value
                Case "収容区分":              If Not IsNull(adoCol.Value) Then .収容区分             = adoCol.Value
                Case "手配区分":              If Not IsNull(adoCol.Value) Then .手配区分             = adoCol.Value
                Case "手配区分名":            If Not IsNull(adoCol.Value) Then .手配区分名           = adoCol.Value
                Case "アドレス":              If Not IsNull(adoCol.Value) Then .アドレス             = adoCol.Value
                Case "IMAGE":                 If Not IsNull(adoCol.Value) Then .IMAGE                = adoCol.Value
                Case "取引先CD":              If Not IsNull(adoCol.Value) Then .取引先CD             = adoCol.Value
                Case "取引先名":              If Not IsNull(adoCol.Value) Then .取引先名             = adoCol.Value
                Case "取引先名1":             If Not IsNull(adoCol.Value) Then .取引先名1            = adoCol.Value
                Case "取引先名2":             If Not IsNull(adoCol.Value) Then .取引先名2            = adoCol.Value
                Case "取引先工場名":          If Not IsNull(adoCol.Value) Then .取引先工場名         = adoCol.Value
                Case "取引先工場名1":         If Not IsNull(adoCol.Value) Then .取引先工場名1        = adoCol.Value
                Case "取引先工場名2":         If Not IsNull(adoCol.Value) Then .取引先工場名2        = adoCol.Value
                Case "納入先固定":            If Not IsNull(adoCol.Value) Then .納入先固定           = adoCol.Value
                Case "納入先":                If Not IsNull(adoCol.Value) Then .納入先               = adoCol.Value
                Case "納入先CD":              If Not IsNull(adoCol.Value) Then .納入先CD             = adoCol.Value
                Case "納入先名":              If Not IsNull(adoCol.Value) Then .納入先名             = adoCol.Value
                Case "納入先名1":             If Not IsNull(adoCol.Value) Then .納入先名1            = adoCol.Value
                Case "納入先名2":             If Not IsNull(adoCol.Value) Then .納入先名2            = adoCol.Value
                Case "納入先工場名":          If Not IsNull(adoCol.Value) Then .納入先工場名         = adoCol.Value
                Case "納入先受入":            If Not IsNull(adoCol.Value) Then .納入先受入           = adoCol.Value
                Case "メーカ固定":            If Not IsNull(adoCol.Value) Then .メーカ固定           = adoCol.Value
                Case "メーカ":                If Not IsNull(adoCol.Value) Then .メーカ               = adoCol.Value
                Case "メーカCD":              If Not IsNull(adoCol.Value) Then .メーカCD             = adoCol.Value
                Case "メーカ名":              If Not IsNull(adoCol.Value) Then .メーカ名             = adoCol.Value
                Case "メーカ名1":             If Not IsNull(adoCol.Value) Then .メーカ名1            = adoCol.Value
                Case "メーカ名2":             If Not IsNull(adoCol.Value) Then .メーカ名2            = adoCol.Value
                Case "メーカ受入":            If Not IsNull(adoCol.Value) Then .メーカ受入           = adoCol.Value
                Case "サイクル":              If Not IsNull(adoCol.Value) Then .サイクル             = adoCol.Value
                Case "置場":                  If Not IsNull(adoCol.Value) Then .置場                 = adoCol.Value
                Case "車型":                  If Not IsNull(adoCol.Value) Then .車型                 = adoCol.Value
                Case "自社CD":                If Not IsNull(adoCol.Value) Then .自社CD               = adoCol.Value
                Case "自社名1":               If Not IsNull(adoCol.Value) Then .自社名1              = adoCol.Value
                Case "自社名2":               If Not IsNull(adoCol.Value) Then .自社名2              = adoCol.Value
                Case "工程":                  If Not IsNull(adoCol.Value) Then .工程                 = adoCol.Value
                Case "工程名":                If Not IsNull(adoCol.Value) Then .工程名               = adoCol.Value
                Case "次工程":                If Not IsNull(adoCol.Value) Then .次工程               = adoCol.Value
                Case "次工程名":              If Not IsNull(adoCol.Value) Then .次工程名             = adoCol.Value
                Case "基準工数":              If Not IsNull(adoCol.Value) Then .基準工数             = adoCol.Value
                Case "生産ロット":            If Not IsNull(adoCol.Value) Then .生産ロット           = adoCol.Value
                Case "基準生産日":            If Not IsNull(adoCol.Value) Then .基準生産日           = adoCol.Value
                Case "管理番号":              If Not IsNull(adoCol.Value) Then .管理番号             = adoCol.Value
                Case "材質":                  If Not IsNull(adoCol.Value) Then .材質                 = adoCol.Value
                Case "材質18":                If Not IsNull(adoCol.Value) Then .材質18               = adoCol.Value
                Case "寸法区分":              If Not IsNull(adoCol.Value) Then .寸法区分             = adoCol.Value
                Case "板厚":                  If Not IsNull(adoCol.Value) Then .板厚                 = adoCol.Value
                Case "幅":                    If Not IsNull(adoCol.Value) Then .幅                   = adoCol.Value
                Case "長さ":                  If Not IsNull(adoCol.Value) Then .長さ                 = adoCol.Value
                Case "材料名":                If Not IsNull(adoCol.Value) Then .材料名               = adoCol.Value
                Case "使用品番":              If Not IsNull(adoCol.Value) Then .使用品番             = adoCol.Value
                Case "単位名":                If Not IsNull(adoCol.Value) Then .単位名               = adoCol.Value
                Case "手配単位名":            If Not IsNull(adoCol.Value) Then .手配単位名           = adoCol.Value
                Case "実数区分":              If Not IsNull(adoCol.Value) Then .実数区分             = adoCol.Value
                Case "材種コード":            If Not IsNull(adoCol.Value) Then .材種コード           = adoCol.Value
                Case "備考":                  If Not IsNull(adoCol.Value) Then .備考                 = adoCol.Value
                Case "StHNo":                 If Not IsNull(adoCol.Value) Then .StHNo                = adoCol.Value
                Case "EdHNo":                 If Not IsNull(adoCol.Value) Then .EdHNo                = adoCol.Value
                Case "Stack1":                If Not IsNull(adoCol.Value) Then .Stack1               = adoCol.Value
                Case "Stack2":                If Not IsNull(adoCol.Value) Then .Stack2               = adoCol.Value
                Case "区分":                  If Not IsNull(adoCol.Value) Then .区分                 = adoCol.Value
                Case "納入時間1":             If Not IsNull(adoCol.Value) Then .納入時間1            = adoCol.Value
                Case "納入時間2":             If Not IsNull(adoCol.Value) Then .納入時間2            = adoCol.Value
                Case "納入時間3":             If Not IsNull(adoCol.Value) Then .納入時間3            = adoCol.Value
                Case "納入時間4":             If Not IsNull(adoCol.Value) Then .納入時間4            = adoCol.Value
                Case "納入時間5":             If Not IsNull(adoCol.Value) Then .納入時間5            = adoCol.Value
                Case "納入時間6":             If Not IsNull(adoCol.Value) Then .納入時間6            = adoCol.Value
                Case "納入時間7":             If Not IsNull(adoCol.Value) Then .納入時間7            = adoCol.Value
                Case "納入時間8":             If Not IsNull(adoCol.Value) Then .納入時間8            = adoCol.Value
                Case "得意先背番1":           If Not IsNull(adoCol.Value) Then .得意先背番1          = adoCol.Value
                Case "得意先背番2":           If Not IsNull(adoCol.Value) Then .得意先背番2          = adoCol.Value
                Case "得意先背番3":           If Not IsNull(adoCol.Value) Then .得意先背番3          = adoCol.Value
                Case "得意先背番4":           If Not IsNull(adoCol.Value) Then .得意先背番4          = adoCol.Value
                Case "得意先工場名1":         If Not IsNull(adoCol.Value) Then .得意先工場名1        = adoCol.Value
                Case "得意先工場名2":         If Not IsNull(adoCol.Value) Then .得意先工場名2        = adoCol.Value
                Case "得意先工場名3":         If Not IsNull(adoCol.Value) Then .得意先工場名3        = adoCol.Value
                Case "得意先工場名4":         If Not IsNull(adoCol.Value) Then .得意先工場名4        = adoCol.Value
                Case "部品区分":              If Not IsNull(adoCol.Value) Then .部品区分             = adoCol.Value
                Case "ポイント":              If Not IsNull(adoCol.Value) Then .ポイント             = adoCol.Value
                Case "ポイントIMAGE":         If Not IsNull(adoCol.Value) Then .ポイントIMAGE        = adoCol.Value
                Case "部署":                  If Not IsNull(adoCol.Value) Then .部署                 = adoCol.Value
                Case "部署名":                If Not IsNull(adoCol.Value) Then .部署名               = adoCol.Value
                Case "返却場所":              If Not IsNull(adoCol.Value) Then .返却場所             = adoCol.Value
                Case "返却場所名":            If Not IsNull(adoCol.Value) Then .返却場所名           = adoCol.Value
                Case "後工程":                If Not IsNull(adoCol.Value) Then .後工程               = adoCol.Value
                Case "後工程名1":             If Not IsNull(adoCol.Value) Then .後工程名1            = adoCol.Value
                Case "後工程名2":             If Not IsNull(adoCol.Value) Then .後工程名2            = adoCol.Value
                Case "得意先":                If Not IsNull(adoCol.Value) Then .得意先               = adoCol.Value
                Case "得意先名":              If Not IsNull(adoCol.Value) Then .得意先名             = adoCol.Value
                Case "生産区分":              If Not IsNull(adoCol.Value) Then .生産区分             = adoCol.Value
                Case "かんばん枠色":          If Not IsNull(adoCol.Value) Then .かんばん枠色         = adoCol.Value
                Case "発行済":                If Not IsNull(adoCol.Value) Then .発行済               = adoCol.Value
                Case "発行FLG":               If Not IsNull(adoCol.Value) Then .発行FLG              = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
                Case "作成者":                If Not IsNull(adoCol.Value) Then .作成者               = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "最終発行連番":	If Not IsNull(adoCol.Value) Then .最終発行連番	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KHTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KHTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From かんばん発行テーブル"
        RTxtSql = RTxtSql &" Where かんばん種類 = " & KHT.かんばん種類 & ""
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(KHT.品番) & "'"
        RTxtSql = RTxtSql &" And   取引先 = '" & RTrim(KHT.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(KHT.受入) & "'"
        RTxtSql = RTxtSql &" And   発行NO = " & KHT.発行NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KHTRead = True
           Call ItemsFromKHT(DBNo)
	   KHT_RDSTS = True
        else
	   KHT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KHTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KHTReadNext(Optional DBNo As Integer = 0)
    If Not KHT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KHT_RDSTS = False
        Else
           Call ItemsFromKHT(DBNo)
           KHT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KHTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KHT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KHT )
'*----------------------------------------------*'
Public Sub ItemsClearKHT()
    With KHT
	.かんばん種類 = 0
	.品番 = ""
	.取引先 = ""
	.受入 = ""
	.発行NO = 0
	.発行種類 = 0
	.材料管理番号 = 0
	.背番号 = ""
	.版数 = 0
	.発行区分 = 0
	.表示品番 = ""
	.品番14 = ""
	.品番30 = ""
	.品目 = 0
	.品名 = ""
	.受入名 = ""
	.収容器 = ""
	.収容数 = 0
	.収容区分 = 0
	.手配区分 = ""
	.手配区分名 = ""
	.アドレス = ""
	.IMAGE = ""
	.取引先CD = ""
	.取引先名 = ""
	.取引先名1 = ""
	.取引先名2 = ""
	.取引先工場名 = ""
	.取引先工場名1 = ""
	.取引先工場名2 = ""
	.納入先固定 = ""
	.納入先 = ""
	.納入先CD = ""
	.納入先名 = ""
	.納入先名1 = ""
	.納入先名2 = ""
	.納入先工場名 = ""
	.納入先受入 = ""
	.メーカ固定 = ""
	.メーカ = ""
	.メーカCD = ""
	.メーカ名 = ""
	.メーカ名1 = ""
	.メーカ名2 = ""
	.メーカ受入 = ""
	.サイクル = ""
	.置場 = ""
	.車型 = ""
	.自社CD = ""
	.自社名1 = ""
	.自社名2 = ""
	.工程 = ""
	.工程名 = ""
	.次工程 = ""
	.次工程名 = ""
	.基準工数 = 0
	.生産ロット = 0
	.基準生産日 = ""
	.管理番号 = ""
	.材質 = ""
	.材質18 = ""
	.寸法区分 = 0
	.板厚 = 0
	.幅 = 0
	.長さ = 0
	.材料名 = ""
	.使用品番 = ""
	.単位名 = ""
	.手配単位名 = ""
	.実数区分 = 0
	.材種コード = ""
	.備考 = ""
	.StHNo = 0
	.EdHNo = 0
	.Stack1 = 0
	.Stack2 = 0
	.区分 = ""
	.納入時間1 = ""
	.納入時間2 = ""
	.納入時間3 = ""
	.納入時間4 = ""
	.納入時間5 = ""
	.納入時間6 = ""
	.納入時間7 = ""
	.納入時間8 = ""
	.得意先背番1 = ""
	.得意先背番2 = ""
	.得意先背番3 = ""
	.得意先背番4 = ""
	.得意先工場名1 = ""
	.得意先工場名2 = ""
	.得意先工場名3 = ""
	.得意先工場名4 = ""
	.部品区分 = 0
	.ポイント = ""
	.ポイントIMAGE = ""
	.部署 = ""
	.部署名 = ""
	.返却場所 = ""
	.返却場所名 = ""
	.後工程 = ""
	.後工程名1 = ""
	.後工程名2 = ""
	.得意先 = ""
	.得意先名 = ""
	.生産区分 = ""
	.かんばん枠色 = 0
	.発行済 = 0
	.発行FLG = ""
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
        .件数 = 0
'   項目追加(自動)
	.最終発行連番 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KHTInsert(Optional DBNo As Integer) As Boolean
    KHTInsert = False
    With KHT
        gSL_Select = "Insert Into かんばん発行テーブル  ( "
        gSL_Select = gSL_Select & "かんばん種類,品番,取引先,"
        gSL_Select = gSL_Select & "受入,発行NO,発行種類,"
        gSL_Select = gSL_Select & "材料管理番号,背番号,版数,"
        gSL_Select = gSL_Select & "発行区分,表示品番,品番14,"
        gSL_Select = gSL_Select & "品番30,品目,品名,"
        gSL_Select = gSL_Select & "受入名,収容器,収容数,"
        gSL_Select = gSL_Select & "収容区分,手配区分,手配区分名,"
        gSL_Select = gSL_Select & "アドレス,IMAGE,取引先CD,"
        gSL_Select = gSL_Select & "取引先名,取引先名1,取引先名2,"
        gSL_Select = gSL_Select & "取引先工場名,取引先工場名1,取引先工場名2,"
        gSL_Select = gSL_Select & "納入先固定,納入先,納入先CD,"
        gSL_Select = gSL_Select & "納入先名,納入先名1,納入先名2,"
        gSL_Select = gSL_Select & "納入先工場名,納入先受入,メーカ固定,"
        gSL_Select = gSL_Select & "メーカ,メーカCD,メーカ名,"
        gSL_Select = gSL_Select & "メーカ名1,メーカ名2,メーカ受入,"
        gSL_Select = gSL_Select & "サイクル,置場,車型,"
        gSL_Select = gSL_Select & "自社CD,自社名1,自社名2,"
        gSL_Select = gSL_Select & "工程,工程名,次工程,"
        gSL_Select = gSL_Select & "次工程名,基準工数,生産ロット,"
        gSL_Select = gSL_Select & "基準生産日,管理番号,材質,"
        gSL_Select = gSL_Select & "材質18,寸法区分,板厚,"
        gSL_Select = gSL_Select & "幅,長さ,材料名,"
        gSL_Select = gSL_Select & "使用品番,単位名,手配単位名,"
        gSL_Select = gSL_Select & "実数区分,材種コード,備考,"
        gSL_Select = gSL_Select & "StHNo,EdHNo,Stack1,"
        gSL_Select = gSL_Select & "Stack2,区分,納入時間1,"
        gSL_Select = gSL_Select & "納入時間2,納入時間3,納入時間4,"
        gSL_Select = gSL_Select & "納入時間5,納入時間6,納入時間7,"
        gSL_Select = gSL_Select & "納入時間8,得意先背番1,得意先背番2,"
        gSL_Select = gSL_Select & "得意先背番3,得意先背番4,得意先工場名1,"
        gSL_Select = gSL_Select & "得意先工場名2,得意先工場名3,得意先工場名4,"
        gSL_Select = gSL_Select & "部品区分,ポイント,ポイントIMAGE,"
        gSL_Select = gSL_Select & "部署,部署名,返却場所,"
        gSL_Select = gSL_Select & "返却場所名,後工程,後工程名1,"
        gSL_Select = gSL_Select & "後工程名2,得意先,得意先名,"
        gSL_Select = gSL_Select & "生産区分,かんばん枠色,発行済,"
        gSL_Select = gSL_Select & "発行FLG,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .かんばん種類 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "" & .発行NO & ","		
        gSL_Select = gSL_Select & "" & .発行種類 & ","		
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "" & .版数 & ","		
        gSL_Select = gSL_Select & "" & .発行区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.表示品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番14) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番30) & "',"		
        gSL_Select = gSL_Select & "" & .品目 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.品名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.収容器) & "',"		
        gSL_Select = gSL_Select & "" & .収容数 & ","		
        gSL_Select = gSL_Select & "" & .収容区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.手配区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配区分名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.アドレス) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.IMAGE) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先CD) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先名1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先名2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先工場名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先工場名1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先工場名2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先固定) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先CD) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先名1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先名2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先工場名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.メーカ固定) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.メーカ) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.メーカCD) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.メーカ名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.メーカ名1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.メーカ名2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.メーカ受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.サイクル) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.置場) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.車型) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.自社CD) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.自社名1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.自社名2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.工程) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.工程名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.次工程) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.次工程名) & "',"		
        gSL_Select = gSL_Select & "" & .基準工数 & ","		
        gSL_Select = gSL_Select & "" & .生産ロット & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.基準生産日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.管理番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.材質18) & "',"		
        gSL_Select = gSL_Select & "" & .寸法区分 & ","		
        gSL_Select = gSL_Select & "" & .板厚 & ","		
        gSL_Select = gSL_Select & "" & .幅 & ","		
        gSL_Select = gSL_Select & "" & .長さ & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.材料名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.使用品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.単位名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配単位名) & "',"		
        gSL_Select = gSL_Select & "" & .実数区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.材種コード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.備考) & "',"		
        gSL_Select = gSL_Select & "" & .StHNo & ","		
        gSL_Select = gSL_Select & "" & .EdHNo & ","		
        gSL_Select = gSL_Select & "" & .Stack1 & ","		
        gSL_Select = gSL_Select & "" & .Stack2 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入時間1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入時間2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入時間3) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入時間4) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入時間5) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入時間6) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入時間7) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入時間8) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.得意先背番1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.得意先背番2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.得意先背番3) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.得意先背番4) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.得意先工場名1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.得意先工場名2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.得意先工場名3) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.得意先工場名4) & "',"		
        gSL_Select = gSL_Select & "" & .部品区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.ポイント) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ポイントIMAGE) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.部署) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.部署名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.返却場所) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.返却場所名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.後工程) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.後工程名1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.後工程名2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.得意先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.得意先名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.生産区分) & "',"		
        gSL_Select = gSL_Select & "" & .かんばん枠色 & ","		
        gSL_Select = gSL_Select & "" & .発行済 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.発行FLG) & "',"		
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
    KHTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KHTUpdate(Optional DBNo As Integer) As Boolean
    KHTUpdate = False	
    With KHT
        gSL_Select = "Update かんばん発行テーブル  Set "
        gSL_Select = gSL_Select & "かんばん種類 = " & .かんばん種類 & ","
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "発行NO = " & .発行NO & ","
        gSL_Select = gSL_Select & "発行種類 = " & .発行種類 & ","
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "版数 = " & .版数 & ","
        gSL_Select = gSL_Select & "発行区分 = " & .発行区分 & ","
        gSL_Select = gSL_Select & "表示品番 = '" & RTrim(.表示品番) & "',"
        gSL_Select = gSL_Select & "品番14 = '" & RTrim(.品番14) & "',"
        gSL_Select = gSL_Select & "品番30 = '" & RTrim(.品番30) & "',"
        gSL_Select = gSL_Select & "品目 = " & .品目 & ","
        gSL_Select = gSL_Select & "品名 = '" & RTrim(.品名) & "',"
        gSL_Select = gSL_Select & "受入名 = '" & RTrim(.受入名) & "',"
        gSL_Select = gSL_Select & "収容器 = '" & RTrim(.収容器) & "',"
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "収容区分 = " & .収容区分 & ","
        gSL_Select = gSL_Select & "手配区分 = '" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "手配区分名 = '" & RTrim(.手配区分名) & "',"
        gSL_Select = gSL_Select & "アドレス = '" & RTrim(.アドレス) & "',"
        gSL_Select = gSL_Select & "IMAGE = '" & RTrim(.IMAGE) & "',"
        gSL_Select = gSL_Select & "取引先CD = '" & RTrim(.取引先CD) & "',"
        gSL_Select = gSL_Select & "取引先名 = '" & RTrim(.取引先名) & "',"
        gSL_Select = gSL_Select & "取引先名1 = '" & RTrim(.取引先名1) & "',"
        gSL_Select = gSL_Select & "取引先名2 = '" & RTrim(.取引先名2) & "',"
        gSL_Select = gSL_Select & "取引先工場名 = '" & RTrim(.取引先工場名) & "',"
        gSL_Select = gSL_Select & "取引先工場名1 = '" & RTrim(.取引先工場名1) & "',"
        gSL_Select = gSL_Select & "取引先工場名2 = '" & RTrim(.取引先工場名2) & "',"
        gSL_Select = gSL_Select & "納入先固定 = '" & RTrim(.納入先固定) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "納入先CD = '" & RTrim(.納入先CD) & "',"
        gSL_Select = gSL_Select & "納入先名 = '" & RTrim(.納入先名) & "',"
        gSL_Select = gSL_Select & "納入先名1 = '" & RTrim(.納入先名1) & "',"
        gSL_Select = gSL_Select & "納入先名2 = '" & RTrim(.納入先名2) & "',"
        gSL_Select = gSL_Select & "納入先工場名 = '" & RTrim(.納入先工場名) & "',"
        gSL_Select = gSL_Select & "納入先受入 = '" & RTrim(.納入先受入) & "',"
        gSL_Select = gSL_Select & "メーカ固定 = '" & RTrim(.メーカ固定) & "',"
        gSL_Select = gSL_Select & "メーカ = '" & RTrim(.メーカ) & "',"
        gSL_Select = gSL_Select & "メーカCD = '" & RTrim(.メーカCD) & "',"
        gSL_Select = gSL_Select & "メーカ名 = '" & RTrim(.メーカ名) & "',"
        gSL_Select = gSL_Select & "メーカ名1 = '" & RTrim(.メーカ名1) & "',"
        gSL_Select = gSL_Select & "メーカ名2 = '" & RTrim(.メーカ名2) & "',"
        gSL_Select = gSL_Select & "メーカ受入 = '" & RTrim(.メーカ受入) & "',"
        gSL_Select = gSL_Select & "サイクル = '" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "置場 = '" & RTrim(.置場) & "',"
        gSL_Select = gSL_Select & "車型 = '" & RTrim(.車型) & "',"
        gSL_Select = gSL_Select & "自社CD = '" & RTrim(.自社CD) & "',"
        gSL_Select = gSL_Select & "自社名1 = '" & RTrim(.自社名1) & "',"
        gSL_Select = gSL_Select & "自社名2 = '" & RTrim(.自社名2) & "',"
        gSL_Select = gSL_Select & "工程 = '" & RTrim(.工程) & "',"
        gSL_Select = gSL_Select & "工程名 = '" & RTrim(.工程名) & "',"
        gSL_Select = gSL_Select & "次工程 = '" & RTrim(.次工程) & "',"
        gSL_Select = gSL_Select & "次工程名 = '" & RTrim(.次工程名) & "',"
        gSL_Select = gSL_Select & "基準工数 = " & .基準工数 & ","
        gSL_Select = gSL_Select & "生産ロット = " & .生産ロット & ","
        gSL_Select = gSL_Select & "基準生産日 = '" & RTrim(.基準生産日) & "',"
        gSL_Select = gSL_Select & "管理番号 = '" & RTrim(.管理番号) & "',"
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "材質18 = '" & RTrim(.材質18) & "',"
        gSL_Select = gSL_Select & "寸法区分 = " & .寸法区分 & ","
        gSL_Select = gSL_Select & "板厚 = " & .板厚 & ","
        gSL_Select = gSL_Select & "幅 = " & .幅 & ","
        gSL_Select = gSL_Select & "長さ = " & .長さ & ","
        gSL_Select = gSL_Select & "材料名 = '" & RTrim(.材料名) & "',"
        gSL_Select = gSL_Select & "使用品番 = '" & RTrim(.使用品番) & "',"
        gSL_Select = gSL_Select & "単位名 = '" & RTrim(.単位名) & "',"
        gSL_Select = gSL_Select & "手配単位名 = '" & RTrim(.手配単位名) & "',"
        gSL_Select = gSL_Select & "実数区分 = " & .実数区分 & ","
        gSL_Select = gSL_Select & "材種コード = '" & RTrim(.材種コード) & "',"
        gSL_Select = gSL_Select & "備考 = '" & RTrim(.備考) & "',"
        gSL_Select = gSL_Select & "StHNo = " & .StHNo & ","
        gSL_Select = gSL_Select & "EdHNo = " & .EdHNo & ","
        gSL_Select = gSL_Select & "Stack1 = " & .Stack1 & ","
        gSL_Select = gSL_Select & "Stack2 = " & .Stack2 & ","
        gSL_Select = gSL_Select & "区分 = '" & RTrim(.区分) & "',"
        gSL_Select = gSL_Select & "納入時間1 = '" & RTrim(.納入時間1) & "',"
        gSL_Select = gSL_Select & "納入時間2 = '" & RTrim(.納入時間2) & "',"
        gSL_Select = gSL_Select & "納入時間3 = '" & RTrim(.納入時間3) & "',"
        gSL_Select = gSL_Select & "納入時間4 = '" & RTrim(.納入時間4) & "',"
        gSL_Select = gSL_Select & "納入時間5 = '" & RTrim(.納入時間5) & "',"
        gSL_Select = gSL_Select & "納入時間6 = '" & RTrim(.納入時間6) & "',"
        gSL_Select = gSL_Select & "納入時間7 = '" & RTrim(.納入時間7) & "',"
        gSL_Select = gSL_Select & "納入時間8 = '" & RTrim(.納入時間8) & "',"
        gSL_Select = gSL_Select & "得意先背番1 = '" & RTrim(.得意先背番1) & "',"
        gSL_Select = gSL_Select & "得意先背番2 = '" & RTrim(.得意先背番2) & "',"
        gSL_Select = gSL_Select & "得意先背番3 = '" & RTrim(.得意先背番3) & "',"
        gSL_Select = gSL_Select & "得意先背番4 = '" & RTrim(.得意先背番4) & "',"
        gSL_Select = gSL_Select & "得意先工場名1 = '" & RTrim(.得意先工場名1) & "',"
        gSL_Select = gSL_Select & "得意先工場名2 = '" & RTrim(.得意先工場名2) & "',"
        gSL_Select = gSL_Select & "得意先工場名3 = '" & RTrim(.得意先工場名3) & "',"
        gSL_Select = gSL_Select & "得意先工場名4 = '" & RTrim(.得意先工場名4) & "',"
        gSL_Select = gSL_Select & "部品区分 = " & .部品区分 & ","
        gSL_Select = gSL_Select & "ポイント = '" & RTrim(.ポイント) & "',"
        gSL_Select = gSL_Select & "ポイントIMAGE = '" & RTrim(.ポイントIMAGE) & "',"
        gSL_Select = gSL_Select & "部署 = '" & RTrim(.部署) & "',"
        gSL_Select = gSL_Select & "部署名 = '" & RTrim(.部署名) & "',"
        gSL_Select = gSL_Select & "返却場所 = '" & RTrim(.返却場所) & "',"
        gSL_Select = gSL_Select & "返却場所名 = '" & RTrim(.返却場所名) & "',"
        gSL_Select = gSL_Select & "後工程 = '" & RTrim(.後工程) & "',"
        gSL_Select = gSL_Select & "後工程名1 = '" & RTrim(.後工程名1) & "',"
        gSL_Select = gSL_Select & "後工程名2 = '" & RTrim(.後工程名2) & "',"
        gSL_Select = gSL_Select & "得意先 = '" & RTrim(.得意先) & "',"
        gSL_Select = gSL_Select & "得意先名 = '" & RTrim(.得意先名) & "',"
        gSL_Select = gSL_Select & "生産区分 = '" & RTrim(.生産区分) & "',"
        gSL_Select = gSL_Select & "かんばん枠色 = " & .かんばん枠色 & ","
        gSL_Select = gSL_Select & "発行済 = " & .発行済 & ","
        gSL_Select = gSL_Select & "発行FLG = '" & RTrim(.発行FLG) & "'"
        gSL_Select = gSL_Select & " Where かんばん種類 = " & .かんばん種類 & ""
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 発行NO = " & .発行NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KHTUpdate = True
End Function

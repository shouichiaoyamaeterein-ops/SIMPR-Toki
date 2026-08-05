Attribute VB_Name = "K_品目マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   K_品目マスタ
'*
'*   モジュールID  : K_品目マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2009/08/26  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_K_品目マスタ
        品目コード                As String
        品目区分                  As String
        品目分類                  As String
        品目名                    As String
        仕様                      As String
        完成品納入先コード        As String
        製作場所                  As String
        エフ発行                  As String
        実質在庫管理区分          As String
        ダミー                    As String
        設変通知NO                As String
        有効開始期間              As String
        有効終了期間              As String
        ローレベルコード          As String
        手配先                    As String
        リードタイム正味          As String
        リードタイム安全          As String
        リードタイム累計          As String
        ロット方式                As String
        ロット決定値1             As String
        ロット決定値2             As String
        ロット丸め区分            As String
        発注単位                  As String
        納入場所                  As String
        注文書発行区分            As String
        棚卸資材区分              As String
        在庫単位                  As String
        在庫管理区分              As String
        在庫引当区分              As String
        安全在庫決定区分          As String
        安全在庫                  As String
        棚卸評価単位              As String
        出荷エフ色                As String
        標準SNP                   As String
        非標準納入場所            As String
        非標準SNP                 As String
        売上                      As String
        原価                      As String
        前期末棚卸評価単位        As String
        車種                      As String
        加工費                    As String
        材料費                    As String
        購入品費                  As String
        型費                      As String
        管理費                    As String
        ETC1                      As String
        ETC2                      As String
        ETC3                      As String
        件数          As Long
    End Type

    Public KKH As T_K_品目マスタ

    Public KKH_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKKH(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKKH	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KKH
            Select Case adoCol.Name
                Case "品目コード":                If Not IsNull(adoCol.Value) Then .品目コード               = adoCol.Value
                Case "品目区分":                  If Not IsNull(adoCol.Value) Then .品目区分                 = adoCol.Value
                Case "品目分類":                  If Not IsNull(adoCol.Value) Then .品目分類                 = adoCol.Value
                Case "品目名":                    If Not IsNull(adoCol.Value) Then .品目名                   = adoCol.Value
                Case "仕様":                      If Not IsNull(adoCol.Value) Then .仕様                     = adoCol.Value
                Case "完成品納入先コード":        If Not IsNull(adoCol.Value) Then .完成品納入先コード       = adoCol.Value
                Case "製作場所":                  If Not IsNull(adoCol.Value) Then .製作場所                 = adoCol.Value
                Case "エフ発行":                  If Not IsNull(adoCol.Value) Then .エフ発行                 = adoCol.Value
                Case "実質在庫管理区分":          If Not IsNull(adoCol.Value) Then .実質在庫管理区分         = adoCol.Value
                Case "ダミー":                    If Not IsNull(adoCol.Value) Then .ダミー                   = adoCol.Value
                Case "設変通知NO":                If Not IsNull(adoCol.Value) Then .設変通知NO               = adoCol.Value
                Case "有効開始期間":              If Not IsNull(adoCol.Value) Then .有効開始期間             = adoCol.Value
                Case "有効終了期間":              If Not IsNull(adoCol.Value) Then .有効終了期間             = adoCol.Value
                Case "ローレベルコード":          If Not IsNull(adoCol.Value) Then .ローレベルコード         = adoCol.Value
                Case "手配先":                    If Not IsNull(adoCol.Value) Then .手配先                   = adoCol.Value
                Case "リードタイム正味":          If Not IsNull(adoCol.Value) Then .リードタイム正味         = adoCol.Value
                Case "リードタイム安全":          If Not IsNull(adoCol.Value) Then .リードタイム安全         = adoCol.Value
                Case "リードタイム累計":          If Not IsNull(adoCol.Value) Then .リードタイム累計         = adoCol.Value
                Case "ロット方式":                If Not IsNull(adoCol.Value) Then .ロット方式               = adoCol.Value
                Case "ロット決定値1":             If Not IsNull(adoCol.Value) Then .ロット決定値1            = adoCol.Value
                Case "ロット決定値2":             If Not IsNull(adoCol.Value) Then .ロット決定値2            = adoCol.Value
                Case "ロット丸め区分":            If Not IsNull(adoCol.Value) Then .ロット丸め区分           = adoCol.Value
                Case "発注単位":                  If Not IsNull(adoCol.Value) Then .発注単位                 = adoCol.Value
                Case "納入場所":                  If Not IsNull(adoCol.Value) Then .納入場所                 = adoCol.Value
                Case "注文書発行区分":            If Not IsNull(adoCol.Value) Then .注文書発行区分           = adoCol.Value
                Case "棚卸資材区分":              If Not IsNull(adoCol.Value) Then .棚卸資材区分             = adoCol.Value
                Case "在庫単位":                  If Not IsNull(adoCol.Value) Then .在庫単位                 = adoCol.Value
                Case "在庫管理区分":              If Not IsNull(adoCol.Value) Then .在庫管理区分             = adoCol.Value
                Case "在庫引当区分":              If Not IsNull(adoCol.Value) Then .在庫引当区分             = adoCol.Value
                Case "安全在庫決定区分":          If Not IsNull(adoCol.Value) Then .安全在庫決定区分         = adoCol.Value
                Case "安全在庫":                  If Not IsNull(adoCol.Value) Then .安全在庫                 = adoCol.Value
                Case "棚卸評価単位":              If Not IsNull(adoCol.Value) Then .棚卸評価単位             = adoCol.Value
                Case "出荷エフ色":                If Not IsNull(adoCol.Value) Then .出荷エフ色               = adoCol.Value
                Case "標準SNP":                   If Not IsNull(adoCol.Value) Then .標準SNP                  = adoCol.Value
                Case "非標準納入場所":            If Not IsNull(adoCol.Value) Then .非標準納入場所           = adoCol.Value
                Case "非標準SNP":                 If Not IsNull(adoCol.Value) Then .非標準SNP                = adoCol.Value
                Case "売上":                      If Not IsNull(adoCol.Value) Then .売上                     = adoCol.Value
                Case "原価":                      If Not IsNull(adoCol.Value) Then .原価                     = adoCol.Value
                Case "前期末棚卸評価単位":        If Not IsNull(adoCol.Value) Then .前期末棚卸評価単位       = adoCol.Value
                Case "車種":                      If Not IsNull(adoCol.Value) Then .車種                     = adoCol.Value
                Case "加工費":                    If Not IsNull(adoCol.Value) Then .加工費                   = adoCol.Value
                Case "材料費":                    If Not IsNull(adoCol.Value) Then .材料費                   = adoCol.Value
                Case "購入品費":                  If Not IsNull(adoCol.Value) Then .購入品費                 = adoCol.Value
                Case "型費":                      If Not IsNull(adoCol.Value) Then .型費                     = adoCol.Value
                Case "管理費":                    If Not IsNull(adoCol.Value) Then .管理費                   = adoCol.Value
                Case "ETC1":                      If Not IsNull(adoCol.Value) Then .ETC1                     = adoCol.Value
                Case "ETC2":                      If Not IsNull(adoCol.Value) Then .ETC2                     = adoCol.Value
                Case "ETC3":                      If Not IsNull(adoCol.Value) Then .ETC3                     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KKHRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KKHRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From K_品目マスタ"
        RTxtSql = RTxtSql &" Where 品目コード = '" & RTrim(KKH.品目コード) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KKHRead = True
           Call ItemsFromKKH(DBNo)
	   KKH_RDSTS = True
        else
	   KKH_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KKHClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KKHReadNext(Optional DBNo As Integer = 0)
    If Not KKH_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KKH_RDSTS = False
        Else
           Call ItemsFromKKH(DBNo)
           KKH_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KKHClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KKH_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KKH )
'*----------------------------------------------*'
Public Sub ItemsClearKKH()
    With KKH
	.品目コード = ""
	.品目区分 = ""
	.品目分類 = ""
	.品目名 = ""
	.仕様 = ""
	.完成品納入先コード = ""
	.製作場所 = ""
	.エフ発行 = ""
	.実質在庫管理区分 = ""
	.ダミー = ""
	.設変通知NO = ""
	.有効開始期間 = ""
	.有効終了期間 = ""
	.ローレベルコード = ""
	.手配先 = ""
	.リードタイム正味 = ""
	.リードタイム安全 = ""
	.リードタイム累計 = ""
	.ロット方式 = ""
	.ロット決定値1 = ""
	.ロット決定値2 = ""
	.ロット丸め区分 = ""
	.発注単位 = ""
	.納入場所 = ""
	.注文書発行区分 = ""
	.棚卸資材区分 = ""
	.在庫単位 = ""
	.在庫管理区分 = ""
	.在庫引当区分 = ""
	.安全在庫決定区分 = ""
	.安全在庫 = ""
	.棚卸評価単位 = ""
	.出荷エフ色 = ""
	.標準SNP = ""
	.非標準納入場所 = ""
	.非標準SNP = ""
	.売上 = ""
	.原価 = ""
	.前期末棚卸評価単位 = ""
	.車種 = ""
	.加工費 = ""
	.材料費 = ""
	.購入品費 = ""
	.型費 = ""
	.管理費 = ""
	.ETC1 = ""
	.ETC2 = ""
	.ETC3 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KKHInsert(Optional DBNo As Integer) As Boolean
    KKHInsert = False
    With KKH
        gSL_Select = "Insert Into K_品目マスタ  ( "
        gSL_Select = gSL_Select & "品目コード,品目区分,品目分類,"
        gSL_Select = gSL_Select & "品目名,仕様,完成品納入先コード,"
        gSL_Select = gSL_Select & "製作場所,エフ発行,実質在庫管理区分,"
        gSL_Select = gSL_Select & "ダミー,設変通知NO,有効開始期間,"
        gSL_Select = gSL_Select & "有効終了期間,ローレベルコード,手配先,"
        gSL_Select = gSL_Select & "リードタイム正味,リードタイム安全,リードタイム累計,"
        gSL_Select = gSL_Select & "ロット方式,ロット決定値1,ロット決定値2,"
        gSL_Select = gSL_Select & "ロット丸め区分,発注単位,納入場所,"
        gSL_Select = gSL_Select & "注文書発行区分,棚卸資材区分,在庫単位,"
        gSL_Select = gSL_Select & "在庫管理区分,在庫引当区分,安全在庫決定区分,"
        gSL_Select = gSL_Select & "安全在庫,棚卸評価単位,出荷エフ色,"
        gSL_Select = gSL_Select & "標準SNP,非標準納入場所,非標準SNP,"
        gSL_Select = gSL_Select & "売上,原価,前期末棚卸評価単位,"
        gSL_Select = gSL_Select & "車種,加工費,材料費,"
        gSL_Select = gSL_Select & "購入品費,型費,管理費,"
        gSL_Select = gSL_Select & "ETC1,ETC2,ETC3"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.品目コード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品目区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品目分類) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品目名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕様) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.完成品納入先コード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.製作場所) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.エフ発行) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.実質在庫管理区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ダミー) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.設変通知NO) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.有効開始期間) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.有効終了期間) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ローレベルコード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.リードタイム正味) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.リードタイム安全) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.リードタイム累計) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ロット方式) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ロット決定値1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ロット決定値2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ロット丸め区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.発注単位) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入場所) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.注文書発行区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.棚卸資材区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.在庫単位) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.在庫管理区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.在庫引当区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.安全在庫決定区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.安全在庫) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.棚卸評価単位) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.出荷エフ色) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.標準SNP) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.非標準納入場所) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.非標準SNP) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.売上) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.原価) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.前期末棚卸評価単位) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.車種) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.加工費) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.材料費) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.購入品費) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.型費) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.管理費) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ETC1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ETC2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ETC3) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KKHInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KKHUpdate(Optional DBNo As Integer) As Boolean
    KKHUpdate = False	
    With KKH
        gSL_Select = "Update K_品目マスタ  Set "
        gSL_Select = gSL_Select & "品目コード = '" & RTrim(.品目コード) & "',"
        gSL_Select = gSL_Select & "品目区分 = '" & RTrim(.品目区分) & "',"
        gSL_Select = gSL_Select & "品目分類 = '" & RTrim(.品目分類) & "',"
        gSL_Select = gSL_Select & "品目名 = '" & RTrim(.品目名) & "',"
        gSL_Select = gSL_Select & "仕様 = '" & RTrim(.仕様) & "',"
        gSL_Select = gSL_Select & "完成品納入先コード = '" & RTrim(.完成品納入先コード) & "',"
        gSL_Select = gSL_Select & "製作場所 = '" & RTrim(.製作場所) & "',"
        gSL_Select = gSL_Select & "エフ発行 = '" & RTrim(.エフ発行) & "',"
        gSL_Select = gSL_Select & "実質在庫管理区分 = '" & RTrim(.実質在庫管理区分) & "',"
        gSL_Select = gSL_Select & "ダミー = '" & RTrim(.ダミー) & "',"
        gSL_Select = gSL_Select & "設変通知NO = '" & RTrim(.設変通知NO) & "',"
        gSL_Select = gSL_Select & "有効開始期間 = '" & RTrim(.有効開始期間) & "',"
        gSL_Select = gSL_Select & "有効終了期間 = '" & RTrim(.有効終了期間) & "',"
        gSL_Select = gSL_Select & "ローレベルコード = '" & RTrim(.ローレベルコード) & "',"
        gSL_Select = gSL_Select & "手配先 = '" & RTrim(.手配先) & "',"
        gSL_Select = gSL_Select & "リードタイム正味 = '" & RTrim(.リードタイム正味) & "',"
        gSL_Select = gSL_Select & "リードタイム安全 = '" & RTrim(.リードタイム安全) & "',"
        gSL_Select = gSL_Select & "リードタイム累計 = '" & RTrim(.リードタイム累計) & "',"
        gSL_Select = gSL_Select & "ロット方式 = '" & RTrim(.ロット方式) & "',"
        gSL_Select = gSL_Select & "ロット決定値1 = '" & RTrim(.ロット決定値1) & "',"
        gSL_Select = gSL_Select & "ロット決定値2 = '" & RTrim(.ロット決定値2) & "',"
        gSL_Select = gSL_Select & "ロット丸め区分 = '" & RTrim(.ロット丸め区分) & "',"
        gSL_Select = gSL_Select & "発注単位 = '" & RTrim(.発注単位) & "',"
        gSL_Select = gSL_Select & "納入場所 = '" & RTrim(.納入場所) & "',"
        gSL_Select = gSL_Select & "注文書発行区分 = '" & RTrim(.注文書発行区分) & "',"
        gSL_Select = gSL_Select & "棚卸資材区分 = '" & RTrim(.棚卸資材区分) & "',"
        gSL_Select = gSL_Select & "在庫単位 = '" & RTrim(.在庫単位) & "',"
        gSL_Select = gSL_Select & "在庫管理区分 = '" & RTrim(.在庫管理区分) & "',"
        gSL_Select = gSL_Select & "在庫引当区分 = '" & RTrim(.在庫引当区分) & "',"
        gSL_Select = gSL_Select & "安全在庫決定区分 = '" & RTrim(.安全在庫決定区分) & "',"
        gSL_Select = gSL_Select & "安全在庫 = '" & RTrim(.安全在庫) & "',"
        gSL_Select = gSL_Select & "棚卸評価単位 = '" & RTrim(.棚卸評価単位) & "',"
        gSL_Select = gSL_Select & "出荷エフ色 = '" & RTrim(.出荷エフ色) & "',"
        gSL_Select = gSL_Select & "標準SNP = '" & RTrim(.標準SNP) & "',"
        gSL_Select = gSL_Select & "非標準納入場所 = '" & RTrim(.非標準納入場所) & "',"
        gSL_Select = gSL_Select & "非標準SNP = '" & RTrim(.非標準SNP) & "',"
        gSL_Select = gSL_Select & "売上 = '" & RTrim(.売上) & "',"
        gSL_Select = gSL_Select & "原価 = '" & RTrim(.原価) & "',"
        gSL_Select = gSL_Select & "前期末棚卸評価単位 = '" & RTrim(.前期末棚卸評価単位) & "',"
        gSL_Select = gSL_Select & "車種 = '" & RTrim(.車種) & "',"
        gSL_Select = gSL_Select & "加工費 = '" & RTrim(.加工費) & "',"
        gSL_Select = gSL_Select & "材料費 = '" & RTrim(.材料費) & "',"
        gSL_Select = gSL_Select & "購入品費 = '" & RTrim(.購入品費) & "',"
        gSL_Select = gSL_Select & "型費 = '" & RTrim(.型費) & "',"
        gSL_Select = gSL_Select & "管理費 = '" & RTrim(.管理費) & "',"
        gSL_Select = gSL_Select & "ETC1 = '" & RTrim(.ETC1) & "',"
        gSL_Select = gSL_Select & "ETC2 = '" & RTrim(.ETC2) & "',"
        gSL_Select = gSL_Select & "ETC3 = '" & RTrim(.ETC3) & "'"
        gSL_Select = gSL_Select & " Where 品目コード = '" & RTrim(.品目コード) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KKHUpdate = True
End Function

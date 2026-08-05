Attribute VB_Name = "CNJ0315_Bas"
Option Explicit
'*---------------------------------------------------------------------------------------*
'*
'*                  かんばん 回転枚数・増減・平準化テーブル　データ作成
'*
'*---------------------------------------------------------------------------------------*

    Global g315_YM              As String       '(IN)   対象年月
    Global g315_ProcKB          As Byte         '(IN)   処理区分    ( 0:初回分, 1:変更分 )
    Global g315_FProc           As Byte         '(IN)   前方処理    ( 0:構成展開, 1:メンテナンス )
    
    Global g315_InsCnt(2)       As Long         '(OUT)  作成件数    ( 0:回転枚数, 1:増減, 2:平準化 )
    Global g315_UpdCnt(2)       As Long         '(OUT)  更新係数    ( 0:回転枚数, 1:増減, 2:平準化 )
    Global g315_ErrCnt          As Long         '(OUT)  エラー件数  ( 平準化設定更新エラー )
    
'*-------------------------------------------------*
'*      平準化設定更新エラーリスト
'*-------------------------------------------------*
Public Function HJKErr_List()
    
    If g315_ErrCnt = 0 Then Exit Function
    
'   < 印刷データ読み込み >
    gSL_Select = "Select HE.*, TR.略称 取引先名"
    gSL_Select = gSL_Select & " From 平準化設定更新エラーテーブル HE"
    gSL_Select = gSL_Select & " Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & " On  TR.取引先CD = HE.取引先"
    gSL_Select = gSL_Select & " And TR.取引先区分 = 1"
    gSL_Select = gSL_Select & " Where HE.作成端末 = Substring( HOST_NAME(), 1, 20 )"
    gSL_Select = gSL_Select & " Order By HE.品番, HE.取引先, HE.受入, HE.データ区分, HE.開始日, HE.終了日"
    Call HUERead(gSL_Select)
    
    Dim CisSimplePrint As New CisSimplePrint
    
    With CisSimplePrint
        .PrintTitle = "平準化設定更新エラーリスト"
        .Preservation = App.Title
        .PrintSize = A4
        .PrintOrientation = Landscape
        
        .pHeadItemSet 0, 左, "品番"
        .pHeadItemSet 1, 左, "取引先"
        .pHeadItemSet 2, 左, "取引先名"
        .pHeadItemSet 3, 左, "受入"
        .pHeadItemSet 4, 中, "開始日"
        .pHeadItemSet 5, 中, "終了日"
        .pHeadItemSet 6, 左, "区分"
        .pHeadItemSet 7, 左, "内容"
        .pHeadItemSet 8, 右, "便上限枚数"
        .pHeadItemSet 9, 右, "日上限枚数"
        .pHeadItemSet 10, 右, "月上限枚数"
    
        .pBodyItemSet 0, 22
        .pBodyItemSet 1, 7
        .pBodyItemSet 2, 20
        .pBodyItemSet 3, 4
        .pBodyItemSet 4, 10, 中
        .pBodyItemSet 5, 10, 中
        .pBodyItemSet 6, 4
        .pBodyItemSet 7, 4
        .pBodyItemSet 8, 10, 右, "#,###"
        .pBodyItemSet 9, 10, 右, "#,###"
        .pBodyItemSet 10, 10, 右, "#,###"
        
        .BodyInit
        
        If Not .CheckPrtCsv Then Exit Function
        
        Do Until (Not HUE_RDSTS)
            
            .pBodyValueSet 0, RTrim$(HUE.品番)
            .pBodyValueSet 1, RTrim$(HUE.取引先)
            .pBodyValueSet 2, RTrim$(HUE.取引先名)
            .pBodyValueSet 3, RTrim$(HUE.受入)
            
            gStr = ""
            If RTrim$(HUE.開始日) <> "" Then gStr = Format(RTrim$(HUE.開始日), "@@@@/@@/@@")
            .pBodyValueSet 4, gStr
            
            gStr = ""
            If RTrim$(HUE.終了日) <> "" Then gStr = Format(RTrim$(HUE.終了日), "@@@@/@@/@@")
            .pBodyValueSet 5, gStr
            
            If HUE.データ区分 = 0 Then gStr = "今回" Else gStr = "既存"
            .pBodyValueSet 6, gStr
            
            If HUE.設定内容区分 = 0 Then gStr = "有効" Else gStr = "無効"
            .pBodyValueSet 7, gStr
            
            .pBodyValueSet 8, HUE.便上限枚数
            .pBodyValueSet 9, HUE.日上限枚数
            .pBodyValueSet 10, HUE.月上限枚数
            .BodySet
            
            Call HUEReadNext
        Loop
        
        .PrintStart
    
    End With
    
End Function


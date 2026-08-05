Attribute VB_Name = "TextCommon"
Option Explicit
'+----------------------------------------------------------+
'+
'+       モジュール名  :テキスト変換用　共通モジュール
'+       モジュールID  :TextCommon.BAS
'+       作成日        :2003/12/10   CIS
'+       更新日        :2004/11/17   CIS 財務ﾃﾞｰﾀ取込処理追加
'+       変更日        :2005/03/31  By CIS 計算方法を取引先マスタの設定に従う
'+       変更日        :2005/06/13  By Cis      レイアウト情報内示 レイアウト番号(2)追加
'+                                                 補給内示ﾃﾞｰﾀ(トヨタ紡織形式)
'+       変更日        :2006/01/25  By Cis      レイアウト情報内示 レイアウト番号(3)追加
'+                                                 内示ﾃﾞｰﾀ(新レイアウト形式)
'+       変更日        :2008/05/09  By CIS 車体精工対応
'+       変更日        :2010/02/15  By CIS 愛知機械対応
'+
'+----------------------------------------------------------+

Global tx_ProcErr       As Boolean
Global tx_GridObj       As vsFlexArray
Global tx_SyoriCnt      As Long
Global tx_InCnt         As Long         ' ﾃﾞｰﾀ件数
Global tx_OutCnt        As Long         ' ﾃﾞｰﾀ出力件数
Global tx_HeadCnt       As Integer      ' ﾍｯﾀﾞｰ件数
Global tx_TailCnt       As Integer      ' ﾄﾚｰﾗｰ件数
Global tx_BunCnt        As Integer      ' 文言件数

Global tx_LogDelDay     As Long
Global tx_InputDir      As String
Global tx_JisyaChk      As Boolean

Global tx_ErMsg         As Long
Global tx_Result        As Long
Global tx_ProcDate      As Long
Global tx_Kensu         As Long

Global tx_KNo           As Long
Global tx_KTorcd        As String
Global tx_TYM           As String
Global cTKK_KYM         As String
Global cTKK_YMD         As String
Global cTTK_Kin         As Currency
'+---------------------------------------- 2005/03/31 Insert
Global cTTK_Syosu       As Currency
'+---------------------------------------- 2005/03/31 Insert End

Global tx_Shiki()       As String

Global ProcListObj      As ListBox

Global tx_HeadHCd       As String
Global tx_HeadYmd       As String
Global tx_HeadSeq       As String
Global tx_HeadFCd       As String
Global tx_HSiirCd       As String
Global tx_DataStr       As String

Global tx_LayoutType    As Long
Global tx_NjKouku()     As String
Global tx_NjKoukuX()    As String
Global tx_DataNo        As Long
Global tx_SetChar       As String
Global tx_SetLeng       As Long
Global tx_SetType       As Byte
Global tx_SetData       As String
Global tx_SetDataLong   As Long
Global tx_SetDataCurr   As Currency
Global tx_MonFlg        As Long

Public Type typTextConvert
    ERRCODE()   As String
    FILECD()    As String
    ProcRow()   As Long
    HFName()    As String
    LogDelDay() As Long
End Type

Public typText  As typTextConvert

Public Type typCon
    kStart      As Long
    kLeng       As Long
    kChar       As String
    kMonCnt     As Long
    kDataCnt    As Long
    kStartRecNo As Integer
    kEndRecNo   As Integer
    
    mStart()    As Long
    mLeng()     As Long
    mChar()     As String
    mDataMax    As Long
    mDStart()   As Long
    mDLeng()    As Long
    mDSetType() As Byte
    mDSetData() As String
    mDSetLeng() As Long

    dStart()    As Long
    dLeng()     As Long
    dChar()     As String
    dDataMax    As Long
    dDStart()   As Long
    dDLeng()    As Long
    dDSetType() As Byte
    dDSetData() As String
    dDSetLeng() As Long
    dDSetLite() As String
    dDSetSyos() As Long
End Type

Public typConvert   As typCon

Public txCnt1       As Long
Public txCnt2       As Long

'+----------------------------+
'+      変換設定エリア　初期化
'+----------------------------+
Public Sub typConvertClear()
    With typConvert
        .kStart = 0
        .kLeng = 0
        .kChar = ""
        .kMonCnt = 0
        .kDataCnt = 0
        .kStartRecNo = 0
        .kEndRecNo = 0
        
        ReDim .mStart(1)
        ReDim .mLeng(1)
        ReDim .mChar(1)
        .mDataMax = 0
        
        Erase .mDStart
        Erase .mDLeng
        Erase .mDSetType
        Erase .mDSetData
        Erase .mDSetLeng
    
        ReDim .dStart(1)
        ReDim .dLeng(1)
        ReDim .dChar(1)
        .dDataMax = 0
        
        Erase .dDStart
        Erase .dDLeng
        Erase .dDSetType
        Erase .dDSetData
        Erase .dDSetLeng
        Erase .dDSetLite
        Erase .dDSetSyos
    
    End With


    
End Sub
'+----------------------------+
'+      変換設定エリア　取得
'+----------------------------+
Public Function typConvertSet(lErr As String) As Boolean
    typConvertSet = False
    lErr = ""
    
    Call typConvertClear
    
    On Error GoTo typConvertSet_Er:
    
    gSL_Select = "Select * from TEXT設定テーブル "
    gSL_Select = gSL_Select & " Where レイアウト番号 = " & TKT.TEXTレイアウト
    gSL_Select = gSL_Select & " Order by 区分,連番,レコード番号"
    If Not TXSRead(gSL_Select) Then
       Call CreateText
       gSL_Select = "Select * from TEXT設定テーブル "
       gSL_Select = gSL_Select & " Where レイアウト番号 = " & TKT.TEXTレイアウト
       gSL_Select = gSL_Select & " Order by 区分,連番,レコード番号"
       If Not TXSRead(gSL_Select) Then
          Call TXSClose
          Call ProcLogPut("テキスト変換設定未登録(" & TKT.TEXTレイアウト & ")")
          lErr = "H4"
          Call TERPut("H4")
          Exit Function
       End If
    End If
    
    Do Until Not TXS_RDSTS
       With typConvert
        ' 基本情報
            If TXS.区分 = 1 Then
               .kStart = TXS.開始位置
               .kLeng = TXS.長さ
               .kChar = TXS.項目
               .kDataCnt = TXS.レコード定義件数
               .kStartRecNo = TXS.開始レコード
               .kEndRecNo = TXS.最終レコード
            End If
        ' 文言情報(基本)
            If TXS.区分 = 2 Then
               If TXS.連番 > UBound(.mStart) Then
                  ReDim Preserve .mStart(TXS.連番)
                  ReDim Preserve .mLeng(TXS.連番)
                  ReDim Preserve .mChar(TXS.連番)
               End If
               .mStart(TXS.連番) = TXS.開始位置
               .mLeng(TXS.連番) = TXS.長さ
               .mChar(TXS.連番) = TXS.項目
               .mDataMax = TXS.レコード定義件数
               .kMonCnt = TXS.連番
            End If
        ' 文言情報(項目)
            If TXS.区分 = 3 Then
               If TXS.連番 = 1 And TXS.レコード番号 = 1 Then
                  ReDim .mDStart(.kMonCnt, .mDataMax)
                  ReDim .mDLeng(.kMonCnt, .mDataMax)
                  ReDim .mDSetType(.kMonCnt, .mDataMax)
                  ReDim .mDSetData(.kMonCnt, .mDataMax)
                  ReDim .mDSetLeng(.kMonCnt, .mDataMax)
               End If
               
               .mDStart(TXS.連番, TXS.レコード番号) = TXS.開始位置
               .mDLeng(TXS.連番, TXS.レコード番号) = TXS.長さ
               .mDSetType(TXS.連番, TXS.レコード番号) = TXS.データタイプ
               .mDSetData(TXS.連番, TXS.レコード番号) = TXS.データ項目
               .mDSetLeng(TXS.連番, TXS.レコード番号) = TXS.出力長さ
            End If
        ' データ情報(基本)
            If TXS.区分 = 10 Then
               If TXS.連番 > UBound(.dStart) Then
                  ReDim Preserve .dStart(TXS.連番)
                  ReDim Preserve .dLeng(TXS.連番)
                  ReDim Preserve .dChar(TXS.連番)
               End If
               .dStart(TXS.連番) = TXS.開始位置
               .dLeng(TXS.連番) = TXS.長さ
               .dChar(TXS.連番) = TXS.項目
               If .dDataMax < TXS.レコード定義件数 Then
                  .dDataMax = TXS.レコード定義件数
               End If
            End If
        ' データ情報(項目)
            If TXS.区分 = 11 Then
               If TXS.連番 = 1 And TXS.レコード番号 = 1 Then
                  ReDim .dDStart(.kDataCnt, .dDataMax)
                  ReDim .dDLeng(.kDataCnt, .dDataMax)
                  ReDim .dDSetType(.kDataCnt, .dDataMax)
                  ReDim .dDSetData(.kDataCnt, .dDataMax)
                  ReDim .dDSetLeng(.kDataCnt, .dDataMax)
                  ReDim .dDSetLite(.kDataCnt, .dDataMax)
                  ReDim .dDSetSyos(.kDataCnt, .dDataMax)
               End If
               
               .dDStart(TXS.連番, TXS.レコード番号) = TXS.開始位置
               .dDLeng(TXS.連番, TXS.レコード番号) = TXS.長さ
               .dDSetType(TXS.連番, TXS.レコード番号) = TXS.データタイプ
               .dDSetData(TXS.連番, TXS.レコード番号) = TXS.データ項目
               .dDSetLeng(TXS.連番, TXS.レコード番号) = TXS.出力長さ
               .dDSetLite(TXS.連番, TXS.レコード番号) = TXS.項目
               .dDSetSyos(TXS.連番, TXS.レコード番号) = TXS.少数桁数
            End If
       End With
       Call TXSReadNext
    Loop

    Call TXSClose
    typConvertSet = True
    
    Exit Function
typConvertSet_Er:
    Static ER_MSG1 As String
    Static ER_MSG2 As String

    ER_MSG1 = "【 " & Err.Description & " 】"

    Call ProcLogPut(ER_MSG1 & "   <変換設定情報取得>")

    ER_MSG2 = Chr(13) & Space(3) & _
               "上記エラーにより、以降の処理を続行出来ません" & _
               Chr(10) & Chr(13) & Space(8) & _
               "内容を確認して再度実行して下さい"
    
    lErr = "H9"
    Call TERPut("H9", ER_MSG1, ER_MSG2)
End Function
'+-----------------------------------+
'+      項目定義の連番を取得する
'+-----------------------------------+
Public Function txDataNOGet() As Boolean
    txDataNOGet = False
    With typConvert
        tx_DataNo = 0
        If .kDataCnt = 1 Then
            tx_DataNo = 1
        Else
            For txCnt1 = 1 To .kDataCnt
                If Mid(tx_DataStr, .dStart(txCnt1), .dLeng(txCnt1)) = .dChar(txCnt1) Then
                   tx_DataNo = txCnt1
                   Exit For
                End If
            Next txCnt1
        End If
        If tx_DataNo = 0 Then Exit Function
    End With
    txDataNOGet = True
End Function
'+-----------------------------------+
'+      項目格納内容を取得する
'+-----------------------------------+
Public Function txDataSetGet(SetCnt As Long) As Boolean
    txDataSetGet = False
    With typConvert
        If Trim(.dDSetData(tx_DataNo, SetCnt)) = "" Then Exit Function
        tx_SetLeng = .dDLeng(tx_DataNo, SetCnt)
        tx_SetType = .dDSetType(tx_DataNo, SetCnt)
        If tx_SetLeng > .dDSetLeng(tx_DataNo, SetCnt) And tx_SetType = 0 Then
           tx_SetLeng = .dDSetLeng(tx_DataNo, SetCnt)
        End If
        tx_SetDataLong = 0: tx_SetDataCurr = 0
        If tx_SetType <> 2 Then
            tx_SetChar = Mid(tx_DataStr, .dDStart(tx_DataNo, SetCnt), .dDLeng(tx_DataNo, SetCnt))
            tx_SetData = Mid(tx_SetChar, 1, tx_SetLeng)
            If tx_SetType = 1 Then
                tx_SetDataLong = CisFun.Val2(Mid(tx_SetChar, 1, tx_SetLeng))
                If .dDSetSyos(tx_DataNo, SetCnt) <> 0 Then
                    tx_SetDataCurr = tx_SetDataLong / .dDSetSyos(tx_DataNo, SetCnt)
                End If
            End If
        Else
            tx_SetData = .dDSetLite(tx_DataNo, SetCnt)
        End If
    End With
    txDataSetGet = True
End Function
'+----------------------------+
'+      変換処理エリア　初期化
'+----------------------------+
Public Sub typTextClear()
    Erase typText.FILECD
    Erase typText.ERRCODE
    Erase typText.ProcRow

    Erase typText.HFName
    Erase typText.LogDelDay
    
End Sub
'+----------------------------+
'+      変換設定初期値　取得
'+----------------------------+
Public Sub TextInit()
    With CisFun
        ' 変換ログ削除日数
        .INI_Section = "TextConvert"
        .INI_Default = "60"
        .INI_Key = "LogDelDay"
        .GetIni
        tx_LogDelDay = .Val2(.INI_String)
        ' 変換ログ出力ファイル設定
        .Log_PutFlg = True
        .Log_FileInitial = "TEXT"
        .Log_FilePath = gIniExe & "Log\"
        .Log_DataHD = ""
        .Log_DataHDTime = あり
        .Log_DelDay = tx_LogDelDay

        .LogPut ""
        ' 変換元フォルダ
        .INI_Section = "TextConvert"
        .INI_Default = gIniExe & "Text\"
        .INI_Key = "InputDir"
        .GetIni
        tx_InputDir = RTrim(.INI_String)
        If Right(tx_InputDir, 1) <> "\" Then tx_InputDir = tx_InputDir & "\"
        ' 自社情報確認チェック
        .INI_Section = "TextConvert"
        .INI_Default = "Yes"
        .INI_Key = "Jisya"
        .GetIni
        If UCase(.INI_String) = "Y" Or UCase(.INI_String) = "YES" Then
           tx_JisyaChk = True
        Else
           tx_JisyaChk = False
        End If
    
    End With
End Sub
'+-----------------------------+
'+      処理中ログ出力
'+-----------------------------+
Public Sub ProcLogPut(LogData As String)
    Call CisFun.LogPut("(" & TKT.FILECD & ")" & vbTab & LogData)
    ProcListObj.AddItem "(" & TKT.FILECD & ")" & LogData
    DoEvents
End Sub
'*---------------------------------------*'
'*      TEXT変換履歴テーブル 更新処理       *
'*---------------------------------------*'
Public Function TextRirekiPut(THKNo As Long, PutKB As String, ErrSet As String, Optional DBNo As Integer) As Boolean

    TextRirekiPut = False
    On Error GoTo TextRirekiPut_ER:
    
    If THKNo = 0 Then
        If tx_LogDelDay <> 0 Then
           gStr = Format(Date, "YYYYMMDD")
           With CisFun
                .Date_RsFormat = fYMD
                .Date_Add gStr, D, tx_LogDelDay * -1, 1
                gStr = .Date_Result
           End With
           With CisDB
                .SQL = "DELETE TEXT変換履歴テーブル "
                .SQL = .SQL & " WHERE FILECD = '" & TKT.FILECD & "' "
                .SQL = .SQL & "   AND CONVERT(CHAR(8),作成日,112) < '" & gStr & "' "
            
                Call .DBExec
            End With
        End If
        
        gSL_Select = "SELECT MAX(履歴管理番号) + 1 AS 履歴管理番号 FROM TEXT変換履歴テーブル "
        If THRRead(gSL_Select, 1, DBNo) Then
            THKNo = THR.履歴管理番号
        End If
    End If
    If THKNo = 0 Then
       THKNo = 1
    End If

    gSL_Select = "Select * From TEXT変換履歴テーブル"
    gSL_Select = gSL_Select & " Where 履歴管理番号 = " & THKNo
    If Not THRRead(gSL_Select, 1, DBNo) Then
        Call ItemsClearTHR
        With THR
            .履歴管理番号 = THKNo
            .開始日時 = Now()
            .FILECD = TKT.FILECD
            .終了日時 = Now()
            .作成者 = gTanto
        End With
        Call THRInsert
        GoTo TextRirekiPut_ED
    End If
    
    With THR
        .FCODE = TKT.FCODE
        .HEAD件数 = tx_HeadCnt
        .TAIL件数 = tx_TailCnt
        .件数 = tx_InCnt
        .変換件数 = tx_SyoriCnt
        .文言件数 = tx_BunCnt
        .出力件数 = tx_OutCnt
        If Trim(ErrSet) = "" Then
          .処理コード = 0
          .受信エラー = ""
        Else
          .処理コード = 1
          .受信エラー = ErrSet
        End If
        .終了日時 = Now()
    End With
    Call THRUpdate
    
    If PutKB = "E" Then
       THKNo = 0
    End If
TextRirekiPut_ED:
    On Error GoTo 0
    TextRirekiPut = True
    Exit Function
TextRirekiPut_ER:
    Call ProcLogPut("TEXT変換履歴テーブル　更新ｴﾗｰ ")
End Function
'+--------------------------------------------------+
'+      内示定義作成
'+--------------------------------------------------+
Private Sub CreateText()
    ' 部品納入内示データ（トヨタ）
    gSL_Select = "Select * from TEXT設定テーブル "
    gSL_Select = gSL_Select & " Where レイアウト番号 = 0"
    If Not TXSRead(gSL_Select, 1) Then
       Call CreateNJToyota
    End If
    ' 部品納入内示データ（トヨタ紡織:旧アラコ）
    gSL_Select = "Select * from TEXT設定テーブル "
    gSL_Select = gSL_Select & " Where レイアウト番号 = 1"
    If Not TXSRead(gSL_Select, 1) Then
       Call CreateNJTB
    End If
'+------------------------------------------------------------------  2005/06/13 Insert
    ' 補給部品納入内示データ（トヨタ紡織:旧アラコ）
    gSL_Select = "Select * from TEXT設定テーブル "
    gSL_Select = gSL_Select & " Where レイアウト番号 = 2"
    If Not TXSRead(gSL_Select, 1) Then
       Call CreateNJTBH
    End If
'+------------------------------------------------------------------  2005/06/13 Insert End
'+------------------------------------------------------------------  2006/01/25 Insert
    ' 内示データ（トヨタ車体）
    gSL_Select = "Select * from TEXT設定テーブル "
    gSL_Select = gSL_Select & " Where レイアウト番号 = 3"
    If Not TXSRead(gSL_Select, 1) Then
       Call CreateNJTYNew
    End If
'+------------------------------------------------------------------  2006/01/25 Insert End
'+------------------------------------------------------------------  2008/05/09 Insert
    ' 内示データ（トヨタ車体精工）
    gSL_Select = "Select * from TEXT設定テーブル "
    gSL_Select = gSL_Select & " Where レイアウト番号 = 6"
    If Not TXSRead(gSL_Select, 1) Then
       Call CreateNJTSK
    End If
'+------------------------------------------------------------------  2008/05/09 Insert End
'+------------------------------------------------------------------  2010/02/15 Insert
    ' 内示データ（愛知機械）
    gSL_Select = "Select * from TEXT設定テーブル "
    gSL_Select = gSL_Select & " Where レイアウト番号 = 7"
    If Not TXSRead(gSL_Select, 1) Then
       Call CreateNJAIT
    End If
'+------------------------------------------------------------------  2010/02/15 Insert End
End Sub
'+------------------------------------------+
'+       内示設定テキスト作成（トヨタ）
'+------------------------------------------+
Private Sub CreateNJToyota()
    With TXS
        ' 基本情報
        Call ItemsClearTXS
        .区分 = 1
        .レコード定義件数 = 3
        .開始位置 = 1
        .長さ = 1
        .項目 = "D"
        .開始レコード = 2
        .最終レコード = 3
        Call TXSInsert
        ' 文言情報
        Call ItemsClearTXS
        .区分 = 2
        .連番 = 1
        .レコード定義件数 = 2
        .開始位置 = 2
        .長さ = 1
        .項目 = "A"
        Call TXSInsert
        Call ItemsClearTXS
        .区分 = 2
        .連番 = 2
        .レコード定義件数 = 2
        .開始位置 = 2
        .長さ = 1
        .項目 = "B"
        Call TXSInsert
        ' 文言情報 項目定義(1)
        Call ItemsClearTXS
        .区分 = 3
        .連番 = 1
        .レコード番号 = 1
        .レコード定義件数 = 2
        .開始位置 = 3
        .長さ = 1
        .項目 = ""
        .データ項目 = "連番"
        .データタイプ = 1
        .出力長さ = 1
        Call TXSInsert
        Call ItemsClearTXS
        .区分 = 3
        .連番 = 1
        .レコード番号 = 2
        .開始位置 = 4
        .長さ = 200
        .項目 = ""
        .データ項目 = "文言"
        .データタイプ = 0
        .出力長さ = 200
        Call TXSInsert
        ' 文言情報 項目定義(2)
        Call ItemsClearTXS
        .区分 = 3
        .連番 = 2
        .レコード番号 = 1
        .開始位置 = 3
        .長さ = 1
        .項目 = ""
        .データ項目 = "連番"
        .データタイプ = 1
        .出力長さ = 1
        Call TXSInsert
        Call ItemsClearTXS
        .区分 = 3
        .連番 = 2
        .レコード番号 = 2
        .開始位置 = 4
        .長さ = 200
        .項目 = ""
        .データ項目 = "文言"
        .データタイプ = 0
        .出力長さ = 200
        Call TXSInsert
        ' 項目定義(1)
        gLong = 0
        Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 10: .長さ = 4: .データ項目 = "納入先"
        .データタイプ = 0:   .出力長さ = 7
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 14: .長さ = 4: .データ項目 = "内示年月"
        .データタイプ = 0:   .出力長さ = 4
        
        Call TXSInsert: Call ItemsClearTXS
        ' レコード定義
        .区分 = 10: .連番 = 1: .レコード定義件数 = gLong: .開始位置 = 2: .長さ = 1: .項目 = "C"
        ' 項目定義(2)
        gLong = 0
        Call TXSInsert: Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 4: .長さ = 1: .データ項目 = "状態区分"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 5: .長さ = 1: .データ項目 = "納品書不作成区分"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 6: .長さ = 1: .データ項目 = "仕入先工区"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 7: .長さ = 12: .データ項目 = "品番"
        .データタイプ = 0:   .出力長さ = 18
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 25: .長さ = 1: .データ項目 = "層別区分"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 25: .長さ = 6: .データ項目 = "管理NO"
        .データタイプ = 0:   .出力長さ = 6
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 32: .長さ = 5: .データ項目 = "納入先"
        .データタイプ = 0:   .出力長さ = 7
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 32: .長さ = 4: .データ項目 = "持込先"
        .データタイプ = 0:   .出力長さ = 4
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 36: .長さ = 1: .データ項目 = "持込先工場"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 37: .長さ = 2: .データ項目 = "支給先受入"
        .データタイプ = 0:   .出力長さ = 2
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 39: .長さ = 3: .データ項目 = "出庫部署"
        .データタイプ = 0:   .出力長さ = 3
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 42: .長さ = 2: .データ項目 = "支給先担当者"
        .データタイプ = 0:   .出力長さ = 2
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 44: .長さ = 2: .データ項目 = "支給係"
        .データタイプ = 0:   .出力長さ = 2
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 46: .長さ = 2: .データ項目 = "受入"
        .データタイプ = 0:   .出力長さ = 2
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 48: .長さ = 2: .データ項目 = "整備"
        .データタイプ = 0:   .出力長さ = 2
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 50: .長さ = 2: .データ項目 = "担当者"
        .データタイプ = 0:   .出力長さ = 2
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 52: .長さ = 5: .データ項目 = "収容数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 57: .長さ = 3: .データ項目 = "背番号"
        .データタイプ = 0:   .出力長さ = 4
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 62: .長さ = 5: .データ項目 = "サイクル"
        .データタイプ = 0:   .出力長さ = 5
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 67: .長さ = 3: .データ項目 = "納品書発行枚数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 70: .長さ = 3: .データ項目 = "かんばん回転枚数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 73: .長さ = 1: .データ項目 = "かんばん増減符号"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 74: .長さ = 3: .データ項目 = "かんばん増減枚数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 77: .長さ = 6: .データ項目 = "適用開始年月日"
        .データタイプ = 0:   .出力長さ = 8
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 83: .長さ = 6: .データ項目 = "適用終了年月日"
        .データタイプ = 0:   .出力長さ = 8
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 89: .長さ = 8: .データ項目 = "内示数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 97: .長さ = 8: .データ項目 = "翌月内示数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 105: .長さ = 8: .データ項目 = "翌々月内示数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 113: .長さ = 8: .データ項目 = "当初内示数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 121: .長さ = 1: .データ項目 = "変更数符号"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 122: .長さ = 8: .データ項目 = "変更数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 130: .長さ = 1: .データ項目 = "調整数符号"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 131: .長さ = 8: .データ項目 = "調整数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 139: .長さ = 1: .データ項目 = "納入区分"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 44: .長さ = 1: .データ項目 = "内外区分"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 45: .長さ = 1: .データ項目 = "現旧区分"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 60: .長さ = 1: .データ項目 = "管理区分"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 61: .長さ = 1: .データ項目 = "生産区分"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 139: .長さ = 1: .データ項目 = "AB区分"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 140: .長さ = 6: .データ項目 = "内示数量01"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 146: .長さ = 6: .データ項目 = "内示数量02"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 152: .長さ = 6: .データ項目 = "内示数量03"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 158: .長さ = 6: .データ項目 = "内示数量04"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 164: .長さ = 6: .データ項目 = "内示数量05"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS
        ' レコード定義
        .区分 = 10: .連番 = 2: .レコード定義件数 = gLong: .開始位置 = 3: .長さ = 1: .項目 = "1"
        ' 項目定義(3)
        gLong = 0
        For gInt = 6 To 31
            Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
            .レイアウト番号 = 0: .区分 = 11: .連番 = 3: .レコード番号 = gLong: .開始位置 = 4 + ((gInt - 6) * 6): .長さ = 6: .データ項目 = "内示数量" & Format(gInt, "00")
            .データタイプ = 1:   .出力長さ = 0
        Next gInt
        
        Call TXSInsert: Call ItemsClearTXS
        ' レコード定義
        .区分 = 10: .連番 = 3: .レコード定義件数 = gLong: .開始位置 = 3: .長さ = 1: .項目 = "2": Call TXSInsert
        
    End With
    With CisDB
        .SQL = "UPDATE TEXT設定テーブル Set 作成者 = 'SIMPR-A',作成端末 = 'SIMPR-A'"
        .SQL = .SQL & " Where レイアウト番号 = 0"
        .DBExec
    End With

End Sub
'+------------------------------------------+
'+       内示設定テキスト作成（紡織）
'+------------------------------------------+
Private Sub CreateNJTB()
    With TXS
        ' 基本情報
        Call ItemsClearTXS
        .レイアウト番号 = 1
        .区分 = 1
        .レコード定義件数 = 1
        .開始位置 = 1
        .長さ = 2
        .項目 = "DN"
        .開始レコード = 1
        .最終レコード = 1
        Call TXSInsert
        ' 項目定義(1)
        gLong = 0
        Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 3: .長さ = 12: .データ項目 = "品番"
        .データタイプ = 0:   .出力長さ = 25
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 19: .長さ = 1: .データ項目 = "仕入先工区"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 20: .長さ = 4: .データ項目 = "受入"
'        .データタイプ = 0:   .出力長さ = 2
        .データタイプ = 0:   .出力長さ = 4
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 0: .長さ = 4: .データ項目 = "納入先"
        .データタイプ = 2:   .出力長さ = 7: .項目 = ""
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 24: .長さ = 2: .データ項目 = "担当者"
        .データタイプ = 0:   .出力長さ = 2
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 26: .長さ = 20: .データ項目 = "品名"
        .データタイプ = 0:   .出力長さ = 20
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 46: .長さ = 1: .データ項目 = "納入区分"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 47: .長さ = 6: .データ項目 = "管理NO"
        .データタイプ = 0:   .出力長さ = 6
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 53: .長さ = 5: .データ項目 = "収容数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
    '+------------------------------------ 2005/06/13 Delete
'        .レイアウト番号 = 1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 47: .長さ = 3: .データ項目 = "背番号"
    '+------------------------------------ 2005/06/13 Delete End
    '+------------------------------------ 2005/06/13 Insert
        .レイアウト番号 = 1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 60: .長さ = 3: .データ項目 = "背番号"
    '+------------------------------------ 2005/06/13 Insert End
        .データタイプ = 0:   .出力長さ = 4
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 69: .長さ = 6: .データ項目 = "翌月内示数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 75: .長さ = 6: .データ項目 = "翌々月内示数"
        .データタイプ = 1:   .出力長さ = 0
        For gInt = 1 To 31
            Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
            .レイアウト番号 = 1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 96 + ((gInt - 1) * 5): .長さ = 5: .データ項目 = "内示数量" & Format(gInt, "00")
            .データタイプ = 1:   .出力長さ = 0
        Next gInt
        Call TXSInsert: Call ItemsClearTXS
        ' レコード定義
        .レイアウト番号 = 1: .区分 = 10: .連番 = 1: .レコード定義件数 = gLong:  .開始位置 = 0: .長さ = 0: .項目 = "": Call TXSInsert
        
    End With
    
    With CisDB
        .SQL = "UPDATE TEXT設定テーブル Set 作成者 = 'SIMPR-A',作成端末 = 'SIMPR-A'"
        .SQL = .SQL & " Where レイアウト番号 = 1"
        .DBExec
    End With

End Sub
'+------------------------------------------------------------------  2005/06/13 Insert
'+------------------------------------------+
'+       内示設定テキスト作成（紡織）
'+------------------------------------------+
Private Sub CreateNJTBH()
    With TXS
        ' 基本情報
        Call ItemsClearTXS
        .レイアウト番号 = 2
        .区分 = 1
        .レコード定義件数 = 1
        .開始位置 = 1
        .長さ = 2
        .項目 = "DJ"
        .開始レコード = 1
        .最終レコード = 1
        Call TXSInsert
        ' 項目定義(1)
        gLong = 0
        Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 2: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 3: .長さ = 12: .データ項目 = "品番"
        .データタイプ = 0:   .出力長さ = 25
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 2: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 19: .長さ = 1: .データ項目 = "仕入先工区"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 2: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 20: .長さ = 2: .データ項目 = "受入"
        .データタイプ = 0:   .出力長さ = 2
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 2: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 0: .長さ = 4: .データ項目 = "納入先"
        .データタイプ = 2:   .出力長さ = 7: .項目 = ""
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 2: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 22: .長さ = 20: .データ項目 = "品名"
        .データタイプ = 0:   .出力長さ = 20
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 2: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 42: .長さ = 6: .データ項目 = "管理NO"
        .データタイプ = 0:   .出力長さ = 6
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 2: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 48: .長さ = 5: .データ項目 = "収容数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 2: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 63: .長さ = 7: .データ項目 = "翌月内示数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = 2: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 70: .長さ = 7: .データ項目 = "翌々月内示数"
        .データタイプ = 1:   .出力長さ = 0
        For gInt = 1 To 31
            Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
            .レイアウト番号 = 2: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 77 + ((gInt - 1) * 5): .長さ = 5: .データ項目 = "内示数量" & Format(gInt, "00")
            .データタイプ = 1:   .出力長さ = 0
        Next gInt
        Call TXSInsert: Call ItemsClearTXS
        ' レコード定義
        .レイアウト番号 = 2: .区分 = 10: .連番 = 1: .レコード定義件数 = gLong:  .開始位置 = 0: .長さ = 0: .項目 = "": Call TXSInsert
        
    End With
    
    With CisDB
        .SQL = "UPDATE TEXT設定テーブル Set 作成者 = 'SIMPR-A',作成端末 = 'SIMPR-A'"
        .SQL = .SQL & " Where レイアウト番号 = 2"
        .DBExec
    End With

End Sub
'+------------------------------------------------------------------  2005/06/13 Insert End



'+------------------------------------------------------------------  2006/01/25 Insert
'+---------------------------------------------+
'+       内示設定テキスト作成（車体:新ﾌｫｰﾏｯﾄ）
'+---------------------------------------------+
Public Sub CreateNJTYNew()
    gCnt1 = 3   ' ﾚｲｱｳﾄ番号
    With TXS
        ' 基本情報
        Call ItemsClearTXS
        .レイアウト番号 = gCnt1
        .区分 = 1
        .レコード定義件数 = 2
        .開始位置 = 1
        .長さ = 1
        .項目 = "D"
        .開始レコード = 2
        .最終レコード = 2
        Call TXSInsert
        ' 文言情報
        Call ItemsClearTXS
        .レイアウト番号 = gCnt1: .区分 = 2: .連番 = 1: .レコード定義件数 = 2: .項目 = "A": .開始位置 = 7: .長さ = 1
        Call TXSInsert: Call ItemsClearTXS
        .レイアウト番号 = gCnt1: .区分 = 2: .連番 = 2: .レコード定義件数 = 2: .項目 = "B": .開始位置 = 7: .長さ = 1
        Call TXSInsert: Call ItemsClearTXS
        ' 文言情報 項目定義(1)
        .レイアウト番号 = gCnt1: .区分 = 3: .連番 = 1: .レコード番号 = 1: .レコード定義件数 = 2: .開始位置 = 8: .長さ = 1
        .データ項目 = "連番": .データタイプ = 1: .出力長さ = 1
        Call TXSInsert: Call ItemsClearTXS
        .レイアウト番号 = gCnt1: .区分 = 3: .連番 = 1: .レコード番号 = 2: .開始位置 = 9: .長さ = 200
        .データ項目 = "文言": .データタイプ = 1: .出力長さ = 200
        Call TXSInsert: Call ItemsClearTXS
        ' 文言情報 項目定義(2)
        .レイアウト番号 = gCnt1: .区分 = 3: .連番 = 2: .レコード番号 = 1: .レコード定義件数 = 2: .開始位置 = 8: .長さ = 1
        .データ項目 = "連番": .データタイプ = 1: .出力長さ = 1
        Call TXSInsert: Call ItemsClearTXS
        .レイアウト番号 = gCnt1: .区分 = 3: .連番 = 2: .レコード番号 = 2: .開始位置 = 9: .長さ = 200
        .データ項目 = "文言": .データタイプ = 1: .出力長さ = 200
        Call TXSInsert: Call ItemsClearTXS
        ' レコード定義(1)
        .レイアウト番号 = gCnt1: .区分 = 10: .連番 = 1: .レコード定義件数 = 1: .開始位置 = 7: .長さ = 1: .項目 = "C"
        Call TXSInsert: Call ItemsClearTXS
        ' レコード定義(2)
        .レイアウト番号 = gCnt1: .区分 = 10: .連番 = 2: .レコード定義件数 = 59: .開始位置 = 7: .長さ = 2: .項目 = "D1"
        Call TXSInsert: Call ItemsClearTXS
        
        ' 項目定義(1)
        gLong = 0
        Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 21: .長さ = 6: .データ項目 = "内示年月"
        .データタイプ = 0:   .出力長さ = 6
        Call TXSInsert
        ' 項目定義(2)
        gLong = 0
        Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 11: .長さ = 12: .データ項目 = "品番"
        .データタイプ = 0:   .出力長さ = 25
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 30: .長さ = 6: .データ項目 = "納入先"
        .データタイプ = 0:   .出力長さ = 7: .項目 = ""
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 51: .長さ = 2: .データ項目 = "受入"
        .データタイプ = 0:   .出力長さ = 2
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 23: .長さ = 1: .データ項目 = "層別区分"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 84: .長さ = 8: .データ項目 = "適用開始年月日"
        .データタイプ = 0:   .出力長さ = 8
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 92: .長さ = 8: .データ項目 = "適用終了年月日"
        .データタイプ = 0:   .出力長さ = 8
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 100: .長さ = 9: .データ項目 = "内示数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 109: .長さ = 9: .データ項目 = "翌月内示数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 118: .長さ = 9: .データ項目 = "翌々月内示数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 24: .長さ = 6: .データ項目 = "管理NO"
        .データタイプ = 0:   .出力長さ = 6
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 55: .長さ = 5: .データ項目 = "収容数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 60: .長さ = 4: .データ項目 = "背番号"
        .データタイプ = 0:   .出力長さ = 4
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 64: .長さ = 10: .データ項目 = "サイクル"
        .データタイプ = 0:   .出力長さ = 10
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 9: .長さ = 1: .データ項目 = "状態区分"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 10: .長さ = 1: .データ項目 = "仕入先工区"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 30: .長さ = 5: .データ項目 = "持込先"
        .データタイプ = 0:   .出力長さ = 4
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 35: .長さ = 1: .データ項目 = "持込先工場"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 74: .長さ = 3: .データ項目 = "納品書発行枚数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 77: .長さ = 3: .データ項目 = "かんばん回転枚数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 80: .長さ = 1: .データ項目 = "かんばん増減符号"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 81: .長さ = 3: .データ項目 = "かんばん増減枚数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 127: .長さ = 9: .データ項目 = "当初内示数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 136: .長さ = 1: .データ項目 = "変更数符号"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 137: .長さ = 9: .データ項目 = "変更数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 146: .長さ = 1: .データ項目 = "調整数符号"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 147: .長さ = 9: .データ項目 = "調整数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 156: .長さ = 1: .データ項目 = "納入区分"
        .データタイプ = 0:   .出力長さ = 1
        For gInt = 1 To 31
            Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
            .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 157 + ((gInt - 1) * 9): .長さ = 9: .データ項目 = "内示数量" & Format(gInt, "00")
            .データタイプ = 1:   .出力長さ = 0
        Next gInt
        Call TXSInsert: Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 2: .レコード番号 = gLong: .開始位置 = 443: .長さ = 40: .データ項目 = "品名"
        .データタイプ = 0:   .出力長さ = 20
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
    End With
    
    With CisDB
        .SQL = "UPDATE TEXT設定テーブル Set 作成者 = 'SIMPR-A',作成端末 = 'SIMPR-A'"
        .SQL = .SQL & " Where レイアウト番号 = " & gCnt1
        .DBExec
    End With

End Sub
'+------------------------------------------------------------------  2006/01/25 Insert End

'+------------------------------------------------------------------  2008/05/09 INSERT START
'+---------------------------------------------+
'+       内示設定テキスト作成（車体精工）
'+---------------------------------------------+
Public Sub CreateNJTSK()
    gCnt1 = 6   ' ﾚｲｱｳﾄ番号
    With TXS
        ' 基本情報
        Call ItemsClearTXS
        .レイアウト番号 = gCnt1: .区分 = 1:             .レコード定義件数 = 1:  .開始位置 = 1: .長さ = 1: .項目 = ""
        .開始レコード = 1: .最終レコード = 1
        Call TXSInsert
        ' 文言情報
        Call ItemsClearTXS
        ' レコード定義(1)
        .レイアウト番号 = gCnt1: .区分 = 10: .連番 = 1: .レコード定義件数 = 14: .開始位置 = 1: .長さ = 1: .項目 = ""
        Call TXSInsert: Call ItemsClearTXS
        
        ' 項目定義(1)
        gLong = 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 1: .長さ = 4:   .データ項目 = "仕入先"
        .データタイプ = 0:   .出力長さ = 4
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 5: .長さ = 5:   .データ項目 = "背番号"
        .データタイプ = 0:   .出力長さ = 5
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 10: .長さ = 6:  .データ項目 = "管理NO"
        .データタイプ = 0:   .出力長さ = 6
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 16: .長さ = 12: .データ項目 = "品番"
        .データタイプ = 0:   .出力長さ = 25
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 28: .長さ = 1: .データ項目 = "工区"
        .データタイプ = 0:   .出力長さ = 1
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 29: .長さ = 2: .データ項目 = "受入"
        .データタイプ = 0:   .出力長さ = 2
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 31: .長さ = 5: .データ項目 = "ダミー"
        .データタイプ = 0:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 36: .長さ = 4: .データ項目 = "後工程"
        .データタイプ = 0:   .出力長さ = 4
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 40: .長さ = 4: .データ項目 = "内示年月"
        .データタイプ = 0:   .出力長さ = 4
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 44: .長さ = 6: .データ項目 = "内示数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 50: .長さ = 6: .データ項目 = "翌月内示数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 56: .長さ = 6: .データ項目 = "翌々月内示数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 62: .長さ = 6: .データ項目 = "作成日"
        .データタイプ = 1:   .出力長さ = 10
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 68: .長さ = 61: .データ項目 = "ダミー"
        .データタイプ = 0:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
    End With
    
    With CisDB
        .SQL = "UPDATE TEXT設定テーブル Set 作成者 = 'SIMPR-A',作成端末 = 'SIMPR-A'"
        .SQL = .SQL & " Where レイアウト番号 = " & gCnt1
        .DBExec
    End With

End Sub
'+------------------------------------------------------------------  2008/05/09 INSERT END

'+------------------------------------------------------------------  2010/02/15 INSERT START
'+---------------------------------------------+
'+       内示設定テキスト作成（愛知機械）
'+---------------------------------------------+
Public Sub CreateNJAIT()
    gCnt1 = 7   ' ﾚｲｱｳﾄ番号
    With TXS
        ' 基本情報
        Call ItemsClearTXS
        .レイアウト番号 = gCnt1: .区分 = 1:             .レコード定義件数 = 1:  .開始位置 = 1: .長さ = 1: .項目 = ""
        .開始レコード = 1: .最終レコード = 1
        Call TXSInsert
        ' 文言情報
        Call ItemsClearTXS
        ' レコード定義(1)
        .レイアウト番号 = gCnt1: .区分 = 10: .連番 = 1: .レコード定義件数 = 14: .開始位置 = 1: .長さ = 1: .項目 = ""
        Call TXSInsert: Call ItemsClearTXS
        
        ' 項目定義(1)
        gLong = 1
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 7: .長さ = 6:   .データ項目 = "納入先"
        .データタイプ = 0:   .出力長さ = 7
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 14: .長さ = 12: .データ項目 = "品番"
        .データタイプ = 0:   .出力長さ = 25
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 75: .長さ = 2: .データ項目 = "受入"
        .データタイプ = 0:   .出力長さ = 2
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 108: .長さ = 7: .データ項目 = "内示数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 129: .長さ = 7: .データ項目 = "翌月内示数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 150: .長さ = 7: .データ項目 = "翌々月内示数"
        .データタイプ = 1:   .出力長さ = 0
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        
        .レイアウト番号 = gCnt1: .区分 = 11: .連番 = 1: .レコード番号 = gLong: .開始位置 = 157: .長さ = 6: .データ項目 = "内示年月"
        .データタイプ = 0:   .出力長さ = 6
        Call TXSInsert:  Call ItemsClearTXS: gLong = gLong + 1
        
    End With
    
    With CisDB
        .SQL = "UPDATE TEXT設定テーブル Set 作成者 = 'SIMPR-A',作成端末 = 'SIMPR-A'"
        .SQL = .SQL & " Where レイアウト番号 = " & gCnt1
        .DBExec
    End With

End Sub
'+------------------------------------------------------------------  2010/02/15 INSERT END


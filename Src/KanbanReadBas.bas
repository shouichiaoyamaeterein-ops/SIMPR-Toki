Attribute VB_Name = "KanbanReadBas"
Option Explicit
'+---------------------------------------------------------------+
'+
'+      モジュール名  :かんばん読取モジュール
'+
'+                    MsComm 使用
'+
'+      作成日        :2004/01/16   By CIS         Version: 1.00
'+      変更日        :2004/09/17   By CIS 材料混在可能  Version: 2.00
'+      変更日        :2004/11/26   By CIS 滞留日付更新(ZEROは更新なしに変更)
'+      変更日        :2005/06/14   By CIS かんばんチェックテーブル削除日数
'+                                         かんばん重複チェックテーブル削除日数　初期値変更
'+      変更日        :2005/06/18   By CIS かんばん流動履歴テーブル削除機能追加
'+      変更日        :2005/11/07   By CIS オプション追加（発注かんばん読取ｴﾗｰﾘｽﾄ発行・発注打切りﾘｽﾄ発行・受注かんばん読取ｴﾗｰﾘｽﾄ発行）
'+      変更日        :2006/02/03   By CIS オプション追加（未流動管理）
'+                                         ポケット最大数より大きいポケットの色設定
'+                                         エラーポケット可変対応
'+      変更日        :2006/03/24   By CIS 片面読み（ﾊｰﾄﾞ）時、読取不能が処理されていないのを修正
'+      変更日        :2006/12/20   By CIS 版数チェック追加
'+
'+---------------------------------------------------------------+
Global kBas_Proc                As String       '  _/部品発注,Z/材料発注,J/受注,S/仕分,H/部品・材料発注
Global kBas_ExecMode            As Boolean      ' 実行モード False:RS232c接続,True:ﾃｽﾄﾓｰﾄﾞ
Global kBas_ReadOk              As Boolean
Global kBas_ReadMsg             As String
Global kBas_DemoFlg             As String
Global kBas_Apparatus           As String       ' 使用機器

Global kBas_PortNo              As Integer
Global kBas_Setting             As String
Global kBas_Hand                As Integer
Global kBas_232cLog             As Boolean

Global kBas_CommData            As String
Global kBas_CommData2           As String
Global kBas_ChkData             As String

Global kBas_ErrSts              As Integer
Global kBas_ReadSTS             As String
Global kBas_ErrMsg              As String

Global kBas_Poket               As Integer
Global kBas_ErrNo               As Integer
Global kBas_ReadData(1)         As String
Global kBas_PCnt                As Long
Global kBas_ReadFlg             As Boolean

Global kBas_DemoFno             As Integer
Global kBas_DemoFName           As String
Global kBas_DemoDir             As String

Global kBas_Event               As String
Global kBas_EventMsg            As String
Global kBas_EventCode           As Integer
Global kBas_EventRtn            As String

Global kBas_PoketMax            As Integer      ' かんばんﾎﾟｹｯﾄ最大数
'---------------------------------------------------------  2006/02/03 Insert
Global kBas_MaxPColor           As Long         ' 最大数より大きい時の色
Global kBas_PoketErr            As Integer      ' かんばんﾎﾟｹｯﾄ外
Global kBas_PoketErrMsg         As String       ' かんばんﾎﾟｹｯﾄ外名称
Global kBas_PoketSave           As Long         ' かんばんﾎﾟｹｯﾄ最大数退避
'---------------------------------------------------------  2006/02/03 Insert End
' 固定ﾎﾟｹｯﾄ
Global Const kBas_PokErr        As Integer = 1  ' 読取ｴﾗｰ
Global kBas_TKey()              As String       ' ﾎﾟｹｯﾄｷｰ(仕分用)
Global kBas_TKey2()             As String       ' ﾎﾟｹｯﾄｷｰ2(仕分用)
Global kBas_TMai()              As Long         ' ﾎﾟｹｯﾄ　指定枚数
Global kBas_TMaiMax(5)          As Long         ' ﾎﾟｹｯﾄ単位　最大指定枚数
Global kBas_TWrork()            As String       ' 可変配列ワーク
'-------------------------------------------------------------------   2006/02/03 Insert
Global kBas_ErSetNo()           As Long         ' ｴﾗｰﾎﾟｹｯﾄ固定(ｴﾗｰ番号)
Global kBas_ErSetMsg()          As String       ' ｴﾗｰﾎﾟｹｯﾄ固定(名称)
Global kBas_ErNo()              As Long         ' ｴﾗｰ番号
Global kBas_ErPno()             As Long         ' ｴﾗｰﾎﾟｹｯﾄ
Global kBas_ErMsg()             As String       ' ｴﾗｰ内容
'-------------------------------------------------------------------   2006/02/03 Insert End

Global kBas_Col                 As Long
Global kBas_Row                 As Long
Global kBas_ColSu               As Long

Global kBas_DelDay1             As Long         ' かんばんﾁｪｯｸﾃｰﾌﾞﾙ削除日数
Global kBas_DelDay2             As Long         ' かんばん重複ﾃｰﾌﾞﾙ削除ﾌﾗｸﾞ

Dim DCnt                        As Long
Dim DId                         As Long

Dim kBas_StoRTN                 As Long

Dim TDay                        As String

' かんばんチェック用
Type Kanban_Check_Type
    ProcKB      As String
    HCKB        As String
    FTime       As String   '振出時間
    Group       As Byte     '読取ｸﾞﾙｰﾌﾟ
    Op1Value    As String   'ｵﾌﾟｼｮﾝ設定(平準化超過の扱い _/発注なし,1/発注あり)
    Op2Value    As String   'ｵﾌﾟｼｮﾝ設定(発注即時処理 _/即時なし,1/即時処理)
'+---------------------------------------------------------------------------------  2005/06/18 Insert
    Op3Value    As String   'ｵﾌﾟｼｮﾝ設定(かんばん流動履歴削除日数 0/削除なし,0以外/削除日数)
'+---------------------------------------------------------------------------------  2005/06/18 Insert End
    HYMD        As String   '発注日
    HBIN        As String   '発注便
    ErrCnt      As Long     'エラー件数
    OutCnt      As Long     '正常出力件数
    ErrCntZ     As Long     'エラー件数
    OutCntZ     As Long     '正常出力件数
'+---------------------------------------------------------------------------------  2005/11/07 Insert
    OpHErrPrt   As Byte   'ｵﾌﾟｼｮﾝ設定(発注かんばん読取エラーリスト 0/発行あり,1/発行なし)
    OpHUchiPrt  As Byte   'ｵﾌﾟｼｮﾝ設定(発注打切リスト 0/発行あり,1/発行なし)
    OpJErrPrt   As Byte   'ｵﾌﾟｼｮﾝ設定(受注かんばん読取エラーリスト 0/発行あり,1/発行なし)
'+---------------------------------------------------------------------------------  2005/11/07 Insert End
'+---------------------------------------------------------------------------------  2006/02/03 Insert
    OpMiryu     As Byte   'ｵﾌﾟｼｮﾝ設定(未流動管理 0/有効,1/無効)
'+---------------------------------------------------------------------------------  2006/02/03 Insert End
End Type
Public kBas_Check               As Kanban_Check_Type

'--------------------------------------------------------------------------
'*      かんばん処理  初期処理
'--------------------------------------------------------------------------
'Public Function cKanbanInit() As Boolean                               ' 2006/02/03 Delete
Public Function cKanbanInit(Optional sGroup As String = "") As Boolean  ' 2006/02/03 Insert
    cKanbanInit = False
    ' --- デバッグ用(かんばん)
    If gDemoPrt <> 0 Then
       kBas_ExecMode = True
    Else
       kBas_ExecMode = False
    End If
    
    With CisFun
        .Log_PutFlg = kBas_ExecMode Or kBas_232cLog
        .Log_FileInitial = "KanRead"
        .Log_FilePath = gIniExe & "Log\"
        .Log_DataHD = ""
        .Log_DataHDTime = あり
        .Log_DelDay = 3

        .LogPut ""
    End With
    
' 自社情報を求める
    Qr_Jisya = RTrim(gCompany)  ' 自社情報セット
' ﾎﾟｹｯﾄ指定最大数を求める
    kBas_MaxPColor = 0              ' 2006/02/03 Insert
    gSL_Select = " SELECT * FROM 名称マスタ "
    gSL_Select = gSL_Select & " WHERE 区分名称 = 'ポケット最大数'"
    gSL_Select = gSL_Select & "   AND 区分タイプ = '' "
    gSL_Select = gSL_Select & "   AND 値 = '1' "
    If Not SYMRead(gSL_Select, 1) Then
        Call ItemsClearSYMCreate
        With SYMCreate
            .区分名称 = "ポケット最大数"
            .区分桁数 = 1
            .値 = "1"
            .値名称 = "ポケット最大数"
            .英数字1 = "70,140,80,80,240"
            .数字1 = 60
            .作成者 = gTanto
        End With
        Call SYMInitCreate(True)
    End If
    kBas_PoketMax = SYM.数字1
    If kBas_PoketMax < 11 Then
       kBas_PoketMax = 11
    End If
'
    kBas_MaxPColor = SYM.数字2              ' 2006/02/03 Insert
    
    ReDim kBas_TKey(kBas_PoketMax)
    ReDim kBas_TKey2(kBas_PoketMax)
    ReDim kBas_TMai(kBas_PoketMax)
'---------------------------------------------------------  2006/02/03 Insert
    ReDim kBas_ErSetMsg(kBas_PoketMax)     ' ｴﾗｰﾎﾟｹｯﾄ固定(名称)
    ReDim kBas_ErSetNo(kBas_PoketMax)      ' ｴﾗｰﾎﾟｹｯﾄ固定(ｴﾗｰ番号)
'---------------------------------------------------------  2006/02/03 Insert End
    kBas_TWrork = Split(SYM.英数字1, ",")
    Erase kBas_TMaiMax
    For gInt = 0 To UBound(kBas_TWrork)
        kBas_TMaiMax(gInt + 1) = CisFun.Val2(kBas_TWrork(gInt))
    Next gInt
'////   発注読取
    If kBas_Proc = "" Or kBas_Proc = "Z" Or kBas_Proc = "H" Then
        
'+---------------------------------------------------------------------------------  2006/02/02 Insert
        ' ｴﾗｰﾎﾟｹｯﾄ情報取得
        Call cKanMeisyoInit(sGroup)
'+---------------------------------------------------------------------------------  2006/02/02 Insert End
        ' かんばん読取初期値セット
        With kBas_Check
            .FTime = Format(Now(), "hhmm")  ' 振出時間
            .Group = 0
            .Op1Value = ""               ' ｵﾌﾟｼｮﾝ設定
            .Op2Value = ""
'+---------------------------------------------------------------------------------  2005/06/18 Insert
            .Op3Value = 20
'+---------------------------------------------------------------------------------  2005/06/18 Insert End
            .ErrCnt = 0                     ' エラー件数
            .OutCnt = 0                     ' 正常出力件数
'+---------------------------------------------------------------------------------  2006/02/03 Insert
            .OpMiryu = 0
'+---------------------------------------------------------------------------------  2006/02/03 Insert End
        End With
    ' オプション設定(発注かんばん)
        ' 平準化月超過発注
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 6
        Call IRNRead("", 1)
        If Mid(IRN.oyian, 1, 1) = "1" Then
           kBas_Check.Op1Value = "X"
        End If
        ' かんばん読取り 即時発注
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 5
        Call IRNRead("", 1)
        If Mid(IRN.oyian, 1, 1) = "1" Then
           kBas_Check.Op2Value = "1"
        End If
'+---------------------------------------------------------------------------------  2005/06/18 Insert
        ' かんばん流動履歴削除日数
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 32
        Call IRNRead("", 1)
        kBas_Check.Op3Value = CisFun.Val2(Mid(IRN.oyian, 1, 3))
'+---------------------------------------------------------------------------------  2005/06/18 Insert End
        '発行元
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 4
        Call IRNRead("", 1)
        Qr_HMoto = RTrim(Mid(IRN.oyian, 1, 7))
        If Trim(Qr_HMoto) = "" Then Qr_HMoto = "SIMPR-A"
'+---------------------------------------------------------------------------------  2005/06/18 Insert
        ' 発注かんばん読取エラーリスト　発行
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 35
        Call IRNRead("", 1)
        kBas_Check.OpHErrPrt = CisFun.Val2(Mid(IRN.oyian, 1, 1))
        ' 発注打切リスト　発行
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 36
        Call IRNRead("", 1)
        kBas_Check.OpHUchiPrt = CisFun.Val2(Mid(IRN.oyian, 1, 1))
'+---------------------------------------------------------------------------------  2005/06/18 Insert End
'+---------------------------------------------------------------------------------  2006/02/03 Insert
        ' 未流動管理
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 37
        Call IRNRead("", 1)
        kBas_Check.OpMiryu = CisFun.Val2(Mid(IRN.oyian, 1, 1))
'+---------------------------------------------------------------------------------  2006/02/03 Insert
        ' 滞留日付設定値・重複削除日付設定値を求める
        gSL_Select = " SELECT * FROM 名称マスタ "
        gSL_Select = gSL_Select & " WHERE 区分名称 = 'データ管理'"
        gSL_Select = gSL_Select & "   AND 区分タイプ = '' "
        gSL_Select = gSL_Select & "   AND 値 = 'KAN' "
        If Not SYMRead(gSL_Select, 1) Then
            Call ItemsClearSYMCreate
            With SYMCreate
                .区分名称 = "データ管理"
                .区分桁数 = 5
                .値 = "KAN"
                .値名称 = "発注かんばん処理削除管理"
                .英数字1 = "滞留期間(稼動日) (1)"
        '+---------------------------------------------------- 2005/06/14 Delete & Insert
'                .英数字2 = "発注かんばん重複削除設定 (2)"          ' Delete
                .英数字2 = "発注かんばん重複削除日数 (2)"           ' Insert
        '+---------------------------------------------------- 2005/06/14 Delete & Insert End
                .英数字3 = "かんばんﾁｪｯｸﾃｰﾌﾞﾙ削除日数(3)"
                .数字1 = 0
        '+---------------------------------------------- 2005/06/14 Delete
'                .数字2 = 1
'                .数字3 = 0
        '+---------------------------------------------- 2005/06/14 Delete End
        '+---------------------------------------------- 2005/06/14 Insert
                .数字2 = 1
                .数字3 = 30
        '+---------------------------------------------- 2005/06/14 Insert End
                .作成者 = gTanto
                .制御区分 = "X"
            End With
            If SYMInitCreate(True) Then
               Call SYMInitCreate(False)
            End If
        '+---------------------------------------------- 2005/06/14 Delete
'            SYM.数字2 = 1
'            SYM.数字3 = 0
        '+---------------------------------------------- 2005/06/14 Delete End
        '+---------------------------------------------- 2005/06/14 Insert
            SYM.数字2 = 1
            SYM.数字3 = 30
        '+---------------------------------------------- 2005/06/14 Insert End
        End If
        kBas_DelDay1 = SYM.数字3 ' かんばんチェックテーブル削除日数
        kBas_DelDay2 = SYM.数字2 ' かんばん重複チェックテーブル削除ﾌﾗｸﾞ　退避
    End If
    ' 部品発注
    If kBas_Proc = "" Or kBas_Proc = "H" Then
        ' 発注かんばん枚数累積テーブル 削除
        With CisDB
            .SQL = "Delete 発注かんばん枚数累積テーブル"
            .SQL = .SQL & "    Where 処理端末 = Host_Name()"
            .SQL = .SQL & "      And 処理区分 = 1"
            .DBExec
        End With
        ' かんばん読取エラーテーブル削除
        With CisDB
            .SQL = "Delete かんばん読取エラーテーブル"
            .SQL = .SQL & " Where かんばん読取エラーテーブル.エラー管理番号"
            .SQL = .SQL & "   = ( Select 管理No From 帳票出力テーブル"
            .SQL = .SQL & "         Where 端末番号 = Host_Name()"
            .SQL = .SQL & "            And 帳票区分 IN (1,2,3,4)"
            .SQL = .SQL & "                  Group by 管理No )"
            .SQL = .SQL & " And かんばん読取エラーテーブル.エラー種類 IN (1,2,3,4) "
            .DBExec
        End With
        ' 帳票出力テーブル削除
        With CisDB
            .SQL = "Delete 帳票出力テーブル"
            .SQL = .SQL & "    Where 端末番号 = Host_Name()"
            .SQL = .SQL & "            And 帳票区分 IN (1,2,3,4)"
            .DBExec
        End With
        TDay = ""
        ' 滞留日付更新(ZEROは更新なしに変更)
        If SYM.数字1 <> 0 Then
            With CisDB
                .SQL = "日付加算"
                .StoadoCount = 4
                .ParaNo = 0: .ParaIO = Return用
                .ParaNo = 1: .ParaIO = Input用
                .ParaNo = 2: .ParaIO = Input用
                .ParaNo = 3: .ParaIO = OutPut用
                .ParaNo = 4: .ParaIO = OutPut用
        
                .ParaNo = 1: .ParaValue = ""
                If SYM.数字1 <> 0 Then SYM.数字1 = SYM.数字1 * -1
                .ParaNo = 2: .ParaValue = SYM.数字1
        
                If .DBStored Then
                   .ParaNo = 4: gLong = CisFun.Val2(.ParaValue)
                   If gLong = 0 Then
                      .ParaNo = 3: TDay = Trim(.ParaValue)
                   End If
                End If
            End With
        ' かんばん流動更新（有効→滞留）
            With CisDB
                .SQL = "Update かんばん流動管理テーブル Set "
                .SQL = .SQL & " 発行区分 = 3,"
                .SQL = .SQL & " 更新日 = GETDATE(),"
                .SQL = .SQL & " 更新者 = '" & gTanto & "',"
                .SQL = .SQL & " 更新端末 = SubString(Host_Name(),1,20)"
                .SQL = .SQL & "    Where 発行区分 = 1"
                .SQL = .SQL & "      And  流動日付 < '" & TDay & "'"
                .SQL = .SQL & "       And  (( 有効指示日 < '" & TDay & "'"
                .SQL = .SQL & "           AND 有効指示日 <> '' )"
                .SQL = .SQL & "              or 有効指示日 = '' )"
                .DBExec
            End With
        End If
        ' 未処理データ存在確認（発注累積テーブル)
        gSL_Select = "SELECT count(*) 件数 FROM 発注かんばん枚数累積テーブル"
        gSL_Select = gSL_Select & " Where 処理端末 = Host_Name()"
        gSL_Select = gSL_Select & "   And 処理区分 = 0"
        Call CHTRead(gSL_Select, 1)
        kBas_Check.OutCnt = CHT.件数
        If kBas_Check.OutCnt <> 0 Then
            With CisFun
                .MB_Lines = 7
                .MB_MSG(1) = "      かんばん読取データ未累積情報が存在します。      "
                .MB_MSG(3) = "              未累積情報の累積処理を行いますか？      "
                .MB_MSG(5) = "          処理を行わない時は、未累積情報の削除を行います。        "
                .MB_Button = Yes_No
                .MB_Title = "未処理確認"
                If Not .MBOX Then
                    ' 発注かんばん枚数累積テーブル 削除
                    With CisDB
                        .SQL = "Delete 発注かんばん枚数累積テーブル"
                        .SQL = .SQL & "    Where 処理端末 = Host_Name()"
                        .DBExec
                    End With
                    kBas_Check.OutCnt = 0
                Else
                    Exit Function
                End If
            End With
        End If
    End If
    ' 材料発注
    If kBas_Proc = "Z" Or kBas_Proc = "H" Then
        ' 発注かんばん枚数累積テーブル 削除
        With CisDB
            .SQL = "Delete 材料発注かんばん枚数累積テーブル"
            .SQL = .SQL & "    Where 処理端末 = Host_Name()"
            .SQL = .SQL & "      And 処理区分 = 1"
            .DBExec
        End With
        ' かんばん読取エラーテーブル削除
        With CisDB
            .SQL = "Delete 材料かんばん読取エラーテーブル"
            .SQL = .SQL & " Where 材料かんばん読取エラーテーブル.エラー管理番号"
            .SQL = .SQL & "   = ( Select 管理No From 帳票出力テーブル"
            .SQL = .SQL & "         Where 端末番号 = Host_Name()"
            .SQL = .SQL & "            And 帳票区分 IN (21,23,24)"
            .SQL = .SQL & "                  Group by 管理No )"
            .SQL = .SQL & " And 材料かんばん読取エラーテーブル.エラー種類 IN (21,23,24) "
            .DBExec
        End With
        ' 帳票出力テーブル削除
        With CisDB
            .SQL = "Delete 帳票出力テーブル"
            .SQL = .SQL & "    Where 端末番号 = Host_Name()"
            .SQL = .SQL & "      And 帳票区分 IN (21,23,24)"
            .DBExec
        End With
        TDay = ""
'        With CisDB
'            .SQL = "日付加算"
'            .StoadoCount = 4
'            .ParaNo = 0: .ParaIO = Return用
'            .ParaNo = 1: .ParaIO = Input用
'            .ParaNo = 2: .ParaIO = Input用
'            .ParaNo = 3: .ParaIO = OutPut用
'            .ParaNo = 4: .ParaIO = OutPut用
'
'            .ParaNo = 1: .ParaValue = ""
'            If SYM.数字1 <> 0 Then SYM.数字1 = SYM.数字1 * -1
'            .ParaNo = 2: .ParaValue = SYM.数字1
'
'            If .DBStored Then
'               .ParaNo = 4: gLong = CisFun.Val2(.ParaValue)
'               If gLong = 0 Then
'                  .ParaNo = 3: TDay = Trim(.ParaValue)
'               End If
'            End If
'        End With
'    ' かんばん流動更新（有効→滞留）
'        With CisDB
'            .SQL = "Update かんばん材料管理テーブル Set "
'            .SQL = .SQL & " 発行区分 = 3,"
'            .SQL = .SQL & " 更新日 = GETDATE(),"
'            .SQL = .SQL & " 更新者 = '" & gTanto & "',"
'            .SQL = .SQL & " 更新端末 = SubString(Host_Name(),1,20)"
'            .SQL = .SQL & "    Where 発行区分 = 1"
'            .SQL = .SQL & "      And  流動日付 < '" & TDay & "'"
'            .SQL = .SQL & "       And  (( 有効指示日 < '" & TDay & "'"
'            .SQL = .SQL & "           AND 有効指示日 <> '' )"
'            .SQL = .SQL & "              or 有効指示日 = '' )"
'            .DBExec
'        End With
        ' 未処理データ存在確認（発注累積テーブル)
        gSL_Select = "SELECT count(*) 件数 FROM 材料発注かんばん枚数累積テーブル"
        gSL_Select = gSL_Select & " Where 処理端末 = Host_Name()"
        gSL_Select = gSL_Select & "   And 処理区分 = 0"
        Call CHTRead(gSL_Select, 1)
        kBas_Check.OutCnt = CHT.件数
        If kBas_Check.OutCnt <> 0 Then
            With CisFun
                .MB_Lines = 7
                .MB_MSG(1) = "      かんばん読取データ未累積情報が存在します。      "
                .MB_MSG(3) = "              未累積情報の累積処理を行いますか？      "
                .MB_MSG(5) = "          処理を行わない時は、未累積情報の削除を行います。        "
                .MB_Button = Yes_No
                .MB_Title = "未処理確認"
                If Not .MBOX Then
                    ' 発注かんばん枚数累積テーブル 削除
                    With CisDB
                        .SQL = "Delete 材料発注かんばん枚数累積テーブル"
                        .SQL = .SQL & "    Where 処理端末 = Host_Name()"
                        .DBExec
                    End With
                    kBas_Check.OutCnt = 0
                Else
                    Exit Function
                End If
            End With
        End If
    End If
    ' 受注
    If kBas_Proc = "J" Then
        ' かんばん重複チェックテーブル 削除
        With CisDB
            .SQL = "Delete かんばん重複チェックテーブル"
            .SQL = .SQL & "    Where 処理区分 = 'J'"
            .DBExec
        End With
        ' 受注かんばん読取テーブル 削除
        With CisDB
            .SQL = "Delete 受注かんばん読取テーブル"
            .DBExec
        End With
        ' かんばん読取エラーテーブル削除
        With CisDB
            .SQL = "Delete かんばん読取エラーテーブル"
            .SQL = .SQL & " Where かんばん読取エラーテーブル.エラー管理番号"
            .SQL = .SQL & "   = ( Select 管理No From 帳票出力テーブル"
            .SQL = .SQL & "         Where 端末番号 = Host_Name()"
            .SQL = .SQL & "            And 帳票区分 = 41"
            .SQL = .SQL & "                  Group by 管理No )"
            .SQL = .SQL & " And かんばん読取エラーテーブル.エラー種類 = 41 "
            .DBExec
        End With
        ' 帳票出力テーブル削除
        With CisDB
            .SQL = "Delete 帳票出力テーブル"
            .SQL = .SQL & "    Where 端末番号 = Host_Name()"
            .SQL = .SQL & "            And 帳票区分 = 41 "
            .DBExec
        End With
'+---------------------------------------------------------------------------------  2005/06/18 Insert
        ' 発注かんばん読取エラーリスト　発行
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 51
        Call IRNRead("", 1)
        kBas_Check.OpJErrPrt = CisFun.Val2(Mid(IRN.oyian, 1, 1))
'+---------------------------------------------------------------------------------  2005/06/18 Insert End
    End If
    cKanbanInit = True
End Function
'--------------------------------------------------------------------------
'*      処理遅延
'--------------------------------------------------------------------------
Public Sub cKanbanLoop()
    Static wLoop As Long
    
    wLoop = 0
    Do Until wLoop > 230000
        wLoop = wLoop + 1
    Loop
End Sub
'--------------------------------------------------------------------------
'*      デモ用データ読取
'--------------------------------------------------------------------------
Public Function GetDemoData() As Boolean
    GetDemoData = False
    
    If kBas_DemoFlg = "S" Then
       kBas_DemoFlg = "R"
       Exit Function
    End If
    
    On Error GoTo GetDemoData_Ed
    
    If kBas_DemoFno = 0 Then
       kBas_DemoFno = FreeFile
       Open kBas_DemoDir & kBas_DemoFName For Input As #kBas_DemoFno
    End If
    
    Line Input #1, kBas_CommData                    ' 行を変数に読み込みます。
    If Len(kBas_CommData) > 20 Then
       kBas_CommData = kBas_CommData & "#"
    End If

    If Mid(kBas_CommData, 1, 1) = "?" Then kBas_CommData = Chr(&H1B) & "?"
    If Mid(kBas_CommData, 1, 1) = "E" Then kBas_CommData = Chr(&H1B) & "E"
    If Mid(kBas_CommData, 1, 1) = "P" Then kBas_CommData = Chr(&H1B) & "P"
    If Mid(kBas_CommData, 1, 1) = "J" Then kBas_CommData = Chr(&H1B) & "J"
    If Mid(kBas_CommData, 1, 1) = "D" Then kBas_CommData = Chr(&H1B) & "D"
    If Mid(kBas_CommData, 1, 1) = "C" Then kBas_CommData = Chr(&H1B) & "C"
    If Mid(kBas_CommData, 1, 1) = "*" Then kBas_CommData = Chr(&H1B) & "*"
    
    GetDemoData = True
    Exit Function
GetDemoData_Ed:
    kBas_CommData = Chr(&H1B) & "E"
' デモ用設定 ファイルクローズ
    If kBas_DemoFno <> 0 Then
       Close kBas_DemoFno
    End If
    kBas_DemoFno = 0
End Function
'------------------------------------------------------------------- 2006/02/03 Insert
'+-------------------------------------------------------+
'+          エラーポケット内容(発注)　作成 & 取得
'+-------------------------------------------------------+
Private Sub cKanMeisyoInit(sGroup As String)
    gSL_Select = " SELECT * FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = 'ポケットグループ発注'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    If Not SYMRead(gSL_Select, 1) Then
        Call ItemsClearSYMCreate
        With SYMCreate
            .区分名称 = "ポケットグループ発注"
            .区分桁数 = 2
            .値 = "01"
            .英数字1 = "指定ﾎﾟｹｯﾄNo"
            .数字1 = 1
            .値名称 = "読取不能"
            .作成者 = gTanto
        End With
        If Not SYMInitCreate(True) Then
            For gInt = 2 To 6
                With SYMCreate
                    .値 = Format(gInt, "00")
                    Select Case gInt
                         Case 2: .値名称 = "重複": .数字1 = 2
                         Case 3: .値名称 = "ﾁｪｯｸｴﾗｰ": .数字1 = 3
                         Case 4: .値名称 = "抜き取り": .数字1 = 4
                         Case 5: .値名称 = "ﾎﾟｹｯﾄｵｰﾊﾞｰ": .数字1 = 5
                         Case 6: .値名称 = "ｸﾞﾙｰﾌﾟ外": .数字1 = 6: .数字2 = 1
                    End Select
                End With
                Call SYMInitCreate
            Next gInt
        End If
    End If
' かんばん読取エラーポケット
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "かんばんエラー設定発注"
        .区分桁数 = 2
        .値 = "01"
        .値名称 = "読取不能"
        .数字2 = 1
        .英数字1 = "1:仕入先ｸﾞﾙｰﾌﾟ"
        .英数字2 = "ｴﾗｰｸﾞﾙｰﾌﾟNo"
        .英数字3 = "ポケットNo"
        .作成者 = gTanto
    End With
    If Not SYMInitCreate(True) Then
        For gInt = 2 To 10
            With SYMCreate
                .値 = Format(gInt, "00")
                Select Case gInt
                     Case 2: .値名称 = "重複読取": .数字2 = 2: .数字3 = 0
                     Case 3: .値名称 = "ﾁｪｯｸｴﾗｰ": .数字2 = 3: .数字3 = 0
                     Case 4: .値名称 = "抜き取り": .数字2 = 4: .数字3 = 0
                     Case 5: .値名称 = "抜き取り(減)": .数字2 = 4: .数字3 = 0
                     Case 6: .値名称 = "打切(完了)": .数字2 = 3: .数字3 = 0
                     Case 7: .値名称 = "打切(端数)": .数字2 = 3: .数字3 = 0
                     Case 8: .値名称 = "平準化": .数字2 = 3: .数字3 = 0
                     Case 9: .値名称 = "ﾎﾟｹｯﾄｵｰﾊﾞｰ": .数字2 = 5: .数字3 = 0
                     Case 10: .値名称 = "グループ外": .数字1 = 1: .数字2 = 6: .数字3 = 0
                End Select
            End With
            Call SYMInitCreate
        Next gInt
    End If
    
    Call cKanErrPoket("発注", sGroup)

End Sub
'+-------------------------------------------------------+
'+          エラーポケット情報取得
'+-------------------------------------------------------+
Private Sub cKanErrPoket(Kubun As String, Optional sGroup As String = "")

    gSL_Select = " SELECT * FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = 'ポケットグループ" & Kubun & "'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    If sGroup = "" Then
       gSL_Select = gSL_Select & "   And 数字2 = 0"
    End If
    If SYMRead(gSL_Select) Then
       Do Until Not SYM_RDSTS
          If SYM.数字1 > 0 And SYM.数字1 <= kBas_PoketMax Then
             kBas_ErSetNo(SYM.数字1) = CisFun.Val2(SYM.値)
             kBas_ErSetMsg(SYM.数字1) = RTrim(SYM.値名称)
          End If
          If SYM.数字1 > kBas_PoketMax Then
             kBas_PoketErr = SYM.数字1
             kBas_PoketErrMsg = Trim(SYM.値名称)
          End If
                  
          Call SYMReadNext
       Loop
    End If
    gSL_Select = " SELECT Max(値) 値 FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = 'かんばんエラー設定" & Kubun & "'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    If SYMRead(gSL_Select, 1) Then
       gLong = CisFun.Val2(SYM.値)
       ReDim kBas_ErNo(gLong)       ' ｴﾗｰ番号
       ReDim kBas_ErPno(gLong)      ' ｴﾗｰﾎﾟｹｯﾄ
       ReDim kBas_ErMsg(gLong)      ' ｴﾗｰ内容
       
       gLong = 1
       gSL_Select = " SELECT * FROM 名称マスタ"
       gSL_Select = gSL_Select & " WHERE 区分名称 = 'かんばんエラー設定" & Kubun & "'"
       gSL_Select = gSL_Select & "   And 区分タイプ = ''"
       If sGroup = "" Then
          gSL_Select = gSL_Select & "   And 数字1 = 0"
       End If
       If SYMRead(gSL_Select) Then
          Do Until Not SYM_RDSTS
             kBas_ErNo(gLong) = CisFun.Val2(SYM.値)     ' ｴﾗｰ番号
             kBas_ErMsg(gLong) = Trim(SYM.値名称)       ' ｴﾗｰ内容
             If CisFun.Val2(SYM.数字2) > 0 And CisFun.Val2(SYM.数字2) <= kBas_PoketMax Then
                If CisFun.Val2(SYM.数字2) <> 0 Then
                   For gInt = 1 To UBound(kBas_ErSetMsg)
                       If kBas_ErSetNo(gInt) = CisFun.Val2(SYM.数字2) Then
                          kBas_ErPno(gLong) = gInt
                          Exit For
                       End If
                   Next gInt
                End If
             End If
             If kBas_ErPno(gLong) = 0 Then
                If CisFun.Val2(SYM.数字3) > 0 And CisFun.Val2(SYM.数字3) <= kBas_PoketMax Then
                   kBas_ErPno(gLong) = SYM.数字3 ' ｴﾗｰﾎﾟｹｯﾄ
                   If kBas_ErSetMsg(SYM.数字3) = "" Then
                      kBas_ErSetMsg(SYM.数字3) = Trim(SYM.値名称)
                   End If
                End If
                If CisFun.Val2(SYM.数字3) > kBas_PoketMax Then
                   kBas_ErPno(0) = SYM.数字3 ' ｴﾗｰﾎﾟｹｯﾄ
                   If kBas_ErSetMsg(0) = "" Then
                      kBas_ErSetMsg(0) = Trim(SYM.値名称)
                   End If
                End If
             End If
             
             Call SYMReadNext
             gLong = gLong + 1
          Loop
       End If
    End If
    Call SYMClose

End Sub
'------------------------------------------------------------------- 2006/02/03 Insert End
'--------------------------------------------------------------------------
'*      かんばんソータ  読取データ編集
'*              バーコード種類 … 0 :
'*              ステータス     … 0 : 読取正常  1 : 読取ｴﾗｰ
'*                               2 : 自社外かんばん(対象外かんばん)
'*                               3 : ?? ﾁｪｯｸﾃﾞｨｼﾞｯﾄｴﾗｰ
'--------------------------------------------------------------------------
Public Function cKanbanValueSet() As Boolean
    cKanbanValueSet = False
    ' かんばん読取情報初期化
    Call QRAreaClear

    If Trim(kBas_CommData) = "" Then Exit Function
' 読取ﾃﾞｰﾀ分割
    DCnt = InStr(1, kBas_CommData, "#")
    
    If DCnt <> 0 Or Mid(kBas_CommData, 1, 1) <> Chr(&H1B) Then
        If DCnt = 0 Then
           DCnt = Len(kBas_CommData) + 1
        End If
    Else
        '----------------------------------------- 2006/03/24 Insert
        If Mid(kBas_CommData, 2, 1) = "?" Then
           QRArea.ステータス = 1
           cKanbanValueSet = True
           Exit Function
        End If
        '----------------------------------------- 2006/03/24 Insert End
        Exit Function
    End If
    
    kBas_ReadData(0) = Mid(kBas_CommData, 1, DCnt - 1)
    kBas_ReadData(1) = Mid(kBas_CommData, DCnt + 1)
'両面読取り かんばん編集
    QRChk_SimplrA = ""
    QRChk_SimplrABar = ""
    
    For DId = 0 To 1
        If Mid(kBas_ReadData(DId), 1, 1) = Chr(&H1B) And _
           Mid(kBas_ReadData(DId), 2, 1) = "?" Then
           kBas_ReadData(DId) = "?"
        Else
           Call KanbanEdit(kBas_ReadData(DId))
        End If
    Next DId

    If Trim(QRChk_SimplrA) <> "" Then
       Call QR_SimprA(QRChk_SimplrA)
       cKanbanValueSet = True
       Exit Function
    End If
    If Trim(QRChk_SimplrABar) <> "" Then
       If Not QR_SimprABar(QRChk_SimplrABar) Then
          QRArea.ステータス = 1
          Exit Function
       Else
          cKanbanValueSet = True
          Exit Function
       End If
    End If
    ' 読取エラーかんばん
    If kBas_ReadData(0) = "?" And kBas_ReadData(1) = "?" Or _
       kBas_ReadData(0) = "?" And kBas_ReadData(1) = "" Or _
       kBas_ReadData(0) = "" And kBas_ReadData(1) = "?" Then
        QRArea.ステータス = 1
        Exit Function
    End If
    ' 処理対象外かんばん
    QRArea.ステータス = 2
End Function
'--------------------------------------------------------------------------
'*      かんばん手差し  読取データ編集
'*              バーコード種類 … 0 :
'*              ステータス     … 0 : 読取正常  1 : 読取ｴﾗｰ
'*                               2 : 自社外かんばん(対象外かんばん)
'*                               3 : ?? ﾁｪｯｸﾃﾞｨｼﾞｯﾄｴﾗｰ
'--------------------------------------------------------------------------
Public Sub cKanbanValueSetT()
    ' かんばん読取情報初期化
    Call QRAreaClear
    QRChk_SimplrA = ""
    QRChk_SimplrABar = ""
    
    Call KanbanEdit(kBas_CommData)
    If Trim(QRChk_SimplrA) <> "" Then
       Call QR_SimprA(QRChk_SimplrA)
       Exit Sub
    End If
    If Trim(QRChk_SimplrABar) <> "" Then
       If QR_SimprABar(QRChk_SimplrABar) Then
          Exit Sub
       End If
    End If
    ' 処理対象外かんばん
    QRArea.ステータス = 2
        
End Sub
'+------------------------------------+
'+      発行元確認
'+------------------------------------+
Private Sub KanbanEdit(ReadData As String)

    kBas_ChkData = Replace(ReadData, vbCr, "")
    
    ' Simpr-A 発行かんばん(自社発行）
    If RTrim(CisFun.Mid2(ReadData, 1, 7)) = Qr_HMoto Then
        QRChk_SimplrA = ReadData
    End If
    ' Simpr-A 発行かんばん(縦バー)
    If Len(kBas_ChkData) = 11 Then
        QRChk_SimplrABar = kBas_ChkData
    End If

End Sub

'//////////////////////////////////////
'//
'//     発注かんばん読取チェック & 更新
'//
'//////////////////////////////////////
Public Function cKanbanChkH() As Boolean
    cKanbanChkH = False
    
    On Error GoTo cKanbanChkH_ER
    gStr = ""
    With CisDB
        .SQL = "発注かんばん処理"
'        .StoadoCount = 19               2006/02/03 Delete
        .StoadoCount = 21
        
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用      ' 処理区分 1:一括,2:指定
        .ParaNo = 2: .ParaIO = Input用      ' 発注区分 '':通常,1:臨時,2:納期指定(処理区分='2'のみ)
        .ParaNo = 3: .ParaIO = Input用      ' かんばん区分
        .ParaNo = 4: .ParaIO = Input用      ' 品番
        .ParaNo = 5: .ParaIO = Input用      ' 背番号
        .ParaNo = 6: .ParaIO = Input用      ' 仕入先
        .ParaNo = 7: .ParaIO = Input用      ' 受入
        .ParaNo = 8: .ParaIO = Input用      ' サイクル
        .ParaNo = 9: .ParaIO = Input用      ' 連番
        .ParaNo = 10: .ParaIO = Input用     ' 収容数
        .ParaNo = 11: .ParaIO = Input用     ' 振出時間(処理区分='1'のみ)
        .ParaNo = 12: .ParaIO = Input用     ' 読取ｸﾞﾙｰﾌﾟ(処理区分='1'のみ)
        .ParaNo = 13: .ParaIO = Input用     ' オプション設定
                                            '   平準化超過の扱い _/発注なし,1/発注あり
'--------------------------------------------------------- 2006/02/03 Insert
        .ParaNo = 14: .ParaIO = Input用     ' オプション設定(未流動管理)
        .ParaNo = 15: .ParaIO = Input用     ' 納入日 指定で通常の時は、発注日を指定する
        .ParaNo = 16: .ParaIO = Input用     ' 納入便 指定で通常の時は、発注便を指定する
        
        .ParaNo = 17: .ParaIO = OutPut用    ' 指定ポケット№
        .ParaNo = 18: .ParaIO = OutPut用    ' エラー番号
        .ParaNo = 19: .ParaIO = OutPut用    ' エラー出力件数
        .ParaNo = 20: .ParaIO = OutPut用    ' 正常出力件数
        .ParaNo = 21: .ParaIO = Input用     ' 版数          2006/12/20追加

        .ParaNo = 1: .ParaValue = kBas_Check.ProcKB
        .ParaNo = 2: .ParaValue = kBas_Check.HCKB
        .ParaNo = 3: .ParaValue = QRArea.手配区分
        .ParaNo = 4: .ParaValue = QRArea.品番
        .ParaNo = 5: .ParaValue = QRArea.背番号
        .ParaNo = 6: .ParaValue = QRArea.取引先
        .ParaNo = 7: .ParaValue = QRArea.受入
        .ParaNo = 8: .ParaValue = QRArea.サイクル
        .ParaNo = 9: .ParaValue = QRArea.連番
        .ParaNo = 10: .ParaValue = QRArea.収容数
        .ParaNo = 11: .ParaValue = kBas_Check.FTime
        .ParaNo = 12: .ParaValue = kBas_Check.Group
        .ParaNo = 13: .ParaValue = kBas_Check.Op1Value
        .ParaNo = 14: .ParaValue = kBas_Check.OpMiryu
        
        .ParaNo = 15: .ParaValue = kBas_Check.HYMD
        .ParaNo = 16: .ParaValue = kBas_Check.HBIN
        
        .ParaNo = 21: .ParaValue = QRArea.版数      ' 版数          2006/12/20追加
        
        If Not .DBStored Then GoTo cKanbanChkH_ER
        
        ' ストアド処理エラー
        .ParaNo = 0: kBas_StoRTN = CisFun.Val2(.ParaValue)
        If kBas_StoRTN <> 0 Then GoTo cKanbanChkH_ER
        
        ' 指定ポケット番号
        .ParaNo = 17: kBas_Poket = CisFun.Val2(.ParaValue)
        ' エラー番号
        .ParaNo = 18: kBas_ErrNo = CisFun.Val2(.ParaValue)
        ' エラー出力件数
        .ParaNo = 19: kBas_Check.ErrCnt = kBas_Check.ErrCnt + CisFun.Val2(.ParaValue)
        ' 正常出力件数
        .ParaNo = 20: kBas_Check.OutCnt = kBas_Check.OutCnt + CisFun.Val2(.ParaValue)
        
'--------------------------------------------------------- 2006/02/03 Insert End
'--------------------------------------------------------- 2006/02/03 Delete
'        .ParaNo = 14: .ParaIO = Input用     ' 納入日 指定で通常の時は、発注日を指定する
'        .ParaNo = 15: .ParaIO = Input用     ' 納入便 指定で通常の時は、発注便を指定する
'
'        .ParaNo = 16: .ParaIO = OutPut用    ' 指定ポケット№
'        .ParaNo = 17: .ParaIO = OutPut用    ' エラー番号
'        .ParaNo = 18: .ParaIO = OutPut用    ' エラー出力件数
'        .ParaNo = 19: .ParaIO = OutPut用    ' 正常出力件数
'
'        .ParaNo = 1: .ParaValue = kBas_Check.ProcKB
'        .ParaNo = 2: .ParaValue = kBas_Check.HCKB
'        .ParaNo = 3: .ParaValue = QRArea.手配区分
'        .ParaNo = 4: .ParaValue = QRArea.品番
'        .ParaNo = 5: .ParaValue = QRArea.背番号
'        .ParaNo = 6: .ParaValue = QRArea.取引先
'        .ParaNo = 7: .ParaValue = QRArea.受入
'        .ParaNo = 8: .ParaValue = QRArea.サイクル
'        .ParaNo = 9: .ParaValue = QRArea.連番
'        .ParaNo = 10: .ParaValue = QRArea.収容数
'        .ParaNo = 11: .ParaValue = kBas_Check.FTime
'        .ParaNo = 12: .ParaValue = kBas_Check.Group
'        .ParaNo = 13: .ParaValue = kBas_Check.Op1Value

'        .ParaNo = 14: .ParaValue = kBas_Check.HYMD
'        .ParaNo = 15: .ParaValue = kBas_Check.HBIN
'
'
'        If Not .DBStored Then GoTo cKanbanChkH_ER
'
'        ' ストアド処理エラー
'        .ParaNo = 0: kBas_StoRTN = CisFun.Val2(.ParaValue)
'        If kBas_StoRTN <> 0 Then GoTo cKanbanChkH_ER
'
'        ' 指定ポケット番号
'        .ParaNo = 16: kBas_Poket = CisFun.Val2(.ParaValue)
'        ' エラー番号
'        .ParaNo = 17: kBas_ErrNo = CisFun.Val2(.ParaValue)
'        ' エラー出力件数
'        .ParaNo = 18: kBas_Check.ErrCnt = kBas_Check.ErrCnt + CisFun.Val2(.ParaValue)
'        ' 正常出力件数
'        .ParaNo = 19: kBas_Check.OutCnt = kBas_Check.OutCnt + CisFun.Val2(.ParaValue)
'--------------------------------------------------------- 2006/02/03 Delete End
        
    End With

    cKanbanChkH = True
    On Error GoTo 0
    Exit Function
cKanbanChkH_ER:
    kBas_ErrMsg = "発注かんばん処理エラー"
    Select Case kBas_StoRTN
        Case 1: gStr = "かんばん重複ﾁｪｯｸﾃｰﾌﾞﾙ 作成エラー"
        Case 2, 3: gStr = "便算出エラー"
        Case 10: gStr = "ｴﾗｰ管理番号　取得ｴﾗｰ"
        Case 21, 22, 23, 24: gStr = "かんばんチェックテーブル　作成エラー"
        Case 26, 27, 28, 29: gStr = "かんばんチェックテーブル　更新エラー"
        
        Case 31, 32, 33, 34: gStr = "かんばん読取ｴﾗｰﾃｰﾌﾞﾙ　作成エラー"
        Case 36, 37: gStr = "かんばん読取ｴﾗｰﾃｰﾌﾞﾙ　更新エラー"
        Case 38: gStr = "かんばん読取ｴﾗｰﾃｰﾌﾞﾙ　削除エラー"
        
        Case 41, 42, 43, 44: gStr = "帳票出力テーブル 作成エラー"
        
        Case 51, 52: gStr = "かんばん流動管理ﾃｰﾌﾞﾙ 更新エラー"
        
        Case 53, 56: gStr = "かんばん増減テーブル 更新エラー"
        
        Case 54, 55: gStr = "発注打切テーブル 更新エラー"
        
        Case 61: gStr = "発注かんばん枚数累積ﾃｰﾌﾞﾙ 作成エラー"
        Case 62: gStr = "発注かんばん枚数累積ﾃｰﾌﾞﾙ 更新エラー"
        Case 63: gStr = "かんばん流動管理ﾃｰﾌﾞﾙ 出力エラー"
        Case Else
             gStr = "その他ｴﾗｰ:" & gStr
    End Select
    kBas_ErrMsg = StrConv(Format(kBas_StoRTN, "00"), vbWide) & ":" & gStr
    kBas_Poket = 1
End Function

'/////////////////////////////////////////////
'//
'//     材料発注かんばん読取チェック & 更新
'//
'/////////////////////////////////////////////
Public Function cKanbanChkHZ() As Boolean
    cKanbanChkHZ = False
    
    On Error GoTo cKanbanChkHZ_ER
    gStr = ""
    With CisDB
        .SQL = "材料発注かんばん処理"
        .StoadoCount = 19
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用      ' 処理区分 1:一括,2:指定
        .ParaNo = 2: .ParaIO = Input用      ' 発注区分 '':通常,1:臨時,2:納期指定(処理区分='2'のみ)
        .ParaNo = 3: .ParaIO = Input用      ' かんばん区分
        .ParaNo = 4: .ParaIO = Input用      ' 材料管理番号
        .ParaNo = 5: .ParaIO = Input用      ' 背番号
        .ParaNo = 6: .ParaIO = Input用      ' 仕入先
        .ParaNo = 7: .ParaIO = Input用      ' 受入
        .ParaNo = 8: .ParaIO = Input用      ' サイクル
        .ParaNo = 9: .ParaIO = Input用      ' 連番
        .ParaNo = 10: .ParaIO = Input用     ' 収容数
        .ParaNo = 11: .ParaIO = Input用     ' 振出時間(処理区分='1'のみ)
        .ParaNo = 12: .ParaIO = Input用     ' 読取ｸﾞﾙｰﾌﾟ(処理区分='1'のみ)
        .ParaNo = 13: .ParaIO = Input用     ' オプション設定
                                            '   平準化超過の扱い _/発注なし,1/発注あり
        .ParaNo = 14: .ParaIO = Input用     ' 納入日 指定で通常の時は、発注日を指定する
        .ParaNo = 15: .ParaIO = Input用     ' 納入便 指定で通常の時は、発注便を指定する
        
        .ParaNo = 16: .ParaIO = OutPut用    ' 指定ポケット№
        .ParaNo = 17: .ParaIO = OutPut用    ' エラー番号
        .ParaNo = 18: .ParaIO = OutPut用    ' エラー出力件数
        .ParaNo = 19: .ParaIO = OutPut用    ' 正常出力件数

        .ParaNo = 1: .ParaValue = kBas_Check.ProcKB
        .ParaNo = 2: .ParaValue = kBas_Check.HCKB
        .ParaNo = 3: .ParaValue = QRArea.手配区分
        .ParaNo = 4: .ParaValue = QRArea.材料管理番号
        .ParaNo = 5: .ParaValue = QRArea.背番号
        .ParaNo = 6: .ParaValue = QRArea.取引先
        .ParaNo = 7: .ParaValue = QRArea.受入
        .ParaNo = 8: .ParaValue = QRArea.サイクル
        .ParaNo = 9: .ParaValue = QRArea.連番
        .ParaNo = 10: .ParaValue = QRArea.材料収容数
        .ParaNo = 11: .ParaValue = kBas_Check.FTime
        .ParaNo = 12: .ParaValue = kBas_Check.Group
        .ParaNo = 13: .ParaValue = kBas_Check.Op1Value
        
        .ParaNo = 14: .ParaValue = kBas_Check.HYMD
        .ParaNo = 15: .ParaValue = kBas_Check.HBIN
        

        If Not .DBStored Then GoTo cKanbanChkHZ_ER
        
        ' ストアド処理エラー
        .ParaNo = 0: kBas_StoRTN = CisFun.Val2(.ParaValue)
        If kBas_StoRTN <> 0 Then GoTo cKanbanChkHZ_ER
        
        ' 指定ポケット番号
        .ParaNo = 16: kBas_Poket = CisFun.Val2(.ParaValue)
        ' エラー番号
        .ParaNo = 17: kBas_ErrNo = CisFun.Val2(.ParaValue)
        ' エラー出力件数
        .ParaNo = 18: kBas_Check.ErrCntZ = kBas_Check.ErrCntZ + CisFun.Val2(.ParaValue)
        ' 正常出力件数
        .ParaNo = 19: kBas_Check.OutCntZ = kBas_Check.OutCntZ + CisFun.Val2(.ParaValue)
        
    End With

    cKanbanChkHZ = True
    On Error GoTo 0
    Exit Function
cKanbanChkHZ_ER:
    kBas_ErrMsg = "材料発注かんばん処理エラー"
    Select Case kBas_StoRTN
        Case 1: gStr = "かんばん重複ﾁｪｯｸﾃｰﾌﾞﾙ 作成エラー"
        Case 2, 3: gStr = "便算出エラー"
        Case 10: gStr = "ｴﾗｰ管理番号　取得ｴﾗｰ"
        Case 21, 22, 23, 24: gStr = "かんばんチェックテーブル　作成エラー"
        Case 26, 27, 28, 29: gStr = "かんばんチェックテーブル　更新エラー"
        
        Case 31, 32, 33, 34: gStr = "かんばん読取ｴﾗｰﾃｰﾌﾞﾙ　作成エラー"
        Case 36, 37: gStr = "かんばん読取ｴﾗｰﾃｰﾌﾞﾙ　更新エラー"
        Case 38: gStr = "かんばん読取ｴﾗｰﾃｰﾌﾞﾙ　削除エラー"
        
        Case 41, 42, 43, 44: gStr = "帳票出力テーブル 作成エラー"
        
        Case 51, 52: gStr = "かんばん材料管理ﾃｰﾌﾞﾙ 更新エラー"
        
        Case 53, 56: gStr = "かんばん増減テーブル 更新エラー"
        
        Case 54, 55: gStr = "材料発注打切テーブル 更新エラー"
        
        Case 61: gStr = "材料発注かんばん枚数累積ﾃｰﾌﾞﾙ 作成エラー"
        Case 62: gStr = "材料発注かんばん枚数累積ﾃｰﾌﾞﾙ 更新エラー"
        Case Else
             gStr = "その他ｴﾗｰ:" & gStr
    End Select
    kBas_ErrMsg = StrConv(Format(kBas_StoRTN, "00"), vbWide) & ":" & gStr
    kBas_Poket = 1
End Function


'//////////////////////////////////////
'//
'//     受注かんばん読取チェック & 更新
'//
'//////////////////////////////////////
Public Function cKanbanChkJ() As Boolean
    cKanbanChkJ = False
    
    On Error GoTo cKanbanChkJ_ER
    gStr = ""
    With CisDB
        .SQL = "受注かんばん処理"
        .StoadoCount = 12
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用      ' 受注区分
        .ParaNo = 2: .ParaIO = Input用      ' 納入日
        .ParaNo = 3: .ParaIO = Input用      ' 納入便
        .ParaNo = 4: .ParaIO = Input用      ' 品番
        .ParaNo = 5: .ParaIO = Input用      ' 背番号
        .ParaNo = 6: .ParaIO = Input用      ' 納入先
        .ParaNo = 7: .ParaIO = Input用      ' 受入
        .ParaNo = 8: .ParaIO = Input用      ' 収容数
        .ParaNo = 9: .ParaIO = Input用      ' サイクル
        .ParaNo = 10: .ParaIO = Input用      ' 連番
        
        .ParaNo = 11: .ParaIO = OutPut用    ' エラー番号
        .ParaNo = 12: .ParaIO = OutPut用    ' 指定ポケット№

        .ParaNo = 1: .ParaValue = CisFun.Val2(kBas_Check.HCKB)
        .ParaNo = 2: .ParaValue = kBas_Check.HYMD
        .ParaNo = 3: .ParaValue = kBas_Check.HBIN
        .ParaNo = 4: .ParaValue = QRArea.品番
        .ParaNo = 5: .ParaValue = QRArea.背番号
        .ParaNo = 6: .ParaValue = QRArea.取引先
        .ParaNo = 7: .ParaValue = QRArea.受入
        .ParaNo = 8: .ParaValue = QRArea.収容数
        .ParaNo = 9: .ParaValue = QRArea.サイクル
        .ParaNo = 10: .ParaValue = QRArea.連番

        If Not .DBStored Then GoTo cKanbanChkJ_ER
        
        ' ストアド処理エラー
        .ParaNo = 0: kBas_StoRTN = CisFun.Val2(.ParaValue)
        If kBas_StoRTN <> 0 Then GoTo cKanbanChkJ_ER
        
        ' 指定ポケット番号
        .ParaNo = 12: kBas_Poket = CisFun.Val2(.ParaValue)
        ' エラー番号
        .ParaNo = 11: kBas_ErrNo = CisFun.Val2(.ParaValue)
        
    End With

    cKanbanChkJ = True
    On Error GoTo 0
    Exit Function
cKanbanChkJ_ER:
    kBas_ErrMsg = "受注かんばん処理エラー"
    Select Case kBas_StoRTN
        Case 1: gStr = "かんばん重複ﾁｪｯｸﾃｰﾌﾞﾙ 作成エラー"
        Case 2: gStr = "受注かんばん読取ﾃｰﾌﾞﾙ 作成ｴﾗｰ"
        Case 3: gStr = "受注かんばん読取ﾃｰﾌﾞﾙ 更新ｴﾗｰ"
        Case 10: gStr = "ｴﾗｰ管理番号　取得ｴﾗｰ"
        Case 11: gStr = "受注管理番号　取得ｴﾗｰ"
        Case 31: gStr = "かんばん読取ｴﾗｰﾃｰﾌﾞﾙ　作成エラー"
        Case 41: gStr = "帳票出力テーブル 作成エラー"
        Case Else
             gStr = "その他ｴﾗｰ:" & gStr
    End Select
    kBas_ErrMsg = StrConv(Format(kBas_StoRTN, "00"), vbWide) & ":" & gStr
    kBas_Poket = 1
End Function

'////////////////////////////////////////////////
'//
'//     既存かんばん対応
'//
'//     かんばん流動管理テーブル作成（有効）
'//
'////////////////////////////////////////////////
Public Function cKanbanOld() As Boolean
    cKanbanOld = False
    If QRArea.バーコード種類 <> 1 Then
       cKanbanOld = True
       Exit Function
    End If
    
    gSL_Select = "Select * From かんばん流動管理テーブル"
    gSL_Select = gSL_Select & " Where 取引先 = '" & QRArea.取引先 & "'"
    gSL_Select = gSL_Select & " And   受入 = '" & QRArea.受入 & "'"
    gSL_Select = gSL_Select & " And   品番 = '" & QRArea.品番 & "'"
    gSL_Select = gSL_Select & " And   発行NO = " & QRArea.連番 & ""
    If Not KRTRead(gSL_Select, 1) Then
       Call ItemsClearKRT
       With KRT
            .取引先 = QRArea.取引先
            .受入 = QRArea.受入
            .品番 = QRArea.品番
            .発行NO = QRArea.連番
            .版数 = 1
            .発行区分 = 0
            .収容区分 = 1
       End With
       If Not KRTInsert Then GoTo cKanbanOld_Er
       
       With CisDB
            .SQL = "Update 品番取引先マスタ Set"
            .SQL = .SQL & " 最終発行連番 = "
            .SQL = .SQL & "   Case When 最終発行連番 < " & QRArea.連番
            .SQL = .SQL & "        Then " & QRArea.連番 & " Else 最終発行連番 End "
            .SQL = .SQL & " Where 取引先 = '" & QRArea.取引先 & "'"
            .SQL = .SQL & " And   受入 = '" & QRArea.受入 & "'"
            .SQL = .SQL & " And   品番 = '" & QRArea.品番 & "'"
                    
            If Not .DBExec Then GoTo cKanbanOld_Er
       End With
    End If
    
    
    cKanbanOld = True
    Exit Function
cKanbanOld_Er:
    kBas_ErrMsg = "既存かんばん対応ｴﾗｰ：流動管理ﾃｰﾌﾞﾙ作成ｴﾗｰ"
    kBas_Poket = 1
End Function

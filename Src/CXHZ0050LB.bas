Attribute VB_Name = "CXHZ0050LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   発注打切対象品番リスト印刷
'*
'*   モジュールID  : CXHZ0050LB
'*
'*                   CXHZ0050LPrint(Ukbn, Hinbn, Torcd)
'*
'*   引数 : Ukbn     : 打切種別
'*   　　   Hinbn    : 品番
'*   　　   Torcd    : 取引先
'*   　　   PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  :
'*   　　   False : 印刷データ無し
'*
'*   使用例:
'
'    If Not CXHZ0050LPrint Then Exit Sub
'
'*   作成日  :2004/02/12   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCXHZ0050        As New CISFormPrint
Public CisSimplePrint       As New CisSimplePrint
Public CXHZ0050_PrtCnt       As Long
Public cUKbnm               As String
Public cUYmd                As String
Public cKYmd                As String
'
Public Function CXHZ0050LPrint(Ukbn As String, _
                             ZKanriNo As String, _
                             Torcd As String, Optional PrintOut As Integer = 0) As Boolean
    CXHZ0050LPrint = False

'   材料取引先マスタ
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT ZTM.材質,ZTM.板厚,ZTM.幅,ZTM.長さ,"
    gSL_Select = gSL_Select & " ZTM.取引先,ZTM.受入,ZTM.背番号,"
    gSL_Select = gSL_Select & " TRM.略称 手配先名,ZRM.表示寸法,"
    gSL_Select = gSL_Select & "(SELECT COUNT(*) FROM 材料品番マスタ ZH"
    gSL_Select = gSL_Select & " WHERE ZH.材料管理番号 = ZTM.材料管理番号)件数"
    gSL_Select = gSL_Select & " FROM"
    gSL_Select = gSL_Select & "  材料取引先マスタ ZTM"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ TRM"
    gSL_Select = gSL_Select & "    on TRM.取引先CD = ZTM.取引先"
    gSL_Select = gSL_Select & "   and TRM.取引先区分 = 1"
    gSL_Select = gSL_Select & "  left outer join 材料マスタ ZRM"
    gSL_Select = gSL_Select & "    on ZRM.材料管理番号 = ZTM.材料管理番号"
    If Ukbn <> "3" Then
        gSL_Select = gSL_Select & " where ZTM.材料管理番号 = " & Val(ZKanriNo) & ""
        If Ukbn = "2" Then
            gSL_Select = gSL_Select & " and ZTM.取引先 = '" & RTrim(Torcd) & "'"
        End If
    Else
        gSL_Select = gSL_Select & " where ZTM.取引先 = '" & RTrim(Torcd) & "'"
    End If
    gSL_Select = gSL_Select & " order by ZTM.取引先,ZTM.受入,ZTM.材質,ZRM.表示寸法"

    If Not ZTMRead(gSL_Select, 0, 1) Then
        Call ZTMClose
        Exit Function
    End If
    CXHZ0050_PrtCnt = CisDB.RecordCount
    With CisSimplePrint
        .MaxCount = CisDB.RecordCount
        .ProgressInit
    End With

    If PrintOut <> 2 Then
        With CisPrtCXHZ0050
             Set .iPrtForm = CXHZ0050L
             .dpMSG = "発注打切対象品番リスト　印刷中"
             Select Case PrintOut
                Case 0: .iPreview = pv_Print
                Case 1: .iPreview = pv_Preview
             End Select
             .dpMSGScreen = モーダル
             .PrintStart
        End With
    Else
        Call CXHZ0050_CsvOut
    End If

    CXHZ0050LPrint = True
End Function
'+----------------------------------------------------------+
'+         CSVファイル出力  処理                        ＋
'+----------------------------------------------------------+
Public Function CXHZ0050_CsvOut() As Boolean
    CXHZ0050_CsvOut = False
    Dim cFNo           As Integer
    Dim PrtHead         As String
    Dim CsvOutName      As String
    
'    pref_NCnt = 0                ' 現在発行ページ初期化
    On Error GoTo CXHZ0050_CsvOut_Err
'------------------------'
'    出力ﾃｷｽﾄ OPEN
'------------------------'
'    CsvOutName = GetSetting(gRegKey, "CXH00050", "Csv", "")
    CsvOutName = CisSimplePrint.GetCsvFileName
    cFNo = FreeFile: gCnt1 = 0
    Open CsvOutName For Output As #cFNo

    PrtHead = "打切種別：," & RTrim(cUKbnm) & ",,警告日," & RTrim(cKYmd) & ",,打切日," & RTrim(cUYmd)
    On Error GoTo CXHZ0050_CsvOut_Err
    Print #cFNo, PrtHead
    
    PrtHead = "№,仕入先,仕入先名,受入,材質,寸法,背番号,構成"
    On Error GoTo CXHZ0050_CsvOut_Err
    Print #cFNo, PrtHead
' 明細を読み込んでセットする (レコード件数分)
    Do Until Not ZTM_RDSTS
        With ZTM
    '------------------------'
    '   レコード編集
    '------------------------'
        gCnt1 = gCnt1 + 1
        gStr = ""
        gStr = Format(gCnt1, "###0") & ","
        gStr = gStr & "'" & RTrim(.取引先) & ","
        gStr = gStr & RTrim(.手配先名) & ","
        gStr = gStr & "'" & RTrim(.受入) & ","
        gStr = gStr & RTrim(.材質) & ","
        gStr = gStr & RTrim(.表示寸法) & ","
        gStr = gStr & RTrim(.背番号) & ","
        gStr = gStr & .件数 & ""
    '------------------------'
    '   レコード出力
    '------------------------'
        On Error GoTo CXHZ0050_CsvOut_Err
        Print #cFNo, gStr
        End With
        
        With CisSimplePrint
            .ProgressCnt = gCnt1
            .ProgressSet
        End With
        '次のﾚｺｰﾄﾞを読み込む
        Call ZTMReadNext(1)
    Loop
'------------------------'
'   出力ﾃｷｽﾄ CLOSE
'------------------------'
    Close #cFNo

    CXHZ0050_CsvOut = True

CXHZ0050_CsvOut_Ed:
    Call ZTMClose(1)
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "    【  CSV出力 正常終了 】            "
        .MB_MSG(3) = "    出 力 先  = " & RTrim(CsvOutName) & "　　　　　"
        .MB_MSG(4) = "    出力件数 = " & Format(gCnt1, "#,##0") & " 件"
        .MB_Title = "【確認】"
        .MB_Button = OK
        .MBOX
    End With
    Exit Function
CXHZ0050_CsvOut_Err:
    CisFun.ErrorBox
    End
End Function




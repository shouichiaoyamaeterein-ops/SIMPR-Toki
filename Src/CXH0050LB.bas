Attribute VB_Name = "CXH0050LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   発注打切対象品番リスト印刷
'*
'*   モジュールID  : CXH0050LB
'*
'*                   CXH0050LPrint(Ukbn, Hinbn, Torcd)
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
'    If Not CXH0050LPrint Then Exit Sub
'
'*   作成日  :2004/02/12   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCXH0050        As New CISFormPrint
Public CisSimplePrint       As New CisSimplePrint
Public CXH0050_PrtCnt       As Long
Public cUKbnm               As String
Public cUYmd                As String
Public cKYmd                As String
'
Public Function CXH0050LPrint(Ukbn As String, _
                             Hinbn As String, _
                             Torcd As String, Optional PrintOut As Integer = 0) As Boolean
    CXH0050LPrint = False

'   品番取引先マスタ
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  htm.品番"
    gSL_Select = gSL_Select & ", htm.取引先"
    gSL_Select = gSL_Select & ", htm.受入"
    gSL_Select = gSL_Select & ", htm.背番号"
    gSL_Select = gSL_Select & ", trm.略称 取引先名"
    gSL_Select = gSL_Select & ", hnm.表示品番"
    gSL_Select = gSL_Select & ",("
    gSL_Select = gSL_Select & "select count(*) from 構成マスタ kom"
    gSL_Select = gSL_Select & " where kom.親品番 = htm.品番"
    gSL_Select = gSL_Select & "    or kom.子品番 = htm.品番"
    gSL_Select = gSL_Select & " ) 件数"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  品番取引先マスタ htm"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ trm"
    gSL_Select = gSL_Select & "    on trm.取引先CD = htm.取引先"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 1"
    gSL_Select = gSL_Select & "  left outer join 品番マスタ hnm"
    gSL_Select = gSL_Select & "    on hnm.品番 = htm.品番"
    If Ukbn <> "3" Then
        gSL_Select = gSL_Select & " where htm.品番 = '" & RTrim(Hinbn) & "'"
        If Ukbn = "2" Then
            gSL_Select = gSL_Select & " and htm.取引先 = '" & RTrim(Torcd) & "'"
        End If
    Else
        gSL_Select = gSL_Select & " where htm.取引先 = '" & RTrim(Torcd) & "'"
    End If
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  htm.取引先"
    gSL_Select = gSL_Select & ", htm.受入"
    gSL_Select = gSL_Select & ", htm.品番"

    If Not HTMRead(gSL_Select, 0, 1) Then
        Call HTMClose
        Exit Function
    End If
    CXH0050_PrtCnt = CisDB.RecordCount
    With CisSimplePrint
        .MaxCount = CisDB.RecordCount
        .ProgressInit
    End With

    If PrintOut <> 2 Then
        With CisPrtCXH0050
             Set .iPrtForm = CXH0050L
             .dpMSG = "発注打切対象品番リスト　印刷中"
             Select Case PrintOut
                Case 0: .iPreview = pv_Print
                Case 1: .iPreview = pv_Preview
             End Select
             .dpMSGScreen = モーダル
             .PrintStart
        End With
    Else
        Call CXH0050_CsvOut
    End If

    CXH0050LPrint = True
End Function
'+----------------------------------------------------------+
'+         CSVファイル出力  処理                        ＋
'+----------------------------------------------------------+
Public Function CXH0050_CsvOut() As Boolean
    CXH0050_CsvOut = False
    Dim cFNo           As Integer
    Dim PrtHead         As String
    Dim CsvOutName      As String
    
'    pref_NCnt = 0                ' 現在発行ページ初期化
    On Error GoTo CXH0050_CsvOut_Err
'------------------------'
'    出力ﾃｷｽﾄ OPEN
'------------------------'
'    CsvOutName = GetSetting(gRegKey, "CXH00050", "Csv", "")
    CsvOutName = CisSimplePrint.GetCsvFileName
    cFNo = FreeFile: gCnt1 = 0
    Open CsvOutName For Output As #cFNo

    PrtHead = "打切種別：," & RTrim(cUKbnm) & ",,警告日," & RTrim(cKYmd) & ",,打切日," & RTrim(cUYmd)
    On Error GoTo CXH0050_CsvOut_Err
    Print #cFNo, PrtHead
    
    PrtHead = "№,仕入先,仕入先名,受入,品番,背番号,構成"
    On Error GoTo CXH0050_CsvOut_Err
    Print #cFNo, PrtHead
' 明細を読み込んでセットする (レコード件数分)
    Do Until Not HTM_RDSTS
        With HTM
    '------------------------'
    '   レコード編集
    '------------------------'
        gCnt1 = gCnt1 + 1
        gStr = ""
        gStr = Format(gCnt1, "###0") & ","
        gStr = gStr & "'" & RTrim(.取引先) & ","
        gStr = gStr & RTrim(.取引先名) & ","
        gStr = gStr & "'" & RTrim(.受入) & ","
        gStr = gStr & RTrim(.表示品番) & ","
        gStr = gStr & RTrim(.背番号) & ","
        gStr = gStr & .件数 & ""
    '------------------------'
    '   レコード出力
    '------------------------'
        On Error GoTo CXH0050_CsvOut_Err
        Print #cFNo, gStr
        End With
        
        With CisSimplePrint
            .ProgressCnt = gCnt1
            .ProgressSet
        End With
        '次のﾚｺｰﾄﾞを読み込む
        Call HTMReadNext(1)
    Loop
'------------------------'
'   出力ﾃｷｽﾄ CLOSE
'------------------------'
    Close #cFNo

    CXH0050_CsvOut = True

CXH0050_CsvOut_Ed:
    Call HTMClose(1)
    
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
CXH0050_CsvOut_Err:
    CisFun.ErrorBox
    End
End Function




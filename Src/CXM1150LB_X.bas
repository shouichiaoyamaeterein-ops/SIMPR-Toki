Attribute VB_Name = "CXM1150LBX"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   構成マスタ照会一覧
'*
'*   モジュールID  : CXM1150LB
'*
'*   作成日  :2019/07/23   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
    Public iRdCnt           As Long         ' 2019/07/18 ADD：品番件数カウント用
    Public wHinbn()         As String       ' 2019/07/18 ADD：品番情報
    Public wSyake()         As String       ' 2019/07/22 ADD：車型情報
    Public wRdKOL()         As Variant      ' 2019/07/18 ADD：構成問合せワーク

    Public bStaPara2        As Byte         ' 2019/07/26 ADD：(H1Op_Hinkb(0).Tag)
    Public bStaPara3        As Byte         ' 2019/07/26 ADD：(H1Op_Tenkb(0).Tag)
    Public bStaPara4        As Byte         ' 2019/07/26 ADD：(H1Op_Oyakb(0).Tag)
    Public wPrtHKbn         As Integer      ' 2019/07/26 ADD
    
    Public iTotalLineCnt    As Long
    
'+++++++++++++++++++++++++++++++++++++++++++++++++++    '++ デバッグ用です +++
'+                                                 +
'+  文字列をイミディエイトウィンドウに表示する     +
'+                                                 +
'+++++++++++++++++++++++++++++++++++++++++++++++++++
Public Sub Str_Monitor(ByVal wStr As String)
    Debug.Print vbCrLf
    Debug.Print Format(Now, "yyyy/MM/dd HH:mm:ss") + _
                " .................................................." + _
                vbCrLf + wStr
End Sub


''<<  KOL 内容  >>
'With KOL
'    wRdKOL(0, iLineCnt) = wHinbn(iCnt)      ' 検索元品番
'    wRdKOL(1, iLineCnt) = .順位
'    wRdKOL(2, iLineCnt) = .レベル
'    wRdKOL(3, iLineCnt) = .親品番
'    wRdKOL(4, iLineCnt) = .子品番
'    wRdKOL(5, iLineCnt) = .構成数
'    wRdKOL(6, iLineCnt) = .展開区分
'    wRdKOL(7, iLineCnt) = .LT
'    wRdKOL(8, iLineCnt) = .適用開始日
'    wRdKOL(9, iLineCnt) = .適用終了日
'    wRdKOL(10, iLineCnt) = .区分
'    wRdKOL(11, iLineCnt) = .表示品番
'    wRdKOL(12, iLineCnt) = .表示寸法
'    wRdKOL(13, iLineCnt) = .品名
'    wRdKOL(14, iLineCnt) = .材料管理番号
'    wRdKOL(15, iLineCnt) = wSyake(iCnt)     ' 検索元車型
'End With

Public Function TestDataToVSV()
    TestDataToVSV = False
    Dim cFNo           As Integer
    Dim PrtHead         As String
    Dim CsvOutName      As String
    Dim Key_Torcd       As String
    
'    pref_NCnt = 0                ' 現在発行ページ初期化
    On Error GoTo TestDataToVSV_Err
'------------------------'
'    出力ﾃｷｽﾄ OPEN
'------------------------'
    CsvOutName = "C:\Users\KCOM101.CISDOM\Desktop\AAAA\TEST_" & Format(Now(), "yyyy.MM.dd_HH.mm.ss") & ".csv" 'CisSimplePrint.GetCsvFileName
    
    cFNo = FreeFile: gCnt1 = 0
    Open CsvOutName For Output As #cFNo

    PrtHead = "検索元品番,順位,レベル,親品番,子品番,構成数,展開区分,LT,適用開始日,適用終了日,区分,表示品番,表示寸法,品名,材料管理番号,検索元車型"
    On Error GoTo TestDataToVSV_Err
    Print #cFNo, PrtHead
    
    
    '===================================================================================================' 2019/07/26 ADD START
    Dim iCnt As Long
    For iCnt = 1 To iTotalLineCnt
        gStr = ""
        
        
        gStr = gStr & wRdKOL(0, iCnt) & ","     ' 検索元品番
        gStr = gStr & wRdKOL(1, iCnt) & ","     ' 順位
        gStr = gStr & wRdKOL(2, iCnt) & ","     ' レベル
        gStr = gStr & wRdKOL(3, iCnt) & ","     ' 親品番
        gStr = gStr & wRdKOL(4, iCnt) & ","     ' 子品番
        gStr = gStr & wRdKOL(5, iCnt) & ","     ' 構成数
        gStr = gStr & wRdKOL(6, iCnt) & ","     ' 展開区分
        gStr = gStr & wRdKOL(7, iCnt) & ","     ' LT
        gStr = gStr & wRdKOL(8, iCnt) & ","     ' 適用開始日
        gStr = gStr & wRdKOL(9, iCnt) & ","     ' 適用終了日
        gStr = gStr & wRdKOL(10, iCnt) & ","    ' 区分
        gStr = gStr & wRdKOL(11, iCnt) & ","    ' 表示品番
        gStr = gStr & wRdKOL(12, iCnt) & ","    ' 表示寸法
        gStr = gStr & wRdKOL(13, iCnt) & ","    ' 品名
        gStr = gStr & wRdKOL(14, iCnt) & ","    ' 材料管理番号
        gStr = gStr & wRdKOL(15, iCnt)          ' 検索元車型
        
        On Error GoTo TestDataToVSV_Err
        Print #cFNo, gStr
   
SKIPLB01:
    Next
    
'------------------------'
'   出力ﾃｷｽﾄ CLOSE
'------------------------'
    Close #cFNo

    TestDataToVSV = True

TestDataToVSV_Ed:
    Call KOLClose(1)
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(1) = "    【  CSV出力 正常終了 】            "
        .MB_MSG(3) = "    出 力 先  = " & RTrim(CsvOutName) & "　　　　　"
        .MB_MSG(4) = "    出力件数 = " & Format(iTotalLineCnt, "#,##0") & " 件"
        .MB_Title = "【確認】"
        .MB_Button = OK
        .MBOX
    End With
    Exit Function
TestDataToVSV_Err:
    CisFun.ErrorBox
    End

End Function



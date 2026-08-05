Attribute VB_Name = "CXM2540LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   スクラップ単価変動額問合せ(支払)
'*
'*   モジュールID  : CXM2540LB
'*
'*                   CXM2540LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:
'*
'*   If Not CXM2540LPrint Then Exit Sub
'*
'*   作成日  :2008/09/08   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
'Public CisPrtCXM2540        As New CISFormPrint
Global CisPrtCXM2540        As CISFormPrint
Global CXM2540_PrtCnt       As Long
Global CXM2540_PrintKb      As Byte
Global CXM2540_PrintKb2     As Byte
Global CXM2540_PrtOut       As String       ' 2004/05/10 Add
Global CXM2545_PrtKb        As Byte
Global CXM2540_Kara         As Boolean
'
'-- 共通変数の定義 ---------------------
Public PrtDat(20)           As String       ' 印刷するGridの内容 (１行分)
Public PrtRow               As Integer      ' 印刷する行の位置
Public PrtCnt               As Integer
Public PrtTim               As String
Public PrtNKi               As String
Public PrtKKn               As String

Public Kei_S                As Double       ' スクラップ単価変動額の合計
Public Kei_U                As Double       ' 売上単価変動額の合計
'

Public Function CXM2540LPrint() As Boolean
    CXM2540LPrint = False
    
    PrtTim = Format(Now(), "yyyy/MM/dd HH:mm")      ' 発行日
    CXM2540_PrtOut = PrtCnt                         ' 件数
    
    Set CisPrtCXM2540 = New CISFormPrint
    With CisPrtCXM2540
        Set .iPrtForm = CXM2540L
        .dpMSG = "スクラップ単価変動額　印刷中"
        .iPreview = pv_Print
       '.iPreview = pv_Preview
        .dpMSGScreen = モードレス
        .PrintStart
    End With
    Set CisPrtCXM2540 = Nothing
    
    CXM2540LPrint = True
End Function

''+-------------------------------------+
''+      Debug用 SQL記録                +
''+-------------------------------------+
Public Sub Test_LogOutput(str As String, Msg As Integer)
    Exit Sub    '=========================================== 実行しないときはこの行を生かすこと ====
    Dim dt01 As String
    Dim dt02 As String
    Dim f_No As Integer
    
    If Msg <> 0 Then MsgBox (str)
    dt01 = CStr(DateTime.Now)
    dt02 = Mid(dt01, 1, 4) & "." & _
           Mid(dt01, 6, 2) & "." & _
           Mid(dt01, 9, 2)

    f_No = FreeFile                             'ファイル番号の取得
    Open "D:\TestLog\SQL_" & dt02 & ".log" _
          For Append As #f_No                   'ファイルを追加モードで開く
    Print #f_No, "[ PXM2500/PXM254x スクラップ単価変動額問合せ(支払) ] " & dt01 & _
                 " ----------------------------------------" & Chr(13) & Chr(10) & _
                 str & Chr(13) & Chr(10)        'ファイルへ書き込む"
    Close #f_No                                 'ファイルを閉じる
End Sub

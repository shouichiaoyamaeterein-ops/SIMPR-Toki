Attribute VB_Name = "CXH0530LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   支給納品書／受領書発行
'*
'*   モジュールID  : CXH0530LB
'*
'*                  CXH0530LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not CXH0530LPrint Then Exit Sub
'
'*   作成日  :2008/08/20   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
'Public CisPrtCXH0530        As New CISFormPrint
Public CisPrtCXH0530        As CISFormPrint
Public CXH0530_PrtCnt       As Long
Global CXH0530_PrintKb      As Byte
Global CXH0530_PrintKb2     As Byte
Public CXH0530_PrtOut       As String       ' 2004/05/10 Add

Public Function CXH0530LPrint() As Boolean
    CXH0530LPrint = False
    
    gSL_Select = "SELECT * " & vbCrLf
    gSL_Select = gSL_Select & " From  帳票出力テーブル        " & vbCrLf
    gSL_Select = gSL_Select & " Where 端末番号 =  HOST_NAME() " & vbCrLf
    gSL_Select = gSL_Select & "   And 帳票区分 =  201         " & vbCrLf
    gSL_Select = gSL_Select & "   And 検索情報 <> ''          " & vbCrLf
    gSL_Select = gSL_Select & " Order By 検索情報, 管理NO " & vbCrLf
    
    If CHTRead(gSL_Select) Then
        CXH0530_PrtOut = CisDB.RecordCount
'       CXH0530_PrtCnt = CisDB.RecordCount
'       CXH0530_Kara = False
       Set CisPrtCXH0530 = New CISFormPrint
       With CisPrtCXH0530
            Set .iPrtForm = CXH0530L
            .dpMSG = "納品書・受領書　印刷中"
            .iPreview = pv_Print
           '.iPreview = pv_Preview
            .dpMSGScreen = モードレス
            .PrintStart
       End With
       Set CisPrtCXH0530 = Nothing
    Else
       Call CHTClose
    End If
    
    CXH0530LPrint = True
End Function

''+-------------------------------------+
''+      Debug用 SQL記録                +
''+-------------------------------------+
Public Sub Test_LogOutput(str As String, Msg As Integer)
'Exit Sub   '=========================================== 実行しないときはこの行を生かすこと =========================
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
    Print #f_No, "[ PXH0530 納品書印刷 ] " & dt01 & _
                 " ----------------------------------------" & Chr(13) & Chr(10) & _
                 str & Chr(13) & Chr(10)        'ファイルへ書き込む"
    Close #f_No                                 'ファイルを閉じる
End Sub

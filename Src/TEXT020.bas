Attribute VB_Name = "TEXT020"
Option Explicit

'*--------------------------------------------------------------------------------
'*
'*   TEXT処理
'*          手配データ
'*
'*   モジュールID  : TEXT020.BAS
'*
'*
'*   作成日  : 2007/03/05   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public T020_SyoriDay        As String

Dim T020_Cnt                As Long
Dim T020_Head               As String

Dim Del_Day                 As Date
Dim T020_File
'+----------------------------------+
'+  テーブル出力
'+----------------------------------+
Public Function T020ConVertPut(ConvData As String, lErr As String) As Boolean
    lErr = ""
    
    T020ConVertPut = False
    
    On Error GoTo T020ConVertPut_Er:
    
    'HEAD部
    If Mid(ConvData, 1, 1) = "H" Then
       T020_SyoriDay = Format(Now, "yyyymmddhhnn")
       T020_Head = Mid(ConvData, 1, 30)
       '--------------------------------------- 同一HEAD内容のデータを削除
       gSL_Select = "DELETE FROM TNS支給送信品番確認テーブル"
       gSL_Select = gSL_Select & " WHERE 契約先 = '" & RTrim(TKT.契約先) & "'"
       gSL_Select = gSL_Select & "   AND FILECD = '" & RTrim(TKT.FILECD) & "'"
       gSL_Select = gSL_Select & "   AND FCD = '" & RTrim(TKT.FCODE) & "'"
       gSL_Select = gSL_Select & "   AND HEADKEY = '" & RTrim(T020_Head) & "'"
       With CisDB
           .SQL = gSL_Select
           .DBExec
       End With
    
       '+--------------------------------------- 過去データ削除
       If TKT.保持期間 <> 0 Then
            Del_Day = DateAdd("d", CDbl(TKT.保持期間 * -1), Now)
            
            gSL_Select = "DELETE FROM TNS支給送信品番確認テーブル"
            gSL_Select = gSL_Select & " WHERE 契約先 = '" & RTrim(TKT.契約先) & "'"
            gSL_Select = gSL_Select & "   AND FILECD = '" & RTrim(TKT.FILECD) & "'"
            gSL_Select = gSL_Select & "   AND FCD = '" & RTrim(TKT.FCODE) & "'"
            gSL_Select = gSL_Select & "   AND 処理日時 < '" & Format(Del_Day, "yyyymmdd") & "'"
            With CisDB
                .SQL = gSL_Select
                .DBExec
            End With
        End If
    End If
    'ﾃﾞｰﾀ部
    If Mid(ConvData, 1, 1) = "D" Then
       '［納入場所］が自社と同一のみ対象
       If tx_JisyaChk Then
          If Mid(ConvData, 39, 4) <> TKT.自社CD Then
             T020ConVertPut = True
             Exit Function
          End If
       End If
        
       tx_InCnt = tx_InCnt + 1
        
       Call ItemsClearTSK
        
       With TSK
            .契約先 = TKT.契約先
            .FILECD = TKT.FILECD
            .FCD = TKT.FCODE
            .処理日時 = T020_SyoriDay
            .HEADKEY = T020_Head
            .区分 = Mid(ConvData, 1, 1)
            .識別子 = Mid(ConvData, 2, 1)
            .ブランク1 = Mid(ConvData, 3, 2)
            .カード = Mid(ConvData, 5, 4)
            .仕入先 = Mid(ConvData, 9, 4)
            .仕入先工区 = Mid(ConvData, 13, 1)
            .管理番号 = Mid(ConvData, 14, 6)
            .ブランク2 = Mid(ConvData, 20, 5)
            .品番 = Mid(ConvData, 25, 12)
            .単位 = Mid(ConvData, 37, 2)
            .納入場所 = Mid(ConvData, 39, 4)
            .納入区分 = Mid(ConvData, 43, 1)
            .支給区分 = Mid(ConvData, 44, 1)
            .部品区分 = Mid(ConvData, 45, 1)
            .背番号 = Mid(ConvData, 46, 3)
            .収容数 = Mid(ConvData, 49, 5)
            .ブランク3 = Mid(ConvData, 54, 14)
            .品名 = Mid(ConvData, 68, 20)
            .担当者 = Mid(ConvData, 88, 2)
            .提供依頼者 = Mid(ConvData, 90, 4)
            .REC = RTrim(ConvData)
       End With
        
       Call TSKInsert

       tx_OutCnt = tx_OutCnt + 1
    End If

    
T020ConVertPut_ED:
    
    T020ConVertPut = True
    On Error GoTo 0
    
    Exit Function
T020ConVertPut_Er:
    Static ER_MSG1 As String
    Static ER_MSG2 As String
    
    ER_MSG1 = "         手配データ変換処理（０２０）      " & vbCrLf & " " & vbCrLf
    ER_MSG1 = ER_MSG1 & "【 " & Err.Description & " 】"

    Call ProcLogPut(ER_MSG1)

    ER_MSG2 = Chr(13) & Space(3) & _
               "上記エラーにより、以降の処理を続行出来ません" & _
               Chr(10) & Chr(13) & Space(8) & _
               "内容を確認して再度実行して下さい"
    lErr = "Z9"
    Call TERPut(lErr, ER_MSG1, ER_MSG2)

End Function


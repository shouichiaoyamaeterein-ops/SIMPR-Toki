Attribute VB_Name = "TEXT112"
Option Explicit

'*--------------------------------------------------------------------------------
'*
'*   TEXT処理
'*          手配データ
'*
'*   モジュールID  : TEXT112.BAS
'*
'*
'*   作成日  : 2007/03/05   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public T112_SyoriDay        As String

Dim T112_Cnt                As Long
Dim T112_Head               As String

Dim Del_Day                 As Date
Dim T112_File
'+----------------------------------+
'+  テーブル出力
'+----------------------------------+
Public Function T112ConVertPut(ConvData As String, lErr As String) As Boolean
    lErr = ""
    
    T112ConVertPut = False
    
    On Error GoTo T112ConVertPut_Er:
    
    'HEAD部
    If Mid(ConvData, 1, 1) = "H" Then
       T112_SyoriDay = Format(Now, "yyyymmddhhnn")
       T112_Head = Mid(ConvData, 1, 30)
       '--------------------------------------- 同一HEAD内容のデータを削除
       gSL_Select = "DELETE FROM TNS支給送信品番確認テーブル"
       gSL_Select = gSL_Select & " WHERE 契約先 = '" & RTrim(TKT.契約先) & "'"
       gSL_Select = gSL_Select & "   AND FILECD = '" & RTrim(TKT.FILECD) & "'"
       gSL_Select = gSL_Select & "   AND FCD = '" & RTrim(TKT.FCODE) & "'"
       gSL_Select = gSL_Select & "   AND HEADKEY = '" & RTrim(T112_Head) & "'"
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
       '［支給先］が自社と同一のみ対象
       If tx_JisyaChk Then
          If Mid(ConvData, 21, 4) <> TKT.自社CD Then
             T112ConVertPut = True
             Exit Function
          End If
       End If
        
       tx_InCnt = tx_InCnt + 1
        
       Call ItemsClearTSK
        
       With TSK
            .契約先 = TKT.契約先
            .FILECD = TKT.FILECD
            .FCD = TKT.FCODE
            .処理日時 = T112_SyoriDay
            .HEADKEY = T112_Head
            .仕入先 = Mid(ConvData, 16, 4)
            .仕入先工区 = Mid(ConvData, 20, 1)
            .管理番号 = Mid(ConvData, 26, 6)
            .品番 = Mid(ConvData, 4, 12)
            .納入場所 = Mid(ConvData, 21, 4)
            .背番号 = Mid(ConvData, 52, 4)
            .収容数 = Mid(ConvData, 38, 5)
            .品名 = Mid(ConvData, 109, 40)
            .REC = RTrim(ConvData)
       End With
        
       Call TSKInsert

       tx_OutCnt = tx_OutCnt + 1
    End If

    
T112ConVertPut_ED:
    
    T112ConVertPut = True
    On Error GoTo 0
    
    Exit Function
T112ConVertPut_Er:
    Static ER_MSG1 As String
    Static ER_MSG2 As String
    
    ER_MSG1 = "         手配データ変換処理（１１２）      " & vbCrLf & " " & vbCrLf
    ER_MSG1 = ER_MSG1 & "【 " & Err.Description & " 】"

    Call ProcLogPut(ER_MSG1)

    ER_MSG2 = Chr(13) & Space(3) & _
               "上記エラーにより、以降の処理を続行出来ません" & _
               Chr(10) & Chr(13) & Space(8) & _
               "内容を確認して再度実行して下さい"
    lErr = "Z9"
    Call TERPut(lErr, ER_MSG1, ER_MSG2)

End Function




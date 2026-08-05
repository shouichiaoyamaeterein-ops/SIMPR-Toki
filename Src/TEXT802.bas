Attribute VB_Name = "TEXT802"
Option Explicit

'*--------------------------------------------------------------------------------
'*
'*   TEXT処理
'*          検収データ（車体精工：支給）
'*
'*   モジュールID  : TEXT802.BAS
'*
'*
'*   作成日  :2008/05/14   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Dim T802_Cnt         As Long
Dim T802_ST          As Long

Global T802_HYMD        As String
Global T802_SEQ         As String

'+----------------------------------+
'+  テーブル出力
'+----------------------------------+
Public Function T802ConVertPut(ConvData As String, lErr As String) As Boolean
    lErr = ""

    On Error GoTo T802ConVertPut_Er:
    
    ' 対象年月・財務管理番号を求める
    If tx_KNo = 0 Then
        ' 契約先
        tx_KTorcd = TKT.契約先
        ' 対象年月
        With CisFun
            .Date_RsFormat = fYM
            .Date_Add tx_HeadYmd, M, -1, 1
            tx_TYM = .Date_Result
        End With
        ' TNS買掛金テーブル削除
        With CisDB
             .SQL = "DELETE TEXT買掛金テーブル "
             .SQL = .SQL & " WHERE 契約先 = '" & tx_KTorcd & "' "
             .SQL = .SQL & "   AND 対象年月 = '" & tx_TYM & "' "
             .SQL = .SQL & "   AND HEAD作成日 = '" & tx_HeadYmd & "' "
             .SQL = .SQL & "   AND HEADSEQ = '" & tx_HeadSeq & "' "
             .SQL = .SQL & "   AND FILECODE = '" & TKT.FILECD & "' "
        
             Call .DBExec
             Call ProcLogPut("TEXT買掛金 削除 契約先=" & TKT.契約先 & ",FILECODE = " & TKT.FILECD & ",HEAD作成日 = " & tx_HeadYmd & ",HEADSEQ = '" & tx_HeadSeq & ",削除件数=" & CisDB.RecordCount)
             If TKT.オプション2 <> 0 Then
                With CisFun
                    .Date_RsFormat = fYM
                    .Date_Add tx_TYM, M, TKT.オプション2 * -1, LastDate
                    gStr = .Date_Result
                End With
                .SQL = "DELETE From TEXT買掛金テーブル"
                .SQL = .SQL & " WHERE 契約先 = '" & tx_KTorcd & "' "
                .SQL = .SQL & "   AND 対象年月 <= '" & gStr & "' "
                .SQL = .SQL & "   AND FILECODE = '" & TKT.FILECD & "' "
                Call .DBExec
                Call ProcLogPut("TEXT買掛金 削除 契約先=" & TKT.契約先 & ",対象年月 <= " & gStr & ",削除件数=" & CisDB.RecordCount)
             End If
        End With
        
        ' 財務管理番号
        gSL_Select = "SELECT MAX(財務管理番号) 財務管理番号 From TEXT買掛金テーブル"
        If TKKRead(gSL_Select, 1) Then
            tx_KNo = TKK.財務管理番号
        End If
    End If
        
    T802ConVertPut = False
'
    Call ItemsClearTKK
'
    With TKK
        .契約先 = tx_KTorcd
        .対象年月 = tx_TYM
        .対象期 = 0
        .対象日 = ""
        .HEAD作成日 = tx_HeadYmd
        .HEADSEQ = tx_HeadSeq
        .FILECODE = TKT.FILECD
        .FCD = TKT.FCODE
        
        .識別子 = ""
        .仕入先 = RTrim(Mid(ConvData, 1, 4))
        .購入区分 = ""
        
        .品番 = RTrim(Mid(ConvData, 11, 12))
        .受入 = ""
        .単価 = CisFun.Val2(Mid(ConvData, 30, 11))
        .訂正理由 = ""
        .連番 = ""
        .部品管理区分 = ""
        .仮単価 = 0
        
        .納入番号 = Trim(Mid(ConvData, 50, 6))
        cTKK_KYM = Mid(ConvData, 56, 4)
        cTKK_YMD = Format(Date, "YYYY") & cTKK_KYM
        If Format(Date, "YYYYMM") < Mid(cTKK_YMD, 1, 6) Then
            cTKK_YMD = Format(Date, "YYYY") - 1 & cTKK_KYM
        End If
        .検収日 = cTKK_YMD
        .数量 = CisFun.Val2(Mid(ConvData, 23, 7))
        .金額 = CisFun.Val2(Mid(ConvData, 41, 9))
        .マイナス = ""
                
        tx_KNo = tx_KNo + 1
        .財務管理番号 = tx_KNo

        If Len(.納入番号) = 6 And Mid(.納入番号, 1, 1) = "0" Then
           .納入番号 = Mid(.納入番号, 2)
        End If
        
        Call TKKInsert
        tx_OutCnt = tx_OutCnt + 1
    End With
'
    
T802ConVertPut_ED:
    
    T802ConVertPut = True
    On Error GoTo 0
    
    Exit Function
T802ConVertPut_Er:
    Static ER_MSG1 As String
    Static ER_MSG2 As String
    
    ER_MSG1 = "         財務データ変換処理（８０２）      " & vbCrLf & " " & vbCrLf
    ER_MSG1 = ER_MSG1 & "【 " & Err.Description & " 】"

    Call ProcLogPut(ER_MSG1)

    ER_MSG2 = Chr(13) & Space(3) & _
               "上記エラーにより、以降の処理を続行出来ません" & _
               Chr(10) & Chr(13) & Space(8) & _
               "内容を確認して再度実行して下さい"
    lErr = "Z1"
    Call TERPut("Z1", ER_MSG1, ER_MSG2)
    
End Function



Attribute VB_Name = "TEXT701"
Option Explicit

'*--------------------------------------------------------------------------------
'*
'*   TEXT処理
'*          検収データ（車体）
'*
'*   モジュールID  : TEXT701.BAS
'*
'*
'*   作成日  :2004/11/08   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Dim T701_Cnt         As Long
Dim T701_ST          As Long

Global T701_HYMD        As String
Global T701_SEQ         As String

'+----------------------------------+
'+  テーブル出力
'+----------------------------------+
Public Function T701ConVertPut(ConvData As String, lErr As String) As Boolean
    lErr = ""

    On Error GoTo T701ConVertPut_Er:
    
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
    
    ConvData = Mid(ConvData, 2)
    
    If Mid(ConvData, 1, 2) <> "11" And Mid(ConvData, 1, 2) <> "12" And _
       Mid(ConvData, 1, 2) <> "13" And Mid(ConvData, 1, 2) <> "21" And _
       Mid(ConvData, 1, 2) <> "14" And Mid(ConvData, 1, 2) <> "15" Then GoTo T701ConVertPut_ED
    
    T701ConVertPut = False
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
        
        .識別子 = RTrim(Mid(ConvData, 1, 2))
        .仕入先 = RTrim(Mid(ConvData, 3, 4))
        .購入区分 = RTrim(Mid(ConvData, 7, 1))
        
        .品番 = RTrim(Mid(ConvData, 19, 12))
        .受入 = Trim(Mid(ConvData, 31, 2))
        .単価 = CisFun.Val2(Mid(ConvData, 34, 8)) / 100
        .訂正理由 = Trim(Mid(ConvData, 42, 2))
        ' 部品A部品表   ﾚｺｰﾄﾞ識別 : 11
        If Mid(ConvData, 1, 2) = "11" Then
            .管理番号 = RTrim(Mid(ConvData, 8, 6))
            .支給先 = RTrim(Mid(ConvData, 14, 4))
            .仮号試 = Trim(Mid(ConvData, 33, 1))
        End If
        ' 一般仕入     ﾚｺｰﾄﾞ識別 : 12
        If Mid(ConvData, 1, 2) = "12" Then
            .連番 = RTrim(Mid(ConvData, 8, 7))
            .決済番号 = RTrim(Mid(ConvData, 15, 4))
            .仮単価 = Trim(Mid(ConvData, 33, 1))
        End If
        ' 補給部品     ﾚｺｰﾄﾞ識別 : 13
        If Mid(ConvData, 1, 2) = "13" Then
            .連番 = RTrim(Mid(ConvData, 8, 7))
            .部品管理区分 = RTrim(Mid(ConvData, 15, 1))
            .仮単価 = Trim(Mid(ConvData, 33, 1))
        End If
        ' 支給部品     ﾚｺｰﾄﾞ識別 : 21
        If Mid(ConvData, 1, 2) = "21" Then
            .管理番号 = RTrim(Mid(ConvData, 8, 6))
            .支給先 = RTrim(Mid(ConvData, 14, 4))
            .仮単価 = Trim(Mid(ConvData, 33, 1))
        End If
        ' 部品A型費    ﾚｺｰﾄﾞ識別 : 14
        If Mid(ConvData, 1, 2) = "14" Then
            .管理番号 = RTrim(Mid(ConvData, 8, 6))
            .支給先 = RTrim(Mid(ConvData, 14, 4))
            .仮単価 = Trim(Mid(ConvData, 33, 1))
        End If
        ' 補給部品A型費    ﾚｺｰﾄﾞ識別 : 15
        If Mid(ConvData, 1, 2) = "15" Then
            .連番 = RTrim(Mid(ConvData, 8, 7))
            .部品管理区分 = RTrim(Mid(ConvData, 15, 1))
            .仮単価 = Trim(Mid(ConvData, 33, 1))
        End If
        
        For T701_Cnt = 1 To 7
            T701_ST = 51 + (T701_Cnt - 1) * 27
            If RTrim(Mid(ConvData, T701_ST, 20)) = "" Then Exit For
            .納入番号 = Trim(Mid(ConvData, T701_ST, 6))
            cTKK_KYM = Mid(ConvData, T701_ST + 6, 4)
            cTKK_YMD = Format(Date, "YYYY") & cTKK_KYM
            If Format(Date, "YYYYMM") < Mid(cTKK_YMD, 1, 6) Then
                cTKK_YMD = Format(Date, "YYYY") - 1 & cTKK_KYM
            End If
            .検収日 = cTKK_YMD
            .数量 = CisFun.Val2(Mid(ConvData, T701_ST + 10, 7))
            .金額 = CisFun.Val2(Mid(ConvData, T701_ST + 17, 9))
            .マイナス = Trim(Mid(ConvData, T701_ST + 26, 1))
            If .マイナス = "-" Then
                .金額 = .金額 * -1
            End If
            
            If .識別子 = "14" Or .識別子 = "15" Then
               .数量 = 1
               .単価 = .金額
               .当月支払額 = .金額
               .検収日 = "20" & Mid(.納入番号, 1, 2) & Mid(ConvData, T701_ST + 6, 4)
            End If
            
            tx_KNo = tx_KNo + 1
            .財務管理番号 = tx_KNo

            If Len(.納入番号) = 6 And Mid(.納入番号, 1, 1) = "0" Then
               .納入番号 = Mid(.納入番号, 2)
            End If
            
            Call TKKInsert
            tx_OutCnt = tx_OutCnt + 1
        Next T701_Cnt
        ' 識別子格納
        gCnt2 = 0
        For gCnt1 = 0 To UBound(tx_Shiki)
            If tx_Shiki(gCnt1) = .識別子 Then
               gCnt2 = 1
               Exit For
            End If
            If tx_Shiki(gCnt1) = "" Then
               tx_Shiki(gCnt1) = .識別子
               gCnt2 = 1
               Exit For
            End If
        Next gCnt1
        If gCnt2 = 0 Then
           ReDim Preserve tx_Shiki(UBound(tx_Shiki) + 1)
           tx_Shiki(UBound(tx_Shiki)) = .識別子
        End If
    End With
'
    
T701ConVertPut_ED:
    
    T701ConVertPut = True
    On Error GoTo 0
    
    Exit Function
T701ConVertPut_Er:
    Static ER_MSG1 As String
    Static ER_MSG2 As String
    
    ER_MSG1 = "         財務データ変換処理（７０１）      " & vbCrLf & " " & vbCrLf
    ER_MSG1 = ER_MSG1 & "【 " & Err.Description & " 】"

    Call ProcLogPut(ER_MSG1)

    ER_MSG2 = Chr(13) & Space(3) & _
               "上記エラーにより、以降の処理を続行出来ません" & _
               Chr(10) & Chr(13) & Space(8) & _
               "内容を確認して再度実行して下さい"
    lErr = "Z1"
    Call TERPut("Z1", ER_MSG1, ER_MSG2)
    
End Function




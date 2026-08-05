Attribute VB_Name = "TEXT060"
Option Explicit

'*--------------------------------------------------------------------------------
'*
'*   TEXT処理
'*          検収データ（買掛金・カード分）
'*
'*   モジュールID  : TEXT060.BAS
'*
'*
'*   作成日  :2004/11/09   By CIS
'*   変更日  : 2005/03/31  By CIS 計算方法を取引先マスタの設定に従う
'*---------------------------------------------------------------------------------
Dim T060_Cnt         As Long
Dim T060_ST          As Long

Global T060_HYMD        As String
Global T060_SEQ         As String

'+----------------------------------+
'+  テーブル出力
'+----------------------------------+
Public Function T060ConVertPut(ConvData As String, lErr As String) As Boolean
    lErr = ""
    
    T060ConVertPut = False
    
    On Error GoTo T060ConVertPut_Er:
    
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
    
    T060ConVertPut = False
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
        
        .識別子 = RTrim(Mid(ConvData, 22, 2))
        .品番 = RTrim(Mid(ConvData, 6, 12))
        .受入 = Trim(Mid(ConvData, 18, 4))
        .単位 = RTrim(Mid(ConvData, 24, 2))
        .重量 = CisFun.Val2(Mid(ConvData, 37, 6)) / 1000
        .単価 = CisFun.Val2(Mid(ConvData, 26, 11)) / 100

        .仕入先 = RTrim(Mid(ConvData, 1, 4))
        .データ元 = RTrim(Mid(ConvData, 5, 1))
        .補助品番 = RTrim(Mid(ConvData, 43, 12))
        
        For T060_Cnt = 1 To 13
            T060_ST = 55 + (T060_Cnt - 1) * 15
            If RTrim(Mid(ConvData, T060_ST, 4)) = "" Then Exit For
            cTKK_KYM = Mid(ConvData, T060_ST, 4)
            cTKK_YMD = Format(Date, "YYYY") & cTKK_KYM
            If Format(Date, "YYYYMM") < Mid(cTKK_YMD, 1, 6) Then
                cTKK_YMD = Format(Date, "YYYY") - 1 & cTKK_KYM
            End If
            .検収日 = cTKK_YMD
            .納入番号 = RTrim(Mid(ConvData, T060_ST + 4, 5))
            .数量 = CisFun.Val2(Mid(ConvData, T060_ST + 9, 6))
    '+-----------------------------------------------  2005/03/31 Delete
'            cTTK_Kin = .単価 * .数量
'            .金額 = Fix(cTTK_Kin)
'            If .金額 <> cTTK_Kin Then
'                .金額 = .金額 + 1
'            End If
    '+-----------------------------------------------  2005/03/31 Delete End
    '+-----------------------------------------------  2005/03/31 Insert
            cTTK_Kin = .数量 * .単価
            .金額 = Fix(cTTK_Kin)        '切り捨て
            cTTK_Syosu = cTTK_Kin - .金額
            
            If TRM.計算単位 = 1 Then '切り上げ
                If .金額 <> cTTK_Kin Then
                    .金額 = .金額 + 1
                End If
            End If
            If TRM.計算単位 = 2 Then '四捨五入
                If cTTK_Syosu > 0.5 Then
                    .金額 = .金額 + 1
                End If
            End If
    '+-----------------------------------------------  2005/03/31 Insert  End
            tx_KNo = tx_KNo + 1
            .財務管理番号 = tx_KNo

            Call TKKInsert
            tx_OutCnt = tx_OutCnt + 1
        Next T060_Cnt
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
   
T060ConVertPut_ED:
    
    T060ConVertPut = True
    On Error GoTo 0
    
    Exit Function
T060ConVertPut_Er:
    Static ER_MSG1 As String
    Static ER_MSG2 As String
    
    ER_MSG1 = "         財務データ変換処理（０６０）      " & vbCrLf & " " & vbCrLf
    ER_MSG1 = ER_MSG1 & "【 " & Err.Description & " 】"

    Call ProcLogPut(ER_MSG1)

    ER_MSG2 = Chr(13) & Space(3) & _
               "上記エラーにより、以降の処理を続行出来ません" & _
               Chr(10) & Chr(13) & Space(8) & _
               "内容を確認して再度実行して下さい"
    lErr = "Z6"
    Call TERPut(lErr, ER_MSG1, ER_MSG2)
End Function


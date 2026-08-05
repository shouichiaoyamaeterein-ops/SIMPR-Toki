Attribute VB_Name = "TEXTG53"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   TEXT処理
'*          財務データ累積
'*
'*   モジュールID  : TEXTG53.BAS
'*
'*
'*   作成日  :2001/10/30   By K.K
'*   変更日  :
'*---------------------------------------------------------------------------------
Global G530_Ki          As Byte
Global G530_SaveHead    As String
Dim G53_ID              As String
'+----------------------------------+
'+  テーブル出力                    +
'+----------------------------------+
Public Function TG53ConVertPut(ConvData As String, lErr As String) As Boolean
    lErr = ""
    
    TG53ConVertPut = False
    
    On Error GoTo TG53ConVertPut_Er:
    
    If Mid(ConvData, 1, 1) = "H" Then
       If Mid(ConvData, 12, 3) = "G53" Then
          G530_SaveHead = Mid(ConvData, 1, 30)
          ' 対象年月・財務管理番号を求める
          ' 契約先
          tx_KTorcd = TKT.契約先
          If CisFun.Val2(Mid(tx_HeadYmd, 7, 2)) > 15 Then            '上・下期判定
             G530_Ki = 0          ' 上期
          Else
             G530_Ki = 1          ' 下期
          End If
          If G530_Ki = 1 Then              '対象年月判定
             With CisFun
                 .Date_RsFormat = fYM
                 .Date_Add tx_HeadYmd, M, -1, 1
                 tx_TYM = .Date_Result
             End With
          Else
             tx_TYM = Mid(tx_HeadYmd, 1, 6)
          End If
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
           End With
       End If
       TG53ConVertPut = True
       Exit Function
    End If
    If tx_KNo = 0 Then
       ' 財務管理番号
       gSL_Select = "SELECT MAX(財務管理番号) 財務管理番号 From TEXT買掛金テーブル"
       If TKKRead(gSL_Select, 1) Then
          tx_KNo = TKK.財務管理番号
       End If
       If TKT.オプション2 <> 0 Then
          With CisFun
              .Date_RsFormat = fYM
              .Date_Add tx_TYM, M, TKT.オプション2 * -1, LastDate
              gStr = .Date_Result
          End With
          With CisDB
              .SQL = "DELETE From TEXT買掛金テーブル"
              .SQL = .SQL & " WHERE 契約先 = '" & tx_KTorcd & "' "
              .SQL = .SQL & "   AND 対象年月 <= '" & gStr & "' "
              .SQL = .SQL & "   AND FILECODE = '" & TKT.FILECD & "' "
              Call .DBExec
          End With
          Call ProcLogPut("TEXT買掛金 削除 契約先=" & TKT.契約先 & ",対象年月 <= " & gStr & ",削除件数=" & CisDB.RecordCount)
       End If
    End If
    
    If Mid(ConvData, 1, 1) <> "D" Then
       TG53ConVertPut = True
       Exit Function
    End If
        
    If TKT.TNSFMT = 0 Then
        ConvData = "D" & Mid(ConvData, 7)
    End If
    G53_ID = Mid(ConvData, 2, 1)

    Call ItemsClearTKK
    With TKK
        .契約先 = tx_KTorcd
        .対象年月 = tx_TYM
        .対象期 = G530_Ki
        If G530_Ki = 0 Then
           .対象日 = "01"
        Else
           .対象日 = "16"
        End If
        .HEAD作成日 = tx_HeadYmd
        .HEADSEQ = tx_HeadSeq
        .FILECODE = TKT.FILECD
        .FCD = TKT.FCODE
        
        .識別子 = G53_ID
        .データ = RTrim(Mid(ConvData, 8, 2))
        .カード = RTrim(Mid(ConvData, 10, 2))
        .購買 = RTrim(Mid(ConvData, 12, 2))
        .品番 = RTrim(Mid(ConvData, 14, 12))
        .管理番号 = RTrim(Mid(ConvData, 26, 11))
        .納入番号 = RTrim(Mid(ConvData, 37, 10))
        .貸借 = RTrim(Mid(ConvData, 47, 1))
        .単価 = CisFun.Val2(RTrim(Mid(ConvData, 58, 13))) / 100
        .金額 = CisFun.Val2(RTrim(Mid(ConvData, 71, 11)))
        ' 売掛金
         If RTrim(Mid(ConvData, 8, 1)) = "2" Then
            .数量 = CisFun.Val2(RTrim(Mid(ConvData, 48, 10))) / 1000
            .支給先 = RTrim(Mid(ConvData, 86, 4))
        Else    ' 買掛金
            .マイナス = RTrim(Mid(ConvData, 48, 1))
            .数量 = CisFun.Val2(RTrim(Mid(ConvData, 49, 9))) / 1000
            .訂正理由 = RTrim(Mid(ConvData, 90, 1))
            .持込先 = RTrim(Mid(ConvData, 91, 5))
        End If
        If .マイナス = "-" Then
            .金額 = .金額 * -1
        End If
        
        cTKK_KYM = Mid(ConvData, 82, 4)
        cTKK_YMD = Format(Date, "YYYY") & cTKK_KYM
        If Format(Date, "YYYYMM") < Mid(cTKK_YMD, 1, 6) Then
            cTKK_YMD = Format(Date, "YYYY") - 1 & cTKK_KYM
        End If
        .検収日 = cTKK_YMD
        .仕入先 = RTrim(Mid(ConvData, 4, 4))
        
        tx_KNo = tx_KNo + 1
        .財務管理番号 = tx_KNo
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

    Call TKKInsert
    tx_OutCnt = tx_OutCnt + 1

TG53ConVertPut_ED:
    
    TG53ConVertPut = True
    On Error GoTo 0
    
    Exit Function
TG53ConVertPut_Er:
    Static ER_MSG1 As String
    Static ER_MSG2 As String
    
    ER_MSG1 = "         財務データ変換処理（Ｇ５３）      " & vbCrLf & " " & vbCrLf
    ER_MSG1 = ER_MSG1 & "【 " & Err.Description & " 】"

    Call ProcLogPut(ER_MSG1)

    ER_MSG2 = Chr(13) & Space(3) & _
               "上記エラーにより、以降の処理を続行出来ません" & _
               Chr(10) & Chr(13) & Space(8) & _
               "内容を確認して再度実行して下さい"
    lErr = "ZG"
    Call TERPut(lErr, ER_MSG1, ER_MSG2)
End Function

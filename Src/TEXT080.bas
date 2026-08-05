Attribute VB_Name = "TEXT080"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   TEXT処理
'*          検収データ（有償支給品）
'*
'*   モジュールID  : TEXT080.BAS
'*
'*
'*   作成日  :2004/11/09   By CIS
'*   変更日  :2004/12/10   By CIS 品番頭２桁が"A5"の時、重量を金額計算に含める
'*   変更日  :2005/02/02   By CIS 2004/12/10変更分の計算式変更
'*   変更日  : 2005/03/31  By CIS 計算方法を取引先マスタの設定に従う
'*---------------------------------------------------------------------------------
Dim T080_Cnt         As Long
Dim T080_ST          As Long
Global T080_RCnt     As Long
'+----------------------------------+
'+  テーブル出力
'+----------------------------------+
Public Function T080ConVertPut(ConvData As String, lErr As String) As Boolean
    lErr = ""
    
    T080ConVertPut = False
    
    On Error GoTo T080ConVertPut_Er:
    
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
        .単位 = RTrim(Mid(ConvData, 24, 2))
        .重量 = CisFun.Val2(Mid(ConvData, 37, 6)) / 1000
        .単価 = CisFun.Val2(Mid(ConvData, 29, 8)) / 100
        .支給先 = Trim(Mid(ConvData, 1, 4))
        .出庫者 = Trim(Mid(ConvData, 18, 4))
        .支給区分 = RTrim(Mid(ConvData, 43, 1))
        .管理番号 = RTrim(Mid(ConvData, 44, 6))
        
        .仕入先 = RTrim(Mid(ConvData, 1, 4))
        .データ元 = RTrim(Mid(ConvData, 5, 1))
        .補助品番 = RTrim(Mid(ConvData, 235, 12))
        
        For T080_Cnt = 1 To 12
            T080_ST = 55 + (T080_Cnt - 1) * 15
            If RTrim(Mid(ConvData, T080_ST, 4)) = "" Then Exit For
            cTKK_KYM = Mid(ConvData, T080_ST, 4)
            cTKK_YMD = Format(Date, "YYYY") & cTKK_KYM
            If Format(Date, "YYYYMM") < Mid(cTKK_YMD, 1, 6) Then
                cTKK_YMD = Format(Date, "YYYY") - 1 & cTKK_KYM
            End If
            .検収日 = cTKK_YMD
            
            .納入番号 = RTrim(Mid(ConvData, T080_ST + 4, 5))
            .自動遡り区分 = RTrim(Mid(ConvData, T080_ST + 9, 1))
            .数量 = CisFun.Val2(Mid(ConvData, T080_ST + 10, 5))
'================================================================ 2005/03/31 Delete
''--------------------------------------------------------------- Insert 2004/12/10
'            If Mid(.品番, 1, 2) = "A5" And .重量 <> 0 Then
'
''-----     -----     -----     -----     -----     -----  < 2005.02.02 MNT START >
''              cTTK_Kin = .重量 * .単価 * .数量
'               cTTK_Kin = Fix(.重量 * .数量) * .単価
''-----     -----     -----     -----     -----     -----  < 2005.02.02 MNT  END  >
'
'            Else
'               cTTK_Kin = .単価 * .数量
'            End If
'            .金額 = Fix(cTTK_Kin)       ' 小数点以下切り捨て
''--------------------------------------------------------------- Insert 2004/12/10 End
''--------------------------------------------------------------- Delete 2004/12/10
''            .金額 = Fix(.単価 * .数量) ' 小数点以下切り捨て
''--------------------------------------------------------------- Delete 2004/12/10 End
''            cTTK_Kin = Fix(.単価 * .数量) ' 小数点以下切り捨て
'            '.金額 = Fix(cTTK_Kin)       ' 小数点以下切り捨て
'================================================================ 2005/03/31 Delete End
            
'================================================================ 2005/03/31 Insert
            If Mid(.品番, 1, 2) = "A5" And .重量 <> 0 Then
               cTTK_Kin = Fix(.重量 * .数量)
               cTTK_Kin = cTTK_Kin * .単価
            Else
               cTTK_Kin = .単価 * .数量
            End If
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
'================================================================ 2005/03/31 Insert End
            
            If .識別子 = "71" Then
               .金額 = .金額 * -1
               .マイナス = "-"
            End If
            tx_KNo = tx_KNo + 1
            .財務管理番号 = tx_KNo

            Call TKKInsert
            
            tx_OutCnt = tx_OutCnt + 1
        Next T080_Cnt

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
T080ConVertPut_ED:
    
    T080ConVertPut = True
    On Error GoTo 0
    
    Exit Function
T080ConVertPut_Er:
    Static ER_MSG1 As String
    Static ER_MSG2 As String
    
    ER_MSG1 = "         財務データ変換処理（０８０）      " & vbCrLf & " " & vbCrLf
    ER_MSG1 = ER_MSG1 & "【 " & Err.Description & " 】"

    Call ProcLogPut(ER_MSG1)

    ER_MSG2 = Chr(13) & Space(3) & _
               "上記エラーにより、以降の処理を続行出来ません" & _
               Chr(10) & Chr(13) & Space(8) & _
               "内容を確認して再度実行して下さい"
    lErr = "Z8"
    Call TERPut(lErr, ER_MSG1, ER_MSG2)
End Function



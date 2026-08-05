Attribute VB_Name = "CMA0010LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   モジュールID  : CMA0010LB.BAS
'*
'*   作成日  :2008.09.11   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public Type typData
    納入先              As String
    納入事業所        As String
    納入指示日          As String
    納入便              As String
    機番                As String
    カード丼            As String
    支給元              As String
    受注者              As String
    受事                As String
    納入番号            As String
    ページNO            As String
    総ページ数          As String
    明細行数            As String
    発注区分            As String
    品番()              As String
    背番号()            As String
    収容数()            As String
    箱数()              As String
    納入数()            As String
End Type
Public typQRAfter2       As typData

Public Sub QRtypDataClear2(ByRef typDataClear2 As typData)
    Dim cnt     As Integer
    
    With typDataClear2
        .納入先 = ""
        .納入事業所 = ""
        .納入指示日 = ""
        .納入便 = ""
        .機番 = ""
        .カード丼 = ""
        .支給元 = ""
        .受注者 = ""
        .受事 = ""
        .納入番号 = ""
        .ページNO = ""
        .総ページ数 = ""
        .明細行数 = ""
        .発注区分 = ""
        ReDim .品番(15)
        ReDim .背番号(15)
        ReDim .収容数(15)
        ReDim .箱数(15)
        ReDim .納入数(15)
        For cnt = 0 To 15
            .品番(cnt) = ""
            .背番号(cnt) = ""
            .収容数(cnt) = ""
            .箱数(cnt) = ""
            .納入数(cnt) = ""
        Next
        End With
End Sub

'+------------------------------------+
'+      QRデータ変換                   +
'+------------------------------------+
Public Function QRConvert_MT(QRString As String) As String

    Dim strRS   As String
    Dim strGS   As String
    Dim strEOT  As String
    Dim QRKind  As Long     ' QR種類 0:ﾁｹｯﾄ,1:JAMA
    Dim wDI     As String
    Dim wValue  As String
    Dim m_Cnt   As Integer
    
    Dim wRowSu  As Integer  ' 行数
    Dim wMojSu  As Integer  ' 明細文字数
    Dim wk_Head As String   ' ヘッダー部抜出データ
    Dim wk_Body As String   ' ボディ部抜出データ
    Dim wk_ct1  As Integer  ' 汎用カウンタ
    Dim wk_ct2  As Integer  ' 汎用カウンタ
    Dim wk      As String   ' 汎用文字列処理変数
    
    wRowSu = 14             ' 行数
    wMojSu = 101            ' 101文字間隔
    QRConvert_MT = ""
    
    strRS = Chr(&H1E)
    strGS = Chr(&H1D)
    strEOT = Chr(&H4)
    
    Call QRtypDataClear2(typQRAfter2)
    
    QRString = Replace(QRString, vbCr, "")
    
    wk_Head = Mid(QRString, 1, 38)      ' 01～34文字目：ヘッダ
    wk_Body = Mid(QRString, 39)         ' 39文字目以降：明細データ
'Call Test_LogOutput(wk_Body, 0)

    With typQRAfter2
        Dim GS_Start    As Long
        Dim GS_End      As Long
        Dim wstrdata    As String
        
        wk = Trim(Mid(wk_Head, 17, 5))
        .納入先 = Mid(wk + Space(4), 1, 4)      ' 納入先    (5→4)
        .納入事業所 = Space(4)                  ' 納入事業所(4)
        .納入指示日 = Mid(wk_Head, 6, 8)        ' 納入指示日(8)
        .納入便 = Format(Trim(Mid(wk_Head, 33, 2)), "00")       ' 納入便　　(2)
        .機番 = Space(3)                        ' 機番　　　(3)
        .カード丼 = Space(1)                    ' ＃　　　　(1)
        .支給元 = Space(4)                      ' 支給元　　(4)
        wk = Trim(Mid(wk_Head, 22, 5))
        .受注者 = Mid(wk + Space(4), 1, 4)      ' 受注者　　(5→4)
        .受事 = Space(1)                        ' 受事　　　(1)
        .納入番号 = Mid(wk_Head, 1, 5)          ' 納入番号　(5)
        .ページNO = "01"                        ' ページ S  (2)
        .総ページ数 = "01"                      ' ページ E  (2)
        .明細行数 = "00"                        ' 入力行数  (2)
        .発注区分 = "0"                         ' 発注区分  (1)

        '明細行データを取り出す
        For wk_ct1 = 0 To wRowSu
            wk_ct2 = wk_ct1 * wMojSu            ' 明細の先頭位置を求める
            If wk_ct2 + 1 > Len(wk_Body) Then   ' 先頭位置が全体の文字数を超えたら終了
                Exit For
            End If
            .品番(wk_ct1) = Mid(wk_Body, wk_ct2 + 16, 12) + Space(6)
            .背番号(wk_ct1) = Mid(wk_Body, wk_ct2 + 28, 3) + Space(1)
            .収容数(wk_ct1) = Format(Mid(wk_Body, wk_ct2 + 31, 5), "#####") ' + Space(1)
            .箱数(wk_ct1) = Mid(wk_Body, wk_ct2 + 36, 4)
            .納入数(wk_ct1) = Mid(wk_Body, wk_ct2 + 50, 6)
        Next
'Call Test_LogOutput(wk_Body, 0)

        'ヘッダの明細行数をセットする
        .明細行数 = Format(wk_ct1, "00")     ' 入力行数  (2)
    End With
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
    Print #f_No, "[ PMA0010 ] " & dt01 & _
                 " ----------------------------------------" & Chr(13) & Chr(10) & _
                 str & Chr(13) & Chr(10)        'ファイルへ書き込む"
    Close #f_No                                 'ファイルを閉じる
End Sub




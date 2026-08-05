Attribute VB_Name = "CKH0940LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   発注打切リスト
'*
'*   モジュールID  : CKH0940LB
'*
'*                  CKH0940LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not CKH0940LPrint Then Exit Sub
'
'*   作成日  :2004/01/07   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCKH0940        As New CISFormPrint
Public CKH0940_PrtCnt       As Long
Public CHH0940_CallType     As String

Public Function CKH0940LPrint() As Boolean
    CKH0940LPrint = False
    
' 出力データ存在確認
    gSL_Select = "SELECT * FROM 帳票出力テーブル"
    gSL_Select = gSL_Select & " Where 端末番号 = Host_Name()"
    gSL_Select = gSL_Select & "   And 帳票区分 = 4"
    
    If Not CHTRead(gSL_Select, 1) Then
       CKH0940LPrint = True
       Exit Function
    End If

    If CHH0940_CallType = "C" Then
    ' かんばん読取エラーテーブル 更新
        With CisDB
            .SQL = "UPDATE かんばん読取エラーテーブル Set "
            .SQL = .SQL & " 今回発注枚数 = CASE WHEN かんばん読取エラーテーブル.打切種別 = 0 Then"
            .SQL = .SQL & "                 ( Select SUM(読取枚数) From 発注かんばん枚数累積テーブル HK "
            .SQL = .SQL & "                        Where HK.処理端末 = Host_Name()"
            .SQL = .SQL & "                            And HK.品番 = かんばん読取エラーテーブル.品番 )"
            .SQL = .SQL & "         Else"
            .SQL = .SQL & "                 ( Select SUM(読取枚数) From 発注かんばん枚数累積テーブル HK "
            .SQL = .SQL & "                        Where HK.処理端末 = Host_Name()"
            .SQL = .SQL & "                            And HK.品番 = かんばん読取エラーテーブル.品番"
            .SQL = .SQL & "                            And HK.取引先 = かんばん読取エラーテーブル.取引先"
            .SQL = .SQL & "                            And HK.受入 = かんばん読取エラーテーブル.受入"
            .SQL = .SQL & "                            And HK.収容数 = かんばん読取エラーテーブル.収容数 ) End "
            .SQL = .SQL & "    Where かんばん読取エラーテーブル.エラー管理番号"
            .SQL = .SQL & "           = ( Select 管理No From 帳票出力テーブル"
            .SQL = .SQL & "                           Where 端末番号 = Host_Name()"
            .SQL = .SQL & "                             And 帳票区分 = 4"
            .SQL = .SQL & "                           Group by 管理No )"
            .DBExec
        End With
    End If
    
    gStr = "SELECT ISNULL(TR.略称,'') 取引先名,"
    gStr = gStr & " CASE WHEN Isnull(HN.表示品番,'') = '' THEN KE.品番 Else HN.表示品番 End 表示品番,"
    gSL_Select = gStr
    gSL_Select = gSL_Select & "KE.取引先,KE.受入,KE.背番号,KE.品番,KE.収容数,"
    gSL_Select = gSL_Select & "KE.打切種別,KE.打切区分,KE.起算日,KE.起算便,KE.完了日,KE.完了便,'' 打切日,"
    gSL_Select = gSL_Select & "KE.打切設定数,KE.今回発注枚数,KE.今回発注端数,KE.打切エラー枚数"
    gSL_Select = gSL_Select & "  From 帳票出力テーブル CH"
    gSL_Select = gSL_Select & "  Left Outer Join かんばん読取エラーテーブル KE"
    gSL_Select = gSL_Select & "      ON  KE.エラー管理番号 = CH.管理NO"
    gSL_Select = gSL_Select & "      And KE.エラー種類 = CH.帳票区分"
    gSL_Select = gSL_Select & "  Left Outer Join 品番マスタ HN"
    gSL_Select = gSL_Select & "      ON  HN.品番 = KE.品番"
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "      ON  TR.取引先CD = KE.取引先"
    gSL_Select = gSL_Select & "      And TR.取引先区分 = 1"
    gSL_Select = gSL_Select & " Where CH.端末番号 = Host_Name()"
    gSL_Select = gSL_Select & "   And CH.帳票区分 = 4"
    gSL_Select = gSL_Select & "   And KE.打切種別 = 0"
    gSL_Select = gSL_Select & " Group by TR.略称,表示品番,"
    gSL_Select = gSL_Select & "KE.取引先,KE.受入,KE.背番号,KE.品番,KE.収容数,"
    gSL_Select = gSL_Select & "KE.打切種別,KE.打切区分,KE.起算日,KE.起算便,KE.完了日,KE.完了便,"
    gSL_Select = gSL_Select & "KE.打切設定数,KE.今回発注枚数,KE.今回発注端数,KE.打切エラー枚数"
    
    gSL_Select = gSL_Select & " Order By"
    gSL_Select = gSL_Select & " KE.取引先,KE.受入,KE.背番号,"
    gSL_Select = gSL_Select & " KE.品番,KE.収容数,KE.打切種別"
    
    If Not KRERead(gSL_Select, 0, 1) Then
       Call KREClose
       CKH0940LPrint = True
       Exit Function
    End If
    
    CKH0940_PrtCnt = CisDB.RecordCount
    
    With CisPrtCKH0940
         Set .iPrtForm = CKH0940L
         .dpMSG = "発注打切リスト　印刷中"
         .iPreview = gDemoPrt
         .dpMSGScreen = モーダル
         .PrintStart
    End With
    
    CKH0940LPrint = True
End Function




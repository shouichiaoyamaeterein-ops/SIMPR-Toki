Attribute VB_Name = "CKHZ0940LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   発注打切リスト
'*
'*   モジュールID  : CKHZ0940LB
'*
'*                  CKHZ0940LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not CKHZ0940LPrint Then Exit Sub
'
'*   作成日  :2004/02/21   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCKHZ0940        As New CISFormPrint
Public CKHZ0940_PrtCnt       As Long
Public CHHZ0940_CallType     As String

Public Function CKHZ0940LPrint() As Boolean
    CKHZ0940LPrint = False
    
' 出力データ存在確認
    gSL_Select = "SELECT * FROM 帳票出力テーブル"
    gSL_Select = gSL_Select & " Where 端末番号 = Host_Name()"
    gSL_Select = gSL_Select & "   And 帳票区分 = 24"
    
    If Not CHTRead(gSL_Select, 1) Then
       CKHZ0940LPrint = True
       Exit Function
    End If

    If CHHZ0940_CallType = "C" Then
    ' 材料かんばん読取エラーテーブル 更新
        With CisDB
            .SQL = "UPDATE 材料かんばん読取エラーテーブル Set "
            .SQL = .SQL & " 今回発注枚数 = ( Select SUM(読取枚数) From 材料発注かんばん枚数累積テーブル HK "
            .SQL = .SQL & "                        Where HK.処理端末 = Host_Name()"
            .SQL = .SQL & "                            And HK.材料管理番号 = 材料かんばん読取エラーテーブル.材料管理番号"
            .SQL = .SQL & "                            And HK.取引先 = 材料かんばん読取エラーテーブル.取引先"
            .SQL = .SQL & "                            And HK.受入 = 材料かんばん読取エラーテーブル.受入"
            .SQL = .SQL & "                            And HK.収容数 = 材料かんばん読取エラーテーブル.収容数 )"
            .SQL = .SQL & "    Where 材料かんばん読取エラーテーブル.エラー管理番号"
            .SQL = .SQL & "           = ( Select 管理No From 帳票出力テーブル"
            .SQL = .SQL & "                           Where 端末番号 = Host_Name()"
            .SQL = .SQL & "                             And 帳票区分 = 24"
            .SQL = .SQL & "                           Group by 管理No )"
            .DBExec
        End With
    End If
    
    gStr = "SELECT ISNULL(TR.略称,'') 取引先名,"
    gStr = gStr & " ZR.表示寸法,KE.材質,ZR.SIZE サイズ,"
    gSL_Select = gStr
    gSL_Select = gSL_Select & "KE.取引先,KE.受入,KE.背番号,KE.収容数,"
    gSL_Select = gSL_Select & "KE.打切種別,KE.打切区分,KE.起算日,KE.起算便,KE.完了日,KE.完了便,'' 打切日,"
    gSL_Select = gSL_Select & "KE.打切設定数,KE.今回発注枚数,KE.今回発注端数,KE.打切エラー枚数"
    gSL_Select = gSL_Select & "  From 帳票出力テーブル CH"
    gSL_Select = gSL_Select & "  Left Outer Join 材料かんばん読取エラーテーブル KE"
    gSL_Select = gSL_Select & "      ON  KE.エラー管理番号 = CH.管理NO"
    gSL_Select = gSL_Select & "      And KE.エラー種類 = CH.帳票区分"
    gSL_Select = gSL_Select & "  Left Outer Join 材料マスタ ZR"
    gSL_Select = gSL_Select & "      ON  ZR.材料管理番号 = KE.材料管理番号"
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "      ON  TR.取引先CD = KE.取引先"
    gSL_Select = gSL_Select & "      And TR.取引先区分 = 1"
    gSL_Select = gSL_Select & " Where CH.端末番号 = Host_Name()"
    gSL_Select = gSL_Select & "   And CH.帳票区分 = 24"
    gSL_Select = gSL_Select & "   And KE.打切種別 = 0"
    gSL_Select = gSL_Select & " Group by TR.略称,ZR.表示寸法,ZR.SIZE,"
    gSL_Select = gSL_Select & "KE.取引先,KE.受入,KE.背番号,KE.材質,KE.収容数,"
    gSL_Select = gSL_Select & "KE.打切種別,KE.打切区分,KE.起算日,KE.起算便,KE.完了日,KE.完了便,"
    gSL_Select = gSL_Select & "KE.打切設定数,KE.今回発注枚数,KE.今回発注端数,KE.打切エラー枚数"
    
    gSL_Select = gSL_Select & " Order By"
    gSL_Select = gSL_Select & " KE.取引先,KE.受入,KE.背番号,"
    gSL_Select = gSL_Select & " KE.材質,ZR.SIZE,KE.収容数,KE.打切種別"
    
    If Not ZRERead(gSL_Select, 0, 1) Then
       Call ZREClose
       CKHZ0940LPrint = True
       Exit Function
    End If
    
    CKHZ0940_PrtCnt = CisDB.RecordCount
    
    With CisPrtCKHZ0940
         Set .iPrtForm = CKHZ0940L
         .dpMSG = "材料発注打切リスト　印刷中"
         .iPreview = gDemoPrt
         .dpMSGScreen = モーダル
         .PrintStart
    End With
    
    CKHZ0940LPrint = True
End Function




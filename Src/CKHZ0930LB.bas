Attribute VB_Name = "CKHZ0930LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   材料発注打切警告リスト
'*
'*   モジュールID  : CKHZ0930LB
'*
'*                  CKHZ0930LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not CKHZ0930LPrint Then Exit Sub
'
'*   作成日  :2004/02/21   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCKHZ0930        As New CISFormPrint
Public CKHZ0930_PrtCnt       As Long
Public CHHZ0930_CallType     As String


Public Function CKHZ0930LPrint() As Boolean
    CKHZ0930LPrint = False
    
' 出力データ存在確認
    gSL_Select = "SELECT * FROM 帳票出力テーブル"
    gSL_Select = gSL_Select & " Where 端末番号 = Host_Name()"
    gSL_Select = gSL_Select & "   And 帳票区分 = 23"
    
    If Not CHTRead(gSL_Select, 1) Then
       CKHZ0930LPrint = True
       Exit Function
    End If

    If CHHZ0930_CallType = "C" Then
    ' かんばん読取エラーテーブル 更新
        With CisDB
            .SQL = "UPDATE 材料かんばん読取エラーテーブル Set "
            .SQL = .SQL & " 今回発注数 = CASE WHEN 材料かんばん読取エラーテーブル.打切種別 = 0 Then"
            .SQL = .SQL & "                 ( Select SUM(発注数) From 材料発注かんばん枚数累積テーブル HK "
            .SQL = .SQL & "                        Where HK.処理端末 = Host_Name()"
            .SQL = .SQL & "                            And HK.材料管理番号 = 材料かんばん読取エラーテーブル.材料管理番号 )"
            .SQL = .SQL & "         Else"
            .SQL = .SQL & "                 ( Select SUM(発注数) From 材料発注かんばん枚数累積テーブル HK "
            .SQL = .SQL & "                        Where HK.処理端末 = Host_Name()"
            .SQL = .SQL & "                            And HK.材料管理番号 = 材料かんばん読取エラーテーブル.材料管理番号"
            .SQL = .SQL & "                            And HK.取引先 = 材料かんばん読取エラーテーブル.取引先"
            .SQL = .SQL & "                            And HK.受入 = 材料かんばん読取エラーテーブル.受入"
            .SQL = .SQL & "                            And HK.収容数 = 材料かんばん読取エラーテーブル.収容数 ) End, "
            .SQL = .SQL & " 発注済数 = 発注済数 + 収容数"
            .SQL = .SQL & "    Where 材料かんばん読取エラーテーブル.エラー管理番号"
            .SQL = .SQL & "           = ( Select 管理No From 帳票出力テーブル"
            .SQL = .SQL & "                           Where 端末番号 = Host_Name()"
            .SQL = .SQL & "                             And 帳票区分 = 23"
            .SQL = .SQL & "                           Group by 管理No )"
            .DBExec
        End With
    End If
    
    gStr = "SELECT ISNULL(TR.略称,'') 取引先名,"
    gStr = gStr & " ZR.表示寸法 ,ZR.SIZE サイズ,"
    gSL_Select = " SELECT '' 取引先名," & vbCr
    gSL_Select = gSL_Select & " ZR.表示寸法 ,ZR.SIZE サイズ," & vbCr
    gSL_Select = gSL_Select & "'' 取引先,'' 受入,KE.背番号,KE.収容数,KE.材質," & vbCr
    gSL_Select = gSL_Select & "KE.打切種別,KE.起算日,KE.起算便,'' 打切日,KE.打切設定数," & vbCr
    gSL_Select = gSL_Select & "KE.警告数,'' 警告日,KE.発注済数,KE.今回発注数,ZU.設定残数" & vbCr
    gSL_Select = gSL_Select & "  From 帳票出力テーブル CH" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 材料かんばん読取エラーテーブル KE" & vbCr
    gSL_Select = gSL_Select & "      ON  KE.エラー管理番号 = CH.管理NO" & vbCr
    gSL_Select = gSL_Select & "      And KE.エラー種類 = CH.帳票区分" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 材料マスタ ZR" & vbCr
    gSL_Select = gSL_Select & "      ON  ZR.材料管理番号 = KE.材料管理番号" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 材料発注打切テーブル ZU" & vbCr
    gSL_Select = gSL_Select & "      ON  ZU.打切種別 = KE.打切種別" & vbCr
    gSL_Select = gSL_Select & "      And ZU.材料管理番号 = KE.材料管理番号" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR" & vbCr
    gSL_Select = gSL_Select & "      ON  TR.取引先CD = KE.取引先" & vbCr
    gSL_Select = gSL_Select & "      And TR.取引先区分 = 1" & vbCr
    gSL_Select = gSL_Select & " Where CH.端末番号 = Host_Name()" & vbCr
    gSL_Select = gSL_Select & "   And CH.帳票区分 = 23" & vbCr
    gSL_Select = gSL_Select & "   And KE.打切種別 = 0" & vbCr
    gSL_Select = gSL_Select & " Group by ZR.表示寸法,ZR.SIZE," & vbCr
    gSL_Select = gSL_Select & "KE.背番号,KE.収容数,KE.材質," & vbCr
    gSL_Select = gSL_Select & "KE.打切種別,KE.起算日,KE.起算便,KE.打切設定数," & vbCr
    gSL_Select = gSL_Select & "KE.警告数,KE.発注済数,KE.今回発注数,ZU.設定残数" & vbCr
    
    gSL_Select = gSL_Select & " Union " & vbCr
    gSL_Select = gSL_Select & " SELECT '' 取引先名," & vbCr
    gSL_Select = gSL_Select & " ZR.表示寸法 ,ZR.SIZE サイズ," & vbCr
    gSL_Select = gSL_Select & "'' 取引先,'' 受入,'' 背番号,0 収容数,KE.材質," & vbCr
    gSL_Select = gSL_Select & "KE.打切種別,'' 起算日,'' 起算便,KE.打切日,0 打切設定数," & vbCr
    gSL_Select = gSL_Select & "0 警告数,KE.警告日 警告日,0 発注済数,SUM(KE.今回発注数) 今回発注数,0 設定残数" & vbCr
    gSL_Select = gSL_Select & "  From 帳票出力テーブル CH" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 材料かんばん読取エラーテーブル KE" & vbCr
    gSL_Select = gSL_Select & "      ON  KE.エラー管理番号 = CH.管理NO" & vbCr
    gSL_Select = gSL_Select & "      And KE.エラー種類 = CH.帳票区分" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 材料マスタ ZR" & vbCr
    gSL_Select = gSL_Select & "      ON  ZR.材料管理番号 = KE.材料管理番号" & vbCr
    gSL_Select = gSL_Select & " Where CH.端末番号 = Host_Name()" & vbCr
    gSL_Select = gSL_Select & "   And CH.帳票区分 = 23" & vbCr
    gSL_Select = gSL_Select & "   And KE.打切種別 = 1" & vbCr
    gSL_Select = gSL_Select & " Group by 表示寸法,ZR.SIZE,KE.材質," & vbCr
    gSL_Select = gSL_Select & "KE.打切種別,KE.打切日,KE.警告日" & vbCr
'
    gSL_Select = gSL_Select & " Union " & vbCr
    gSL_Select = gSL_Select & gStr & vbCr
    gSL_Select = gSL_Select & "KE.取引先,'' 受入,'' 背番号,0 収容数,KE.材質," & vbCr
    gSL_Select = gSL_Select & "KE.打切種別,'' 起算日,'' 起算便,KE.打切日,0 打切設定数," & vbCr
    gSL_Select = gSL_Select & "0 警告数,KE.警告日 警告日,SUM(KE.発注済数) 発注済数,SUM(KE.今回発注数) 今回発注数,0 設定残数" & vbCr
    gSL_Select = gSL_Select & "  From 帳票出力テーブル CH" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 材料かんばん読取エラーテーブル KE" & vbCr
    gSL_Select = gSL_Select & "      ON  KE.エラー管理番号 = CH.管理NO" & vbCr
    gSL_Select = gSL_Select & "      And KE.エラー種類 = CH.帳票区分" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 材料マスタ ZR" & vbCr
    gSL_Select = gSL_Select & "      ON  ZR.材料管理番号 = KE.材料管理番号" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR" & vbCr
    gSL_Select = gSL_Select & "      ON  TR.取引先CD = KE.取引先" & vbCr
    gSL_Select = gSL_Select & "      And TR.取引先区分 = 1" & vbCr
    gSL_Select = gSL_Select & " Where CH.端末番号 = Host_Name()" & vbCr
    gSL_Select = gSL_Select & "   And CH.帳票区分 = 23" & vbCr
    gSL_Select = gSL_Select & "   And KE.打切種別 = 2" & vbCr
    gSL_Select = gSL_Select & " Group by TR.略称,表示寸法,ZR.SIZE,KE.材質," & vbCr
    gSL_Select = gSL_Select & "KE.取引先," & vbCr
    gSL_Select = gSL_Select & "KE.打切種別,KE.打切日,KE.警告日" & vbCr
'
    gSL_Select = gSL_Select & " Union " & vbCr
    gSL_Select = gSL_Select & "SELECT ISNULL(TR.略称,'') 取引先名," & vbCr
    gSL_Select = gSL_Select & " '' 表示寸法,'' サイズ," & vbCr
    gSL_Select = gSL_Select & "KE.取引先,'' 受入,'' 背番号,0 収容数,'' 材質," & vbCr
    gSL_Select = gSL_Select & "KE.打切種別,'' 起算日,'' 起算便,KE.打切日,0 打切設定数," & vbCr
    gSL_Select = gSL_Select & "0 警告数,KE.警告日 警告日,SUM(KE.発注済数) 発注済数,SUM(KE.今回発注数) 今回発注数,0 設定残数" & vbCr
    gSL_Select = gSL_Select & "  From 帳票出力テーブル CH" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 材料かんばん読取エラーテーブル KE" & vbCr
    gSL_Select = gSL_Select & "      ON  KE.エラー管理番号 = CH.管理NO" & vbCr
    gSL_Select = gSL_Select & "      And KE.エラー種類 = CH.帳票区分" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 材料マスタ ZR" & vbCr
    gSL_Select = gSL_Select & "      ON  ZR.材料管理番号 = KE.材料管理番号" & vbCr
    gSL_Select = gSL_Select & "  Left Outer Join 取引先マスタ TR" & vbCr
    gSL_Select = gSL_Select & "      ON  TR.取引先CD = KE.取引先" & vbCr
    gSL_Select = gSL_Select & "      And TR.取引先区分 = 1" & vbCr
    gSL_Select = gSL_Select & " Where CH.端末番号 = Host_Name()" & vbCr
    gSL_Select = gSL_Select & "   And CH.帳票区分 = 23" & vbCr
    gSL_Select = gSL_Select & "   And KE.打切種別 = 3" & vbCr
    gSL_Select = gSL_Select & " Group by TR.略称," & vbCr
    gSL_Select = gSL_Select & "KE.取引先," & vbCr
    gSL_Select = gSL_Select & "KE.打切種別,KE.打切日,KE.警告日" & vbCr
    
    gSL_Select = gSL_Select & " Order By" & vbCr
    gSL_Select = gSL_Select & " KE.取引先,KE.受入,KE.背番号," & vbCr
    gSL_Select = gSL_Select & " KE.材質,ZR.SIZE,KE.収容数,KE.打切種別" & vbCr
    
    If Not ZRERead(gSL_Select, 0, 1) Then
       Call ZREClose
       CKHZ0930LPrint = True
       Exit Function
    End If
    
    
    CKHZ0930_PrtCnt = CisDB.RecordCount
    
    With CisPrtCKHZ0930
         Set .iPrtForm = CKHZ0930L
         .dpMSG = "材料発注打切警告リスト　印刷中"
         .iPreview = gDemoPrt
         .dpMSGScreen = モーダル
         .PrintStart
    End With
    
    CKHZ0930LPrint = True
End Function




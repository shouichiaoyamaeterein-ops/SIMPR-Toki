Module MST_単価売上
'*--------------------------------------------------------------------------------
'*
'*   .Net  仕様
'*
'*   MST_単価売上
'*
'*   モジュールID  : MST_単価売上.vb
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*                    ※ 追加時は"ﾘﾝｸとして追加"を行って下さい。	
'*					/項目追加含む/
'*
'*   作成日  :  2019/09/10  By Cis(CDH00301)
'*   変更日  :
'*---------------------------------------------------------------------------------
Structure  T_MST_単価売上
    Dim 品番  As String
    Dim 契約先  As String
    Dim 履歴NO  As Integer
    Dim 号補区分  As String
    Dim 納入番号  As String
    Dim 適用開始日  As String
    Dim 適用終了日  As String
    Dim 旧型区分  As Byte
    Dim 仮区分  As String
    Dim 単価  As Decimal
    Dim 単価変更理由  As String
    Dim 梱包区分  As Byte
    Dim 梱包単価  As Decimal
    Dim 備考  As String
    Dim 設変符号  As String
    Dim 納入先  As String
    Dim 作成日  As Date
    Dim 作成者  As String
    Dim 作成端末  As String
    Dim 更新日  As Date
    Dim 更新者  As String
    Dim 更新端末  As String
        Dim 件数 As Long
'   項目追加(自動)
        Dim 内訳単価1      As Decimal
        Dim 内訳単価2      As Decimal
        Dim 内訳単価3      As Decimal
        Dim 内訳単価4      As Decimal
        Dim 内訳単価5      As Decimal
        Dim 内訳単価6      As Decimal
        Dim 内訳単価7      As Decimal
        Dim 内訳単価8      As Decimal
        Dim 内訳単価9      As Decimal
        Dim 内訳単価10     As Decimal
        Dim 納入先名       As String
        Dim 今期単価       As Decimal
        Dim 来期単価       As Decimal
        Dim 取引先名       As String
        Dim 今期区分       As String
        Dim 来期区分       As String
        Dim 契約先名       As String
        Dim 現旧区分名     As String
        Dim 単価変更理由名 As String
        Dim 号補区分名     As String
        Dim 表示品番       As String
        Dim Warp単価       As Decimal
        Dim Warp比較       As Decimal
        Dim 改定単価       As Decimal
        Dim 内訳単価計     As Decimal
        Dim 品名           As String
        Dim 現在内訳単価1  As Decimal
        Dim 現在内訳単価2  As Decimal
        Dim 現在内訳単価3  As Decimal
        Dim 現在内訳単価4  As Decimal
        Dim 現在内訳単価5  As Decimal
        Dim 現在内訳単価6  As Decimal
        Dim 現在内訳単価7  As Decimal
        Dim 現在内訳単価8  As Decimal
        Dim 現在内訳単価9  As Decimal
        Dim 現在内訳単価10 As Decimal
        Dim 現在仮区分     As String
End Structure

Public JTM As T_MST_単価売上

Public JTM_RDSTS    As Boolean     	    ' 専用読取フラグ
Public JTM_DBNAME   As String = ""           ' 指定DBNAME
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ        *
'*----------------------------------------*'
Public Sub ItemsFromJTM()
'   退避エリアクリア
    Call ItemsClearJTM	
    With CisDB
         For .ItemCnt = 0 To .SqlDataRow.Table.Columns.Count - 1 	
             Select Case .SqlDataRow.Table.Columns(.ItemCnt).Caption
                    Case "品番":          If Not IsDBNull(.SqlDataRow("品番")) Then JTM.品番 = CStr(.SqlDataRow("品番"))
                    Case "契約先":        If Not IsDBNull(.SqlDataRow("契約先")) Then JTM.契約先 = CStr(.SqlDataRow("契約先"))
                    Case "履歴NO":        If Not IsDBNull(.SqlDataRow("履歴NO")) Then JTM.履歴NO = CInt(.SqlDataRow("履歴NO"))
                    Case "号補区分":      If Not IsDBNull(.SqlDataRow("号補区分")) Then JTM.号補区分 = CStr(.SqlDataRow("号補区分"))
                    Case "納入番号":      If Not IsDBNull(.SqlDataRow("納入番号")) Then JTM.納入番号 = CStr(.SqlDataRow("納入番号"))
                    Case "適用開始日":    If Not IsDBNull(.SqlDataRow("適用開始日")) Then JTM.適用開始日 = CStr(.SqlDataRow("適用開始日"))
                    Case "適用終了日":    If Not IsDBNull(.SqlDataRow("適用終了日")) Then JTM.適用終了日 = CStr(.SqlDataRow("適用終了日"))
                    Case "旧型区分":      If Not IsDBNull(.SqlDataRow("旧型区分")) Then JTM.旧型区分 = CByte(.SqlDataRow("旧型区分"))
                    Case "仮区分":        If Not IsDBNull(.SqlDataRow("仮区分")) Then JTM.仮区分 = CStr(.SqlDataRow("仮区分"))
                    Case "単価":          If Not IsDBNull(.SqlDataRow("単価")) Then JTM.単価 = CDec(.SqlDataRow("単価"))
                    Case "単価変更理由":  If Not IsDBNull(.SqlDataRow("単価変更理由")) Then JTM.単価変更理由 = CStr(.SqlDataRow("単価変更理由"))
                    Case "梱包区分":      If Not IsDBNull(.SqlDataRow("梱包区分")) Then JTM.梱包区分 = CByte(.SqlDataRow("梱包区分"))
                    Case "梱包単価":      If Not IsDBNull(.SqlDataRow("梱包単価")) Then JTM.梱包単価 = CDec(.SqlDataRow("梱包単価"))
                    Case "備考":          If Not IsDBNull(.SqlDataRow("備考")) Then JTM.備考 = CStr(.SqlDataRow("備考"))
                    Case "設変符号":      If Not IsDBNull(.SqlDataRow("設変符号")) Then JTM.設変符号 = CStr(.SqlDataRow("設変符号"))
                    Case "納入先":        If Not IsDBNull(.SqlDataRow("納入先")) Then JTM.納入先 = CStr(.SqlDataRow("納入先"))
                    Case "作成日":        If Not IsDBNull(.SqlDataRow("作成日")) Then JTM.作成日 = CDate(.SqlDataRow("作成日"))
                    Case "作成者":        If Not IsDBNull(.SqlDataRow("作成者")) Then JTM.作成者 = CStr(.SqlDataRow("作成者"))
                    Case "作成端末":      If Not IsDBNull(.SqlDataRow("作成端末")) Then JTM.作成端末 = CStr(.SqlDataRow("作成端末"))
                    Case "更新日":        If Not IsDBNull(.SqlDataRow("更新日")) Then JTM.更新日 = CDate(.SqlDataRow("更新日"))
                    Case "更新者":        If Not IsDBNull(.SqlDataRow("更新者")) Then JTM.更新者 = CStr(.SqlDataRow("更新者"))
                    Case "更新端末":      If Not IsDBNull(.SqlDataRow("更新端末")) Then JTM.更新端末 = CStr(.SqlDataRow("更新端末"))
                Case "件数":         If Not IsDBNull(.SqlDataRow("件数")) Then JTM.件数 = CLng(.SqlDataRow("件数"))
'   項目追加(自動)
                Case "内訳単価1":	If Not IsDBNull(.SqlDataRow("内訳単価1")) Then JTM.内訳単価1	= CDec(.SqlDataRow("内訳単価1"))
                Case "内訳単価2":	If Not IsDBNull(.SqlDataRow("内訳単価2")) Then JTM.内訳単価2	= CDec(.SqlDataRow("内訳単価2"))
                Case "内訳単価3":	If Not IsDBNull(.SqlDataRow("内訳単価3")) Then JTM.内訳単価3	= CDec(.SqlDataRow("内訳単価3"))
                Case "内訳単価4":	If Not IsDBNull(.SqlDataRow("内訳単価4")) Then JTM.内訳単価4	= CDec(.SqlDataRow("内訳単価4"))
                Case "内訳単価5":	If Not IsDBNull(.SqlDataRow("内訳単価5")) Then JTM.内訳単価5	= CDec(.SqlDataRow("内訳単価5"))
                Case "内訳単価6":	If Not IsDBNull(.SqlDataRow("内訳単価6")) Then JTM.内訳単価6	= CDec(.SqlDataRow("内訳単価6"))
                Case "内訳単価7":	If Not IsDBNull(.SqlDataRow("内訳単価7")) Then JTM.内訳単価7	= CDec(.SqlDataRow("内訳単価7"))
                Case "内訳単価8":	If Not IsDBNull(.SqlDataRow("内訳単価8")) Then JTM.内訳単価8	= CDec(.SqlDataRow("内訳単価8"))
                Case "内訳単価9":	If Not IsDBNull(.SqlDataRow("内訳単価9")) Then JTM.内訳単価9	= CDec(.SqlDataRow("内訳単価9"))
                Case "内訳単価10":	If Not IsDBNull(.SqlDataRow("内訳単価10")) Then JTM.内訳単価10	= CDec(.SqlDataRow("内訳単価10"))
                Case "納入先名":	If Not IsDBNull(.SqlDataRow("納入先名")) Then JTM.納入先名	= CStr(.SqlDataRow("納入先名"))
                Case "今期単価":	If Not IsDBNull(.SqlDataRow("今期単価")) Then JTM.今期単価	= CDec(.SqlDataRow("今期単価"))
                Case "来期単価":	If Not IsDBNull(.SqlDataRow("来期単価")) Then JTM.来期単価	= CDec(.SqlDataRow("来期単価"))
                Case "取引先名":	If Not IsDBNull(.SqlDataRow("取引先名")) Then JTM.取引先名	= CStr(.SqlDataRow("取引先名"))
                Case "今期区分":	If Not IsDBNull(.SqlDataRow("今期区分")) Then JTM.今期区分	= CStr(.SqlDataRow("今期区分"))
                Case "来期区分":	If Not IsDBNull(.SqlDataRow("来期区分")) Then JTM.来期区分	= CStr(.SqlDataRow("来期区分"))
                Case "契約先名":	If Not IsDBNull(.SqlDataRow("契約先名")) Then JTM.契約先名	= CStr(.SqlDataRow("契約先名"))
                Case "現旧区分名":	If Not IsDBNull(.SqlDataRow("現旧区分名")) Then JTM.現旧区分名	= CStr(.SqlDataRow("現旧区分名"))
                Case "単価変更理由名":	If Not IsDBNull(.SqlDataRow("単価変更理由名")) Then JTM.単価変更理由名	= CStr(.SqlDataRow("単価変更理由名"))
                Case "号補区分名":	If Not IsDBNull(.SqlDataRow("号補区分名")) Then JTM.号補区分名	= CStr(.SqlDataRow("号補区分名"))
                Case "表示品番":	If Not IsDBNull(.SqlDataRow("表示品番")) Then JTM.表示品番	= CStr(.SqlDataRow("表示品番"))
                Case "Warp単価":	If Not IsDBNull(.SqlDataRow("Warp単価")) Then JTM.Warp単価	= CDec(.SqlDataRow("Warp単価"))
                Case "Warp比較":	If Not IsDBNull(.SqlDataRow("Warp比較")) Then JTM.Warp比較	= CDec(.SqlDataRow("Warp比較"))
                Case "改定単価":	If Not IsDBNull(.SqlDataRow("改定単価")) Then JTM.改定単価	= CDec(.SqlDataRow("改定単価"))
                Case "内訳単価計":	If Not IsDBNull(.SqlDataRow("内訳単価計")) Then JTM.内訳単価計	= CDec(.SqlDataRow("内訳単価計"))
                Case "品名":	If Not IsDBNull(.SqlDataRow("品名")) Then JTM.品名	= CStr(.SqlDataRow("品名"))
                Case "現在内訳単価1":	If Not IsDBNull(.SqlDataRow("現在内訳単価1")) Then JTM.現在内訳単価1	= CDec(.SqlDataRow("現在内訳単価1"))
                Case "現在内訳単価2":	If Not IsDBNull(.SqlDataRow("現在内訳単価2")) Then JTM.現在内訳単価2	= CDec(.SqlDataRow("現在内訳単価2"))
                Case "現在内訳単価3":	If Not IsDBNull(.SqlDataRow("現在内訳単価3")) Then JTM.現在内訳単価3	= CDec(.SqlDataRow("現在内訳単価3"))
                Case "現在内訳単価4":	If Not IsDBNull(.SqlDataRow("現在内訳単価4")) Then JTM.現在内訳単価4	= CDec(.SqlDataRow("現在内訳単価4"))
                Case "現在内訳単価5":	If Not IsDBNull(.SqlDataRow("現在内訳単価5")) Then JTM.現在内訳単価5	= CDec(.SqlDataRow("現在内訳単価5"))
                Case "現在内訳単価6":	If Not IsDBNull(.SqlDataRow("現在内訳単価6")) Then JTM.現在内訳単価6	= CDec(.SqlDataRow("現在内訳単価6"))
                Case "現在内訳単価7":	If Not IsDBNull(.SqlDataRow("現在内訳単価7")) Then JTM.現在内訳単価7	= CDec(.SqlDataRow("現在内訳単価7"))
                Case "現在内訳単価8":	If Not IsDBNull(.SqlDataRow("現在内訳単価8")) Then JTM.現在内訳単価8	= CDec(.SqlDataRow("現在内訳単価8"))
                Case "現在内訳単価9":	If Not IsDBNull(.SqlDataRow("現在内訳単価9")) Then JTM.現在内訳単価9	= CDec(.SqlDataRow("現在内訳単価9"))
                Case "現在内訳単価10":	If Not IsDBNull(.SqlDataRow("現在内訳単価10")) Then JTM.現在内訳単価10	= CDec(.SqlDataRow("現在内訳単価10"))
                Case "現在仮区分":	If Not IsDBNull(.SqlDataRow("現在仮区分")) Then JTM.現在仮区分	= CStr(.SqlDataRow("現在仮区分"))
             End Select
         Next .ItemCnt
    End With	
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ   *
'*----------------------------------------*'
    Public Function JTMRead(Optional ByVal ReadgSql As String = "", Optional ByVal ReadFirstGyo As Boolean = False, Optional ByVal ConnectNo As Byte = 1,Optional ByVal Init As Boolean = False) As Boolean
        With CisDB
            If ReadgSql = "" And .Sql.Trim = "" Then
               .Sql = "Select * From " 
               If JTM_DBNAME <> "" Then
                  .Sql += JTM_DBNAME & ".."
               End If
               .Sql += 	"MST_単価売上"
                .Sql += " Where 品番 = '" & RTrim(JTM.品番) & "'"
                .Sql += " And   契約先 = '" & RTrim(JTM.契約先) & "'"
                .Sql += " And   履歴NO = " & JTM.履歴NO & ""
                .Sql += " And   号補区分 = '" & RTrim(JTM.号補区分) & "'"
                .Sql += " And   納入番号 = '" & RTrim(JTM.納入番号) & "'"
                .Sql += " And   納入先 = '" & RTrim(JTM.納入先) & "'"
            End If

            If ReadFirstGyo Then
                JTM_RDSTS = .DBRead(True, ConnectNo)
            Else
                JTM_RDSTS = .DBRead(ConnectNo)
            End If
            If JTM_RDSTS Then
                Call ItemsFromJTM()
            Else
                If Init Then
                   Call ItemsClearJTM	
                End If
            End If
            Return JTM_RDSTS
        End With
    End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)   *
'*----------------------------------------------*'
    Public Function JTMReadNext(Optional ByVal ConnectNo As Byte = 1,Optional ByVal Init As Boolean = False) As Boolean
        If CisDB.DBReadNext(ConnectNo) Then
            Call ItemsFromJTM()
            JTM_RDSTS = True
        Else
            JTM_RDSTS = False
            If Init Then
               Call ItemsClearJTM	
            End If
        End If
        Return JTM_RDSTS
    End Function
'*----------------------------------------------*'
'*      ワークエリアクリア  
'*----------------------------------------------*'
Public Sub ItemsClearJTM()
    With JTM
        .品番 = ""
        .契約先 = ""
        .履歴NO = 0
        .号補区分 = ""
        .納入番号 = ""
        .適用開始日 = ""
        .適用終了日 = ""
        .旧型区分 = 0
        .仮区分 = ""
        .単価 = 0
        .単価変更理由 = ""
        .梱包区分 = 0
        .梱包単価 = 0
        .備考 = ""
        .設変符号 = ""
        .納入先 = ""
        .作成日 = Nothing
        .作成者 = ""
        .作成端末 = ""
        .更新日 = Nothing
        .更新者 = ""
        .更新端末 = ""
        .件数 = 0
'   項目追加(自動)
	.内訳単価1 = 0
	.内訳単価2 = 0
	.内訳単価3 = 0
	.内訳単価4 = 0
	.内訳単価5 = 0
	.内訳単価6 = 0
	.内訳単価7 = 0
	.内訳単価8 = 0
	.内訳単価9 = 0
	.内訳単価10 = 0
	.納入先名 = ""
	.今期単価 = 0
	.来期単価 = 0
	.取引先名 = ""
	.今期区分 = ""
	.来期区分 = ""
	.契約先名 = ""
	.現旧区分名 = ""
	.単価変更理由名 = ""
	.号補区分名 = ""
	.表示品番 = ""
	.Warp単価 = 0
	.Warp比較 = 0
	.改定単価 = 0
	.内訳単価計 = 0
	.品名 = ""
	.現在内訳単価1 = 0
	.現在内訳単価2 = 0
	.現在内訳単価3 = 0
	.現在内訳単価4 = 0
	.現在内訳単価5 = 0
	.現在内訳単価6 = 0
	.現在内訳単価7 = 0
	.現在内訳単価8 = 0
	.現在内訳単価9 = 0
	.現在内訳単価10 = 0
	.現在仮区分 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理      *
'*--------------------*'
Public Function JTMInsert(Optional ByVal ConnectNo As Byte = 1) As Boolean
    If JTM.作成者 = Nothing orElse JTM.作成者.Trim = "" Then 
       JTM.作成者 = User_Tanto 
    End If
    With CisDB
         .Sql = "Insert Into "
         If JTM_DBNAME <> "" Then
            .Sql += JTM_DBNAME & ".."
         End If
         .Sql += "MST_単価売上  ( "
         .Sql += "品番,契約先,履歴NO,"
         .Sql += "号補区分,納入番号,適用開始日,"
         .Sql += "適用終了日,旧型区分,仮区分,"
         .Sql += "単価,単価変更理由,梱包区分,"
         .Sql += "梱包単価,備考,設変符号,"
         .Sql += "納入先,"
         .Sql += "作成日,作成者,作成端末"
         .Sql += " ) Values ( "
         .Sql += "'" & RTrim(JTM.品番) & "',"		
         .Sql += "'" & RTrim(JTM.契約先) & "',"		
         .Sql += "" & JTM.履歴NO & ","		
         .Sql += "'" & RTrim(JTM.号補区分) & "',"		
         .Sql += "'" & RTrim(JTM.納入番号) & "',"		
         .Sql += "'" & RTrim(JTM.適用開始日) & "',"		
         .Sql += "'" & RTrim(JTM.適用終了日) & "',"		
         .Sql += "" & JTM.旧型区分 & ","		
         .Sql += "'" & RTrim(JTM.仮区分) & "',"		
         .Sql += "" & JTM.単価 & ","		
         .Sql += "'" & RTrim(JTM.単価変更理由) & "',"		
         .Sql += "" & JTM.梱包区分 & ","		
         .Sql += "" & JTM.梱包単価 & ","		
         .Sql += "'" & RTrim(JTM.備考) & "',"		
         .Sql += "'" & RTrim(JTM.設変符号) & "',"		
         .Sql += "'" & RTrim(JTM.納入先) & "',"		
         .Sql += "GETDATE(),"		
         .Sql += "'" & RTrim(JTM.作成者) & "',"		
         .Sql += "SubString(Host_Name(),1,20)"		
         .Sql +=  " )"

         Return .DBExec(ConnectNo)
    End With
End Function
'*--------------------*'
'*      更新処理      *
'*--------------------*'
Public Function JTMUpdate(Optional ByVal ConnectNo As Byte = 1) As Boolean
    With CisDB
         .Sql = "Update "
         If JTM_DBNAME <> "" Then
            .Sql += JTM_DBNAME & ".."
         End If
	 .Sql += "MST_単価売上 Set "
         .Sql += "品番 = '" & RTrim(JTM.品番) & "',"
         .Sql += "契約先 = '" & RTrim(JTM.契約先) & "',"
         .Sql += "履歴NO = " & JTM.履歴NO & ","
         .Sql += "号補区分 = '" & RTrim(JTM.号補区分) & "',"
         .Sql += "納入番号 = '" & RTrim(JTM.納入番号) & "',"
         .Sql += "適用開始日 = '" & RTrim(JTM.適用開始日) & "',"
         .Sql += "適用終了日 = '" & RTrim(JTM.適用終了日) & "',"
         .Sql += "旧型区分 = " & JTM.旧型区分 & ","
         .Sql += "仮区分 = '" & RTrim(JTM.仮区分) & "',"
         .Sql += "単価 = " & JTM.単価 & ","
         .Sql += "単価変更理由 = '" & RTrim(JTM.単価変更理由) & "',"
         .Sql += "梱包区分 = " & JTM.梱包区分 & ","
         .Sql += "梱包単価 = " & JTM.梱包単価 & ","
         .Sql += "備考 = '" & RTrim(JTM.備考) & "',"
         .Sql += "設変符号 = '" & RTrim(JTM.設変符号) & "',"
         .Sql += "納入先 = '" & RTrim(JTM.納入先) & "',"
         .Sql += "更新日 = GETDATE(),"
         .Sql += "更新者 = '" & RTrim(User_Tanto) & "',"
         .Sql += "更新端末 = SubString(Host_Name(),1,20)"
         .Sql += " Where 品番 = '" & RTrim(JTM.品番) & "'"
         .Sql += "   And 契約先 = '" & RTrim(JTM.契約先) & "'"
         .Sql += "   And 履歴NO = " & JTM.履歴NO & ""
         .Sql += "   And 号補区分 = '" & RTrim(JTM.号補区分) & "'"
         .Sql += "   And 納入番号 = '" & RTrim(JTM.納入番号) & "'"
         .Sql += "   And 納入先 = '" & RTrim(JTM.納入先) & "'"

         Return .DBExec(ConnectNo)
    End With
End Function
End Module

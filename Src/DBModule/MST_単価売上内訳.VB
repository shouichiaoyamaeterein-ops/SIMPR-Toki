Module MST_単価売上内訳
'*--------------------------------------------------------------------------------
'*
'*   .Net  仕様
'*
'*   MST_単価売上内訳
'*
'*   モジュールID  : MST_単価売上内訳.vb
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*                    ※ 追加時は"ﾘﾝｸとして追加"を行って下さい。	
'*					
'*
'*   作成日  :  2019/09/10  By Cis(CDH00301)
'*   変更日  :
'*---------------------------------------------------------------------------------
Structure  T_MST_単価売上内訳
    Dim 品番  As String
    Dim 契約先  As String
    Dim 履歴NO  As Integer
    Dim 号補区分  As String
    Dim 納入番号  As String
    Dim 適用開始日  As String
    Dim 内訳単価1  As Decimal
    Dim 内訳単価2  As Decimal
    Dim 内訳単価3  As Decimal
    Dim 内訳単価4  As Decimal
    Dim 内訳単価5  As Decimal
    Dim 内訳単価6  As Decimal
    Dim 内訳単価7  As Decimal
    Dim 内訳単価8  As Decimal
    Dim 内訳単価9  As Decimal
    Dim 内訳単価10  As Decimal
    Dim 納入先  As String
    Dim 作成日  As Date
    Dim 作成者  As String
    Dim 作成端末  As String
    Dim 更新日  As Date
    Dim 更新者  As String
    Dim 更新端末  As String
        Dim 件数 As Long
End Structure

Public JTMU As T_MST_単価売上内訳

Public JTMU_RDSTS    As Boolean     	    ' 専用読取フラグ
Public JTMU_DBNAME   As String = ""           ' 指定DBNAME
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ        *
'*----------------------------------------*'
Public Sub ItemsFromJTMU()
'   退避エリアクリア
    Call ItemsClearJTMU	
    With CisDB
         For .ItemCnt = 0 To .SqlDataRow.Table.Columns.Count - 1 	
             Select Case .SqlDataRow.Table.Columns(.ItemCnt).Caption
                    Case "品番":              If Not IsDBNull(.SqlDataRow("品番")) Then JTMU.品番 = CStr(.SqlDataRow("品番"))
                    Case "契約先":            If Not IsDBNull(.SqlDataRow("契約先")) Then JTMU.契約先 = CStr(.SqlDataRow("契約先"))
                    Case "履歴NO":            If Not IsDBNull(.SqlDataRow("履歴NO")) Then JTMU.履歴NO = CInt(.SqlDataRow("履歴NO"))
                    Case "号補区分":          If Not IsDBNull(.SqlDataRow("号補区分")) Then JTMU.号補区分 = CStr(.SqlDataRow("号補区分"))
                    Case "納入番号":          If Not IsDBNull(.SqlDataRow("納入番号")) Then JTMU.納入番号 = CStr(.SqlDataRow("納入番号"))
                    Case "適用開始日":        If Not IsDBNull(.SqlDataRow("適用開始日")) Then JTMU.適用開始日 = CStr(.SqlDataRow("適用開始日"))
                    Case "内訳単価1":         If Not IsDBNull(.SqlDataRow("内訳単価1")) Then JTMU.内訳単価1 = CDec(.SqlDataRow("内訳単価1"))
                    Case "内訳単価2":         If Not IsDBNull(.SqlDataRow("内訳単価2")) Then JTMU.内訳単価2 = CDec(.SqlDataRow("内訳単価2"))
                    Case "内訳単価3":         If Not IsDBNull(.SqlDataRow("内訳単価3")) Then JTMU.内訳単価3 = CDec(.SqlDataRow("内訳単価3"))
                    Case "内訳単価4":         If Not IsDBNull(.SqlDataRow("内訳単価4")) Then JTMU.内訳単価4 = CDec(.SqlDataRow("内訳単価4"))
                    Case "内訳単価5":         If Not IsDBNull(.SqlDataRow("内訳単価5")) Then JTMU.内訳単価5 = CDec(.SqlDataRow("内訳単価5"))
                    Case "内訳単価6":         If Not IsDBNull(.SqlDataRow("内訳単価6")) Then JTMU.内訳単価6 = CDec(.SqlDataRow("内訳単価6"))
                    Case "内訳単価7":         If Not IsDBNull(.SqlDataRow("内訳単価7")) Then JTMU.内訳単価7 = CDec(.SqlDataRow("内訳単価7"))
                    Case "内訳単価8":         If Not IsDBNull(.SqlDataRow("内訳単価8")) Then JTMU.内訳単価8 = CDec(.SqlDataRow("内訳単価8"))
                    Case "内訳単価9":         If Not IsDBNull(.SqlDataRow("内訳単価9")) Then JTMU.内訳単価9 = CDec(.SqlDataRow("内訳単価9"))
                    Case "内訳単価10":        If Not IsDBNull(.SqlDataRow("内訳単価10")) Then JTMU.内訳単価10 = CDec(.SqlDataRow("内訳単価10"))
                    Case "納入先":            If Not IsDBNull(.SqlDataRow("納入先")) Then JTMU.納入先 = CStr(.SqlDataRow("納入先"))
                    Case "作成日":            If Not IsDBNull(.SqlDataRow("作成日")) Then JTMU.作成日 = CDate(.SqlDataRow("作成日"))
                    Case "作成者":            If Not IsDBNull(.SqlDataRow("作成者")) Then JTMU.作成者 = CStr(.SqlDataRow("作成者"))
                    Case "作成端末":          If Not IsDBNull(.SqlDataRow("作成端末")) Then JTMU.作成端末 = CStr(.SqlDataRow("作成端末"))
                    Case "更新日":            If Not IsDBNull(.SqlDataRow("更新日")) Then JTMU.更新日 = CDate(.SqlDataRow("更新日"))
                    Case "更新者":            If Not IsDBNull(.SqlDataRow("更新者")) Then JTMU.更新者 = CStr(.SqlDataRow("更新者"))
                    Case "更新端末":          If Not IsDBNull(.SqlDataRow("更新端末")) Then JTMU.更新端末 = CStr(.SqlDataRow("更新端末"))
                Case "件数":         If Not IsDBNull(.SqlDataRow("件数")) Then JTMU.件数 = CLng(.SqlDataRow("件数"))
             End Select
         Next .ItemCnt
    End With	
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ   *
'*----------------------------------------*'
    Public Function JTMURead(Optional ByVal ReadgSql As String = "", Optional ByVal ReadFirstGyo As Boolean = False, Optional ByVal ConnectNo As Byte = 1,Optional ByVal Init As Boolean = False) As Boolean
        With CisDB
            If ReadgSql = "" And .Sql.Trim = "" Then
               .Sql = "Select * From " 
               If JTMU_DBNAME <> "" Then
                  .Sql += JTMU_DBNAME & ".."
               End If
               .Sql += 	"MST_単価売上内訳"
                .Sql += " Where 品番 = '" & RTrim(JTMU.品番) & "'"
                .Sql += " And   契約先 = '" & RTrim(JTMU.契約先) & "'"
                .Sql += " And   履歴NO = " & JTMU.履歴NO & ""
                .Sql += " And   号補区分 = '" & RTrim(JTMU.号補区分) & "'"
                .Sql += " And   納入番号 = '" & RTrim(JTMU.納入番号) & "'"
                .Sql += " And   納入先 = '" & RTrim(JTMU.納入先) & "'"
            End If

            If ReadFirstGyo Then
                JTMU_RDSTS = .DBRead(True, ConnectNo)
            Else
                JTMU_RDSTS = .DBRead(ConnectNo)
            End If
            If JTMU_RDSTS Then
                Call ItemsFromJTMU()
            Else
                If Init Then
                   Call ItemsClearJTMU	
                End If
            End If
            Return JTMU_RDSTS
        End With
    End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)   *
'*----------------------------------------------*'
    Public Function JTMUReadNext(Optional ByVal ConnectNo As Byte = 1,Optional ByVal Init As Boolean = False) As Boolean
        If CisDB.DBReadNext(ConnectNo) Then
            Call ItemsFromJTMU()
            JTMU_RDSTS = True
        Else
            JTMU_RDSTS = False
            If Init Then
               Call ItemsClearJTMU	
            End If
        End If
        Return JTMU_RDSTS
    End Function
'*----------------------------------------------*'
'*      ワークエリアクリア  
'*----------------------------------------------*'
Public Sub ItemsClearJTMU()
    With JTMU
        .品番 = ""
        .契約先 = ""
        .履歴NO = 0
        .号補区分 = ""
        .納入番号 = ""
        .適用開始日 = ""
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
        .納入先 = ""
        .作成日 = Nothing
        .作成者 = ""
        .作成端末 = ""
        .更新日 = Nothing
        .更新者 = ""
        .更新端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理      *
'*--------------------*'
Public Function JTMUInsert(Optional ByVal ConnectNo As Byte = 1) As Boolean
    If JTMU.作成者 = Nothing orElse JTMU.作成者.Trim = "" Then 
       JTMU.作成者 = User_Tanto 
    End If
    With CisDB
         .Sql = "Insert Into "
         If JTMU_DBNAME <> "" Then
            .Sql += JTMU_DBNAME & ".."
         End If
         .Sql += "MST_単価売上内訳  ( "
         .Sql += "品番,契約先,履歴NO,"
         .Sql += "号補区分,納入番号,適用開始日,"
         .Sql += "内訳単価1,内訳単価2,内訳単価3,"
         .Sql += "内訳単価4,内訳単価5,内訳単価6,"
         .Sql += "内訳単価7,内訳単価8,内訳単価9,"
         .Sql += "内訳単価10,納入先,"
         .Sql += "作成日,作成者,作成端末"
         .Sql += " ) Values ( "
         .Sql += "'" & RTrim(JTMU.品番) & "',"		
         .Sql += "'" & RTrim(JTMU.契約先) & "',"		
         .Sql += "" & JTMU.履歴NO & ","		
         .Sql += "'" & RTrim(JTMU.号補区分) & "',"		
         .Sql += "'" & RTrim(JTMU.納入番号) & "',"		
         .Sql += "'" & RTrim(JTMU.適用開始日) & "',"		
         .Sql += "" & JTMU.内訳単価1 & ","		
         .Sql += "" & JTMU.内訳単価2 & ","		
         .Sql += "" & JTMU.内訳単価3 & ","		
         .Sql += "" & JTMU.内訳単価4 & ","		
         .Sql += "" & JTMU.内訳単価5 & ","		
         .Sql += "" & JTMU.内訳単価6 & ","		
         .Sql += "" & JTMU.内訳単価7 & ","		
         .Sql += "" & JTMU.内訳単価8 & ","		
         .Sql += "" & JTMU.内訳単価9 & ","		
         .Sql += "" & JTMU.内訳単価10 & ","		
         .Sql += "'" & RTrim(JTMU.納入先) & "',"		
         .Sql += "GETDATE(),"		
         .Sql += "'" & RTrim(JTMU.作成者) & "',"		
         .Sql += "SubString(Host_Name(),1,20)"		
         .Sql +=  " )"

         Return .DBExec(ConnectNo)
    End With
End Function
'*--------------------*'
'*      更新処理      *
'*--------------------*'
Public Function JTMUUpdate(Optional ByVal ConnectNo As Byte = 1) As Boolean
    With CisDB
         .Sql = "Update "
         If JTMU_DBNAME <> "" Then
            .Sql += JTMU_DBNAME & ".."
         End If
	 .Sql += "MST_単価売上内訳 Set "
         .Sql += "品番 = '" & RTrim(JTMU.品番) & "',"
         .Sql += "契約先 = '" & RTrim(JTMU.契約先) & "',"
         .Sql += "履歴NO = " & JTMU.履歴NO & ","
         .Sql += "号補区分 = '" & RTrim(JTMU.号補区分) & "',"
         .Sql += "納入番号 = '" & RTrim(JTMU.納入番号) & "',"
         .Sql += "適用開始日 = '" & RTrim(JTMU.適用開始日) & "',"
         .Sql += "内訳単価1 = " & JTMU.内訳単価1 & ","
         .Sql += "内訳単価2 = " & JTMU.内訳単価2 & ","
         .Sql += "内訳単価3 = " & JTMU.内訳単価3 & ","
         .Sql += "内訳単価4 = " & JTMU.内訳単価4 & ","
         .Sql += "内訳単価5 = " & JTMU.内訳単価5 & ","
         .Sql += "内訳単価6 = " & JTMU.内訳単価6 & ","
         .Sql += "内訳単価7 = " & JTMU.内訳単価7 & ","
         .Sql += "内訳単価8 = " & JTMU.内訳単価8 & ","
         .Sql += "内訳単価9 = " & JTMU.内訳単価9 & ","
         .Sql += "内訳単価10 = " & JTMU.内訳単価10 & ","
         .Sql += "納入先 = '" & RTrim(JTMU.納入先) & "',"
         .Sql += "更新日 = GETDATE(),"
         .Sql += "更新者 = '" & RTrim(User_Tanto) & "',"
         .Sql += "更新端末 = SubString(Host_Name(),1,20)"
         .Sql += " Where 品番 = '" & RTrim(JTMU.品番) & "'"
         .Sql += "   And 契約先 = '" & RTrim(JTMU.契約先) & "'"
         .Sql += "   And 履歴NO = " & JTMU.履歴NO & ""
         .Sql += "   And 号補区分 = '" & RTrim(JTMU.号補区分) & "'"
         .Sql += "   And 納入番号 = '" & RTrim(JTMU.納入番号) & "'"
         .Sql += "   And 納入先 = '" & RTrim(JTMU.納入先) & "'"

         Return .DBExec(ConnectNo)
    End With
End Function
End Module

Attribute VB_Name = "内示テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   内示テーブル
'*
'*   モジュールID  : 内示テーブル.Bas
'*
'*   作成日  :  2003/12/15  By Cis 内示数　配列化
'*   変更日  :  2004/03/05  By Cis NJTInsert & NJTUpdate を Function に変更
'*   変更日  :  2008/05/02  By CIS 内示金額　配列化
'*   変更日  :  2009/10/15  By CIS 表示品番/納入先名/契約先名/比率 追加
'*---------------------------------------------------------------------------------
    Type T_内示テーブル
        内示年月          As String
        品番              As String
        納入先            As String
        受入              As String
        内示区分          As String
        契約先            As String
        工区              As String
        層別区分          As String
        号補区分          As String
        検索品番          As String
        適用開始年月日    As String
        適用終了年月日    As String
        内示数            As Long
        翌月内示数        As Long
        翌々月内示数      As Long
        管理NO            As String
        収容数            As Long
        背番号            As String
        サイクル          As String
        担当              As String
        整備              As String
        納入区分          As String
        納品書発行枚数    As Integer
        かんばん回転枚数  As Integer
        かんばん増減符号  As String
        かんばん増減枚数  As Integer
        変更数符号        As String
        変更数            As Long
        調整数符号        As String
        調整数            As Long
        処理区分          As Byte
        翌月処理区分      As Byte
        削除区分          As Byte
        変更回数          As Byte
        登録区分          As Byte
        登録FCODE         As String
        登録H作成日       As String
        登録HSEQ          As String
        登録連番          As Long
        内示数量(31)      As Long
        内示金額(31)      As Currency
        コメント          As String
        作成者            As String
'        作成日            As Date
'        作成端末          As String
        更新者            As String
'        更新日            As Date
'        更新端末          As String

        件数          As Long
'---------------------------------------' 2009.10.15 add start
        表示品番        As String
        納入先名        As String
        契約先名        As String
        内示比率        As Double
'---------------------------------------' 2009.10.15 add end
    End Type

    Public NJT As T_内示テーブル

    Public NJT_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
    Dim idx             As Long
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromNJT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearNJT
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With NJT
            Select Case adoCol.Name
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "納入先":            If Not IsNull(adoCol.Value) Then .納入先 = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "内示区分":          If Not IsNull(adoCol.Value) Then .内示区分 = adoCol.Value
                Case "内示年月":          If Not IsNull(adoCol.Value) Then .内示年月 = adoCol.Value
                Case "契約先":            If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
                Case "工区":              If Not IsNull(adoCol.Value) Then .工区 = adoCol.Value
                Case "層別区分":          If Not IsNull(adoCol.Value) Then .層別区分 = adoCol.Value
                Case "号補区分":          If Not IsNull(adoCol.Value) Then .号補区分 = adoCol.Value
                Case "検索品番":          If Not IsNull(adoCol.Value) Then .検索品番 = adoCol.Value
                Case "適用開始年月日":    If Not IsNull(adoCol.Value) Then .適用開始年月日 = adoCol.Value
                Case "適用終了年月日":    If Not IsNull(adoCol.Value) Then .適用終了年月日 = adoCol.Value
                Case "内示数":            If Not IsNull(adoCol.Value) Then .内示数 = adoCol.Value
                Case "翌月内示数":        If Not IsNull(adoCol.Value) Then .翌月内示数 = adoCol.Value
                Case "翌々月内示数":      If Not IsNull(adoCol.Value) Then .翌々月内示数 = adoCol.Value
                Case "管理NO":            If Not IsNull(adoCol.Value) Then .管理NO = adoCol.Value
                Case "収容数":            If Not IsNull(adoCol.Value) Then .収容数 = adoCol.Value
                Case "背番号":            If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "サイクル":          If Not IsNull(adoCol.Value) Then .サイクル = adoCol.Value
                Case "担当":              If Not IsNull(adoCol.Value) Then .担当 = adoCol.Value
                Case "整備":              If Not IsNull(adoCol.Value) Then .整備 = adoCol.Value
                Case "納入区分":          If Not IsNull(adoCol.Value) Then .納入区分 = adoCol.Value
                Case "納品書発行枚数":    If Not IsNull(adoCol.Value) Then .納品書発行枚数 = adoCol.Value
                Case "かんばん回転枚数":  If Not IsNull(adoCol.Value) Then .かんばん回転枚数 = adoCol.Value
                Case "かんばん増減符号":  If Not IsNull(adoCol.Value) Then .かんばん増減符号 = adoCol.Value
                Case "かんばん増減枚数":  If Not IsNull(adoCol.Value) Then .かんばん増減枚数 = adoCol.Value
                Case "変更数符号":        If Not IsNull(adoCol.Value) Then .変更数符号 = adoCol.Value
                Case "変更数":            If Not IsNull(adoCol.Value) Then .変更数 = adoCol.Value
                Case "調整数符号":        If Not IsNull(adoCol.Value) Then .調整数符号 = adoCol.Value
                Case "調整数":            If Not IsNull(adoCol.Value) Then .調整数 = adoCol.Value
                Case "処理区分":          If Not IsNull(adoCol.Value) Then .処理区分 = adoCol.Value
                Case "翌月処理区分":      If Not IsNull(adoCol.Value) Then .翌月処理区分 = adoCol.Value
                Case "削除区分":          If Not IsNull(adoCol.Value) Then .削除区分 = adoCol.Value
                Case "変更回数":          If Not IsNull(adoCol.Value) Then .変更回数 = adoCol.Value
                Case "登録区分":          If Not IsNull(adoCol.Value) Then .登録区分 = adoCol.Value
                Case "登録FCODE":         If Not IsNull(adoCol.Value) Then .登録FCODE = adoCol.Value
                Case "登録H作成日":       If Not IsNull(adoCol.Value) Then .登録H作成日 = adoCol.Value
                Case "登録HSEQ":          If Not IsNull(adoCol.Value) Then .登録HSEQ = adoCol.Value
                Case "登録連番":          If Not IsNull(adoCol.Value) Then .登録連番 = adoCol.Value
                Case "コメント":          If Not IsNull(adoCol.Value) Then .コメント = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
                Case "更新者":            If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'-----------------------------------------------------------------------------------' 2009.10.15 add start
                Case "表示品番":        If Not IsNull(adoCol.Value) Then .表示品番 = adoCol.Value
                Case "納入先名":        If Not IsNull(adoCol.Value) Then .納入先名 = adoCol.Value
                Case "契約先名":        If Not IsNull(adoCol.Value) Then .契約先名 = adoCol.Value
                Case "内示比率":        If Not IsNull(adoCol.Value) Then .内示比率 = adoCol.Value
'-----------------------------------------------------------------------------------' 2009.10.15 add end
                Case Else
                    For idx = 1 To 31
                        If "内示数量" & Format(idx, "00") = adoCol.Name Then
                            .内示数量(idx) = adoCol.Value
                            Exit For
                        End If
                        If "内示金額" & Format(idx, "00") = adoCol.Name Then
                            .内示金額(idx) = adoCol.Value
                            Exit For
                        End If
                    Next idx
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function NJTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    NJTRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 内示テーブル"
        RTxtSql = RTxtSql & " Where 内示年月 = '" & RTrim(NJT.内示年月) & "'"
        RTxtSql = RTxtSql & " And   品番 = '" & RTrim(NJT.品番) & "'"
        RTxtSql = RTxtSql & " And   納入先 = '" & RTrim(NJT.納入先) & "'"
        RTxtSql = RTxtSql & " And   受入 = '" & RTrim(NJT.受入) & "'"
        RTxtSql = RTxtSql & " And   内示区分 = '" & RTrim(NJT.内示区分) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           NJTRead = True
           Call ItemsFromNJT(DBNo)
           NJT_RDSTS = True
        Else
           NJT_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call NJTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub NJTReadNext(Optional DBNo As Integer = 0)
    If Not NJT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           NJT_RDSTS = False
        Else
           Call ItemsFromNJT(DBNo)
           NJT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub NJTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        NJT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( NJT )
'*----------------------------------------------*'
Public Sub ItemsClearNJT()
    With NJT
        .品番 = ""
        .納入先 = ""
        .受入 = ""
        .内示区分 = ""
        .内示年月 = ""
        .契約先 = ""
        .工区 = ""
        .層別区分 = ""
        .号補区分 = ""
        .検索品番 = ""
        .適用開始年月日 = ""
        .適用終了年月日 = ""
        .内示数 = 0
        .翌月内示数 = 0
        .翌々月内示数 = 0
        .管理NO = ""
        .収容数 = 0
        .背番号 = ""
        .サイクル = ""
        .担当 = ""
        .整備 = ""
        .納入区分 = ""
        .納品書発行枚数 = 0
        .かんばん回転枚数 = 0
        .かんばん増減符号 = ""
        .かんばん増減枚数 = 0
        .変更数符号 = ""
        .変更数 = 0
        .調整数符号 = ""
        .調整数 = 0
        .処理区分 = 0
        .翌月処理区分 = 0
        .削除区分 = 0
        .変更回数 = 0
        .登録区分 = 0
        .登録FCODE = ""
        .登録H作成日 = ""
        .登録HSEQ = ""
        .登録連番 = 0
        For idx = 1 To 31
            .内示数量(idx) = 0
            .内示金額(idx) = 0
        Next idx
        .コメント = ""
        .作成者 = ""
'       .作成日 = 0
'       .作成端末 = ""
        .更新者 = ""
'       .更新日 = 0
'       .更新端末 = ""
        .件数 = 0
'---------------------------------------' 2009.10.15 add start
        .表示品番 = ""
        .納入先名 = ""
        .契約先名 = ""
        .内示比率 = 0
'---------------------------------------' 2009.10.15 add end
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function NJTInsert(Optional DBNo As Integer) As Boolean
    NJTInsert = False
    With NJT
        gSL_Select = "Insert Into 内示テーブル  ( "
        gSL_Select = gSL_Select & "内示年月,品番,納入先,"
        gSL_Select = gSL_Select & "受入,内示区分,"
        gSL_Select = gSL_Select & "契約先,工区,層別区分,"
        gSL_Select = gSL_Select & "号補区分,検索品番,適用開始年月日,"
        gSL_Select = gSL_Select & "適用終了年月日,内示数,翌月内示数,"
        gSL_Select = gSL_Select & "翌々月内示数,管理NO,収容数,"
        gSL_Select = gSL_Select & "背番号,サイクル,担当,"
        gSL_Select = gSL_Select & "整備,納入区分,納品書発行枚数,"
        gSL_Select = gSL_Select & "かんばん回転枚数,かんばん増減符号,かんばん増減枚数,"
        gSL_Select = gSL_Select & "変更数符号,変更数,調整数符号,"
        gSL_Select = gSL_Select & "調整数,処理区分,翌月処理区分,"
        gSL_Select = gSL_Select & "削除区分,変更回数,登録区分,"
        gSL_Select = gSL_Select & "登録FCODE,登録H作成日,登録HSEQ,"
        gSL_Select = gSL_Select & "登録連番,"
        For idx = 1 To 31
            gSL_Select = gSL_Select & "内示数量" & Format(idx, "00") & ","
        Next idx
        gSL_Select = gSL_Select & "コメント,"
        gSL_Select = gSL_Select & "作成者,作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.内示年月) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.内示区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.工区) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.層別区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.号補区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.検索品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.適用開始年月日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.適用終了年月日) & "',"
        gSL_Select = gSL_Select & "" & .内示数 & ","
        gSL_Select = gSL_Select & "" & .翌月内示数 & ","
        gSL_Select = gSL_Select & "" & .翌々月内示数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.管理NO) & "',"
        gSL_Select = gSL_Select & "" & .収容数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.担当) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.整備) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入区分) & "',"
        gSL_Select = gSL_Select & "" & .納品書発行枚数 & ","
        gSL_Select = gSL_Select & "" & .かんばん回転枚数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.かんばん増減符号) & "',"
        gSL_Select = gSL_Select & "" & .かんばん増減枚数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.変更数符号) & "',"
        gSL_Select = gSL_Select & "" & .変更数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.調整数符号) & "',"
        gSL_Select = gSL_Select & "" & .調整数 & ","
        gSL_Select = gSL_Select & "" & .処理区分 & ","
        gSL_Select = gSL_Select & "" & .翌月処理区分 & ","
        gSL_Select = gSL_Select & "" & .削除区分 & ","
        gSL_Select = gSL_Select & "" & .変更回数 & ","
        gSL_Select = gSL_Select & "" & .登録区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.登録FCODE) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.登録H作成日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.登録HSEQ) & "',"
        gSL_Select = gSL_Select & "" & .登録連番 & ","
        For idx = 1 To 31
            gSL_Select = gSL_Select & "" & .内示数量(idx) & ","
        Next idx
        gSL_Select = gSL_Select & "'" & RTrim(.コメント) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "',"
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    NJTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function NJTUpdate(Optional DBNo As Integer) As Boolean
    NJTUpdate = False
    With NJT
        gSL_Select = "Update 内示テーブル  Set "
        gSL_Select = gSL_Select & "内示年月 = '" & RTrim(.内示年月) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "内示区分 = '" & RTrim(.内示区分) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "工区 = '" & RTrim(.工区) & "',"
        gSL_Select = gSL_Select & "層別区分 = '" & RTrim(.層別区分) & "',"
        gSL_Select = gSL_Select & "号補区分 = '" & RTrim(.号補区分) & "',"
        gSL_Select = gSL_Select & "検索品番 = '" & RTrim(.検索品番) & "',"
        gSL_Select = gSL_Select & "適用開始年月日 = '" & RTrim(.適用開始年月日) & "',"
        gSL_Select = gSL_Select & "適用終了年月日 = '" & RTrim(.適用終了年月日) & "',"
        gSL_Select = gSL_Select & "内示数 = " & .内示数 & ","
        gSL_Select = gSL_Select & "翌月内示数 = " & .翌月内示数 & ","
        gSL_Select = gSL_Select & "翌々月内示数 = " & .翌々月内示数 & ","
        gSL_Select = gSL_Select & "管理NO = '" & RTrim(.管理NO) & "',"
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "サイクル = '" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "担当 = '" & RTrim(.担当) & "',"
        gSL_Select = gSL_Select & "整備 = '" & RTrim(.整備) & "',"
        gSL_Select = gSL_Select & "納入区分 = '" & RTrim(.納入区分) & "',"
        gSL_Select = gSL_Select & "納品書発行枚数 = " & .納品書発行枚数 & ","
        gSL_Select = gSL_Select & "かんばん回転枚数 = " & .かんばん回転枚数 & ","
        gSL_Select = gSL_Select & "かんばん増減符号 = '" & RTrim(.かんばん増減符号) & "',"
        gSL_Select = gSL_Select & "かんばん増減枚数 = " & .かんばん増減枚数 & ","
        gSL_Select = gSL_Select & "変更数符号 = '" & RTrim(.変更数符号) & "',"
        gSL_Select = gSL_Select & "変更数 = " & .変更数 & ","
        gSL_Select = gSL_Select & "調整数符号 = '" & RTrim(.調整数符号) & "',"
        gSL_Select = gSL_Select & "調整数 = " & .調整数 & ","
        gSL_Select = gSL_Select & "処理区分 = " & .処理区分 & ","
        gSL_Select = gSL_Select & "翌月処理区分 = " & .翌月処理区分 & ","
        gSL_Select = gSL_Select & "削除区分 = " & .削除区分 & ","
        gSL_Select = gSL_Select & "変更回数 = " & .変更回数 & ","
        gSL_Select = gSL_Select & "登録区分 = " & .登録区分 & ","
        gSL_Select = gSL_Select & "登録FCODE = '" & RTrim(.登録FCODE) & "',"
        gSL_Select = gSL_Select & "登録H作成日 = '" & RTrim(.登録H作成日) & "',"
        gSL_Select = gSL_Select & "登録HSEQ = '" & RTrim(.登録HSEQ) & "',"
        gSL_Select = gSL_Select & "登録連番 = " & .登録連番 & ","
        For idx = 1 To 31
            gSL_Select = gSL_Select & "内示数量" & Format(idx, "00") & " = " & .内示数量(idx) & ","
        Next idx
        gSL_Select = gSL_Select & "コメント = '" & RTrim(.コメント) & "',"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 内示年月 = '" & RTrim(.内示年月) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 納入先 = '" & RTrim(.納入先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 内示区分 = '" & RTrim(.内示区分) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    NJTUpdate = True
End Function

Attribute VB_Name = "内示TEXTテーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   内示TEXTテーブル
'*
'*   モジュールID  : 内示TEXTテーブル.Bas
'*
'*   作成日  :  2003/12/11  By Cis  内示数　配列化
'*   変更日  :  2003/12/17  By Cis  項目追加(契約先･正式名称･内示形態・コメント･品番編集)
'*   変更日  :  2004/11/12  By Cis  項目追加(品名･契約先)
'*   変更日  :  2005/06/07  By Cis  項目追加(車型)
'*---------------------------------------------------------------------------------
    Type T_内示TEXTテーブル
        FILECODE          As String
        HEAD作成日        As String
        HEADSEQ           As String
        連番              As Long
        対象年月          As String
        内示年月          As String
        品番              As String
        納入先            As String
        受入              As String
        層別区分          As String
        号補区分          As String
        適用開始年月日    As String
        適用終了年月日    As String
        内示数            As Long
        翌月内示数        As Long
        翌々月内示数      As Long
        管理NO            As String
        収容数            As Long
        背番号            As String
        サイクル          As String
        状態区分          As String
        納品書不作成区分  As String
        仕入先工区        As String
        置換工区          As String
        持込先            As String
        持込先工場        As String
        出庫部署          As String
        担当              As String
        支給係            As String
        内外区分          As String
        現旧区分          As String
        整備              As String
        担当者            As String
        支給先受入        As String
        支給先整備        As String
        支給先担当者      As String
        管理区分          As String
        生産区分          As String
        納品書発行枚数    As Integer
        かんばん回転枚数  As Integer
        かんばん増減符号  As String
        かんばん増減枚数  As Integer
        当初内示数        As Long
        変更数符号        As String
        変更数            As Long
        調整数符号        As String
        調整数            As Long
        納入区分          As String
        AB区分            As String
        変更回数           As Long
        内示数量(31)       As Long
'        作成日            As Date
'        作成端末          As String
        品名              As String
        件数            As Long
        
        契約先          As String
        正式名称        As String
        内示形態        As Integer
        コメント        As String
        品番編集        As String
        車型            As String
    End Type

    Public NTT As T_内示TEXTテーブル

    Public NTT_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
    Dim idx             As Long
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromNTT(Optional DBNo As Integer)

'   退避エリアクリア
    Call ItemsClearNTT
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With NTT
            Select Case adoCol.Name
                Case "FILECODE":          If Not IsNull(adoCol.Value) Then .FILECODE = adoCol.Value
                Case "HEAD作成日":        If Not IsNull(adoCol.Value) Then .HEAD作成日 = adoCol.Value
                Case "HEADSEQ":           If Not IsNull(adoCol.Value) Then .HEADSEQ = adoCol.Value
                Case "連番":              If Not IsNull(adoCol.Value) Then .連番 = adoCol.Value
                Case "対象年月":          If Not IsNull(adoCol.Value) Then .対象年月 = adoCol.Value
                Case "内示年月":          If Not IsNull(adoCol.Value) Then .内示年月 = adoCol.Value
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "納入先":            If Not IsNull(adoCol.Value) Then .納入先 = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "層別区分":          If Not IsNull(adoCol.Value) Then .層別区分 = adoCol.Value
                Case "号補区分":          If Not IsNull(adoCol.Value) Then .号補区分 = adoCol.Value
                Case "適用開始年月日":    If Not IsNull(adoCol.Value) Then .適用開始年月日 = adoCol.Value
                Case "適用終了年月日":    If Not IsNull(adoCol.Value) Then .適用終了年月日 = adoCol.Value
                Case "内示数":            If Not IsNull(adoCol.Value) Then .内示数 = adoCol.Value
                Case "翌月内示数":        If Not IsNull(adoCol.Value) Then .翌月内示数 = adoCol.Value
                Case "翌々月内示数":      If Not IsNull(adoCol.Value) Then .翌々月内示数 = adoCol.Value
                Case "管理NO":            If Not IsNull(adoCol.Value) Then .管理NO = adoCol.Value
                Case "収容数":            If Not IsNull(adoCol.Value) Then .収容数 = adoCol.Value
                Case "背番号":            If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "サイクル":          If Not IsNull(adoCol.Value) Then .サイクル = adoCol.Value
                Case "状態区分":          If Not IsNull(adoCol.Value) Then .状態区分 = adoCol.Value
                Case "納品書不作成区分":  If Not IsNull(adoCol.Value) Then .納品書不作成区分 = adoCol.Value
                Case "仕入先工区":        If Not IsNull(adoCol.Value) Then .仕入先工区 = adoCol.Value
                Case "置換工区":          If Not IsNull(adoCol.Value) Then .置換工区 = adoCol.Value
                Case "持込先":            If Not IsNull(adoCol.Value) Then .持込先 = adoCol.Value
                Case "持込先工場":        If Not IsNull(adoCol.Value) Then .持込先工場 = adoCol.Value
                Case "出庫部署":          If Not IsNull(adoCol.Value) Then .出庫部署 = adoCol.Value
                Case "担当":              If Not IsNull(adoCol.Value) Then .担当 = adoCol.Value
                Case "支給係":            If Not IsNull(adoCol.Value) Then .支給係 = adoCol.Value
                Case "内外区分":          If Not IsNull(adoCol.Value) Then .内外区分 = adoCol.Value
                Case "現旧区分":          If Not IsNull(adoCol.Value) Then .現旧区分 = adoCol.Value
                Case "整備":              If Not IsNull(adoCol.Value) Then .整備 = adoCol.Value
                Case "担当者":            If Not IsNull(adoCol.Value) Then .担当者 = adoCol.Value
                Case "支給先受入":        If Not IsNull(adoCol.Value) Then .支給先受入 = adoCol.Value
                Case "支給先整備":        If Not IsNull(adoCol.Value) Then .支給先整備 = adoCol.Value
                Case "支給先担当者":      If Not IsNull(adoCol.Value) Then .支給先担当者 = adoCol.Value
                Case "管理区分":          If Not IsNull(adoCol.Value) Then .管理区分 = adoCol.Value
                Case "生産区分":          If Not IsNull(adoCol.Value) Then .生産区分 = adoCol.Value
                Case "納品書発行枚数":    If Not IsNull(adoCol.Value) Then .納品書発行枚数 = adoCol.Value
                Case "かんばん回転枚数":  If Not IsNull(adoCol.Value) Then .かんばん回転枚数 = adoCol.Value
                Case "かんばん増減符号":  If Not IsNull(adoCol.Value) Then .かんばん増減符号 = adoCol.Value
                Case "かんばん増減枚数":  If Not IsNull(adoCol.Value) Then .かんばん増減枚数 = adoCol.Value
                Case "当初内示数":        If Not IsNull(adoCol.Value) Then .当初内示数 = adoCol.Value
                Case "変更数符号":        If Not IsNull(adoCol.Value) Then .変更数符号 = adoCol.Value
                Case "変更数":            If Not IsNull(adoCol.Value) Then .変更数 = adoCol.Value
                Case "調整数符号":        If Not IsNull(adoCol.Value) Then .調整数符号 = adoCol.Value
                Case "調整数":            If Not IsNull(adoCol.Value) Then .調整数 = adoCol.Value
                Case "納入区分":          If Not IsNull(adoCol.Value) Then .納入区分 = adoCol.Value
                Case "ＡＢ区分":          If Not IsNull(adoCol.Value) Then .AB区分 = adoCol.Value
                Case "変更回数":          If Not IsNull(adoCol.Value) Then .変更回数 = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
                Case "品名":         If Not IsNull(adoCol.Value) Then .品名 = adoCol.Value
                
                Case "契約先":         If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
                Case "正式名称":       If Not IsNull(adoCol.Value) Then .正式名称 = adoCol.Value
                Case "内示形態":       If Not IsNull(adoCol.Value) Then .内示形態 = adoCol.Value
                Case "コメント":       If Not IsNull(adoCol.Value) Then .コメント = adoCol.Value
                Case "品番編集":       If Not IsNull(adoCol.Value) Then .品番編集 = adoCol.Value
                Case "車型":           If Not IsNull(adoCol.Value) Then .車型 = adoCol.Value
                Case Else
                    For idx = 1 To 31
                        If "内示数量" & Format(idx, "00") = adoCol.Name Then
                            .内示数量(idx) = adoCol.Value
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
Public Function NTTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    NTTRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 内示TEXTテーブル"
        RTxtSql = RTxtSql & " Where FILECODE = '" & RTrim(NTT.FILECODE) & "'"
        RTxtSql = RTxtSql & " And   HEAD作成日 = '" & RTrim(NTT.HEAD作成日) & "'"
        RTxtSql = RTxtSql & " And   HEADSEQ = '" & RTrim(NTT.HEADSEQ) & "'"
        RTxtSql = RTxtSql & " And   連番 = " & NTT.連番 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           NTTRead = True
           Call ItemsFromNTT(DBNo)
           NTT_RDSTS = True
        Else
           NTT_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call NTTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub NTTReadNext(Optional DBNo As Integer = 0)
    If Not NTT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           NTT_RDSTS = False
        Else
           Call ItemsFromNTT(DBNo)
           NTT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub NTTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        NTT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( NTT )
'*----------------------------------------------*'
Public Sub ItemsClearNTT()
    With NTT
        .FILECODE = ""
        .HEAD作成日 = ""
        .HEADSEQ = ""
        .連番 = 0
        .対象年月 = ""
        .内示年月 = ""
        .品番 = ""
        .納入先 = ""
        .受入 = ""
        .層別区分 = ""
        .号補区分 = ""
        .適用開始年月日 = ""
        .適用終了年月日 = ""
        .内示数 = 0
        .翌月内示数 = 0
        .翌々月内示数 = 0
        .管理NO = ""
        .収容数 = 0
        .背番号 = ""
        .サイクル = ""
        .状態区分 = ""
        .納品書不作成区分 = ""
        .仕入先工区 = ""
        .置換工区 = ""
        .持込先 = ""
        .持込先工場 = ""
        .出庫部署 = ""
        .担当 = ""
        .支給係 = ""
        .内外区分 = ""
        .現旧区分 = ""
        .整備 = ""
        .担当者 = ""
        .支給先受入 = ""
        .支給先整備 = ""
        .支給先担当者 = ""
        .管理区分 = ""
        .生産区分 = ""
        .納品書発行枚数 = 0
        .かんばん回転枚数 = 0
        .かんばん増減符号 = ""
        .かんばん増減枚数 = 0
        .当初内示数 = 0
        .変更数符号 = ""
        .変更数 = 0
        .調整数符号 = ""
        .調整数 = 0
        .納入区分 = ""
        .AB区分 = ""
        .変更回数 = 0
        For idx = 1 To 31
            .内示数量(idx) = 0
        Next idx
'       .作成日 = 0
'       .作成端末 = ""
        .件数 = 0
        .品名 = ""
        
        .契約先 = ""
        .正式名称 = ""
        .内示形態 = 0
        .コメント = ""
        .品番編集 = ""
        .車型 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Sub NTTInsert(Optional DBNo As Integer)
    With NTT
        gSL_Select = "Insert Into 内示TEXTテーブル  ( "
        gSL_Select = gSL_Select & "FILECODE,HEAD作成日,HEADSEQ,"
        gSL_Select = gSL_Select & "連番,対象年月,内示年月,"
        gSL_Select = gSL_Select & "品番,納入先,受入,"
        gSL_Select = gSL_Select & "層別区分,号補区分,適用開始年月日,"
        gSL_Select = gSL_Select & "適用終了年月日,内示数,翌月内示数,"
        gSL_Select = gSL_Select & "翌々月内示数,管理NO,収容数,"
        gSL_Select = gSL_Select & "背番号,サイクル,状態区分,"
        gSL_Select = gSL_Select & "納品書不作成区分,仕入先工区,置換工区,持込先,"
        gSL_Select = gSL_Select & "持込先工場,出庫部署,担当,"
        gSL_Select = gSL_Select & "支給係,内外区分,現旧区分,"
        gSL_Select = gSL_Select & "整備,担当者,支給先受入,"
        gSL_Select = gSL_Select & "支給先整備,支給先担当者,管理区分,"
        gSL_Select = gSL_Select & "生産区分,納品書発行枚数,かんばん回転枚数,"
        gSL_Select = gSL_Select & "かんばん増減符号,かんばん増減枚数,当初内示数,"
        gSL_Select = gSL_Select & "変更数符号,変更数,調整数符号,"
        gSL_Select = gSL_Select & "調整数,納入区分,ＡＢ区分,変更回数,"
        For idx = 1 To 31
            gSL_Select = gSL_Select & "内示数量" & Format(idx, "00") & ","
        Next idx
        gSL_Select = gSL_Select & "作成日,作成端末,品名,契約先"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.FILECODE) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.HEAD作成日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.HEADSEQ) & "',"
        gSL_Select = gSL_Select & "" & .連番 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.対象年月) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.内示年月) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.層別区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.号補区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.適用開始年月日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.適用終了年月日) & "',"
        gSL_Select = gSL_Select & "" & .内示数 & ","
        gSL_Select = gSL_Select & "" & .翌月内示数 & ","
        gSL_Select = gSL_Select & "" & .翌々月内示数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.管理NO) & "',"
        gSL_Select = gSL_Select & "" & .収容数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.状態区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納品書不作成区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先工区) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.置換工区) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.持込先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.持込先工場) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.出庫部署) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.担当) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.支給係) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.内外区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.現旧区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.整備) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.担当者) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.支給先受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.支給先整備) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.支給先担当者) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.管理区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.生産区分) & "',"
        gSL_Select = gSL_Select & "" & .納品書発行枚数 & ","
        gSL_Select = gSL_Select & "" & .かんばん回転枚数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.かんばん増減符号) & "',"
        gSL_Select = gSL_Select & "" & .かんばん増減枚数 & ","
        gSL_Select = gSL_Select & "" & .当初内示数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.変更数符号) & "',"
        gSL_Select = gSL_Select & "" & .変更数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.調整数符号) & "',"
        gSL_Select = gSL_Select & "" & .調整数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.納入区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.AB区分) & "',"
        gSL_Select = gSL_Select & "" & RTrim(.変更回数) & ","
        For idx = 1 To 31
            gSL_Select = gSL_Select & "" & .内示数量(idx) & ","
        Next idx
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,15),"
        gSL_Select = gSL_Select & "'" & RTrim(.品名) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "'"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        .DBExec
    End With
End Sub
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Sub NTTUpdate(Optional DBNo As Integer)
    With NTT
        gSL_Select = "Update 内示TEXTテーブル  Set "
        gSL_Select = gSL_Select & "FILECODE = '" & RTrim(.FILECODE) & "',"
        gSL_Select = gSL_Select & "HEAD作成日 = '" & RTrim(.HEAD作成日) & "',"
        gSL_Select = gSL_Select & "HEADSEQ = '" & RTrim(.HEADSEQ) & "',"
        gSL_Select = gSL_Select & "連番 = " & .連番 & ","
        gSL_Select = gSL_Select & "対象年月 = '" & RTrim(.対象年月) & "',"
        gSL_Select = gSL_Select & "内示年月 = '" & RTrim(.内示年月) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "層別区分 = '" & RTrim(.層別区分) & "',"
        gSL_Select = gSL_Select & "号補区分 = '" & RTrim(.号補区分) & "',"
        gSL_Select = gSL_Select & "適用開始年月日 = '" & RTrim(.適用開始年月日) & "',"
        gSL_Select = gSL_Select & "適用終了年月日 = '" & RTrim(.適用終了年月日) & "',"
        gSL_Select = gSL_Select & "内示数 = " & .内示数 & ","
        gSL_Select = gSL_Select & "翌月内示数 = " & .翌月内示数 & ","
        gSL_Select = gSL_Select & "翌々月内示数 = " & .翌々月内示数 & ","
        gSL_Select = gSL_Select & "管理NO = '" & RTrim(.管理NO) & "',"
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "サイクル = '" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "状態区分 = '" & RTrim(.状態区分) & "',"
        gSL_Select = gSL_Select & "納品書不作成区分 = '" & RTrim(.納品書不作成区分) & "',"
        gSL_Select = gSL_Select & "仕入先工区 = '" & RTrim(.仕入先工区) & "',"
        gSL_Select = gSL_Select & "置換工区 = '" & RTrim(.置換工区) & "',"
        gSL_Select = gSL_Select & "持込先 = '" & RTrim(.持込先) & "',"
        gSL_Select = gSL_Select & "持込先工場 = '" & RTrim(.持込先工場) & "',"
        gSL_Select = gSL_Select & "出庫部署 = '" & RTrim(.出庫部署) & "',"
        gSL_Select = gSL_Select & "担当 = '" & RTrim(.担当) & "',"
        gSL_Select = gSL_Select & "支給係 = '" & RTrim(.支給係) & "',"
        gSL_Select = gSL_Select & "内外区分 = '" & RTrim(.内外区分) & "',"
        gSL_Select = gSL_Select & "現旧区分 = '" & RTrim(.現旧区分) & "',"
        gSL_Select = gSL_Select & "整備 = '" & RTrim(.整備) & "',"
        gSL_Select = gSL_Select & "担当者 = '" & RTrim(.担当者) & "',"
        gSL_Select = gSL_Select & "支給先受入 = '" & RTrim(.支給先受入) & "',"
        gSL_Select = gSL_Select & "支給先整備 = '" & RTrim(.支給先整備) & "',"
        gSL_Select = gSL_Select & "支給先担当者 = '" & RTrim(.支給先担当者) & "',"
        gSL_Select = gSL_Select & "管理区分 = '" & RTrim(.管理区分) & "',"
        gSL_Select = gSL_Select & "生産区分 = '" & RTrim(.生産区分) & "',"
        gSL_Select = gSL_Select & "納品書発行枚数 = " & .納品書発行枚数 & ","
        gSL_Select = gSL_Select & "かんばん回転枚数 = " & .かんばん回転枚数 & ","
        gSL_Select = gSL_Select & "かんばん増減符号 = '" & RTrim(.かんばん増減符号) & "',"
        gSL_Select = gSL_Select & "かんばん増減枚数 = " & .かんばん増減枚数 & ","
        gSL_Select = gSL_Select & "当初内示数 = " & .当初内示数 & ","
        gSL_Select = gSL_Select & "変更数符号 = '" & RTrim(.変更数符号) & "',"
        gSL_Select = gSL_Select & "変更数 = " & .変更数 & ","
        gSL_Select = gSL_Select & "調整数符号 = '" & RTrim(.調整数符号) & "',"
        gSL_Select = gSL_Select & "調整数 = " & .調整数 & ","
        gSL_Select = gSL_Select & "納入区分 = '" & RTrim(.納入区分) & "',"
        For idx = 1 To 31
            gSL_Select = gSL_Select & "内示数量" & Format(idx, "00") & " = " & .内示数量(idx) & ","
        Next idx
        gSL_Select = gSL_Select & "変更回数 = " & .変更回数 & ","
        gSL_Select = gSL_Select & "ＡＢ区分 = '" & RTrim(.AB区分) & "',"
        gSL_Select = gSL_Select & "品名 = '" & RTrim(.品名) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "'"
        gSL_Select = gSL_Select & " Where FILECODE = '" & RTrim(.FILECODE) & "'"
        gSL_Select = gSL_Select & "   And HEAD作成日 = '" & RTrim(.HEAD作成日) & "'"
        gSL_Select = gSL_Select & "   And HEADSEQ = '" & RTrim(.HEADSEQ) & "'"
        gSL_Select = gSL_Select & "   And 連番 = " & .連番 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        .DBExec
    End With

End Sub

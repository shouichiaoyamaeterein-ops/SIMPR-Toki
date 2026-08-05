Attribute VB_Name = "発注内示テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   発注内示テーブル
'*
'*   モジュールID  : 発注内示テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*                                      /項目追加含む/
'*   作成日  :  2004/12/20  By Cis
'*   作成日  :  2008/06/09  By Cis(FMV-E630)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_発注内示テーブル
        内示年月          As String
        品番              As String
        手配先            As String
        受入              As String
        直送先            As String
        納期              As String
        仕掛日            As String
        手配日            As String
        数量              As Long
        処理区分          As Byte
        削除区分          As Byte
        変更回数          As Byte
        ローレベル        As Byte
        計画年月          As String
        作成者            As String
'        作成日            As Date
'        作成端末          As String
        更新者            As String
'        更新日            As Date
'        更新端末          As String
        件数          As Long
'   項目追加(自動)
        当月内示数     As Long
        翌月内示数     As Long
        翌々月内示数   As Long
        KB             As String
        DD             As Byte
        手配先名       As String
        品名           As String
        背番号         As String
        サイクル       As String
        収容数         As Long
        品番編集       As String
        手配区分名     As String
        増減率         As Integer
        今回計画数     As Long
        前回計画数     As Long
        計画差         As Long
        今回必要数     As Long
        前回必要数     As Long
        必要差         As Long
        表示品番       As String
        機械           As String
        ロット         As Long
        金額           As Currency
'-----------------------------------------------' テスト追加
        工順            As Integer
        機械名          As String
        機械NO          As String
        機械NO名        As String
        工程数          As Currency
        基準工数1       As Currency
        基準工数2       As Currency
       
    End Type

    Public HNJ As T_発注内示テーブル

    Public HNJ_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHNJ(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHNJ
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With HNJ
            Select Case adoCol.Name
                Case "内示年月":          If Not IsNull(adoCol.Value) Then .内示年月 = adoCol.Value
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "手配先":            If Not IsNull(adoCol.Value) Then .手配先 = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "直送先":            If Not IsNull(adoCol.Value) Then .直送先 = adoCol.Value
                Case "納期":              If Not IsNull(adoCol.Value) Then .納期 = adoCol.Value
                Case "仕掛日":            If Not IsNull(adoCol.Value) Then .仕掛日 = adoCol.Value
                Case "手配日":            If Not IsNull(adoCol.Value) Then .手配日 = adoCol.Value
                Case "数量":              If Not IsNull(adoCol.Value) Then .数量 = adoCol.Value
                Case "処理区分":          If Not IsNull(adoCol.Value) Then .処理区分 = adoCol.Value
                Case "削除区分":          If Not IsNull(adoCol.Value) Then .削除区分 = adoCol.Value
                Case "変更回数":          If Not IsNull(adoCol.Value) Then .変更回数 = adoCol.Value
                Case "ローレベル":        If Not IsNull(adoCol.Value) Then .ローレベル = adoCol.Value
                Case "計画年月":          If Not IsNull(adoCol.Value) Then .計画年月 = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
                Case "更新者":            If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "当月内示数":      If Not IsNull(adoCol.Value) Then .当月内示数 = adoCol.Value
                Case "翌月内示数":      If Not IsNull(adoCol.Value) Then .翌月内示数 = adoCol.Value
                Case "翌々月内示数":    If Not IsNull(adoCol.Value) Then .翌々月内示数 = adoCol.Value
                Case "KB":      If Not IsNull(adoCol.Value) Then .KB = adoCol.Value
                Case "DD":      If Not IsNull(adoCol.Value) Then .DD = adoCol.Value
                Case "手配先名":        If Not IsNull(adoCol.Value) Then .手配先名 = adoCol.Value
                Case "品名":    If Not IsNull(adoCol.Value) Then .品名 = adoCol.Value
                Case "背番号":  If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "サイクル":        If Not IsNull(adoCol.Value) Then .サイクル = adoCol.Value
                Case "収容数":  If Not IsNull(adoCol.Value) Then .収容数 = adoCol.Value
                Case "品番編集":        If Not IsNull(adoCol.Value) Then .品番編集 = adoCol.Value
                Case "手配区分名":      If Not IsNull(adoCol.Value) Then .手配区分名 = adoCol.Value
                Case "増減率":  If Not IsNull(adoCol.Value) Then .増減率 = adoCol.Value
                Case "今回計画数":      If Not IsNull(adoCol.Value) Then .今回計画数 = adoCol.Value
                Case "前回計画数":      If Not IsNull(adoCol.Value) Then .前回計画数 = adoCol.Value
                Case "計画差":  If Not IsNull(adoCol.Value) Then .計画差 = adoCol.Value
                Case "今回必要数":      If Not IsNull(adoCol.Value) Then .今回必要数 = adoCol.Value
                Case "前回必要数":      If Not IsNull(adoCol.Value) Then .前回必要数 = adoCol.Value
                Case "必要差":  If Not IsNull(adoCol.Value) Then .必要差 = adoCol.Value
                Case "表示品番":        If Not IsNull(adoCol.Value) Then .表示品番 = adoCol.Value
                Case "機械":    If Not IsNull(adoCol.Value) Then .機械 = adoCol.Value
                Case "ロット":  If Not IsNull(adoCol.Value) Then .ロット = adoCol.Value
                Case "金額":    If Not IsNull(adoCol.Value) Then .金額 = adoCol.Value
'-----------------------------------------------' テスト追加
                Case "工順":    If Not IsNull(adoCol.Value) Then .工順 = adoCol.Value '        工順            As integr
                Case "機械名":  If Not IsNull(adoCol.Value) Then .機械名 = adoCol.Value '        機械名          As String
                Case "機械NO":  If Not IsNull(adoCol.Value) Then .機械NO = adoCol.Value '        機械NO          As String
                Case "機械NO名":        If Not IsNull(adoCol.Value) Then .機械NO名 = adoCol.Value '        機械NO名        As String
                Case "工程数":          If Not IsNull(adoCol.Value) Then .工程数 = adoCol.Value
                Case "基準工数1":       If Not IsNull(adoCol.Value) Then .基準工数1 = adoCol.Value
                Case "基準工数2":       If Not IsNull(adoCol.Value) Then .基準工数2 = adoCol.Value
                
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function HNJRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HNJRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 発注内示テーブル"
        RTxtSql = RTxtSql & " Where 内示年月 = '" & RTrim(HNJ.内示年月) & "'"
        RTxtSql = RTxtSql & " And   品番 = '" & RTrim(HNJ.品番) & "'"
        RTxtSql = RTxtSql & " And   手配先 = '" & RTrim(HNJ.手配先) & "'"
        RTxtSql = RTxtSql & " And   受入 = '" & RTrim(HNJ.受入) & "'"
        RTxtSql = RTxtSql & " And   直送先 = '" & RTrim(HNJ.直送先) & "'"
        RTxtSql = RTxtSql & " And   納期 = '" & RTrim(HNJ.納期) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HNJRead = True
           Call ItemsFromHNJ(DBNo)
           HNJ_RDSTS = True
        Else
           HNJ_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call HNJClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HNJReadNext(Optional DBNo As Integer = 0)
    If Not HNJ_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HNJ_RDSTS = False
        Else
           Call ItemsFromHNJ(DBNo)
           HNJ_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HNJClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HNJ_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HNJ )
'*----------------------------------------------*'
Public Sub ItemsClearHNJ()
    With HNJ
        .内示年月 = ""
        .品番 = ""
        .手配先 = ""
        .受入 = ""
        .直送先 = ""
        .納期 = ""
        .仕掛日 = ""
        .手配日 = ""
        .数量 = 0
        .処理区分 = 0
        .削除区分 = 0
        .変更回数 = 0
        .ローレベル = 0
        .計画年月 = ""
        .作成者 = ""
'       .作成日 = 0
'       .作成端末 = ""
        .更新者 = ""
'       .更新日 = 0
'       .更新端末 = ""
        .件数 = 0
'   項目追加(自動)
        .当月内示数 = 0
        .翌月内示数 = 0
        .翌々月内示数 = 0
        .KB = ""
        .DD = 0
        .手配先名 = ""
        .品名 = ""
        .背番号 = ""
        .サイクル = ""
        .収容数 = 0
        .品番編集 = ""
        .手配区分名 = ""
        .増減率 = 0
        .今回計画数 = 0
        .前回計画数 = 0
        .計画差 = 0
        .今回必要数 = 0
        .前回必要数 = 0
        .必要差 = 0
        .表示品番 = ""
        .機械 = ""
        .ロット = 0
        .金額 = 0
'-----------------------------------------------' テスト追加
        .工順 = 0
        .機械名 = ""
        .機械NO = ""
        .機械NO名 = ""
        .工程数 = 0
        .基準工数1 = 0
        .基準工数2 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function HNJInsert(Optional DBNo As Integer) As Boolean
    HNJInsert = False
    With HNJ
        gSL_Select = "Insert Into 発注内示テーブル  ( "
        gSL_Select = gSL_Select & "内示年月,品番,手配先,"
        gSL_Select = gSL_Select & "受入,直送先,納期,"
        gSL_Select = gSL_Select & "仕掛日,手配日,数量,"
        gSL_Select = gSL_Select & "処理区分,削除区分,変更回数,"
        gSL_Select = gSL_Select & "ローレベル,計画年月,"
        gSL_Select = gSL_Select & "作成者,作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.内示年月) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.手配先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.直送先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納期) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.仕掛日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.手配日) & "',"
        gSL_Select = gSL_Select & "" & .数量 & ","
        gSL_Select = gSL_Select & "" & .処理区分 & ","
        gSL_Select = gSL_Select & "" & .削除区分 & ","
        gSL_Select = gSL_Select & "" & .変更回数 & ","
        gSL_Select = gSL_Select & "" & .ローレベル & ","
        gSL_Select = gSL_Select & "'" & RTrim(.計画年月) & "',"
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
    HNJInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function HNJUpdate(Optional DBNo As Integer) As Boolean
    HNJUpdate = False
    With HNJ
        gSL_Select = "Update 発注内示テーブル  Set "
        gSL_Select = gSL_Select & "内示年月 = '" & RTrim(.内示年月) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "手配先 = '" & RTrim(.手配先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "直送先 = '" & RTrim(.直送先) & "',"
        gSL_Select = gSL_Select & "納期 = '" & RTrim(.納期) & "',"
        gSL_Select = gSL_Select & "仕掛日 = '" & RTrim(.仕掛日) & "',"
        gSL_Select = gSL_Select & "手配日 = '" & RTrim(.手配日) & "',"
        gSL_Select = gSL_Select & "数量 = " & .数量 & ","
        gSL_Select = gSL_Select & "処理区分 = " & .処理区分 & ","
        gSL_Select = gSL_Select & "削除区分 = " & .削除区分 & ","
        gSL_Select = gSL_Select & "変更回数 = " & .変更回数 & ","
        gSL_Select = gSL_Select & "ローレベル = " & .ローレベル & ","
        gSL_Select = gSL_Select & "計画年月 = '" & RTrim(.計画年月) & "',"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 内示年月 = '" & RTrim(.内示年月) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 手配先 = '" & RTrim(.手配先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 直送先 = '" & RTrim(.直送先) & "'"
        gSL_Select = gSL_Select & "   And 納期 = '" & RTrim(.納期) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HNJUpdate = True
End Function

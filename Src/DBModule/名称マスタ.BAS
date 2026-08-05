Attribute VB_Name = "名称マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   名称マスタ
'*
'*   モジュールID  : 名称マスタ.Bas
'*
'*   作成日  :  2003/12/08  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_名称マスタ
        区分名称          As String
        区分タイプ        As String
        値                As String
        値名称            As String
        区分桁数          As Integer
        英数字1           As String
        英数字2           As String
        英数字3           As String
        数字1             As Currency
        数字2             As Currency
        数字3             As Currency
        制御区分          As String
        作成日            As Date
        作成者            As String
'        作成端末          As String
        更新日            As Date
        更新者            As String
'        更新端末          As String
        件数          As Long
    
        検索名          As String
        コード名        As String
        名称名          As String
    End Type

    Public SYM As T_名称マスタ

    Public SYM_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object

    Public SYMCreate As T_名称マスタ
    Public SYM_Meisyo   As String
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromSYM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearSYM
    
    CisDB.ConnectNo = DBNo
    
    For Each adoCol In CisDB.Recordset.Fields
        With SYM
            Select Case adoCol.Name
                Case "区分名称":          If Not IsNull(adoCol.Value) Then .区分名称 = adoCol.Value
                Case "区分タイプ":        If Not IsNull(adoCol.Value) Then .区分タイプ = adoCol.Value
                Case "値":                If Not IsNull(adoCol.Value) Then .値 = adoCol.Value
                Case "値名称":            If Not IsNull(adoCol.Value) Then .値名称 = adoCol.Value
                Case "区分桁数":          If Not IsNull(adoCol.Value) Then .区分桁数 = adoCol.Value
                Case "英数字1":           If Not IsNull(adoCol.Value) Then .英数字1 = adoCol.Value
                Case "英数字2":           If Not IsNull(adoCol.Value) Then .英数字2 = adoCol.Value
                Case "英数字3":           If Not IsNull(adoCol.Value) Then .英数字3 = adoCol.Value
                Case "数字1":             If Not IsNull(adoCol.Value) Then .数字1 = adoCol.Value
                Case "数字2":             If Not IsNull(adoCol.Value) Then .数字2 = adoCol.Value
                Case "数字3":             If Not IsNull(adoCol.Value) Then .数字3 = adoCol.Value
                Case "制御区分":          If Not IsNull(adoCol.Value) Then .制御区分 = adoCol.Value
                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日 = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日 = adoCol.Value
                Case "更新者":            If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
                
                Case "検索名":       If Not IsNull(adoCol.Value) Then .検索名 = adoCol.Value
                Case "コード名":       If Not IsNull(adoCol.Value) Then .コード名 = adoCol.Value
                Case "名称名":       If Not IsNull(adoCol.Value) Then .名称名 = adoCol.Value
                
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function SYMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    SYMRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 名称マスタ"
        RTxtSql = RTxtSql & " Where 区分名称 = '" & RTrim(SYM.区分名称) & "'"
        RTxtSql = RTxtSql & " And   区分タイプ = '" & RTrim(SYM.区分タイプ) & "'"
        RTxtSql = RTxtSql & " And   値 = '" & RTrim(SYM.値) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           SYMRead = True
           Call ItemsFromSYM(DBNo)
           SYM_RDSTS = True
        Else
           SYM_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call SYMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub SYMReadNext(Optional DBNo As Integer = 0)
    If Not SYM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           SYM_RDSTS = False
        Else
           Call ItemsFromSYM(DBNo)
           SYM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub SYMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        SYM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( SYM )
'*----------------------------------------------*'
Public Sub ItemsClearSYM()
    With SYM
        .区分名称 = ""
        .区分タイプ = ""
        .値 = ""
        .値名称 = ""
        .区分桁数 = 0
        .英数字1 = ""
        .英数字2 = ""
        .英数字3 = ""
        .数字1 = 0
        .数字2 = 0
        .数字3 = 0
        .制御区分 = ""
       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
        .件数 = 0
    
        .検索名 = ""
        .コード名 = ""
        .名称名 = ""
    
    End With
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( SYM )
'*----------------------------------------------*'
Public Sub ItemsClearSYMCreate()
    With SYMCreate
        .区分名称 = ""
        .区分タイプ = ""
        .値 = ""
        .値名称 = ""
        .区分桁数 = 0
        .英数字1 = ""
        .英数字2 = ""
        .英数字3 = ""
        .数字1 = 0
        .数字2 = 0
        .数字3 = 0
        .制御区分 = ""
       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
        .件数 = 0
    
        .検索名 = ""
        .コード名 = ""
        .名称名 = ""
    End With
End Sub

'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Sub SYMInsert(Optional DBNo As Integer)
    With SYM
        gSL_Select = "Insert Into 名称マスタ  ( "
        gSL_Select = gSL_Select & "区分名称,区分タイプ,値,"
        gSL_Select = gSL_Select & "値名称,区分桁数,英数字1,"
        gSL_Select = gSL_Select & "英数字2,英数字3,数字1,"
        gSL_Select = gSL_Select & "数字2,数字3,制御区分,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.区分名称) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.区分タイプ) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.値) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.値名称) & "',"
        gSL_Select = gSL_Select & "" & .区分桁数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.英数字1) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.英数字2) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.英数字3) & "',"
        gSL_Select = gSL_Select & "" & .数字1 & ","
        gSL_Select = gSL_Select & "" & .数字2 & ","
        gSL_Select = gSL_Select & "" & .数字3 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.制御区分) & "',"
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "',"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"
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
Public Sub SYMUpdate(Optional DBNo As Integer)
    With SYM
        gSL_Select = "Update 名称マスタ  Set "
        gSL_Select = gSL_Select & "区分名称 = '" & RTrim(.区分名称) & "',"
        gSL_Select = gSL_Select & "区分タイプ = '" & RTrim(.区分タイプ) & "',"
        gSL_Select = gSL_Select & "値 = '" & RTrim(.値) & "',"
        gSL_Select = gSL_Select & "値名称 = '" & RTrim(.値名称) & "',"
        gSL_Select = gSL_Select & "区分桁数 = " & .区分桁数 & ","
        gSL_Select = gSL_Select & "英数字1 = '" & RTrim(.英数字1) & "',"
        gSL_Select = gSL_Select & "英数字2 = '" & RTrim(.英数字2) & "',"
        gSL_Select = gSL_Select & "英数字3 = '" & RTrim(.英数字3) & "',"
        gSL_Select = gSL_Select & "数字1 = " & .数字1 & ","
        gSL_Select = gSL_Select & "数字2 = " & .数字2 & ","
        gSL_Select = gSL_Select & "数字3 = " & .数字3 & ","
        gSL_Select = gSL_Select & "制御区分 = '" & RTrim(.制御区分) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 区分名称 = '" & RTrim(.区分名称) & "'"
        gSL_Select = gSL_Select & "   And 区分タイプ = '" & RTrim(.区分タイプ) & "'"
        gSL_Select = gSL_Select & "   And 値 = '" & RTrim(.値) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        .DBExec
    End With

End Sub
'*--------------------------------------------------------------------------------
'*
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ
'*              ( 管理レコードを確認し、なければ作成後指定データを書き込む )
'*
'*      モジュールID    : SYMReadChk
'*
'*      初期設定項目    : ユーザ定義(T_区分名称マスタ)で作成されたSYM
'*                       ( 管理レコードチェックを行う場合には、
'*                          区分タイプ・区分桁数を忘れずに・・・！！ )
'*
'*          引数        :  Chk       =   管理レコードチェックの有無 ( Optional )
'*                                      < 0 : 管理ﾚｺｰﾄﾞ、ﾃﾞｰﾀﾚｺｰﾄﾞ > 初期値
'*                                      < 1 : 管理ﾚｺｰﾄﾞ >
'*                                      < 2 : ﾃﾞｰﾀﾚｺｰﾄﾞ >
'*          　　        : DBNo      =   データベースNO ( Optional )
'*
'*          戻値        : True Or False
'*
'*      作成日          :   2001/05/21   By Cis
'*      変更日          :
'*
' 使用例
'' 名称マスタ検索　確認
'    Call ItemsClearSYM
'    With SYM
'        .区分名称 = "名称検索"
'        .区分桁数 = 2
'        .値 = "01"
'        .値名称 = "名称検索"
'        .英数字1 = "区   分   検   索"
'        .英数字2 = "値"
'        .英数字3 = "区分名"
'        .制御区分 = "X"
'    End With
'    Call SYMReadChk
'*---------------------------------------------------------------------------------
Public Function SYMReadChk(Optional DBNo As Integer = 0, Optional DataPut As Boolean = True)
    
    SYMReadChk = False
    
    With CisDB
        .SQL = "SELECT * FROM 名称マスタ"
        .SQL = .SQL & " WHERE 区分名称 = '" & RTrim(SYM.区分名称) & "'"
        .SQL = .SQL & "   AND 区分タイプ  = ''"
        .SQL = .SQL & "   AND 値   = '" & RTrim(SYM.値) & "'"
        If .DBRead(0, DBNo) Then
            Call ItemsFromSYM
            .Recordset.Close
            SYMReadChk = True
            Exit Function
        End If
    End With
' データレコード作成
    If DataPut Then
       Call SYMInsert
    End If
' 管理レコード作成
    With CisDB
        .SQL = "SELECT * FROM 名称マスタ"
        .SQL = .SQL & " WHERE 区分名称 = '" & RTrim(SYM.区分名称) & "'"
        .SQL = .SQL & "   AND 区分タイプ  = 'K'"
        If Not .DBRead(1, DBNo) Then
            With SYM
                .区分タイプ = "K"
                .値 = ""
                .値名称 = ""
                .英数字1 = ""
                .英数字2 = ""
                .英数字3 = ""
                .数字1 = 0
                .数字2 = 0
                .数字3 = 0
            End With
            Call SYMInsert
        End If
        On Error Resume Next
        .Recordset.Close
        On Error GoTo 0
    End With
    SYMReadChk = True
End Function
'*--------------------------------------------------------------------------------
'*
'*      名称マスタ　初期内容設定
'*              ( 管理レコードを確認し、なければ作成後指定データを書き込む )
'*              ( 検索名が空白以外の時、検索用データの出力も行う<LockFlg=Trueの時>)
'*      モジュールID    : SYMInitCreate
'*
'*      初期設定項目    : ユーザ定義(T_区分名称マスタ)で作成されたSYMCreate
'*                       ( 管理レコードチェックを行う場合には、
'*                          区分タイプ・区分桁数を忘れずに・・・！！ )
'*
'*          引数        : LockFlg   =   管理レコード検索 ( Optional )
'*          　　        : DBNo      =   データベースNO ( Optional )
'*
'*          戻値        : True      : 管理レコードあり
'*                      : False     : 管理レコードなし
'*                  ※ LockFlg = True の時有効
'*      作成日          :   2003/11/18   By Cis
'*      変更日          :
'*
' 使用例
'' 名称マスタ検索　確認
'    Call ItemsClearSYMCreate
'    With SYMCreate
'        .区分名称 = "伝票種類"
'        .区分桁数 = 1
'        .値 = "1"
'        .値名称 = "チケット"
'        .作成者 = gTanto
'        .検索名 = "伝 票 種 類 検 索"
'        .コード名 = "伝票種"
'        .名称名 = "伝票種類名"
'    End With
'    If Not SYMInitCreate(True) Then
'        With SYMCreate
'            .値 = "2"
'            .値名称 = "ＪＡＭＡ"
'        End With
'        Call SYMInitCreate
'    End If
'*---------------------------------------------------------------------------------
Public Function SYMInitCreate(Optional LockFlg As Boolean = False, Optional DBNo As Integer = 0) As Boolean
    Dim CrtCode         As String
    
    SYMInitCreate = False
    
    If LockFlg Then
        gSL_Select = " SELECT 値名称 FROM 名称マスタ "
        gSL_Select = gSL_Select & " WHERE 区分名称 = '" & SYMCreate.区分名称 & "' "
        gSL_Select = gSL_Select & "   AND 区分タイプ = 'K' "
        If SYMRead(gSL_Select, 1, DBNo) Then
            SYMInitCreate = True
            Exit Function
        End If
    End If
    Call ItemsClearSYM
    With SYM
        .区分名称 = SYMCreate.区分名称
        .区分桁数 = SYMCreate.区分桁数
        .値 = SYMCreate.値
        .値名称 = SYMCreate.値名称
        .英数字1 = SYMCreate.英数字1
        .英数字2 = SYMCreate.英数字2
        .英数字3 = SYMCreate.英数字3
        .数字1 = SYMCreate.数字1
        .数字2 = SYMCreate.数字2
        .数字3 = SYMCreate.数字3
        .作成者 = SYMCreate.作成者
        .制御区分 = SYMCreate.制御区分
    End With
    Call SYMReadChk(DBNo)
    
    If LockFlg And Trim(SYMCreate.検索名) <> "" Then
        gSL_Select = "SELECT 値名称 FROM 名称マスタ"
        gSL_Select = gSL_Select & " WHERE 区分名称 = '名称検索'"
        gSL_Select = gSL_Select & "   And 値名称 = '" & SYMCreate.区分名称 & "'"
        If Not SYMRead(gSL_Select, 1, DBNo) Then
            gSL_Select = "SELECT max(値)  値 FROM 名称マスタ"
            gSL_Select = gSL_Select & " WHERE 区分名称 = '名称検索'"
            If SYMRead(gSL_Select, 1, DBNo) Then
               CrtCode = Format(CisFun.Val2(SYM.値) + 1, "00")
            Else
               CrtCode = "01"
            End If
            Call ItemsClearSYM
            With SYM
                .区分名称 = "名称検索"
                .区分桁数 = 2
                .値 = CrtCode
                .値名称 = SYMCreate.区分名称
                .英数字1 = SYMCreate.検索名
                .英数字2 = SYMCreate.コード名
                .英数字3 = SYMCreate.名称名
                .数字1 = SYMCreate.数字1
                .数字2 = SYMCreate.数字2
                .数字3 = SYMCreate.数字3
                .制御区分 = "X"
                .作成者 = SYMCreate.作成者
            End With
            Call SYMReadChk(DBNo)
        End If
    End If

End Function
'*--------------------------------------------------------------------------------
'*
'*      名称を求める
'*
'*      モジュールID    : MeisyoGet
'*
'*
'*          引数        : Kubun   =   区分名
'*                      : Atai    =   値
'*          　　        : DBNo    =   データベースNO ( Optional )
'*
'*          戻値        : True Or False
'*                      : SYM_Meisyo : 値名称
'*
'*      作成日          :   2001/06/11   By Cis
'*      変更日          :
'*
'*---------------------------------------------------------------------------------
Public Function MeisyoGet(Kubun As String, Atai As String, Optional DBNo As Integer = 0) As Boolean
    MeisyoGet = False
    SYM_Meisyo = ""
    
    gSL_Select = "SELECT * FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = '" & RTrim(Kubun) & "'"
    gSL_Select = gSL_Select & "   AND 値 = '" & RTrim(Atai) & "'"
    gSL_Select = gSL_Select & "   AND 区分ﾀｲﾌﾟ = ''"
    If Not SYMRead(gSL_Select, 1, DBNo) Then Exit Function

    MeisyoGet = True
    SYM_Meisyo = RTrim(SYM.値名称)
    
End Function


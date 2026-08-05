Attribute VB_Name = "TEXT管理テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TEXT管理テーブル
'*
'*   モジュールID  : TEXT管理テーブル.Bas
'*
'*                                      /項目追加含む/
'*   作成日  :  2003/12/12  By Cis      GRP名･レイアウト名 項目追加
'*   変更日  :  2003/12/13  By Cis      初期設定内容 初期データセット ﾓｼﾞｭｰﾙ追加
'*   変更日  :  2005/06/13  By Cis      レイアウト情報内示 レイアウト番号(2)追加
'*                                                 補給内示ﾃﾞｰﾀ(トヨタ紡織形式)
'*   変更日  :  2006/01/25  By Cis      レイアウト情報内示 レイアウト番号(3)追加
'*                                                 内示ﾃﾞｰﾀ(新レイアウト形式)
'*   作成日  :  2007/03/05  By Cis
'*   変更日  :  2008/05/08  By CIS      車体精工コード追加
'*---------------------------------------------------------------------------------
    Type T_TEXT管理テーブル
        FILECD                As String
        発信元                As String
        FCODE                 As String
        F名称                 As String
        識別名                As String
        変換元ファイル名      As String
        累積元ファイル名      As String
        TNS区分               As Integer
        TNSタイプ             As Integer
        受信処理              As Integer
        受信方法              As Integer
        TNSFMT                As Integer
        TEXTレイアウト        As Byte
        契約先                As String
        最終処理日時          As Date
        処理件数              As Long
        最終処理              As String
        処理結果              As String
        有効長                As Long
        連続実行              As Integer
        実行GRP               As String
        チェックリスト        As Integer
        リスト累積期間        As Integer
        オプション1           As Integer
        オプション2           As Integer
        オプション3           As Integer
        オプション4           As Integer
        オプション5           As Integer
        オプション6           As Integer
        保持期間              As Long
        自社CD                As String
'        作成日                As Date
'        作成端末              As String
        作成者                As String
'        更新日                As Date
'        更新端末              As String
        更新者                As String
        件数          As Long
'   項目追加(自動)
        GRP名          As String
        レイアウト名   As String
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT START
'   項目追加(手動)
        車体精工コード        As String
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT END
    End Type

    Public TKT As T_TEXT管理テーブル

    Public TKT_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTKT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTKT
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With TKT
            Select Case adoCol.Name
                Case "FILECD":                If Not IsNull(adoCol.Value) Then .FILECD = adoCol.Value
                Case "発信元":                If Not IsNull(adoCol.Value) Then .発信元 = adoCol.Value
                Case "FCODE":                 If Not IsNull(adoCol.Value) Then .FCODE = adoCol.Value
                Case "F名称":                 If Not IsNull(adoCol.Value) Then .F名称 = adoCol.Value
                Case "識別名":                If Not IsNull(adoCol.Value) Then .識別名 = adoCol.Value
                Case "変換元ﾌｧｲﾙ名":          If Not IsNull(adoCol.Value) Then .変換元ファイル名 = adoCol.Value
                Case "累積元ﾌｧｲﾙ名":          If Not IsNull(adoCol.Value) Then .累積元ファイル名 = adoCol.Value
                Case "TNS区分":               If Not IsNull(adoCol.Value) Then .TNS区分 = adoCol.Value
                Case "TNSタイプ":             If Not IsNull(adoCol.Value) Then .TNSタイプ = adoCol.Value
                Case "受信処理":              If Not IsNull(adoCol.Value) Then .受信処理 = adoCol.Value
                Case "受信方法":              If Not IsNull(adoCol.Value) Then .受信方法 = adoCol.Value
                Case "TNSFMT":                If Not IsNull(adoCol.Value) Then .TNSFMT = adoCol.Value
                Case "TEXTレイアウト":        If Not IsNull(adoCol.Value) Then .TEXTレイアウト = adoCol.Value
                Case "契約先":                If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
                Case "最終処理日時":          If Not IsNull(adoCol.Value) Then .最終処理日時 = adoCol.Value
                Case "処理件数":              If Not IsNull(adoCol.Value) Then .処理件数 = adoCol.Value
                Case "最終処理":              If Not IsNull(adoCol.Value) Then .最終処理 = adoCol.Value
                Case "処理結果":              If Not IsNull(adoCol.Value) Then .処理結果 = adoCol.Value
                Case "有効長":                If Not IsNull(adoCol.Value) Then .有効長 = adoCol.Value
                Case "連続実行":              If Not IsNull(adoCol.Value) Then .連続実行 = adoCol.Value
                Case "実行GRP":               If Not IsNull(adoCol.Value) Then .実行GRP = adoCol.Value
                Case "チェックリスト":        If Not IsNull(adoCol.Value) Then .チェックリスト = adoCol.Value
                Case "リスト累積期間":        If Not IsNull(adoCol.Value) Then .リスト累積期間 = adoCol.Value
                Case "ｵﾌﾟｼｮﾝ1":               If Not IsNull(adoCol.Value) Then .オプション1 = adoCol.Value
                Case "ｵﾌﾟｼｮﾝ2":               If Not IsNull(adoCol.Value) Then .オプション2 = adoCol.Value
                Case "ｵﾌﾟｼｮﾝ3":               If Not IsNull(adoCol.Value) Then .オプション3 = adoCol.Value
                Case "ｵﾌﾟｼｮﾝ4":               If Not IsNull(adoCol.Value) Then .オプション4 = adoCol.Value
                Case "ｵﾌﾟｼｮﾝ5":               If Not IsNull(adoCol.Value) Then .オプション5 = adoCol.Value
                Case "ｵﾌﾟｼｮﾝ6":               If Not IsNull(adoCol.Value) Then .オプション6 = adoCol.Value
                Case "保持期間":              If Not IsNull(adoCol.Value) Then .保持期間 = adoCol.Value
                Case "自社CD":                If Not IsNull(adoCol.Value) Then .自社CD = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "作成者":                If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "更新日":                If Not IsNull(adoCol.Value) Then .更新日               = adoCol.Value
'                Case "更新端末":              If Not IsNull(adoCol.Value) Then .更新端末             = adoCol.Value
                Case "更新者":                If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "GRP名":   If Not IsNull(adoCol.Value) Then .GRP名 = adoCol.Value
                Case "レイアウト名":    If Not IsNull(adoCol.Value) Then .レイアウト名 = adoCol.Value
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT START
'   項目追加(自動)
                Case "車体精工コード":        If Not IsNull(adoCol.Value) Then .車体精工コード = adoCol.Value
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT END
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TKTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TKTRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From TEXT管理テーブル"
        RTxtSql = RTxtSql & " Where FILECD = '" & RTrim(TKT.FILECD) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TKTRead = True
           Call ItemsFromTKT(DBNo)
           TKT_RDSTS = True
        Else
           TKT_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call TKTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TKTReadNext(Optional DBNo As Integer = 0)
    If Not TKT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TKT_RDSTS = False
        Else
           Call ItemsFromTKT(DBNo)
           TKT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TKTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TKT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TKT )
'*----------------------------------------------*'
Public Sub ItemsClearTKT()
    With TKT
        .FILECD = ""
        .発信元 = ""
        .FCODE = ""
        .F名称 = ""
        .識別名 = ""
        .変換元ファイル名 = ""
        .累積元ファイル名 = ""
        .TNS区分 = 0
        .TNSタイプ = 0
        .受信処理 = 0
        .受信方法 = 0
        .TNSFMT = 0
        .TEXTレイアウト = 0
        .契約先 = ""
        .最終処理日時 = 0
        .処理件数 = 0
        .最終処理 = ""
        .処理結果 = ""
        .有効長 = 0
        .連続実行 = 0
        .実行GRP = ""
        .チェックリスト = 0
        .リスト累積期間 = 0
        .オプション1 = 0
        .オプション2 = 0
        .オプション3 = 0
        .オプション4 = 0
        .オプション5 = 0
        .オプション6 = 0
        .保持期間 = 0
        .自社CD = ""
'       .作成日 = 0
'       .作成端末 = ""
        .作成者 = ""
'       .更新日 = 0
'       .更新端末 = ""
        .更新者 = ""
        .件数 = 0
'   項目追加(自動)
        .GRP名 = ""
        .レイアウト名 = ""
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT START
'   項目追加(手動)
        .車体精工コード = ""
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT END
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function TKTInsert(Optional DBNo As Integer) As Boolean
    TKTInsert = True
    With TKT
        gSL_Select = "Insert Into TEXT管理テーブル  ( "
        gSL_Select = gSL_Select & "FILECD,発信元,FCODE,"
        gSL_Select = gSL_Select & "F名称,識別名,変換元ﾌｧｲﾙ名,"
        gSL_Select = gSL_Select & "累積元ﾌｧｲﾙ名,TNS区分,TNSタイプ,"
        gSL_Select = gSL_Select & "受信処理,受信方法,TNSFMT,"
        gSL_Select = gSL_Select & "TEXTレイアウト,契約先,最終処理日時,"
        gSL_Select = gSL_Select & "処理件数,最終処理,処理結果,"
        gSL_Select = gSL_Select & "有効長,連続実行,実行GRP,"
        gSL_Select = gSL_Select & "チェックリスト,リスト累積期間,ｵﾌﾟｼｮﾝ1,"
        gSL_Select = gSL_Select & "ｵﾌﾟｼｮﾝ2,ｵﾌﾟｼｮﾝ3,ｵﾌﾟｼｮﾝ4,"
        gSL_Select = gSL_Select & "ｵﾌﾟｼｮﾝ5,ｵﾌﾟｼｮﾝ6,保持期間,"
        gSL_Select = gSL_Select & "自社CD,"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT START
        gSL_Select = gSL_Select & "車体精工コード,"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT END
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.FILECD) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発信元) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.FCODE) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.F名称) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.識別名) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.変換元ファイル名) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.累積元ファイル名) & "',"
        gSL_Select = gSL_Select & "" & .TNS区分 & ","
        gSL_Select = gSL_Select & "" & .TNSタイプ & ","
        gSL_Select = gSL_Select & "" & .受信処理 & ","
        gSL_Select = gSL_Select & "" & .受信方法 & ","
        gSL_Select = gSL_Select & "" & .TNSFMT & ","
        gSL_Select = gSL_Select & "" & .TEXTレイアウト & ","
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.最終処理日時) & "',"
        gSL_Select = gSL_Select & "" & .処理件数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.最終処理) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.処理結果) & "',"
        gSL_Select = gSL_Select & "" & .有効長 & ","
        gSL_Select = gSL_Select & "" & .連続実行 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.実行GRP) & "',"
        gSL_Select = gSL_Select & "" & .チェックリスト & ","
        gSL_Select = gSL_Select & "" & .リスト累積期間 & ","
        gSL_Select = gSL_Select & "" & .オプション1 & ","
        gSL_Select = gSL_Select & "" & .オプション2 & ","
        gSL_Select = gSL_Select & "" & .オプション3 & ","
        gSL_Select = gSL_Select & "" & .オプション4 & ","
        gSL_Select = gSL_Select & "" & .オプション5 & ","
        gSL_Select = gSL_Select & "" & .オプション6 & ","
        gSL_Select = gSL_Select & "" & .保持期間 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.自社CD) & "',"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT START
        gSL_Select = gSL_Select & "'" & RTrim(.車体精工コード) & "',"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT END
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "'"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TKTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function TKTUpdate(Optional DBNo As Integer) As Boolean
    TKTUpdate = False
    With TKT
        gSL_Select = "Update TEXT管理テーブル  Set "
        gSL_Select = gSL_Select & "FILECD = '" & RTrim(.FILECD) & "',"
        gSL_Select = gSL_Select & "発信元 = '" & RTrim(.発信元) & "',"
        gSL_Select = gSL_Select & "FCODE = '" & RTrim(.FCODE) & "',"
        gSL_Select = gSL_Select & "F名称 = '" & RTrim(.F名称) & "',"
        gSL_Select = gSL_Select & "識別名 = '" & RTrim(.識別名) & "',"
        gSL_Select = gSL_Select & "変換元ﾌｧｲﾙ名 = '" & RTrim(.変換元ファイル名) & "',"
        gSL_Select = gSL_Select & "累積元ﾌｧｲﾙ名 = '" & RTrim(.累積元ファイル名) & "',"
        gSL_Select = gSL_Select & "TNS区分 = " & .TNS区分 & ","
        gSL_Select = gSL_Select & "TNSタイプ = " & .TNSタイプ & ","
        gSL_Select = gSL_Select & "受信処理 = " & .受信処理 & ","
        gSL_Select = gSL_Select & "受信方法 = " & .受信方法 & ","
        gSL_Select = gSL_Select & "TNSFMT = " & .TNSFMT & ","
        gSL_Select = gSL_Select & "TEXTレイアウト = " & .TEXTレイアウト & ","
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "最終処理日時 = '" & RTrim(.最終処理日時) & "',"
        gSL_Select = gSL_Select & "処理件数 = " & .処理件数 & ","
        gSL_Select = gSL_Select & "最終処理 = '" & RTrim(.最終処理) & "',"
        gSL_Select = gSL_Select & "処理結果 = '" & RTrim(.処理結果) & "',"
        gSL_Select = gSL_Select & "有効長 = " & .有効長 & ","
        gSL_Select = gSL_Select & "連続実行 = " & .連続実行 & ","
        gSL_Select = gSL_Select & "実行GRP = '" & RTrim(.実行GRP) & "',"
        gSL_Select = gSL_Select & "チェックリスト = " & .チェックリスト & ","
        gSL_Select = gSL_Select & "リスト累積期間 = " & .リスト累積期間 & ","
        gSL_Select = gSL_Select & "ｵﾌﾟｼｮﾝ1 = " & .オプション1 & ","
        gSL_Select = gSL_Select & "ｵﾌﾟｼｮﾝ2 = " & .オプション2 & ","
        gSL_Select = gSL_Select & "ｵﾌﾟｼｮﾝ3 = " & .オプション3 & ","
        gSL_Select = gSL_Select & "ｵﾌﾟｼｮﾝ4 = " & .オプション4 & ","
        gSL_Select = gSL_Select & "ｵﾌﾟｼｮﾝ5 = " & .オプション5 & ","
        gSL_Select = gSL_Select & "ｵﾌﾟｼｮﾝ6 = " & .オプション6 & ","
        gSL_Select = gSL_Select & "保持期間 = " & .保持期間 & ","
        gSL_Select = gSL_Select & "自社CD = '" & RTrim(.自社CD) & "',"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT START
        gSL_Select = gSL_Select & "車体精工コード = '" & RTrim(.車体精工コード) & "',"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT END
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "'"
        gSL_Select = gSL_Select & " Where FILECD = '" & RTrim(.FILECD) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TKTUpdate = True
End Function
'+-----------------------------------+
'+      初期設定内容 初期データセット
'+-----------------------------------+
Public Sub TEXTInitAdd()
'-------------------------- 名称マスタ初期設定
' レイアウト情報確認
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "レイアウト情報内示"
        .区分桁数 = 1
        .値 = "0"
        .値名称 = "部品納入内示表(トヨタ形式)"
        .数字1 = 1
        .制御区分 = "X"
        .作成者 = gTanto
        If Not SYMInitCreate(True) Then
            .値 = "1"
            .値名称 = "納入内示ﾃﾞｰﾀB(トヨタ紡織形式)"
            .数字1 = 1
            Call SYMInitCreate
        
'            .値 = "2"
'            .値名称 = "納入内示ﾃﾞｰﾀ(トヨタ車体形式)"
'            .数字1 = 1
'            Call SYMInitCreate
        End If
    End With
    
'+------------------------------------------------------------- 2005/06/13 Insert
' レイアウト情報確認
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "レイアウト情報内示"
        .区分桁数 = 1
        .値 = "2"
        .値名称 = "補給内示ﾃﾞｰﾀ(トヨタ紡織形式)"
        .数字1 = 1
        .制御区分 = "X"
        .作成者 = gTanto
        If Not SYMInitCreate(False) Then
            .値 = "2"
            .値名称 = "補給内示ﾃﾞｰﾀ(トヨタ紡織形式)"
            .数字1 = 1
            Call SYMInitCreate
        End If
    End With
'+------------------------------------------------------------- 2005/06/13 Insert End
'+------------------------------------------------------------- 2006/01/25 Insert
' レイアウト情報確認
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "レイアウト情報内示"
        .区分桁数 = 1
        .値 = "3"
        .値名称 = "納入内示ﾃﾞｰﾀ(新ﾚｲｱｳﾄ形式)"
        .数字1 = 1
        .制御区分 = "X"
        .作成者 = gTanto
        If Not SYMInitCreate(False) Then
        End If
    End With
'+------------------------------------------------------------- 2006/01/25 Insert End
'+------------------------------------------------------------- 2007/03/05 Insert
' レイアウト情報確認
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "レイアウト情報手配"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "納品書手配ﾃﾞｰﾀ(トヨタ紡織形式)"
        .数字1 = 1
        .制御区分 = "X"
        .作成者 = gTanto
        If Not SYMInitCreate(False) Then
            .値 = "2"
            .値名称 = "納品書手配ﾃﾞｰﾀ(トヨタ車体形式)"
            .数字1 = 2
            Call SYMInitCreate
        End If
    End With
'--------------------------------------------------------------- 2007/03/12 Insert
    Exit Sub
'--------------------------------------------------------------- 2007/03/12 Insert End
'+------------------------------------------------------------- 2007/03/05 Insert End
    gSL_Select = "SELECT COUNT(*) AS 件数 FROM TEXT管理テーブル "
    Call TKTRead(gSL_Select, 1)
    If TKT.件数 <> 0 Then Exit Sub
' 内示データ（月次）
    Call ItemsClearTKT
    With TKT
        .FILECD = "T1000D0A"
        .発信元 = "T1000"
        .FCODE = "D0A"
        .F名称 = "部品納入内示表(月次)" & "(T1000D0A)"
        .識別名 = "ITNSD0A"
        .TNSタイプ = 0
        .変換元ファイル名 = "T1000D0A.Apl"
        .連続実行 = 0
        .実行GRP = "01"
        .受信処理 = 1
        .TNSFMT = 1
        .有効長 = 250
        .TNS区分 = 0
        .TEXTレイアウト = 0
        .契約先 = "1000"
        .作成者 = "SIMPR-A"
        
        .オプション2 = 3
    End With
    Call TKTInsert
' 内示データ（計画変更）
    Call ItemsClearTKT
    With TKT
        .FILECD = "T1000D0B"
        .発信元 = "T1000"
        .FCODE = "D0B"
        .F名称 = "部品納入内示表(計変)" & "(T1000D0B)"
        .識別名 = "ITNSD0B"
        .TNSタイプ = 0
        .変換元ファイル名 = "T1000D0B.Apl"
        .連続実行 = 0
        .実行GRP = "01"
        .受信処理 = 1
        .TNSFMT = 1
        .有効長 = 250
        .TNS区分 = 0
        .TEXTレイアウト = 0
        .契約先 = "1000"
        .作成者 = "SIMPR-A"

        .オプション2 = 3
        .オプション3 = 1
    End With
    Call TKTInsert
' 内示データ（補給）
    Call ItemsClearTKT
    With TKT
        .FILECD = "T1000V36"
        .発信元 = "T1000"
        .FCODE = "D0B"
        .F名称 = "部品納入内示表(補給)" & "(T1000V36)"
        .識別名 = "ITNSV36"
        .TNSタイプ = 0
        .変換元ファイル名 = "T1000V36.Apl"
        .連続実行 = 0
        .実行GRP = ""
        .受信処理 = 1
        .TNSFMT = 1
        .有効長 = 250
        .TNS区分 = 0
        .TEXTレイアウト = 0
        .契約先 = "1000"
        .作成者 = "SIMPR-A"
        .オプション2 = 3
    End With
    Call TKTInsert
' 内示データ（アラコ）
    Call ItemsClearTKT
    With TKT
        .FILECD = "NARC030"
        .発信元 = " 7364"
        .FCODE = "030"
        .F名称 = "納入内示データＢ" & "(NARC030)"
        .識別名 = "NARC030"
        .TNSタイプ = 0
        .変換元ファイル名 = "NARC030.*"
        .連続実行 = 0
        .実行GRP = ""
        .受信処理 = 1
        .TNSFMT = 1
        .有効長 = 250
        .TNS区分 = 0
        .TEXTレイアウト = 1
        '.契約先 = "1097"
        .契約先 = "7364"
        .作成者 = "SIMPR-A"
        .オプション2 = 3
    End With
    Call TKTInsert
' グループ情報確認
    gSL_Select = "Select * From Irnaksic"
    gSL_Select = gSL_Select & " Where bk = 'Group'"
    gSL_Select = gSL_Select & "   AND oyian = '01'"

    If Not IRNRead(gSL_Select, 1) Then
        With IRN
            Call ItemsClearIRN
            .bk = "Group"
            .rno = 0
            .oyian = "01"
            .oyian2 = "内示処理"
        End With
        Call IRNInsert
    End If
''-------------------------- 名称マスタ初期設定
'' レイアウト情報確認
'    Call ItemsClearSYMCreate
'    With SYMCreate
'        .区分名称 = "レイアウト情報"
'        .区分桁数 = 1
'        .値 = "0"
'        .値名称 = "部品納入内示表(トヨタ形式)"
'        .数字1 = 1
'        .制御区分 = "X"
'        .作成者 = gTanto
'        If Not SYMInitCreate(True) Then
'            .値 = "1"
'            .値名称 = "納入内示ﾃﾞｰﾀB(トヨタ紡織形式)"
'            .数字1 = 1
'            Call SYMInitCreate
'        End If
'    End With

End Sub


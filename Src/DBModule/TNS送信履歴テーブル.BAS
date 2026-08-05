Attribute VB_Name = "TNS送信履歴テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TNS送信履歴テーブル
'*
'*   モジュールID  : TNS送信履歴テーブル.Bas
'*
'*                                      /項目追加含む/
'*   作成日  :  2006/02/13  By Cis      履歴モジュール追加
'*   作成日  :  2007/03/09  By Cis
'*   変更日  :  2007/12/27  契約先･送信先区分の追加
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_TNS送信履歴テーブル
        履歴管理NO            As Long
        送信先                As String
        FCODE                 As String
        契約先                As String
        送信先区分            As Byte
        枝番                  As Byte
        HTYPE                 As String
        開始日時              As Date
        終了日時              As Date
        データ件数            As Long
        送信件数              As Long
        処理コード            As Byte
        処理結果              As String
        ERR                   As String
        送信開始日時          As Date
        送信終了日時          As Date
        送信日                As String
        送信連番              As Byte
        送信管理NO            As Long
        処理連番              As Byte
        実行区分              As Byte
        自動区分              As Byte
        対象ファイル          As String
        対象ファイルUNC       As String
        ERR件数               As Long
'        作成日                As Date
'        作成端末              As String
        件数          As Long
'   項目追加(自動)
        送信先名       As String
        ファイル名称   As String
        代行送信先     As String
        代行送信先名   As String
    End Type

    Public TSR As T_TNS送信履歴テーブル

    Public TSR_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'-----------------------------------------------------  2006/02/13    Insert
    Public TSR_ErrMsg   As String
    Public TSR_RNo      As Long
    Public TSR_SEQNO    As Long
'-----------------------------------------------------  2006/02/13    Insert End
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTSR(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTSR
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With TSR
            Select Case adoCol.Name
                Case "履歴管理NO":            If Not IsNull(adoCol.Value) Then .履歴管理NO = adoCol.Value
                Case "送信先":                If Not IsNull(adoCol.Value) Then .送信先 = adoCol.Value
                Case "FCODE":                 If Not IsNull(adoCol.Value) Then .FCODE = adoCol.Value
                Case "契約先":                If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
                Case "送信先区分":            If Not IsNull(adoCol.Value) Then .送信先区分 = adoCol.Value
                Case "枝番":                  If Not IsNull(adoCol.Value) Then .枝番 = adoCol.Value
                Case "HTYPE":                 If Not IsNull(adoCol.Value) Then .HTYPE = adoCol.Value
                Case "開始日時":              If Not IsNull(adoCol.Value) Then .開始日時 = adoCol.Value
                Case "終了日時":              If Not IsNull(adoCol.Value) Then .終了日時 = adoCol.Value
                Case "データ件数":            If Not IsNull(adoCol.Value) Then .データ件数 = adoCol.Value
                Case "送信件数":              If Not IsNull(adoCol.Value) Then .送信件数 = adoCol.Value
                Case "処理コード":            If Not IsNull(adoCol.Value) Then .処理コード = adoCol.Value
                Case "処理結果":              If Not IsNull(adoCol.Value) Then .処理結果 = adoCol.Value
                Case "ERR":                   If Not IsNull(adoCol.Value) Then .ERR = adoCol.Value
                Case "送信開始日時":          If Not IsNull(adoCol.Value) Then .送信開始日時 = adoCol.Value
                Case "送信終了日時":          If Not IsNull(adoCol.Value) Then .送信終了日時 = adoCol.Value
                Case "送信日":                If Not IsNull(adoCol.Value) Then .送信日 = adoCol.Value
                Case "送信連番":              If Not IsNull(adoCol.Value) Then .送信連番 = adoCol.Value
                Case "送信管理NO":            If Not IsNull(adoCol.Value) Then .送信管理NO = adoCol.Value
                Case "処理連番":              If Not IsNull(adoCol.Value) Then .処理連番 = adoCol.Value
                Case "実行区分":              If Not IsNull(adoCol.Value) Then .実行区分 = adoCol.Value
                Case "自動区分":              If Not IsNull(adoCol.Value) Then .自動区分 = adoCol.Value
                Case "対象ファイル":          If Not IsNull(adoCol.Value) Then .対象ファイル = adoCol.Value
                Case "対象ファイルUNC":       If Not IsNull(adoCol.Value) Then .対象ファイルUNC = adoCol.Value
                Case "ERR件数":               If Not IsNull(adoCol.Value) Then .ERR件数 = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "送信先名":        If Not IsNull(adoCol.Value) Then .送信先名 = adoCol.Value
                Case "ファイル名称":    If Not IsNull(adoCol.Value) Then .ファイル名称 = adoCol.Value
                Case "代行送信先":        If Not IsNull(adoCol.Value) Then .代行送信先 = adoCol.Value
                Case "代行送信先名":        If Not IsNull(adoCol.Value) Then .代行送信先名 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TSRRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TSRRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From TNS送信履歴テーブル"
        RTxtSql = RTxtSql & " Where 履歴管理NO = " & TSR.履歴管理NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TSRRead = True
           Call ItemsFromTSR(DBNo)
           TSR_RDSTS = True
        Else
           TSR_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call TSRClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TSRReadNext(Optional DBNo As Integer = 0)
    If Not TSR_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TSR_RDSTS = False
        Else
           Call ItemsFromTSR(DBNo)
           TSR_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TSRClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TSR_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TSR )
'*----------------------------------------------*'
Public Sub ItemsClearTSR()
    With TSR
        .履歴管理NO = 0
        .送信先 = ""
        .FCODE = ""
        .契約先 = ""
        .送信先区分 = 0
        .枝番 = 0
        .HTYPE = ""
        .開始日時 = 0
        .終了日時 = 0
        .データ件数 = 0
        .送信件数 = 0
        .処理コード = 0
        .処理結果 = ""
        .ERR = ""
        .送信開始日時 = 0
        .送信終了日時 = 0
        .送信日 = ""
        .送信連番 = 0
        .送信管理NO = 0
        .処理連番 = 0
        .実行区分 = 0
        .自動区分 = 0
        .対象ファイル = ""
        .対象ファイルUNC = ""
        .ERR件数 = 0
'       .作成日 = 0
'       .作成端末 = ""
        .件数 = 0
'   項目追加(自動)
        .送信先名 = ""
        .ファイル名称 = ""
        .代行送信先 = ""
        .代行送信先名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function TSRInsert(Optional DBNo As Integer) As Boolean
    TSRInsert = True
    With TSR
        gSL_Select = "Insert Into TNS送信履歴テーブル  ( "
        gSL_Select = gSL_Select & "履歴管理NO,送信先,FCODE,"
        gSL_Select = gSL_Select & "契約先,送信先区分,"
        gSL_Select = gSL_Select & "枝番,HTYPE,開始日時,"
        gSL_Select = gSL_Select & "終了日時,データ件数,送信件数,"
        gSL_Select = gSL_Select & "処理コード,処理結果,ERR,"
        gSL_Select = gSL_Select & "送信開始日時,送信終了日時,送信日,"
        gSL_Select = gSL_Select & "送信連番,送信管理NO,処理連番,"
        gSL_Select = gSL_Select & "実行区分,自動区分,対象ファイル,"
        gSL_Select = gSL_Select & "対象ファイルUNC,ERR件数,"
        gSL_Select = gSL_Select & "作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .履歴管理NO & ","
        gSL_Select = gSL_Select & "'" & RTrim(.送信先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.FCODE) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "" & .送信先区分 & ","
        gSL_Select = gSL_Select & "" & .枝番 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.HTYPE) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.開始日時) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.終了日時) & "',"
        gSL_Select = gSL_Select & "" & .データ件数 & ","
        gSL_Select = gSL_Select & "" & .送信件数 & ","
        gSL_Select = gSL_Select & "" & .処理コード & ","
        gSL_Select = gSL_Select & "'" & RTrim(.処理結果) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.ERR) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.送信開始日時) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.送信終了日時) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.送信日) & "',"
        gSL_Select = gSL_Select & "" & .送信連番 & ","
        gSL_Select = gSL_Select & "" & .送信管理NO & ","
        gSL_Select = gSL_Select & "" & .処理連番 & ","
        gSL_Select = gSL_Select & "" & .実行区分 & ","
        gSL_Select = gSL_Select & "" & .自動区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.対象ファイル) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.対象ファイルUNC) & "',"
        gSL_Select = gSL_Select & "" & .ERR件数 & ","
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TSRInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function TSRUpdate(Optional DBNo As Integer) As Boolean
    TSRUpdate = False
    With TSR
        gSL_Select = "Update TNS送信履歴テーブル  Set "
        gSL_Select = gSL_Select & "履歴管理NO = " & .履歴管理NO & ","
        gSL_Select = gSL_Select & "送信先 = '" & RTrim(.送信先) & "',"
        gSL_Select = gSL_Select & "FCODE = '" & RTrim(.FCODE) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "送信先区分 = " & .送信先区分 & ","
        gSL_Select = gSL_Select & "枝番 = " & .枝番 & ","
        gSL_Select = gSL_Select & "HTYPE = '" & RTrim(.HTYPE) & "',"
        gSL_Select = gSL_Select & "開始日時 = '" & RTrim(.開始日時) & "',"
        gSL_Select = gSL_Select & "終了日時 = '" & RTrim(.終了日時) & "',"
        gSL_Select = gSL_Select & "データ件数 = " & .データ件数 & ","
        gSL_Select = gSL_Select & "送信件数 = " & .送信件数 & ","
        gSL_Select = gSL_Select & "処理コード = " & .処理コード & ","
        gSL_Select = gSL_Select & "処理結果 = '" & RTrim(.処理結果) & "',"
        gSL_Select = gSL_Select & "ERR = '" & RTrim(.ERR) & "',"
        gSL_Select = gSL_Select & "送信開始日時 = '" & RTrim(.送信開始日時) & "',"
        gSL_Select = gSL_Select & "送信終了日時 = '" & RTrim(.送信終了日時) & "',"
        gSL_Select = gSL_Select & "送信日 = '" & RTrim(.送信日) & "',"
        gSL_Select = gSL_Select & "送信連番 = " & .送信連番 & ","
        gSL_Select = gSL_Select & "送信管理NO = " & .送信管理NO & ","
        gSL_Select = gSL_Select & "処理連番 = " & .処理連番 & ","
        gSL_Select = gSL_Select & "実行区分 = " & .実行区分 & ","
        gSL_Select = gSL_Select & "自動区分 = " & .自動区分 & ","
        gSL_Select = gSL_Select & "対象ファイル = '" & RTrim(.対象ファイル) & "',"
        gSL_Select = gSL_Select & "対象ファイルUNC = '" & RTrim(.対象ファイルUNC) & "',"
        gSL_Select = gSL_Select & "ERR件数 = " & .ERR件数 & ""
        gSL_Select = gSL_Select & " Where 履歴管理NO = " & .履歴管理NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TSRUpdate = True
End Function
'*------------------------------------------------------------------*'
'*  処理開始(履歴作成）Isert                                         *
'*
'*  Torcd : 送信先(契約先)
'*  FCode : ファイルコード
'*------------------------------------------------------------------*'
Public Function Start_TSRInsert(Torcd As String, FCODE As String, Optional SendNo As Long = 0) As Boolean
    
    Start_TSRInsert = True
    
    Call Get_TSRNumber(Torcd, FCODE)
    
    Call ItemsClearTSR
    With TSR
        .履歴管理NO = TSR_RNo + 1
        .送信先 = Torcd
        .FCODE = FCODE
        .契約先 = Ts_KTorcd
        .送信先区分 = Ts_SendKb
        
        .枝番 = 0
        .HTYPE = ""
        .開始日時 = Now()
        .終了日時 = Now()
        .処理コード = 99
        .処理結果 = "99"
        If SendNo = 0 Then
           .送信管理NO = .履歴管理NO
        Else
           .送信管理NO = SendNo
        End If
        If Ts_Auto Then
           .自動区分 = 1
        End If
    End With
    
    If Not TSRInsert Then Exit Function
    
    Start_TSRInsert = True
    
End Function
'*------------------------------------------------------------------*'
'*
'*                   送信SEQ・履歴管理Noを求める
'*
'*------------------------------------------------------------------*'
Public Sub Get_TSRNumber(Torcd As String, FCODE As String)
    
    TSR_SEQNO = 0
    TSR_RNo = 0
    On Error GoTo Get_TSRNumber_Er:
    
'送信SEQを求める
    gSL_Select = "SELECT MAX(送信連番) 送信連番 FROM TNS送信履歴テーブル"
    gSL_Select = gSL_Select & " WHERE 送信先 = '" & Torcd & "'"
    gSL_Select = gSL_Select & "   AND FCODE = '" & FCODE & "'"
    gSL_Select = gSL_Select & "   AND Convert(char(8),開始日時,112) = '" & Format(Now(), "yyyymmdd") & "'"
    With CisDB
        .SQL = gSL_Select
        .ConnectNo = 0
        If .DBRead(1, 0) Then
            If Not IsNull(.Recordset("送信連番")) Then
                TSR_SEQNO = .Recordset("送信連番")
            End If
        End If
        .Recordset.Close
    End With
    
'送信信管理NO
    gSL_Select = "SELECT MAX(履歴管理NO) 履歴管理NO FROM TNS送信履歴テーブル"
    With CisDB
        .SQL = gSL_Select
        .ConnectNo = 0
        If .DBRead(1, 0) Then
            If Not IsNull(.Recordset("履歴管理NO")) Then
                TSR_RNo = .Recordset("履歴管理NO")
            End If
        End If
        .Recordset.Close
    End With
    
Get_TSRNumber_Er:

End Sub

'*------------------------------------------------------------------*'
'*
'*                      処　理　経　過　更　新
'*
'*------------------------------------------------------------------*'
Public Function Middle_TSRUpDate(ProcCode As String, Optional ErrMsg As String)

    Middle_TSRUpDate = False
    
    If Trim(ProcCode) <> "UPD" Then
        Select Case RTrim$(UCase(ProcCode))
            Case "00"
                TSR_ErrMsg = "正常終了"
                If TSR.処理結果 > "10" Then
                   TSR.処理結果 = "00"
                End If
            Case Else
                TSR.処理結果 = ProcCode
                If Trim(ErrMsg) = "" Or Trim(ProcCode) = "" Then
                   TSR_ErrMsg = "その他エラー"
                   ProcCode = "99"
                Else
                   TSR_ErrMsg = ErrMsg
                   ProcCode = ProcCode
                End If
        End Select
        
        Call ItemsClearSYM
        With SYM
            .区分名称 = "送受信処理結果"
            .区分桁数 = 2
            .値 = ProcCode
            .値名称 = TSR_ErrMsg
            .制御区分 = "X"
        End With
        Call SYMReadChk
        If TSR.処理結果 <> "00" Then
           TSR.処理結果 = ProcCode
        End If
    End If
    
    TSR.終了日時 = Now()
    If TSR.処理結果 > "10" Then
       TSR.処理コード = 1
    Else
       TSR.処理コード = 0
    End If
    Call TSRUpdate

    Middle_TSRUpDate = True
End Function

Attribute VB_Name = "TEXTERRMSG"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TEXTERRMSG
'*
'*   モジュールID  : TEXTERRMSG.Bas
'*
'*   作成日  :  2003/12/11  By Cis エラー内容追加処理 ﾓｼﾞｭｰﾙ追加
'*   変更日  :  2008/09/08  By CIS 千代田工業カスタマイズ
'*---------------------------------------------------------------------------------
    Type T_TEXTERRMSG
        ERRCODE   As String
        ERRLEBEL  As String
        ERRMSG1   As String
        ERRMSG2   As String
        件数          As Long
    End Type

    Public TER As T_TEXTERRMSG

    Public TER_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTER(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTER
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With TER
            Select Case adoCol.Name
                Case "ERRCODE":   If Not IsNull(adoCol.Value) Then .ERRCODE = adoCol.Value
                Case "ERRLEBEL":  If Not IsNull(adoCol.Value) Then .ERRLEBEL = adoCol.Value
                Case "ERRMSG1":   If Not IsNull(adoCol.Value) Then .ERRMSG1 = adoCol.Value
                Case "ERRMSG2":   If Not IsNull(adoCol.Value) Then .ERRMSG2 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TERRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TERRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From TEXTERRMSG"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TERRead = True
           Call ItemsFromTER(DBNo)
           TER_RDSTS = True
        Else
           TER_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call TERClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TERReadNext(Optional DBNo As Integer = 0)
    If Not TER_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TER_RDSTS = False
        Else
           Call ItemsFromTER(DBNo)
           TER_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TERClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TER_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TER )
'*----------------------------------------------*'
Public Sub ItemsClearTER()
    With TER
        .ERRCODE = ""
        .ERRLEBEL = ""
        .ERRMSG1 = ""
        .ERRMSG2 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Sub TERInsert(Optional DBNo As Integer)
    With TER
        gSL_Select = "Insert Into TEXTERRMSG  ( "
        gSL_Select = gSL_Select & "ERRCODE,ERRLEBEL,ERRMSG1,"
        gSL_Select = gSL_Select & "ERRMSG2"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.ERRCODE) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.ERRLEBEL) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.ERRMSG1) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.ERRMSG2) & "'"
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
Public Sub TERUpdate(Optional DBNo As Integer)
    With TER
        gSL_Select = "Update TEXTERRMSG  Set "
        gSL_Select = gSL_Select & "ERRLEBEL = '" & RTrim(.ERRLEBEL) & "',"
        gSL_Select = gSL_Select & "ERRMSG1 = '" & RTrim(.ERRMSG1) & "',"
        gSL_Select = gSL_Select & "ERRMSG2 = '" & RTrim(.ERRMSG2) & "'"
        gSL_Select = gSL_Select & " Where ERRCODE = '" & RTrim(.ERRCODE) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        .DBExec
    End With

End Sub
'*--------------------------*'
'*   エラー内容追加処理      *
'*--------------------------*'
Public Sub TERPut(ErrCD As String, Optional ErrStr1 As String, Optional ErrStr2 As String)
    Dim wPutFlg     As Boolean
    
    gSL_Select = "SELECT * FROM TEXTERRMSG "
    gSL_Select = gSL_Select & " Where ERRCODE = '" & RTrim(ErrCD) & "'"
    If TERRead(gSL_Select, 1, 1) Then
       wPutFlg = False
    Else
       wPutFlg = True
       Call ItemsClearTER
       TER.ERRCODE = RTrim(ErrCD)
    End If
    
    With TER
        .ERRMSG1 = "": .ERRMSG2 = ""
        If Trim(ErrStr1) <> "" Then
           .ERRMSG1 = CisFun.Mid2(ErrStr1, 1, 100): .ERRMSG2 = CisFun.Mid2(ErrStr2, 1, 100)
        Else
            Select Case ErrCD
                Case "H1"
                   .ERRMSG1 = "     テキストデータファイルが異常です。     "
                   .ERRMSG2 = "           (長さ０バイト)              "
                Case "H2"
                   .ERRMSG1 = "      変換元テキストファイル未登録(変換処理)     "
                   .ERRMSG2 = "  テキストファイル設定を確認後、再度処理を行って下さい！      "
                Case "H3"
                   .ERRMSG1 = "     ファイルコード未設定（処理異常）       "
                   .ERRMSG2 = "  テキストファイル設定を確認後、再度処理を行って下さい！！      "
                Case "H4"
                   .ERRMSG1 = "     テキスト変換設定未登録（処理異常）       "
                   .ERRMSG2 = "  テキスト変換設定を確認後、再度処理を行って下さい！！      "
                Case "H5"
                   .ERRMSG1 = "     テキストファイルが異常です。(改行なし)            "
                   .ERRMSG2 = "  テキスト変換設定を確認後、再度処理を行って下さい！！      "
                Case "N1"
                   .ERRMSG1 = "     内示データ更新異常。            "
                   .ERRMSG2 = "  "
            End Select
        End If
    End With
    If wPutFlg Then
       Call TERInsert(1)
    Else
       Call TERUpdate(1)
    End If
End Sub
'*--------------------------*'
'*   エラー内容追加処理      *
'*--------------------------*'
Public Sub TERPutSak(ErrCD As String, Optional ErrStr1 As String, Optional ErrStr2 As String)
    Dim wPutFlg     As Boolean
    
    gSL_Select = "SELECT * FROM TEXTERRMSG "
    gSL_Select = gSL_Select & " Where ERRCODE = '" & RTrim(ErrCD) & "'"
    If TERRead(gSL_Select, 1, 1) Then
       wPutFlg = False
    Else
       wPutFlg = True
       Call ItemsClearTER
       TER.ERRCODE = RTrim(ErrCD)
    End If
    
    With TER
        .ERRMSG1 = "": .ERRMSG2 = ""
        If Trim(ErrStr1) <> "" Then
           .ERRMSG1 = CisFun.Mid2(ErrStr1, 1, 100): .ERRMSG2 = CisFun.Mid2(ErrStr2, 1, 100)
        Else
            Select Case ErrCD
                Case "H1"
                   .ERRMSG1 = "     テキストデータファイルが異常です。     "
                   .ERRMSG2 = "           (長さ０バイト)              "
                Case "H2"
                   .ERRMSG1 = "      変換元テキストファイル未登録(変換処理)     "
                   .ERRMSG2 = "  テキストファイル設定を確認後、再度処理を行って下さい！      "
                Case "H3"
                   .ERRMSG1 = "     ファイルコード未設定（処理異常）       "
                   .ERRMSG2 = "  テキストファイル設定を確認後、再度処理を行って下さい！！      "
                Case "H4"
                   .ERRMSG1 = "     テキスト変換設定未登録（処理異常）       "
                   .ERRMSG2 = "  テキスト変換設定を確認後、再度処理を行って下さい！！      "
                Case "S1"
                   .ERRMSG1 = "     処理中止(データ変換)       "
                   .ERRMSG2 = "       "
                Case "S2"
                   .ERRMSG1 = "     処理中止(テキスト変換）       "
                   .ERRMSG2 = "       "
                Case "S3"
                   .ERRMSG1 = "          テキスト変換エラー             "
                   .ERRMSG2 = "       内示テキストワーク出力エラー      "
                Case "S4"
                   .ERRMSG1 = "          データ変換エラー             "
                   .ERRMSG2 = "       内示TEXTテーブル出力エラー      "
                Case "J1"
                   .ERRMSG1 = "          データ変換エラー             "
                   .ERRMSG2 = "       　受注情報出力エラー            "
                Case "J2"
                   .ERRMSG1 = "     処理中止(データ変換)       "
                   .ERRMSG2 = "       "
            End Select
        End If
    End With
    If wPutFlg Then
       Call TERInsert(1)
    Else
       Call TERUpdate(1)
    End If
End Sub







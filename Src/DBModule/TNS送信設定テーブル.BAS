Attribute VB_Name = "TNS送信設定テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TNS送信設定テーブル
'*
'*   モジュールID  : TNS送信設定テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*                                      /項目追加含む/
'*   作成日  :  2006/02/10  By Cis
'*   作成日  :  2007/02/08  By Cis
'*   作成日  :  2007/03/08  By Cis
'*   作成日  :  2007/03/09  By Cis
'*   作成日  :  2007/12/26  By Cis  代行送信先･送信先区分･送信開始日･受入CHK を追加
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_TNS送信設定テーブル
        送信先                    As String
        FCODE                     As String
        代行送信先                As String
        送信先区分                As Byte
        枝番                      As Byte
        送信                      As Byte
        TNS宛先                   As String
        宛先                      As String
        発信元                    As String
        ファイル名称              As String
        処理区分                  As String
        削除日数                  As Byte
        識別子                    As String
        TNSタイプ                 As Byte
        TNSフォーマット           As Byte
        TNS送信ファイル名         As String
        レコード長                As Long
        作成先区分1               As Byte
        作成先区分2               As Byte
        ゼロ件データ              As Byte
        出力区分                  As Byte
        送信開始日                As String
        TNSシステム起動           As Byte
        品番CHK                   As Byte
        受入CHK                   As Byte
        品番ERRLIST               As Byte
'        作成日                    As Date
'        作成端末                  As String
'        更新日                    As Date
'        更新端末                  As String
        件数          As Long
'   項目追加(自動)
        送信先名       As String
        代行送信先名   As String
    End Type

    Public TSS As T_TNS送信設定テーブル

    Public TSS_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTSS(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTSS
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With TSS
            Select Case adoCol.Name
                Case "送信先":                    If Not IsNull(adoCol.Value) Then .送信先 = adoCol.Value
                Case "FCODE":                     If Not IsNull(adoCol.Value) Then .FCODE = adoCol.Value
                Case "代行送信先":                If Not IsNull(adoCol.Value) Then .代行送信先 = adoCol.Value
                Case "送信先区分":                If Not IsNull(adoCol.Value) Then .送信先区分 = adoCol.Value
                Case "枝番":                      If Not IsNull(adoCol.Value) Then .枝番 = adoCol.Value
                Case "送信":                      If Not IsNull(adoCol.Value) Then .送信 = adoCol.Value
                Case "TNS宛先":                   If Not IsNull(adoCol.Value) Then .TNS宛先 = adoCol.Value
                Case "宛先":                      If Not IsNull(adoCol.Value) Then .宛先 = adoCol.Value
                Case "発信元":                    If Not IsNull(adoCol.Value) Then .発信元 = adoCol.Value
                Case "ファイル名称":              If Not IsNull(adoCol.Value) Then .ファイル名称 = adoCol.Value
                Case "処理区分":                  If Not IsNull(adoCol.Value) Then .処理区分 = adoCol.Value
                Case "削除日数":                  If Not IsNull(adoCol.Value) Then .削除日数 = adoCol.Value
                Case "識別子":                    If Not IsNull(adoCol.Value) Then .識別子 = adoCol.Value
                Case "TNSタイプ":                 If Not IsNull(adoCol.Value) Then .TNSタイプ = adoCol.Value
                Case "TNSフォーマット":           If Not IsNull(adoCol.Value) Then .TNSフォーマット = adoCol.Value
                Case "TNS送信ファイル名":         If Not IsNull(adoCol.Value) Then .TNS送信ファイル名 = adoCol.Value
                Case "レコード長":                If Not IsNull(adoCol.Value) Then .レコード長 = adoCol.Value
                Case "作成先区分1":               If Not IsNull(adoCol.Value) Then .作成先区分1 = adoCol.Value
                Case "作成先区分2":               If Not IsNull(adoCol.Value) Then .作成先区分2 = adoCol.Value
                Case "ゼロ件データ":              If Not IsNull(adoCol.Value) Then .ゼロ件データ = adoCol.Value
                Case "出力区分":                  If Not IsNull(adoCol.Value) Then .出力区分 = adoCol.Value
                Case "送信開始日":                If Not IsNull(adoCol.Value) Then .送信開始日 = adoCol.Value
                Case "TNSシステム起動":           If Not IsNull(adoCol.Value) Then .TNSシステム起動 = adoCol.Value
                Case "品番CHK":                   If Not IsNull(adoCol.Value) Then .品番CHK = adoCol.Value
                Case "受入CHK":                   If Not IsNull(adoCol.Value) Then .受入CHK = adoCol.Value
                Case "品番ERRLIST":               If Not IsNull(adoCol.Value) Then .品番ERRLIST = adoCol.Value
'                Case "作成日":                    If Not IsNull(adoCol.Value) Then .作成日                   = adoCol.Value
'                Case "作成端末":                  If Not IsNull(adoCol.Value) Then .作成端末                 = adoCol.Value
'                Case "更新日":                    If Not IsNull(adoCol.Value) Then .更新日                   = adoCol.Value
'                Case "更新端末":                  If Not IsNull(adoCol.Value) Then .更新端末                 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "送信先名":        If Not IsNull(adoCol.Value) Then .送信先名 = adoCol.Value
                Case "代行送信先名":    If Not IsNull(adoCol.Value) Then .代行送信先名 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TSSRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TSSRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From TNS送信設定テーブル"
        RTxtSql = RTxtSql & " Where 送信先 = '" & RTrim(TSS.送信先) & "'"
        RTxtSql = RTxtSql & " And   FCODE = '" & RTrim(TSS.FCODE) & "'"
        RTxtSql = RTxtSql & " And   代行送信先 = '" & RTrim(TSS.代行送信先) & "'"
        RTxtSql = RTxtSql & " And   送信先区分 = " & TSS.送信先区分 & ""
        RTxtSql = RTxtSql & " And   枝番 = " & TSS.枝番 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TSSRead = True
           Call ItemsFromTSS(DBNo)
           TSS_RDSTS = True
        Else
           TSS_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call TSSClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TSSReadNext(Optional DBNo As Integer = 0)
    If Not TSS_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TSS_RDSTS = False
        Else
           Call ItemsFromTSS(DBNo)
           TSS_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TSSClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TSS_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TSS )
'*----------------------------------------------*'
Public Sub ItemsClearTSS()
    With TSS
        .送信先 = ""
        .FCODE = ""
        .代行送信先 = ""
        .送信先区分 = 0
        .枝番 = 0
        .送信 = 0
        .TNS宛先 = ""
        .宛先 = ""
        .発信元 = ""
        .ファイル名称 = ""
        .処理区分 = ""
        .削除日数 = 0
        .識別子 = ""
        .TNSタイプ = 0
        .TNSフォーマット = 0
        .TNS送信ファイル名 = ""
        .レコード長 = 0
        .作成先区分1 = 0
        .作成先区分2 = 0
        .ゼロ件データ = 0
        .出力区分 = 0
        .送信開始日 = ""
        .TNSシステム起動 = 0
        .品番CHK = 0
        .受入CHK = 0
        .品番ERRLIST = 0
'       .作成日 = 0
'       .作成端末 = ""
'       .更新日 = 0
'       .更新端末 = ""
        .件数 = 0
'   項目追加(自動)
        .送信先名 = ""
        .代行送信先名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function TSSInsert(Optional DBNo As Integer) As Boolean
    TSSInsert = False
    With TSS
        gSL_Select = "Insert Into TNS送信設定テーブル  ( "
        gSL_Select = gSL_Select & "送信先,FCODE,代行送信先,"
        gSL_Select = gSL_Select & "送信先区分,枝番,送信,"
        gSL_Select = gSL_Select & "TNS宛先,宛先,発信元,"
        gSL_Select = gSL_Select & "ファイル名称,処理区分,削除日数,"
        gSL_Select = gSL_Select & "識別子,TNSタイプ,TNSフォーマット,"
        gSL_Select = gSL_Select & "TNS送信ファイル名,レコード長,作成先区分1,"
        gSL_Select = gSL_Select & "作成先区分2,ゼロ件データ,出力区分,"
        gSL_Select = gSL_Select & "送信開始日,TNSシステム起動,品番CHK,"
        gSL_Select = gSL_Select & "受入CHK,品番ERRLIST,"
        gSL_Select = gSL_Select & "作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.送信先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.FCODE) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.代行送信先) & "',"
        gSL_Select = gSL_Select & "" & .送信先区分 & ","
        gSL_Select = gSL_Select & "" & .枝番 & ","
        gSL_Select = gSL_Select & "" & .送信 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.TNS宛先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.宛先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発信元) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.ファイル名称) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.処理区分) & "',"
        gSL_Select = gSL_Select & "" & .削除日数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.識別子) & "',"
        gSL_Select = gSL_Select & "" & .TNSタイプ & ","
        gSL_Select = gSL_Select & "" & .TNSフォーマット & ","
        gSL_Select = gSL_Select & "'" & RTrim(.TNS送信ファイル名) & "',"
        gSL_Select = gSL_Select & "" & .レコード長 & ","
        gSL_Select = gSL_Select & "" & .作成先区分1 & ","
        gSL_Select = gSL_Select & "" & .作成先区分2 & ","
        gSL_Select = gSL_Select & "" & .ゼロ件データ & ","
        gSL_Select = gSL_Select & "" & .出力区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.送信開始日) & "',"
        gSL_Select = gSL_Select & "" & .TNSシステム起動 & ","
        gSL_Select = gSL_Select & "" & .品番CHK & ","
        gSL_Select = gSL_Select & "" & .受入CHK & ","
        gSL_Select = gSL_Select & "" & .品番ERRLIST & ","
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TSSInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function TSSUpdate(Optional DBNo As Integer) As Boolean
    TSSUpdate = False
    With TSS
        gSL_Select = "Update TNS送信設定テーブル  Set "
        gSL_Select = gSL_Select & "送信先 = '" & RTrim(.送信先) & "',"
        gSL_Select = gSL_Select & "FCODE = '" & RTrim(.FCODE) & "',"
        gSL_Select = gSL_Select & "代行送信先 = '" & RTrim(.代行送信先) & "',"
        gSL_Select = gSL_Select & "送信先区分 = " & .送信先区分 & ","
        gSL_Select = gSL_Select & "枝番 = " & .枝番 & ","
        gSL_Select = gSL_Select & "送信 = " & .送信 & ","
        gSL_Select = gSL_Select & "TNS宛先 = '" & RTrim(.TNS宛先) & "',"
        gSL_Select = gSL_Select & "宛先 = '" & RTrim(.宛先) & "',"
        gSL_Select = gSL_Select & "発信元 = '" & RTrim(.発信元) & "',"
        gSL_Select = gSL_Select & "ファイル名称 = '" & RTrim(.ファイル名称) & "',"
        gSL_Select = gSL_Select & "処理区分 = '" & RTrim(.処理区分) & "',"
        gSL_Select = gSL_Select & "削除日数 = " & .削除日数 & ","
        gSL_Select = gSL_Select & "識別子 = '" & RTrim(.識別子) & "',"
        gSL_Select = gSL_Select & "TNSタイプ = " & .TNSタイプ & ","
        gSL_Select = gSL_Select & "TNSフォーマット = " & .TNSフォーマット & ","
        gSL_Select = gSL_Select & "TNS送信ファイル名 = '" & RTrim(.TNS送信ファイル名) & "',"
        gSL_Select = gSL_Select & "レコード長 = " & .レコード長 & ","
        gSL_Select = gSL_Select & "作成先区分1 = " & .作成先区分1 & ","
        gSL_Select = gSL_Select & "作成先区分2 = " & .作成先区分2 & ","
        gSL_Select = gSL_Select & "ゼロ件データ = " & .ゼロ件データ & ","
        gSL_Select = gSL_Select & "出力区分 = " & .出力区分 & ","
        gSL_Select = gSL_Select & "送信開始日 = '" & RTrim(.送信開始日) & "',"
        gSL_Select = gSL_Select & "TNSシステム起動 = " & .TNSシステム起動 & ","
        gSL_Select = gSL_Select & "品番CHK = " & .品番CHK & ","
        gSL_Select = gSL_Select & "受入CHK = " & .受入CHK & ","
        gSL_Select = gSL_Select & "品番ERRLIST = " & .品番ERRLIST & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,15)"
        gSL_Select = gSL_Select & " Where 送信先 = '" & RTrim(.送信先) & "'"
        gSL_Select = gSL_Select & "   And FCODE = '" & RTrim(.FCODE) & "'"
        gSL_Select = gSL_Select & "   And 代行送信先 = '" & RTrim(.代行送信先) & "'"
        gSL_Select = gSL_Select & "   And 送信先区分 = " & .送信先区分 & ""
        gSL_Select = gSL_Select & "   And 枝番 = " & .枝番 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TSSUpdate = True
End Function

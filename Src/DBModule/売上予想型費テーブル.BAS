Attribute VB_Name = "売上予想型費テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   売上予想型費テーブル
'*
'*   モジュールID  : 売上予想型費テーブル.Bas
'*
'*
'*                                      /項目追加含む/
'*   作成日  :  2004/11/06  By Cis
'*   変更日  :  2004/11/06  金額の配列化
'*---------------------------------------------------------------------------------
    Type T_売上予想型費テーブル
        年度      As String
        予想年月  As String
        契約先    As String
        品番      As String
        開始年月  As String
        車型      As String
        工区      As String
        金額(14)    As Currency
        内示位置  As Byte
        集約品番  As String
        品番編集  As String
'        作成日    As Date
'        作成端末  As String
'        更新日    As Date
'        更新端末  As String
        件数          As Long
'   項目追加(自動)
        契約先名      As String
        上期計        As Currency
        下期計        As Currency
        期合計        As Currency
    End Type

    Public UYK As T_売上予想型費テーブル

    Public UYK_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromUYK(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearUYK
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With UYK
            Select Case adoCol.Name
                Case "年度":      If Not IsNull(adoCol.Value) Then .年度 = adoCol.Value
                Case "予想年月":  If Not IsNull(adoCol.Value) Then .予想年月 = adoCol.Value
                Case "契約先":    If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
                Case "品番":      If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "開始年月":  If Not IsNull(adoCol.Value) Then .開始年月 = adoCol.Value
                Case "車型":      If Not IsNull(adoCol.Value) Then .車型 = adoCol.Value
                Case "工区":      If Not IsNull(adoCol.Value) Then .工区 = adoCol.Value
                Case "内示位置":  If Not IsNull(adoCol.Value) Then .内示位置 = adoCol.Value
                Case "集約品番":  If Not IsNull(adoCol.Value) Then .集約品番 = adoCol.Value
                Case "品番編集":  If Not IsNull(adoCol.Value) Then .品番編集 = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
'                Case "作成端末":  If Not IsNull(adoCol.Value) Then .作成端末 = adoCol.Value
'                Case "更新日":    If Not IsNull(adoCol.Value) Then .更新日   = adoCol.Value
'                Case "更新端末":  If Not IsNull(adoCol.Value) Then .更新端末 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "契約先名":        If Not IsNull(adoCol.Value) Then .契約先名 = adoCol.Value
                Case "上期計":  If Not IsNull(adoCol.Value) Then .上期計 = adoCol.Value
                Case "下期計":  If Not IsNull(adoCol.Value) Then .下期計 = adoCol.Value
                Case "期合計":  If Not IsNull(adoCol.Value) Then .期合計 = adoCol.Value
                Case Else
                    If Not IsNull(adoCol.Value) Then
                        If adoCol.Name Like "金額*" Then
                            .金額(CInt(Mid(adoCol.Name, 3, 2))) = adoCol.Value
                        End If
                    End If
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function UYKRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    UYKRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 売上予想型費テーブル"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           UYKRead = True
           Call ItemsFromUYK(DBNo)
           UYK_RDSTS = True
        Else
           UYK_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call UYKClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub UYKReadNext(Optional DBNo As Integer = 0)
    If Not UYK_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           UYK_RDSTS = False
        Else
           Call ItemsFromUYK(DBNo)
           UYK_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub UYKClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        UYK_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( UYK )
'*----------------------------------------------*'
Public Sub ItemsClearUYK()
    With UYK
        .年度 = ""
        .予想年月 = ""
        .契約先 = ""
        .品番 = ""
        .開始年月 = ""
        .車型 = ""
        .工区 = ""
        For gInt = 1 To 14
            .金額(gInt) = 0
        Next
        .内示位置 = 0
        .集約品番 = ""
        .品番編集 = ""
'       .作成日 = 0
'       .作成端末 = ""
'       .更新日 = 0
'       .更新端末 = ""
        .件数 = 0
'   項目追加(自動)
        .契約先名 = ""
        .上期計 = 0
        .下期計 = 0
        .期合計 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function UYKInsert(Optional DBNo As Integer) As Boolean
    UYKInsert = True
    With UYK
        gSL_Select = "Insert Into 売上予想型費テーブル  ( "
        gSL_Select = gSL_Select & "年度,予想年月,契約先,"
        gSL_Select = gSL_Select & "品番,開始年月,車型,"
        gSL_Select = gSL_Select & "工区,"
        For gInt = 1 To 14
            gSL_Select = gSL_Select & "金額" & gInt & ","
        Next
        gSL_Select = gSL_Select & "内示位置,集約品番,品番編集,"
        gSL_Select = gSL_Select & "作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.年度) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.予想年月) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.開始年月) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.車型) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.工区) & "',"
        For gInt = 1 To 14
            gSL_Select = gSL_Select & "" & .金額(gInt) & ","
        Next
        gSL_Select = gSL_Select & "" & .内示位置 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.集約品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品番編集) & "',"
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,15)"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    UYKInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function UYKUpdate(Optional DBNo As Integer) As Boolean
    UYKUpdate = False
    With UYK
        gSL_Select = "Update 売上予想型費テーブル  Set "
        gSL_Select = gSL_Select & "年度 = '" & RTrim(.年度) & "',"
        gSL_Select = gSL_Select & "予想年月 = '" & RTrim(.予想年月) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "開始年月 = '" & RTrim(.開始年月) & "',"
        gSL_Select = gSL_Select & "車型 = '" & RTrim(.車型) & "',"
        gSL_Select = gSL_Select & "工区 = '" & RTrim(.工区) & "',"
        For gInt = 1 To 14
            gSL_Select = gSL_Select & "金額" & gInt & " = " & .金額(gInt) & ","
        Next
        gSL_Select = gSL_Select & "内示位置 = " & .内示位置 & ","
        gSL_Select = gSL_Select & "集約品番 = '" & RTrim(.集約品番) & "',"
        gSL_Select = gSL_Select & "品番編集 = '" & RTrim(.品番編集) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,15)"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    UYKUpdate = True
End Function

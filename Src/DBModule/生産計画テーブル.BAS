Attribute VB_Name = "生産計画テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   生産計画テーブル
'*
'*   モジュールID  : 生産計画テーブル.Bas
'*
'*   作成日  :  2004/02/19  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_生産計画テーブル
        内示年月  As String
        品番      As String
        手配先    As String
        受入      As String
        工区      As String
        納期      As String
        仕掛日    As String
        手配日    As String
        数量      As Long
        処理区分  As Byte
        削除区分  As Byte
        変更回数  As Byte
        計画年月    As String
        作成者    As String
'        作成日    As Date
'        作成端末  As String
        更新者    As String
'        更新日    As Date
'        更新端末  As String
        件数          As Long
    
        発注比率      As Byte
        手配先名      As String
        表示品番      As String
    
'[ 変更一覧 ]
        増減率          As Integer
        今回計画数      As Long
        前回計画数      As Long
        計画差          As Long
        今回内示数      As Long
        前回内示数      As Long
        内示差          As Long
    End Type

    Public SKK As T_生産計画テーブル

    Public SKK_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromSKK(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearSKK
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With SKK
            Select Case adoCol.Name
                Case "内示年月":  If Not IsNull(adoCol.Value) Then .内示年月 = adoCol.Value
                Case "品番":      If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "手配先":    If Not IsNull(adoCol.Value) Then .手配先 = adoCol.Value
                Case "受入":      If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "工区":      If Not IsNull(adoCol.Value) Then .工区 = adoCol.Value
                Case "納期":      If Not IsNull(adoCol.Value) Then .納期 = adoCol.Value
                Case "仕掛日":    If Not IsNull(adoCol.Value) Then .仕掛日 = adoCol.Value
                Case "手配日":    If Not IsNull(adoCol.Value) Then .手配日 = adoCol.Value
                Case "数量":      If Not IsNull(adoCol.Value) Then .数量 = adoCol.Value
                Case "処理区分":  If Not IsNull(adoCol.Value) Then .処理区分 = adoCol.Value
                Case "削除区分":  If Not IsNull(adoCol.Value) Then .削除区分 = adoCol.Value
                Case "変更回数":  If Not IsNull(adoCol.Value) Then .変更回数 = adoCol.Value
                Case "計画年月":  If Not IsNull(adoCol.Value) Then .計画年月 = adoCol.Value
                Case "作成者":    If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
'                Case "作成端末":  If Not IsNull(adoCol.Value) Then .作成端末 = adoCol.Value
                Case "更新者":    If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新日":    If Not IsNull(adoCol.Value) Then .更新日   = adoCol.Value
'                Case "更新端末":  If Not IsNull(adoCol.Value) Then .更新端末 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
            
                Case "発注比率":         If Not IsNull(adoCol.Value) Then .発注比率 = adoCol.Value
                Case "手配先名":         If Not IsNull(adoCol.Value) Then .手配先名 = adoCol.Value
                Case "表示品番":         If Not IsNull(adoCol.Value) Then .表示品番 = adoCol.Value
            
                Case "増減率":           If Not IsNull(adoCol.Value) Then .増減率 = adoCol.Value
                Case "今回計画数":       If Not IsNull(adoCol.Value) Then .今回計画数 = adoCol.Value
                Case "前回計画数":       If Not IsNull(adoCol.Value) Then .前回計画数 = adoCol.Value
                Case "計画差":           If Not IsNull(adoCol.Value) Then .計画差 = adoCol.Value
                Case "今回内示数":       If Not IsNull(adoCol.Value) Then .今回内示数 = adoCol.Value
                Case "前回内示数":       If Not IsNull(adoCol.Value) Then .前回内示数 = adoCol.Value
                Case "内示差":           If Not IsNull(adoCol.Value) Then .内示差 = adoCol.Value
            
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function SKKRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    SKKRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 生産計画テーブル"
        RTxtSql = RTxtSql & " Where 内示年月 = '" & RTrim(SKK.内示年月) & "'"
        RTxtSql = RTxtSql & " And   品番 = '" & RTrim(SKK.品番) & "'"
        RTxtSql = RTxtSql & " And   手配先 = '" & RTrim(SKK.手配先) & "'"
        RTxtSql = RTxtSql & " And   受入 = '" & RTrim(SKK.受入) & "'"
        RTxtSql = RTxtSql & " And   納期 = '" & RTrim(SKK.納期) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           SKKRead = True
           Call ItemsFromSKK(DBNo)
           SKK_RDSTS = True
        Else
           SKK_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call SKKClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub SKKReadNext(Optional DBNo As Integer = 0)
    If Not SKK_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           SKK_RDSTS = False
        Else
           Call ItemsFromSKK(DBNo)
           SKK_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub SKKClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        SKK_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( SKK )
'*----------------------------------------------*'
Public Sub ItemsClearSKK()
    With SKK
        .内示年月 = ""
        .品番 = ""
        .手配先 = ""
        .受入 = ""
        .工区 = ""
        .納期 = ""
        .仕掛日 = ""
        .手配日 = ""
        .数量 = 0
        .処理区分 = 0
        .削除区分 = 0
        .変更回数 = 0
        .計画年月 = ""
        .作成者 = ""
'       .作成日 = 0
'       .作成端末 = ""
        .更新者 = ""
'       .更新日 = 0
'       .更新端末 = ""
        .件数 = 0
    
        .発注比率 = 0
        .手配先名 = ""
    
        .表示品番 = ""
        .増減率 = 0
        .今回計画数 = 0
        .前回計画数 = 0
        .計画差 = 0
        .今回内示数 = 0
        .前回内示数 = 0
        .内示差 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function SKKInsert(Optional DBNo As Integer) As Boolean
    SKKInsert = True
    With SKK
        gSL_Select = "Insert Into 生産計画テーブル  ( "
        gSL_Select = gSL_Select & "内示年月,品番,手配先,"
        gSL_Select = gSL_Select & "受入,工区,納期,"
        gSL_Select = gSL_Select & "仕掛日,手配日,数量,"
        gSL_Select = gSL_Select & "処理区分,削除区分,変更回数,"
        gSL_Select = gSL_Select & "計画年月,"
        gSL_Select = gSL_Select & "作成者,作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.内示年月) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.手配先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.工区) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納期) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.仕掛日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.手配日) & "',"
        gSL_Select = gSL_Select & "" & .数量 & ","
        gSL_Select = gSL_Select & "" & .処理区分 & ","
        gSL_Select = gSL_Select & "" & .削除区分 & ","
        gSL_Select = gSL_Select & "" & .変更回数 & ","
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
    SKKInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function SKKUpdate(Optional DBNo As Integer) As Boolean
    SKKUpdate = False
    With SKK
        gSL_Select = "Update 生産計画テーブル  Set "
        gSL_Select = gSL_Select & "内示年月 = '" & RTrim(.内示年月) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "手配先 = '" & RTrim(.手配先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "工区 = '" & RTrim(.工区) & "',"
        gSL_Select = gSL_Select & "納期 = '" & RTrim(.納期) & "',"
        gSL_Select = gSL_Select & "仕掛日 = '" & RTrim(.仕掛日) & "',"
        gSL_Select = gSL_Select & "手配日 = '" & RTrim(.手配日) & "',"
        gSL_Select = gSL_Select & "数量 = " & .数量 & ","
        gSL_Select = gSL_Select & "処理区分 = " & .処理区分 & ","
        gSL_Select = gSL_Select & "削除区分 = " & .削除区分 & ","
        gSL_Select = gSL_Select & "変更回数 = " & .変更回数 & ","
        gSL_Select = gSL_Select & "計画年月 = '" & RTrim(.計画年月) & "',"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 内示年月 = '" & RTrim(.内示年月) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 手配先 = '" & RTrim(.手配先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 納期 = '" & RTrim(.納期) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SKKUpdate = True
End Function

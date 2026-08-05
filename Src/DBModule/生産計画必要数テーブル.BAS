Attribute VB_Name = "生産計画必要数テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   生産計画必要数テーブル
'*
'*   モジュールID  : 生産計画必要数テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*
'*   作成日  :  2004/03/01  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_生産計画必要数テーブル
        内示年月          As String
        品番              As String
        手配先            As String
        受入              As String
        工区              As String
        納期              As String
        数量              As Long
        分配前数量        As Long
        手配区分          As String
        発注比率          As Byte
        発注比率計        As Integer
        件数              As Byte
        計画年月          As String
        作成者            As String
'        作成日            As Date
'        作成端末          As String
    End Type

    Public SHH As T_生産計画必要数テーブル

    Public SHH_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromSHH(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearSHH
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With SHH
            Select Case adoCol.Name
                Case "内示年月":          If Not IsNull(adoCol.Value) Then .内示年月 = adoCol.Value
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "手配先":            If Not IsNull(adoCol.Value) Then .手配先 = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "工区":              If Not IsNull(adoCol.Value) Then .工区 = adoCol.Value
                Case "納期":              If Not IsNull(adoCol.Value) Then .納期 = adoCol.Value
                Case "数量":              If Not IsNull(adoCol.Value) Then .数量 = adoCol.Value
                Case "分配前数量":        If Not IsNull(adoCol.Value) Then .分配前数量 = adoCol.Value
                Case "手配区分":          If Not IsNull(adoCol.Value) Then .手配区分 = adoCol.Value
                Case "発注比率":          If Not IsNull(adoCol.Value) Then .発注比率 = adoCol.Value
                Case "発注比率計":        If Not IsNull(adoCol.Value) Then .発注比率計 = adoCol.Value
                Case "件数":              If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
                Case "計画年月":          If Not IsNull(adoCol.Value) Then .計画年月 = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function SHHRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    SHHRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 生産計画必要数テーブル"
        RTxtSql = RTxtSql & " Where 内示年月 = '" & RTrim(SHH.内示年月) & "'"
        RTxtSql = RTxtSql & " And   品番 = '" & RTrim(SHH.品番) & "'"
        RTxtSql = RTxtSql & " And   手配先 = '" & RTrim(SHH.手配先) & "'"
        RTxtSql = RTxtSql & " And   受入 = '" & RTrim(SHH.受入) & "'"
        RTxtSql = RTxtSql & " And   納期 = '" & RTrim(SHH.納期) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           SHHRead = True
           Call ItemsFromSHH(DBNo)
           SHH_RDSTS = True
        Else
           SHH_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call SHHClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub SHHReadNext(Optional DBNo As Integer = 0)
    If Not SHH_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           SHH_RDSTS = False
        Else
           Call ItemsFromSHH(DBNo)
           SHH_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub SHHClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        SHH_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( SHH )
'*----------------------------------------------*'
Public Sub ItemsClearSHH()
    With SHH
        .内示年月 = ""
        .品番 = ""
        .手配先 = ""
        .受入 = ""
        .工区 = ""
        .納期 = ""
        .数量 = 0
        .分配前数量 = 0
        .手配区分 = ""
        .発注比率 = 0
        .発注比率計 = 0
        .件数 = 0
        .計画年月 = ""
        .作成者 = ""
'       .作成日 = 0
'       .作成端末 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function SHHInsert(Optional DBNo As Integer) As Boolean
    SHHInsert = True
    With SHH
        gSL_Select = "Insert Into 生産計画必要数テーブル  ( "
        gSL_Select = gSL_Select & "内示年月,品番,手配先,"
        gSL_Select = gSL_Select & "受入,工区,納期,"
        gSL_Select = gSL_Select & "数量,分配前数量,手配区分,"
        gSL_Select = gSL_Select & "発注比率,発注比率計,件数,"
        gSL_Select = gSL_Select & "計画年月,"
        gSL_Select = gSL_Select & "作成者,作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.内示年月) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.手配先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.工区) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納期) & "',"
        gSL_Select = gSL_Select & "" & .数量 & ","
        gSL_Select = gSL_Select & "" & .分配前数量 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "" & .発注比率 & ","
        gSL_Select = gSL_Select & "" & .発注比率計 & ","
        gSL_Select = gSL_Select & "" & .件数 & ","
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
    SHHInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function SHHUpdate(Optional DBNo As Integer) As Boolean
    SHHUpdate = False
    With SHH
        gSL_Select = "Update 生産計画必要数テーブル  Set "
        gSL_Select = gSL_Select & "内示年月 = '" & RTrim(.内示年月) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "手配先 = '" & RTrim(.手配先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "工区 = '" & RTrim(.工区) & "',"
        gSL_Select = gSL_Select & "納期 = '" & RTrim(.納期) & "',"
        gSL_Select = gSL_Select & "数量 = " & .数量 & ","
        gSL_Select = gSL_Select & "分配前数量 = " & .分配前数量 & ","
        gSL_Select = gSL_Select & "手配区分 = '" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "発注比率 = " & .発注比率 & ","
        gSL_Select = gSL_Select & "発注比率計 = " & .発注比率計 & ","
        gSL_Select = gSL_Select & "件数 = " & .件数 & ","
        gSL_Select = gSL_Select & "計画年月 = '" & RTrim(.計画年月) & "'"
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
    SHHUpdate = True
End Function

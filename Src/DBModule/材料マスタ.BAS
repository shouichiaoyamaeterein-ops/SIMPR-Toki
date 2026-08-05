Attribute VB_Name = "材料マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料マスタ
'*
'*   モジュールID  : 材料マスタ.Bas
'*
'*   作成日  :  2004/02/06  By Cis
'*   変更日  :  2004/02/18  By Cis 材料表示編集ﾓｼﾞｭｰﾙ追加
'*   変更日  :  2008/02/06  By Cis 受給品番/発注品番/設定数　追加
'*   変更日  :  2009/02/26  By Cis 受給品番/発注品番　削除、設定数/所要量区分　追加
'*---------------------------------------------------------------------------------
    Type T_材料マスタ
        材料管理番号          As Long
        材質                  As String
        寸法区分              As Byte
        板厚                  As Currency
        幅                    As Currency
        長さ                  As Currency
        材料名                As String
        材種                  As String
        形状                  As String
        単位                  As String
        換算値                As Currency
        手配単位              As String
        実数区分              As Byte
        在庫管理区分          As Byte
        在庫管理数値          As Currency
        在庫管理係数          As Currency
        在庫管理日数          As Currency
        在庫管理発注点        As Long
        在庫区分              As Byte
        材種コード            As String
        仮区分                As Byte
        表示寸法              As String
        SIZE                 As String
'        作成日                As Date
        作成者                As String
'        作成端末              As String
'        更新日                As Date
        更新者                As String
'        更新端末              As String
        件数          As Long
        
'-------------------------------------------------------------------------------' 2009.02.27 del start
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'        受給品番            As String
'        発注品番            As String
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'-------------------------------------------------------------------------------' 2009.02.27 del end

'-------------------------------------------------------------------------------' 2009.02.26 add start
        設定数              As Currency
        所要量区分          As Byte
'-------------------------------------------------------------------------------' 2009.02.26 add end

    End Type

    Public ZRM As T_材料マスタ

    Public ZRM_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZRM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZRM
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With ZRM
            Select Case adoCol.Name
                Case "材料管理番号":          If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "材質":                  If Not IsNull(adoCol.Value) Then .材質 = adoCol.Value
                Case "寸法区分":              If Not IsNull(adoCol.Value) Then .寸法区分 = adoCol.Value
                Case "板厚":                  If Not IsNull(adoCol.Value) Then .板厚 = adoCol.Value
                Case "幅":                    If Not IsNull(adoCol.Value) Then .幅 = adoCol.Value
                Case "長さ":                  If Not IsNull(adoCol.Value) Then .長さ = adoCol.Value
                Case "材料名":                If Not IsNull(adoCol.Value) Then .材料名 = adoCol.Value
                Case "材種":                  If Not IsNull(adoCol.Value) Then .材種 = adoCol.Value
                Case "形状":                  If Not IsNull(adoCol.Value) Then .形状 = adoCol.Value
                Case "単位":                  If Not IsNull(adoCol.Value) Then .単位 = adoCol.Value
                Case "換算値":                If Not IsNull(adoCol.Value) Then .換算値 = adoCol.Value
                Case "手配単位":              If Not IsNull(adoCol.Value) Then .手配単位 = adoCol.Value
                Case "実数区分":              If Not IsNull(adoCol.Value) Then .実数区分 = adoCol.Value
                Case "在庫管理区分":          If Not IsNull(adoCol.Value) Then .在庫管理区分 = adoCol.Value
                Case "在庫管理数値":          If Not IsNull(adoCol.Value) Then .在庫管理数値 = adoCol.Value
                Case "在庫管理係数":          If Not IsNull(adoCol.Value) Then .在庫管理係数 = adoCol.Value
                Case "在庫管理日数":          If Not IsNull(adoCol.Value) Then .在庫管理日数 = adoCol.Value
                Case "在庫管理発注点":        If Not IsNull(adoCol.Value) Then .在庫管理発注点 = adoCol.Value
                Case "在庫区分":              If Not IsNull(adoCol.Value) Then .在庫区分 = adoCol.Value
                Case "材種コード":            If Not IsNull(adoCol.Value) Then .材種コード = adoCol.Value
                Case "仮区分":                If Not IsNull(adoCol.Value) Then .仮区分 = adoCol.Value
                Case "表示寸法":              If Not IsNull(adoCol.Value) Then .表示寸法 = adoCol.Value
                Case "SIZE":                  If Not IsNull(adoCol.Value) Then .SIZE = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
                Case "作成者":                If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
'                Case "更新日":                If Not IsNull(adoCol.Value) Then .更新日               = adoCol.Value
                Case "更新者":                If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新端末":              If Not IsNull(adoCol.Value) Then .更新端末             = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'---------------------------------------------------------------------------------------' 2009.02.27 del start
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'                Case "受給品番":         If Not IsNull(adoCol.Value) Then .受給品番 = adoCol.Value
'                Case "発注品番":         If Not IsNull(adoCol.Value) Then .発注品番 = adoCol.Value
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'---------------------------------------------------------------------------------------' 2009.02.27 del end
'---------------------------------------------------------------' 2009.02.26 add start
                Case "設定数":                If Not IsNull(adoCol.Value) Then .設定数 = adoCol.Value
                Case "所要量区分":            If Not IsNull(adoCol.Value) Then .所要量区分 = adoCol.Value
'---------------------------------------------------------------' 2009.02.26 add end
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZRMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZRMRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 材料マスタ"
        RTxtSql = RTxtSql & " Where 材料管理番号 = " & ZRM.材料管理番号 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZRMRead = True
           Call ItemsFromZRM(DBNo)
           ZRM_RDSTS = True
        Else
           ZRM_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call ZRMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZRMReadNext(Optional DBNo As Integer = 0)
    If Not ZRM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZRM_RDSTS = False
        Else
           Call ItemsFromZRM(DBNo)
           ZRM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZRMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZRM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZRM )
'*----------------------------------------------*'
Public Sub ItemsClearZRM()
    With ZRM
        .材料管理番号 = 0
        .材質 = ""
        .寸法区分 = 0
        .板厚 = 0
        .幅 = 0
        .長さ = 0
        .材料名 = ""
        .材種 = ""
        .形状 = ""
        .単位 = ""
        .換算値 = 0
        .手配単位 = ""
        .実数区分 = 0
        .在庫管理区分 = 0
        .在庫管理数値 = 0
        .在庫管理係数 = 0
        .在庫管理日数 = 0
        .在庫管理発注点 = 0
        .在庫区分 = 0
        .材種コード = ""
        .仮区分 = 0
        .表示寸法 = ""
        .SIZE = ""
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
'       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
        .件数 = 0
'-------------------------------------------------------------------------------' 2009.02.27 del start
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'        .受給品番 = ""
'        .発注品番 = ""
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'-------------------------------------------------------------------------------' 2009.02.27 del end
'-------------------------------------------------------------------------------' 2009.02.26 add start
        .設定数 = 0
'-------------------------------------------------------------------------------' 2009.02.26 add end
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZRMInsert(Optional DBNo As Integer) As Boolean
    ZRMInsert = True
    With ZRM
        gSL_Select = "Insert Into 材料マスタ  ( "
        gSL_Select = gSL_Select & "材料管理番号,材質,寸法区分,"
        gSL_Select = gSL_Select & "板厚,幅,長さ,"
        gSL_Select = gSL_Select & "材料名,材種,形状,"
        gSL_Select = gSL_Select & "単位,換算値,手配単位,"
        gSL_Select = gSL_Select & "実数区分,在庫管理区分,在庫管理数値,"
        gSL_Select = gSL_Select & "在庫管理係数,在庫管理日数,在庫管理発注点,"
        gSL_Select = gSL_Select & "在庫区分,材種コード,仮区分,"
        gSL_Select = gSL_Select & "表示寸法,SIZE,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
'-------------------------------------------------------------------------------' 2009.02.27 del start
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'        gSL_Select = gSL_Select & ",受給品番,発注品番"
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'-------------------------------------------------------------------------------' 2009.02.27 del end
'-------------------------------------------------------------------------------' 2009.02.26 add start
        gSL_Select = gSL_Select & ",設定数"
        gSL_Select = gSL_Select & ",所要量区分"
'-------------------------------------------------------------------------------' 2009.02.26 add end
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "" & .寸法区分 & ","
        gSL_Select = gSL_Select & "" & .板厚 & ","
        gSL_Select = gSL_Select & "" & .幅 & ","
        gSL_Select = gSL_Select & "" & .長さ & ","
        gSL_Select = gSL_Select & "'" & RTrim(.材料名) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.材種) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.形状) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.単位) & "',"
        gSL_Select = gSL_Select & "" & .換算値 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.手配単位) & "',"
        gSL_Select = gSL_Select & "" & .実数区分 & ","
        gSL_Select = gSL_Select & "" & .在庫管理区分 & ","
        gSL_Select = gSL_Select & "" & .在庫管理数値 & ","
        gSL_Select = gSL_Select & "" & .在庫管理係数 & ","
        gSL_Select = gSL_Select & "" & .在庫管理日数 & ","
        gSL_Select = gSL_Select & "" & .在庫管理発注点 & ","
        gSL_Select = gSL_Select & "" & .在庫区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.材種コード) & "',"
        gSL_Select = gSL_Select & "" & .仮区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.表示寸法) & "',"
        gSL_Select = gSL_Select & "'" & Format(.板厚, "0000.00")
        gSL_Select = gSL_Select & "" & Format(.幅, "0000.00")
        gSL_Select = gSL_Select & "" & Format(.長さ, "0000.00") & "',"
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "',"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"
'-------------------------------------------------------------------------------' 2009.02.27 del start
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'        gSL_Select = gSL_Select & ",'" & RTrim(.受給品番) & "'"
'        gSL_Select = gSL_Select & ",'" & RTrim(.発注品番) & "'"
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'-------------------------------------------------------------------------------' 2009.02.27 del end
'---------------------------------------------------------------' 2009.02.26 add start
        gSL_Select = gSL_Select & "," & .設定数 & ""
        gSL_Select = gSL_Select & "," & .所要量区分 & ""
'---------------------------------------------------------------' 2009.02.26 add end
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZRMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZRMUpdate(Optional DBNo As Integer) As Boolean
    ZRMUpdate = False
    With ZRM
        gSL_Select = "Update 材料マスタ  Set "
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "寸法区分 = " & .寸法区分 & ","
        gSL_Select = gSL_Select & "板厚 = " & .板厚 & ","
        gSL_Select = gSL_Select & "幅 = " & .幅 & ","
        gSL_Select = gSL_Select & "長さ = " & .長さ & ","
        gSL_Select = gSL_Select & "材料名 = '" & RTrim(.材料名) & "',"
        gSL_Select = gSL_Select & "材種 = '" & RTrim(.材種) & "',"
        gSL_Select = gSL_Select & "形状 = '" & RTrim(.形状) & "',"
        gSL_Select = gSL_Select & "単位 = '" & RTrim(.単位) & "',"
        gSL_Select = gSL_Select & "換算値 = " & .換算値 & ","
        gSL_Select = gSL_Select & "手配単位 = '" & RTrim(.手配単位) & "',"
        gSL_Select = gSL_Select & "実数区分 = " & .実数区分 & ","
        gSL_Select = gSL_Select & "在庫管理区分 = " & .在庫管理区分 & ","
        gSL_Select = gSL_Select & "在庫管理数値 = " & .在庫管理数値 & ","
        gSL_Select = gSL_Select & "在庫管理係数 = " & .在庫管理係数 & ","
        gSL_Select = gSL_Select & "在庫管理日数 = " & .在庫管理日数 & ","
        gSL_Select = gSL_Select & "在庫管理発注点 = " & .在庫管理発注点 & ","
        gSL_Select = gSL_Select & "在庫区分 = " & .在庫区分 & ","
        gSL_Select = gSL_Select & "材種コード = '" & RTrim(.材種コード) & "',"
        gSL_Select = gSL_Select & "仮区分 = " & .仮区分 & ","
        gSL_Select = gSL_Select & "表示寸法 = '" & RTrim(.表示寸法) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
'-------------------------------------------------------------------------------' 2009.02.27 del start
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT START
'        gSL_Select = gSL_Select & ",受給品番 = '" & RTrim(.受給品番) & "'"
'        gSL_Select = gSL_Select & ",発注品番 = '" & RTrim(.発注品番) & "'"
''-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/06 INSERT END
'-------------------------------------------------------------------------------' 2009.02.27 del end
'---------------------------------------------------------------' 2009.02.26 add start
        gSL_Select = gSL_Select & ",設定数     = " & .設定数 & ""
        gSL_Select = gSL_Select & ",所要量区分 = " & .所要量区分 & ""
'---------------------------------------------------------------' 2009.02.26 add end
        gSL_Select = gSL_Select & " Where 材料管理番号 = " & .材料管理番号 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZRMUpdate = True
End Function
'+--------------------------+
'+  材料表示編集
'+--------------------------+
Public Function EditZairyo(Optional Zai As String = "", Optional Ita As Currency = 0, Optional Haba As Currency = 0, Optional Nagasa As Currency = 0) As String
    Dim eZairyo     As String
    Dim eSunpou     As Byte
    
    eZairyo = ""
    eSunpou = 1
    
    If Trim(Zai) = "" Then
       eSunpou = ZRM.寸法区分
       Zai = ZRM.材質
       Ita = ZRM.板厚
       Haba = ZRM.幅
       Nagasa = ZRM.長さ
    End If
    If Ita <> 0 Or Haba <> 0 Or Nagasa <> 0 Then
       eSunpou = 0
    End If
    
    If Trim(Zai) = "" Then
       EditZairyo = ""
       Exit Function
    End If
    
    eZairyo = RTrim(ZRM.材質) & Space(2)
    If eSunpou = 0 Then
       eZairyo = eZairyo & _
                 Format(Ita, "###0.00") & "x" & _
                 Format(Haba, "###0.00") & "x" & _
                 Format(Nagasa, "###0.00")
    End If
    
    EditZairyo = eZairyo

End Function
                

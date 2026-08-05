Attribute VB_Name = "材料品番マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料品番マスタ
'*
'*   モジュールID  : 材料品番マスタ.Bas
'*
'*                                      /項目追加含む/
'*   作成日  :  2004/01/23  By Cis
'*   変更日  :  2004/04/05          ピッチ タイプ変更( Integer → Currency )
'*   変更日  :  2004/11/15  By Cis  表示寸法　項目追加
'*   変更日  :  2008/02/13  By Cis  単位　　　項目追加
'*---------------------------------------------------------------------------------
    Type T_材料品番マスタ
        材料管理番号  As Long
        使用品番      As String
' *CHG* 2004/04/05
'        ピッチ        As Integer
        ピッチ        As Currency
        
        取数          As Long
        使用量        As Currency
        未使用量      As Currency
        有効期限開始  As String
        有効期限終了  As String
        展開区分      As Byte
        LT            As Byte
'        作成日        As Date
        作成者        As String
'        作成端末      As String
'        更新日        As Date
        更新者        As String
'        更新端末      As String
        件数          As Long
'   項目追加(自動)
        品名          As String
        材質          As String
        板厚          As Double
        幅            As Double
        長さ          As Double
        背番号        As String
        寸法区分      As Byte
        品番編集      As String
    
'   項目追加
        表示寸法        As String           ' 2004/11/15
        単位            As String           ' 2008/02/13
    End Type

    Public ZHM As T_材料品番マスタ

    Public ZHM_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZHM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZHM
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With ZHM
            Select Case adoCol.Name
                Case "材料管理番号":  If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "使用品番":      If Not IsNull(adoCol.Value) Then .使用品番 = adoCol.Value
                Case "ピッチ":        If Not IsNull(adoCol.Value) Then .ピッチ = adoCol.Value
                Case "取数":          If Not IsNull(adoCol.Value) Then .取数 = adoCol.Value
                Case "使用量":        If Not IsNull(adoCol.Value) Then .使用量 = adoCol.Value
                Case "未使用量":      If Not IsNull(adoCol.Value) Then .未使用量 = adoCol.Value
                Case "有効期限開始":  If Not IsNull(adoCol.Value) Then .有効期限開始 = adoCol.Value
                Case "有効期限終了":  If Not IsNull(adoCol.Value) Then .有効期限終了 = adoCol.Value
                Case "展開区分":      If Not IsNull(adoCol.Value) Then .展開区分 = adoCol.Value
                Case "LT":            If Not IsNull(adoCol.Value) Then .LT = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "品名":    If Not IsNull(adoCol.Value) Then .品名 = adoCol.Value
                Case "材質":    If Not IsNull(adoCol.Value) Then .材質 = adoCol.Value
                Case "板厚":    If Not IsNull(adoCol.Value) Then .板厚 = adoCol.Value
                Case "幅":      If Not IsNull(adoCol.Value) Then .幅 = adoCol.Value
                Case "長さ":    If Not IsNull(adoCol.Value) Then .長さ = adoCol.Value
                Case "背番号":  If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "寸法区分":        If Not IsNull(adoCol.Value) Then .寸法区分 = adoCol.Value
                Case "品番編集":        If Not IsNull(adoCol.Value) Then .品番編集 = adoCol.Value
            
'   項目追加
                Case "表示寸法":    If Not IsNull(adoCol.Value) Then .表示寸法 = adoCol.Value
                Case "単位":    If Not IsNull(adoCol.Value) Then .単位 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZHMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZHMRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 材料品番マスタ"
        RTxtSql = RTxtSql & " Where 材料管理番号 = " & ZHM.材料管理番号 & ""
        RTxtSql = RTxtSql & " And   使用品番 = '" & RTrim(ZHM.使用品番) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZHMRead = True
           Call ItemsFromZHM(DBNo)
           ZHM_RDSTS = True
        Else
           ZHM_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call ZHMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZHMReadNext(Optional DBNo As Integer = 0)
    If Not ZHM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZHM_RDSTS = False
        Else
           Call ItemsFromZHM(DBNo)
           ZHM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZHMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZHM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZHM )
'*----------------------------------------------*'
Public Sub ItemsClearZHM()
    With ZHM
        .材料管理番号 = 0
        .使用品番 = ""
        .ピッチ = 0
        .取数 = 0
        .使用量 = 0
        .未使用量 = 0
        .有効期限開始 = ""
        .有効期限終了 = ""
        .展開区分 = 0
        .LT = 0
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
'       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
        .件数 = 0
'   項目追加(自動)
        .品名 = ""
        .材質 = ""
        .板厚 = 0
        .幅 = 0
        .長さ = 0
        .背番号 = ""
        .寸法区分 = 0
        .品番編集 = ""
    
'   項目追加
        .表示寸法 = ""
        .単位 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZHMInsert(Optional DBNo As Integer) As Boolean
    ZHMInsert = True
    With ZHM
        gSL_Select = "Insert Into 材料品番マスタ  ( "
        gSL_Select = gSL_Select & "材料管理番号,使用品番,ピッチ,"
        gSL_Select = gSL_Select & "取数,使用量,未使用量,"
        gSL_Select = gSL_Select & "有効期限開始,有効期限終了,展開区分,"
        gSL_Select = gSL_Select & "LT,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.使用品番) & "',"
        gSL_Select = gSL_Select & "" & .ピッチ & ","
        gSL_Select = gSL_Select & "" & .取数 & ","
        gSL_Select = gSL_Select & "" & .使用量 & ","
        gSL_Select = gSL_Select & "" & .未使用量 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.有効期限開始) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.有効期限終了) & "',"
        gSL_Select = gSL_Select & "" & .展開区分 & ","
        gSL_Select = gSL_Select & "" & .LT & ","
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "',"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZHMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZHMUpdate(Optional DBNo As Integer) As Boolean
    ZHMUpdate = False
    With ZHM
        gSL_Select = "Update 材料品番マスタ  Set "
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "使用品番 = '" & RTrim(.使用品番) & "',"
        gSL_Select = gSL_Select & "ピッチ = " & .ピッチ & ","
        gSL_Select = gSL_Select & "取数 = " & .取数 & ","
        gSL_Select = gSL_Select & "使用量 = " & .使用量 & ","
        gSL_Select = gSL_Select & "未使用量 = " & .未使用量 & ","
        gSL_Select = gSL_Select & "有効期限開始 = '" & RTrim(.有効期限開始) & "',"
        gSL_Select = gSL_Select & "有効期限終了 = '" & RTrim(.有効期限終了) & "',"
        gSL_Select = gSL_Select & "展開区分 = " & .展開区分 & ","
        gSL_Select = gSL_Select & "LT = " & .LT & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 材料管理番号 = " & .材料管理番号 & ""
        gSL_Select = gSL_Select & "   And 使用品番 = '" & RTrim(.使用品番) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZHMUpdate = True
End Function

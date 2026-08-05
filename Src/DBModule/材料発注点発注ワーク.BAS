Attribute VB_Name = "材料発注点発注ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料発注点発注ワーク
'*
'*   モジュールID  : 材料発注点発注ワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*                                      /項目追加含む/
'*   作成日  :  2004/11/22  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料発注点発注ワーク
        処理端末      As String
        材料管理番号  As Long
        仕入先        As String
        受入          As String
        現在庫数      As Currency
        発注予定数    As Currency
        所要量        As Currency
'        作成日        As Date
        件数          As Long
'   項目追加(自動)
        設定数        As Currency
        在庫管理区分  As Byte
        在庫管理係数  As Currency
        在庫管理日数  As Currency
        在庫管理発注点 As Currency
        在庫管理名    As String
        表示寸法      As String
        収容数1       As Currency
        発注ロット1   As Currency
        切替日        As String
        切替便        As String
        収容数2       As Currency
        発注ロット2   As Currency
        Week1         As Byte
        Week2         As Byte
        Week3         As Byte
        Week4         As Byte
        Week5         As Byte
        Day1          As Byte
        Day2          As Byte
        Day3          As Byte
        Day4          As Byte
        Day5          As Byte
        背番号        As String
        仕入先名      As String
        担当者        As String
        担当名        As String
        手配区分      As String
        SK区分        As Byte
        材質          As String
    End Type

    Public ZHW As T_材料発注点発注ワーク

    Public ZHW_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZHW(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZHW
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With ZHW
            Select Case adoCol.Name
                Case "処理端末":      If Not IsNull(adoCol.Value) Then .処理端末 = adoCol.Value
                Case "材料管理番号":  If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "仕入先":        If Not IsNull(adoCol.Value) Then .仕入先 = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "現在庫数":      If Not IsNull(adoCol.Value) Then .現在庫数 = adoCol.Value
                Case "発注予定数":    If Not IsNull(adoCol.Value) Then .発注予定数 = adoCol.Value
                Case "所要量":        If Not IsNull(adoCol.Value) Then .所要量 = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "設定数":  If Not IsNull(adoCol.Value) Then .設定数 = adoCol.Value
                Case "在庫管理区分":    If Not IsNull(adoCol.Value) Then .在庫管理区分 = adoCol.Value
                Case "在庫管理係数":    If Not IsNull(adoCol.Value) Then .在庫管理係数 = adoCol.Value
                Case "在庫管理日数":    If Not IsNull(adoCol.Value) Then .在庫管理日数 = adoCol.Value
                Case "在庫管理発注点":  If Not IsNull(adoCol.Value) Then .在庫管理発注点 = adoCol.Value
                Case "在庫管理名":      If Not IsNull(adoCol.Value) Then .在庫管理名 = adoCol.Value
                Case "表示寸法":        If Not IsNull(adoCol.Value) Then .表示寸法 = adoCol.Value
                Case "収容数1": If Not IsNull(adoCol.Value) Then .収容数1 = adoCol.Value
                Case "発注ロット1":     If Not IsNull(adoCol.Value) Then .発注ロット1 = adoCol.Value
                Case "切替日":  If Not IsNull(adoCol.Value) Then .切替日 = adoCol.Value
                Case "切替便":  If Not IsNull(adoCol.Value) Then .切替便 = adoCol.Value
                Case "収容数2": If Not IsNull(adoCol.Value) Then .収容数2 = adoCol.Value
                Case "発注ロット2":     If Not IsNull(adoCol.Value) Then .発注ロット2 = adoCol.Value
                Case "Week1":   If Not IsNull(adoCol.Value) Then .Week1 = adoCol.Value
                Case "Week2":   If Not IsNull(adoCol.Value) Then .Week2 = adoCol.Value
                Case "Week3":   If Not IsNull(adoCol.Value) Then .Week3 = adoCol.Value
                Case "Week4":   If Not IsNull(adoCol.Value) Then .Week4 = adoCol.Value
                Case "Week5":   If Not IsNull(adoCol.Value) Then .Week5 = adoCol.Value
                Case "Day1":    If Not IsNull(adoCol.Value) Then .Day1 = adoCol.Value
                Case "Day2":    If Not IsNull(adoCol.Value) Then .Day2 = adoCol.Value
                Case "Day3":    If Not IsNull(adoCol.Value) Then .Day3 = adoCol.Value
                Case "Day4":    If Not IsNull(adoCol.Value) Then .Day4 = adoCol.Value
                Case "Day5":    If Not IsNull(adoCol.Value) Then .Day5 = adoCol.Value
                Case "背番号":  If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "仕入先名":        If Not IsNull(adoCol.Value) Then .仕入先名 = adoCol.Value
                Case "担当者":  If Not IsNull(adoCol.Value) Then .担当者 = adoCol.Value
                Case "担当名":  If Not IsNull(adoCol.Value) Then .担当名 = adoCol.Value
                Case "手配区分":        If Not IsNull(adoCol.Value) Then .手配区分 = adoCol.Value
                Case "SK区分":  If Not IsNull(adoCol.Value) Then .SK区分 = adoCol.Value
                Case "材質":    If Not IsNull(adoCol.Value) Then .材質 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZHWRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZHWRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 材料発注点発注ワーク"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZHWRead = True
           Call ItemsFromZHW(DBNo)
           ZHW_RDSTS = True
        Else
           ZHW_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call ZHWClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZHWReadNext(Optional DBNo As Integer = 0)
    If Not ZHW_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZHW_RDSTS = False
        Else
           Call ItemsFromZHW(DBNo)
           ZHW_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZHWClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZHW_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZHW )
'*----------------------------------------------*'
Public Sub ItemsClearZHW()
    With ZHW
        .処理端末 = ""
        .材料管理番号 = 0
        .仕入先 = ""
        .受入 = ""
        .現在庫数 = 0
        .発注予定数 = 0
        .所要量 = 0
'       .作成日 = 0
        .件数 = 0
'   項目追加(自動)
        .設定数 = 0
        .在庫管理区分 = 0
        .在庫管理係数 = 0
        .在庫管理日数 = 0
        .在庫管理発注点 = 0
        .在庫管理名 = ""
        .表示寸法 = ""
        .収容数1 = 0
        .発注ロット1 = 0
        .切替日 = ""
        .切替便 = ""
        .収容数2 = 0
        .発注ロット2 = 0
        .Week1 = 0
        .Week2 = 0
        .Week3 = 0
        .Week4 = 0
        .Week5 = 0
        .Day1 = 0
        .Day2 = 0
        .Day3 = 0
        .Day4 = 0
        .Day5 = 0
        .背番号 = ""
        .仕入先名 = ""
        .担当者 = ""
        .担当名 = ""
        .手配区分 = ""
        .SK区分 = 0
        .材質 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZHWInsert(Optional DBNo As Integer) As Boolean
    ZHWInsert = True
    With ZHW
        gSL_Select = "Insert Into 材料発注点発注ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,材料管理番号,仕入先,"
        gSL_Select = gSL_Select & "受入,現在庫数,発注予定数,"
        gSL_Select = gSL_Select & "所要量,"
        gSL_Select = gSL_Select & "作成日"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "" & .現在庫数 & ","
        gSL_Select = gSL_Select & "" & .発注予定数 & ","
        gSL_Select = gSL_Select & "" & .所要量 & ","
        gSL_Select = gSL_Select & "GETDATE()"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZHWInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZHWUpdate(Optional DBNo As Integer) As Boolean
    ZHWUpdate = False
    With ZHW
        gSL_Select = "Update 材料発注点発注ワーク  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "現在庫数 = " & .現在庫数 & ","
        gSL_Select = gSL_Select & "発注予定数 = " & .発注予定数 & ","
        gSL_Select = gSL_Select & "所要量 = " & .所要量 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZHWUpdate = True
End Function

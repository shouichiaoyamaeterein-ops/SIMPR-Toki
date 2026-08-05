Attribute VB_Name = "カレンダマスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   カレンダマスタ
'*
'*   モジュールID  : カレンダマスタ.Bas
'*
'*
'*
'*   作成日  :  2003/12/25  By Cis
'*   変更日  :  2003/12/25  稼働日チェック追加
'*   変更日  :  2005/02/28  臨時昼区分・臨時夜区分、追加
'*---------------------------------------------------------------------------------
    Type T_カレンダマスタ
        取引先区分        As Byte
        取引先            As String
        年                As String
        月                As String
        受入              As String
        末日              As Byte
        稼動日数          As Currency
        昼勤日数          As Byte
        夜勤日数          As Byte
        稼動区分          As String
        夜勤区分          As String
'        作成日            As Date
'        作成端末          As String
'        更新日            As Date
'        更新端末          As String
        臨時昼区分      As String
        臨時夜区分      As String
        件数          As Long
    
        略称              As String
        工場名            As String
    End Type

    Public CLM As T_カレンダマスタ

    Public CLM_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object

    Global clHiruKB     As Boolean
    Global clYoruKB     As Boolean
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromCLM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearCLM
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With CLM
            Select Case adoCol.Name
                Case "取引先区分":        If Not IsNull(adoCol.Value) Then .取引先区分 = adoCol.Value
                Case "取引先":            If Not IsNull(adoCol.Value) Then .取引先 = adoCol.Value
                Case "年":                If Not IsNull(adoCol.Value) Then .年 = adoCol.Value
                Case "月":                If Not IsNull(adoCol.Value) Then .月 = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "末日":              If Not IsNull(adoCol.Value) Then .末日 = adoCol.Value
                Case "稼動日数":          If Not IsNull(adoCol.Value) Then .稼動日数 = adoCol.Value
                Case "昼勤日数":          If Not IsNull(adoCol.Value) Then .昼勤日数 = adoCol.Value
                Case "夜勤日数":          If Not IsNull(adoCol.Value) Then .夜勤日数 = adoCol.Value
                Case "稼動区分":          If Not IsNull(adoCol.Value) Then .稼動区分 = adoCol.Value
                Case "夜勤区分":          If Not IsNull(adoCol.Value) Then .夜勤区分 = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "臨時昼区分":          If Not IsNull(adoCol.Value) Then .臨時昼区分 = adoCol.Value
                Case "臨時夜区分":          If Not IsNull(adoCol.Value) Then .臨時夜区分 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
            
                Case "略称":         If Not IsNull(adoCol.Value) Then .略称 = adoCol.Value
                Case "工場名":         If Not IsNull(adoCol.Value) Then .工場名 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function CLMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    CLMRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From カレンダマスタ"
        RTxtSql = RTxtSql & " Where 取引先区分 = " & CLM.取引先区分 & ""
        RTxtSql = RTxtSql & " And   取引先 = '" & RTrim(CLM.取引先) & "'"
        RTxtSql = RTxtSql & " And   年 = '" & RTrim(CLM.年) & "'"
        RTxtSql = RTxtSql & " And   月 = '" & RTrim(CLM.月) & "'"
        RTxtSql = RTxtSql & " And   受入 = '" & RTrim(CLM.受入) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           CLMRead = True
           Call ItemsFromCLM(DBNo)
           CLM_RDSTS = True
        Else
           CLM_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call CLMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub CLMReadNext(Optional DBNo As Integer = 0)
    If Not CLM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           CLM_RDSTS = False
        Else
           Call ItemsFromCLM(DBNo)
           CLM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub CLMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        CLM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( CLM )
'*----------------------------------------------*'
Public Sub ItemsClearCLM()
    With CLM
        .取引先区分 = 0
        .取引先 = ""
        .年 = ""
        .月 = ""
        .受入 = ""
        .末日 = 0
        .稼動日数 = 0
        .昼勤日数 = 0
        .夜勤日数 = 0
        .稼動区分 = ""
        .夜勤区分 = ""
'       .作成日 = 0
'       .作成端末 = ""
'       .更新日 = 0
'       .更新端末 = ""
        .臨時昼区分 = ""
        .臨時夜区分 = ""
        .件数 = 0
    
        .略称 = ""
        .工場名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function CLMInsert(Optional DBNo As Integer) As Boolean
    CLMInsert = True
    With CLM
        gSL_Select = "Insert Into カレンダマスタ  ( "
        gSL_Select = gSL_Select & "取引先区分,取引先,年,"
        gSL_Select = gSL_Select & "月,受入,末日,"
        gSL_Select = gSL_Select & "稼動日数,昼勤日数,夜勤日数,"
        gSL_Select = gSL_Select & "稼動区分,夜勤区分,臨時昼区分,臨時夜区分,"
        gSL_Select = gSL_Select & "作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .取引先区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.年) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.月) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "" & .末日 & ","
        gSL_Select = gSL_Select & "" & .稼動日数 & ","
        gSL_Select = gSL_Select & "" & .昼勤日数 & ","
        gSL_Select = gSL_Select & "" & .夜勤日数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.稼動区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.夜勤区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.臨時昼区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.臨時夜区分) & "',"
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    CLMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function CLMUpdate(Optional DBNo As Integer) As Boolean
    CLMUpdate = False
    With CLM
        gSL_Select = "Update カレンダマスタ  Set "
        gSL_Select = gSL_Select & "取引先区分 = " & .取引先区分 & ","
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "年 = '" & RTrim(.年) & "',"
        gSL_Select = gSL_Select & "月 = '" & RTrim(.月) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "末日 = " & .末日 & ","
        gSL_Select = gSL_Select & "稼動日数 = " & .稼動日数 & ","
        gSL_Select = gSL_Select & "昼勤日数 = " & .昼勤日数 & ","
        gSL_Select = gSL_Select & "夜勤日数 = " & .夜勤日数 & ","
        gSL_Select = gSL_Select & "稼動区分 = '" & RTrim(.稼動区分) & "',"
        gSL_Select = gSL_Select & "夜勤区分 = '" & RTrim(.夜勤区分) & "',"
        gSL_Select = gSL_Select & "臨時昼区分 = '" & RTrim(.臨時昼区分) & "',"
        gSL_Select = gSL_Select & "臨時夜区分 = '" & RTrim(.臨時夜区分) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 取引先区分 = " & .取引先区分 & ""
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 年 = '" & RTrim(.年) & "'"
        gSL_Select = gSL_Select & "   And 月 = '" & RTrim(.月) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    CLMUpdate = True
End Function
'*---------------------------------------------------------------------*'
'*    稼働日チェック                                                    *
'*
'*      戻り値  Kdou_Check = 0 :  稼動日
'*                           1 :  非稼働日
'*                           2 :  マスタ未登録
'*              clHiruKB = True : 昼勤稼働日、 false : 昼勤非稼働日
'*              clYoruKB = True : 夜勤稼働日、 false : 夜勤非稼働日
'*---------------------------------------------------------------------*'
Public Function Kdou_Check(ChkYmd As String, Optional ChkTorkb As Byte = 1, Optional ChkTorcd As String = "", Optional ChkUkeir As String = "") As Byte
    Kdou_Check = 1
    
    gSL_Select = "Select * From カレンダマスタ"
    gSL_Select = gSL_Select & " Where 取引先区分 = " & ChkTorkb & ""
    gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(ChkTorcd) & "'"
    gSL_Select = gSL_Select & "   And 年 = '" & Left(ChkYmd, 4) & "'"
    gSL_Select = gSL_Select & "   And 月 = '" & Mid(ChkYmd, 5, 2) & "'"
    gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(ChkUkeir) & "'"
    If Not CLMRead(gSL_Select, 1) Then
        gSL_Select = "Select * From カレンダマスタ"
        gSL_Select = gSL_Select & " Where 取引先区分 = " & ChkTorkb & ""
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(ChkTorcd) & "'"
        gSL_Select = gSL_Select & "   And 年 = '" & Left(ChkYmd, 4) & "'"
        gSL_Select = gSL_Select & "   And 月 = '" & Mid(ChkYmd, 5, 2) & "'"
        gSL_Select = gSL_Select & "   And 受入 = ''"
        If Not CLMRead(gSL_Select, 1) Then
            gSL_Select = "Select * From カレンダマスタ"
            gSL_Select = gSL_Select & " Where 取引先区分 = " & ChkTorkb & ""
            gSL_Select = gSL_Select & "   And 取引先 = ''"
            gSL_Select = gSL_Select & "   And 年 = '" & Left(ChkYmd, 4) & "'"
            gSL_Select = gSL_Select & "   And 月 = '" & Mid(ChkYmd, 5, 2) & "'"
            gSL_Select = gSL_Select & "   And 受入 = ''"
            If Not CLMRead(gSL_Select, 1) Then
                Kdou_Check = 2
                Exit Function
            End If
        End If
    End If
    If Mid(CLM.稼動区分, CisFun.Val2(Mid(ChkYmd, 7, 2)), 1) = "0" Then
        clHiruKB = True
    Else
        clHiruKB = False
    End If
    If Mid(CLM.夜勤区分, CisFun.Val2(Mid(ChkYmd, 7, 2)), 1) = "0" Then
        clYoruKB = True
    Else
        clYoruKB = False
    End If
    
    If clHiruKB Or clYoruKB Then
        Kdou_Check = 0
    End If
End Function


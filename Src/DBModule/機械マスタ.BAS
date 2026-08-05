Attribute VB_Name = "機械マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   機械マスタ
'*
'*   モジュールID  : 機械マスタ.Bas
'*
'*   作成日  :  2004/12/13  By Cis 名称取得ﾓｼﾞｭｰﾙ追加
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_機械マスタ
        機械      As String
        機械名称  As String
'        作成日    As Date
        作成者    As String
'        作成端末  As String
'        更新日    As Date
        更新者    As String
'        更新端末  As String
        件数          As Long
    End Type

    Public KMM As T_機械マスタ

    Public KMM_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object

    Global KMM_Meisyo       As String
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKMM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKMM
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With KMM
            Select Case adoCol.Name
                Case "機械":      If Not IsNull(adoCol.Value) Then .機械 = adoCol.Value
                Case "機械名称":  If Not IsNull(adoCol.Value) Then .機械名称 = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
                Case "作成者":    If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":  If Not IsNull(adoCol.Value) Then .作成端末 = adoCol.Value
'                Case "更新日":    If Not IsNull(adoCol.Value) Then .更新日   = adoCol.Value
                Case "更新者":    If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新端末":  If Not IsNull(adoCol.Value) Then .更新端末 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KMMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KMMRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 機械マスタ"
        RTxtSql = RTxtSql & " Where 機械 = '" & RTrim(KMM.機械) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KMMRead = True
           Call ItemsFromKMM(DBNo)
           KMM_RDSTS = True
        Else
           KMM_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call KMMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KMMReadNext(Optional DBNo As Integer = 0)
    If Not KMM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KMM_RDSTS = False
        Else
           Call ItemsFromKMM(DBNo)
           KMM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KMMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KMM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KMM )
'*----------------------------------------------*'
Public Sub ItemsClearKMM()
    With KMM
        .機械 = ""
        .機械名称 = ""
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
'       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KMMInsert(Optional DBNo As Integer) As Boolean
    KMMInsert = True
    With KMM
        gSL_Select = "Insert Into 機械マスタ  ( "
        gSL_Select = gSL_Select & "機械,機械名称,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.機械) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.機械名称) & "',"
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
    KMMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KMMUpdate(Optional DBNo As Integer) As Boolean
    KMMUpdate = False
    With KMM
        gSL_Select = "Update 機械マスタ  Set "
        gSL_Select = gSL_Select & "機械 = '" & RTrim(.機械) & "',"
        gSL_Select = gSL_Select & "機械名称 = '" & RTrim(.機械名称) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 機械 = '" & RTrim(.機械) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KMMUpdate = True
End Function
'*--------------------------------------------------------------------------------
'*
'*      名称を求める
'*
'*      モジュールID    : KMMMeisyoGet
'*
'*
'*          引数        : Kikai   =   機械№
'*          　　        : DBNo    =   データベースNO ( Optional )
'*
'*          戻値        : True Or False
'*                      : KMM_Meisyo : 機械名称
'*
'*      作成日          :   2004/12/13   By Cis
'*      変更日          :
'*
'*---------------------------------------------------------------------------------
Public Function KMMMeisyoGet(Kikai As String, Optional DBNo As Integer = 0) As Boolean
    KMMMeisyoGet = False
    KMM_Meisyo = ""
    
    gSL_Select = "SELECT * FROM 機械マスタ"
    gSL_Select = gSL_Select & " WHERE 機械 = '" & RTrim(Kikai) & "'"
    If Not KMMRead(gSL_Select, 1, DBNo) Then Exit Function

    KMMMeisyoGet = True
    KMM_Meisyo = RTrim(KMM.機械名称)
    
End Function



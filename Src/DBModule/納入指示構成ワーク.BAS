Attribute VB_Name = "納入指示構成ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   納入指示構成ワーク
'*
'*   モジュールID  : 納入指示構成ワーク.Bas
'*
'*   作成日  :  2010/01/24  By Cis(FMVW5220)
'*   変更日  :  2010/06/02  By Cis  納入指示進捗ワーク項目を追加
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_納入指示構成ワーク
        処理端末      As String
        順位          As String
        レベル        As Byte
        部品材料区分  As String
        大親品番      As String
        大親納入先    As String
        品番          As String
        材料管理番号  As Long
        材質          As String
        板厚          As Currency
        幅            As Currency
        長さ          As Currency
        使用品番      As String
        仕入先        As String
        加工CD        As String
        展開区分      As Byte
        構成数        As Currency
        SEQNO         As Byte
        受入          As String
        作業順        As Byte
        件数          As Long
'   項目追加(自動)
        子品番        As String

'   項目追加(手動)
        値(72)        As String         ' 2010/06/02 Insert
        区分          As Integer        ' 2010/06/02 Insert
    End Type

    Public NSK As T_納入指示構成ワーク

    Public NSK_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol          As Object
    Dim idx             As Long
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromNSK(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearNSK
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With NSK
            Select Case adoCol.Name
                Case "処理端末":      If Not IsNull(adoCol.Value) Then .処理端末 = adoCol.Value
                Case "順位":          If Not IsNull(adoCol.Value) Then .順位 = adoCol.Value
                Case "レベル":        If Not IsNull(adoCol.Value) Then .レベル = adoCol.Value
                Case "部品材料区分":  If Not IsNull(adoCol.Value) Then .部品材料区分 = adoCol.Value
                Case "大親品番":      If Not IsNull(adoCol.Value) Then .大親品番 = adoCol.Value
                Case "大親納入先":    If Not IsNull(adoCol.Value) Then .大親納入先 = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "材料管理番号":  If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "材質":          If Not IsNull(adoCol.Value) Then .材質 = adoCol.Value
                Case "板厚":          If Not IsNull(adoCol.Value) Then .板厚 = adoCol.Value
                Case "幅":            If Not IsNull(adoCol.Value) Then .幅 = adoCol.Value
                Case "長さ":          If Not IsNull(adoCol.Value) Then .長さ = adoCol.Value
                Case "使用品番":      If Not IsNull(adoCol.Value) Then .使用品番 = adoCol.Value
                Case "仕入先":        If Not IsNull(adoCol.Value) Then .仕入先 = adoCol.Value
                Case "加工CD":        If Not IsNull(adoCol.Value) Then .加工CD = adoCol.Value
                Case "展開区分":      If Not IsNull(adoCol.Value) Then .展開区分 = adoCol.Value
                Case "構成数":        If Not IsNull(adoCol.Value) Then .構成数 = adoCol.Value
                Case "SEQNO":         If Not IsNull(adoCol.Value) Then .SEQNO = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "作業順":        If Not IsNull(adoCol.Value) Then .作業順 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "子品番":  If Not IsNull(adoCol.Value) Then .子品番 = adoCol.Value
'--------------------------------------------- 2010/06/02 Insert Start
'   項目追加(手動)
                Case "区分":         If Not IsNull(adoCol.Value) Then .区分 = adoCol.Value
                Case Else
                    For idx = 1 To 72
                        If "値" & Format(idx, "##") = adoCol.Name Then
                            .値(idx) = adoCol.Value
                            Exit For
                        End If
                    Next idx
'--------------------------------------------- 2010/06/02 Insert End
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function NSKRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    NSKRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 納入指示構成ワーク"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           NSKRead = True
           Call ItemsFromNSK(DBNo)
           NSK_RDSTS = True
        Else
           NSK_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call NSKClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub NSKReadNext(Optional DBNo As Integer = 0)
    If Not NSK_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           NSK_RDSTS = False
        Else
           Call ItemsFromNSK(DBNo)
           NSK_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub NSKClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        NSK_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( NSK )
'*----------------------------------------------*'
Public Sub ItemsClearNSK()
    With NSK
        .処理端末 = ""
        .順位 = ""
        .レベル = 0
        .部品材料区分 = ""
        .大親品番 = ""
        .大親納入先 = ""
        .品番 = ""
        .材料管理番号 = 0
        .材質 = ""
        .板厚 = 0
        .幅 = 0
        .長さ = 0
        .使用品番 = ""
        .仕入先 = ""
        .加工CD = ""
        .展開区分 = 0
        .構成数 = 0
        .SEQNO = 0
        .受入 = ""
        .作業順 = 0
        .件数 = 0
'   項目追加(自動)
        .子品番 = ""
'--------------------------------------------- 2010/06/02 Insert Start
'   項目追加(手動)
        .区分 = 0
        For idx = 1 To 72
                .値(idx) = ""
        Next idx
    End With
'--------------------------------------------- 2010/06/02 Insert End
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function NSKInsert(Optional DBNo As Integer) As Boolean
    NSKInsert = False
    With NSK
        gSL_Select = "Insert Into 納入指示構成ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,順位,レベル,"
        gSL_Select = gSL_Select & "部品材料区分,大親品番,大親納入先,"
        gSL_Select = gSL_Select & "品番,材料管理番号,材質,"
        gSL_Select = gSL_Select & "板厚,幅,長さ,"
        gSL_Select = gSL_Select & "使用品番,仕入先,加工CD,"
        gSL_Select = gSL_Select & "展開区分,構成数,SEQNO,"
        gSL_Select = gSL_Select & "受入,作業順"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.順位) & "',"
        gSL_Select = gSL_Select & "" & .レベル & ","
        gSL_Select = gSL_Select & "'" & RTrim(.部品材料区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.大親品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.大親納入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "" & .板厚 & ","
        gSL_Select = gSL_Select & "" & .幅 & ","
        gSL_Select = gSL_Select & "" & .長さ & ","
        gSL_Select = gSL_Select & "'" & RTrim(.使用品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.加工CD) & "',"
        gSL_Select = gSL_Select & "" & .展開区分 & ","
        gSL_Select = gSL_Select & "" & .構成数 & ","
        gSL_Select = gSL_Select & "" & .SEQNO & ","
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "" & .作業順 & ""
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    NSKInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function NSKUpdate(Optional DBNo As Integer) As Boolean
    NSKUpdate = False
    With NSK
        gSL_Select = "Update 納入指示構成ワーク  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "順位 = '" & RTrim(.順位) & "',"
        gSL_Select = gSL_Select & "レベル = " & .レベル & ","
        gSL_Select = gSL_Select & "部品材料区分 = '" & RTrim(.部品材料区分) & "',"
        gSL_Select = gSL_Select & "大親品番 = '" & RTrim(.大親品番) & "',"
        gSL_Select = gSL_Select & "大親納入先 = '" & RTrim(.大親納入先) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "板厚 = " & .板厚 & ","
        gSL_Select = gSL_Select & "幅 = " & .幅 & ","
        gSL_Select = gSL_Select & "長さ = " & .長さ & ","
        gSL_Select = gSL_Select & "使用品番 = '" & RTrim(.使用品番) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "加工CD = '" & RTrim(.加工CD) & "',"
        gSL_Select = gSL_Select & "展開区分 = " & .展開区分 & ","
        gSL_Select = gSL_Select & "構成数 = " & .構成数 & ","
        gSL_Select = gSL_Select & "SEQNO = " & .SEQNO & ","
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "作業順 = " & .作業順 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    NSKUpdate = True
End Function

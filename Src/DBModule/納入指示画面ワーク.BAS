Attribute VB_Name = "納入指示画面ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   納入指示画面ワーク
'*
'*   モジュールID  : 納入指示画面ワーク.Bas
'*
'*   作成日  :  2004/03/29  By Cis
'*   変更日  :  2004/06/03  By Cis - 展開区分＆構成数＆SEQNOの追加
'*   変更日  :  2004/06/09  By Cis - 寸法区分の追加
'*   変更日  :  2004/06/17  By Cis - 社内区分＆備考の追加
'*   変更日  :  2004/09/03  By Cis - 機械名の追加
'*   変更日  :  2004/11/10  By Cis - 取数の追加
'*   変更日  :  2004/12/13  By Cis - 受入、作業順の追加
'*   変更日  :  2010/02/05  By Cis - 機械名称追加
'*
'*---------------------------------------------------------------------------------
    Type T_納入指示画面ワーク
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
        
'-----------------------------------< 2004.12.13 ADD START >
        受入          As String
        作業順        As Byte
'-----------------------------------< 2004.12.13 ADD  END  >
        
        件数          As Long
    '追加
        子品番        As String
        親品番        As String
        仕入先名      As String
        作業名        As String
        形状名        As String
        寸法区分      As Byte
        社内区分      As Byte
        備考          As String
        機械名        As String
        取数          As Long
        機械名称      As String
    End Type

    Public NSG As T_納入指示画面ワーク

    Public NSG_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromNSG(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearNSG
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With NSG
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

'-----------------------------------< 2004.12.13 ADD START >
                Case "受入":        If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "作業順":        If Not IsNull(adoCol.Value) Then .作業順 = adoCol.Value
'-----------------------------------< 2004.12.13 ADD  END  >
                
                Case "SEQNO":         If Not IsNull(adoCol.Value) Then .SEQNO = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
                Case "子品番":       If Not IsNull(adoCol.Value) Then .子品番 = adoCol.Value
                Case "親品番":       If Not IsNull(adoCol.Value) Then .親品番 = adoCol.Value
                Case "仕入先名":     If Not IsNull(adoCol.Value) Then .仕入先名 = adoCol.Value
                Case "作業名":       If Not IsNull(adoCol.Value) Then .作業名 = adoCol.Value
                Case "形状名":       If Not IsNull(adoCol.Value) Then .形状名 = adoCol.Value
                Case "寸法区分":      If Not IsNull(adoCol.Value) Then .寸法区分 = adoCol.Value
                Case "社内区分":      If Not IsNull(adoCol.Value) Then .社内区分 = adoCol.Value
                Case "備考":          If Not IsNull(adoCol.Value) Then .備考 = adoCol.Value
                Case "機械名":       If Not IsNull(adoCol.Value) Then .機械名 = adoCol.Value
                Case "取数":       If Not IsNull(adoCol.Value) Then .取数 = adoCol.Value
                Case "機械名称":     If Not IsNull(adoCol.Value) Then .機械名称 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function NSGRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    NSGRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 納入指示構成ワーク"
        RTxtSql = RTxtSql & " Where 処理端末 = '" & RTrim(NSG.処理端末) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           NSGRead = True
           Call ItemsFromNSG(DBNo)
           NSG_RDSTS = True
        Else
           NSG_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call NSGClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub NSGReadNext(Optional DBNo As Integer = 0)
    If Not NSG_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           NSG_RDSTS = False
        Else
           Call ItemsFromNSG(DBNo)
           NSG_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub NSGClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        NSG_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( NSG )
'*----------------------------------------------*'
Public Sub ItemsClearNSG()
    With NSG
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

'-----------------------------------< 2004.12.13 ADD START >
        .受入 = ""
        .作業順 = 0
'-----------------------------------< 2004.12.13 ADD  END  >
        
        .件数 = 0
        .子品番 = ""
        .親品番 = ""
        .仕入先名 = ""
        .作業名 = ""
        .形状名 = ""
        .寸法区分 = 0
        .社内区分 = 0
        .備考 = ""
        .機械名 = ""
        .取数 = 0
        .機械名称 = ""
    End With
End Sub

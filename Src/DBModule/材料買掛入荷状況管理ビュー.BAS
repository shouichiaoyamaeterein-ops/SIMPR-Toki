Attribute VB_Name = "材料買掛入荷状況管理ビュー"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料買掛入荷状況管理ビュー
'*
'*   モジュールID  : 材料買掛入荷状況管理ビュー.Bas
'*
'*   作成日  :  2009/02/12  By Cis(HP-DC7900)
'*   変更日  :  2009/02/12  By Cis(HP-DC7900) 実績金額配列化
'*---------------------------------------------------------------------------------
    Type T_材料買掛入荷状況管理ビュー
        区分              As String
        編集契約先        As String
        契約先            As String
        契約先名          As String
        予測金額          As Currency
        日当り数          As Currency
        実績金額          As Currency
        進捗              As Currency
        実績金額T(31)     As Currency
        年月              As String
        自社工区          As String
        受入              As String
        件数          As Long
    End Type

    Public ZKNV As T_材料買掛入荷状況管理ビュー

    Public ZKNV_RDSTS    As Boolean                 ' 専用読取フラグ
    Dim adoCol        As Object
    Dim idx             As Integer
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZKNV(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZKNV
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With ZKNV
            Select Case adoCol.Name
                Case "区分":              If Not IsNull(adoCol.Value) Then .区分 = adoCol.Value
                Case "編集契約先":        If Not IsNull(adoCol.Value) Then .編集契約先 = adoCol.Value
                Case "契約先":            If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
                Case "契約先名":          If Not IsNull(adoCol.Value) Then .契約先名 = adoCol.Value
                Case "予測金額":          If Not IsNull(adoCol.Value) Then .予測金額 = adoCol.Value
                Case "日当り数":          If Not IsNull(adoCol.Value) Then .日当り数 = adoCol.Value
                Case "実績金額":          If Not IsNull(adoCol.Value) Then .実績金額 = adoCol.Value
                Case "進捗":              If Not IsNull(adoCol.Value) Then .進捗 = adoCol.Value
                Case "年月":              If Not IsNull(adoCol.Value) Then .年月 = adoCol.Value
                Case "自社工区":          If Not IsNull(adoCol.Value) Then .自社工区 = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
                Case Else
                   For idx = 1 To 31
                     If "実績金額" & idx = adoCol.Name Then
                        .実績金額T(idx) = adoCol.Value
                        Exit For
                     End If
                   Next idx
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZKNVRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZKNVRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 材料買掛入荷状況管理ビュー"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZKNVRead = True
           Call ItemsFromZKNV(DBNo)
           ZKNV_RDSTS = True
        Else
           ZKNV_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call ZKNVClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZKNVReadNext(Optional DBNo As Integer = 0)
    If Not ZKNV_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZKNV_RDSTS = False
        Else
           Call ItemsFromZKNV(DBNo)
           ZKNV_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZKNVClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZKNV_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZKNV )
'*----------------------------------------------*'
Public Sub ItemsClearZKNV()
    With ZKNV
        .区分 = ""
        .編集契約先 = ""
        .契約先 = ""
        .契約先名 = ""
        .予測金額 = 0
        .日当り数 = 0
        .実績金額 = 0
        .進捗 = 0
        For idx = 1 To 31
           .実績金額T(idx) = 0
        Next idx
        .年月 = ""
        .自社工区 = ""
        .受入 = ""
        .件数 = 0
    End With
End Sub

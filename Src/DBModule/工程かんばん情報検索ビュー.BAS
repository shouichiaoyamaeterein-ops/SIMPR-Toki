Attribute VB_Name = "工程かんばん情報検索ビュー"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   工程かんばん情報検索ビュー
'*
'*   モジュールID  : 工程かんばん情報検索ビュー.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2010/07/23  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_工程かんばん情報検索ビュー
        品番          As String
        表示品番      As String
        社内背番      As String
        表示社内背番  As String
        品名          As String
        車型          As String
        取引先        As String
        受入          As String
        表示取引先    As String
        表示契約先    As String
        表示納入先    As String
        背番号        As String
        表示背番号    As String
        収容数1       As Long
        SK区分        As Byte
        SK            As String
        表示サイクル  As String
        手配区分      As String
        手配区分名    As String
        品目          As Byte
        品目名        As String
        収容器1       As String
        変更日便      As String
        収容器2       As String
        収容数2       As Long
        納入先受入    As String
        置場          As String
        アドレス      As String
        取引先名      As String
        最終発行連番  As Integer
        部品区分名    As String
        件数          As Long
    End Type

    Public KJV As T_工程かんばん情報検索ビュー

    Public KJV_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKJV(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKJV	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KJV
            Select Case adoCol.Name
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "表示品番":      If Not IsNull(adoCol.Value) Then .表示品番     = adoCol.Value
                Case "社内背番":      If Not IsNull(adoCol.Value) Then .社内背番     = adoCol.Value
                Case "表示社内背番":  If Not IsNull(adoCol.Value) Then .表示社内背番 = adoCol.Value
                Case "品名":          If Not IsNull(adoCol.Value) Then .品名         = adoCol.Value
                Case "車型":          If Not IsNull(adoCol.Value) Then .車型         = adoCol.Value
                Case "取引先":        If Not IsNull(adoCol.Value) Then .取引先       = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "表示取引先":    If Not IsNull(adoCol.Value) Then .表示取引先   = adoCol.Value
                Case "表示契約先":    If Not IsNull(adoCol.Value) Then .表示契約先   = adoCol.Value
                Case "表示納入先":    If Not IsNull(adoCol.Value) Then .表示納入先   = adoCol.Value
                Case "背番号":        If Not IsNull(adoCol.Value) Then .背番号       = adoCol.Value
                Case "表示背番号":    If Not IsNull(adoCol.Value) Then .表示背番号   = adoCol.Value
                Case "収容数1":       If Not IsNull(adoCol.Value) Then .収容数1      = adoCol.Value
                Case "SK区分":        If Not IsNull(adoCol.Value) Then .SK区分       = adoCol.Value
                Case "SK":            If Not IsNull(adoCol.Value) Then .SK           = adoCol.Value
                Case "表示サイクル":  If Not IsNull(adoCol.Value) Then .表示サイクル = adoCol.Value
                Case "手配区分":      If Not IsNull(adoCol.Value) Then .手配区分     = adoCol.Value
                Case "手配区分名":    If Not IsNull(adoCol.Value) Then .手配区分名   = adoCol.Value
                Case "品目":          If Not IsNull(adoCol.Value) Then .品目         = adoCol.Value
                Case "品目名":        If Not IsNull(adoCol.Value) Then .品目名       = adoCol.Value
                Case "収容器1":       If Not IsNull(adoCol.Value) Then .収容器1      = adoCol.Value
                Case "変更日便":      If Not IsNull(adoCol.Value) Then .変更日便     = adoCol.Value
                Case "収容器2":       If Not IsNull(adoCol.Value) Then .収容器2      = adoCol.Value
                Case "収容数2":       If Not IsNull(adoCol.Value) Then .収容数2      = adoCol.Value
                Case "納入先受入":    If Not IsNull(adoCol.Value) Then .納入先受入   = adoCol.Value
                Case "置場":          If Not IsNull(adoCol.Value) Then .置場         = adoCol.Value
                Case "アドレス":      If Not IsNull(adoCol.Value) Then .アドレス     = adoCol.Value
                Case "取引先名":      If Not IsNull(adoCol.Value) Then .取引先名     = adoCol.Value
                Case "最終発行連番":  If Not IsNull(adoCol.Value) Then .最終発行連番 = adoCol.Value
                Case "部品区分名":    If Not IsNull(adoCol.Value) Then .部品区分名   = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KJVRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KJVRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 工程かんばん情報検索ビュー"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KJVRead = True
           Call ItemsFromKJV(DBNo)
	   KJV_RDSTS = True
        else
	   KJV_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KJVClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KJVReadNext(Optional DBNo As Integer = 0)
    If Not KJV_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KJV_RDSTS = False
        Else
           Call ItemsFromKJV(DBNo)
           KJV_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KJVClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KJV_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KJV )
'*----------------------------------------------*'
Public Sub ItemsClearKJV()
    With KJV
	.品番 = ""
	.表示品番 = ""
	.社内背番 = ""
	.表示社内背番 = ""
	.品名 = ""
	.車型 = ""
	.取引先 = ""
	.受入 = ""
	.表示取引先 = ""
	.表示契約先 = ""
	.表示納入先 = ""
	.背番号 = ""
	.表示背番号 = ""
	.収容数1 = 0
	.SK区分 = 0
	.SK = ""
	.表示サイクル = ""
	.手配区分 = ""
	.手配区分名 = ""
	.品目 = 0
	.品目名 = ""
	.収容器1 = ""
	.変更日便 = ""
	.収容器2 = ""
	.収容数2 = 0
	.納入先受入 = ""
	.置場 = ""
	.アドレス = ""
	.取引先名 = ""
	.最終発行連番 = 0
	.部品区分名 = ""
        .件数 = 0
    End With
End Sub

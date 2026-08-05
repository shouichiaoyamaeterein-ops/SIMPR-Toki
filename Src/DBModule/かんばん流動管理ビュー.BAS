Attribute VB_Name = "かんばん流動管理ビュー"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   かんばん流動管理ビュー
'*
'*   モジュールID  : かんばん流動管理ビュー.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2006/02/19  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_かんばん流動管理ビュー
        取引先    As String
        取引先名  As String
        受入      As String
        品番      As String
        件数          As Long
'   項目追加(自動)
	収容数         As Long
	当月内示数     As Long
	前回回転枚数   As Integer
	今回回転枚数   As Integer
	有効枚数       As Integer
	増減区分       As Byte
	設定枚数       As Integer
	便当枚数       As Integer
	平準化区分     As Byte
	平準設定枚数   As Integer
	背番号         As String
	表示品番       As String
	累積枚数       As Integer
	開始日         As String
	終了日         As String
	完了日         As String
	増減設定内容区分	 As Byte
	平準設定内容区分	 As Byte
	データ更新日   As Date
	設定内容区分   As Byte
    End Type

    Public KRV As T_かんばん流動管理ビュー

    Public KRV_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKRV(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKRV	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KRV
            Select Case adoCol.Name
                Case "取引先":    If Not IsNull(adoCol.Value) Then .取引先   = adoCol.Value
                Case "取引先名":  If Not IsNull(adoCol.Value) Then .取引先名 = adoCol.Value
                Case "受入":      If Not IsNull(adoCol.Value) Then .受入     = adoCol.Value
                Case "品番":      If Not IsNull(adoCol.Value) Then .品番     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "収容数":	If Not IsNull(adoCol.Value) Then .収容数	= adoCol.Value
		Case "当月内示数":	If Not IsNull(adoCol.Value) Then .当月内示数	= adoCol.Value
		Case "前回回転枚数":	If Not IsNull(adoCol.Value) Then .前回回転枚数	= adoCol.Value
		Case "今回回転枚数":	If Not IsNull(adoCol.Value) Then .今回回転枚数	= adoCol.Value
		Case "有効枚数":	If Not IsNull(adoCol.Value) Then .有効枚数	= adoCol.Value
		Case "増減区分":	If Not IsNull(adoCol.Value) Then .増減区分	= adoCol.Value
		Case "設定枚数":	If Not IsNull(adoCol.Value) Then .設定枚数	= adoCol.Value
		Case "便当枚数":	If Not IsNull(adoCol.Value) Then .便当枚数	= adoCol.Value
		Case "平準化区分":	If Not IsNull(adoCol.Value) Then .平準化区分	= adoCol.Value
		Case "平準設定枚数":	If Not IsNull(adoCol.Value) Then .平準設定枚数	= adoCol.Value
		Case "背番号":	If Not IsNull(adoCol.Value) Then .背番号	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "累積枚数":	If Not IsNull(adoCol.Value) Then .累積枚数	= adoCol.Value
		Case "開始日":	If Not IsNull(adoCol.Value) Then .開始日	= adoCol.Value
		Case "終了日":	If Not IsNull(adoCol.Value) Then .終了日	= adoCol.Value
		Case "完了日":	If Not IsNull(adoCol.Value) Then .完了日	= adoCol.Value
		Case "増減設定内容区分":	If Not IsNull(adoCol.Value) Then .増減設定内容区分	= adoCol.Value
		Case "平準設定内容区分":	If Not IsNull(adoCol.Value) Then .平準設定内容区分	= adoCol.Value
		Case "データ更新日":	If Not IsNull(adoCol.Value) Then .データ更新日	= adoCol.Value
		Case "設定内容区分":	If Not IsNull(adoCol.Value) Then .設定内容区分	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KRVRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KRVRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From かんばん流動管理ビュー"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KRVRead = True
           Call ItemsFromKRV(DBNo)
	   KRV_RDSTS = True
        else
	   KRV_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KRVClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KRVReadNext(Optional DBNo As Integer = 0)
    If Not KRV_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KRV_RDSTS = False
        Else
           Call ItemsFromKRV(DBNo)
           KRV_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KRVClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KRV_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KRV )
'*----------------------------------------------*'
Public Sub ItemsClearKRV()
    With KRV
	.取引先 = ""
	.取引先名 = ""
	.受入 = ""
	.品番 = ""
        .件数 = 0
'   項目追加(自動)
	.収容数 = 0
	.当月内示数 = 0
	.前回回転枚数 = 0
	.今回回転枚数 = 0
	.有効枚数 = 0
	.増減区分 = 0
	.設定枚数 = 0
	.便当枚数 = 0
	.平準化区分 = 0
	.平準設定枚数 = 0
	.背番号 = ""
	.表示品番 = ""
	.累積枚数 = 0
	.開始日 = ""
	.終了日 = ""
	.完了日 = ""
	.増減設定内容区分 = 0
	.平準設定内容区分 = 0
	.データ更新日 = 0
	.設定内容区分 = 0
    End With
End Sub

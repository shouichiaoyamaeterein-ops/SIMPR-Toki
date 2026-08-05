Attribute VB_Name = "平準化設定更新エラーテーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   平準化設定更新エラーテーブル
'*
'*   モジュールID  : 平準化設定更新エラーテーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/03/15  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_平準化設定更新エラーテーブル
        取引先        As String
        受入          As String
        品番          As String
        開始日        As String
        終了日        As String
        データ区分    As Byte
        設定内容区分  As Byte
        平準化区分    As Byte
        便上限枚数    As Long
        日上限枚数    As Long
        月上限枚数    As Long
'        作成日        As Date
'        作成端末      As String
        件数          As Long
'   項目追加(自動)
	取引先名      As String
    End Type

    Public HUE As T_平準化設定更新エラーテーブル

    Public HUE_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHUE(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHUE	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With HUE
            Select Case adoCol.Name
                Case "取引先":        If Not IsNull(adoCol.Value) Then .取引先       = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "開始日":        If Not IsNull(adoCol.Value) Then .開始日       = adoCol.Value
                Case "終了日":        If Not IsNull(adoCol.Value) Then .終了日       = adoCol.Value
                Case "データ区分":    If Not IsNull(adoCol.Value) Then .データ区分   = adoCol.Value
                Case "設定内容区分":  If Not IsNull(adoCol.Value) Then .設定内容区分 = adoCol.Value
                Case "平準化区分":    If Not IsNull(adoCol.Value) Then .平準化区分   = adoCol.Value
                Case "便上限枚数":    If Not IsNull(adoCol.Value) Then .便上限枚数   = adoCol.Value
                Case "日上限枚数":    If Not IsNull(adoCol.Value) Then .日上限枚数   = adoCol.Value
                Case "月上限枚数":    If Not IsNull(adoCol.Value) Then .月上限枚数   = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "取引先名":	If Not IsNull(adoCol.Value) Then .取引先名	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function HUERead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HUERead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 平準化設定更新エラーテーブル"
        RTxtSql = RTxtSql &" Where 取引先 = '" & RTrim(HUE.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(HUE.受入) & "'"
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(HUE.品番) & "'"
        RTxtSql = RTxtSql &" And   開始日 = '" & RTrim(HUE.開始日) & "'"
        RTxtSql = RTxtSql &" And   データ区分 = " & HUE.データ区分 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HUERead = True
           Call ItemsFromHUE(DBNo)
	   HUE_RDSTS = True
        else
	   HUE_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call HUEClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HUEReadNext(Optional DBNo As Integer = 0)
    If Not HUE_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HUE_RDSTS = False
        Else
           Call ItemsFromHUE(DBNo)
           HUE_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HUEClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HUE_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HUE )
'*----------------------------------------------*'
Public Sub ItemsClearHUE()
    With HUE
	.取引先 = ""
	.受入 = ""
	.品番 = ""
	.開始日 = ""
	.終了日 = ""
	.データ区分 = 0
	.設定内容区分 = 0
	.平準化区分 = 0
	.便上限枚数 = 0
	.日上限枚数 = 0
	.月上限枚数 = 0
'	.作成日 = 0
'	.作成端末 = ""
        .件数 = 0
'   項目追加(自動)
	.取引先名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function HUEInsert(Optional DBNo As Integer) As Boolean
    HUEInsert = True
    With HUE
        gSL_Select = "Insert Into 平準化設定更新エラーテーブル  ( "
        gSL_Select = gSL_Select & "取引先,受入,品番,"
        gSL_Select = gSL_Select & "開始日,終了日,データ区分,"
        gSL_Select = gSL_Select & "設定内容区分,平準化区分,便上限枚数,"
        gSL_Select = gSL_Select & "日上限枚数,月上限枚数,"
        gSL_Select = gSL_Select & "作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.開始日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.終了日) & "',"		
        gSL_Select = gSL_Select & "" & .データ区分 & ","		
        gSL_Select = gSL_Select & "" & .設定内容区分 & ","		
        gSL_Select = gSL_Select & "" & .平準化区分 & ","		
        gSL_Select = gSL_Select & "" & .便上限枚数 & ","		
        gSL_Select = gSL_Select & "" & .日上限枚数 & ","		
        gSL_Select = gSL_Select & "" & .月上限枚数 & ","		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HUEInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function HUEUpdate(Optional DBNo As Integer) As Boolean
    HUEUpdate = False	
    With HUE
        gSL_Select = "Update 平準化設定更新エラーテーブル  Set "
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "開始日 = '" & RTrim(.開始日) & "',"
        gSL_Select = gSL_Select & "終了日 = '" & RTrim(.終了日) & "',"
        gSL_Select = gSL_Select & "データ区分 = " & .データ区分 & ","
        gSL_Select = gSL_Select & "設定内容区分 = " & .設定内容区分 & ","
        gSL_Select = gSL_Select & "平準化区分 = " & .平準化区分 & ","
        gSL_Select = gSL_Select & "便上限枚数 = " & .便上限枚数 & ","
        gSL_Select = gSL_Select & "日上限枚数 = " & .日上限枚数 & ","
        gSL_Select = gSL_Select & "月上限枚数 = " & .月上限枚数 & ""
        gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 開始日 = '" & RTrim(.開始日) & "'"
        gSL_Select = gSL_Select & "   And データ区分 = " & .データ区分 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HUEUpdate = True
End Function

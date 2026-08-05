Attribute VB_Name = "かんばん平準化テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   かんばん平準化テーブル
'*
'*   モジュールID  : かんばん平準化テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/01/26  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_かんばん平準化テーブル
        取引先        As String
        受入          As String
        品番          As String
        開始日        As String
        終了日        As String
        設定内容区分  As Byte
        平準化区分    As Byte
        便上限枚数    As Long
        日上限枚数    As Long
        月上限枚数    As Long
'        作成日        As Date
        作成者        As String
'        作成端末      As String
'        更新日        As Date
        更新者        As String
'        更新端末      As String
        件数          As Long
    End Type

    Public KJT As T_かんばん平準化テーブル

    Public KJT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKJT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKJT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KJT
            Select Case adoCol.Name
                Case "取引先":        If Not IsNull(adoCol.Value) Then .取引先       = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "開始日":        If Not IsNull(adoCol.Value) Then .開始日       = adoCol.Value
                Case "終了日":        If Not IsNull(adoCol.Value) Then .終了日       = adoCol.Value
                Case "設定内容区分":  If Not IsNull(adoCol.Value) Then .設定内容区分 = adoCol.Value
                Case "平準化区分":    If Not IsNull(adoCol.Value) Then .平準化区分   = adoCol.Value
                Case "便上限枚数":    If Not IsNull(adoCol.Value) Then .便上限枚数   = adoCol.Value
                Case "日上限枚数":    If Not IsNull(adoCol.Value) Then .日上限枚数   = adoCol.Value
                Case "月上限枚数":    If Not IsNull(adoCol.Value) Then .月上限枚数   = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者       = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KJTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KJTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From かんばん平準化テーブル"
        RTxtSql = RTxtSql &" Where 取引先 = '" & RTrim(KJT.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(KJT.受入) & "'"
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(KJT.品番) & "'"
        RTxtSql = RTxtSql &" And   開始日 = '" & RTrim(KJT.開始日) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KJTRead = True
           Call ItemsFromKJT(DBNo)
	   KJT_RDSTS = True
        else
	   KJT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KJTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KJTReadNext(Optional DBNo As Integer = 0)
    If Not KJT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KJT_RDSTS = False
        Else
           Call ItemsFromKJT(DBNo)
           KJT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KJTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KJT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KJT )
'*----------------------------------------------*'
Public Sub ItemsClearKJT()
    With KJT
	.取引先 = ""
	.受入 = ""
	.品番 = ""
	.開始日 = ""
	.終了日 = ""
	.設定内容区分 = 0
	.平準化区分 = 0
	.便上限枚数 = 0
	.日上限枚数 = 0
	.月上限枚数 = 0
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
'	.更新日 = 0
	.更新者 = ""
'	.更新端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KJTInsert(Optional DBNo As Integer) As Boolean
    KJTInsert = True
    With KJT
        gSL_Select = "Insert Into かんばん平準化テーブル  ( "
        gSL_Select = gSL_Select & "取引先,受入,品番,"
        gSL_Select = gSL_Select & "開始日,終了日,設定内容区分,"
        gSL_Select = gSL_Select & "平準化区分,便上限枚数,日上限枚数,"
        gSL_Select = gSL_Select & "月上限枚数,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.開始日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.終了日) & "',"		
        gSL_Select = gSL_Select & "" & .設定内容区分 & ","		
        gSL_Select = gSL_Select & "" & .平準化区分 & ","		
        gSL_Select = gSL_Select & "" & .便上限枚数 & ","		
        gSL_Select = gSL_Select & "" & .日上限枚数 & ","		
        gSL_Select = gSL_Select & "" & .月上限枚数 & ","		
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
    KJTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KJTUpdate(Optional DBNo As Integer) As Boolean
    KJTUpdate = False	
    With KJT
        gSL_Select = "Update かんばん平準化テーブル  Set "
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "開始日 = '" & RTrim(.開始日) & "',"
        gSL_Select = gSL_Select & "終了日 = '" & RTrim(.終了日) & "',"
        gSL_Select = gSL_Select & "設定内容区分 = " & .設定内容区分 & ","
        gSL_Select = gSL_Select & "平準化区分 = " & .平準化区分 & ","
        gSL_Select = gSL_Select & "便上限枚数 = " & .便上限枚数 & ","
        gSL_Select = gSL_Select & "日上限枚数 = " & .日上限枚数 & ","
        gSL_Select = gSL_Select & "月上限枚数 = " & .月上限枚数 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 開始日 = '" & RTrim(.開始日) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KJTUpdate = True
End Function

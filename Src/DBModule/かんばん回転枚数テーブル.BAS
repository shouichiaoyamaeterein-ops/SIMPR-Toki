Attribute VB_Name = "かんばん回転枚数テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   かんばん回転枚数テーブル
'*
'*   モジュールID  : かんばん回転枚数テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/03/17  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_かんばん回転枚数テーブル
        年月          As String
        仕入先        As String
        品番          As String
        受入          As String
        背番号        As String
        回転枚数      As Long
        前回回転枚数  As Long
        回転枚数差    As Long
        増発枚数      As Long
        抜取枚数      As Long
        増減区分      As Byte
        前月回転枚数  As Long
        平準化区分    As Byte
        便上限枚数    As Long
        日上限枚数    As Long
        月上限枚数    As Long
'        作成日        As Date
        作成者        As String
'        作成端末      As String
        件数          As Long
'   項目追加(自動)
	仕入先名      As String
	表示品番      As String
    End Type

    Public KKT As T_かんばん回転枚数テーブル

    Public KKT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKKT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKKT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KKT
            Select Case adoCol.Name
                Case "年月":          If Not IsNull(adoCol.Value) Then .年月         = adoCol.Value
                Case "仕入先":        If Not IsNull(adoCol.Value) Then .仕入先       = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "背番号":        If Not IsNull(adoCol.Value) Then .背番号       = adoCol.Value
                Case "回転枚数":      If Not IsNull(adoCol.Value) Then .回転枚数     = adoCol.Value
                Case "前回回転枚数":  If Not IsNull(adoCol.Value) Then .前回回転枚数 = adoCol.Value
                Case "回転枚数差":    If Not IsNull(adoCol.Value) Then .回転枚数差   = adoCol.Value
                Case "増発枚数":      If Not IsNull(adoCol.Value) Then .増発枚数     = adoCol.Value
                Case "抜取枚数":      If Not IsNull(adoCol.Value) Then .抜取枚数     = adoCol.Value
                Case "増減区分":      If Not IsNull(adoCol.Value) Then .増減区分     = adoCol.Value
                Case "前月回転枚数":  If Not IsNull(adoCol.Value) Then .前月回転枚数 = adoCol.Value
                Case "平準化区分":    If Not IsNull(adoCol.Value) Then .平準化区分   = adoCol.Value
                Case "便上限枚数":    If Not IsNull(adoCol.Value) Then .便上限枚数   = adoCol.Value
                Case "日上限枚数":    If Not IsNull(adoCol.Value) Then .日上限枚数   = adoCol.Value
                Case "月上限枚数":    If Not IsNull(adoCol.Value) Then .月上限枚数   = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "仕入先名":	If Not IsNull(adoCol.Value) Then .仕入先名	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KKTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KKTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From かんばん回転枚数テーブル"
        RTxtSql = RTxtSql &" Where 年月 = '" & RTrim(KKT.年月) & "'"
        RTxtSql = RTxtSql &" And   仕入先 = '" & RTrim(KKT.仕入先) & "'"
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(KKT.品番) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(KKT.受入) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KKTRead = True
           Call ItemsFromKKT(DBNo)
	   KKT_RDSTS = True
        else
	   KKT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KKTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KKTReadNext(Optional DBNo As Integer = 0)
    If Not KKT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KKT_RDSTS = False
        Else
           Call ItemsFromKKT(DBNo)
           KKT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KKTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KKT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KKT )
'*----------------------------------------------*'
Public Sub ItemsClearKKT()
    With KKT
	.年月 = ""
	.仕入先 = ""
	.品番 = ""
	.受入 = ""
	.背番号 = ""
	.回転枚数 = 0
	.前回回転枚数 = 0
	.回転枚数差 = 0
	.増発枚数 = 0
	.抜取枚数 = 0
	.増減区分 = 0
	.前月回転枚数 = 0
	.平準化区分 = 0
	.便上限枚数 = 0
	.日上限枚数 = 0
	.月上限枚数 = 0
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
        .件数 = 0
'   項目追加(自動)
	.仕入先名 = ""
	.表示品番 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KKTInsert(Optional DBNo As Integer) As Boolean
    KKTInsert = True
    With KKT
        gSL_Select = "Insert Into かんばん回転枚数テーブル  ( "
        gSL_Select = gSL_Select & "年月,仕入先,品番,"
        gSL_Select = gSL_Select & "受入,背番号,回転枚数,"
        gSL_Select = gSL_Select & "前回回転枚数,回転枚数差,増発枚数,"
        gSL_Select = gSL_Select & "抜取枚数,増減区分,前月回転枚数,"
        gSL_Select = gSL_Select & "平準化区分,便上限枚数,日上限枚数,"
        gSL_Select = gSL_Select & "月上限枚数,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "" & .回転枚数 & ","		
        gSL_Select = gSL_Select & "" & .前回回転枚数 & ","		
        gSL_Select = gSL_Select & "" & .回転枚数差 & ","		
        gSL_Select = gSL_Select & "" & .増発枚数 & ","		
        gSL_Select = gSL_Select & "" & .抜取枚数 & ","		
        gSL_Select = gSL_Select & "" & .増減区分 & ","		
        gSL_Select = gSL_Select & "" & .前月回転枚数 & ","		
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
    KKTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KKTUpdate(Optional DBNo As Integer) As Boolean
    KKTUpdate = False	
    With KKT
        gSL_Select = "Update かんばん回転枚数テーブル  Set "
        gSL_Select = gSL_Select & "年月 = '" & RTrim(.年月) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "回転枚数 = " & .回転枚数 & ","
        gSL_Select = gSL_Select & "前回回転枚数 = " & .前回回転枚数 & ","
        gSL_Select = gSL_Select & "回転枚数差 = " & .回転枚数差 & ","
        gSL_Select = gSL_Select & "増発枚数 = " & .増発枚数 & ","
        gSL_Select = gSL_Select & "抜取枚数 = " & .抜取枚数 & ","
        gSL_Select = gSL_Select & "増減区分 = " & .増減区分 & ","
        gSL_Select = gSL_Select & "前月回転枚数 = " & .前月回転枚数 & ","
        gSL_Select = gSL_Select & "平準化区分 = " & .平準化区分 & ","
        gSL_Select = gSL_Select & "便上限枚数 = " & .便上限枚数 & ","
        gSL_Select = gSL_Select & "日上限枚数 = " & .日上限枚数 & ","
        gSL_Select = gSL_Select & "月上限枚数 = " & .月上限枚数 & ""
        gSL_Select = gSL_Select & " Where 年月 = '" & RTrim(.年月) & "'"
        gSL_Select = gSL_Select & "   And 仕入先 = '" & RTrim(.仕入先) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KKTUpdate = True
End Function

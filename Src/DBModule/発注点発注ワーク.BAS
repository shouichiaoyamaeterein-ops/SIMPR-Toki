Attribute VB_Name = "発注点発注ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   発注点発注ワーク
'*
'*   モジュールID  : 発注点発注ワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/11/22  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_発注点発注ワーク
        処理端末          As String
        品番              As String
        仕入先            As String
        受入              As String
        現在庫数          As Long
        発注予定数        As Long
        所要量            As Long
'        作成日            As Date
        件数          As Long
'   項目追加(自動)
	設定数        As Long
	在庫管理区分  As Byte
	在庫係数      As Currency
	在庫日数      As Currency
	在庫発注点    As Long
	在庫管理名    As String
	表示品番      As String
	収容数1       As Long
	発注ロット1   As Long
	切替日        As String
	切替便        As String
	収容数2       As Long
	発注ロット2   As Long
	Week1         As Byte
	Week2         As Byte
	Week3         As Byte
	Week4         As Byte
	Week5         As Byte
	Day1          As Byte
	Day2          As Byte
	Day3          As Byte
	Day4          As Byte
	Day5          As Byte
	背番号        As String
	仕入先名      As String
	担当          As String
	担当名        As String
	手配区分      As String
	SK区分        As Byte
    End Type

    Public HHW As T_発注点発注ワーク

    Public HHW_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHHW(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHHW	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With HHW
            Select Case adoCol.Name
                Case "処理端末":          If Not IsNull(adoCol.Value) Then .処理端末         = adoCol.Value
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番             = adoCol.Value
                Case "仕入先":            If Not IsNull(adoCol.Value) Then .仕入先           = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入             = adoCol.Value
                Case "現在庫数":          If Not IsNull(adoCol.Value) Then .現在庫数         = adoCol.Value
                Case "発注予定数":        If Not IsNull(adoCol.Value) Then .発注予定数       = adoCol.Value
                Case "所要量":            If Not IsNull(adoCol.Value) Then .所要量           = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "設定数":	If Not IsNull(adoCol.Value) Then .設定数	= adoCol.Value
		Case "在庫管理区分":	If Not IsNull(adoCol.Value) Then .在庫管理区分	= adoCol.Value
		Case "在庫係数":	If Not IsNull(adoCol.Value) Then .在庫係数	= adoCol.Value
		Case "在庫日数":	If Not IsNull(adoCol.Value) Then .在庫日数	= adoCol.Value
		Case "在庫発注点":	If Not IsNull(adoCol.Value) Then .在庫発注点	= adoCol.Value
		Case "在庫管理名":	If Not IsNull(adoCol.Value) Then .在庫管理名	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "収容数1":	If Not IsNull(adoCol.Value) Then .収容数1	= adoCol.Value
		Case "発注ロット1":	If Not IsNull(adoCol.Value) Then .発注ロット1	= adoCol.Value
		Case "切替日":	If Not IsNull(adoCol.Value) Then .切替日	= adoCol.Value
		Case "切替便":	If Not IsNull(adoCol.Value) Then .切替便	= adoCol.Value
		Case "収容数2":	If Not IsNull(adoCol.Value) Then .収容数2	= adoCol.Value
		Case "発注ロット2":	If Not IsNull(adoCol.Value) Then .発注ロット2	= adoCol.Value
		Case "Week1":	If Not IsNull(adoCol.Value) Then .Week1	= adoCol.Value
		Case "Week2":	If Not IsNull(adoCol.Value) Then .Week2	= adoCol.Value
		Case "Week3":	If Not IsNull(adoCol.Value) Then .Week3	= adoCol.Value
		Case "Week4":	If Not IsNull(adoCol.Value) Then .Week4	= adoCol.Value
		Case "Week5":	If Not IsNull(adoCol.Value) Then .Week5	= adoCol.Value
		Case "Day1":	If Not IsNull(adoCol.Value) Then .Day1	= adoCol.Value
		Case "Day2":	If Not IsNull(adoCol.Value) Then .Day2	= adoCol.Value
		Case "Day3":	If Not IsNull(adoCol.Value) Then .Day3	= adoCol.Value
		Case "Day4":	If Not IsNull(adoCol.Value) Then .Day4	= adoCol.Value
		Case "Day5":	If Not IsNull(adoCol.Value) Then .Day5	= adoCol.Value
		Case "背番号":	If Not IsNull(adoCol.Value) Then .背番号	= adoCol.Value
		Case "仕入先名":	If Not IsNull(adoCol.Value) Then .仕入先名	= adoCol.Value
		Case "担当":	If Not IsNull(adoCol.Value) Then .担当	= adoCol.Value
		Case "担当名":	If Not IsNull(adoCol.Value) Then .担当名	= adoCol.Value
		Case "手配区分":	If Not IsNull(adoCol.Value) Then .手配区分	= adoCol.Value
		Case "SK区分":	If Not IsNull(adoCol.Value) Then .SK区分	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function HHWRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HHWRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 発注点発注ワーク"
        RTxtSql = RTxtSql &" Where 処理端末 = '" & RTrim(HHW.処理端末) & "'"
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(HHW.品番) & "'"
        RTxtSql = RTxtSql &" And   仕入先 = '" & RTrim(HHW.仕入先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(HHW.受入) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HHWRead = True
           Call ItemsFromHHW(DBNo)
	   HHW_RDSTS = True
        else
	   HHW_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call HHWClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HHWReadNext(Optional DBNo As Integer = 0)
    If Not HHW_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HHW_RDSTS = False
        Else
           Call ItemsFromHHW(DBNo)
           HHW_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HHWClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HHW_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HHW )
'*----------------------------------------------*'
Public Sub ItemsClearHHW()
    With HHW
	.処理端末 = ""
	.品番 = ""
	.仕入先 = ""
	.受入 = ""
	.現在庫数 = 0
	.発注予定数 = 0
	.所要量 = 0
'	.作成日 = 0
        .件数 = 0
'   項目追加(自動)
	.設定数 = 0
	.在庫管理区分 = 0
	.在庫係数 = 0
	.在庫日数 = 0
	.在庫発注点 = 0
	.在庫管理名 = ""
	.表示品番 = ""
	.収容数1 = 0
	.発注ロット1 = 0
	.切替日 = ""
	.切替便 = ""
	.収容数2 = 0
	.発注ロット2 = 0
	.Week1 = 0
	.Week2 = 0
	.Week3 = 0
	.Week4 = 0
	.Week5 = 0
	.Day1 = 0
	.Day2 = 0
	.Day3 = 0
	.Day4 = 0
	.Day5 = 0
	.背番号 = ""
	.仕入先名 = ""
	.担当 = ""
	.担当名 = ""
	.手配区分 = ""
	.SK区分 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function HHWInsert(Optional DBNo As Integer) As Boolean
    HHWInsert = True
    With HHW
        gSL_Select = "Insert Into 発注点発注ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,品番,仕入先,"
        gSL_Select = gSL_Select & "受入,現在庫数,発注予定数,"
        gSL_Select = gSL_Select & "所要量,"
        gSL_Select = gSL_Select & "作成日"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "" & .現在庫数 & ","		
        gSL_Select = gSL_Select & "" & .発注予定数 & ","		
        gSL_Select = gSL_Select & "" & .所要量 & ","		
        gSL_Select = gSL_Select & "GETDATE()"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HHWInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function HHWUpdate(Optional DBNo As Integer) As Boolean
    HHWUpdate = False	
    With HHW
        gSL_Select = "Update 発注点発注ワーク  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "現在庫数 = " & .現在庫数 & ","
        gSL_Select = gSL_Select & "発注予定数 = " & .発注予定数 & ","
        gSL_Select = gSL_Select & "所要量 = " & .所要量 & ""
        gSL_Select = gSL_Select & " Where 処理端末 = '" & RTrim(.処理端末) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 仕入先 = '" & RTrim(.仕入先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HHWUpdate = True
End Function

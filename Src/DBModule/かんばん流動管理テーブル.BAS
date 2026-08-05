Attribute VB_Name = "かんばん流動管理テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   かんばん流動管理テーブル
'*
'*   モジュールID  : かんばん流動管理テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2005/06/15  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_かんばん流動管理テーブル
        取引先            As String
        受入              As String
        品番              As String
        発行NO            As Integer
        流動日付          As String
        発行区分          As Byte
        版数              As Byte
        有効指示日        As String
        無効指示日        As String
        回収日            As String
        収容区分          As Byte
'        作成日            As Date
        作成者            As String
'        作成端末          As String
'        更新日            As Date
        更新者            As String
'        更新端末          As String
        件数          As Long
'   項目追加(自動)
	未流件数      As Long
	有効件数      As Long
	無効件数      As Long
	滞留件数      As Long
	取引先名      As String
	背番号        As String
	サイクル      As String
	変更日        As String
	収容数1       As Long
	収容数2       As Long
	回転枚数      As Long
	変更便        As String
	置場          As String
	収容器1       As String
	収容器2       As String
	前月回転枚数  As Long
    End Type

    Public KRT As T_かんばん流動管理テーブル

    Public KRT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKRT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKRT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KRT
            Select Case adoCol.Name
                Case "取引先":            If Not IsNull(adoCol.Value) Then .取引先           = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入             = adoCol.Value
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番             = adoCol.Value
                Case "発行NO":            If Not IsNull(adoCol.Value) Then .発行NO           = adoCol.Value
                Case "流動日付":          If Not IsNull(adoCol.Value) Then .流動日付         = adoCol.Value
                Case "発行区分":          If Not IsNull(adoCol.Value) Then .発行区分         = adoCol.Value
                Case "版数":              If Not IsNull(adoCol.Value) Then .版数             = adoCol.Value
                Case "有効指示日":        If Not IsNull(adoCol.Value) Then .有効指示日       = adoCol.Value
                Case "無効指示日":        If Not IsNull(adoCol.Value) Then .無効指示日       = adoCol.Value
                Case "回収日":            If Not IsNull(adoCol.Value) Then .回収日           = adoCol.Value
                Case "収容区分":          If Not IsNull(adoCol.Value) Then .収容区分         = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
'                Case "更新日":            If Not IsNull(adoCol.Value) Then .更新日           = adoCol.Value
                Case "更新者":            If Not IsNull(adoCol.Value) Then .更新者           = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "未流件数":	If Not IsNull(adoCol.Value) Then .未流件数	= adoCol.Value
		Case "有効件数":	If Not IsNull(adoCol.Value) Then .有効件数	= adoCol.Value
		Case "無効件数":	If Not IsNull(adoCol.Value) Then .無効件数	= adoCol.Value
		Case "滞留件数":	If Not IsNull(adoCol.Value) Then .滞留件数	= adoCol.Value
		Case "取引先名":	If Not IsNull(adoCol.Value) Then .取引先名	= adoCol.Value
		Case "背番号":	If Not IsNull(adoCol.Value) Then .背番号	= adoCol.Value
		Case "サイクル":	If Not IsNull(adoCol.Value) Then .サイクル	= adoCol.Value
		Case "変更日":	If Not IsNull(adoCol.Value) Then .変更日	= adoCol.Value
		Case "収容数1":	If Not IsNull(adoCol.Value) Then .収容数1	= adoCol.Value
		Case "収容数2":	If Not IsNull(adoCol.Value) Then .収容数2	= adoCol.Value
		Case "回転枚数":	If Not IsNull(adoCol.Value) Then .回転枚数	= adoCol.Value
		Case "変更便":	If Not IsNull(adoCol.Value) Then .変更便	= adoCol.Value
		Case "置場":	If Not IsNull(adoCol.Value) Then .置場	= adoCol.Value
		Case "収容器1":	If Not IsNull(adoCol.Value) Then .収容器1	= adoCol.Value
		Case "収容器2":	If Not IsNull(adoCol.Value) Then .収容器2	= adoCol.Value
		Case "前月回転枚数":	If Not IsNull(adoCol.Value) Then .前月回転枚数	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KRTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KRTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From かんばん流動管理テーブル"
        RTxtSql = RTxtSql &" Where 取引先 = '" & RTrim(KRT.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(KRT.受入) & "'"
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(KRT.品番) & "'"
        RTxtSql = RTxtSql &" And   発行NO = " & KRT.発行NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KRTRead = True
           Call ItemsFromKRT(DBNo)
	   KRT_RDSTS = True
        else
	   KRT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KRTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KRTReadNext(Optional DBNo As Integer = 0)
    If Not KRT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KRT_RDSTS = False
        Else
           Call ItemsFromKRT(DBNo)
           KRT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KRTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KRT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KRT )
'*----------------------------------------------*'
Public Sub ItemsClearKRT()
    With KRT
	.取引先 = ""
	.受入 = ""
	.品番 = ""
	.発行NO = 0
	.流動日付 = ""
	.発行区分 = 0
	.版数 = 0
	.有効指示日 = ""
	.無効指示日 = ""
	.回収日 = ""
	.収容区分 = 0
'	.作成日 = Nothing
	.作成者 = ""
'	.作成端末 = ""
'	.更新日 = Nothing
	.更新者 = ""
'	.更新端末 = ""
        .件数 = 0
'   項目追加(自動)
	.未流件数 = 0
	.有効件数 = 0
	.無効件数 = 0
	.滞留件数 = 0
	.取引先名 = ""
	.背番号 = ""
	.サイクル = ""
	.変更日 = ""
	.収容数1 = 0
	.収容数2 = 0
	.回転枚数 = 0
	.変更便 = ""
	.置場 = ""
	.収容器1 = ""
	.収容器2 = ""
	.前月回転枚数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KRTInsert(Optional DBNo As Integer) As Boolean
    KRTInsert = True
    With KRT
        gSL_Select = "Insert Into かんばん流動管理テーブル  ( "
        gSL_Select = gSL_Select & "取引先,受入,品番,"
        gSL_Select = gSL_Select & "発行NO,流動日付,発行区分,"
        gSL_Select = gSL_Select & "版数,有効指示日,無効指示日,"
        gSL_Select = gSL_Select & "回収日,収容区分,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "" & .発行NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.流動日付) & "',"		
        gSL_Select = gSL_Select & "" & .発行区分 & ","		
        gSL_Select = gSL_Select & "" & .版数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.有効指示日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.無効指示日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.回収日) & "',"		
        gSL_Select = gSL_Select & "" & .収容区分 & ","		
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
    KRTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KRTUpdate(Optional DBNo As Integer) As Boolean
    KRTUpdate = False	
    With KRT
        gSL_Select = "Update かんばん流動管理テーブル  Set "
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "発行NO = " & .発行NO & ","
        gSL_Select = gSL_Select & "流動日付 = '" & RTrim(.流動日付) & "',"
        gSL_Select = gSL_Select & "発行区分 = " & .発行区分 & ","
        gSL_Select = gSL_Select & "版数 = " & .版数 & ","
        gSL_Select = gSL_Select & "有効指示日 = '" & RTrim(.有効指示日) & "',"
        gSL_Select = gSL_Select & "無効指示日 = '" & RTrim(.無効指示日) & "',"
        gSL_Select = gSL_Select & "回収日 = '" & RTrim(.回収日) & "',"
        gSL_Select = gSL_Select & "収容区分 = " & .収容区分 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 発行NO = " & .発行NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KRTUpdate = True
End Function

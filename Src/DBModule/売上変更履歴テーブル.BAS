Attribute VB_Name = "売上変更履歴テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   売上変更履歴テーブル
'*
'*   モジュールID  : 売上変更履歴テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2005/06/09  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_売上変更履歴テーブル
        連番              As Long
        処理区分          As String
        削除区分          As String
        売上年月          As String
        日付              As String
        便                As String
        契約先            As String
        納入先            As String
        受入              As String
        処理件数          As Long
        変更前年月        As String
        作成元区分        As String
'        作成日            As Date
'        作成端末          As String
        作成者            As String
        件数          As Long
'   項目追加(自動)
	契約先名      As String
	納入先名      As String
	処理日時      As String
    End Type

    Public UHR As T_売上変更履歴テーブル

    Public UHR_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromUHR(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearUHR	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With UHR
            Select Case adoCol.Name
                Case "連番":              If Not IsNull(adoCol.Value) Then .連番             = adoCol.Value
                Case "処理区分":          If Not IsNull(adoCol.Value) Then .処理区分         = adoCol.Value
                Case "削除区分":          If Not IsNull(adoCol.Value) Then .削除区分         = adoCol.Value
                Case "売上年月":          If Not IsNull(adoCol.Value) Then .売上年月         = adoCol.Value
                Case "日付":              If Not IsNull(adoCol.Value) Then .日付             = adoCol.Value
                Case "便":                If Not IsNull(adoCol.Value) Then .便               = adoCol.Value
                Case "契約先":            If Not IsNull(adoCol.Value) Then .契約先           = adoCol.Value
                Case "納入先":            If Not IsNull(adoCol.Value) Then .納入先           = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入             = adoCol.Value
                Case "処理件数":          If Not IsNull(adoCol.Value) Then .処理件数         = adoCol.Value
                Case "変更前年月":        If Not IsNull(adoCol.Value) Then .変更前年月       = adoCol.Value
                Case "作成元区分":        If Not IsNull(adoCol.Value) Then .作成元区分       = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者           = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "契約先名":	If Not IsNull(adoCol.Value) Then .契約先名	= adoCol.Value
		Case "納入先名":	If Not IsNull(adoCol.Value) Then .納入先名	= adoCol.Value
		Case "処理日時":	If Not IsNull(adoCol.Value) Then .処理日時	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function UHRRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    UHRRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 売上変更履歴テーブル"
        RTxtSql = RTxtSql &" Where 連番 = " & UHR.連番 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           UHRRead = True
           Call ItemsFromUHR(DBNo)
	   UHR_RDSTS = True
        else
	   UHR_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call UHRClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub UHRReadNext(Optional DBNo As Integer = 0)
    If Not UHR_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           UHR_RDSTS = False
        Else
           Call ItemsFromUHR(DBNo)
           UHR_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub UHRClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        UHR_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( UHR )
'*----------------------------------------------*'
Public Sub ItemsClearUHR()
    With UHR
	.連番 = 0
	.処理区分 = ""
	.削除区分 = ""
	.売上年月 = ""
	.日付 = ""
	.便 = ""
	.契約先 = ""
	.納入先 = ""
	.受入 = ""
	.処理件数 = 0
	.変更前年月 = ""
	.作成元区分 = ""
'	.作成日 = Nothing
'	.作成端末 = ""
	.作成者 = ""
        .件数 = 0
'   項目追加(自動)
	.契約先名 = ""
	.納入先名 = ""
	.処理日時 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function UHRInsert(Optional DBNo As Integer) As Boolean
    UHRInsert = True
    With UHR
        gSL_Select = "Insert Into 売上変更履歴テーブル  ( "
        gSL_Select = gSL_Select & "処理区分,削除区分,売上年月,"
        gSL_Select = gSL_Select & "日付,便,契約先,"
        gSL_Select = gSL_Select & "納入先,受入,処理件数,"
        gSL_Select = gSL_Select & "変更前年月,作成元区分,"
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.削除区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.売上年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.日付) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "" & .処理件数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.変更前年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.作成元区分) & "',"		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20),"		
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    UHRInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function UHRUpdate(Optional DBNo As Integer) As Boolean
    UHRUpdate = False	
    With UHR
        gSL_Select = "Update 売上変更履歴テーブル  Set "
        gSL_Select = gSL_Select & "処理区分 = '" & RTrim(.処理区分) & "',"
        gSL_Select = gSL_Select & "削除区分 = '" & RTrim(.削除区分) & "',"
        gSL_Select = gSL_Select & "売上年月 = '" & RTrim(.売上年月) & "',"
        gSL_Select = gSL_Select & "日付 = '" & RTrim(.日付) & "',"
        gSL_Select = gSL_Select & "便 = '" & RTrim(.便) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "処理件数 = " & .処理件数 & ","
        gSL_Select = gSL_Select & "変更前年月 = '" & RTrim(.変更前年月) & "',"
        gSL_Select = gSL_Select & "作成元区分 = '" & RTrim(.作成元区分) & "'"
        gSL_Select = gSL_Select & " Where 連番 = " & .連番 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    UHRUpdate = True
End Function

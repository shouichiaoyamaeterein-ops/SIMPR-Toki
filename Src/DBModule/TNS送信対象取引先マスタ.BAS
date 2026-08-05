Attribute VB_Name = "TNS送信対象取引先マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TNS送信対象取引先マスタ
'*
'*   モジュールID  : TNS送信対象取引先マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2007/07/20  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_TNS送信対象取引先マスタ
        契約先    As String
        取引先    As String
        受入      As String
        切替日付  As String
        実績区分  As Byte
'        作成日    As Date
        作成者    As String
'        作成端末  As String
        件数          As Long
'   項目追加(自動)
	取引先名       As String
	対象フラグ     As String
	仕入先         As String
    End Type

    Public TST As T_TNS送信対象取引先マスタ

    Public TST_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTST(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTST	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With TST
            Select Case adoCol.Name
                Case "契約先":    If Not IsNull(adoCol.Value) Then .契約先   = adoCol.Value
                Case "取引先":    If Not IsNull(adoCol.Value) Then .取引先   = adoCol.Value
                Case "受入":      If Not IsNull(adoCol.Value) Then .受入     = adoCol.Value
                Case "切替日付":  If Not IsNull(adoCol.Value) Then .切替日付 = adoCol.Value
                Case "実績区分":  If Not IsNull(adoCol.Value) Then .実績区分 = adoCol.Value
'                Case "作成日":    If Not IsNull(adoCol.Value) Then .作成日   = adoCol.Value
                Case "作成者":    If Not IsNull(adoCol.Value) Then .作成者   = adoCol.Value
'                Case "作成端末":  If Not IsNull(adoCol.Value) Then .作成端末 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "取引先名":	If Not IsNull(adoCol.Value) Then .取引先名	= adoCol.Value
		Case "対象フラグ":	If Not IsNull(adoCol.Value) Then .対象フラグ	= adoCol.Value
		Case "仕入先":	If Not IsNull(adoCol.Value) Then .仕入先	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TSTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TSTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From TNS送信対象取引先マスタ"
        RTxtSql = RTxtSql &" Where 契約先 = '" & RTrim(TST.契約先) & "'"
        RTxtSql = RTxtSql &" And   取引先 = '" & RTrim(TST.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(TST.受入) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TSTRead = True
           Call ItemsFromTST(DBNo)
	   TST_RDSTS = True
        else
	   TST_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call TSTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TSTReadNext(Optional DBNo As Integer = 0)
    If Not TST_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TST_RDSTS = False
        Else
           Call ItemsFromTST(DBNo)
           TST_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TSTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TST_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TST )
'*----------------------------------------------*'
Public Sub ItemsClearTST()
    With TST
	.契約先 = ""
	.取引先 = ""
	.受入 = ""
	.切替日付 = ""
	.実績区分 = 0
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
        .件数 = 0
'   項目追加(自動)
	.取引先名 = ""
	.対象フラグ = ""
	.仕入先 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function TSTInsert(Optional DBNo As Integer) As Boolean
    TSTInsert = False
    With TST
        gSL_Select = "Insert Into TNS送信対象取引先マスタ  ( "
        gSL_Select = gSL_Select & "契約先,取引先,受入,"
        gSL_Select = gSL_Select & "切替日付,実績区分,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.切替日付) & "',"		
        gSL_Select = gSL_Select & "" & .実績区分 & ","		
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
    TSTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function TSTUpdate(Optional DBNo As Integer) As Boolean
    TSTUpdate = False	
    With TST
        gSL_Select = "Update TNS送信対象取引先マスタ  Set "
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "切替日付 = '" & RTrim(.切替日付) & "',"
        gSL_Select = gSL_Select & "実績区分 = " & .実績区分 & ""
        gSL_Select = gSL_Select & " Where 契約先 = '" & RTrim(.契約先) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TSTUpdate = True
End Function

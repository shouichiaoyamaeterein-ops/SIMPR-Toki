Attribute VB_Name = "かんばん増減開始日マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   かんばん増減開始日マスタ
'*
'*   モジュールID  : かんばん増減開始日マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2006/02/20  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_かんばん増減開始日マスタ
        取引先        As String
        受入          As String
        品番          As String
        開始日        As String
'        作成日        As Date
        作成者        As String
'        作成端末      As String
'        更新日        As Date
        更新者        As String
'        更新端末      As String
        件数          As Long
'   項目追加(自動)
	取引先名       As String
	表示品番       As String
    End Type

    Public KZS As T_かんばん増減開始日マスタ

    Public KZS_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKZS(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKZS	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KZS
            Select Case adoCol.Name
                Case "取引先":        If Not IsNull(adoCol.Value) Then .取引先       = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "開始日":        If Not IsNull(adoCol.Value) Then .開始日       = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者       = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "取引先名":	If Not IsNull(adoCol.Value) Then .取引先名	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KZSRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KZSRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From かんばん増減開始日マスタ"
        RTxtSql = RTxtSql &" Where 取引先 = '" & RTrim(KZS.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(KZS.受入) & "'"
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(KZS.品番) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KZSRead = True
           Call ItemsFromKZS(DBNo)
	   KZS_RDSTS = True
        else
	   KZS_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KZSClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KZSReadNext(Optional DBNo As Integer = 0)
    If Not KZS_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KZS_RDSTS = False
        Else
           Call ItemsFromKZS(DBNo)
           KZS_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KZSClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KZS_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KZS )
'*----------------------------------------------*'
Public Sub ItemsClearKZS()
    With KZS
	.取引先 = ""
	.受入 = ""
	.品番 = ""
	.開始日 = ""
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
'	.更新日 = 0
	.更新者 = ""
'	.更新端末 = ""
        .件数 = 0
'   項目追加(自動)
	.取引先名 = ""
	.表示品番 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KZSInsert(Optional DBNo As Integer) As Boolean
    KZSInsert = True
    With KZS
        gSL_Select = "Insert Into かんばん増減開始日マスタ  ( "
        gSL_Select = gSL_Select & "取引先,受入,品番,"
        gSL_Select = gSL_Select & "開始日,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.開始日) & "',"		
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
    KZSInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KZSUpdate(Optional DBNo As Integer) As Boolean
    KZSUpdate = False	
    With KZS
        gSL_Select = "Update かんばん増減開始日マスタ  Set "
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "開始日 = '" & RTrim(.開始日) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KZSUpdate = True
End Function

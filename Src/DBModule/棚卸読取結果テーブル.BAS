Attribute VB_Name = "棚卸読取結果テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   棚卸読取結果テーブル
'*
'*   モジュールID  : 棚卸読取結果テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2019/07/02  By Cis(CDH00301)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_棚卸読取結果テーブル
        対象年月              As String
        処理日                As String
        処理時間              As String
        機種名                As String
        担当者                As String
        置場                  As String
        読取連番              As Long
        かんばんQR内容        As String
        品番                  As String
        実棚数                As Long
'        作成日                As Date
'        作成端末              As String
        作成者                As String
        件数          As Long
'   項目追加(自動)
	取引先名       As String
	品名           As String
    End Type

    Public HHR As T_棚卸読取結果テーブル

    Public HHR_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHHR(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHHR	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With HHR
            Select Case adoCol.Name
                Case "対象年月":              If Not IsNull(adoCol.Value) Then .対象年月             = adoCol.Value
                Case "処理日":                If Not IsNull(adoCol.Value) Then .処理日               = adoCol.Value
                Case "処理時間":              If Not IsNull(adoCol.Value) Then .処理時間             = adoCol.Value
                Case "機種名":                If Not IsNull(adoCol.Value) Then .機種名               = adoCol.Value
                Case "担当者":                If Not IsNull(adoCol.Value) Then .担当者               = adoCol.Value
                Case "置場":                  If Not IsNull(adoCol.Value) Then .置場                 = adoCol.Value
                Case "読取連番":              If Not IsNull(adoCol.Value) Then .読取連番             = adoCol.Value
                Case "かんばんQR内容":        If Not IsNull(adoCol.Value) Then .かんばんQR内容       = adoCol.Value
                Case "品番":                  If Not IsNull(adoCol.Value) Then .品番                 = adoCol.Value
                Case "実棚数":                If Not IsNull(adoCol.Value) Then .実棚数               = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "作成者":                If Not IsNull(adoCol.Value) Then .作成者               = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "取引先名":	If Not IsNull(adoCol.Value) Then .取引先名	= adoCol.Value
		Case "品名":	If Not IsNull(adoCol.Value) Then .品名	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function HHRRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HHRRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 棚卸読取結果テーブル"
        RTxtSql = RTxtSql &" Where 対象年月 = '" & RTrim(HHR.対象年月) & "'"
        RTxtSql = RTxtSql &" And   処理日 = '" & RTrim(HHR.処理日) & "'"
        RTxtSql = RTxtSql &" And   処理時間 = '" & RTrim(HHR.処理時間) & "'"
        RTxtSql = RTxtSql &" And   機種名 = '" & RTrim(HHR.機種名) & "'"
        RTxtSql = RTxtSql &" And   担当者 = '" & RTrim(HHR.担当者) & "'"
        RTxtSql = RTxtSql &" And   置場 = '" & RTrim(HHR.置場) & "'"
        RTxtSql = RTxtSql &" And   読取連番 = " & HHR.読取連番 & ""
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(HHR.品番) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HHRRead = True
           Call ItemsFromHHR(DBNo)
	   HHR_RDSTS = True
        else
	   HHR_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call HHRClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HHRReadNext(Optional DBNo As Integer = 0)
    If Not HHR_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HHR_RDSTS = False
        Else
           Call ItemsFromHHR(DBNo)
           HHR_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HHRClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HHR_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HHR )
'*----------------------------------------------*'
Public Sub ItemsClearHHR()
    With HHR
	.対象年月 = ""
	.処理日 = ""
	.処理時間 = ""
	.機種名 = ""
	.担当者 = ""
	.置場 = ""
	.読取連番 = 0
	.かんばんQR内容 = ""
	.品番 = ""
	.実棚数 = 0
'	.作成日 = 0
'	.作成端末 = ""
	.作成者 = ""
        .件数 = 0
'   項目追加(自動)
	.取引先名 = ""
	.品名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function HHRInsert(Optional DBNo As Integer) As Boolean
    HHRInsert = False
    With HHR
        gSL_Select = "Insert Into 棚卸読取結果テーブル  ( "
        gSL_Select = gSL_Select & "対象年月,処理日,処理時間,"
        gSL_Select = gSL_Select & "機種名,担当者,置場,"
        gSL_Select = gSL_Select & "読取連番,かんばんQR内容,品番,"
        gSL_Select = gSL_Select & "実棚数,"
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.対象年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.処理日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.処理時間) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.機種名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.担当者) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.置場) & "',"		
        gSL_Select = gSL_Select & "" & .読取連番 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.かんばんQR内容) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "" & .実棚数 & ","		
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
    HHRInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function HHRUpdate(Optional DBNo As Integer) As Boolean
    HHRUpdate = False	
    With HHR
        gSL_Select = "Update 棚卸読取結果テーブル  Set "
        gSL_Select = gSL_Select & "対象年月 = '" & RTrim(.対象年月) & "',"
        gSL_Select = gSL_Select & "処理日 = '" & RTrim(.処理日) & "',"
        gSL_Select = gSL_Select & "処理時間 = '" & RTrim(.処理時間) & "',"
        gSL_Select = gSL_Select & "機種名 = '" & RTrim(.機種名) & "',"
        gSL_Select = gSL_Select & "担当者 = '" & RTrim(.担当者) & "',"
        gSL_Select = gSL_Select & "置場 = '" & RTrim(.置場) & "',"
        gSL_Select = gSL_Select & "読取連番 = " & .読取連番 & ","
        gSL_Select = gSL_Select & "かんばんQR内容 = '" & RTrim(.かんばんQR内容) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "実棚数 = " & .実棚数 & ""
        gSL_Select = gSL_Select & " Where 対象年月 = '" & RTrim(.対象年月) & "'"
        gSL_Select = gSL_Select & "   And 処理日 = '" & RTrim(.処理日) & "'"
        gSL_Select = gSL_Select & "   And 処理時間 = '" & RTrim(.処理時間) & "'"
        gSL_Select = gSL_Select & "   And 機種名 = '" & RTrim(.機種名) & "'"
        gSL_Select = gSL_Select & "   And 担当者 = '" & RTrim(.担当者) & "'"
        gSL_Select = gSL_Select & "   And 置場 = '" & RTrim(.置場) & "'"
        gSL_Select = gSL_Select & "   And 読取連番 = " & .読取連番 & ""
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HHRUpdate = True
End Function

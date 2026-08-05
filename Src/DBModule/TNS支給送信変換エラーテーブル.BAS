Attribute VB_Name = "TNS支給送信変換エラーテーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TNS支給送信変換エラーテーブル
'*
'*   モジュールID  : TNS支給送信変換エラーテーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2007/03/08  By Cis
'*   作成日  :  2007/03/09  By Cis
'*   作成日  :  2007/03/12  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_TNS支給送信変換エラーテーブル
        履歴管理NO        As Long
        契約先            As String
        FCD               As String
        HEADKEY           As String
        品番              As String
        仕入先            As String
        受入              As String
        REC               As String
        エラー            As Byte
        エラー内容        As String
        納入日            As String
        納入番号          As String
        発注管理NO        As Long
        変換日付          As String
        変換時刻          As String
        変換端末          As String
        更新時刻          As String
'        更新日付          As String
'        更新端末          As String
        件数          As Long
'   項目追加(自動)
	仕入先名       As String
	背番号         As String
	表示品番       As String
	品名           As String
	車型           As String
	サイクル       As String
	収容数         As Long
    End Type

    Public THE As T_TNS支給送信変換エラーテーブル

    Public THE_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTHE(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTHE	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With THE
            Select Case adoCol.Name
                Case "履歴管理NO":        If Not IsNull(adoCol.Value) Then .履歴管理NO       = adoCol.Value
                Case "契約先":            If Not IsNull(adoCol.Value) Then .契約先           = adoCol.Value
                Case "FCD":               If Not IsNull(adoCol.Value) Then .FCD              = adoCol.Value
                Case "HEADKEY":           If Not IsNull(adoCol.Value) Then .HEADKEY          = adoCol.Value
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番             = adoCol.Value
                Case "仕入先":            If Not IsNull(adoCol.Value) Then .仕入先           = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入             = adoCol.Value
                Case "REC":               If Not IsNull(adoCol.Value) Then .REC              = adoCol.Value
                Case "エラー":            If Not IsNull(adoCol.Value) Then .エラー           = adoCol.Value
                Case "エラー内容":        If Not IsNull(adoCol.Value) Then .エラー内容       = adoCol.Value
                Case "納入日":            If Not IsNull(adoCol.Value) Then .納入日           = adoCol.Value
                Case "納入番号":          If Not IsNull(adoCol.Value) Then .納入番号         = adoCol.Value
                Case "発注管理NO":        If Not IsNull(adoCol.Value) Then .発注管理NO       = adoCol.Value
                Case "変換日付":          If Not IsNull(adoCol.Value) Then .変換日付         = adoCol.Value
                Case "変換時刻":          If Not IsNull(adoCol.Value) Then .変換時刻         = adoCol.Value
                Case "変換端末":          If Not IsNull(adoCol.Value) Then .変換端末         = adoCol.Value
                Case "更新時刻":          If Not IsNull(adoCol.Value) Then .更新時刻         = adoCol.Value
'                Case "更新日付":          If Not IsNull(adoCol.Value) Then .更新日付         = adoCol.Value
'                Case "更新端末":          If Not IsNull(adoCol.Value) Then .更新端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "仕入先名":	If Not IsNull(adoCol.Value) Then .仕入先名	= adoCol.Value
		Case "背番号":	If Not IsNull(adoCol.Value) Then .背番号	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "品名":	If Not IsNull(adoCol.Value) Then .品名	= adoCol.Value
		Case "車型":	If Not IsNull(adoCol.Value) Then .車型	= adoCol.Value
		Case "サイクル":	If Not IsNull(adoCol.Value) Then .サイクル	= adoCol.Value
		Case "収容数":	If Not IsNull(adoCol.Value) Then .収容数	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function THERead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    THERead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From TNS支給送信変換エラーテーブル"
        RTxtSql = RTxtSql &" Where 契約先 = '" & RTrim(THE.契約先) & "'"
        RTxtSql = RTxtSql &" And   HEADKEY = '" & RTrim(THE.HEADKEY) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           THERead = True
           Call ItemsFromTHE(DBNo)
	   THE_RDSTS = True
        else
	   THE_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call THEClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub THEReadNext(Optional DBNo As Integer = 0)
    If Not THE_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           THE_RDSTS = False
        Else
           Call ItemsFromTHE(DBNo)
           THE_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub THEClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        THE_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( THE )
'*----------------------------------------------*'
Public Sub ItemsClearTHE()
    With THE
	.履歴管理NO = 0
	.契約先 = ""
	.FCD = ""
	.HEADKEY = ""
	.品番 = ""
	.仕入先 = ""
	.受入 = ""
	.REC = ""
	.エラー = 0
	.エラー内容 = ""
	.納入日 = ""
	.納入番号 = ""
	.発注管理NO = 0
	.変換日付 = ""
	.変換時刻 = ""
	.変換端末 = ""
	.更新時刻 = ""
'	.更新日付 = ""
'	.更新端末 = ""
        .件数 = 0
'   項目追加(自動)
	.仕入先名 = ""
	.背番号 = ""
	.表示品番 = ""
	.品名 = ""
	.車型 = ""
	.サイクル = ""
	.収容数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function THEInsert(Optional DBNo As Integer) As Boolean
    THEInsert = False
    With THE
        gSL_Select = "Insert Into TNS支給送信変換エラーテーブル  ( "
        gSL_Select = gSL_Select & "履歴管理NO,契約先,FCD,"
        gSL_Select = gSL_Select & "HEADKEY,品番,仕入先,"
        gSL_Select = gSL_Select & "受入,REC,エラー,"
        gSL_Select = gSL_Select & "エラー内容,納入日,納入番号,"
        gSL_Select = gSL_Select & "発注管理NO,変換日付,変換時刻,"
        gSL_Select = gSL_Select & "変換端末,更新時刻"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .履歴管理NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.FCD) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.HEADKEY) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.REC) & "',"		
        gSL_Select = gSL_Select & "" & .エラー & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.エラー内容) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"		
        gSL_Select = gSL_Select & "" & .発注管理NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.変換日付) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.変換時刻) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.変換端末) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.更新時刻) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    THEInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function THEUpdate(Optional DBNo As Integer) As Boolean
    THEUpdate = False	
    With THE
        gSL_Select = "Update TNS支給送信変換エラーテーブル  Set "
        gSL_Select = gSL_Select & "履歴管理NO = " & .履歴管理NO & ","
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "FCD = '" & RTrim(.FCD) & "',"
        gSL_Select = gSL_Select & "HEADKEY = '" & RTrim(.HEADKEY) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "REC = '" & RTrim(.REC) & "',"
        gSL_Select = gSL_Select & "エラー = " & .エラー & ","
        gSL_Select = gSL_Select & "エラー内容 = '" & RTrim(.エラー内容) & "',"
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "発注管理NO = " & .発注管理NO & ","
        gSL_Select = gSL_Select & "変換日付 = '" & RTrim(.変換日付) & "',"
        gSL_Select = gSL_Select & "変換時刻 = '" & RTrim(.変換時刻) & "',"
        gSL_Select = gSL_Select & "変換端末 = '" & RTrim(.変換端末) & "',"
        gSL_Select = gSL_Select & "更新時刻 = '" & RTrim(.更新時刻) & "',"
        gSL_Select = gSL_Select & "更新日付 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 契約先 = '" & RTrim(.契約先) & "'"
        gSL_Select = gSL_Select & "   And HEADKEY = '" & RTrim(.HEADKEY) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    THEUpdate = True
End Function

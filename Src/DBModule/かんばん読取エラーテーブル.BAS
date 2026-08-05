Attribute VB_Name = "かんばん読取エラーテーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   かんばん読取エラーテーブル
'*
'*   モジュールID  : かんばん読取エラーテーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/06/01  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_かんばん読取エラーテーブル
        エラー管理番号        As Long
        エラー種類            As Byte
        品番                  As String
        取引先                As String
        受入                  As String
        発行NO                As Integer
        収容数                As Long
        背番号                As String
        手配区分              As String
        サイクル              As String
        エラー要因            As Byte
        平準化区分            As Byte
        平準化設定枚数        As Long
        平準化超過枚数        As Long
        打切種別              As Byte
        打切区分              As Byte
        完了区分              As String
        起算日                As String
        起算便                As String
        完了日                As String
        完了便                As String
        打切日                As String
        打切設定数            As Long
        警告数                As Long
        警告日                As String
        発注済数              As Long
        今回発注数            As Long
        設定残数              As Long
        今回発注枚数          As Long
        今回発注端数          As Long
        打切エラー枚数        As Long
'        作成日                As Date
'        作成端末              As String
        件数          As Long
'   項目追加(自動)
	取引先名      As String
	手配区分名    As String
	表示品番      As String
    End Type

    Public KRE As T_かんばん読取エラーテーブル

    Public KRE_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKRE(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKRE	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KRE
            Select Case adoCol.Name
                Case "エラー管理番号":        If Not IsNull(adoCol.Value) Then .エラー管理番号       = adoCol.Value
                Case "エラー種類":            If Not IsNull(adoCol.Value) Then .エラー種類           = adoCol.Value
                Case "品番":                  If Not IsNull(adoCol.Value) Then .品番                 = adoCol.Value
                Case "取引先":                If Not IsNull(adoCol.Value) Then .取引先               = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入                 = adoCol.Value
                Case "発行NO":                If Not IsNull(adoCol.Value) Then .発行NO               = adoCol.Value
                Case "収容数":                If Not IsNull(adoCol.Value) Then .収容数               = adoCol.Value
                Case "背番号":                If Not IsNull(adoCol.Value) Then .背番号               = adoCol.Value
                Case "手配区分":              If Not IsNull(adoCol.Value) Then .手配区分             = adoCol.Value
                Case "サイクル":              If Not IsNull(adoCol.Value) Then .サイクル             = adoCol.Value
                Case "エラー要因":            If Not IsNull(adoCol.Value) Then .エラー要因           = adoCol.Value
                Case "平準化区分":            If Not IsNull(adoCol.Value) Then .平準化区分           = adoCol.Value
                Case "平準化設定枚数":        If Not IsNull(adoCol.Value) Then .平準化設定枚数       = adoCol.Value
                Case "平準化超過枚数":        If Not IsNull(adoCol.Value) Then .平準化超過枚数       = adoCol.Value
                Case "打切種別":              If Not IsNull(adoCol.Value) Then .打切種別             = adoCol.Value
                Case "打切区分":              If Not IsNull(adoCol.Value) Then .打切区分             = adoCol.Value
                Case "完了区分":              If Not IsNull(adoCol.Value) Then .完了区分             = adoCol.Value
                Case "起算日":                If Not IsNull(adoCol.Value) Then .起算日               = adoCol.Value
                Case "起算便":                If Not IsNull(adoCol.Value) Then .起算便               = adoCol.Value
                Case "完了日":                If Not IsNull(adoCol.Value) Then .完了日               = adoCol.Value
                Case "完了便":                If Not IsNull(adoCol.Value) Then .完了便               = adoCol.Value
                Case "打切日":                If Not IsNull(adoCol.Value) Then .打切日               = adoCol.Value
                Case "打切設定数":            If Not IsNull(adoCol.Value) Then .打切設定数           = adoCol.Value
                Case "警告数":                If Not IsNull(adoCol.Value) Then .警告数               = adoCol.Value
                Case "警告日":                If Not IsNull(adoCol.Value) Then .警告日               = adoCol.Value
                Case "発注済数":              If Not IsNull(adoCol.Value) Then .発注済数             = adoCol.Value
                Case "今回発注数":            If Not IsNull(adoCol.Value) Then .今回発注数           = adoCol.Value
                Case "設定残数":              If Not IsNull(adoCol.Value) Then .設定残数             = adoCol.Value
                Case "今回発注枚数":          If Not IsNull(adoCol.Value) Then .今回発注枚数         = adoCol.Value
                Case "今回発注端数":          If Not IsNull(adoCol.Value) Then .今回発注端数         = adoCol.Value
                Case "打切エラー枚数":        If Not IsNull(adoCol.Value) Then .打切エラー枚数       = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "取引先名":	If Not IsNull(adoCol.Value) Then .取引先名	= adoCol.Value
		Case "手配区分名":	If Not IsNull(adoCol.Value) Then .手配区分名	= adoCol.Value
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KRERead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KRERead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From かんばん読取エラーテーブル"
        RTxtSql = RTxtSql &" Where エラー管理番号 = " & KRE.エラー管理番号 & ""
        RTxtSql = RTxtSql &" And   エラー種類 = " & KRE.エラー種類 & ""
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(KRE.品番) & "'"
        RTxtSql = RTxtSql &" And   取引先 = '" & RTrim(KRE.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(KRE.受入) & "'"
        RTxtSql = RTxtSql &" And   発行NO = " & KRE.発行NO & ""
        RTxtSql = RTxtSql &" And   収容数 = " & KRE.収容数 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KRERead = True
           Call ItemsFromKRE(DBNo)
	   KRE_RDSTS = True
        else
	   KRE_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KREClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KREReadNext(Optional DBNo As Integer = 0)
    If Not KRE_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KRE_RDSTS = False
        Else
           Call ItemsFromKRE(DBNo)
           KRE_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KREClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KRE_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KRE )
'*----------------------------------------------*'
Public Sub ItemsClearKRE()
    With KRE
	.エラー管理番号 = 0
	.エラー種類 = 0
	.品番 = ""
	.取引先 = ""
	.受入 = ""
	.発行NO = 0
	.収容数 = 0
	.背番号 = ""
	.手配区分 = ""
	.サイクル = ""
	.エラー要因 = 0
	.平準化区分 = 0
	.平準化設定枚数 = 0
	.平準化超過枚数 = 0
	.打切種別 = 0
	.打切区分 = 0
	.完了区分 = ""
	.起算日 = ""
	.起算便 = ""
	.完了日 = ""
	.完了便 = ""
	.打切日 = ""
	.打切設定数 = 0
	.警告数 = 0
	.警告日 = ""
	.発注済数 = 0
	.今回発注数 = 0
	.設定残数 = 0
	.今回発注枚数 = 0
	.今回発注端数 = 0
	.打切エラー枚数 = 0
'	.作成日 = 0
'	.作成端末 = ""
        .件数 = 0
'   項目追加(自動)
	.取引先名 = ""
	.手配区分名 = ""
	.表示品番 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KREInsert(Optional DBNo As Integer) As Boolean
    KREInsert = True
    With KRE
        gSL_Select = "Insert Into かんばん読取エラーテーブル  ( "
        gSL_Select = gSL_Select & "エラー管理番号,エラー種類,品番,"
        gSL_Select = gSL_Select & "取引先,受入,発行NO,"
        gSL_Select = gSL_Select & "収容数,背番号,手配区分,"
        gSL_Select = gSL_Select & "サイクル,エラー要因,平準化区分,"
        gSL_Select = gSL_Select & "平準化設定枚数,平準化超過枚数,打切種別,"
        gSL_Select = gSL_Select & "打切区分,完了区分,起算日,"
        gSL_Select = gSL_Select & "起算便,完了日,完了便,"
        gSL_Select = gSL_Select & "打切日,打切設定数,警告数,"
        gSL_Select = gSL_Select & "警告日,発注済数,今回発注数,"
        gSL_Select = gSL_Select & "設定残数,今回発注枚数,今回発注端数,"
        gSL_Select = gSL_Select & "打切エラー枚数,"
        gSL_Select = gSL_Select & "作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .エラー管理番号 & ","		
        gSL_Select = gSL_Select & "" & .エラー種類 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "" & .発行NO & ","		
        gSL_Select = gSL_Select & "" & .収容数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.サイクル) & "',"		
        gSL_Select = gSL_Select & "" & .エラー要因 & ","		
        gSL_Select = gSL_Select & "" & .平準化区分 & ","		
        gSL_Select = gSL_Select & "" & .平準化設定枚数 & ","		
        gSL_Select = gSL_Select & "" & .平準化超過枚数 & ","		
        gSL_Select = gSL_Select & "" & .打切種別 & ","		
        gSL_Select = gSL_Select & "" & .打切区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.完了区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.起算日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.起算便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.完了日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.完了便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.打切日) & "',"		
        gSL_Select = gSL_Select & "" & .打切設定数 & ","		
        gSL_Select = gSL_Select & "" & .警告数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.警告日) & "',"		
        gSL_Select = gSL_Select & "" & .発注済数 & ","		
        gSL_Select = gSL_Select & "" & .今回発注数 & ","		
        gSL_Select = gSL_Select & "" & .設定残数 & ","		
        gSL_Select = gSL_Select & "" & .今回発注枚数 & ","		
        gSL_Select = gSL_Select & "" & .今回発注端数 & ","		
        gSL_Select = gSL_Select & "" & .打切エラー枚数 & ","		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KREInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KREUpdate(Optional DBNo As Integer) As Boolean
    KREUpdate = False	
    With KRE
        gSL_Select = "Update かんばん読取エラーテーブル  Set "
        gSL_Select = gSL_Select & "エラー管理番号 = " & .エラー管理番号 & ","
        gSL_Select = gSL_Select & "エラー種類 = " & .エラー種類 & ","
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "発行NO = " & .発行NO & ","
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "手配区分 = '" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "サイクル = '" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "エラー要因 = " & .エラー要因 & ","
        gSL_Select = gSL_Select & "平準化区分 = " & .平準化区分 & ","
        gSL_Select = gSL_Select & "平準化設定枚数 = " & .平準化設定枚数 & ","
        gSL_Select = gSL_Select & "平準化超過枚数 = " & .平準化超過枚数 & ","
        gSL_Select = gSL_Select & "打切種別 = " & .打切種別 & ","
        gSL_Select = gSL_Select & "打切区分 = " & .打切区分 & ","
        gSL_Select = gSL_Select & "完了区分 = '" & RTrim(.完了区分) & "',"
        gSL_Select = gSL_Select & "起算日 = '" & RTrim(.起算日) & "',"
        gSL_Select = gSL_Select & "起算便 = '" & RTrim(.起算便) & "',"
        gSL_Select = gSL_Select & "完了日 = '" & RTrim(.完了日) & "',"
        gSL_Select = gSL_Select & "完了便 = '" & RTrim(.完了便) & "',"
        gSL_Select = gSL_Select & "打切日 = '" & RTrim(.打切日) & "',"
        gSL_Select = gSL_Select & "打切設定数 = " & .打切設定数 & ","
        gSL_Select = gSL_Select & "警告数 = " & .警告数 & ","
        gSL_Select = gSL_Select & "警告日 = '" & RTrim(.警告日) & "',"
        gSL_Select = gSL_Select & "発注済数 = " & .発注済数 & ","
        gSL_Select = gSL_Select & "今回発注数 = " & .今回発注数 & ","
        gSL_Select = gSL_Select & "設定残数 = " & .設定残数 & ","
        gSL_Select = gSL_Select & "今回発注枚数 = " & .今回発注枚数 & ","
        gSL_Select = gSL_Select & "今回発注端数 = " & .今回発注端数 & ","
        gSL_Select = gSL_Select & "打切エラー枚数 = " & .打切エラー枚数 & ""
        gSL_Select = gSL_Select & " Where エラー管理番号 = " & .エラー管理番号 & ""
        gSL_Select = gSL_Select & "   And エラー種類 = " & .エラー種類 & ""
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 発行NO = " & .発行NO & ""
        gSL_Select = gSL_Select & "   And 収容数 = " & .収容数 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KREUpdate = True
End Function

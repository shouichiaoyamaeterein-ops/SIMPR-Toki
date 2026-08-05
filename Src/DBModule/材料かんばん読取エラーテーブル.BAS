Attribute VB_Name = "材料かんばん読取エラーテーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料かんばん読取エラーテーブル
'*
'*   モジュールID  : 材料かんばん読取エラーテーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/02/21  By Cis
'*   作成日  :  2008/10/01  By Cis(FMV-E630)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料かんばん読取エラーテーブル
        エラー管理番号        As Long
        エラー種類            As Byte
        材料管理番号          As Long
        取引先                As String
        受入                  As String
        発行NO                As Integer
        収容数                As Currency
        材質                  As String
        板厚                  As Currency
        幅                    As Currency
        長さ                  As Currency
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
        打切設定数            As Currency
        警告数                As Currency
        警告日                As String
        発注済数              As Currency
        今回発注数            As Currency
        設定残数              As Currency
        今回発注枚数          As Long
        今回発注端数          As Currency
        打切エラー枚数        As Long
'        作成日                As Date
'        作成端末              As String
        件数          As Long
'   項目追加(自動)
	取引先名       As String
	手配区分名     As String
	表示寸法       As String
	サイズ         As String
    End Type

    Public ZRE As T_材料かんばん読取エラーテーブル

    Public ZRE_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZRE(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZRE	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZRE
            Select Case adoCol.Name
                Case "エラー管理番号":        If Not IsNull(adoCol.Value) Then .エラー管理番号       = adoCol.Value
                Case "エラー種類":            If Not IsNull(adoCol.Value) Then .エラー種類           = adoCol.Value
                Case "材料管理番号":          If Not IsNull(adoCol.Value) Then .材料管理番号         = adoCol.Value
                Case "取引先":                If Not IsNull(adoCol.Value) Then .取引先               = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入                 = adoCol.Value
                Case "発行NO":                If Not IsNull(adoCol.Value) Then .発行NO               = adoCol.Value
                Case "収容数":                If Not IsNull(adoCol.Value) Then .収容数               = adoCol.Value
                Case "材質":                  If Not IsNull(adoCol.Value) Then .材質                 = adoCol.Value
                Case "板厚":                  If Not IsNull(adoCol.Value) Then .板厚                 = adoCol.Value
                Case "幅":                    If Not IsNull(adoCol.Value) Then .幅                   = adoCol.Value
                Case "長さ":                  If Not IsNull(adoCol.Value) Then .長さ                 = adoCol.Value
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
		Case "表示寸法":	If Not IsNull(adoCol.Value) Then .表示寸法	= adoCol.Value
		Case "サイズ":	If Not IsNull(adoCol.Value) Then .サイズ	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZRERead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZRERead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 材料かんばん読取エラーテーブル"
        RTxtSql = RTxtSql &" Where エラー管理番号 = " & ZRE.エラー管理番号 & ""
        RTxtSql = RTxtSql &" And   エラー種類 = " & ZRE.エラー種類 & ""
        RTxtSql = RTxtSql &" And   材料管理番号 = " & ZRE.材料管理番号 & ""
        RTxtSql = RTxtSql &" And   取引先 = '" & RTrim(ZRE.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(ZRE.受入) & "'"
        RTxtSql = RTxtSql &" And   発行NO = " & ZRE.発行NO & ""
        RTxtSql = RTxtSql &" And   収容数 = " & ZRE.収容数 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZRERead = True
           Call ItemsFromZRE(DBNo)
	   ZRE_RDSTS = True
        else
	   ZRE_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZREClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZREReadNext(Optional DBNo As Integer = 0)
    If Not ZRE_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZRE_RDSTS = False
        Else
           Call ItemsFromZRE(DBNo)
           ZRE_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZREClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZRE_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZRE )
'*----------------------------------------------*'
Public Sub ItemsClearZRE()
    With ZRE
	.エラー管理番号 = 0
	.エラー種類 = 0
	.材料管理番号 = 0
	.取引先 = ""
	.受入 = ""
	.発行NO = 0
	.収容数 = 0
	.材質 = ""
	.板厚 = 0
	.幅 = 0
	.長さ = 0
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
	.表示寸法 = ""
	.サイズ = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZREInsert(Optional DBNo As Integer) As Boolean
    ZREInsert = False
    With ZRE
        gSL_Select = "Insert Into 材料かんばん読取エラーテーブル  ( "
        gSL_Select = gSL_Select & "エラー管理番号,エラー種類,材料管理番号,"
        gSL_Select = gSL_Select & "取引先,受入,発行NO,"
        gSL_Select = gSL_Select & "収容数,材質,板厚,"
        gSL_Select = gSL_Select & "幅,長さ,背番号,"
        gSL_Select = gSL_Select & "手配区分,サイクル,エラー要因,"
        gSL_Select = gSL_Select & "平準化区分,平準化設定枚数,平準化超過枚数,"
        gSL_Select = gSL_Select & "打切種別,打切区分,完了区分,"
        gSL_Select = gSL_Select & "起算日,起算便,完了日,"
        gSL_Select = gSL_Select & "完了便,打切日,打切設定数,"
        gSL_Select = gSL_Select & "警告数,警告日,発注済数,"
        gSL_Select = gSL_Select & "今回発注数,設定残数,今回発注枚数,"
        gSL_Select = gSL_Select & "今回発注端数,打切エラー枚数,"
        gSL_Select = gSL_Select & "作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .エラー管理番号 & ","		
        gSL_Select = gSL_Select & "" & .エラー種類 & ","		
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "" & .発行NO & ","		
        gSL_Select = gSL_Select & "" & .収容数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"		
        gSL_Select = gSL_Select & "" & .板厚 & ","		
        gSL_Select = gSL_Select & "" & .幅 & ","		
        gSL_Select = gSL_Select & "" & .長さ & ","		
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
    ZREInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZREUpdate(Optional DBNo As Integer) As Boolean
    ZREUpdate = False	
    With ZRE
        gSL_Select = "Update 材料かんばん読取エラーテーブル  Set "
        gSL_Select = gSL_Select & "エラー管理番号 = " & .エラー管理番号 & ","
        gSL_Select = gSL_Select & "エラー種類 = " & .エラー種類 & ","
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "発行NO = " & .発行NO & ","
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "板厚 = " & .板厚 & ","
        gSL_Select = gSL_Select & "幅 = " & .幅 & ","
        gSL_Select = gSL_Select & "長さ = " & .長さ & ","
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
        gSL_Select = gSL_Select & "   And 材料管理番号 = " & .材料管理番号 & ""
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
    ZREUpdate = True
End Function

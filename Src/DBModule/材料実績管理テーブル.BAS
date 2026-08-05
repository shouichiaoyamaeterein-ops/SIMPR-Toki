Attribute VB_Name = "材料実績管理テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料実績管理テーブル
'*
'*   モジュールID  : 材料実績管理テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/02/27  By Cis
'*   作成日  :  2008/12/18  By Cis(FMV-E630)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料実績管理テーブル
        実績管理NO    As Long
        実績区分      As Byte
        納入番号      As String
        納品書頁      As Byte
        納品書行      As Byte
        分納区分      As Byte
        実績先        As String
        受入          As String
        実績元        As String
        材質          As String
        寸法区分      As Byte
        板厚          As Currency
        幅            As Currency
        長さ          As Currency
        背番号        As String
        実績日        As String
        実績便        As String
        収容数        As Currency
        実績枚数      As Currency
        実績端数      As Currency
        実績数        As Currency
        分割数        As Currency
        挽回日        As String
        挽回便        As String
        有無償区分    As Byte
        分割納品書NO  As String
        検収区分      As String
        検収年月      As String
        検収処理日    As String
        入力種類      As String
'        作成日        As Date
'        作成端末      As String
        作成者        As String
'        更新日        As Date
'        更新端末      As String
        更新者        As String
        件数          As Long
'   項目追加(自動)
	実績先名       As String
	実績元名       As String
	取引先         As String
	工場CD         As String
	工場名         As String
	材料管理番号   As Long
	表示寸法       As String
	単位           As String
	品番           As String
    End Type

    Public ZJK As T_材料実績管理テーブル

    Public ZJK_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZJK(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZJK	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZJK
            Select Case adoCol.Name
                Case "実績管理NO":    If Not IsNull(adoCol.Value) Then .実績管理NO   = adoCol.Value
                Case "実績区分":      If Not IsNull(adoCol.Value) Then .実績区分     = adoCol.Value
                Case "納入番号":      If Not IsNull(adoCol.Value) Then .納入番号     = adoCol.Value
                Case "納品書頁":      If Not IsNull(adoCol.Value) Then .納品書頁     = adoCol.Value
                Case "納品書行":      If Not IsNull(adoCol.Value) Then .納品書行     = adoCol.Value
                Case "分納区分":      If Not IsNull(adoCol.Value) Then .分納区分     = adoCol.Value
                Case "実績先":        If Not IsNull(adoCol.Value) Then .実績先       = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "実績元":        If Not IsNull(adoCol.Value) Then .実績元       = adoCol.Value
                Case "材質":          If Not IsNull(adoCol.Value) Then .材質         = adoCol.Value
                Case "寸法区分":      If Not IsNull(adoCol.Value) Then .寸法区分     = adoCol.Value
                Case "板厚":          If Not IsNull(adoCol.Value) Then .板厚         = adoCol.Value
                Case "幅":            If Not IsNull(adoCol.Value) Then .幅           = adoCol.Value
                Case "長さ":          If Not IsNull(adoCol.Value) Then .長さ         = adoCol.Value
                Case "背番号":        If Not IsNull(adoCol.Value) Then .背番号       = adoCol.Value
                Case "実績日":        If Not IsNull(adoCol.Value) Then .実績日       = adoCol.Value
                Case "実績便":        If Not IsNull(adoCol.Value) Then .実績便       = adoCol.Value
                Case "収容数":        If Not IsNull(adoCol.Value) Then .収容数       = adoCol.Value
                Case "実績枚数":      If Not IsNull(adoCol.Value) Then .実績枚数     = adoCol.Value
                Case "実績端数":      If Not IsNull(adoCol.Value) Then .実績端数     = adoCol.Value
                Case "実績数":        If Not IsNull(adoCol.Value) Then .実績数       = adoCol.Value
                Case "分割数":        If Not IsNull(adoCol.Value) Then .分割数       = adoCol.Value
                Case "挽回日":        If Not IsNull(adoCol.Value) Then .挽回日       = adoCol.Value
                Case "挽回便":        If Not IsNull(adoCol.Value) Then .挽回便       = adoCol.Value
                Case "有無償区分":    If Not IsNull(adoCol.Value) Then .有無償区分   = adoCol.Value
                Case "分割納品書NO":  If Not IsNull(adoCol.Value) Then .分割納品書NO = adoCol.Value
                Case "検収区分":      If Not IsNull(adoCol.Value) Then .検収区分     = adoCol.Value
                Case "検収年月":      If Not IsNull(adoCol.Value) Then .検収年月     = adoCol.Value
                Case "検収処理日":    If Not IsNull(adoCol.Value) Then .検収処理日   = adoCol.Value
                Case "入力種類":      If Not IsNull(adoCol.Value) Then .入力種類     = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者       = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "実績先名":	If Not IsNull(adoCol.Value) Then .実績先名	= adoCol.Value
		Case "実績元名":	If Not IsNull(adoCol.Value) Then .実績元名	= adoCol.Value
		Case "取引先":	If Not IsNull(adoCol.Value) Then .取引先	= adoCol.Value
		Case "工場CD":	If Not IsNull(adoCol.Value) Then .工場CD	= adoCol.Value
		Case "工場名":	If Not IsNull(adoCol.Value) Then .工場名	= adoCol.Value
		Case "材料管理番号":	If Not IsNull(adoCol.Value) Then .材料管理番号	= adoCol.Value
		Case "表示寸法":	If Not IsNull(adoCol.Value) Then .表示寸法	= adoCol.Value
		Case "単位":	If Not IsNull(adoCol.Value) Then .単位	= adoCol.Value
		Case "品番":	If Not IsNull(adoCol.Value) Then .品番	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZJKRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZJKRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 材料実績管理テーブル"
        RTxtSql = RTxtSql &" Where 実績管理NO = " & ZJK.実績管理NO & ""
        RTxtSql = RTxtSql &" And   実績区分 = " & ZJK.実績区分 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZJKRead = True
           Call ItemsFromZJK(DBNo)
	   ZJK_RDSTS = True
        else
	   ZJK_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZJKClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZJKReadNext(Optional DBNo As Integer = 0)
    If Not ZJK_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZJK_RDSTS = False
        Else
           Call ItemsFromZJK(DBNo)
           ZJK_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZJKClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZJK_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZJK )
'*----------------------------------------------*'
Public Sub ItemsClearZJK()
    With ZJK
	.実績管理NO = 0
	.実績区分 = 0
	.納入番号 = ""
	.納品書頁 = 0
	.納品書行 = 0
	.分納区分 = 0
	.実績先 = ""
	.受入 = ""
	.実績元 = ""
	.材質 = ""
	.寸法区分 = 0
	.板厚 = 0
	.幅 = 0
	.長さ = 0
	.背番号 = ""
	.実績日 = ""
	.実績便 = ""
	.収容数 = 0
	.実績枚数 = 0
	.実績端数 = 0
	.実績数 = 0
	.分割数 = 0
	.挽回日 = ""
	.挽回便 = ""
	.有無償区分 = 0
	.分割納品書NO = ""
	.検収区分 = ""
	.検収年月 = ""
	.検収処理日 = ""
	.入力種類 = ""
'	.作成日 = 0
'	.作成端末 = ""
	.作成者 = ""
'	.更新日 = 0
'	.更新端末 = ""
	.更新者 = ""
        .件数 = 0
'   項目追加(自動)
	.実績先名 = ""
	.実績元名 = ""
	.取引先 = ""
	.工場CD = ""
	.工場名 = ""
	.材料管理番号 = 0
	.表示寸法 = ""
	.単位 = ""
	.品番 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZJKInsert(Optional DBNo As Integer) As Boolean
    ZJKInsert = False
    With ZJK
        gSL_Select = "Insert Into 材料実績管理テーブル  ( "
        gSL_Select = gSL_Select & "実績管理NO,実績区分,納入番号,"
        gSL_Select = gSL_Select & "納品書頁,納品書行,分納区分,"
        gSL_Select = gSL_Select & "実績先,受入,実績元,"
        gSL_Select = gSL_Select & "材質,寸法区分,板厚,"
        gSL_Select = gSL_Select & "幅,長さ,背番号,"
        gSL_Select = gSL_Select & "実績日,実績便,収容数,"
        gSL_Select = gSL_Select & "実績枚数,実績端数,実績数,"
        gSL_Select = gSL_Select & "分割数,挽回日,挽回便,"
        gSL_Select = gSL_Select & "有無償区分,分割納品書NO,検収区分,"
        gSL_Select = gSL_Select & "検収年月,検収処理日,入力種類,"
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .実績管理NO & ","		
        gSL_Select = gSL_Select & "" & .実績区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"		
        gSL_Select = gSL_Select & "" & .納品書頁 & ","		
        gSL_Select = gSL_Select & "" & .納品書行 & ","		
        gSL_Select = gSL_Select & "" & .分納区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.実績先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.実績元) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"		
        gSL_Select = gSL_Select & "" & .寸法区分 & ","		
        gSL_Select = gSL_Select & "" & .板厚 & ","		
        gSL_Select = gSL_Select & "" & .幅 & ","		
        gSL_Select = gSL_Select & "" & .長さ & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.実績日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.実績便) & "',"		
        gSL_Select = gSL_Select & "" & .収容数 & ","		
        gSL_Select = gSL_Select & "" & .実績枚数 & ","		
        gSL_Select = gSL_Select & "" & .実績端数 & ","		
        gSL_Select = gSL_Select & "" & .実績数 & ","		
        gSL_Select = gSL_Select & "" & .分割数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.挽回日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.挽回便) & "',"		
        gSL_Select = gSL_Select & "" & .有無償区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.分割納品書NO) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.検収区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.検収年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.検収処理日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.入力種類) & "',"		
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
    ZJKInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZJKUpdate(Optional DBNo As Integer) As Boolean
    ZJKUpdate = False	
    With ZJK
        gSL_Select = "Update 材料実績管理テーブル  Set "
        gSL_Select = gSL_Select & "実績管理NO = " & .実績管理NO & ","
        gSL_Select = gSL_Select & "実績区分 = " & .実績区分 & ","
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "納品書頁 = " & .納品書頁 & ","
        gSL_Select = gSL_Select & "納品書行 = " & .納品書行 & ","
        gSL_Select = gSL_Select & "分納区分 = " & .分納区分 & ","
        gSL_Select = gSL_Select & "実績先 = '" & RTrim(.実績先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "実績元 = '" & RTrim(.実績元) & "',"
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "寸法区分 = " & .寸法区分 & ","
        gSL_Select = gSL_Select & "板厚 = " & .板厚 & ","
        gSL_Select = gSL_Select & "幅 = " & .幅 & ","
        gSL_Select = gSL_Select & "長さ = " & .長さ & ","
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "実績日 = '" & RTrim(.実績日) & "',"
        gSL_Select = gSL_Select & "実績便 = '" & RTrim(.実績便) & "',"
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "実績枚数 = " & .実績枚数 & ","
        gSL_Select = gSL_Select & "実績端数 = " & .実績端数 & ","
        gSL_Select = gSL_Select & "実績数 = " & .実績数 & ","
        gSL_Select = gSL_Select & "分割数 = " & .分割数 & ","
        gSL_Select = gSL_Select & "挽回日 = '" & RTrim(.挽回日) & "',"
        gSL_Select = gSL_Select & "挽回便 = '" & RTrim(.挽回便) & "',"
        gSL_Select = gSL_Select & "有無償区分 = " & .有無償区分 & ","
        gSL_Select = gSL_Select & "分割納品書NO = '" & RTrim(.分割納品書NO) & "',"
        gSL_Select = gSL_Select & "検収区分 = '" & RTrim(.検収区分) & "',"
        gSL_Select = gSL_Select & "検収年月 = '" & RTrim(.検収年月) & "',"
        gSL_Select = gSL_Select & "検収処理日 = '" & RTrim(.検収処理日) & "',"
        gSL_Select = gSL_Select & "入力種類 = '" & RTrim(.入力種類) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "'"
        gSL_Select = gSL_Select & " Where 実績管理NO = " & .実績管理NO & ""
        gSL_Select = gSL_Select & "   And 実績区分 = " & .実績区分 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZJKUpdate = True
End Function

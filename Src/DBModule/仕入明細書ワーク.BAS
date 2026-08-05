Attribute VB_Name = "仕入明細書ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   仕入明細書ワーク
'*
'*   モジュールID  : 仕入明細書ワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2009/02/20  By Cis(FMV-E630)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_仕入明細書ワーク
        処理端末      As String
        区分          As String
        区分名称      As String
        契約先        As String
        契約先名      As String
        品番          As String
        材料管理番号  As Long
        仕入先        As String
        仕入先名      As String
        入荷数        As Long
        単価          As Currency
        有償支給      As Currency
        外注材料      As Currency
        外注購入      As Currency
        有償支給計    As Currency
        外注材料計    As Currency
        外注購入計    As Currency
        実績換算値    As Currency
        金額          As Currency
        自社工区      As Byte
        受入          As String
        件数          As Long
    End Type

    Public SMW As T_仕入明細書ワーク

    Public SMW_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromSMW(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearSMW	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With SMW
            Select Case adoCol.Name
                Case "処理端末":      If Not IsNull(adoCol.Value) Then .処理端末     = adoCol.Value
                Case "区分":          If Not IsNull(adoCol.Value) Then .区分         = adoCol.Value
                Case "区分名称":      If Not IsNull(adoCol.Value) Then .区分名称     = adoCol.Value
                Case "契約先":        If Not IsNull(adoCol.Value) Then .契約先       = adoCol.Value
                Case "契約先名":      If Not IsNull(adoCol.Value) Then .契約先名     = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "材料管理番号":  If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "仕入先":        If Not IsNull(adoCol.Value) Then .仕入先       = adoCol.Value
                Case "仕入先名":      If Not IsNull(adoCol.Value) Then .仕入先名     = adoCol.Value
                Case "入荷数":        If Not IsNull(adoCol.Value) Then .入荷数       = adoCol.Value
                Case "単価":          If Not IsNull(adoCol.Value) Then .単価         = adoCol.Value
                Case "有償支給":      If Not IsNull(adoCol.Value) Then .有償支給     = adoCol.Value
                Case "外注材料":      If Not IsNull(adoCol.Value) Then .外注材料     = adoCol.Value
                Case "外注購入":      If Not IsNull(adoCol.Value) Then .外注購入     = adoCol.Value
                Case "有償支給計":    If Not IsNull(adoCol.Value) Then .有償支給計   = adoCol.Value
                Case "外注材料計":    If Not IsNull(adoCol.Value) Then .外注材料計   = adoCol.Value
                Case "外注購入計":    If Not IsNull(adoCol.Value) Then .外注購入計   = adoCol.Value
                Case "実績換算値":    If Not IsNull(adoCol.Value) Then .実績換算値   = adoCol.Value
                Case "金額":          If Not IsNull(adoCol.Value) Then .金額         = adoCol.Value
                Case "自社工区":      If Not IsNull(adoCol.Value) Then .自社工区     = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function SMWRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    SMWRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 仕入明細書ワーク"
        RTxtSql = RTxtSql &" Where 処理端末 = '" & RTrim(SMW.処理端末) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           SMWRead = True
           Call ItemsFromSMW(DBNo)
	   SMW_RDSTS = True
        else
	   SMW_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call SMWClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub SMWReadNext(Optional DBNo As Integer = 0)
    If Not SMW_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           SMW_RDSTS = False
        Else
           Call ItemsFromSMW(DBNo)
           SMW_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub SMWClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        SMW_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( SMW )
'*----------------------------------------------*'
Public Sub ItemsClearSMW()
    With SMW
	.処理端末 = ""
	.区分 = ""
	.区分名称 = ""
	.契約先 = ""
	.契約先名 = ""
	.品番 = ""
	.材料管理番号 = 0
	.仕入先 = ""
	.仕入先名 = ""
	.入荷数 = 0
	.単価 = 0
	.有償支給 = 0
	.外注材料 = 0
	.外注購入 = 0
	.有償支給計 = 0
	.外注材料計 = 0
	.外注購入計 = 0
	.実績換算値 = 0
	.金額 = 0
	.自社工区 = 0
	.受入 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function SMWInsert(Optional DBNo As Integer) As Boolean
    SMWInsert = False
    With SMW
        gSL_Select = "Insert Into 仕入明細書ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,区分,区分名称,"
        gSL_Select = gSL_Select & "契約先,契約先名,品番,"
        gSL_Select = gSL_Select & "材料管理番号,仕入先,仕入先名,"
        gSL_Select = gSL_Select & "入荷数,単価,有償支給,"
        gSL_Select = gSL_Select & "外注材料,外注購入,有償支給計,"
        gSL_Select = gSL_Select & "外注材料計,外注購入計,実績換算値,"
        gSL_Select = gSL_Select & "金額,自社工区,受入"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.区分名称) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先名) & "',"		
        gSL_Select = gSL_Select & "" & .入荷数 & ","		
        gSL_Select = gSL_Select & "" & .単価 & ","		
        gSL_Select = gSL_Select & "" & .有償支給 & ","		
        gSL_Select = gSL_Select & "" & .外注材料 & ","		
        gSL_Select = gSL_Select & "" & .外注購入 & ","		
        gSL_Select = gSL_Select & "" & .有償支給計 & ","		
        gSL_Select = gSL_Select & "" & .外注材料計 & ","		
        gSL_Select = gSL_Select & "" & .外注購入計 & ","		
        gSL_Select = gSL_Select & "" & .実績換算値 & ","		
        gSL_Select = gSL_Select & "" & .金額 & ","		
        gSL_Select = gSL_Select & "" & .自社工区 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SMWInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function SMWUpdate(Optional DBNo As Integer) As Boolean
    SMWUpdate = False	
    With SMW
        gSL_Select = "Update 仕入明細書ワーク  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "区分 = '" & RTrim(.区分) & "',"
        gSL_Select = gSL_Select & "区分名称 = '" & RTrim(.区分名称) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "契約先名 = '" & RTrim(.契約先名) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "仕入先名 = '" & RTrim(.仕入先名) & "',"
        gSL_Select = gSL_Select & "入荷数 = " & .入荷数 & ","
        gSL_Select = gSL_Select & "単価 = " & .単価 & ","
        gSL_Select = gSL_Select & "有償支給 = " & .有償支給 & ","
        gSL_Select = gSL_Select & "外注材料 = " & .外注材料 & ","
        gSL_Select = gSL_Select & "外注購入 = " & .外注購入 & ","
        gSL_Select = gSL_Select & "有償支給計 = " & .有償支給計 & ","
        gSL_Select = gSL_Select & "外注材料計 = " & .外注材料計 & ","
        gSL_Select = gSL_Select & "外注購入計 = " & .外注購入計 & ","
        gSL_Select = gSL_Select & "実績換算値 = " & .実績換算値 & ","
        gSL_Select = gSL_Select & "金額 = " & .金額 & ","
        gSL_Select = gSL_Select & "自社工区 = " & .自社工区 & ","
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & " Where 処理端末 = '" & RTrim(.処理端末) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    SMWUpdate = True
End Function

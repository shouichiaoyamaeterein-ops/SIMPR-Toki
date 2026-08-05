Attribute VB_Name = "かんばん設定テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   かんばん設定テーブル
'*
'*   モジュールID  : かんばん設定テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/02/26  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_かんばん設定テーブル
        かんばん種類  As Byte
        かんばん名    As String
        ﾗﾍﾞﾙﾋﾟｯﾁ      As Long
        有効印字幅    As Long
        有効印字長    As Long
        印字スピード  As Byte
        カット間隔    As Long
        センサ        As Byte
        リボン        As Byte
        ホッパ        As Byte
        バー種類      As String
        誤り訂正      As String
        セル幅        As Byte
        モード        As String
        回転          As Byte
        モデル        As Byte
'        作成日        As Date
        作成者        As String
'        作成端末      As String
'        更新日        As Date
        更新者        As String
'        更新端末      As String
        件数          As Long
    End Type

    Public KST As T_かんばん設定テーブル

    Public KST_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKST(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKST	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KST
            Select Case adoCol.Name
                Case "かんばん種類":  If Not IsNull(adoCol.Value) Then .かんばん種類 = adoCol.Value
                Case "かんばん名":    If Not IsNull(adoCol.Value) Then .かんばん名   = adoCol.Value
                Case "ﾗﾍﾞﾙﾋﾟｯﾁ":      If Not IsNull(adoCol.Value) Then .ﾗﾍﾞﾙﾋﾟｯﾁ     = adoCol.Value
                Case "有効印字幅":    If Not IsNull(adoCol.Value) Then .有効印字幅   = adoCol.Value
                Case "有効印字長":    If Not IsNull(adoCol.Value) Then .有効印字長   = adoCol.Value
                Case "印字スピード":  If Not IsNull(adoCol.Value) Then .印字スピード = adoCol.Value
                Case "カット間隔":    If Not IsNull(adoCol.Value) Then .カット間隔   = adoCol.Value
                Case "センサ":        If Not IsNull(adoCol.Value) Then .センサ       = adoCol.Value
                Case "リボン":        If Not IsNull(adoCol.Value) Then .リボン       = adoCol.Value
                Case "ホッパ":        If Not IsNull(adoCol.Value) Then .ホッパ       = adoCol.Value
                Case "バー種類":      If Not IsNull(adoCol.Value) Then .バー種類     = adoCol.Value
                Case "誤り訂正":      If Not IsNull(adoCol.Value) Then .誤り訂正     = adoCol.Value
                Case "セル幅":        If Not IsNull(adoCol.Value) Then .セル幅       = adoCol.Value
                Case "モード":        If Not IsNull(adoCol.Value) Then .モード       = adoCol.Value
                Case "回転":          If Not IsNull(adoCol.Value) Then .回転         = adoCol.Value
                Case "モデル":        If Not IsNull(adoCol.Value) Then .モデル       = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者       = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KSTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KSTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From かんばん設定テーブル"
        RTxtSql = RTxtSql &" Where かんばん種類 = " & KST.かんばん種類 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KSTRead = True
           Call ItemsFromKST(DBNo)
	   KST_RDSTS = True
        else
	   KST_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KSTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KSTReadNext(Optional DBNo As Integer = 0)
    If Not KST_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KST_RDSTS = False
        Else
           Call ItemsFromKST(DBNo)
           KST_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KSTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KST_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KST )
'*----------------------------------------------*'
Public Sub ItemsClearKST()
    With KST
	.かんばん種類 = 0
	.かんばん名 = ""
	.ﾗﾍﾞﾙﾋﾟｯﾁ = 0
	.有効印字幅 = 0
	.有効印字長 = 0
	.印字スピード = 0
	.カット間隔 = 0
	.センサ = 0
	.リボン = 0
	.ホッパ = 0
	.バー種類 = ""
	.誤り訂正 = ""
	.セル幅 = 0
	.モード = ""
	.回転 = 0
	.モデル = 0
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
'	.更新日 = 0
	.更新者 = ""
'	.更新端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KSTInsert(Optional DBNo As Integer) As Boolean
    KSTInsert = True
    With KST
        gSL_Select = "Insert Into かんばん設定テーブル  ( "
        gSL_Select = gSL_Select & "かんばん種類,かんばん名,ﾗﾍﾞﾙﾋﾟｯﾁ,"
        gSL_Select = gSL_Select & "有効印字幅,有効印字長,印字スピード,"
        gSL_Select = gSL_Select & "カット間隔,センサ,リボン,"
        gSL_Select = gSL_Select & "ホッパ,バー種類,誤り訂正,"
        gSL_Select = gSL_Select & "セル幅,モード,回転,"
        gSL_Select = gSL_Select & "モデル,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .かんばん種類 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.かんばん名) & "',"		
        gSL_Select = gSL_Select & "" & .ﾗﾍﾞﾙﾋﾟｯﾁ & ","		
        gSL_Select = gSL_Select & "" & .有効印字幅 & ","		
        gSL_Select = gSL_Select & "" & .有効印字長 & ","		
        gSL_Select = gSL_Select & "" & .印字スピード & ","		
        gSL_Select = gSL_Select & "" & .カット間隔 & ","		
        gSL_Select = gSL_Select & "" & .センサ & ","		
        gSL_Select = gSL_Select & "" & .リボン & ","		
        gSL_Select = gSL_Select & "" & .ホッパ & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.バー種類) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.誤り訂正) & "',"		
        gSL_Select = gSL_Select & "" & .セル幅 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.モード) & "',"		
        gSL_Select = gSL_Select & "" & .回転 & ","		
        gSL_Select = gSL_Select & "" & .モデル & ","		
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
    KSTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KSTUpdate(Optional DBNo As Integer) As Boolean
    KSTUpdate = False	
    With KST
        gSL_Select = "Update かんばん設定テーブル  Set "
        gSL_Select = gSL_Select & "かんばん種類 = " & .かんばん種類 & ","
        gSL_Select = gSL_Select & "かんばん名 = '" & RTrim(.かんばん名) & "',"
        gSL_Select = gSL_Select & "ﾗﾍﾞﾙﾋﾟｯﾁ = " & .ﾗﾍﾞﾙﾋﾟｯﾁ & ","
        gSL_Select = gSL_Select & "有効印字幅 = " & .有効印字幅 & ","
        gSL_Select = gSL_Select & "有効印字長 = " & .有効印字長 & ","
        gSL_Select = gSL_Select & "印字スピード = " & .印字スピード & ","
        gSL_Select = gSL_Select & "カット間隔 = " & .カット間隔 & ","
        gSL_Select = gSL_Select & "センサ = " & .センサ & ","
        gSL_Select = gSL_Select & "リボン = " & .リボン & ","
        gSL_Select = gSL_Select & "ホッパ = " & .ホッパ & ","
        gSL_Select = gSL_Select & "バー種類 = '" & RTrim(.バー種類) & "',"
        gSL_Select = gSL_Select & "誤り訂正 = '" & RTrim(.誤り訂正) & "',"
        gSL_Select = gSL_Select & "セル幅 = " & .セル幅 & ","
        gSL_Select = gSL_Select & "モード = '" & RTrim(.モード) & "',"
        gSL_Select = gSL_Select & "回転 = " & .回転 & ","
        gSL_Select = gSL_Select & "モデル = " & .モデル & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where かんばん種類 = " & .かんばん種類 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KSTUpdate = True
End Function

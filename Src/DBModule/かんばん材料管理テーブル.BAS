Attribute VB_Name = "かんばん材料管理テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   かんばん材料管理テーブル
'*
'*   モジュールID  : かんばん材料管理テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/02/21  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_かんばん材料管理テーブル
        材料管理番号  As Long
        取引先        As String
        受入          As String
        発行NO        As Integer
        流動日付      As String
        発行区分      As Byte
        版数          As Byte
        有効指示日    As String
        無効指示日    As String
        回収日        As String
        収容区分      As Byte
'        作成日        As Date
        作成者        As String
'        作成端末      As String
'        更新日        As Date
        更新者        As String
'        更新端末      As String
        件数          As Long
    End Type

    Public KZK As T_かんばん材料管理テーブル

    Public KZK_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKZK(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKZK	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KZK
            Select Case adoCol.Name
                Case "材料管理番号":  If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "取引先":        If Not IsNull(adoCol.Value) Then .取引先       = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "発行NO":        If Not IsNull(adoCol.Value) Then .発行NO       = adoCol.Value
                Case "流動日付":      If Not IsNull(adoCol.Value) Then .流動日付     = adoCol.Value
                Case "発行区分":      If Not IsNull(adoCol.Value) Then .発行区分     = adoCol.Value
                Case "版数":          If Not IsNull(adoCol.Value) Then .版数         = adoCol.Value
                Case "有効指示日":    If Not IsNull(adoCol.Value) Then .有効指示日   = adoCol.Value
                Case "無効指示日":    If Not IsNull(adoCol.Value) Then .無効指示日   = adoCol.Value
                Case "回収日":        If Not IsNull(adoCol.Value) Then .回収日       = adoCol.Value
                Case "収容区分":      If Not IsNull(adoCol.Value) Then .収容区分     = adoCol.Value
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
Public Function KZKRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KZKRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From かんばん材料管理テーブル"
        RTxtSql = RTxtSql &" Where 材料管理番号 = " & KZK.材料管理番号 & ""
        RTxtSql = RTxtSql &" And   取引先 = '" & RTrim(KZK.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(KZK.受入) & "'"
        RTxtSql = RTxtSql &" And   発行NO = " & KZK.発行NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KZKRead = True
           Call ItemsFromKZK(DBNo)
	   KZK_RDSTS = True
        else
	   KZK_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KZKClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KZKReadNext(Optional DBNo As Integer = 0)
    If Not KZK_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KZK_RDSTS = False
        Else
           Call ItemsFromKZK(DBNo)
           KZK_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KZKClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KZK_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KZK )
'*----------------------------------------------*'
Public Sub ItemsClearKZK()
    With KZK
	.材料管理番号 = 0
	.取引先 = ""
	.受入 = ""
	.発行NO = 0
	.流動日付 = ""
	.発行区分 = 0
	.版数 = 0
	.有効指示日 = ""
	.無効指示日 = ""
	.回収日 = ""
	.収容区分 = 0
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
Public Function KZKInsert(Optional DBNo As Integer) As Boolean
    KZKInsert = True
    With KZK
        gSL_Select = "Insert Into かんばん材料管理テーブル  ( "
        gSL_Select = gSL_Select & "材料管理番号,取引先,受入,"
        gSL_Select = gSL_Select & "発行NO,流動日付,発行区分,"
        gSL_Select = gSL_Select & "版数,有効指示日,無効指示日,"
        gSL_Select = gSL_Select & "回収日,収容区分,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "" & .発行NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.流動日付) & "',"		
        gSL_Select = gSL_Select & "" & .発行区分 & ","		
        gSL_Select = gSL_Select & "" & .版数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.有効指示日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.無効指示日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.回収日) & "',"		
        gSL_Select = gSL_Select & "" & .収容区分 & ","		
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
    KZKInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KZKUpdate(Optional DBNo As Integer) As Boolean
    KZKUpdate = False	
    With KZK
        gSL_Select = "Update かんばん材料管理テーブル  Set "
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "発行NO = " & .発行NO & ","
        gSL_Select = gSL_Select & "流動日付 = '" & RTrim(.流動日付) & "',"
        gSL_Select = gSL_Select & "発行区分 = " & .発行区分 & ","
        gSL_Select = gSL_Select & "版数 = " & .版数 & ","
        gSL_Select = gSL_Select & "有効指示日 = '" & RTrim(.有効指示日) & "',"
        gSL_Select = gSL_Select & "無効指示日 = '" & RTrim(.無効指示日) & "',"
        gSL_Select = gSL_Select & "回収日 = '" & RTrim(.回収日) & "',"
        gSL_Select = gSL_Select & "収容区分 = " & .収容区分 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 材料管理番号 = " & .材料管理番号 & ""
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 発行NO = " & .発行NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KZKUpdate = True
End Function

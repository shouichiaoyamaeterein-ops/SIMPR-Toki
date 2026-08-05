Attribute VB_Name = "材料代用マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料代用マスタ
'*
'*   モジュールID  : 材料代用マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2009/09/23  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料代用マスタ
        材料管理番号  As Long
        品番          As String
        板厚          As Currency
        幅            As Currency
        長さ          As Currency
'        作成日        As Date
        作成者        As String
'        作成端末      As String
        件数          As Long
    End Type

    Public ZDA As T_材料代用マスタ

    Public ZDA_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZDA(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZDA	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZDA
            Select Case adoCol.Name
                Case "材料管理番号":  If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "板厚":          If Not IsNull(adoCol.Value) Then .板厚         = adoCol.Value
                Case "幅":            If Not IsNull(adoCol.Value) Then .幅           = adoCol.Value
                Case "長さ":          If Not IsNull(adoCol.Value) Then .長さ         = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function ZDARead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZDARead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 材料代用マスタ"
        RTxtSql = RTxtSql &" Where 材料管理番号 = " & ZDA.材料管理番号 & ""
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(ZDA.品番) & "'"
        RTxtSql = RTxtSql &" And   板厚 = " & ZDA.板厚 & ""
        RTxtSql = RTxtSql &" And   幅 = " & ZDA.幅 & ""
        RTxtSql = RTxtSql &" And   長さ = " & ZDA.長さ & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZDARead = True
           Call ItemsFromZDA(DBNo)
	   ZDA_RDSTS = True
        else
	   ZDA_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZDAClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZDAReadNext(Optional DBNo As Integer = 0)
    If Not ZDA_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZDA_RDSTS = False
        Else
           Call ItemsFromZDA(DBNo)
           ZDA_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZDAClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZDA_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZDA )
'*----------------------------------------------*'
Public Sub ItemsClearZDA()
    With ZDA
	.材料管理番号 = 0
	.品番 = ""
	.板厚 = 0
	.幅 = 0
	.長さ = 0
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZDAInsert(Optional DBNo As Integer) As Boolean
    ZDAInsert = False
    With ZDA
        gSL_Select = "Insert Into 材料代用マスタ  ( "
        gSL_Select = gSL_Select & "材料管理番号,品番,板厚,"
        gSL_Select = gSL_Select & "幅,長さ,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "" & .板厚 & ","		
        gSL_Select = gSL_Select & "" & .幅 & ","		
        gSL_Select = gSL_Select & "" & .長さ & ","		
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
    ZDAInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZDAUpdate(Optional DBNo As Integer) As Boolean
    ZDAUpdate = False	
    With ZDA
        gSL_Select = "Update 材料代用マスタ  Set "
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "板厚 = " & .板厚 & ","
        gSL_Select = gSL_Select & "幅 = " & .幅 & ","
        gSL_Select = gSL_Select & "長さ = " & .長さ & ""
        gSL_Select = gSL_Select & " Where 材料管理番号 = " & .材料管理番号 & ""
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 板厚 = " & .板厚 & ""
        gSL_Select = gSL_Select & "   And 幅 = " & .幅 & ""
        gSL_Select = gSL_Select & "   And 長さ = " & .長さ & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZDAUpdate = True
End Function

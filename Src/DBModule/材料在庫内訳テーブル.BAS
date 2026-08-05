Attribute VB_Name = "材料在庫内訳テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料在庫内訳テーブル
'*
'*   モジュールID  : 材料在庫内訳テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2010/01/14  By Cis(HP-DX7500)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料在庫内訳テーブル
        年月度        As String
        材料管理番号  As Long
        置場          As String
        材質          As String
        寸法区分      As Byte
        板厚          As Currency
        幅            As Currency
        長さ          As Currency
        在庫数        As Currency
        実棚日        As String
'        作成日        As Date
        作成者        As String
'        作成端末      As String
        件数          As Long
    End Type

    Public ZZU As T_材料在庫内訳テーブル

    Public ZZU_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZZU(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZZU	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZZU
            Select Case adoCol.Name
                Case "年月度":        If Not IsNull(adoCol.Value) Then .年月度       = adoCol.Value
                Case "材料管理番号":  If Not IsNull(adoCol.Value) Then .材料管理番号 = adoCol.Value
                Case "置場":          If Not IsNull(adoCol.Value) Then .置場         = adoCol.Value
                Case "材質":          If Not IsNull(adoCol.Value) Then .材質         = adoCol.Value
                Case "寸法区分":      If Not IsNull(adoCol.Value) Then .寸法区分     = adoCol.Value
                Case "板厚":          If Not IsNull(adoCol.Value) Then .板厚         = adoCol.Value
                Case "幅":            If Not IsNull(adoCol.Value) Then .幅           = adoCol.Value
                Case "長さ":          If Not IsNull(adoCol.Value) Then .長さ         = adoCol.Value
                Case "在庫数":        If Not IsNull(adoCol.Value) Then .在庫数       = adoCol.Value
                Case "実棚日":        If Not IsNull(adoCol.Value) Then .実棚日       = adoCol.Value
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
Public Function ZZURead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZZURead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 材料在庫内訳テーブル"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZZURead = True
           Call ItemsFromZZU(DBNo)
	   ZZU_RDSTS = True
        else
	   ZZU_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZZUClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZZUReadNext(Optional DBNo As Integer = 0)
    If Not ZZU_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZZU_RDSTS = False
        Else
           Call ItemsFromZZU(DBNo)
           ZZU_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZZUClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZZU_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZZU )
'*----------------------------------------------*'
Public Sub ItemsClearZZU()
    With ZZU
	.年月度 = ""
	.材料管理番号 = 0
	.置場 = ""
	.材質 = ""
	.寸法区分 = 0
	.板厚 = 0
	.幅 = 0
	.長さ = 0
	.在庫数 = 0
	.実棚日 = ""
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZZUInsert(Optional DBNo As Integer) As Boolean
    ZZUInsert = False
    With ZZU
        gSL_Select = "Insert Into 材料在庫内訳テーブル  ( "
        gSL_Select = gSL_Select & "年月度,材料管理番号,置場,"
        gSL_Select = gSL_Select & "材質,寸法区分,板厚,"
        gSL_Select = gSL_Select & "幅,長さ,在庫数,"
        gSL_Select = gSL_Select & "実棚日,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.年月度) & "',"		
        gSL_Select = gSL_Select & "" & .材料管理番号 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.置場) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"		
        gSL_Select = gSL_Select & "" & .寸法区分 & ","		
        gSL_Select = gSL_Select & "" & .板厚 & ","		
        gSL_Select = gSL_Select & "" & .幅 & ","		
        gSL_Select = gSL_Select & "" & .長さ & ","		
        gSL_Select = gSL_Select & "" & .在庫数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.実棚日) & "',"		
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
    ZZUInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZZUUpdate(Optional DBNo As Integer) As Boolean
    ZZUUpdate = False	
    With ZZU
        gSL_Select = "Update 材料在庫内訳テーブル  Set "
        gSL_Select = gSL_Select & "年月度 = '" & RTrim(.年月度) & "',"
        gSL_Select = gSL_Select & "材料管理番号 = " & .材料管理番号 & ","
        gSL_Select = gSL_Select & "置場 = '" & RTrim(.置場) & "',"
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "寸法区分 = " & .寸法区分 & ","
        gSL_Select = gSL_Select & "板厚 = " & .板厚 & ","
        gSL_Select = gSL_Select & "幅 = " & .幅 & ","
        gSL_Select = gSL_Select & "長さ = " & .長さ & ","
        gSL_Select = gSL_Select & "在庫数 = " & .在庫数 & ","
        gSL_Select = gSL_Select & "実棚日 = '" & RTrim(.実棚日) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZZUUpdate = True
End Function

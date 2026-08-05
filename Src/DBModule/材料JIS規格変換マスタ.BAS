Attribute VB_Name = "材料JIS規格変換マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   材料JIS規格変換マスタ
'*
'*   モジュールID  : 材料JIS規格変換マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2009/09/23  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_材料JIS規格変換マスタ
        材質          As String
        JIS規格       As String
'        作成日        As Date
        作成者        As String
'        作成端末      As String
        件数          As Long
    End Type

    Public ZJS As T_材料JIS規格変換マスタ

    Public ZJS_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromZJS(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearZJS	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With ZJS
            Select Case adoCol.Name
                Case "材質":          If Not IsNull(adoCol.Value) Then .材質         = adoCol.Value
                Case "JIS規格":       If Not IsNull(adoCol.Value) Then .JIS規格      = adoCol.Value
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
Public Function ZJSRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    ZJSRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 材料JIS規格変換マスタ"
        RTxtSql = RTxtSql &" Where 材質 = '" & RTrim(ZJS.材質) & "'"
        RTxtSql = RTxtSql &" And   JIS規格 = '" & RTrim(ZJS.JIS規格) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           ZJSRead = True
           Call ItemsFromZJS(DBNo)
	   ZJS_RDSTS = True
        else
	   ZJS_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call ZJSClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub ZJSReadNext(Optional DBNo As Integer = 0)
    If Not ZJS_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           ZJS_RDSTS = False
        Else
           Call ItemsFromZJS(DBNo)
           ZJS_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub ZJSClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        ZJS_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( ZJS )
'*----------------------------------------------*'
Public Sub ItemsClearZJS()
    With ZJS
	.材質 = ""
	.JIS規格 = ""
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function ZJSInsert(Optional DBNo As Integer) As Boolean
    ZJSInsert = False
    With ZJS
        gSL_Select = "Insert Into 材料JIS規格変換マスタ  ( "
        gSL_Select = gSL_Select & "材質,JIS規格,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.材質) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.JIS規格) & "',"		
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
    ZJSInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function ZJSUpdate(Optional DBNo As Integer) As Boolean
    ZJSUpdate = False	
    With ZJS
        gSL_Select = "Update 材料JIS規格変換マスタ  Set "
        gSL_Select = gSL_Select & "材質 = '" & RTrim(.材質) & "',"
        gSL_Select = gSL_Select & "JIS規格 = '" & RTrim(.JIS規格) & "'"
        gSL_Select = gSL_Select & " Where 材質 = '" & RTrim(.材質) & "'"
        gSL_Select = gSL_Select & "   And JIS規格 = '" & RTrim(.JIS規格) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    ZJSUpdate = True
End Function

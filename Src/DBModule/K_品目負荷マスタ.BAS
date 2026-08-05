Attribute VB_Name = "K_品目負荷マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   K_品目負荷マスタ
'*
'*   モジュールID  : K_品目負荷マスタ.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2010/04/22  By Cis(FMVW5220)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_K_品目負荷マスタ
        品目コード            As String
        手配先コード          As String
        負荷係数1             As String
        負荷係数2             As String
        負荷係数3             As String
        負荷係数4             As String
        負荷係数5             As String
        負荷係数6             As String
        負荷係数7             As String
        負荷係数8             As String
        負荷係数9             As String
        負荷係数10            As String
        負荷係数11            As String
        負荷係数12            As String
        負荷係数13            As String
        負荷係数14            As String
        負荷係数15            As String
        内示負荷数            As String
        出荷負荷数            As String
        生産負荷数_日次       As String
        生産負荷数_月次       As String
        負荷数5               As String
        空き                  As String
        件数          As Long
    End Type

    Public KKF As T_K_品目負荷マスタ

    Public KKF_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKKF(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKKF	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KKF
            Select Case adoCol.Name
                Case "品目コード":            If Not IsNull(adoCol.Value) Then .品目コード           = adoCol.Value
                Case "手配先コード":          If Not IsNull(adoCol.Value) Then .手配先コード         = adoCol.Value
                Case "負荷係数1":             If Not IsNull(adoCol.Value) Then .負荷係数1            = adoCol.Value
                Case "負荷係数2":             If Not IsNull(adoCol.Value) Then .負荷係数2            = adoCol.Value
                Case "負荷係数3":             If Not IsNull(adoCol.Value) Then .負荷係数3            = adoCol.Value
                Case "負荷係数4":             If Not IsNull(adoCol.Value) Then .負荷係数4            = adoCol.Value
                Case "負荷係数5":             If Not IsNull(adoCol.Value) Then .負荷係数5            = adoCol.Value
                Case "負荷係数6":             If Not IsNull(adoCol.Value) Then .負荷係数6            = adoCol.Value
                Case "負荷係数7":             If Not IsNull(adoCol.Value) Then .負荷係数7            = adoCol.Value
                Case "負荷係数8":             If Not IsNull(adoCol.Value) Then .負荷係数8            = adoCol.Value
                Case "負荷係数9":             If Not IsNull(adoCol.Value) Then .負荷係数9            = adoCol.Value
                Case "負荷係数10":            If Not IsNull(adoCol.Value) Then .負荷係数10           = adoCol.Value
                Case "負荷係数11":            If Not IsNull(adoCol.Value) Then .負荷係数11           = adoCol.Value
                Case "負荷係数12":            If Not IsNull(adoCol.Value) Then .負荷係数12           = adoCol.Value
                Case "負荷係数13":            If Not IsNull(adoCol.Value) Then .負荷係数13           = adoCol.Value
                Case "負荷係数14":            If Not IsNull(adoCol.Value) Then .負荷係数14           = adoCol.Value
                Case "負荷係数15":            If Not IsNull(adoCol.Value) Then .負荷係数15           = adoCol.Value
                Case "内示負荷数":            If Not IsNull(adoCol.Value) Then .内示負荷数           = adoCol.Value
                Case "出荷負荷数":            If Not IsNull(adoCol.Value) Then .出荷負荷数           = adoCol.Value
                Case "生産負荷数_日次":       If Not IsNull(adoCol.Value) Then .生産負荷数_日次      = adoCol.Value
                Case "生産負荷数_月次":       If Not IsNull(adoCol.Value) Then .生産負荷数_月次      = adoCol.Value
                Case "負荷数5":               If Not IsNull(adoCol.Value) Then .負荷数5              = adoCol.Value
                Case "空き":                  If Not IsNull(adoCol.Value) Then .空き                 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KKFRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KKFRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From K_品目負荷マスタ"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KKFRead = True
           Call ItemsFromKKF(DBNo)
	   KKF_RDSTS = True
        else
	   KKF_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KKFClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KKFReadNext(Optional DBNo As Integer = 0)
    If Not KKF_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KKF_RDSTS = False
        Else
           Call ItemsFromKKF(DBNo)
           KKF_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KKFClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KKF_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KKF )
'*----------------------------------------------*'
Public Sub ItemsClearKKF()
    With KKF
	.品目コード = ""
	.手配先コード = ""
	.負荷係数1 = ""
	.負荷係数2 = ""
	.負荷係数3 = ""
	.負荷係数4 = ""
	.負荷係数5 = ""
	.負荷係数6 = ""
	.負荷係数7 = ""
	.負荷係数8 = ""
	.負荷係数9 = ""
	.負荷係数10 = ""
	.負荷係数11 = ""
	.負荷係数12 = ""
	.負荷係数13 = ""
	.負荷係数14 = ""
	.負荷係数15 = ""
	.内示負荷数 = ""
	.出荷負荷数 = ""
	.生産負荷数_日次 = ""
	.生産負荷数_月次 = ""
	.負荷数5 = ""
	.空き = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KKFInsert(Optional DBNo As Integer) As Boolean
    KKFInsert = False
    With KKF
        gSL_Select = "Insert Into K_品目負荷マスタ  ( "
        gSL_Select = gSL_Select & "品目コード,手配先コード,負荷係数1,"
        gSL_Select = gSL_Select & "負荷係数2,負荷係数3,負荷係数4,"
        gSL_Select = gSL_Select & "負荷係数5,負荷係数6,負荷係数7,"
        gSL_Select = gSL_Select & "負荷係数8,負荷係数9,負荷係数10,"
        gSL_Select = gSL_Select & "負荷係数11,負荷係数12,負荷係数13,"
        gSL_Select = gSL_Select & "負荷係数14,負荷係数15,内示負荷数,"
        gSL_Select = gSL_Select & "出荷負荷数,生産負荷数_日次,生産負荷数_月次,"
        gSL_Select = gSL_Select & "負荷数5,空き"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.品目コード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.手配先コード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.負荷係数1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.負荷係数2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.負荷係数3) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.負荷係数4) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.負荷係数5) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.負荷係数6) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.負荷係数7) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.負荷係数8) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.負荷係数9) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.負荷係数10) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.負荷係数11) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.負荷係数12) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.負荷係数13) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.負荷係数14) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.負荷係数15) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.内示負荷数) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.出荷負荷数) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.生産負荷数_日次) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.生産負荷数_月次) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.負荷数5) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.空き) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KKFInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KKFUpdate(Optional DBNo As Integer) As Boolean
    KKFUpdate = False	
    With KKF
        gSL_Select = "Update K_品目負荷マスタ  Set "
        gSL_Select = gSL_Select & "品目コード = '" & RTrim(.品目コード) & "',"
        gSL_Select = gSL_Select & "手配先コード = '" & RTrim(.手配先コード) & "',"
        gSL_Select = gSL_Select & "負荷係数1 = '" & RTrim(.負荷係数1) & "',"
        gSL_Select = gSL_Select & "負荷係数2 = '" & RTrim(.負荷係数2) & "',"
        gSL_Select = gSL_Select & "負荷係数3 = '" & RTrim(.負荷係数3) & "',"
        gSL_Select = gSL_Select & "負荷係数4 = '" & RTrim(.負荷係数4) & "',"
        gSL_Select = gSL_Select & "負荷係数5 = '" & RTrim(.負荷係数5) & "',"
        gSL_Select = gSL_Select & "負荷係数6 = '" & RTrim(.負荷係数6) & "',"
        gSL_Select = gSL_Select & "負荷係数7 = '" & RTrim(.負荷係数7) & "',"
        gSL_Select = gSL_Select & "負荷係数8 = '" & RTrim(.負荷係数8) & "',"
        gSL_Select = gSL_Select & "負荷係数9 = '" & RTrim(.負荷係数9) & "',"
        gSL_Select = gSL_Select & "負荷係数10 = '" & RTrim(.負荷係数10) & "',"
        gSL_Select = gSL_Select & "負荷係数11 = '" & RTrim(.負荷係数11) & "',"
        gSL_Select = gSL_Select & "負荷係数12 = '" & RTrim(.負荷係数12) & "',"
        gSL_Select = gSL_Select & "負荷係数13 = '" & RTrim(.負荷係数13) & "',"
        gSL_Select = gSL_Select & "負荷係数14 = '" & RTrim(.負荷係数14) & "',"
        gSL_Select = gSL_Select & "負荷係数15 = '" & RTrim(.負荷係数15) & "',"
        gSL_Select = gSL_Select & "内示負荷数 = '" & RTrim(.内示負荷数) & "',"
        gSL_Select = gSL_Select & "出荷負荷数 = '" & RTrim(.出荷負荷数) & "',"
        gSL_Select = gSL_Select & "生産負荷数_日次 = '" & RTrim(.生産負荷数_日次) & "',"
        gSL_Select = gSL_Select & "生産負荷数_月次 = '" & RTrim(.生産負荷数_月次) & "',"
        gSL_Select = gSL_Select & "負荷数5 = '" & RTrim(.負荷数5) & "',"
        gSL_Select = gSL_Select & "空き = '" & RTrim(.空き) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KKFUpdate = True
End Function

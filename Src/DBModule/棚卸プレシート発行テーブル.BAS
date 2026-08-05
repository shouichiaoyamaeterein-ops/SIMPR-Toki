Attribute VB_Name = "棚卸プレシート発行テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   棚卸プレシート発行テーブル
'*
'*   モジュールID  : 棚卸プレシート発行テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2009/12/14  By Cis(FMVW5220)
'*   作成日  :  2010/01/06  By Cis(HP-DX7500)
'*   作成日  :  2010/01/18  By Cis(HP-DX7500)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_棚卸プレシート発行テーブル
        発行年月          As String
        伝票NO            As Long
        伝票行NO          As Long
        在庫場所1         As String
        在庫場所2         As String
        棚番              As String
        背番号            As String
        品番              As String
        取引先            As String
        受入              As String
        工順              As Byte
        在庫数            As Long
        棚卸入力済フラグ  As Byte
        部品区分          As Byte
        自動車ガス        As Byte
'        作成日            As Date
        作成者            As String
'        作成端末          As String
        件数          As Long
'   項目追加(自動)
	表示品番       As String
	表示材料       As String
	材質           As String
	板厚           As Currency
	幅             As Currency
	長さ           As Currency
	在庫場所名1    As String
	在庫場所名2    As String
	表示棚番       As String
	品名           As String
	材料単位       As String
	材料単位名     As String
	プレシートNO   As String
	部署           As String
	機械           As String
	機械NO         As String
	置場           As String
	理論在庫数     As Long
	基準単価       As Double
    End Type

    Public TPH As T_棚卸プレシート発行テーブル

    Public TPH_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTPH(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTPH	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With TPH
            Select Case adoCol.Name
                Case "発行年月":          If Not IsNull(adoCol.Value) Then .発行年月         = adoCol.Value
                Case "伝票NO":            If Not IsNull(adoCol.Value) Then .伝票NO           = adoCol.Value
                Case "伝票行NO":          If Not IsNull(adoCol.Value) Then .伝票行NO         = adoCol.Value
                Case "在庫場所1":         If Not IsNull(adoCol.Value) Then .在庫場所1        = adoCol.Value
                Case "在庫場所2":         If Not IsNull(adoCol.Value) Then .在庫場所2        = adoCol.Value
                Case "棚番":              If Not IsNull(adoCol.Value) Then .棚番             = adoCol.Value
                Case "背番号":            If Not IsNull(adoCol.Value) Then .背番号           = adoCol.Value
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番             = adoCol.Value
                Case "取引先":            If Not IsNull(adoCol.Value) Then .取引先           = adoCol.Value
                Case "受入":              If Not IsNull(adoCol.Value) Then .受入             = adoCol.Value
                Case "工順":              If Not IsNull(adoCol.Value) Then .工順             = adoCol.Value
                Case "在庫数":            If Not IsNull(adoCol.Value) Then .在庫数           = adoCol.Value
                Case "棚卸入力済フラグ":  If Not IsNull(adoCol.Value) Then .棚卸入力済フラグ = adoCol.Value
                Case "部品区分":          If Not IsNull(adoCol.Value) Then .部品区分         = adoCol.Value
                Case "自動車ガス":        If Not IsNull(adoCol.Value) Then .自動車ガス       = adoCol.Value
'                Case "作成日":            If Not IsNull(adoCol.Value) Then .作成日           = adoCol.Value
                Case "作成者":            If Not IsNull(adoCol.Value) Then .作成者           = adoCol.Value
'                Case "作成端末":          If Not IsNull(adoCol.Value) Then .作成端末         = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "表示材料":	If Not IsNull(adoCol.Value) Then .表示材料	= adoCol.Value
		Case "材質":	If Not IsNull(adoCol.Value) Then .材質	= adoCol.Value
		Case "板厚":	If Not IsNull(adoCol.Value) Then .板厚	= adoCol.Value
		Case "幅":	If Not IsNull(adoCol.Value) Then .幅	= adoCol.Value
		Case "長さ":	If Not IsNull(adoCol.Value) Then .長さ	= adoCol.Value
		Case "在庫場所名1":	If Not IsNull(adoCol.Value) Then .在庫場所名1	= adoCol.Value
		Case "在庫場所名2":	If Not IsNull(adoCol.Value) Then .在庫場所名2	= adoCol.Value
		Case "表示棚番":	If Not IsNull(adoCol.Value) Then .表示棚番	= adoCol.Value
		Case "品名":	If Not IsNull(adoCol.Value) Then .品名	= adoCol.Value
		Case "材料単位":	If Not IsNull(adoCol.Value) Then .材料単位	= adoCol.Value
		Case "材料単位名":	If Not IsNull(adoCol.Value) Then .材料単位名	= adoCol.Value
		Case "プレシートNO":	If Not IsNull(adoCol.Value) Then .プレシートNO	= adoCol.Value
		Case "部署":	If Not IsNull(adoCol.Value) Then .部署	= adoCol.Value
		Case "機械":	If Not IsNull(adoCol.Value) Then .機械	= adoCol.Value
		Case "機械NO":	If Not IsNull(adoCol.Value) Then .機械NO	= adoCol.Value
		Case "置場":	If Not IsNull(adoCol.Value) Then .置場	= adoCol.Value
		Case "理論在庫数":	If Not IsNull(adoCol.Value) Then .理論在庫数	= adoCol.Value
		Case "基準単価":	If Not IsNull(adoCol.Value) Then .基準単価	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TPHRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TPHRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From 棚卸プレシート発行テーブル"
        RTxtSql = RTxtSql &" Where 発行年月 = '" & RTrim(TPH.発行年月) & "'"
        RTxtSql = RTxtSql &" And   伝票NO = " & TPH.伝票NO & ""
        RTxtSql = RTxtSql &" And   伝票行NO = " & TPH.伝票行NO & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TPHRead = True
           Call ItemsFromTPH(DBNo)
	   TPH_RDSTS = True
        else
	   TPH_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call TPHClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TPHReadNext(Optional DBNo As Integer = 0)
    If Not TPH_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TPH_RDSTS = False
        Else
           Call ItemsFromTPH(DBNo)
           TPH_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TPHClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TPH_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TPH )
'*----------------------------------------------*'
Public Sub ItemsClearTPH()
    With TPH
	.発行年月 = ""
	.伝票NO = 0
	.伝票行NO = 0
	.在庫場所1 = ""
	.在庫場所2 = ""
	.棚番 = ""
	.背番号 = ""
	.品番 = ""
	.取引先 = ""
	.受入 = ""
	.工順 = 0
	.在庫数 = 0
	.棚卸入力済フラグ = 0
	.部品区分 = 0
	.自動車ガス = 0
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
        .件数 = 0
'   項目追加(自動)
	.表示品番 = ""
	.表示材料 = ""
	.材質 = ""
	.板厚 = 0
	.幅 = 0
	.長さ = 0
	.在庫場所名1 = ""
	.在庫場所名2 = ""
	.表示棚番 = ""
	.品名 = ""
	.材料単位 = ""
	.材料単位名 = ""
	.プレシートNO = ""
	.部署 = ""
	.機械 = ""
	.機械NO = ""
	.置場 = ""
	.理論在庫数 = 0
	.基準単価 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function TPHInsert(Optional DBNo As Integer) As Boolean
    TPHInsert = False
    With TPH
        gSL_Select = "Insert Into 棚卸プレシート発行テーブル  ( "
        gSL_Select = gSL_Select & "発行年月,伝票NO,伝票行NO,"
        gSL_Select = gSL_Select & "在庫場所1,在庫場所2,棚番,"
        gSL_Select = gSL_Select & "背番号,品番,取引先,"
        gSL_Select = gSL_Select & "受入,工順,在庫数,"
        gSL_Select = gSL_Select & "棚卸入力済フラグ,部品区分,自動車ガス,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.発行年月) & "',"		
        gSL_Select = gSL_Select & "" & .伝票NO & ","		
        gSL_Select = gSL_Select & "" & .伝票行NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.在庫場所1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.在庫場所2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.棚番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "" & .工順 & ","		
        gSL_Select = gSL_Select & "" & .在庫数 & ","		
        gSL_Select = gSL_Select & "" & .棚卸入力済フラグ & ","		
        gSL_Select = gSL_Select & "" & .部品区分 & ","		
        gSL_Select = gSL_Select & "" & .自動車ガス & ","		
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
    TPHInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function TPHUpdate(Optional DBNo As Integer) As Boolean
    TPHUpdate = False	
    With TPH
        gSL_Select = "Update 棚卸プレシート発行テーブル  Set "
        gSL_Select = gSL_Select & "発行年月 = '" & RTrim(.発行年月) & "',"
        gSL_Select = gSL_Select & "伝票NO = " & .伝票NO & ","
        gSL_Select = gSL_Select & "伝票行NO = " & .伝票行NO & ","
        gSL_Select = gSL_Select & "在庫場所1 = '" & RTrim(.在庫場所1) & "',"
        gSL_Select = gSL_Select & "在庫場所2 = '" & RTrim(.在庫場所2) & "',"
        gSL_Select = gSL_Select & "棚番 = '" & RTrim(.棚番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "工順 = " & .工順 & ","
        gSL_Select = gSL_Select & "在庫数 = " & .在庫数 & ","
        gSL_Select = gSL_Select & "棚卸入力済フラグ = " & .棚卸入力済フラグ & ","
        gSL_Select = gSL_Select & "部品区分 = " & .部品区分 & ","
        gSL_Select = gSL_Select & "自動車ガス = " & .自動車ガス & ""
        gSL_Select = gSL_Select & " Where 発行年月 = '" & RTrim(.発行年月) & "'"
        gSL_Select = gSL_Select & "   And 伝票NO = " & .伝票NO & ""
        gSL_Select = gSL_Select & "   And 伝票行NO = " & .伝票行NO & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TPHUpdate = True
End Function

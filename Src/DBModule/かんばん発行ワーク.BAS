Attribute VB_Name = "かんばん発行ワーク"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   かんばん発行ワーク
'*
'*   モジュールID  : かんばん発行ワーク.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2006/01/18  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_かんばん発行ワーク
        処理端末      As String
        かんばん種類  As Byte
        品番          As String
        取引先        As String
        受入          As String
        連番          As Long
        発行NO        As Integer
        表示品番      As String
        背番号        As String
        発行種類      As Byte
        発行区分      As Byte
        収容区分      As Byte
        所番地区分    As Byte
        収容器        As String
        収容数        As Currency
        所番地        As String
        発行済        As Byte
        サイクル      As String
        曜日          As String
        職場          As String
'        作成日        As Date
        作成者        As String
        件数          As Long
'   項目追加(自動)
	取引先名       As String
	取引先C        As String
	工場C          As String
	工場名         As String
	アドレス       As String
	機番イメージ   As String
	ポイントイメージ	 As String
	ポイント       As String
	行先名1        As String
	行先名2        As String
	行先名3        As String
	品番イメージ   As String
	チェック1      As String
	チェック2      As String
	納入先         As String
	手配区分       As String
	職場名         As String
	最終発行連番   As Long
	職場C          As String
	職場KJ         As String
	行先1          As String
	行先2          As String
	行先3          As String
	機番           As Long
	次工程C        As String
	次工程工場C    As String
	次工程名       As String
    End Type

    Public KHW As T_かんばん発行ワーク

    Public KHW_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKHW(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKHW	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KHW
            Select Case adoCol.Name
                Case "処理端末":      If Not IsNull(adoCol.Value) Then .処理端末     = adoCol.Value
                Case "かんばん種類":  If Not IsNull(adoCol.Value) Then .かんばん種類 = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "取引先":        If Not IsNull(adoCol.Value) Then .取引先       = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "連番":          If Not IsNull(adoCol.Value) Then .連番         = adoCol.Value
                Case "発行NO":        If Not IsNull(adoCol.Value) Then .発行NO       = adoCol.Value
                Case "表示品番":      If Not IsNull(adoCol.Value) Then .表示品番     = adoCol.Value
                Case "背番号":        If Not IsNull(adoCol.Value) Then .背番号       = adoCol.Value
                Case "発行種類":      If Not IsNull(adoCol.Value) Then .発行種類     = adoCol.Value
                Case "発行区分":      If Not IsNull(adoCol.Value) Then .発行区分     = adoCol.Value
                Case "収容区分":      If Not IsNull(adoCol.Value) Then .収容区分     = adoCol.Value
                Case "所番地区分":    If Not IsNull(adoCol.Value) Then .所番地区分   = adoCol.Value
                Case "収容器":        If Not IsNull(adoCol.Value) Then .収容器       = adoCol.Value
                Case "収容数":        If Not IsNull(adoCol.Value) Then .収容数       = adoCol.Value
                Case "所番地":        If Not IsNull(adoCol.Value) Then .所番地       = adoCol.Value
                Case "発行済":        If Not IsNull(adoCol.Value) Then .発行済       = adoCol.Value
                Case "サイクル":      If Not IsNull(adoCol.Value) Then .サイクル     = adoCol.Value
                Case "曜日":          If Not IsNull(adoCol.Value) Then .曜日         = adoCol.Value
                Case "職場":          If Not IsNull(adoCol.Value) Then .職場         = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "取引先名":	If Not IsNull(adoCol.Value) Then .取引先名	= adoCol.Value
		Case "取引先C":	If Not IsNull(adoCol.Value) Then .取引先C	= adoCol.Value
		Case "工場C":	If Not IsNull(adoCol.Value) Then .工場C	= adoCol.Value
		Case "工場名":	If Not IsNull(adoCol.Value) Then .工場名	= adoCol.Value
		Case "アドレス":	If Not IsNull(adoCol.Value) Then .アドレス	= adoCol.Value
		Case "機番イメージ":	If Not IsNull(adoCol.Value) Then .機番イメージ	= adoCol.Value
		Case "ポイントイメージ":	If Not IsNull(adoCol.Value) Then .ポイントイメージ	= adoCol.Value
		Case "ポイント":	If Not IsNull(adoCol.Value) Then .ポイント	= adoCol.Value
		Case "行先名1":	If Not IsNull(adoCol.Value) Then .行先名1	= adoCol.Value
		Case "行先名2":	If Not IsNull(adoCol.Value) Then .行先名2	= adoCol.Value
		Case "行先名3":	If Not IsNull(adoCol.Value) Then .行先名3	= adoCol.Value
		Case "品番イメージ":	If Not IsNull(adoCol.Value) Then .品番イメージ	= adoCol.Value
		Case "チェック1":	If Not IsNull(adoCol.Value) Then .チェック1	= adoCol.Value
		Case "チェック2":	If Not IsNull(adoCol.Value) Then .チェック2	= adoCol.Value
		Case "納入先":	If Not IsNull(adoCol.Value) Then .納入先	= adoCol.Value
		Case "手配区分":	If Not IsNull(adoCol.Value) Then .手配区分	= adoCol.Value
		Case "職場名":	If Not IsNull(adoCol.Value) Then .職場名	= adoCol.Value
		Case "最終発行連番":	If Not IsNull(adoCol.Value) Then .最終発行連番	= adoCol.Value
		Case "職場C":	If Not IsNull(adoCol.Value) Then .職場C	= adoCol.Value
		Case "職場KJ":	If Not IsNull(adoCol.Value) Then .職場KJ	= adoCol.Value
		Case "行先1":	If Not IsNull(adoCol.Value) Then .行先1	= adoCol.Value
		Case "行先2":	If Not IsNull(adoCol.Value) Then .行先2	= adoCol.Value
		Case "行先3":	If Not IsNull(adoCol.Value) Then .行先3	= adoCol.Value
		Case "機番":	If Not IsNull(adoCol.Value) Then .機番	= adoCol.Value
		Case "次工程C":	If Not IsNull(adoCol.Value) Then .次工程C	= adoCol.Value
		Case "次工程工場C":	If Not IsNull(adoCol.Value) Then .次工程工場C	= adoCol.Value
		Case "次工程名":	If Not IsNull(adoCol.Value) Then .次工程名	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KHWRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KHWRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From かんばん発行ワーク"
        RTxtSql = RTxtSql &" Where 処理端末 = '" & RTrim(KHW.処理端末) & "'"
        RTxtSql = RTxtSql &" And   かんばん種類 = " & KHW.かんばん種類 & ""
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(KHW.品番) & "'"
        RTxtSql = RTxtSql &" And   取引先 = '" & RTrim(KHW.取引先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(KHW.受入) & "'"
        RTxtSql = RTxtSql &" And   連番 = " & KHW.連番 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KHWRead = True
           Call ItemsFromKHW(DBNo)
	   KHW_RDSTS = True
        else
	   KHW_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KHWClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KHWReadNext(Optional DBNo As Integer = 0)
    If Not KHW_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KHW_RDSTS = False
        Else
           Call ItemsFromKHW(DBNo)
           KHW_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KHWClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KHW_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KHW )
'*----------------------------------------------*'
Public Sub ItemsClearKHW()
    With KHW
	.処理端末 = ""
	.かんばん種類 = 0
	.品番 = ""
	.取引先 = ""
	.受入 = ""
	.連番 = 0
	.発行NO = 0
	.表示品番 = ""
	.背番号 = ""
	.発行種類 = 0
	.発行区分 = 0
	.収容区分 = 0
	.所番地区分 = 0
	.収容器 = ""
	.収容数 = 0
	.所番地 = ""
	.発行済 = 0
	.サイクル = ""
	.曜日 = ""
	.職場 = ""
'	.作成日 = 0
	.作成者 = ""
        .件数 = 0
'   項目追加(自動)
	.取引先名 = ""
	.取引先C = ""
	.工場C = ""
	.工場名 = ""
	.アドレス = ""
	.機番イメージ = ""
	.ポイントイメージ = ""
	.ポイント = ""
	.行先名1 = ""
	.行先名2 = ""
	.行先名3 = ""
	.品番イメージ = ""
	.チェック1 = ""
	.チェック2 = ""
	.納入先 = ""
	.手配区分 = ""
	.職場名 = ""
	.最終発行連番 = 0
	.職場C = ""
	.職場KJ = ""
	.行先1 = ""
	.行先2 = ""
	.行先3 = ""
	.機番 = 0
	.次工程C = ""
	.次工程工場C = ""
	.次工程名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KHWInsert(Optional DBNo As Integer) As Boolean
    KHWInsert = True
    With KHW
        gSL_Select = "Insert Into かんばん発行ワーク  ( "
        gSL_Select = gSL_Select & "処理端末,かんばん種類,品番,"
        gSL_Select = gSL_Select & "取引先,受入,連番,"
        gSL_Select = gSL_Select & "発行NO,表示品番,背番号,"
        gSL_Select = gSL_Select & "発行種類,発行区分,収容区分,"
        gSL_Select = gSL_Select & "所番地区分,収容器,収容数,"
        gSL_Select = gSL_Select & "所番地,発行済,サイクル,"
        gSL_Select = gSL_Select & "曜日,職場,"
        gSL_Select = gSL_Select & "作成日,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"		
        gSL_Select = gSL_Select & "" & .かんばん種類 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.取引先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "" & .連番 & ","		
        gSL_Select = gSL_Select & "" & .発行NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.表示品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "" & .発行種類 & ","		
        gSL_Select = gSL_Select & "" & .発行区分 & ","		
        gSL_Select = gSL_Select & "" & .収容区分 & ","		
        gSL_Select = gSL_Select & "" & .所番地区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.収容器) & "',"		
        gSL_Select = gSL_Select & "" & .収容数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.所番地) & "',"		
        gSL_Select = gSL_Select & "" & .発行済 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.サイクル) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.曜日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.職場) & "',"		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KHWInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KHWUpdate(Optional DBNo As Integer) As Boolean
    KHWUpdate = False	
    With KHW
        gSL_Select = "Update かんばん発行ワーク  Set "
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "かんばん種類 = " & .かんばん種類 & ","
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "取引先 = '" & RTrim(.取引先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "連番 = " & .連番 & ","
        gSL_Select = gSL_Select & "発行NO = " & .発行NO & ","
        gSL_Select = gSL_Select & "表示品番 = '" & RTrim(.表示品番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "発行種類 = " & .発行種類 & ","
        gSL_Select = gSL_Select & "発行区分 = " & .発行区分 & ","
        gSL_Select = gSL_Select & "収容区分 = " & .収容区分 & ","
        gSL_Select = gSL_Select & "所番地区分 = " & .所番地区分 & ","
        gSL_Select = gSL_Select & "収容器 = '" & RTrim(.収容器) & "',"
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "所番地 = '" & RTrim(.所番地) & "',"
        gSL_Select = gSL_Select & "発行済 = " & .発行済 & ","
        gSL_Select = gSL_Select & "サイクル = '" & RTrim(.サイクル) & "',"
        gSL_Select = gSL_Select & "曜日 = '" & RTrim(.曜日) & "',"
        gSL_Select = gSL_Select & "職場 = '" & RTrim(.職場) & "'"
        gSL_Select = gSL_Select & " Where 処理端末 = '" & RTrim(.処理端末) & "'"
        gSL_Select = gSL_Select & "   And かんばん種類 = " & .かんばん種類 & ""
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 連番 = " & .連番 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KHWUpdate = True
End Function

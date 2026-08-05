Attribute VB_Name = "Eかんばん変換テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   eかんばん変換テーブル
'*
'*   モジュールID  : Eかんばん変換テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2008/07/16  By Cis(FMV-W5240U)
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_eかんばん変換テーブル
        処理日                As String
        処理回数              As Long
        処理区分              As Byte
        ファイル名            As String
        納入先                As String
        受入                  As String
        納入番号              As String
        納入日                As String
        納入便                As String
        品番                  As String
        便内分割区分          As String
        退避連番              As Long
        受注管理番号          As Long
        受注管理行番号        As Byte
        出荷指示書NO          As Long
        出荷指示書行NO        As Byte
        受注日                As String
        受注便                As String
        背番号                As String
        収容数                As Long
        かんばん枚数          As Long
        納入数                As Long
        打切端数              As Long
        打切区分              As Byte
        打切残数              As Long
        出荷数                As Long
        挽回日                As String
        挽回便                As String
        ERR重複               As Byte
        ERR収容数             As Byte
        ERR背番号             As Byte
        ERR受注               As Byte
        品番仮                As Byte
        品取仮                As Byte
        取引先仮              As Byte
        受入仮                As Byte
        納入先名              As String
        品名                  As String
        行連番                As Integer
'        作成日                As Date
'        作成端末              As String
        件数          As Long
'   項目追加(自動)
	表示品番       As String
	納入先名称     As String
    End Type

    Public EHT As T_eかんばん変換テーブル

    Public EHT_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromEHT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearEHT	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With EHT
            Select Case adoCol.Name
                Case "処理日":                If Not IsNull(adoCol.Value) Then .処理日               = adoCol.Value
                Case "処理回数":              If Not IsNull(adoCol.Value) Then .処理回数             = adoCol.Value
                Case "処理区分":              If Not IsNull(adoCol.Value) Then .処理区分             = adoCol.Value
                Case "ファイル名":            If Not IsNull(adoCol.Value) Then .ファイル名           = adoCol.Value
                Case "納入先":                If Not IsNull(adoCol.Value) Then .納入先               = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入                 = adoCol.Value
                Case "納入番号":              If Not IsNull(adoCol.Value) Then .納入番号             = adoCol.Value
                Case "納入日":                If Not IsNull(adoCol.Value) Then .納入日               = adoCol.Value
                Case "納入便":                If Not IsNull(adoCol.Value) Then .納入便               = adoCol.Value
                Case "品番":                  If Not IsNull(adoCol.Value) Then .品番                 = adoCol.Value
                Case "便内分割区分":          If Not IsNull(adoCol.Value) Then .便内分割区分         = adoCol.Value
                Case "退避連番":              If Not IsNull(adoCol.Value) Then .退避連番             = adoCol.Value
                Case "受注管理番号":          If Not IsNull(adoCol.Value) Then .受注管理番号         = adoCol.Value
                Case "受注管理行番号":        If Not IsNull(adoCol.Value) Then .受注管理行番号       = adoCol.Value
                Case "出荷指示書NO":          If Not IsNull(adoCol.Value) Then .出荷指示書NO         = adoCol.Value
                Case "出荷指示書行NO":        If Not IsNull(adoCol.Value) Then .出荷指示書行NO       = adoCol.Value
                Case "受注日":                If Not IsNull(adoCol.Value) Then .受注日               = adoCol.Value
                Case "受注便":                If Not IsNull(adoCol.Value) Then .受注便               = adoCol.Value
                Case "背番号":                If Not IsNull(adoCol.Value) Then .背番号               = adoCol.Value
                Case "収容数":                If Not IsNull(adoCol.Value) Then .収容数               = adoCol.Value
                Case "かんばん枚数":          If Not IsNull(adoCol.Value) Then .かんばん枚数         = adoCol.Value
                Case "納入数":                If Not IsNull(adoCol.Value) Then .納入数               = adoCol.Value
                Case "打切端数":              If Not IsNull(adoCol.Value) Then .打切端数             = adoCol.Value
                Case "打切区分":              If Not IsNull(adoCol.Value) Then .打切区分             = adoCol.Value
                Case "打切残数":              If Not IsNull(adoCol.Value) Then .打切残数             = adoCol.Value
                Case "出荷数":                If Not IsNull(adoCol.Value) Then .出荷数               = adoCol.Value
                Case "挽回日":                If Not IsNull(adoCol.Value) Then .挽回日               = adoCol.Value
                Case "挽回便":                If Not IsNull(adoCol.Value) Then .挽回便               = adoCol.Value
                Case "ERR重複":               If Not IsNull(adoCol.Value) Then .ERR重複              = adoCol.Value
                Case "ERR収容数":             If Not IsNull(adoCol.Value) Then .ERR収容数            = adoCol.Value
                Case "ERR背番号":             If Not IsNull(adoCol.Value) Then .ERR背番号            = adoCol.Value
                Case "ERR受注":               If Not IsNull(adoCol.Value) Then .ERR受注              = adoCol.Value
                Case "品番仮":                If Not IsNull(adoCol.Value) Then .品番仮               = adoCol.Value
                Case "品取仮":                If Not IsNull(adoCol.Value) Then .品取仮               = adoCol.Value
                Case "取引先仮":              If Not IsNull(adoCol.Value) Then .取引先仮             = adoCol.Value
                Case "受入仮":                If Not IsNull(adoCol.Value) Then .受入仮               = adoCol.Value
                Case "納入先名":              If Not IsNull(adoCol.Value) Then .納入先名             = adoCol.Value
                Case "品名":                  If Not IsNull(adoCol.Value) Then .品名                 = adoCol.Value
                Case "行連番":                If Not IsNull(adoCol.Value) Then .行連番               = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "表示品番":	If Not IsNull(adoCol.Value) Then .表示品番	= adoCol.Value
		Case "納入先名称":	If Not IsNull(adoCol.Value) Then .納入先名称	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function EHTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    EHTRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From eかんばん変換テーブル"
        RTxtSql = RTxtSql &" Where 処理日 = '" & RTrim(EHT.処理日) & "'"
        RTxtSql = RTxtSql &" And   処理回数 = " & EHT.処理回数 & ""
        RTxtSql = RTxtSql &" And   処理区分 = " & EHT.処理区分 & ""
        RTxtSql = RTxtSql &" And   ファイル名 = '" & RTrim(EHT.ファイル名) & "'"
        RTxtSql = RTxtSql &" And   納入先 = '" & RTrim(EHT.納入先) & "'"
        RTxtSql = RTxtSql &" And   受入 = '" & RTrim(EHT.受入) & "'"
        RTxtSql = RTxtSql &" And   納入番号 = '" & RTrim(EHT.納入番号) & "'"
        RTxtSql = RTxtSql &" And   納入日 = '" & RTrim(EHT.納入日) & "'"
        RTxtSql = RTxtSql &" And   納入便 = '" & RTrim(EHT.納入便) & "'"
        RTxtSql = RTxtSql &" And   品番 = '" & RTrim(EHT.品番) & "'"
        RTxtSql = RTxtSql &" And   便内分割区分 = '" & RTrim(EHT.便内分割区分) & "'"
        RTxtSql = RTxtSql &" And   行連番 = " & EHT.行連番 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           EHTRead = True
           Call ItemsFromEHT(DBNo)
	   EHT_RDSTS = True
        else
	   EHT_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call EHTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub EHTReadNext(Optional DBNo As Integer = 0)
    If Not EHT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           EHT_RDSTS = False
        Else
           Call ItemsFromEHT(DBNo)
           EHT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub EHTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        EHT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( EHT )
'*----------------------------------------------*'
Public Sub ItemsClearEHT()
    With EHT
	.処理日 = ""
	.処理回数 = 0
	.処理区分 = 0
	.ファイル名 = ""
	.納入先 = ""
	.受入 = ""
	.納入番号 = ""
	.納入日 = ""
	.納入便 = ""
	.品番 = ""
	.便内分割区分 = ""
	.退避連番 = 0
	.受注管理番号 = 0
	.受注管理行番号 = 0
	.出荷指示書NO = 0
	.出荷指示書行NO = 0
	.受注日 = ""
	.受注便 = ""
	.背番号 = ""
	.収容数 = 0
	.かんばん枚数 = 0
	.納入数 = 0
	.打切端数 = 0
	.打切区分 = 0
	.打切残数 = 0
	.出荷数 = 0
	.挽回日 = ""
	.挽回便 = ""
	.ERR重複 = 0
	.ERR収容数 = 0
	.ERR背番号 = 0
	.ERR受注 = 0
	.品番仮 = 0
	.品取仮 = 0
	.取引先仮 = 0
	.受入仮 = 0
	.納入先名 = ""
	.品名 = ""
	.行連番 = 0
'	.作成日 = 0
'	.作成端末 = ""
        .件数 = 0
'   項目追加(自動)
	.表示品番 = ""
	.納入先名称 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function EHTInsert(Optional DBNo As Integer) As Boolean
    EHTInsert = False
    With EHT
        gSL_Select = "Insert Into eかんばん変換テーブル  ( "
        gSL_Select = gSL_Select & "処理日,処理回数,処理区分,"
        gSL_Select = gSL_Select & "ファイル名,納入先,受入,"
        gSL_Select = gSL_Select & "納入番号,納入日,納入便,"
        gSL_Select = gSL_Select & "品番,便内分割区分,退避連番,"
        gSL_Select = gSL_Select & "受注管理番号,受注管理行番号,出荷指示書NO,"
        gSL_Select = gSL_Select & "出荷指示書行NO,受注日,受注便,"
        gSL_Select = gSL_Select & "背番号,収容数,かんばん枚数,"
        gSL_Select = gSL_Select & "納入数,打切端数,打切区分,"
        gSL_Select = gSL_Select & "打切残数,出荷数,挽回日,"
        gSL_Select = gSL_Select & "挽回便,ERR重複,ERR収容数,"
        gSL_Select = gSL_Select & "ERR背番号,ERR受注,品番仮,"
        gSL_Select = gSL_Select & "品取仮,取引先仮,受入仮,"
        gSL_Select = gSL_Select & "納入先名,品名,行連番,"
        gSL_Select = gSL_Select & "作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.処理日) & "',"		
        gSL_Select = gSL_Select & "" & .処理回数 & ","		
        gSL_Select = gSL_Select & "" & .処理区分 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.ファイル名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.便内分割区分) & "',"		
        gSL_Select = gSL_Select & "" & .退避連番 & ","		
        gSL_Select = gSL_Select & "" & .受注管理番号 & ","		
        gSL_Select = gSL_Select & "" & .受注管理行番号 & ","		
        gSL_Select = gSL_Select & "" & .出荷指示書NO & ","		
        gSL_Select = gSL_Select & "" & .出荷指示書行NO & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.受注日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受注便) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "" & .収容数 & ","		
        gSL_Select = gSL_Select & "" & .かんばん枚数 & ","		
        gSL_Select = gSL_Select & "" & .納入数 & ","		
        gSL_Select = gSL_Select & "" & .打切端数 & ","		
        gSL_Select = gSL_Select & "" & .打切区分 & ","		
        gSL_Select = gSL_Select & "" & .打切残数 & ","		
        gSL_Select = gSL_Select & "" & .出荷数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.挽回日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.挽回便) & "',"		
        gSL_Select = gSL_Select & "" & .ERR重複 & ","		
        gSL_Select = gSL_Select & "" & .ERR収容数 & ","		
        gSL_Select = gSL_Select & "" & .ERR背番号 & ","		
        gSL_Select = gSL_Select & "" & .ERR受注 & ","		
        gSL_Select = gSL_Select & "" & .品番仮 & ","		
        gSL_Select = gSL_Select & "" & .品取仮 & ","		
        gSL_Select = gSL_Select & "" & .取引先仮 & ","		
        gSL_Select = gSL_Select & "" & .受入仮 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.納入先名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品名) & "',"		
        gSL_Select = gSL_Select & "" & .行連番 & ","		
        gSL_Select = gSL_Select & "GETDATE(),"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    EHTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function EHTUpdate(Optional DBNo As Integer) As Boolean
    EHTUpdate = False	
    With EHT
        gSL_Select = "Update eかんばん変換テーブル  Set "
        gSL_Select = gSL_Select & "処理日 = '" & RTrim(.処理日) & "',"
        gSL_Select = gSL_Select & "処理回数 = " & .処理回数 & ","
        gSL_Select = gSL_Select & "処理区分 = " & .処理区分 & ","
        gSL_Select = gSL_Select & "ファイル名 = '" & RTrim(.ファイル名) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "便内分割区分 = '" & RTrim(.便内分割区分) & "',"
        gSL_Select = gSL_Select & "退避連番 = " & .退避連番 & ","
        gSL_Select = gSL_Select & "受注管理番号 = " & .受注管理番号 & ","
        gSL_Select = gSL_Select & "受注管理行番号 = " & .受注管理行番号 & ","
        gSL_Select = gSL_Select & "出荷指示書NO = " & .出荷指示書NO & ","
        gSL_Select = gSL_Select & "出荷指示書行NO = " & .出荷指示書行NO & ","
        gSL_Select = gSL_Select & "受注日 = '" & RTrim(.受注日) & "',"
        gSL_Select = gSL_Select & "受注便 = '" & RTrim(.受注便) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "収容数 = " & .収容数 & ","
        gSL_Select = gSL_Select & "かんばん枚数 = " & .かんばん枚数 & ","
        gSL_Select = gSL_Select & "納入数 = " & .納入数 & ","
        gSL_Select = gSL_Select & "打切端数 = " & .打切端数 & ","
        gSL_Select = gSL_Select & "打切区分 = " & .打切区分 & ","
        gSL_Select = gSL_Select & "打切残数 = " & .打切残数 & ","
        gSL_Select = gSL_Select & "出荷数 = " & .出荷数 & ","
        gSL_Select = gSL_Select & "挽回日 = '" & RTrim(.挽回日) & "',"
        gSL_Select = gSL_Select & "挽回便 = '" & RTrim(.挽回便) & "',"
        gSL_Select = gSL_Select & "ERR重複 = " & .ERR重複 & ","
        gSL_Select = gSL_Select & "ERR収容数 = " & .ERR収容数 & ","
        gSL_Select = gSL_Select & "ERR背番号 = " & .ERR背番号 & ","
        gSL_Select = gSL_Select & "ERR受注 = " & .ERR受注 & ","
        gSL_Select = gSL_Select & "品番仮 = " & .品番仮 & ","
        gSL_Select = gSL_Select & "品取仮 = " & .品取仮 & ","
        gSL_Select = gSL_Select & "取引先仮 = " & .取引先仮 & ","
        gSL_Select = gSL_Select & "受入仮 = " & .受入仮 & ","
        gSL_Select = gSL_Select & "納入先名 = '" & RTrim(.納入先名) & "',"
        gSL_Select = gSL_Select & "品名 = '" & RTrim(.品名) & "',"
        gSL_Select = gSL_Select & "行連番 = " & .行連番 & ""
        gSL_Select = gSL_Select & " Where 処理日 = '" & RTrim(.処理日) & "'"
        gSL_Select = gSL_Select & "   And 処理回数 = " & .処理回数 & ""
        gSL_Select = gSL_Select & "   And 処理区分 = " & .処理区分 & ""
        gSL_Select = gSL_Select & "   And ファイル名 = '" & RTrim(.ファイル名) & "'"
        gSL_Select = gSL_Select & "   And 納入先 = '" & RTrim(.納入先) & "'"
        gSL_Select = gSL_Select & "   And 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   And 納入番号 = '" & RTrim(.納入番号) & "'"
        gSL_Select = gSL_Select & "   And 納入日 = '" & RTrim(.納入日) & "'"
        gSL_Select = gSL_Select & "   And 納入便 = '" & RTrim(.納入便) & "'"
        gSL_Select = gSL_Select & "   And 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   And 便内分割区分 = '" & RTrim(.便内分割区分) & "'"
        gSL_Select = gSL_Select & "   And 行連番 = " & .行連番 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    EHTUpdate = True
End Function

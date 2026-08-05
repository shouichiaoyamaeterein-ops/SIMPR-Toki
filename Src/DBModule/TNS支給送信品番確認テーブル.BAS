Attribute VB_Name = "TNS支給送信品番確認テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TNS支給送信品番確認テーブル
'*
'*   モジュールID  : TNS支給送信品番確認テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2007/03/08  By Cis
'*   作成日  :  2007/03/09  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_TNS支給送信品番確認テーブル
        契約先            As String
        FILECD            As String
        FCD               As String
        処理日時          As String
        HEADKEY           As String
        区分              As String
        識別子            As String
        ブランク1         As String
        カード            As String
        仕入先            As String
        仕入先工区        As String
        管理番号          As String
        ブランク2         As String
        品番              As String
        単位              As String
        納入場所          As String
        納入区分          As String
        支給区分          As String
        部品区分          As String
        背番号            As String
        収容数            As String
        ブランク3         As String
        品名              As String
        担当者            As String
        提供依頼者        As String
        REC               As String
        件数          As Long
    End Type

    Public TSK As T_TNS支給送信品番確認テーブル

    Public TSK_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTSK(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTSK	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With TSK
            Select Case adoCol.Name
                Case "契約先":            If Not IsNull(adoCol.Value) Then .契約先           = adoCol.Value
                Case "FILECD":            If Not IsNull(adoCol.Value) Then .FILECD           = adoCol.Value
                Case "FCD":               If Not IsNull(adoCol.Value) Then .FCD              = adoCol.Value
                Case "処理日時":          If Not IsNull(adoCol.Value) Then .処理日時         = adoCol.Value
                Case "HEADKEY":           If Not IsNull(adoCol.Value) Then .HEADKEY          = adoCol.Value
                Case "区分":              If Not IsNull(adoCol.Value) Then .区分             = adoCol.Value
                Case "識別子":            If Not IsNull(adoCol.Value) Then .識別子           = adoCol.Value
                Case "ブランク1":         If Not IsNull(adoCol.Value) Then .ブランク1        = adoCol.Value
                Case "カード":            If Not IsNull(adoCol.Value) Then .カード           = adoCol.Value
                Case "仕入先":            If Not IsNull(adoCol.Value) Then .仕入先           = adoCol.Value
                Case "仕入先工区":        If Not IsNull(adoCol.Value) Then .仕入先工区       = adoCol.Value
                Case "管理番号":          If Not IsNull(adoCol.Value) Then .管理番号         = adoCol.Value
                Case "ブランク2":         If Not IsNull(adoCol.Value) Then .ブランク2        = adoCol.Value
                Case "品番":              If Not IsNull(adoCol.Value) Then .品番             = adoCol.Value
                Case "単位":              If Not IsNull(adoCol.Value) Then .単位             = adoCol.Value
                Case "納入場所":          If Not IsNull(adoCol.Value) Then .納入場所         = adoCol.Value
                Case "納入区分":          If Not IsNull(adoCol.Value) Then .納入区分         = adoCol.Value
                Case "支給区分":          If Not IsNull(adoCol.Value) Then .支給区分         = adoCol.Value
                Case "部品区分":          If Not IsNull(adoCol.Value) Then .部品区分         = adoCol.Value
                Case "背番号":            If Not IsNull(adoCol.Value) Then .背番号           = adoCol.Value
                Case "収容数":            If Not IsNull(adoCol.Value) Then .収容数           = adoCol.Value
                Case "ブランク3":         If Not IsNull(adoCol.Value) Then .ブランク3        = adoCol.Value
                Case "品名":              If Not IsNull(adoCol.Value) Then .品名             = adoCol.Value
                Case "担当者":            If Not IsNull(adoCol.Value) Then .担当者           = adoCol.Value
                Case "提供依頼者":        If Not IsNull(adoCol.Value) Then .提供依頼者       = adoCol.Value
                Case "REC":               If Not IsNull(adoCol.Value) Then .REC              = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TSKRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TSKRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From TNS支給送信品番確認テーブル"
        RTxtSql = RTxtSql &" Where 契約先 = '" & RTrim(TSK.契約先) & "'"
        RTxtSql = RTxtSql &" And   FILECD = '" & RTrim(TSK.FILECD) & "'"
        RTxtSql = RTxtSql &" And   処理日時 = '" & RTrim(TSK.処理日時) & "'"
        RTxtSql = RTxtSql &" And   HEADKEY = '" & RTrim(TSK.HEADKEY) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TSKRead = True
           Call ItemsFromTSK(DBNo)
	   TSK_RDSTS = True
        else
	   TSK_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call TSKClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TSKReadNext(Optional DBNo As Integer = 0)
    If Not TSK_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TSK_RDSTS = False
        Else
           Call ItemsFromTSK(DBNo)
           TSK_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TSKClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TSK_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TSK )
'*----------------------------------------------*'
Public Sub ItemsClearTSK()
    With TSK
	.契約先 = ""
	.FILECD = ""
	.FCD = ""
	.処理日時 = ""
	.HEADKEY = ""
	.区分 = ""
	.識別子 = ""
	.ブランク1 = ""
	.カード = ""
	.仕入先 = ""
	.仕入先工区 = ""
	.管理番号 = ""
	.ブランク2 = ""
	.品番 = ""
	.単位 = ""
	.納入場所 = ""
	.納入区分 = ""
	.支給区分 = ""
	.部品区分 = ""
	.背番号 = ""
	.収容数 = ""
	.ブランク3 = ""
	.品名 = ""
	.担当者 = ""
	.提供依頼者 = ""
	.REC = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function TSKInsert(Optional DBNo As Integer) As Boolean
    TSKInsert = True
    With TSK
        gSL_Select = "Insert Into TNS支給送信品番確認テーブル  ( "
        gSL_Select = gSL_Select & "契約先,FILECD,FCD,"
        gSL_Select = gSL_Select & "処理日時,HEADKEY,区分,"
        gSL_Select = gSL_Select & "識別子,ブランク1,カード,"
        gSL_Select = gSL_Select & "仕入先,仕入先工区,管理番号,"
        gSL_Select = gSL_Select & "ブランク2,品番,単位,"
        gSL_Select = gSL_Select & "納入場所,納入区分,支給区分,"
        gSL_Select = gSL_Select & "部品区分,背番号,収容数,"
        gSL_Select = gSL_Select & "ブランク3,品名,担当者,"
        gSL_Select = gSL_Select & "提供依頼者,REC"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.FILECD) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.FCD) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.処理日時) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.HEADKEY) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.識別子) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ブランク1) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.カード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先工区) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.管理番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ブランク2) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.単位) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入場所) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.部品区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.収容数) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.ブランク3) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品名) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.担当者) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.提供依頼者) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.REC) & "'"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TSKInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function TSKUpdate(Optional DBNo As Integer) As Boolean
    TSKUpdate = False	
    With TSK
        gSL_Select = "Update TNS支給送信品番確認テーブル  Set "
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "FILECD = '" & RTrim(.FILECD) & "',"
        gSL_Select = gSL_Select & "FCD = '" & RTrim(.FCD) & "',"
        gSL_Select = gSL_Select & "処理日時 = '" & RTrim(.処理日時) & "',"
        gSL_Select = gSL_Select & "HEADKEY = '" & RTrim(.HEADKEY) & "',"
        gSL_Select = gSL_Select & "区分 = '" & RTrim(.区分) & "',"
        gSL_Select = gSL_Select & "識別子 = '" & RTrim(.識別子) & "',"
        gSL_Select = gSL_Select & "ブランク1 = '" & RTrim(.ブランク1) & "',"
        gSL_Select = gSL_Select & "カード = '" & RTrim(.カード) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "仕入先工区 = '" & RTrim(.仕入先工区) & "',"
        gSL_Select = gSL_Select & "管理番号 = '" & RTrim(.管理番号) & "',"
        gSL_Select = gSL_Select & "ブランク2 = '" & RTrim(.ブランク2) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "単位 = '" & RTrim(.単位) & "',"
        gSL_Select = gSL_Select & "納入場所 = '" & RTrim(.納入場所) & "',"
        gSL_Select = gSL_Select & "納入区分 = '" & RTrim(.納入区分) & "',"
        gSL_Select = gSL_Select & "支給区分 = '" & RTrim(.支給区分) & "',"
        gSL_Select = gSL_Select & "部品区分 = '" & RTrim(.部品区分) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "収容数 = '" & RTrim(.収容数) & "',"
        gSL_Select = gSL_Select & "ブランク3 = '" & RTrim(.ブランク3) & "',"
        gSL_Select = gSL_Select & "品名 = '" & RTrim(.品名) & "',"
        gSL_Select = gSL_Select & "担当者 = '" & RTrim(.担当者) & "',"
        gSL_Select = gSL_Select & "提供依頼者 = '" & RTrim(.提供依頼者) & "',"
        gSL_Select = gSL_Select & "REC = '" & RTrim(.REC) & "'"
        gSL_Select = gSL_Select & " Where 契約先 = '" & RTrim(.契約先) & "'"
        gSL_Select = gSL_Select & "   And FILECD = '" & RTrim(.FILECD) & "'"
        gSL_Select = gSL_Select & "   And 処理日時 = '" & RTrim(.処理日時) & "'"
        gSL_Select = gSL_Select & "   And HEADKEY = '" & RTrim(.HEADKEY) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TSKUpdate = True
End Function

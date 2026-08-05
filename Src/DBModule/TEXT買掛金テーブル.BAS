Attribute VB_Name = "TEXT買掛金テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   TEXT買掛金テーブル
'*
'*   モジュールID  : TEXT買掛金テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					/項目追加含む/
'*   作成日  :  2004/11/15  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_TEXT買掛金テーブル
        財務管理番号  As Long
        FILECODE      As String
        HEAD作成日    As String
        HEADSEQ       As String
        FCD           As String
        契約先        As String
        対象年月      As String
        対象期        As Byte
        対象日        As String
        識別子        As String
        データ        As String
        カード        As String
        購買          As String
        貸借          As String
        工区          As String
        品番          As String
        納入番号      As String
        受入          As String
        依頼部署      As String
        単位          As String
        重量          As Currency
        繰越          As String
        品名          As String
        数量          As Currency
        単価          As Currency
        金額          As Long
        検収日        As String
        支給先        As String
        持込先        As String
        出庫者        As String
        支給区分      As String
        管理番号      As String
        型台帳NO      As String
        支払開始年月  As String
        設定月数      As Byte
        経過月数      As Byte
        型投資額      As Long
        当月支払額    As Currency
        科目コード    As String
        仕入先        As String
        データ元      As String
        補助品番      As String
        自動遡り区分  As String
        購入区分      As String
        仮号試        As String
        訂正理由      As String
        連番          As String
        決済番号      As String
        部品管理区分  As String
        仮単価        As String
        マイナス      As String
        照合年月      As String
        照合処理日    As String
        照合日付      As String
        照合品番      As String
        照合納番      As String
        照合管理      As String
        照合結果      As String
        照合区分      As String
'        作成端末      As String
'        作成日        As Date
'        更新端末      As String
'        更新日        As Date
        件数          As Long
'   項目追加(自動)
	契約先名      As String
	識別子名      As String
	帳票種類      As String
    End Type

    Public TKK As T_TEXT買掛金テーブル

    Public TKK_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromTKK(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearTKK	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With TKK
            Select Case adoCol.Name
                Case "財務管理番号":  If Not IsNull(adoCol.Value) Then .財務管理番号 = adoCol.Value
                Case "FILECODE":      If Not IsNull(adoCol.Value) Then .FILECODE     = adoCol.Value
                Case "HEAD作成日":    If Not IsNull(adoCol.Value) Then .HEAD作成日   = adoCol.Value
                Case "HEADSEQ":       If Not IsNull(adoCol.Value) Then .HEADSEQ      = adoCol.Value
                Case "FCD":           If Not IsNull(adoCol.Value) Then .FCD          = adoCol.Value
                Case "契約先":        If Not IsNull(adoCol.Value) Then .契約先       = adoCol.Value
                Case "対象年月":      If Not IsNull(adoCol.Value) Then .対象年月     = adoCol.Value
                Case "対象期":        If Not IsNull(adoCol.Value) Then .対象期       = adoCol.Value
                Case "対象日":        If Not IsNull(adoCol.Value) Then .対象日       = adoCol.Value
                Case "識別子":        If Not IsNull(adoCol.Value) Then .識別子       = adoCol.Value
                Case "データ":        If Not IsNull(adoCol.Value) Then .データ       = adoCol.Value
                Case "カード":        If Not IsNull(adoCol.Value) Then .カード       = adoCol.Value
                Case "購買":          If Not IsNull(adoCol.Value) Then .購買         = adoCol.Value
                Case "貸借":          If Not IsNull(adoCol.Value) Then .貸借         = adoCol.Value
                Case "工区":          If Not IsNull(adoCol.Value) Then .工区         = adoCol.Value
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番         = adoCol.Value
                Case "納入番号":      If Not IsNull(adoCol.Value) Then .納入番号     = adoCol.Value
                Case "受入":          If Not IsNull(adoCol.Value) Then .受入         = adoCol.Value
                Case "依頼部署":      If Not IsNull(adoCol.Value) Then .依頼部署     = adoCol.Value
                Case "単位":          If Not IsNull(adoCol.Value) Then .単位         = adoCol.Value
                Case "重量":          If Not IsNull(adoCol.Value) Then .重量         = adoCol.Value
                Case "繰越":          If Not IsNull(adoCol.Value) Then .繰越         = adoCol.Value
                Case "品名":          If Not IsNull(adoCol.Value) Then .品名         = adoCol.Value
                Case "数量":          If Not IsNull(adoCol.Value) Then .数量         = adoCol.Value
                Case "単価":          If Not IsNull(adoCol.Value) Then .単価         = adoCol.Value
                Case "金額":          If Not IsNull(adoCol.Value) Then .金額         = adoCol.Value
                Case "検収日":        If Not IsNull(adoCol.Value) Then .検収日       = adoCol.Value
                Case "支給先":        If Not IsNull(adoCol.Value) Then .支給先       = adoCol.Value
                Case "持込先":        If Not IsNull(adoCol.Value) Then .持込先       = adoCol.Value
                Case "出庫者":        If Not IsNull(adoCol.Value) Then .出庫者       = adoCol.Value
                Case "支給区分":      If Not IsNull(adoCol.Value) Then .支給区分     = adoCol.Value
                Case "管理番号":      If Not IsNull(adoCol.Value) Then .管理番号     = adoCol.Value
                Case "型台帳NO":      If Not IsNull(adoCol.Value) Then .型台帳NO     = adoCol.Value
                Case "支払開始年月":  If Not IsNull(adoCol.Value) Then .支払開始年月 = adoCol.Value
                Case "設定月数":      If Not IsNull(adoCol.Value) Then .設定月数     = adoCol.Value
                Case "経過月数":      If Not IsNull(adoCol.Value) Then .経過月数     = adoCol.Value
                Case "型投資額":      If Not IsNull(adoCol.Value) Then .型投資額     = adoCol.Value
                Case "当月支払額":    If Not IsNull(adoCol.Value) Then .当月支払額   = adoCol.Value
                Case "科目コード":    If Not IsNull(adoCol.Value) Then .科目コード   = adoCol.Value
                Case "仕入先":        If Not IsNull(adoCol.Value) Then .仕入先       = adoCol.Value
                Case "データ元":      If Not IsNull(adoCol.Value) Then .データ元     = adoCol.Value
                Case "補助品番":      If Not IsNull(adoCol.Value) Then .補助品番     = adoCol.Value
                Case "自動遡り区分":  If Not IsNull(adoCol.Value) Then .自動遡り区分 = adoCol.Value
                Case "購入区分":      If Not IsNull(adoCol.Value) Then .購入区分     = adoCol.Value
                Case "仮号試":        If Not IsNull(adoCol.Value) Then .仮号試       = adoCol.Value
                Case "訂正理由":      If Not IsNull(adoCol.Value) Then .訂正理由     = adoCol.Value
                Case "連番":          If Not IsNull(adoCol.Value) Then .連番         = adoCol.Value
                Case "決済番号":      If Not IsNull(adoCol.Value) Then .決済番号     = adoCol.Value
                Case "部品管理区分":  If Not IsNull(adoCol.Value) Then .部品管理区分 = adoCol.Value
                Case "仮単価":        If Not IsNull(adoCol.Value) Then .仮単価       = adoCol.Value
                Case "マイナス":      If Not IsNull(adoCol.Value) Then .マイナス     = adoCol.Value
                Case "照合年月":      If Not IsNull(adoCol.Value) Then .照合年月     = adoCol.Value
                Case "照合処理日":    If Not IsNull(adoCol.Value) Then .照合処理日   = adoCol.Value
                Case "照合日付":      If Not IsNull(adoCol.Value) Then .照合日付     = adoCol.Value
                Case "照合品番":      If Not IsNull(adoCol.Value) Then .照合品番     = adoCol.Value
                Case "照合納番":      If Not IsNull(adoCol.Value) Then .照合納番     = adoCol.Value
                Case "照合管理":      If Not IsNull(adoCol.Value) Then .照合管理     = adoCol.Value
                Case "照合結果":      If Not IsNull(adoCol.Value) Then .照合結果     = adoCol.Value
                Case "照合区分":      If Not IsNull(adoCol.Value) Then .照合区分     = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
'   項目追加(自動)
		Case "契約先名":	If Not IsNull(adoCol.Value) Then .契約先名	= adoCol.Value
		Case "識別子名":	If Not IsNull(adoCol.Value) Then .識別子名	= adoCol.Value
		Case "帳票種類":	If Not IsNull(adoCol.Value) Then .帳票種類	= adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function TKKRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    TKKRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From TEXT買掛金テーブル"
        RTxtSql = RTxtSql &" Where 財務管理番号 = " & TKK.財務管理番号 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           TKKRead = True
           Call ItemsFromTKK(DBNo)
	   TKK_RDSTS = True
        else
	   TKK_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call TKKClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub TKKReadNext(Optional DBNo As Integer = 0)
    If Not TKK_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           TKK_RDSTS = False
        Else
           Call ItemsFromTKK(DBNo)
           TKK_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub TKKClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        TKK_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( TKK )
'*----------------------------------------------*'
Public Sub ItemsClearTKK()
    With TKK
	.財務管理番号 = 0
	.FILECODE = ""
	.HEAD作成日 = ""
	.HEADSEQ = ""
	.FCD = ""
	.契約先 = ""
	.対象年月 = ""
	.対象期 = 0
	.対象日 = ""
	.識別子 = ""
	.データ = ""
	.カード = ""
	.購買 = ""
	.貸借 = ""
	.工区 = ""
	.品番 = ""
	.納入番号 = ""
	.受入 = ""
	.依頼部署 = ""
	.単位 = ""
	.重量 = 0
	.繰越 = ""
	.品名 = ""
	.数量 = 0
	.単価 = 0
	.金額 = 0
	.検収日 = ""
	.支給先 = ""
	.持込先 = ""
	.出庫者 = ""
	.支給区分 = ""
	.管理番号 = ""
	.型台帳NO = ""
	.支払開始年月 = ""
	.設定月数 = 0
	.経過月数 = 0
	.型投資額 = 0
	.当月支払額 = 0
	.科目コード = ""
	.仕入先 = ""
	.データ元 = ""
	.補助品番 = ""
	.自動遡り区分 = ""
	.購入区分 = ""
	.仮号試 = ""
	.訂正理由 = ""
	.連番 = ""
	.決済番号 = ""
	.部品管理区分 = ""
	.仮単価 = ""
	.マイナス = ""
	.照合年月 = ""
	.照合処理日 = ""
	.照合日付 = ""
	.照合品番 = ""
	.照合納番 = ""
	.照合管理 = ""
	.照合結果 = ""
	.照合区分 = ""
'	.作成端末 = ""
'	.作成日 = 0
'	.更新端末 = ""
'	.更新日 = 0
        .件数 = 0
'   項目追加(自動)
	.契約先名 = ""
	.識別子名 = ""
	.帳票種類 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function TKKInsert(Optional DBNo As Integer) As Boolean
    TKKInsert = True
    With TKK
        gSL_Select = "Insert Into TEXT買掛金テーブル  ( "
        gSL_Select = gSL_Select & "財務管理番号,FILECODE,HEAD作成日,"
        gSL_Select = gSL_Select & "HEADSEQ,FCD,契約先,"
        gSL_Select = gSL_Select & "対象年月,対象期,対象日,"
        gSL_Select = gSL_Select & "識別子,データ,カード,"
        gSL_Select = gSL_Select & "購買,貸借,工区,"
        gSL_Select = gSL_Select & "品番,納入番号,受入,"
        gSL_Select = gSL_Select & "依頼部署,単位,重量,"
        gSL_Select = gSL_Select & "繰越,品名,数量,"
        gSL_Select = gSL_Select & "単価,金額,検収日,"
        gSL_Select = gSL_Select & "支給先,持込先,出庫者,"
        gSL_Select = gSL_Select & "支給区分,管理番号,型台帳NO,"
        gSL_Select = gSL_Select & "支払開始年月,設定月数,経過月数,"
        gSL_Select = gSL_Select & "型投資額,当月支払額,科目コード,"
        gSL_Select = gSL_Select & "仕入先,データ元,補助品番,"
        gSL_Select = gSL_Select & "自動遡り区分,購入区分,仮号試,"
        gSL_Select = gSL_Select & "訂正理由,連番,決済番号,"
        gSL_Select = gSL_Select & "部品管理区分,仮単価,マイナス,"
        gSL_Select = gSL_Select & "照合年月,照合処理日,照合日付,"
        gSL_Select = gSL_Select & "照合品番,照合納番,照合管理,"
        gSL_Select = gSL_Select & "照合結果,照合区分,"
        gSL_Select = gSL_Select & "作成端末,作成日"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .財務管理番号 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.FILECODE) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.HEAD作成日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.HEADSEQ) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.FCD) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.対象年月) & "',"		
        gSL_Select = gSL_Select & "" & .対象期 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.対象日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.識別子) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.データ) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.カード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.購買) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.貸借) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.工区) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.依頼部署) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.単位) & "',"		
        gSL_Select = gSL_Select & "" & .重量 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.繰越) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.品名) & "',"		
        gSL_Select = gSL_Select & "" & .数量 & ","		
        gSL_Select = gSL_Select & "" & .単価 & ","		
        gSL_Select = gSL_Select & "" & .金額 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.検収日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.持込先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.出庫者) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支給区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.管理番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.型台帳NO) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.支払開始年月) & "',"		
        gSL_Select = gSL_Select & "" & .設定月数 & ","		
        gSL_Select = gSL_Select & "" & .経過月数 & ","		
        gSL_Select = gSL_Select & "" & .型投資額 & ","		
        gSL_Select = gSL_Select & "" & .当月支払額 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.科目コード) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.データ元) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.補助品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.自動遡り区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.購入区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仮号試) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.訂正理由) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.連番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.決済番号) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.部品管理区分) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.仮単価) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.マイナス) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合年月) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合処理日) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合日付) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合品番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合納番) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合管理) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合結果) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.照合区分) & "',"		
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20),"		
        gSL_Select = gSL_Select & "GETDATE()"		
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TKKInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function TKKUpdate(Optional DBNo As Integer) As Boolean
    TKKUpdate = False	
    With TKK
        gSL_Select = "Update TEXT買掛金テーブル  Set "
        gSL_Select = gSL_Select & "財務管理番号 = " & .財務管理番号 & ","
        gSL_Select = gSL_Select & "FILECODE = '" & RTrim(.FILECODE) & "',"
        gSL_Select = gSL_Select & "HEAD作成日 = '" & RTrim(.HEAD作成日) & "',"
        gSL_Select = gSL_Select & "HEADSEQ = '" & RTrim(.HEADSEQ) & "',"
        gSL_Select = gSL_Select & "FCD = '" & RTrim(.FCD) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "対象年月 = '" & RTrim(.対象年月) & "',"
        gSL_Select = gSL_Select & "対象期 = " & .対象期 & ","
        gSL_Select = gSL_Select & "対象日 = '" & RTrim(.対象日) & "',"
        gSL_Select = gSL_Select & "識別子 = '" & RTrim(.識別子) & "',"
        gSL_Select = gSL_Select & "データ = '" & RTrim(.データ) & "',"
        gSL_Select = gSL_Select & "カード = '" & RTrim(.カード) & "',"
        gSL_Select = gSL_Select & "購買 = '" & RTrim(.購買) & "',"
        gSL_Select = gSL_Select & "貸借 = '" & RTrim(.貸借) & "',"
        gSL_Select = gSL_Select & "工区 = '" & RTrim(.工区) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "依頼部署 = '" & RTrim(.依頼部署) & "',"
        gSL_Select = gSL_Select & "単位 = '" & RTrim(.単位) & "',"
        gSL_Select = gSL_Select & "重量 = " & .重量 & ","
        gSL_Select = gSL_Select & "繰越 = '" & RTrim(.繰越) & "',"
        gSL_Select = gSL_Select & "品名 = '" & RTrim(.品名) & "',"
        gSL_Select = gSL_Select & "数量 = " & .数量 & ","
        gSL_Select = gSL_Select & "単価 = " & .単価 & ","
        gSL_Select = gSL_Select & "金額 = " & .金額 & ","
        gSL_Select = gSL_Select & "検収日 = '" & RTrim(.検収日) & "',"
        gSL_Select = gSL_Select & "支給先 = '" & RTrim(.支給先) & "',"
        gSL_Select = gSL_Select & "持込先 = '" & RTrim(.持込先) & "',"
        gSL_Select = gSL_Select & "出庫者 = '" & RTrim(.出庫者) & "',"
        gSL_Select = gSL_Select & "支給区分 = '" & RTrim(.支給区分) & "',"
        gSL_Select = gSL_Select & "管理番号 = '" & RTrim(.管理番号) & "',"
        gSL_Select = gSL_Select & "型台帳NO = '" & RTrim(.型台帳NO) & "',"
        gSL_Select = gSL_Select & "支払開始年月 = '" & RTrim(.支払開始年月) & "',"
        gSL_Select = gSL_Select & "設定月数 = " & .設定月数 & ","
        gSL_Select = gSL_Select & "経過月数 = " & .経過月数 & ","
        gSL_Select = gSL_Select & "型投資額 = " & .型投資額 & ","
        gSL_Select = gSL_Select & "当月支払額 = " & .当月支払額 & ","
        gSL_Select = gSL_Select & "科目コード = '" & RTrim(.科目コード) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "データ元 = '" & RTrim(.データ元) & "',"
        gSL_Select = gSL_Select & "補助品番 = '" & RTrim(.補助品番) & "',"
        gSL_Select = gSL_Select & "自動遡り区分 = '" & RTrim(.自動遡り区分) & "',"
        gSL_Select = gSL_Select & "購入区分 = '" & RTrim(.購入区分) & "',"
        gSL_Select = gSL_Select & "仮号試 = '" & RTrim(.仮号試) & "',"
        gSL_Select = gSL_Select & "訂正理由 = '" & RTrim(.訂正理由) & "',"
        gSL_Select = gSL_Select & "連番 = '" & RTrim(.連番) & "',"
        gSL_Select = gSL_Select & "決済番号 = '" & RTrim(.決済番号) & "',"
        gSL_Select = gSL_Select & "部品管理区分 = '" & RTrim(.部品管理区分) & "',"
        gSL_Select = gSL_Select & "仮単価 = '" & RTrim(.仮単価) & "',"
        gSL_Select = gSL_Select & "マイナス = '" & RTrim(.マイナス) & "',"
        gSL_Select = gSL_Select & "照合年月 = '" & RTrim(.照合年月) & "',"
        gSL_Select = gSL_Select & "照合処理日 = '" & RTrim(.照合処理日) & "',"
        gSL_Select = gSL_Select & "照合日付 = '" & RTrim(.照合日付) & "',"
        gSL_Select = gSL_Select & "照合品番 = '" & RTrim(.照合品番) & "',"
        gSL_Select = gSL_Select & "照合納番 = '" & RTrim(.照合納番) & "',"
        gSL_Select = gSL_Select & "照合管理 = '" & RTrim(.照合管理) & "',"
        gSL_Select = gSL_Select & "照合結果 = '" & RTrim(.照合結果) & "',"
        gSL_Select = gSL_Select & "照合区分 = '" & RTrim(.照合区分) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "更新日 = GETDATE()"
        gSL_Select = gSL_Select & " Where 財務管理番号 = " & .財務管理番号 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    TKKUpdate = True
End Function

Attribute VB_Name = "購入明細テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   購入明細テーブル
'*
'*   モジュールID  : 購入明細テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*                                      /項目追加含む/
'*   作成日  :  2006/08/30  By Cis
'*   作成日  :  2008/01/21  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_購入明細テーブル
        検収管理NO            As Long
        実績年月              As String
        入力区分              As String
        識別子                As String
        契約先                As String
        品番                  As String
        仕入先                As String
        受入                  As String
        背番号                As String
        車型                  As String
        納入番号              As String
        検収日                As String
        赤黒区分              As Byte
        単価                  As Currency
        仮単価区分            As String
        数量                  As Long
        金額                  As Currency
        部品区分              As String
        単位                  As String
        持込先                As String
        購入区分              As String
        訂正理由              As String
        科目                  As String
        発注管理NO            As Long
        遡り区分              As String
        実績区分              As Byte
        実績処理日            As String
        計算方法              As Byte
        計算単位              As Byte
        締日                  As Byte
        明細件数              As Integer
        集約検収管理NO        As Long
        送信区分              As Byte
        送信処理日            As String
        照合年月              As String
        照合処理日            As String
        照合日付              As String
        照合品番              As String
        照合納番              As String
        照合管理              As String
        照合赤黒              As Byte
        照合結果              As String
        照合区分              As String
'        作成日                As Date
'        作成端末              As String
        作成者                As String
'        更新日                As Date
'        更新端末              As String
        件数          As Long
'   項目追加(自動)
        仕入先名       As String
    End Type

    Public KMT As T_購入明細テーブル

    Public KMT_RDSTS    As Boolean                 ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKMT(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKMT
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With KMT
            Select Case adoCol.Name
                Case "検収管理NO":            If Not IsNull(adoCol.Value) Then .検収管理NO = adoCol.Value
                Case "実績年月":              If Not IsNull(adoCol.Value) Then .実績年月 = adoCol.Value
                Case "入力区分":              If Not IsNull(adoCol.Value) Then .入力区分 = adoCol.Value
                Case "識別子":                If Not IsNull(adoCol.Value) Then .識別子 = adoCol.Value
                Case "契約先":                If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
                Case "品番":                  If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "仕入先":                If Not IsNull(adoCol.Value) Then .仕入先 = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "背番号":                If Not IsNull(adoCol.Value) Then .背番号 = adoCol.Value
                Case "車型":                  If Not IsNull(adoCol.Value) Then .車型 = adoCol.Value
                Case "納入番号":              If Not IsNull(adoCol.Value) Then .納入番号 = adoCol.Value
                Case "検収日":                If Not IsNull(adoCol.Value) Then .検収日 = adoCol.Value
                Case "赤黒区分":              If Not IsNull(adoCol.Value) Then .赤黒区分 = adoCol.Value
                Case "単価":                  If Not IsNull(adoCol.Value) Then .単価 = adoCol.Value
                Case "仮単価区分":            If Not IsNull(adoCol.Value) Then .仮単価区分 = adoCol.Value
                Case "数量":                  If Not IsNull(adoCol.Value) Then .数量 = adoCol.Value
                Case "金額":                  If Not IsNull(adoCol.Value) Then .金額 = adoCol.Value
                Case "部品区分":              If Not IsNull(adoCol.Value) Then .部品区分 = adoCol.Value
                Case "単位":                  If Not IsNull(adoCol.Value) Then .単位 = adoCol.Value
                Case "持込先":                If Not IsNull(adoCol.Value) Then .持込先 = adoCol.Value
                Case "購入区分":              If Not IsNull(adoCol.Value) Then .購入区分 = adoCol.Value
                Case "訂正理由":              If Not IsNull(adoCol.Value) Then .訂正理由 = adoCol.Value
                Case "科目":                  If Not IsNull(adoCol.Value) Then .科目 = adoCol.Value
                Case "発注管理NO":            If Not IsNull(adoCol.Value) Then .発注管理NO = adoCol.Value
                Case "遡り区分":              If Not IsNull(adoCol.Value) Then .遡り区分 = adoCol.Value
                Case "実績区分":              If Not IsNull(adoCol.Value) Then .実績区分 = adoCol.Value
                Case "実績処理日":            If Not IsNull(adoCol.Value) Then .実績処理日 = adoCol.Value
                Case "計算方法":              If Not IsNull(adoCol.Value) Then .計算方法 = adoCol.Value
                Case "計算単位":              If Not IsNull(adoCol.Value) Then .計算単位 = adoCol.Value
                Case "締日":                  If Not IsNull(adoCol.Value) Then .締日 = adoCol.Value
                Case "明細件数":              If Not IsNull(adoCol.Value) Then .明細件数 = adoCol.Value
                Case "集約検収管理NO":        If Not IsNull(adoCol.Value) Then .集約検収管理NO = adoCol.Value
                Case "送信区分":              If Not IsNull(adoCol.Value) Then .送信区分 = adoCol.Value
                Case "送信処理日":            If Not IsNull(adoCol.Value) Then .送信処理日 = adoCol.Value
                Case "照合年月":              If Not IsNull(adoCol.Value) Then .照合年月 = adoCol.Value
                Case "照合処理日":            If Not IsNull(adoCol.Value) Then .照合処理日 = adoCol.Value
                Case "照合日付":              If Not IsNull(adoCol.Value) Then .照合日付 = adoCol.Value
                Case "照合品番":              If Not IsNull(adoCol.Value) Then .照合品番 = adoCol.Value
                Case "照合納番":              If Not IsNull(adoCol.Value) Then .照合納番 = adoCol.Value
                Case "照合管理":              If Not IsNull(adoCol.Value) Then .照合管理 = adoCol.Value
                Case "照合赤黒":              If Not IsNull(adoCol.Value) Then .照合赤黒 = adoCol.Value
                Case "照合結果":              If Not IsNull(adoCol.Value) Then .照合結果 = adoCol.Value
                Case "照合区分":              If Not IsNull(adoCol.Value) Then .照合区分 = adoCol.Value
'                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日               = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "作成者":                If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "更新日":                If Not IsNull(adoCol.Value) Then .更新日               = adoCol.Value
'                Case "更新端末":              If Not IsNull(adoCol.Value) Then .更新端末             = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
'   項目追加(自動)
                Case "仕入先名":        If Not IsNull(adoCol.Value) Then .仕入先名 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KMTRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KMTRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 購入明細テーブル"
        RTxtSql = RTxtSql & " Where 検収管理NO = " & KMT.検収管理NO & ""
        RTxtSql = RTxtSql & " And   実績年月 = '" & RTrim(KMT.実績年月) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KMTRead = True
           Call ItemsFromKMT(DBNo)
           KMT_RDSTS = True
        Else
           KMT_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call KMTClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KMTReadNext(Optional DBNo As Integer = 0)
    If Not KMT_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KMT_RDSTS = False
        Else
           Call ItemsFromKMT(DBNo)
           KMT_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KMTClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KMT_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KMT )
'*----------------------------------------------*'
Public Sub ItemsClearKMT()
    With KMT
        .検収管理NO = 0
        .実績年月 = ""
        .入力区分 = ""
        .識別子 = ""
        .契約先 = ""
        .品番 = ""
        .仕入先 = ""
        .受入 = ""
        .背番号 = ""
        .車型 = ""
        .納入番号 = ""
        .検収日 = ""
        .赤黒区分 = 0
        .単価 = 0
        .仮単価区分 = ""
        .数量 = 0
        .金額 = 0
        .部品区分 = ""
        .単位 = ""
        .持込先 = ""
        .購入区分 = ""
        .訂正理由 = ""
        .科目 = ""
        .発注管理NO = 0
        .遡り区分 = ""
        .実績区分 = 0
        .実績処理日 = ""
        .計算方法 = 0
        .計算単位 = 0
        .締日 = 0
        .明細件数 = 0
        .集約検収管理NO = 0
        .送信区分 = 0
        .送信処理日 = ""
        .照合年月 = ""
        .照合処理日 = ""
        .照合日付 = ""
        .照合品番 = ""
        .照合納番 = ""
        .照合管理 = ""
        .照合赤黒 = 0
        .照合結果 = ""
        .照合区分 = ""
'       .作成日 = 0
'       .作成端末 = ""
        .作成者 = ""
'       .更新日 = 0
'       .更新端末 = ""
        .件数 = 0
'   項目追加(自動)
        .仕入先名 = ""
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KMTInsert(Optional DBNo As Integer) As Boolean
    KMTInsert = False
    With KMT
        gSL_Select = "Insert Into 購入明細テーブル  ( "
        gSL_Select = gSL_Select & "検収管理NO,実績年月,入力区分,"
        gSL_Select = gSL_Select & "識別子,契約先,品番,"
        gSL_Select = gSL_Select & "仕入先,受入,背番号,"
        gSL_Select = gSL_Select & "車型,納入番号,検収日,"
        gSL_Select = gSL_Select & "赤黒区分,単価,仮単価区分,"
        gSL_Select = gSL_Select & "数量,金額,部品区分,"
        gSL_Select = gSL_Select & "単位,持込先,購入区分,"
        gSL_Select = gSL_Select & "訂正理由,科目,発注管理NO,"
        gSL_Select = gSL_Select & "遡り区分,実績区分,実績処理日,"
        gSL_Select = gSL_Select & "計算方法,計算単位,締日,"
        gSL_Select = gSL_Select & "明細件数,集約検収管理NO,送信区分,"
        gSL_Select = gSL_Select & "送信処理日,照合年月,照合処理日,"
        gSL_Select = gSL_Select & "照合日付,照合品番,照合納番,"
        gSL_Select = gSL_Select & "照合管理,照合赤黒,照合結果,"
        gSL_Select = gSL_Select & "照合区分,"
        gSL_Select = gSL_Select & "作成日,作成端末,作成者"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .検収管理NO & ","
        gSL_Select = gSL_Select & "'" & RTrim(.実績年月) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.入力区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.識別子) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.車型) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.検収日) & "',"
        gSL_Select = gSL_Select & "" & .赤黒区分 & ","
        gSL_Select = gSL_Select & "" & .単価 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.仮単価区分) & "',"
        gSL_Select = gSL_Select & "" & .数量 & ","
        gSL_Select = gSL_Select & "" & .金額 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.部品区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.単位) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.持込先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.購入区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.訂正理由) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.科目) & "',"
        gSL_Select = gSL_Select & "" & .発注管理NO & ","
        gSL_Select = gSL_Select & "'" & RTrim(.遡り区分) & "',"
        gSL_Select = gSL_Select & "" & .実績区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.実績処理日) & "',"
        gSL_Select = gSL_Select & "" & .計算方法 & ","
        gSL_Select = gSL_Select & "" & .計算単位 & ","
        gSL_Select = gSL_Select & "" & .締日 & ","
        gSL_Select = gSL_Select & "" & .明細件数 & ","
        gSL_Select = gSL_Select & "" & .集約検収管理NO & ","
        gSL_Select = gSL_Select & "" & .送信区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.送信処理日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.照合年月) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.照合処理日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.照合日付) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.照合品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.照合納番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.照合管理) & "',"
        gSL_Select = gSL_Select & "" & .照合赤黒 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.照合結果) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.照合区分) & "',"
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20),"
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "'"
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KMTInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KMTUpdate(Optional DBNo As Integer) As Boolean
    KMTUpdate = False
    With KMT
        gSL_Select = "Update 購入明細テーブル  Set "
        gSL_Select = gSL_Select & "検収管理NO = " & .検収管理NO & ","
        gSL_Select = gSL_Select & "実績年月 = '" & RTrim(.実績年月) & "',"
        gSL_Select = gSL_Select & "入力区分 = '" & RTrim(.入力区分) & "',"
        gSL_Select = gSL_Select & "識別子 = '" & RTrim(.識別子) & "',"
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "車型 = '" & RTrim(.車型) & "',"
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "検収日 = '" & RTrim(.検収日) & "',"
        gSL_Select = gSL_Select & "赤黒区分 = " & .赤黒区分 & ","
        gSL_Select = gSL_Select & "単価 = " & .単価 & ","
        gSL_Select = gSL_Select & "仮単価区分 = '" & RTrim(.仮単価区分) & "',"
        gSL_Select = gSL_Select & "数量 = " & .数量 & ","
        gSL_Select = gSL_Select & "金額 = " & .金額 & ","
        gSL_Select = gSL_Select & "部品区分 = '" & RTrim(.部品区分) & "',"
        gSL_Select = gSL_Select & "単位 = '" & RTrim(.単位) & "',"
        gSL_Select = gSL_Select & "持込先 = '" & RTrim(.持込先) & "',"
        gSL_Select = gSL_Select & "購入区分 = '" & RTrim(.購入区分) & "',"
        gSL_Select = gSL_Select & "訂正理由 = '" & RTrim(.訂正理由) & "',"
        gSL_Select = gSL_Select & "科目 = '" & RTrim(.科目) & "',"
        gSL_Select = gSL_Select & "発注管理NO = " & .発注管理NO & ","
        gSL_Select = gSL_Select & "遡り区分 = '" & RTrim(.遡り区分) & "',"
        gSL_Select = gSL_Select & "実績区分 = " & .実績区分 & ","
        gSL_Select = gSL_Select & "実績処理日 = '" & RTrim(.実績処理日) & "',"
        gSL_Select = gSL_Select & "計算方法 = " & .計算方法 & ","
        gSL_Select = gSL_Select & "計算単位 = " & .計算単位 & ","
        gSL_Select = gSL_Select & "締日 = " & .締日 & ","
        gSL_Select = gSL_Select & "明細件数 = " & .明細件数 & ","
        gSL_Select = gSL_Select & "集約検収管理NO = " & .集約検収管理NO & ","
        gSL_Select = gSL_Select & "送信区分 = " & .送信区分 & ","
        gSL_Select = gSL_Select & "送信処理日 = '" & RTrim(.送信処理日) & "',"
        gSL_Select = gSL_Select & "照合年月 = '" & RTrim(.照合年月) & "',"
        gSL_Select = gSL_Select & "照合処理日 = '" & RTrim(.照合処理日) & "',"
        gSL_Select = gSL_Select & "照合日付 = '" & RTrim(.照合日付) & "',"
        gSL_Select = gSL_Select & "照合品番 = '" & RTrim(.照合品番) & "',"
        gSL_Select = gSL_Select & "照合納番 = '" & RTrim(.照合納番) & "',"
        gSL_Select = gSL_Select & "照合管理 = '" & RTrim(.照合管理) & "',"
        gSL_Select = gSL_Select & "照合赤黒 = " & .照合赤黒 & ","
        gSL_Select = gSL_Select & "照合結果 = '" & RTrim(.照合結果) & "',"
        gSL_Select = gSL_Select & "照合区分 = '" & RTrim(.照合区分) & "',"
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where 検収管理NO = " & .検収管理NO & ""
        gSL_Select = gSL_Select & "   And 実績年月 = '" & RTrim(.実績年月) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KMTUpdate = True
End Function

Attribute VB_Name = "品番マスタ"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   品番マスタ
'*
'*   モジュールID  : 品番マスタ.Bas
'*
'*   作成日  :  2004/02/12  By Cis
'*   変更日  :  2003/12/08  By Cis      品番編集ﾓｼﾞｭｰﾙ追加
'*   変更日  :  2003/12/17  By Cis      品番編集 (システム初期値取得)ﾓｼﾞｭｰﾙ追加
'*   変更日  :  2004/03/01  By Cis      項目追加:在庫管理区分名
'*   変更日  :  2008/02/04  By Cis      部品区分追加
'*   変更日  :  2008/06/12  By CIS      背景色コード、枠色コード追加
'*   変更日  :  2009/10/19  By CIS      得意先追加
'*---------------------------------------------------------------------------------
    Type T_品番マスタ
        品番          As String
        社内背番      As String
        車型          As String
        品名          As String
        品番編集      As String
        表示品番      As String
        基本収容器    As String
        基本収容数    As Long
        在庫管理区分  As Byte
        在庫管理数値  As Currency
        在庫係数      As Currency
        在庫日数      As Currency
        在庫発注点    As Long
        工程          As String
        次工程        As String
        担当          As String
        型NO          As String
        型所在地      As String
        IMAGE         As String
        実数区分      As Byte
        ABC           As String
        品番区分      As String
        仮区分        As Byte
'        作成日        As Date
        作成者        As String
'        作成端末      As String
'        更新日        As Date
        更新者        As String
'        更新端末      As String
        得意先       As String          '2009/10/19

' *ADD* 2008/02/04 S
        部品区分      As Byte
' *ADD* 2008/02/04 E

        件数          As Long
' *ADD* 2004/03/01
        在庫管理区分名  As String
' *ADD* 2008/06/12
        背景色コード  As Long
        枠色コード    As Long
        
        設定数      As Integer
        
        
    End Type

    Public HNM As T_品番マスタ

    Public HNM_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
    Public HNM_Edit     As String
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHNM(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHNM
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With HNM
            Select Case adoCol.Name
                Case "品番":          If Not IsNull(adoCol.Value) Then .品番 = adoCol.Value
                Case "社内背番":      If Not IsNull(adoCol.Value) Then .社内背番 = adoCol.Value
                Case "車型":          If Not IsNull(adoCol.Value) Then .車型 = adoCol.Value
                Case "品名":          If Not IsNull(adoCol.Value) Then .品名 = adoCol.Value
                Case "品番編集":      If Not IsNull(adoCol.Value) Then .品番編集 = adoCol.Value
                Case "表示品番":      If Not IsNull(adoCol.Value) Then .表示品番 = adoCol.Value
                Case "基本収容器":    If Not IsNull(adoCol.Value) Then .基本収容器 = adoCol.Value
                Case "基本収容数":    If Not IsNull(adoCol.Value) Then .基本収容数 = adoCol.Value
                Case "在庫管理区分":  If Not IsNull(adoCol.Value) Then .在庫管理区分 = adoCol.Value
                Case "在庫管理数値":  If Not IsNull(adoCol.Value) Then .在庫管理数値 = adoCol.Value
                Case "在庫係数":      If Not IsNull(adoCol.Value) Then .在庫係数 = adoCol.Value
                Case "在庫日数":      If Not IsNull(adoCol.Value) Then .在庫日数 = adoCol.Value
                Case "在庫発注点":    If Not IsNull(adoCol.Value) Then .在庫発注点 = adoCol.Value
                Case "工程":          If Not IsNull(adoCol.Value) Then .工程 = adoCol.Value
                Case "次工程":        If Not IsNull(adoCol.Value) Then .次工程 = adoCol.Value
                Case "担当":          If Not IsNull(adoCol.Value) Then .担当 = adoCol.Value
                Case "型NO":          If Not IsNull(adoCol.Value) Then .型NO = adoCol.Value
                Case "型所在地":      If Not IsNull(adoCol.Value) Then .型所在地 = adoCol.Value
                Case "IMAGE":         If Not IsNull(adoCol.Value) Then .IMAGE = adoCol.Value
                Case "実数区分":      If Not IsNull(adoCol.Value) Then .実数区分 = adoCol.Value
                Case "ABC":           If Not IsNull(adoCol.Value) Then .ABC = adoCol.Value
                Case "品番区分":      If Not IsNull(adoCol.Value) Then .品番区分 = adoCol.Value
                Case "仮区分":        If Not IsNull(adoCol.Value) Then .仮区分 = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者 = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者 = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
                Case "部品区分":      If Not IsNull(adoCol.Value) Then .部品区分 = adoCol.Value
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
                Case "得意先":       If Not IsNull(adoCol.Value) Then .得意先 = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
' *ADD* 2004/03/01
                Case "在庫管理区分名":         If Not IsNull(adoCol.Value) Then .在庫管理区分名 = adoCol.Value
' *ADD* 2008/06/12
                Case "背景色コード":       If Not IsNull(adoCol.Value) Then .背景色コード = adoCol.Value
                Case "枠色コード":         If Not IsNull(adoCol.Value) Then .枠色コード = adoCol.Value
                Case "設定数":             If Not IsNull(adoCol.Value) Then .設定数 = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function HNMRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HNMRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 品番マスタ"
        RTxtSql = RTxtSql & " Where 品番 = '" & RTrim(HNM.品番) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HNMRead = True
           Call ItemsFromHNM(DBNo)
           HNM_RDSTS = True
        Else
           HNM_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call HNMClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HNMReadNext(Optional DBNo As Integer = 0)
    If Not HNM_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HNM_RDSTS = False
        Else
           Call ItemsFromHNM(DBNo)
           HNM_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HNMClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HNM_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HNM )
'*----------------------------------------------*'
Public Sub ItemsClearHNM()
    With HNM
        .品番 = ""
        .社内背番 = ""
        .車型 = ""
        .品名 = ""
        .品番編集 = ""
        .表示品番 = ""
        .基本収容器 = ""
        .基本収容数 = 0
        .在庫管理区分 = 0
        .在庫管理数値 = 0
        .在庫係数 = 0
        .在庫日数 = 0
        .在庫発注点 = 0
        .工程 = ""
        .次工程 = ""
        .担当 = ""
        .型NO = ""
        .型所在地 = ""
        .IMAGE = ""
        .実数区分 = 0
        .ABC = ""
        .品番区分 = ""
        .得意先 = ""
        .仮区分 = 0
'       .作成日 = 0
        .作成者 = ""
'       .作成端末 = ""
'       .更新日 = 0
        .更新者 = ""
'       .更新端末 = ""
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
        .部品区分 = 0
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
        .件数 = 0
    
' *ADD* 2004/03/01
        .在庫管理区分名 = ""
' *ADD* 2008/06/12
        .背景色コード = 0
        .枠色コード = 0
    
        .設定数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function HNMInsert(Optional DBNo As Integer) As Boolean
    HNMInsert = True
    With HNM
        gSL_Select = "Insert Into 品番マスタ  ( "
        gSL_Select = gSL_Select & "品番,社内背番,車型,"
        gSL_Select = gSL_Select & "品名,品番編集,表示品番,"
        gSL_Select = gSL_Select & "基本収容器,基本収容数,在庫管理区分,"
        gSL_Select = gSL_Select & "在庫管理数値,在庫係数,在庫日数,"
        gSL_Select = gSL_Select & "在庫発注点,工程,次工程,"
        gSL_Select = gSL_Select & "担当,型NO,型所在地,"
        gSL_Select = gSL_Select & "IMAGE,実数区分,ABC,"
        gSL_Select = gSL_Select & "品番区分,得意先,仮区分,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
        gSL_Select = gSL_Select & ",部品区分,設定数"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "'" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.社内背番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.車型) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品名) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品番編集) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.表示品番) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.基本収容器) & "',"
        gSL_Select = gSL_Select & "" & .基本収容数 & ","
        gSL_Select = gSL_Select & "" & .在庫管理区分 & ","
        gSL_Select = gSL_Select & "" & .在庫管理数値 & ","
        gSL_Select = gSL_Select & "" & .在庫係数 & ","
        gSL_Select = gSL_Select & "" & .在庫日数 & ","
        gSL_Select = gSL_Select & "" & .在庫発注点 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.工程) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.次工程) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.担当) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.型NO) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.型所在地) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.IMAGE) & "',"
        gSL_Select = gSL_Select & "" & .実数区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.ABC) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.品番区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.得意先) & "',"
        gSL_Select = gSL_Select & "" & .仮区分 & ","
        gSL_Select = gSL_Select & "GETDATE(),"
        gSL_Select = gSL_Select & "'" & RTrim(.作成者) & "',"
        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
        gSL_Select = gSL_Select & "," & .部品区分
        gSL_Select = gSL_Select & "," & .設定数
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HNMInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function HNMUpdate(Optional DBNo As Integer) As Boolean
    HNMUpdate = False
    With HNM
        gSL_Select = "Update 品番マスタ  Set "
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "社内背番 = '" & RTrim(.社内背番) & "',"
        gSL_Select = gSL_Select & "車型 = '" & RTrim(.車型) & "',"
        gSL_Select = gSL_Select & "品名 = '" & RTrim(.品名) & "',"
        gSL_Select = gSL_Select & "品番編集 = '" & RTrim(.品番編集) & "',"
        gSL_Select = gSL_Select & "表示品番 = '" & RTrim(.表示品番) & "',"
        gSL_Select = gSL_Select & "基本収容器 = '" & RTrim(.基本収容器) & "',"
        gSL_Select = gSL_Select & "基本収容数 = " & .基本収容数 & ","
        gSL_Select = gSL_Select & "在庫管理区分 = " & .在庫管理区分 & ","
        gSL_Select = gSL_Select & "在庫管理数値 = " & .在庫管理数値 & ","
        gSL_Select = gSL_Select & "在庫係数 = " & .在庫係数 & ","
        gSL_Select = gSL_Select & "在庫日数 = " & .在庫日数 & ","
        gSL_Select = gSL_Select & "在庫発注点 = " & .在庫発注点 & ","
        gSL_Select = gSL_Select & "工程 = '" & RTrim(.工程) & "',"
        gSL_Select = gSL_Select & "次工程 = '" & RTrim(.次工程) & "',"
        gSL_Select = gSL_Select & "担当 = '" & RTrim(.担当) & "',"
        gSL_Select = gSL_Select & "型NO = '" & RTrim(.型NO) & "',"
        gSL_Select = gSL_Select & "型所在地 = '" & RTrim(.型所在地) & "',"
        gSL_Select = gSL_Select & "IMAGE = '" & RTrim(.IMAGE) & "',"
        gSL_Select = gSL_Select & "実数区分 = " & .実数区分 & ","
        gSL_Select = gSL_Select & "ABC = '" & RTrim(.ABC) & "',"
        gSL_Select = gSL_Select & "品番区分 = '" & RTrim(.品番区分) & "',"
        gSL_Select = gSL_Select & "得意先 = '" & RTrim(.得意先) & "',"
        gSL_Select = gSL_Select & "仮区分 = " & .仮区分 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT START
        gSL_Select = gSL_Select & ",部品区分 = " & .部品区分
        gSL_Select = gSL_Select & ",設定数 = " & .設定数
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/02/04 INSERT END
        gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(.品番) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HNMUpdate = True
End Function
'+--------------------------+
'+  品番編集(ｼｽﾃﾑ初期値取得)
'+--------------------------+
Public Sub EditHinbnInit()
' 品番編集初期値　取得
    HNM_Edit = "552"
    gSL_Select = " SELECT * FROM 名称マスタ "
    gSL_Select = gSL_Select & " WHERE 区分名称 = '品番編集'"
    gSL_Select = gSL_Select & "   AND 区分タイプ = '' "
    gSL_Select = gSL_Select & "   AND 値 = '1' "
    If SYMRead(gSL_Select, 1, 1) Then
       HNM_Edit = SYM.英数字1
    End If

End Sub
'+--------------------------+
'+  品番編集
'+--------------------------+
Public Function EditHinbn(Hinbn As String, Edit As String) As String
    Dim eId         As Long
    Dim eSt         As Long
    Dim eEd         As Long
    Dim eStr        As String
    Dim eEdit       As String
    Dim eLeng       As String
    Dim eEnd        As Long
    
    EditHinbn = Hinbn
    If Trim(Edit) = "" Then Exit Function
'       Edit = HNM_Edit
'    End If
    If Trim(Edit) = "" Then Exit Function
    
    If Len(Edit) < 5 Then
       Edit = Edit & String(5 - Len(Edit), "0")
    End If
    eEnd = 0
    For eId = 1 To 5
        eLeng = Mid(Edit, eId, 1)
        If IsNumeric(eLeng) Then
            eEd = CLng(eLeng)
            If eEd <> 0 Then
               eEnd = eId
            End If
        End If
    Next eId
    
    If eEnd = 0 Then Exit Function
    
    eSt = 1: eEd = 0: eStr = "": eEdit = ""
    For eId = 1 To eEnd
        eLeng = Mid(Edit, eId, 1)
        If IsNumeric(eLeng) Then
            eEd = CLng(eLeng)
            If eEd <> 0 Then
               If eId = eEnd Then
                  eStr = eStr & eEdit & Mid(Hinbn, eSt)
               Else
                  eStr = eStr & eEdit & Mid(Hinbn, eSt, eEd)
               End If
               eEdit = "-"
            End If
            eSt = eSt + eEd
            If Len(Hinbn) < eSt Then Exit For
        End If
    Next eId
    
    EditHinbn = RTrim(eStr)
    
End Function


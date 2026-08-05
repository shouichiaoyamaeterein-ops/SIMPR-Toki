Attribute VB_Name = "かんばん設定詳細テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   かんばん設定詳細テーブル
'*
'*   モジュールID  : かんばん設定詳細テーブル.Bas
'*
'*   ( 当モジュールは自動生成されたモジュールです。)
'*					
'*   作成日  :  2004/04/16  By Cis
'*   変更日  :
'*---------------------------------------------------------------------------------
    Type T_かんばん設定詳細テーブル
        かんばん種類  As Byte
        設定区分      As Byte
        設定連番      As Long
        表示          As String
        始点X         As Long
        始点Y         As Long
        終点X         As Long
        終点Y         As Long
        種別          As Long
        幅            As Long
        丸み          As Long
        横倍          As Byte
        縦倍          As Byte
        フォント      As String
        文字間        As Integer
        回転          As Byte
        背景          As String
        フィールド    As String
        固有設定      As String
        固定          As String
        桁数          As Long
        少数          As Byte
        編集          As String
        位置          As Byte
        文字幅        As Long
        文字高        As Long
'        作成日        As Date
        作成者        As String
'        作成端末      As String
'        更新日        As Date
        更新者        As String
'        更新端末      As String
        件数          As Long
    End Type

    Public KSS As T_かんばん設定詳細テーブル

    Public KSS_RDSTS    As Boolean     	    ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromKSS(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearKSS	
    
    CisDB.ConnectNo = DBNo
	
    For Each adoCol In CisDB.Recordset.Fields
        With KSS
            Select Case adoCol.Name
                Case "かんばん種類":  If Not IsNull(adoCol.Value) Then .かんばん種類 = adoCol.Value
                Case "設定区分":      If Not IsNull(adoCol.Value) Then .設定区分     = adoCol.Value
                Case "設定連番":      If Not IsNull(adoCol.Value) Then .設定連番     = adoCol.Value
                Case "表示":          If Not IsNull(adoCol.Value) Then .表示         = adoCol.Value
                Case "始点X":         If Not IsNull(adoCol.Value) Then .始点X        = adoCol.Value
                Case "始点Y":         If Not IsNull(adoCol.Value) Then .始点Y        = adoCol.Value
                Case "終点X":         If Not IsNull(adoCol.Value) Then .終点X        = adoCol.Value
                Case "終点Y":         If Not IsNull(adoCol.Value) Then .終点Y        = adoCol.Value
                Case "種別":          If Not IsNull(adoCol.Value) Then .種別         = adoCol.Value
                Case "幅":            If Not IsNull(adoCol.Value) Then .幅           = adoCol.Value
                Case "丸み":          If Not IsNull(adoCol.Value) Then .丸み         = adoCol.Value
                Case "横倍":          If Not IsNull(adoCol.Value) Then .横倍         = adoCol.Value
                Case "縦倍":          If Not IsNull(adoCol.Value) Then .縦倍         = adoCol.Value
                Case "フォント":      If Not IsNull(adoCol.Value) Then .フォント     = adoCol.Value
                Case "文字間":        If Not IsNull(adoCol.Value) Then .文字間       = adoCol.Value
                Case "回転":          If Not IsNull(adoCol.Value) Then .回転         = adoCol.Value
                Case "背景":          If Not IsNull(adoCol.Value) Then .背景         = adoCol.Value
                Case "フィールド":    If Not IsNull(adoCol.Value) Then .フィールド   = adoCol.Value
                Case "固有設定":      If Not IsNull(adoCol.Value) Then .固有設定     = adoCol.Value
                Case "固定":          If Not IsNull(adoCol.Value) Then .固定         = adoCol.Value
                Case "桁数":          If Not IsNull(adoCol.Value) Then .桁数         = adoCol.Value
                Case "少数":          If Not IsNull(adoCol.Value) Then .少数         = adoCol.Value
                Case "編集":          If Not IsNull(adoCol.Value) Then .編集         = adoCol.Value
                Case "位置":          If Not IsNull(adoCol.Value) Then .位置         = adoCol.Value
                Case "文字幅":        If Not IsNull(adoCol.Value) Then .文字幅       = adoCol.Value
                Case "文字高":        If Not IsNull(adoCol.Value) Then .文字高       = adoCol.Value
'                Case "作成日":        If Not IsNull(adoCol.Value) Then .作成日       = adoCol.Value
                Case "作成者":        If Not IsNull(adoCol.Value) Then .作成者       = adoCol.Value
'                Case "作成端末":      If Not IsNull(adoCol.Value) Then .作成端末     = adoCol.Value
'                Case "更新日":        If Not IsNull(adoCol.Value) Then .更新日       = adoCol.Value
                Case "更新者":        If Not IsNull(adoCol.Value) Then .更新者       = adoCol.Value
'                Case "更新端末":      If Not IsNull(adoCol.Value) Then .更新端末     = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数        = adoCol.Value
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function KSSRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    KSSRead = False
    If RTxtSql = "" Then 
	RTxtSql = "Select * From かんばん設定詳細テーブル"
        RTxtSql = RTxtSql &" Where かんばん種類 = " & KSS.かんばん種類 & ""
        RTxtSql = RTxtSql &" And   設定区分 = " & KSS.設定区分 & ""
        RTxtSql = RTxtSql &" And   設定連番 = " & KSS.設定連番 & ""
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           KSSRead = True
           Call ItemsFromKSS(DBNo)
	   KSS_RDSTS = True
        else
	   KSS_RDSTS = False
        End If
    End With		

    If RGyo = 1 Then
       Call KSSClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub KSSReadNext(Optional DBNo As Integer = 0)
    If Not KSS_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           KSS_RDSTS = False
        Else
           Call ItemsFromKSS(DBNo)
           KSS_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub KSSClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        KSS_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( KSS )
'*----------------------------------------------*'
Public Sub ItemsClearKSS()
    With KSS
	.かんばん種類 = 0
	.設定区分 = 0
	.設定連番 = 0
	.表示 = ""
	.始点X = 0
	.始点Y = 0
	.終点X = 0
	.終点Y = 0
	.種別 = 0
	.幅 = 0
	.丸み = 0
	.横倍 = 0
	.縦倍 = 0
	.フォント = ""
	.文字間 = 0
	.回転 = 0
	.背景 = ""
	.フィールド = ""
	.固有設定 = ""
	.固定 = ""
	.桁数 = 0
	.少数 = 0
	.編集 = ""
	.位置 = 0
	.文字幅 = 0
	.文字高 = 0
'	.作成日 = 0
	.作成者 = ""
'	.作成端末 = ""
'	.更新日 = 0
	.更新者 = ""
'	.更新端末 = ""
        .件数 = 0
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function KSSInsert(Optional DBNo As Integer) As Boolean
    KSSInsert = True
    With KSS
        gSL_Select = "Insert Into かんばん設定詳細テーブル  ( "
        gSL_Select = gSL_Select & "かんばん種類,設定区分,設定連番,"
        gSL_Select = gSL_Select & "表示,始点X,始点Y,"
        gSL_Select = gSL_Select & "終点X,終点Y,種別,"
        gSL_Select = gSL_Select & "幅,丸み,横倍,"
        gSL_Select = gSL_Select & "縦倍,フォント,文字間,"
        gSL_Select = gSL_Select & "回転,背景,フィールド,"
        gSL_Select = gSL_Select & "固有設定,固定,桁数,"
        gSL_Select = gSL_Select & "少数,編集,位置,"
        gSL_Select = gSL_Select & "文字幅,文字高,"
        gSL_Select = gSL_Select & "作成日,作成者,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .かんばん種類 & ","		
        gSL_Select = gSL_Select & "" & .設定区分 & ","		
        gSL_Select = gSL_Select & "" & .設定連番 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.表示) & "',"		
        gSL_Select = gSL_Select & "" & .始点X & ","		
        gSL_Select = gSL_Select & "" & .始点Y & ","		
        gSL_Select = gSL_Select & "" & .終点X & ","		
        gSL_Select = gSL_Select & "" & .終点Y & ","		
        gSL_Select = gSL_Select & "" & .種別 & ","		
        gSL_Select = gSL_Select & "" & .幅 & ","		
        gSL_Select = gSL_Select & "" & .丸み & ","		
        gSL_Select = gSL_Select & "" & .横倍 & ","		
        gSL_Select = gSL_Select & "" & .縦倍 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.フォント) & "',"		
        gSL_Select = gSL_Select & "" & .文字間 & ","		
        gSL_Select = gSL_Select & "" & .回転 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.背景) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.フィールド) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.固有設定) & "',"		
        gSL_Select = gSL_Select & "'" & RTrim(.固定) & "',"		
        gSL_Select = gSL_Select & "" & .桁数 & ","		
        gSL_Select = gSL_Select & "" & .少数 & ","		
        gSL_Select = gSL_Select & "'" & RTrim(.編集) & "',"		
        gSL_Select = gSL_Select & "" & .位置 & ","		
        gSL_Select = gSL_Select & "" & .文字幅 & ","		
        gSL_Select = gSL_Select & "" & .文字高 & ","		
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
    KSSInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function KSSUpdate(Optional DBNo As Integer) As Boolean
    KSSUpdate = False	
    With KSS
        gSL_Select = "Update かんばん設定詳細テーブル  Set "
        gSL_Select = gSL_Select & "かんばん種類 = " & .かんばん種類 & ","
        gSL_Select = gSL_Select & "設定区分 = " & .設定区分 & ","
        gSL_Select = gSL_Select & "設定連番 = " & .設定連番 & ","
        gSL_Select = gSL_Select & "表示 = '" & RTrim(.表示) & "',"
        gSL_Select = gSL_Select & "始点X = " & .始点X & ","
        gSL_Select = gSL_Select & "始点Y = " & .始点Y & ","
        gSL_Select = gSL_Select & "終点X = " & .終点X & ","
        gSL_Select = gSL_Select & "終点Y = " & .終点Y & ","
        gSL_Select = gSL_Select & "種別 = " & .種別 & ","
        gSL_Select = gSL_Select & "幅 = " & .幅 & ","
        gSL_Select = gSL_Select & "丸み = " & .丸み & ","
        gSL_Select = gSL_Select & "横倍 = " & .横倍 & ","
        gSL_Select = gSL_Select & "縦倍 = " & .縦倍 & ","
        gSL_Select = gSL_Select & "フォント = '" & RTrim(.フォント) & "',"
        gSL_Select = gSL_Select & "文字間 = " & .文字間 & ","
        gSL_Select = gSL_Select & "回転 = " & .回転 & ","
        gSL_Select = gSL_Select & "背景 = '" & RTrim(.背景) & "',"
        gSL_Select = gSL_Select & "フィールド = '" & RTrim(.フィールド) & "',"
        gSL_Select = gSL_Select & "固有設定 = '" & RTrim(.固有設定) & "',"
        gSL_Select = gSL_Select & "固定 = '" & RTrim(.固定) & "',"
        gSL_Select = gSL_Select & "桁数 = " & .桁数 & ","
        gSL_Select = gSL_Select & "少数 = " & .少数 & ","
        gSL_Select = gSL_Select & "編集 = '" & RTrim(.編集) & "',"
        gSL_Select = gSL_Select & "位置 = " & .位置 & ","
        gSL_Select = gSL_Select & "文字幅 = " & .文字幅 & ","
        gSL_Select = gSL_Select & "文字高 = " & .文字高 & ","
        gSL_Select = gSL_Select & "更新日 = GETDATE(),"
        gSL_Select = gSL_Select & "更新者 = '" & RTrim(.更新者) & "',"
        gSL_Select = gSL_Select & "更新端末 = SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & " Where かんばん種類 = " & .かんばん種類 & ""
        gSL_Select = gSL_Select & "   And 設定区分 = " & .設定区分 & ""
        gSL_Select = gSL_Select & "   And 設定連番 = " & .設定連番 & ""
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    KSSUpdate = True
End Function

Attribute VB_Name = "発注帳票発行テーブル"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   ActiveX Data Objects (ADO<DLL>) 仕様
'*
'*   発注帳票発行テーブル
'*
'*   モジュールID  : 発注帳票発行テーブル.Bas
'*
'*
'*
'*   作成日  :  2004/02/07  By Cis
'*   変更日  :  2004/02/07  項目追加
'*   変更日  :  2004/09/16  項目追加（支給元・支給元2＆代行・代行工場・代行受入・代行名称＆後工程符号）
'*   変更日  :  2005/02/21  項目追加（発注日・発注便・打切区分1～10）
'*   変更日  :  2005/02/22  項目追加（分納区分・担当者・分納時初回行）
'*   変更日  :  2007/12/11  項目追加（発行区分:0=発行無し、1:発行有り）
'*   変更日  :  2007/12/18  項目追加（代行契約先）
'*   変更日  :  2008/11/14  項目追加（自社受入名）
'*---------------------------------------------------------------------------------
    Type T_発注帳票発行テーブル
        納品書管理NO          As Long
        納入番号              As String
        頁NO                  As Byte
        カード                As String
        発行順                As Byte
        契約先                As String
        仕入先                As String
        受入                  As String
        納入先                As String
        納入先受入            As String
        発注区分              As Byte
        納入日                As String
        納入便                As String
        手配区分              As String
        SK区分                As Byte
        代行契約先            As String             '2007/12/18 ADD
        箱総数                As Long
        総頁数                As Byte
        作成日                As String
        作成時間              As String
        グループNO            As String
        物理現頁数            As Byte
        物理総頁数            As Byte
        END表示区分           As String
        ALLEND表示区分        As String
        明細行数              As Byte
        行(10)                As Byte
        アドレス(10)          As String
        背番号(10)            As String
        部品番号(10)          As String
        表示品番(10)          As String
        部品名称(10)          As String
        収容数(10)            As Long
        箱数(10)              As Long
        端数(10)              As Long
        納入数(10)            As Long
        処理端末              As String
        処理区分              As Byte
        TEXTNAME              As String
        発行区分              As Byte
        発行種類              As String
        発行形態              As String
        発行日                As String
'        作成端末              As String
        件数          As Long
    
        発注区分名            As String
        手配区分名            As String
        仕入先名              As String
        納入先名              As String
        仕入先CD              As String
        仕入先工場            As String
        仕入先工場名          As String
        納入先CD              As String
        納入先工場            As String
        納入先工場名          As String
        
        契約取引先            As String
        契約工場              As String
        被自社                As String
        被自社工場            As String
        
        自社工場              As String
        機番                  As String
        発行元                As String
'*--------------------------------------------------<< 2004/09/16 >> Insert Start
        支給元                As String
        支給元2               As String
        契約納入先            As String
        契約納入先工場        As String
        契約納入先受入        As String
        契約納入先名称        As String
        直直区分              As Byte
        後工程符号            As String
'*--------------------------------------------------<< 2004/09/16 >> Insert End
'*--------------------------------------------------<< 2005/02/21 >> Insert Start
        発注日                As String
        発注便                As String
        打切区分(10)          As String
'*--------------------------------------------------<< 2005/02/21 >> Insert End
'*--------------------------------------------------<< 2005/02/22 >> Insert Start
        分納区分              As String
        担当者                As String
        担当者名              As String
        分納時初回行          As Byte
'*--------------------------------------------------<< 2005/02/22 >> Insert End
    
'================= *** 2008/11/14 Start
        自社受入名            As String
'================= *** 2008/11/14 End
    
    End Type

    Public HCH As T_発注帳票発行テーブル

    Public HCH_RDSTS    As Boolean          ' 専用読取フラグ
    Dim adoCol        As Object
'*----------------------------------------*'
'*      結果ｾｯﾄより構造体へ値をｾｯﾄ         *
'*----------------------------------------*'
Public Sub ItemsFromHCH(Optional DBNo As Integer)
'   退避エリアクリア
    Call ItemsClearHCH
    
    CisDB.ConnectNo = DBNo
        
    For Each adoCol In CisDB.Recordset.Fields
        With HCH
            Select Case adoCol.Name
                Case "納品書管理NO":          If Not IsNull(adoCol.Value) Then .納品書管理NO = adoCol.Value
                Case "納入番号":              If Not IsNull(adoCol.Value) Then .納入番号 = adoCol.Value
                Case "頁NO":                  If Not IsNull(adoCol.Value) Then .頁NO = adoCol.Value
                Case "カード":                If Not IsNull(adoCol.Value) Then .カード = adoCol.Value
                Case "発行順":                If Not IsNull(adoCol.Value) Then .発行順 = adoCol.Value
                Case "契約先":                If Not IsNull(adoCol.Value) Then .契約先 = adoCol.Value
                Case "仕入先":                If Not IsNull(adoCol.Value) Then .仕入先 = adoCol.Value
                Case "受入":                  If Not IsNull(adoCol.Value) Then .受入 = adoCol.Value
                Case "納入先":                If Not IsNull(adoCol.Value) Then .納入先 = adoCol.Value
                Case "納入先受入":            If Not IsNull(adoCol.Value) Then .納入先受入 = adoCol.Value
                Case "発注区分":              If Not IsNull(adoCol.Value) Then .発注区分 = adoCol.Value
                Case "納入日":                If Not IsNull(adoCol.Value) Then .納入日 = adoCol.Value
                Case "納入便":                If Not IsNull(adoCol.Value) Then .納入便 = adoCol.Value
                Case "手配区分":              If Not IsNull(adoCol.Value) Then .手配区分 = adoCol.Value
                Case "SK区分":                If Not IsNull(adoCol.Value) Then .SK区分 = adoCol.Value
                Case "代行契約先":            If Not IsNull(adoCol.Value) Then .代行契約先 = adoCol.Value
                Case "箱総数":                If Not IsNull(adoCol.Value) Then .箱総数 = adoCol.Value
                Case "総頁数":                If Not IsNull(adoCol.Value) Then .総頁数 = adoCol.Value
                Case "作成日":                If Not IsNull(adoCol.Value) Then .作成日 = adoCol.Value
                Case "作成時間":              If Not IsNull(adoCol.Value) Then .作成時間 = adoCol.Value
                Case "グループNO":            If Not IsNull(adoCol.Value) Then .グループNO = adoCol.Value
                Case "物理現頁数":            If Not IsNull(adoCol.Value) Then .物理現頁数 = adoCol.Value
                Case "物理総頁数":            If Not IsNull(adoCol.Value) Then .物理総頁数 = adoCol.Value
                Case "END表示区分":           If Not IsNull(adoCol.Value) Then .END表示区分 = adoCol.Value
                Case "ALLEND表示区分":        If Not IsNull(adoCol.Value) Then .ALLEND表示区分 = adoCol.Value
                Case "明細行数":              If Not IsNull(adoCol.Value) Then .明細行数 = adoCol.Value
                Case "処理端末":              If Not IsNull(adoCol.Value) Then .処理端末 = adoCol.Value
                Case "処理区分":              If Not IsNull(adoCol.Value) Then .処理区分 = adoCol.Value
                Case "TEXTNAME":              If Not IsNull(adoCol.Value) Then .TEXTNAME = adoCol.Value
                Case "発行区分":              If Not IsNull(adoCol.Value) Then .発行区分 = adoCol.Value
                Case "発行種類":              If Not IsNull(adoCol.Value) Then .発行種類 = adoCol.Value
                Case "発行形態":              If Not IsNull(adoCol.Value) Then .発行形態 = adoCol.Value
                Case "発行日":                If Not IsNull(adoCol.Value) Then .発行日 = adoCol.Value
'                Case "作成端末":              If Not IsNull(adoCol.Value) Then .作成端末             = adoCol.Value
                Case "件数":         If Not IsNull(adoCol.Value) Then .件数 = adoCol.Value
                
                Case "発注区分名":              If Not IsNull(adoCol.Value) Then .発注区分名 = adoCol.Value
                Case "手配区分名":              If Not IsNull(adoCol.Value) Then .手配区分名 = adoCol.Value
                Case "仕入先名":                If Not IsNull(adoCol.Value) Then .仕入先名 = adoCol.Value
                Case "納入先名":                If Not IsNull(adoCol.Value) Then .納入先名 = adoCol.Value
                Case "仕入先CD":                  If Not IsNull(adoCol.Value) Then .仕入先CD = adoCol.Value
                Case "仕入先工場":              If Not IsNull(adoCol.Value) Then .仕入先工場 = adoCol.Value
                Case "仕入先工場名":            If Not IsNull(adoCol.Value) Then .仕入先工場名 = adoCol.Value
                Case "納入先CD":                  If Not IsNull(adoCol.Value) Then .納入先CD = adoCol.Value
                Case "納入先工場":              If Not IsNull(adoCol.Value) Then .納入先工場 = adoCol.Value
                Case "納入先工場名":            If Not IsNull(adoCol.Value) Then .納入先工場名 = adoCol.Value
                
                Case "契約取引先":              If Not IsNull(adoCol.Value) Then .契約取引先 = adoCol.Value
                Case "契約工場":                If Not IsNull(adoCol.Value) Then .契約工場 = adoCol.Value
                Case "被自社":                  If Not IsNull(adoCol.Value) Then .被自社 = adoCol.Value
                Case "被自社工場":              If Not IsNull(adoCol.Value) Then .被自社工場 = adoCol.Value
'*--------------------------------------------------<< 2004/09/16 >> Insert Start
                Case "支給元":                If Not IsNull(adoCol.Value) Then .支給元 = adoCol.Value
                Case "支給元2":               If Not IsNull(adoCol.Value) Then .支給元2 = adoCol.Value
                Case "契約納入先":            If Not IsNull(adoCol.Value) Then .契約納入先 = adoCol.Value
                Case "契約納入先工場":        If Not IsNull(adoCol.Value) Then .契約納入先工場 = adoCol.Value
                Case "契約納入先受入":        If Not IsNull(adoCol.Value) Then .契約納入先受入 = adoCol.Value
                Case "契約納入先名称":        If Not IsNull(adoCol.Value) Then .契約納入先名称 = adoCol.Value
                Case "直直区分":              If Not IsNull(adoCol.Value) Then .直直区分 = adoCol.Value
                Case "後工程符号":            If Not IsNull(adoCol.Value) Then .後工程符号 = adoCol.Value
'*--------------------------------------------------<< 2004/09/16 >> Insert End
'*--------------------------------------------------<< 2005/02/21 >> Insert Start
                Case "発注日":                If Not IsNull(adoCol.Value) Then .発注日 = adoCol.Value
                Case "発注便":                If Not IsNull(adoCol.Value) Then .発注便 = adoCol.Value
'*--------------------------------------------------<< 2005/02/21 >> Insert End
'*--------------------------------------------------<< 2005/02/22 >> Insert Start
                Case "分納区分":              If Not IsNull(adoCol.Value) Then .分納区分 = adoCol.Value
                Case "担当者":                If Not IsNull(adoCol.Value) Then .担当者 = adoCol.Value
                Case "担当者名":              If Not IsNull(adoCol.Value) Then .担当者名 = adoCol.Value
                Case "分納時初回行":          If Not IsNull(adoCol.Value) Then .分納時初回行 = adoCol.Value
'========================= *** 2008/11/14 Start
                Case "自社受入名":            If Not IsNull(adoCol.Value) Then .自社受入名 = adoCol.Value
'========================= *** 2008/11/14 Start
'*--------------------------------------------------<< 2005/02/22 >> Insert End
                Case Else:
                    If Not IsNull(adoCol.Value) Then
                        If adoCol.Name Like "行*" Then .行(Val(Mid(adoCol.Name, 2, 2))) = adoCol.Value
                        If adoCol.Name Like "アドレス*" Then .アドレス(Val(Mid(adoCol.Name, 5, 2))) = adoCol.Value
                        If adoCol.Name Like "背番号*" Then .背番号(Val(Mid(adoCol.Name, 4, 2))) = adoCol.Value
                        If adoCol.Name Like "部品番号*" Then .部品番号(Val(Mid(adoCol.Name, 5, 2))) = adoCol.Value
                        If adoCol.Name Like "表示品番*" Then .表示品番(Val(Mid(adoCol.Name, 5, 2))) = adoCol.Value
                        If adoCol.Name Like "部品名称*" Then .部品名称(Val(Mid(adoCol.Name, 5, 2))) = adoCol.Value
                        If adoCol.Name Like "収容数*" Then .収容数(Val(Mid(adoCol.Name, 4, 2))) = adoCol.Value
                        If adoCol.Name Like "箱数*" Then .箱数(Val(Mid(adoCol.Name, 3, 2))) = adoCol.Value
                        If adoCol.Name Like "端数*" Then .端数(Val(Mid(adoCol.Name, 3, 2))) = adoCol.Value
                        If adoCol.Name Like "納入数*" Then .納入数(Val(Mid(adoCol.Name, 4, 2))) = adoCol.Value
                    
'*--------------------------------------------------<< 2005/02/21 >> Insert Start
                        If adoCol.Name Like "打切区分*" Then .打切区分(Val(Mid(adoCol.Name, 5, 2))) = adoCol.Value
'*--------------------------------------------------<< 2005/02/21 >> Insert End
                    End If
            End Select
        End With
    Next adoCol
End Sub
'*----------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ    *
'*----------------------------------------*'
Public Function HCHRead(RTxtSql As String, Optional RGyo As Long, Optional DBNo As Integer)
    HCHRead = False
    If RTxtSql = "" Then
        RTxtSql = "Select * From 発注帳票発行テーブル"
        RTxtSql = RTxtSql & " Where 納入番号 = '" & RTrim(HCH.納入番号) & "'"
    End If
    With CisDB
        .SQL = RTxtSql
        If .DBRead(RGyo, DBNo) Then
           HCHRead = True
           Call ItemsFromHCH(DBNo)
           HCH_RDSTS = True
        Else
           HCH_RDSTS = False
        End If
    End With

    If RGyo = 1 Then
       Call HCHClose(DBNo)
    End If
End Function
'*----------------------------------------------*'
'*      選択ｸｴﾘｰ(SELECT)実行＆構造体ｾｯﾄ(Next)    *
'*----------------------------------------------*'
Public Sub HCHReadNext(Optional DBNo As Integer = 0)
    If Not HCH_RDSTS Then Exit Sub
    With CisDB
        .ConnectNo = DBNo
        .Recordset.MoveNext
        If .Recordset.EOF Then
           HCH_RDSTS = False
        Else
           Call ItemsFromHCH(DBNo)
           HCH_RDSTS = True
        End If
    End With
End Sub
'*-------------------------*'
'*     専用結果ｾｯﾄｸﾛｰｽﾞ     *
'*-------------------------*'
Public Sub HCHClose(Optional DBNo As Integer = 0)
    On Error Resume Next
    With CisDB
        .ConnectNo = DBNo
        .Recordset.Close
        HCH_RDSTS = False
    End With
    On Error GoTo 0
End Sub
'*----------------------------------------------*'
'*      ワークエリアクリア  ( HCH )
'*----------------------------------------------*'
Public Sub ItemsClearHCH()
    With HCH
        .納品書管理NO = 0
        .納入番号 = ""
        .頁NO = 0
        .カード = ""
        .発行順 = 0
        .契約先 = ""
        .仕入先 = ""
        .受入 = ""
        .納入先 = ""
        .納入先受入 = ""
        .発注区分 = 0
        .納入日 = ""
        .納入便 = ""
        .手配区分 = ""
        .SK区分 = 0
        .代行契約先 = ""
        .箱総数 = 0
        .総頁数 = 0
        .作成日 = ""
        .作成時間 = ""
        .グループNO = ""
        .物理現頁数 = 0
        .物理総頁数 = 0
        .END表示区分 = ""
        .ALLEND表示区分 = ""
        .明細行数 = 0
        For gInt = 1 To 10
            .行(gInt) = 0
            .アドレス(gInt) = ""
            .背番号(gInt) = ""
            .部品番号(gInt) = ""
            .表示品番(gInt) = ""
            .部品名称(gInt) = ""
            .収容数(gInt) = 0
            .箱数(gInt) = 0
            .端数(gInt) = 0
            .納入数(gInt) = 0
            .打切区分(gInt) = 0
        Next
        .処理端末 = ""
        .処理区分 = 0
        .TEXTNAME = ""
        .発行区分 = 0
        .発行種類 = ""
        .発行形態 = ""
        .発行日 = ""
'       .作成端末 = ""
        .発注日 = ""
        .発注便 = ""
        .件数 = 0
    
        .発注区分名 = ""
        .手配区分名 = ""
        .仕入先名 = ""
        .納入先名 = ""
        .仕入先CD = ""
        .仕入先工場 = ""
        .仕入先工場名 = ""
        .納入先 = ""
        .納入先工場 = ""
        .納入先工場名 = ""
    
        .契約取引先 = ""
        .契約工場 = ""
        .被自社 = ""
        .被自社工場 = ""
'*--------------------------------------------------<< 2004/09/16 >> Insert Start
        .支給元 = ""
        .支給元2 = ""
        .契約納入先 = ""
        .契約納入先工場 = ""
        .契約納入先受入 = ""
        .契約納入先名称 = ""
        .直直区分 = 0
        .後工程符号 = ""
'*--------------------------------------------------<< 2004/09/16 >> Insert End
'*--------------------------------------------------<< 2005/02/22 >> Insert Start
        .分納区分 = ""
        .担当者 = ""
        .担当者名 = ""
        .分納時初回行 = 0
'*--------------------------------------------------<< 2005/02/22 >> Insert End
    End With
End Sub
'*--------------------*'
'*      追加処理       *
'*--------------------*'
Public Function HCHInsert(Optional DBNo As Integer) As Boolean
    HCHInsert = True
    With HCH
        gSL_Select = "Insert Into 発注帳票発行テーブル  ( "
        gSL_Select = gSL_Select & "納品書管理NO,納入番号,頁NO,"
        gSL_Select = gSL_Select & "カード,発行順,契約先,"
        gSL_Select = gSL_Select & "仕入先,受入,納入先,"
        gSL_Select = gSL_Select & "納入先受入,発注区分,納入日,"
        gSL_Select = gSL_Select & "納入便,手配区分,SK区分,"
        gSL_Select = gSL_Select & "代行契約先,"
        gSL_Select = gSL_Select & "箱総数,総頁数,作成日,作成時間,"
        gSL_Select = gSL_Select & "グループNO,物理現頁数,物理総頁数,"
        gSL_Select = gSL_Select & "END表示区分,ALLEND表示区分,明細行数,"
        For gInt = 1 To 10
            gSL_Select = gSL_Select & "行" & gInt & ","
            gSL_Select = gSL_Select & "アドレス" & gInt & ","
            gSL_Select = gSL_Select & "背番号" & gInt & ","
            gSL_Select = gSL_Select & "部品番号" & gInt & ","
            gSL_Select = gSL_Select & "表示品番" & gInt & ","
            gSL_Select = gSL_Select & "部品名称" & gInt & ","
            gSL_Select = gSL_Select & "収容数" & gInt & ","
            gSL_Select = gSL_Select & "箱数" & gInt & ","
            gSL_Select = gSL_Select & "端数" & gInt & ","
            gSL_Select = gSL_Select & "納入数" & gInt & ","
            gSL_Select = gSL_Select & "打切区分" & gInt & ","
        Next gInt
        gSL_Select = gSL_Select & "処理端末,処理区分,"
        gSL_Select = gSL_Select & "TEXTNAME,発行区分,発行種類,"
        gSL_Select = gSL_Select & "発行形態,発行日,発注日,発注便,担当者,分納時初回行"
'        gSL_Select = gSL_Select & "作成日,作成端末"
        gSL_Select = gSL_Select & " ) Values ( "
        gSL_Select = gSL_Select & "" & .納品書管理NO & ","
        gSL_Select = gSL_Select & "'" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "" & .頁NO & ","
        gSL_Select = gSL_Select & "'" & RTrim(.カード) & "',"
        gSL_Select = gSL_Select & "" & .発行順 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入先受入) & "',"
        gSL_Select = gSL_Select & "" & .発注区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "" & .SK区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.代行契約先) & "',"
        gSL_Select = gSL_Select & "" & .箱総数 & ","
        gSL_Select = gSL_Select & "" & .総頁数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.作成日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.作成時間) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.グループNO) & "',"
        gSL_Select = gSL_Select & "" & .物理現頁数 & ","
        gSL_Select = gSL_Select & "" & .物理総頁数 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.END表示区分) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.ALLEND表示区分) & "',"
        gSL_Select = gSL_Select & "" & .明細行数 & ","
        For gInt = 1 To 10
            gSL_Select = gSL_Select & "" & .行(gInt) & ","
            gSL_Select = gSL_Select & "'" & RTrim(.アドレス(gInt)) & "',"
            gSL_Select = gSL_Select & "'" & RTrim(.背番号(gInt)) & "',"
            gSL_Select = gSL_Select & "'" & RTrim(.部品番号(gInt)) & "',"
            gSL_Select = gSL_Select & "'" & RTrim(.表示品番(gInt)) & "',"
            gSL_Select = gSL_Select & "'" & RTrim(.部品名称(gInt)) & "',"
            gSL_Select = gSL_Select & "" & .収容数(gInt) & ","
            gSL_Select = gSL_Select & "" & .箱数(gInt) & ","
            gSL_Select = gSL_Select & "" & .端数(gInt) & ","
            gSL_Select = gSL_Select & "" & .納入数(gInt) & ","
            gSL_Select = gSL_Select & "" & .打切区分(gInt) & ","
        Next
        gSL_Select = gSL_Select & "'" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "" & .処理区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.TEXTNAME) & "',"
        gSL_Select = gSL_Select & "" & .発行区分 & ","
        gSL_Select = gSL_Select & "'" & RTrim(.発行種類) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発行形態) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発行日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発注日) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.発注便) & "',"
        gSL_Select = gSL_Select & "'" & RTrim(.担当者) & "',"
'        gSL_Select = gSL_Select & "GETDATE(),"
'        gSL_Select = gSL_Select & "SubString(Host_Name(),1,20)"
        gSL_Select = gSL_Select & "" & .分納時初回行 & ""
        gSL_Select = gSL_Select & " )"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HCHInsert = True
End Function
'*--------------------*'
'*      更新処理       *
'*--------------------*'
Public Function HCHUpdate(Optional DBNo As Integer) As Boolean
    HCHUpdate = False
    With HCH
        gSL_Select = "Update 発注帳票発行テーブル  Set "
        gSL_Select = gSL_Select & "納品書管理NO = " & .納品書管理NO & ","
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "頁NO = " & .頁NO & ","
        gSL_Select = gSL_Select & "カード = '" & RTrim(.カード) & "',"
        gSL_Select = gSL_Select & "発行順 = " & .発行順 & ","
        gSL_Select = gSL_Select & "契約先 = '" & RTrim(.契約先) & "',"
        gSL_Select = gSL_Select & "仕入先 = '" & RTrim(.仕入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "納入先受入 = '" & RTrim(.納入先受入) & "',"
        gSL_Select = gSL_Select & "発注区分 = " & .発注区分 & ","
        gSL_Select = gSL_Select & "納入日 = '" & RTrim(.納入日) & "',"
        gSL_Select = gSL_Select & "納入便 = '" & RTrim(.納入便) & "',"
        gSL_Select = gSL_Select & "手配区分 = '" & RTrim(.手配区分) & "',"
        gSL_Select = gSL_Select & "SK区分 = " & .SK区分 & ","
        gSL_Select = gSL_Select & "代行契約先 = '" & RTrim(.代行契約先) & "',"
        gSL_Select = gSL_Select & "箱総数 = " & .箱総数 & ","
        gSL_Select = gSL_Select & "総頁数 = " & .総頁数 & ","
        gSL_Select = gSL_Select & "作成日 = '" & RTrim(.作成日) & "',"
        gSL_Select = gSL_Select & "作成時間 = '" & RTrim(.作成時間) & "',"
        gSL_Select = gSL_Select & "グループNO = '" & RTrim(.グループNO) & "',"
        gSL_Select = gSL_Select & "物理現頁数 = " & .物理現頁数 & ","
        gSL_Select = gSL_Select & "物理総頁数 = " & .物理総頁数 & ","
        gSL_Select = gSL_Select & "END表示区分 = '" & RTrim(.END表示区分) & "',"
        gSL_Select = gSL_Select & "ALLEND表示区分 = '" & RTrim(.ALLEND表示区分) & "',"
        gSL_Select = gSL_Select & "明細行数 = " & .明細行数 & ","
        For gInt = 1 To 10
            gSL_Select = gSL_Select & "行" & gInt & " = " & .行(gInt) & ","
            gSL_Select = gSL_Select & "アドレス" & gInt & " = '" & RTrim(.アドレス(gInt)) & "',"
            gSL_Select = gSL_Select & "背番号" & gInt & " = '" & RTrim(.背番号(gInt)) & "',"
            gSL_Select = gSL_Select & "部品番号" & gInt & " = '" & RTrim(.部品番号(gInt)) & "',"
            gSL_Select = gSL_Select & "表示品番" & gInt & " = '" & RTrim(.表示品番(gInt)) & "',"
            gSL_Select = gSL_Select & "部品名称" & gInt & " = '" & RTrim(.部品名称(gInt)) & "',"
            gSL_Select = gSL_Select & "収容数" & gInt & " = " & .収容数(gInt) & ","
            gSL_Select = gSL_Select & "箱数" & gInt & " = " & .箱数(gInt) & ","
            gSL_Select = gSL_Select & "端数" & gInt & " = " & .端数(gInt) & ","
            gSL_Select = gSL_Select & "納入数" & gInt & " = " & .納入数(gInt) & ","
            gSL_Select = gSL_Select & "打切区分" & gInt & " = " & .打切区分(gInt) & ","
        Next
        gSL_Select = gSL_Select & "処理端末 = '" & RTrim(.処理端末) & "',"
        gSL_Select = gSL_Select & "処理区分 = " & .処理区分 & ","
        gSL_Select = gSL_Select & "TEXTNAME = '" & RTrim(.TEXTNAME) & "',"
        gSL_Select = gSL_Select & "発行区分 = " & .発行区分 & ","
        gSL_Select = gSL_Select & "発行種類 = '" & RTrim(.発行種類) & "',"
        gSL_Select = gSL_Select & "発行形態 = '" & RTrim(.発行形態) & "',"
        gSL_Select = gSL_Select & "発行日 = '" & RTrim(.発行日) & "',"
        gSL_Select = gSL_Select & "発注日 = '" & RTrim(.発注日) & "',"
        gSL_Select = gSL_Select & "発注便 = '" & RTrim(.発注便) & "'"
        gSL_Select = gSL_Select & "担当者 = '" & RTrim(.担当者) & "'"
        gSL_Select = gSL_Select & "分納時初回行 = " & .分納時初回行 & ","
        gSL_Select = gSL_Select & " Where 納入番号 = '" & RTrim(.納入番号) & "'"
    End With
    
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    HCHUpdate = True
End Function

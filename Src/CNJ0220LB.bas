Attribute VB_Name = "CNJ0220LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   内示表　印刷
'*
'*   モジュールID  : CNJ0220.Bas
'*
'*   戻値 : なし
'*
'*   作成日  :2003/12/17   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrt           As New CISFormPrint

Public gNJYM            As String
Public wKadouFlg        As Boolean
Public wMM(4)           As String
Public wKad(4)          As String
Public wKadsu(4)        As Single

'*********************************************************
'*      稼動日数 退避
'*********************************************************
Public Sub GetKadou()
    
    If wKadouFlg = True Then Exit Sub
    
' 前月
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add gNJYM, M, -1
        gStr = .Date_Result
    End With
    gSL_Select = "SELECT * FROM カレンダマスタ "
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先 = '' "
    gSL_Select = gSL_Select & "   AND 年 = '" & Mid(gStr, 1, 4) & "'"
    gSL_Select = gSL_Select & "   AND 月 = '" & Mid(gStr, 5, 2) & "'"
    If CLMRead(gSL_Select, 1, 1) Then
        wMM(0) = Format(CLM.月, "00") & "月稼働日"
        wKad(0) = Format(CLM.稼動日数, "00.0") & " 日"
        wKadsu(0) = CLM.稼動日数
    End If
' 当月
    gSL_Select = "SELECT * FROM カレンダマスタ "
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先 = '' "
    gSL_Select = gSL_Select & "   AND 年 = '" & Mid(gNJYM, 1, 4) & "'"
    gSL_Select = gSL_Select & "   AND 月 = '" & Mid(gNJYM, 5, 2) & "'"
    If CLMRead(gSL_Select, 1, 1) Then
        wMM(1) = Format(CLM.月, "00") & "月稼働日"
        wKad(1) = Format(CLM.稼動日数, "00.0") & " 日"
        wKadsu(1) = CLM.稼動日数
    End If
' 翌月
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add gNJYM, M, 1
        gStr = .Date_Result
    End With
    gSL_Select = "SELECT * FROM カレンダマスタ "
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先 = '' "
    gSL_Select = gSL_Select & "   AND 年 = '" & Mid(gStr, 1, 4) & "'"
    gSL_Select = gSL_Select & "   AND 月 = '" & Mid(gStr, 5, 2) & "'"
    If CLMRead(gSL_Select, 1, 1) Then
        wMM(2) = Format(CLM.月, "00") & "月稼働日"
        wKad(2) = Format(CLM.稼動日数, "00.0") & " 日"
        wKadsu(2) = CLM.稼動日数
    End If
' 翌々月
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add gNJYM, M, 2
        gStr = .Date_Result
    End With
    gSL_Select = "SELECT * FROM カレンダマスタ "
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先 = '' "
    gSL_Select = gSL_Select & "   AND 年 = '" & Mid(gStr, 1, 4) & "'"
    gSL_Select = gSL_Select & "   AND 月 = '" & Mid(gStr, 5, 2) & "'"
    If CLMRead(gSL_Select, 1, 1) Then
        wMM(3) = Format(CLM.月, "00") & "月稼働日"
        wKad(3) = Format(CLM.稼動日数, "00.0") & " 日"
        wKadsu(3) = CLM.稼動日数
    End If

    wKadouFlg = False
End Sub


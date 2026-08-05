Attribute VB_Name = "NJModule"
Option Explicit
'*-----------------------------------------------------------------------------------*
'*
'*      内　示　処　理　系　共　通　処　理　モ　ジ　ュ　ー　ル
'*
'*          作  成  日    :   2004/01/19  By CIS
'*          変  更  日    :
'*
'*-----------------------------------------------------------------------------------*
    Global gn_NJYM              As String       ' 　　内示対象年月
    Global gn_SKYM              As String       ' 生産計画対象年月
    Global gn_HNYM              As String       ' 発注内示対象年月

'   ( モジュール内ワーク )
    Dim wSelect                 As String
    Dim wCnt1                   As Integer
    Dim wCnt2                   As Integer
'*----------------------------------------------------------------------------------------------*
'*
'*              対 象 年 月 獲 得   [ pKB 0:全て、1:内示、2:生産計画、3:発注内示 ]
'*
'*----------------------------------------------------------------------------------------------*
Public Function GetNJProc_YM(pKB As Byte)

    If pKB = 0 Then
        wCnt1 = 1: wCnt2 = 3
    Else
        wCnt1 = pKB: wCnt2 = pKB
    End If
    
    Do
        Call ItemsClearSYMCreate
        With SYMCreate
            .区分名称 = "内示年月"
            .区分桁数 = 1
            .値 = wCnt1
            .制御区分 = "X"
            
            Select Case wCnt1
                Case 1:     .値名称 = "内示対象年月"
                Case 2:     .値名称 = "生産計画年月"
                Case 3:     .値名称 = "発注内示年月"
            End Select
            
            .英数字1 = ""
            .作成者 = gTanto
        End With
        If wCnt1 <> 1 Then Call SYMInitCreate(False)
        
        wCnt1 = wCnt1 + 1
    Loop Until wCnt1 > wCnt2
    
    Select Case pKB
        Case 1:     gn_NJYM = ""
        Case 2:     gn_SKYM = ""
        Case 3:     gn_HNYM = ""
        Case Else
                    gn_NJYM = "": gn_SKYM = "": gn_HNYM = ""
    End Select

    wSelect = "Select 値, 英数字1 From 名称マスタ"
    wSelect = wSelect & " Where 区分名称 = '内示年月'"
    If pKB = 0 Then
        wSelect = wSelect & " And   区分タイプ <> 'K'"
    Else
        wSelect = wSelect & " And   値 = " & pKB
    End If

    If SYMRead(wSelect) Then
        Do Until Not SYM_RDSTS
            Select Case SYM.値
                Case "1":     gn_NJYM = RTrim$(SYM.英数字1)
                Case "2":     gn_SKYM = RTrim$(SYM.英数字1)
                Case "3":     gn_HNYM = RTrim$(SYM.英数字1)
            End Select
        
            If CisDB.RecordCount = 1 Then Exit Do
            Call SYMReadNext
        Loop
    End If
    
    Call SYMClose

End Function

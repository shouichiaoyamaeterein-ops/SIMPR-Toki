Attribute VB_Name = "PrintSelectM"
Option Explicit

'*--------------------------------------------------------------------------------
'*
'*   印刷先選択用　ワーク定義
'*
'*   モジュールID  : PrintSelectM.BAS
'*
'*
'*   作成日  : 2004/09/10   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Global Psel_PrtName()       As String
Global Psel_PrtNomal()      As Boolean
Global Psel_NowSelect       As String
Global Psel_ChgPrtName      As String
Global Psel_ChgFlg          As Boolean

Dim bPrintObj               As Printer
Dim bCnt                    As Long

'+--------------------------+
'+ 　プリンタ情報取得
'+--------------------------+
Public Sub PrtSelect_Init()

    ReDim Psel_PrtName(1)
    ReDim Psel_PrtNomal(1)
    If Printers.Count < 1 Then
       Exit Sub
    End If

    ReDim Psel_PrtName(Printers.Count - 1)
    ReDim Psel_PrtNomal(Printers.Count - 1)
    bCnt = 0
    ' 通常使うﾌﾟﾘﾝﾀ
    For Each bPrintObj In Printers
        If bPrintObj.DeviceName = Printer.DeviceName Then
           Psel_PrtNomal(bCnt) = True
           Psel_PrtName(bCnt) = bPrintObj.DeviceName
           bCnt = bCnt + 1
           Exit For
        End If
    Next bPrintObj
    
    For Each bPrintObj In Printers
        If bPrintObj.DeviceName <> Printer.DeviceName Then
           Psel_PrtName(bCnt) = bPrintObj.DeviceName
           bCnt = bCnt + 1
        End If
    Next bPrintObj
End Sub


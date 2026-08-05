Attribute VB_Name = "CTA0510LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   棚卸ＨＨＴ読取結果取込エラーリスト　発行
'*
'*   モジュールID  : CTA0510LB (Based CXH0250LB：入荷一覧表　発行)
'*
'*   作成日  :2019/07/05   By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
Public CisPrtCTA0510        As New CISFormPrint

    Public wErrorTbl()      As String
    Public iErrorCnt        As Long
    Public bErrorHak        As Boolean

Public Sub Insert_ErrorData(ByVal iCount As Long, ByVal wData As String)
    ReDim Preserve wErrorTbl(iCount)
    wErrorTbl(iCount) = wData
    iErrorCnt = iCount
End Sub


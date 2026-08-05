Attribute VB_Name = "CNJ3330LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   生産予定表　印刷（キンテック向けレイアウト）
'*
'*   モジュールID  : CNJ3330.Bas
'*
'*   戻値 : なし
'*
'*   作成日  :2005/06/06  By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
    Public CisPrt           As New CISFormPrint
    Public TGet_YM          As String
    Public TGet_YMD         As String
    Public opOutKB          As Integer

    Public PgBk_Torcd       As String       ' ページブレイクキー
    Public PrtK_Hinbn       As String       ' 明細行出力キー    ( 品番 )
    Public PrtK_Torcd       As String       ' 明細行出力キー    ( 手配先 )
    Public PrtK_Seban       As String       ' 明細行出力キー    ( 背番号 )
'*-------------------------------------------------------*
'*  稼働日数取得(社内＆発注分カレンダーマスタ)
'*-------------------------------------------------------*
Public Sub GetKadou()
    ItemsClearCLM
    With CLM
        .取引先区分 = "1"                   '発注
        .取引先 = ""                        '社内カレンダー
        .年 = Mid(RTrim(TGet_YM), 1, 4)
        .月 = Mid(RTrim(TGet_YM), 5, 2)
        .受入 = ""                          '社内カレンダー
        If Not CLMRead("", 1, 1) Then Exit Sub
    End With
End Sub

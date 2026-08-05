Attribute VB_Name = "CNJ0330LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   生産予定表　印刷
'*
'*   モジュールID  : CNJ0330.Bas
'*
'*   戻値 : なし
'*
'*   作成日  :2004/03/25  By CIS
'*   変更日  :
'*---------------------------------------------------------------------------------
    Public CisPrt           As New CISFormPrint
    Public TGet_YM          As String
    Public opOutKB          As Integer

    Public PgBk_Torcd       As String       ' ページブレイクキー
    Public PrtK_Hinbn       As String       ' 明細行出力キー    ( 品番 )
    Public PrtK_Torcd       As String       ' 明細行出力キー    ( 手配先 )
    Public PrtK_Ukeir       As String       ' 明細行出力キー    ( 受入 )


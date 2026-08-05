Attribute VB_Name = "CNJZ0320LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   材料発注内示表　印刷
'*
'*   モジュールID  : CNJZ0320.Bas
'*
'*   戻値 : なし
'*
'*   作成日  :2004/03/25  By CIS
'*   変更日  :2010/02/24              YMAdd:内示対象年月(加算値)をPublic項目に変更 ⇒ CNJZ0320.frm より移動
'*   変更日  :
'*---------------------------------------------------------------------------------
    Public CisPrt           As New CISFormPrint
    Public TGet_NJYM        As String           ' 内示年月
    Public TGet_YM          As String           ' 対象年月
    Public opOutKB          As Integer
    
    Public PgBk_Torcd       As String       ' ページブレイクキー
    Public PrtK_Hinbn       As Long         ' 明細行出力キー    ( 材料管理番号 )
    Public PrtK_Torcd       As String       ' 明細行出力キー    ( 手配先 )
    Public PrtK_Ukeir       As String       ' 明細行出力キー    ( 受入 )

'## ADD-S 2010/02/24 ##
    Public YMAdd            As Byte         ' 内示対象年月(加算値)
'## ADD-E 2010/02/24 ##


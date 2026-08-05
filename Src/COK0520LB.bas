Attribute VB_Name = "COK0520LB"
Option Explicit
'*--------------------------------------------------------------------------------
'*
'*   仕入先別売上実績表発行
'*
'*   モジュールID  : COK0520LB
'*
'*                  COK0520LPrint(SelStr)
'*
'*   引数 : PrintOut : 印刷先 (Optional)
'*
'*   戻値 : True  : 印刷ﾃﾞｰﾀあり（印刷済）
'*        : False : 印刷データなし
'*
'*   使用例:

'    If Not COK0520LPrint Then Exit Sub
'
'*   作成日  :2008/05/22   By CIS
'*   作成日  :2009/02/10   By CIS
'*   作成日  :2009/02/12   By CIS
'*---------------------------------------------------------------------------------
Public CisPrt               As CISFormPrint
Public COK0520_PrtCnt       As Long
Global COK0520_PrtKb      As Byte
Global COK0520_PrtKb2     As Byte
Public COK0520_PrtOut       As String       ' 2004/05/10 Add
Public COK0520_FaxObj       As Object 'MfRemote
Public COK0520_Frm          As CISFormContorl
Public COK0520_PrtChg       As Boolean
Public COK0520_Hikae        As Boolean
Public COK0520_Kara         As Boolean
Public COK0520_FaxConnect   As Boolean
Public COK0520_FaxProc      As Boolean
Public COK0520_FaxServer    As String
Public COK0520_BasePrinter  As String
Public COK0520_SQL1         As String
Public COK0520_SQL2         As String

Public COK0520_SQL3         As String       '2009/02/10
Public COK0520_SQL4         As String       '2009/02/12

Public COK0520_YM           As String
Public COK0520_MD1          As String
Public COK0520_MD2          As String
Public COK0520_MD3          As String
Public COK0520_MD4          As String
Public COK0520_MD5          As String
Public COK0520_MD6          As String

    'SendFrom(電話番号,氏名,会社名,所属,住所1,住所2,FAX番号,リモートID,携帯電話(30),メールアドレス(40))
Dim jw_Tel      As String
Dim jw_Name     As String
Dim jw_Cmpany   As String
Dim jw_Busyo    As String
Dim jw_Fax      As String
Dim Fax_Shigai  As String
Dim Fax_No      As String
Dim TextRec     As String
Dim wTorcd()    As String
Dim wTId        As Long
Dim blnReturn   As Boolean


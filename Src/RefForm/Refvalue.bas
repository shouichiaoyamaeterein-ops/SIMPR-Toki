Attribute VB_Name = "RefValue"
Option Explicit
'   検索用(PB_LOOK)ﾌｫｰｶｽ位置
    Global Rv_GotFocus  As String
' SELECT
    Global Nm_Select    As String
    Global Rv_Select    As String
' CALL
    Global RV_Call      As String
    Global Rv_Call2     As String
    Global Rv_Call3     As String
    Global Rv_CallRtn   As Integer
    Global Rv_Call3Rtn  As Integer
    Global Rv_CallObj   As Object
    Global Rv_CallObj2  As Object
'   RETURN
    Global RV_Rtn       As Boolean
'   検索画面表示位置
    Global RV_Left      As Integer
    Global RV_Top       As Integer

'   日付入力用
    Global RV_SYmd      As String
    Global RV_EYmd      As String
    Global RV_Date      As Date
'   名称取得用
    Global RV_Kubun      As String
    Global RV_Code       As String
    Global RV_Name       As String
    Global RV_NameR      As String
    Global RV_Su1        As Double
' 名称マスタ検索
    Global RVI_Kubun    As String
    Global RVI_Seigyo   As String
' 取引先マスタ検索
    Global RV_TorcdK    As String
    Global RV_Torcd     As String
    Global RV_Kojyo     As String
    Global RV_TorKb     As Long
    Global RV_TorRName  As String
    Global RV_TorKb1    As Long
    Global RV_TorKb2    As Long
    Global RV_TorKb3    As Long
    Global RV_TorKb4    As Long
    Global RV_Ukeir     As String
' 品番マスタ検索
    Global RV_Hinbn     As String
    Global RV_SSeban    As String
    Global RV_Hinnm     As String
    Global RV_Seban     As String
    Global RV_Hinmoku   As String
    Global RV_Tehai     As String
    Global RVI_Torcd    As String
    Global RVI_Ukeir    As String
    Global RVI_Seban    As String
    Global RVI_Hinmoku  As Integer
    Global RVI_SK       As Integer
    Global RVI_Tehai    As String
' 材料マスタ検索
    Global RV_Zaist     As String
    Global RV_Itatu     As Double
    Global RV_Width     As Double
    Global RV_Long      As Double
    Global RV_Size      As Byte
    Global RV_Zainm     As String
    Global RV_HinZai    As Byte     '******* 2008/02/13 ADD
    Global RV_JHinbn    As String   '******* 2008/02/13 ADD
    Global RV_HHinbn    As String   '******* 2008/02/13 ADD
    Global RV_SHinbn    As String   '******* 2008/02/13 ADD
    Global RV_Tani      As String   '******* 2008/02/13 ADD
' 単価検索
    Global RV_Tanka     As Double
' 入荷実績用
    Global Rv_HNo()     As Long
    Global Rv_HPage()   As Long
    Global Rv_NExec()   As Boolean
    Global Rv_NRow()    As Long
    Global Rv_NCall     As String
    Global Rv_NEndSu    As Long
    Global Rv_NTtlSu    As Long
    Global Rv_NYmd      As String
    Global Rv_NBin      As String
    Global Rvc_NKNo     As Long
' 生産計画
    Global RV_NJYM      As String
    Global RV_YM        As String
    Global RV_KOKBN     As String
' 出荷実績用
    Global Rv_SJNo()    As Long
    Global Rv_SExec()   As Boolean
    Global Rv_SRow()    As Long
    Global Rv_STtlSu    As Long
    '+------------------------------------------------ 2004/11/30 Insert
    Global Rv_LongStr   As String
    '+------------------------------------------------ 2004/11/30 Insert End
    Global Rv_KanriNo    As Long
' 取引先変換マスタ品番選択用
    '+------------------------------------------------ 2006/02/09 Insert
    Global RV_STorcd    As String
    Global RV_KTorcd    As String
    Global RV_KTorkj    As String
    Global RV_STorcd2   As String
    Global RV_STorkj2   As String
    Global RV_STorUke2  As String
    Global RV_CyokuKb   As Byte             ' 直直
    Global RV_DTorcd    As String
    Global RV_DTorkj    As String
    Global RV_DUke      As String
    Global RV_DTorNm    As String
    Global RV_NTorcd    As String
    Global RV_NTorkj    As String
    Global RV_NUkeir    As String
    Global RV_NTorNm    As String
    Global RV_HiJisya   As String
    Global RV_JTorcd    As String
    Global RV_TyokTorcd As String                   '2006/03/15
    Global RV_SelCnt    As Long             ' 品番選択件数
    Global RV_CrtRep    As Byte             ' 代表作成(0:無し,1:有り)
    '+------------------------------------------------ 2006/02/09 Insert End
    Global RV_DaikoK    As String           ' 代行契約先 2007/12/17 ADD
    '+------------------------------------------------ 2006/03/11 Insert Start
' 増減一括入力用
    Global RV_ZG_DSPHin As String
    Global RV_ZG_Syuyo  As Long
    Global RV_ZG_Naisu  As Long
    Global RV_ZG_ZenKai As Long
    Global RV_ZG_KonKai As Long
    Global RV_ZG_Yuko   As Long
    Global RV_ZG_ZGKB   As Byte
    Global RV_ZG_ZGSet  As Integer
    Global RV_ZG_ZGBin  As Integer
    Global RV_ZG_HJKB   As Byte
    Global RV_ZG_HJSet  As Integer
    
    Global RVS_ZG_SYMD()   As String
    Global RVS_ZG_EYMD()   As String
    Global RVS_ZG_ZGKB()   As Byte
    Global RVS_ZG_ZGSet()  As Integer
    Global RVS_ZG_ZGBin()  As Integer
    Global RVS_ZG_ZGNKB()  As Byte
    Global RVS_ZG_HJKB()   As Byte
    Global RVS_ZG_HJSet()  As Integer
    Global RVS_ZG_HJNKB()  As Byte
    Global RVS_ZG_Count    As Integer
    Global RVS_ZG_Kan()    As String
    Global RVS_ZG_RuiMai() As Integer
    '+------------------------------------------------ 2006/03/11 Insert End
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/16 INSERT START
    Global RV_Y             As String
    Global RV_Ki            As Byte
    Global RV_Zaisyu        As String
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/16 INSERT END
'-----------------------------------------------------------------------' 2009.02.26 add start
' JIS規格変換
    Global RV_JIS_Zai       As String       ' 材質
    Global RV_JIS_Num       As String       ' JIS規格
'-----------------------------------------------------------------------' 2009.02.26 add start
'************************************'
'*                                  *
'*        各  種  検   索            *
'*                                  *
'************************************'
''   ( 各種検索 )
'           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
'Private Sub PB_Look_Click()
'    If PB_Look.Tag = "iH1_Kubun" Then Call Look_Kubun: Exit Sub
'End Sub
''+----------------------------+
''+      区分名
''+----------------------------+
'Private Sub iH1_Kubun_GotFocus()
'    PB_Look.Visible = True
'    PB_Look.Tag = ActiveControl.Name
'End Sub
'Private Sub iH1_Kubun_LostFocus()
'    If ActiveControl.Name <> "PB_Look" Then
'       PB_Look.Visible = False
'       PB_Look.Tag = ""
'    End If
'End Sub
'

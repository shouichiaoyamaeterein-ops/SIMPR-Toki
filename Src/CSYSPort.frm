VERSION 5.00
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Begin VB.Form CSYSPort 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "通信ポート設定"
   ClientHeight    =   3795
   ClientLeft      =   5670
   ClientTop       =   3120
   ClientWidth     =   4485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   0
      ScaleHeight     =   3795
      ScaleWidth      =   4485
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4485
      Begin VB.ComboBox B1Comb_HandShak 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1995
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   7
         Top             =   3315
         Width           =   2265
      End
      Begin VB.ComboBox B1Comb_StopLen 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1995
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   6
         Top             =   2805
         Width           =   2265
      End
      Begin VB.ComboBox B1Comb_DataLen 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1995
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   5
         Top             =   2310
         Width           =   2265
      End
      Begin VB.ComboBox B1Comb_Parity 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1995
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   4
         Top             =   1830
         Width           =   2265
      End
      Begin VB.ComboBox B1Comb_Speed 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1995
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   3
         Top             =   1335
         Width           =   2265
      End
      Begin VB.ComboBox B1Comb_Port 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1995
         Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
         TabIndex        =   2
         Top             =   825
         Width           =   2265
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00808080&
         Height          =   615
         Left            =   2865
         ScaleHeight     =   555
         ScaleWidth      =   1515
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   1575
         Begin CisBtn_60.CisBtn PB_END 
            Height          =   555
            Left            =   750
            Top             =   0
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   979
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BtnCaption      =   "戻る"
            BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BtnForeColor    =   16711680
            BtnShadowColor  =   16777215
            PFCaption       =   "( F12 )"
            BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PFForeColor     =   12648447
            PFShadowColor   =   0
            PFBackColor     =   16711680
         End
         Begin CisBtn_60.CisBtn PB_ENT 
            Height          =   555
            Left            =   0
            Top             =   0
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   979
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BtnCaption      =   "更新"
            BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BtnForeColor    =   128
            BtnShadowColor  =   16777215
            PFCaption       =   "(Enter)"
            BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PFForeColor     =   16777215
            PFShadowColor   =   0
            PFBackColor     =   128
         End
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   -15
         X2              =   4515
         Y1              =   675
         Y2              =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "(６) ハンドシェイク"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   15
         Left            =   165
         TabIndex        =   14
         Top             =   3375
         Width           =   1605
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "(５) ストップビット長"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   13
         Left            =   165
         TabIndex        =   13
         Top             =   2880
         Width           =   1740
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "(４) データビット長"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   11
         Left            =   165
         TabIndex        =   12
         Top             =   2385
         Width           =   1650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "(３) パリティ"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   9
         Left            =   165
         TabIndex        =   11
         Top             =   1890
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "(２) 通信速度"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   7
         Left            =   165
         TabIndex        =   10
         Top             =   1410
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "(１) 通信ポート"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   165
         TabIndex        =   9
         Top             =   900
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   4530
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Label lb_Section 
         BackStyle       =   0  '透明
         Caption         =   "かんばんソーター"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   165
         TabIndex        =   8
         Top             =   210
         Width           =   2475
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  '不透明
         BorderColor     =   &H00000000&
         Height          =   495
         Left            =   45
         Shape           =   4  '丸みのある長方形
         Top             =   75
         Width           =   2730
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "(１) 通信ポート"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   195
         TabIndex        =   15
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "(２) 通信速度"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   195
         TabIndex        =   16
         Top             =   1410
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "(３) パリティ"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   195
         TabIndex        =   17
         Top             =   1890
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "(４) データビット長"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   10
         Left            =   195
         TabIndex        =   18
         Top             =   2385
         Width           =   1650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "(５) ストップビット長"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   12
         Left            =   195
         TabIndex        =   19
         Top             =   2880
         Width           =   1740
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "(６) ハンドシェイク"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   14
         Left            =   195
         TabIndex        =   20
         Top             =   3375
         Width           =   1605
      End
   End
End
Attribute VB_Name = "CSYSPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  通信ポート設定
'**       フォームID    :  CSYSPort
'**       処理概要      :
'**
'**       作  成  日    :  2004/02/09  By CIS
'**       変  更  日    :  2006/02/02  By CIS  INIファイル-クライアント対応
'**       変  更  日    :
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    
    Dim CisFrm          As CISFormContorl
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Call CisFrm.InitFld
'
    Call CisFrm.Frm_Center(False)
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

'   *-----------------------------------------------*
'   * 通信装置　初期設定情報 確認 & 通信装置コンボ作成 *
'   *-----------------------------------------------*
'   * 各コンボ作成 *
    
'   < ポート >
    With B1Comb_Port
        .AddItem Edit_Item_Value("")
        .AddItem Edit_Item_Value("COM1")
        .AddItem Edit_Item_Value("COM2")
        .AddItem Edit_Item_Value("COM3")
        .AddItem Edit_Item_Value("COM4")
        .AddItem Edit_Item_Value("COM5")
        .AddItem Edit_Item_Value("COM6")
        .AddItem Edit_Item_Value("COM7")
        .AddItem Edit_Item_Value("COM8")
        .AddItem Edit_Item_Value("COM9")
    End With

'   < 通信速度 >
    With B1Comb_Speed
        .AddItem Edit_Item_Value("")
        .AddItem Edit_Item_Value("2400")
        .AddItem Edit_Item_Value("4800")
        .AddItem Edit_Item_Value("9600")
        .AddItem Edit_Item_Value("14400")
        .AddItem Edit_Item_Value("19200")
        .AddItem Edit_Item_Value("38400")
        .AddItem Edit_Item_Value("57600")
        .AddItem Edit_Item_Value("115200")
        .AddItem Edit_Item_Value("256000")
    End With
    
'   < パリティ >
    With B1Comb_Parity
        .AddItem Edit_Item_Value("")
        .AddItem Edit_Item_Value("N:パリティなし")
        .AddItem Edit_Item_Value("E:偶数パリティ")
        .AddItem Edit_Item_Value("O:奇数パリティ")
        .AddItem Edit_Item_Value("S:スペースパリティ")
        .AddItem Edit_Item_Value("M:マークパリティ")
    End With
    
'   < データビット長 >
    With B1Comb_DataLen
        .AddItem Edit_Item_Value("")
        .AddItem Edit_Item_Value("4")
        .AddItem Edit_Item_Value("5")
        .AddItem Edit_Item_Value("6")
        .AddItem Edit_Item_Value("7")
        .AddItem Edit_Item_Value("8")
    End With
    
'   < ストップビット長 >
    With B1Comb_StopLen
        .AddItem Edit_Item_Value("")
        .AddItem Edit_Item_Value("1")
        .AddItem Edit_Item_Value("1.5")
        .AddItem Edit_Item_Value("2")
    End With
    
'   < ハンドシェイク >
    With B1Comb_HandShak
        .AddItem Edit_Item_Value("")
        .AddItem Edit_Item_Value("0:なし")
        .AddItem Edit_Item_Value("1:XON/XOFF")
        .AddItem Edit_Item_Value("2:RTS/CTS")
        .AddItem Edit_Item_Value("3:XON/XOFF･RTS/CTS")
    End With
    
    Call BodySet
    
    On Error Resume Next
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》       +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    If Not FormAct Then Exit Function
    
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    
    Select Case Key_Code
            Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
            Case vbKeyReturn
                                If PB_ENT.Visible Then                      '【入力】
                                    Call PB_ENT_Click
                                End If
            Case Else:          Exit Function
    End Select
    Key_Acc = True
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    Me.Hide
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
   Call ReturnPress
End Sub
'+---------------------+
'+   B1Comb_Port
'+---------------------+
Private Sub B1Comb_Port_GotFocus()
    B1Comb_Port.BackColor = gIGotSel
End Sub
Private Sub B1Comb_Port_LostFocus()
    B1Comb_Port.BackColor = gILostSel
End Sub
'+---------------------+
'+   B1Comb_Speed
'+---------------------+
Private Sub B1Comb_Speed_GotFocus()
    B1Comb_Speed.BackColor = gIGotSel
End Sub
Private Sub B1Comb_Speed_LostFocus()
    B1Comb_Speed.BackColor = gILostSel
End Sub
'+---------------------+
'+   B1Comb_Parity
'+---------------------+
Private Sub B1Comb_Parity_GotFocus()
    B1Comb_Parity.BackColor = gIGotSel
End Sub
Private Sub B1Comb_Parity_LostFocus()
    B1Comb_Parity.BackColor = gILostSel
End Sub
'+---------------------+
'+   B1Comb_DataLen
'+---------------------+
Private Sub B1Comb_DataLen_GotFocus()
    B1Comb_DataLen.BackColor = gIGotSel
End Sub
Private Sub B1Comb_DataLen_LostFocus()
    B1Comb_DataLen.BackColor = gILostSel
End Sub
'+---------------------+
'+   B1Comb_StopLen
'+---------------------+
Private Sub B1Comb_StopLen_GotFocus()
    B1Comb_StopLen.BackColor = gIGotSel
End Sub
Private Sub B1Comb_StopLen_LostFocus()
    B1Comb_StopLen.BackColor = gILostSel
End Sub
'+---------------------+
'+   B1Comb_HandShak
'+---------------------+
Private Sub B1Comb_HandShak_GotFocus()
    B1Comb_HandShak.BackColor = gIGotSel
End Sub
Private Sub B1Comb_HandShak_LostFocus()
    B1Comb_HandShak.BackColor = gILostSel
End Sub
'****************************
'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Sub BodySet()
    Dim wApparatus  As String
    Dim wSetting()  As String
    
    With CisFun
        wApparatus = IRN_LinkNo + 1
        .INI_Section = "Apparatus"
        .INI_FileID = gIniExe & gIniName
        .INI_Key = wApparatus
        .INI_Default = ""
        .GetIni
        
        .INI_Section = Trim(.INI_String)
        lb_Section.Tag = Trim(.INI_String)
        lb_Section = IRN_LinkStr
        
        ' '----- CommPort獲得 -----
'                                           < 2006/02/02 ADD START >
        .INI_FileID = gIniExe & gwsIniName
'                                           < 2006/02/02 ADD  END  >
        .INI_Default = ""
        .INI_Key = "CommPort"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_String = "1"
           .PutIni
        End If
        gLong = CisFun.Val2(.INI_String)
        If gLong = 0 Then gLong = 1
        B1Comb_Port.ListIndex = gLong
        ' '----- 通信速度獲得 -----
'                                           < 2006/02/02 ADD START >
        .INI_FileID = gIniExe & gIniName
'                                           < 2006/02/02 ADD  END  >
        .INI_Default = ""
        .INI_Key = "Setting"
        .GetIni
        If Trim(.INI_String) = "" Then
           Select Case wApparatus
                Case "2": .INI_String = "19200,N,8,1"
                Case "3": .INI_String = "9600,E,8,1"
                Case "4": .INI_String = "38400,N,8,1"
                Case Else: .INI_String = "9600,n,8,1"
           End Select
        End If
        wSetting = Split(RTrim$(.INI_String), ",", 4)
        B1Comb_Speed = Edit_Item_Value(RTrim$(wSetting(0)))
'       * パリティ *
        Select Case UCase(RTrim$(wSetting(1)))
            Case "N":   gStr = RTrim$(wSetting(1)) & ":パリティなし"
            Case "E":   gStr = RTrim$(wSetting(1)) & ":偶数パリティ"
            Case "O":   gStr = RTrim$(wSetting(1)) & ":奇数パリティ"
            Case "S":   gStr = RTrim$(wSetting(1)) & ":スペースパリティ"
            Case "M":   gStr = RTrim$(wSetting(1)) & ":マークパリティ"
            Case Else:  gStr = ""
        End Select
        B1Comb_Parity = Edit_Item_Value(RTrim$(gStr))
        B1Comb_DataLen = Edit_Item_Value(RTrim$(wSetting(2)))
        B1Comb_StopLen = Edit_Item_Value(RTrim$(wSetting(3)))
'       * ハンドシェイク *
        ' '----- HandShaking獲得 -----
        .INI_Default = ""
        .INI_Key = "HandShaking"
        .GetIni
        If Trim(.INI_String) = "" Then
           Select Case wApparatus
                Case "3": .INI_String = "3"
                Case "2", "4", 6: .INI_String = "0"
                Case Else: .INI_String = "1"
           End Select
           .PutIni
        End If
        Select Case RTrim$(.INI_String)
            Case 0:   gStr = "0:なし"
            Case 1:   gStr = "1:XON/XOFF"
            Case 2:   gStr = "2:RTS/CTS"
            Case 3:   gStr = "3:XON/XOFF･RTS/CTS"
        End Select
        B1Comb_HandShak = Edit_Item_Value(RTrim$(gStr))
    End With
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If Not Body1Chk Then GoTo ReturnPress_Ed
    If Not DBPut Then GoTo ReturnPress_Ed
    Call PB_END_Click
ReturnPress_Ed:
    Call CisFrm.MousePT(1)
End Sub
'****************************
'*      BODY1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
 
    If RTrim$(B1Comb_Port) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "     通信ポートを選択して下さい           "
            .MB_MSG(3) = ""
            .MB_Title = "": .MB_Button = Error
            .MBOX
        End With
        B1Comb_Port.SetFocus
        Exit Function
    End If
 
    If RTrim$(B1Comb_Speed) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "     通信速度を選択して下さい         "
            .MB_MSG(3) = ""
            .MB_Title = "": .MB_Button = Error
            .MBOX
        End With
        B1Comb_Speed.SetFocus
        Exit Function
    End If
 
    If RTrim$(B1Comb_Parity) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "     パリティを選択して下さい         "
            .MB_MSG(3) = ""
            .MB_Title = "": .MB_Button = Error
            .MBOX
        End With
        B1Comb_Parity.SetFocus
        Exit Function
    End If
 
    If RTrim$(B1Comb_DataLen) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "     データビット長を選択して下さい         "
            .MB_MSG(3) = ""
            .MB_Title = "": .MB_Button = Error
            .MBOX
        End With
        B1Comb_DataLen.SetFocus
        Exit Function
    End If
 
    If RTrim$(B1Comb_StopLen) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "     ストップビット長を選択して下さい         "
            .MB_MSG(3) = ""
            .MB_Title = "": .MB_Button = Error
            .MBOX
        End With
        B1Comb_StopLen.SetFocus
        Exit Function
    End If
 
    If RTrim$(B1Comb_HandShak) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "     ハンドシェイクを選択して下さい         "
            .MB_MSG(3) = ""
            .MB_Title = "": .MB_Button = Error
            .MBOX
        End With
        B1Comb_HandShak.SetFocus
        Exit Function
    End If
    
    Body1Chk = True
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    
    On Error GoTo DBPut_Err
    
    With CisFun
        .INI_FileID = gIniExe & gIniName
        .INI_Section = Trim(lb_Section.Tag)
        ' '----- CommPort獲得 -----
'                                           < 2006/02/02 ADD START >
        .INI_FileID = gIniExe & gwsIniName
'                                           < 2006/02/02 ADD  END  >
        .INI_Key = "CommPort"
        .INI_String = B1Comb_Port.ListIndex
        .PutIni
    
        gStr = ""
        Select Case B1Comb_Speed.ListIndex
            Case 1: gStr = "2400,"
            Case 2: gStr = "4800,"
            Case 3: gStr = "9600,"
            Case 4: gStr = "14400,"
            Case 5: gStr = "19200,"
            Case 6: gStr = "38400,"
            Case 7: gStr = "57600,"
            Case 8: gStr = "115200,"
            Case 9: gStr = "256000,"
        End Select
        Select Case B1Comb_Parity.ListIndex
            Case 1: gStr = gStr & "N,"
            Case 2: gStr = gStr & "E,"
            Case 3: gStr = gStr & "O,"
            Case 4: gStr = gStr & "S,"
            Case 5: gStr = gStr & "M,"
        End Select
        Select Case B1Comb_DataLen.ListIndex
            Case 1: gStr = gStr & "4,"
            Case 2: gStr = gStr & "5,"
            Case 3: gStr = gStr & "6,"
            Case 4: gStr = gStr & "7,"
            Case 5: gStr = gStr & "8,"
        End Select
        Select Case B1Comb_StopLen.ListIndex
            Case 1: gStr = gStr & "1"
            Case 2: gStr = gStr & "1.5"
            Case 3: gStr = gStr & "2"
        End Select
        
'                                           < 2006/02/02 ADD START >
        .INI_FileID = gIniExe & gIniName
'                                           < 2006/02/02 ADD  END  >
        .INI_Key = "Setting"
        .INI_String = gStr
        .PutIni
    
        .INI_Key = "HandShaking"
        .INI_String = B1Comb_HandShak.ListIndex - 1
        .PutIni
    End With
    
    On Error GoTo 0
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'*---------------------------------------------------------------------------------------------*
'*---------------------------------------------------------------------------------------------*
'*---------------------------------------------------------------------------------------------*
'*---------------------------------------------------------------------------------------------*
Private Function Edit_Item_Value(pVal As String)
    Edit_Item_Value = RTrim$(pVal) & Space(30 - CisFun.Len2(RTrim$(pVal)))
End Function

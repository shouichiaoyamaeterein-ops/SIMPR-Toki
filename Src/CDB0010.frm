VERSION 5.00
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Begin VB.Form CDB0010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "データベース退避復元"
   ClientHeight    =   10950
   ClientLeft      =   -45
   ClientTop       =   300
   ClientWidth     =   15090
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10950
   ScaleWidth      =   15090
   Begin VB.Timer Timer1 
      Left            =   735
      Top             =   540
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   11430
      ScaleHeight     =   735
      ScaleWidth      =   3240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   495
      Width           =   3300
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   675
         Left            =   1095
         Top             =   30
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1191
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   3
         BtnCaption      =   "開始"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
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
            Name            =   "ＭＳ 明朝"
            Size            =   9
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
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   675
         Left            =   2160
         Top             =   30
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1191
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   3
         BtnType         =   1
         BtnCaption      =   "終了"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
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
            Name            =   "ＭＳ 明朝"
            Size            =   9
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   675
         Left            =   30
         Top             =   30
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1191
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   3
         BtnType         =   11
         BtnCaption      =   "検索"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   33023
         BtnShadowColor  =   16777215
         PFCaption       =   "( F7 )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   12648447
         PFShadowColor   =   0
         PFBackColor     =   33023
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   1005
      Left            =   555
      ScaleHeight     =   945
      ScaleWidth      =   14130
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1425
      Width           =   14190
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   855
         Left            =   45
         Top             =   45
         Width           =   14040
         _ExtentX        =   24765
         _ExtentY        =   1508
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   795
            Left            =   2925
            Top             =   30
            Width           =   11085
            _ExtentX        =   19553
            _ExtentY        =   1402
            ForeColor       =   12582912
            Caption         =   "退避先・復元元　フォルダ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   90
            Begin Cis3D_v60.CIS3D PB_Dir 
               Height          =   390
               Left            =   9540
               Top             =   195
               Width           =   450
               _ExtentX        =   794
               _ExtentY        =   688
               Caption         =   "…"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cBoderWidth     =   3
               cAlingnment     =   7
               cButton         =   -1  'True
               cPositionX      =   20
               cPositionY      =   20
            End
            Begin CisText_V60.CisText iH1_Dir 
               Height          =   405
               Left            =   3105
               TabIndex        =   0
               Top             =   195
               Width           =   6525
               _ExtentX        =   11509
               _ExtentY        =   714
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   $"CDB0010.frx":0000
               MaxLength       =   255
            End
            Begin Cis3D_v60.CIS3D PB_Open 
               Height          =   390
               Left            =   10050
               Top             =   195
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   688
               Caption         =   "開く"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cBoderWidth     =   3
               cAlingnment     =   7
               cButton         =   -1  'True
               cPositionX      =   20
               cPositionY      =   20
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   795
            Left            =   30
            Top             =   30
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   1402
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cBoderStyle     =   1
            Begin VB.OptionButton H1Op_Proc 
               BackColor       =   &H0080FF80&
               Caption         =   "退  避"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   750
               Index           =   0
               Left            =   30
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   1
               Top             =   30
               Value           =   -1  'True
               Width           =   1410
            End
            Begin VB.OptionButton H1Op_Proc 
               Caption         =   "復 元"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   750
               Index           =   1
               Left            =   1455
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   2
               Top             =   30
               Width           =   1410
            End
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8265
      Left            =   555
      ScaleHeight     =   8205
      ScaleWidth      =   14100
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2580
      Width           =   14160
      Begin vsFlexLib.vsFlexArray VsGrid1 
         Height          =   8025
         Left            =   75
         TabIndex        =   3
         Top             =   75
         Width           =   13950
         _Version        =   196608
         _ExtentX        =   24606
         _ExtentY        =   14155
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "No | 処理日時 | 退避ファイル名 | 退避処理|処理 | 復元日時 |  退避先|存在|処理CD|日付|終了日時|秒|SELECT|"
         Rows            =   23
         Cols            =   14
         BackColor       =   12648447
         BackColorBkg    =   14737632
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   350
         SelectionMode   =   1
         BackColorAlternate=   12648447
      End
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 データベース退避復元 】"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cBoderWidth     =   3
      cAlingnment     =   1
      cPositionX      =   1200
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14355
         Top             =   75
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   450
         BackColor       =   14737632
         ForeColor       =   16711680
         Caption         =   "SIMPR-A"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   14737632
         cBoderColor2    =   14737632
         cFont3DColor1   =   8454143
         cFont3DColor2   =   16777215
         cFont3D         =   1
         cAlingnment     =   7
         cBoderStyle     =   2
      End
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   315
         Left            =   11775
         Top             =   30
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   556
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXXXXXXX2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cBoderStyle     =   1
         cPositionX      =   50
      End
   End
   Begin VB.Label Del_Color 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Label1"
      Height          =   330
      Left            =   3540
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   2190
   End
End
Attribute VB_Name = "CDB0010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   データベース退避復元
'**       フォームID    :   CDB0010
'**       処理概要      :
'**
'**       作  成  日    :   2004/03/15  By CIS
'**       変　更　日    :   2005.04.01   By CIS データ退避時にタイムアウトを1800に
'**       変  更  日    :
'**
'**********************************************************'

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 4   ' 入力域切替
    Dim CISFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    
    Dim mGotColor       As Long
    Dim mLostColor      As Long
    Dim mRstFlg         As Boolean
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   # 二 重 起 動 防 止  #
'   #-------------------#
    If Not Dupli_Start Then End
'   #-------------------#
'   # 初 期 内 容 取 得  #
'   #-------------------#
'    Call IniGet
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Set CISFrm = New CISFormContorl
    
    Set CISFrm.ThisForm = Me
    Call CISFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
'
'    Dummy.Left = -1000
'    Dummy.Enabled = False
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    
    mGotColor = H1Op_Proc(0).BackColor
    mLostColor = H1Op_Proc(1).BackColor
'   #------------------#
'   #  ｸﾞﾘｯﾄﾞ初期処理   #
'   #------------------#
    With CisVsGrid
        Set .GridObj = VsGrid1              ' ｸﾞﾘｯﾄﾞｵﾌﾞｼﾞｪｸﾄ名 ｾｯﾄ
        .FixedHeight = 360                  ' 固定行(見出し)高さ設定(省略時は300)
        .PatanMax = 1                       ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ最大数
        .PatanInit = 0                      ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ初期値
        .SelectCol = 13                     ' セル選択処理ﾜｰｸ列
        .NoSpaceCol = 1
        .NumberSet = NoSpace
'        Set .SelectDispObj = B1lb_SKensu    ' セル選択件数表示ｵﾌﾞｼﾞｪｸﾄｾｯﾄ
' ﾌﾟﾛｸﾞﾗﾑ作成時、以下のﾌﾗｸﾞをｾｯﾄし、ｸﾘｯﾌﾟﾎﾞｰﾄﾞよりｺｰﾄﾞの貼り付けを行う。(GridInit)
        .InitGet = False                    ' ｸﾞﾘｯﾄﾞ内容初期設定 ｸﾘｯﾌﾟﾎﾞｰﾄﾞ作成ﾌﾗｸﾞ

        .Init                               ' ｸﾞﾘｯﾄﾞ初期処理 実行
    End With
    Call GridInit                           ' ｸﾞﾘｯﾄﾞ初期表示
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "BACK"
    Call DispChange(ProcHB)
    
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CISFrm.Frm_Center(True, x1024y768)

    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    H1Op_Proc(db_Proc).Value = True
    
    Call BodyDisp
    
    FormAct = Not FormAct
'
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》                 +
'+-------------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Cancel = 1
    Call PB_END_Click
End Sub
'+-------------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》         +
'+-------------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    Select Case Key_Code
           Case vbKeyF1:    H1Op_Proc(0).Value = True
           Case vbKeyF2:    H1Op_Proc(1).Value = True
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      '【入力】
'                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
'                                  Dummy.Enabled = False
                               End If
           Case Else:          Exit Function
    End Select
    Key_Acc = True
End Function


'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
'+---------------------+
'+   履歴出力(Text)
'+---------------------+
    Call TextPut
'+---------------------+
'+   ﾃﾞｰﾀﾍﾞｰｽ切断
'+---------------------+
    Call CisDB.DBDISConnect
'+---------------------+
'+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
'+---------------------+
    Call CISFrm.UnLoadDisp
    
    With CisFun
        gLong = CisFun.Val2(Command)
        If gLong <> 0 Then
            If .TackCheck("SIMPR-A.exe") = 9 Then
               Call Shell(gIniExe & "SIMPR-A.exe " & Format(gLong, "0"), 1)
            End If
        End If
    End With
    
    End
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
   Call ReturnPress
End Sub

Private Sub PB_Open_Click()
    Dim resL As Long
    resL = ShellExecute(Me.hWnd, "explore", iH1_Dir, vbNullString, vbNullString, 1)
    If resL = 2 Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "      退避先・復元元　フォルダが存在しません。         "
            .MB_Title = "ﾌｫﾙﾀﾞ確認"
            .MB_Button = OK
            .MBOX
        End With
    End If
End Sub

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Dir" Then Call PB_Dir_Click: Exit Sub
End Sub
'+---------------------+
'+   実行ﾃﾞｨﾚｸﾄﾘ　参照 +
'+---------------------+
Private Sub PB_Dir_Click()
    With CisFun
        .Folder_Flags = ffDirOnly
        .Folder_DefDir = iH1_Dir
        .FolderSel
        gStr = RTrim(.Folder_Result)
    End With
    If Trim(gStr) <> "" Then
        iH1_Dir = gStr
        db_BackDir = RTrim(iH1_Dir)
        If Right(db_BackDir, 1) <> "\" Then db_BackDir = db_BackDir & "\"
    
        Call GridSet
    End If
    
    iH1_Dir.SetFocus
End Sub
'+---------------------------------------------+
'+  iH1_Dir    フォルダ
'+---------------------------------------------+
Private Sub iH1_Dir_GotFocus()
    PB_Look.Visible = True
    PB_Look.Tag = ActiveControl.Name
End Sub
Private Sub iH1_Dir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
       PB_Look.Visible = False
       PB_Look.Tag = ""
    End If
End Sub
'+----------------------------------+
'+      退避・復元選択
'+----------------------------------+
Private Sub H1Op_Proc_Click(Index As Integer)
    If Index = 0 Then
       ProcHB = "BACK"
       Call DispChange(ProcHB)
       H1Op_Proc(0).BackColor = mGotColor
       H1Op_Proc(1).BackColor = mLostColor
       iH1_Dir.SetFocus
    Else
       ProcHB = "RST"
       Call DispChange(ProcHB)
       H1Op_Proc(0).BackColor = mLostColor
       H1Op_Proc(1).BackColor = mGotColor
       VsGrid1.SetFocus
    End If
End Sub
'****************************
'*                          *
'*      HEAD/BODY 処理      *
'*                          *
'****************************
'****************************
'*      入力切替 処理        *
'****************************
Private Function DispChange(Pro As String)
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
'    If Dummy.Visible Then
'       Dummy.Enabled = True
'       Dummy.SetFocus
'    End If
'
    B1_Area1.BackColor = gPGotSel
    B1_Area1.Enabled = True
    Select Case Pro
        Case "BACK"
            B1_Area1.BackColor = gPLostSel
            B1_Area1.Enabled = False
    End Select
'    PB_Look.Visible = H1Mode
'    PB_END.Visible = H1Mode
'' ﾀﾞﾐｰ項目(入力禁止)
'    Dummy.Enabled = False
'
End Function
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CISFrm.MousePT(11)
    
    db_BackDir = RTrim(iH1_Dir)
    If Right(db_BackDir, 1) <> "\" Then db_BackDir = db_BackDir & "\"
    With CisFun
        .INI_Section = "DataBase"
        .INI_FileID = gIniExe & gIniName
        .INI_Key = "BackDir"
        .INI_String = db_BackDir
        .PutIni
    End With

    If ProcHB = "BACK" Then
       gStr = vbCr
       gStr = gStr & "     データベースの退避処理を行います         "
       gStr = gStr & vbCr
       gStr = gStr & vbCr
       gStr = gStr & vbCr
       gStr = gStr & "           よろしいですか？                   "
       gStr = gStr & vbCr
       gStr = gStr & vbCr
       gStr = gStr & vbCr
       gStr = gStr & vbCr
       
       If MsgBox(gStr, vbExclamation + vbOKCancel + vbDefaultButton2, "退避確認") <> vbOK Then GoTo ReturnPress_Ed
       
       db_ProcType = ""
       
       Call db_BackUp
       
       Call GridSet
       
       GoTo ReturnPress_Ed
    End If
    If Trim(ProcHB) = "RST" Then
       If Trim(VsGrid1.TextMatrix(1, CisVsGrid.FixedGet(1))) = "" Then
          With CisFun
               .MB_Lines = 5
               .MB_MSG(3) = "         復元ファイルが存在しません              "
               .MB_Button = Error
               .MBOX
          End With
          GoTo ReturnPress_Ed
       End If
       If Trim(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid.FixedGet(1))) = "" Then
          With CisFun
               .MB_Lines = 5
               .MB_MSG(3) = "         復元ファイルを選択して下さい              "
               .MB_Button = Error
               .MBOX
          End With
          GoTo ReturnPress_Ed
       End If
       If Trim(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid.FixedGet(7))) = "" Then
          Call CisVsGrid.vsColor(消去全部)
          Call CisVsGrid.vsColor
       Else
          With CisFun
               .MB_Lines = 5
               .MB_MSG(1) = "         退避ファイルが存在しません              "
               .MB_MSG(3) = "     退避ファイルを確認後、再度行って下さい      "
               .MB_Button = Error
               .MBOX
          End With
          GoTo ReturnPress_Ed
       End If
       gStr = vbCr
       gStr = gStr & "     データベースの復元処理を行います         "
       gStr = gStr & vbCr
       gStr = gStr & vbCr
       gStr = gStr & "           よろしいですか？                   "
       gStr = gStr & vbCr
       
       If MsgBox(gStr, vbExclamation + vbOKCancel + vbDefaultButton2, "復元確認") <> vbOK Then
          Call CisVsGrid.vsColor(消去全部)
          GoTo ReturnPress_Ed
       End If
       
       db_ProcType = ""
       With VsGrid1
            db_SaveFile = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(6))) & _
                          Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(2)))
            db_FileCopy = True
            If Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(8))) = "S" Then
               db_FileCopy = False
            End If
       End With
       If Trim(db_BackPath) = "" Then
           Call db_crtDevice("", True)
       End If
        ' DB接続解除
       Call CisDB.DBDISConnect
       With CisDB
            db_RstDBName = .DBName
            .DBName = "Master"
        '( DB パスワード )
            gCipher = 1
            gStr2 = GetSetting(gRegKey, "Environment", "DrowsSap", "")
            If RTrim$(gStr2) = "" Then gStr2 = CisDB.PassWord("")
            CisDB.PASS = RTrim(gStr2)
            .ConnectCount = 2
            Call .DBConnect(App.Title)
       End With

       If db_ReStore Then
          VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid.FixedGet(5)) = Format(Now(), "YYYY/MM/DD hh:nn:ss")
          mRstFlg = True
       Else
          VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid.FixedGet(5)) = "復元エラー"
       End If
       CisDB.DBName = db_RstDBName
       Call CisDB.DBDISConnect
       
       If Not IniGet Then End
       
       Call CisDB.DBDISConnect
       
       Call CisVsGrid.vsColor(消去全部)
       GoTo ReturnPress_Ed
    End If
ReturnPress_Ed:
   Call CISFrm.MousePT(1)
End Sub
'+-----------------------------------+
'+      内容表示
'+-----------------------------------+
Private Sub BodyDisp()

    iH1_Dir = db_BackDir
    
    Call GridSet
End Sub
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
         Call .FixedSet(中中, 右中, あり, 0, 0, "No")
         Call .FixedSet(中中, 中中, あり, 21, 1, "処理日時", 英字)
         Call .FixedSet(中中, 左中, あり, 36, 2, "退避ファイル名")
         Call .FixedSet(中中, 中中, あり, 8, 3, "退避処理")
         Call .FixedSet(中中, 中中, あり, 9, 4, "処理時間")
         Call .FixedSet(中中, 中中, あり, 21, 5, "復元日時")
         Call .FixedSet(中中, 左中, あり, 50, 6, "退避先")
         
         Call .FixedSet(中中, 左中, あり, 0, 7, "存在")
         Call .FixedSet(中中, 左中, あり, 0, 8, "処理CD")
         Call .FixedSet(中中, 左中, あり, 0, 9, "日付")
         Call .FixedSet(中中, 左中, あり, 0, 10, "終了日時")
         Call .FixedSet(中中, 左中, あり, 0, 11, "秒")
         Call .FixedSet(中中, 右中, あり, 10, 12, "サイズ")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
         Call .FixedSet(中中, 左中, あり, 0, 13, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,12")
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
    Dim GridSetData() As String
    Dim wDelDate    As String
    Dim wDelDate2   As String
    Dim sqlAgentFlg As Boolean
    Dim sqlAgentFile As String
    Dim wDate       As Date
    Dim wSize       As Currency
'  初期ｾｯﾄ
    CisVsGrid.Clear
'
    db_BackDir = RTrim(iH1_Dir)
    If Right(db_BackDir, 1) <> "\" Then db_BackDir = db_BackDir & "\"

    With db_FileCtl
         .FE_FullPath = db_BackDir & db_LogFName
         Ttl_Kensu = .FileLine
    End With
    sqlAgentFlg = False
    With CisFun
         .INI_Section = "DataBase"
         .INI_FileID = gIniExe & gIniName
         .INI_Default = ""
         .INI_Key = "SqlAgent"
         .GetIni
         sqlAgentFile = RTrim(.INI_String)
         If Trim(sqlAgentFile) <> "" Then
            sqlAgentFlg = True
            Ttl_Kensu = Ttl_Kensu + 1
         Else
            .INI_String = ""
            .PutIni
         End If
    End With
    
    If Not db_Rireki("Input") Then Exit Sub
    
    wDelDate = Format(Now(), "yyyymmdd")
    With CisFun
        .Date_RsFormat = fDate
        .Date_Add wDelDate, D, db_LogDelDay * -1
        wDelDate2 = .Date_Result
    End With
    
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid1
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(ID)
        Next ID
        .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 1
    With VsGrid1
        Do Until Not db_Rireki("Read", gStr)
            GridSetData = Split(gStr, ",")
            If UBound(GridSetData) > 5 Then
               If wDelDate2 <= GridSetData(0) And Dir(GridSetData(5)) <> "" Then
                  .TextMatrix(ID, GridCol(1)) = GridSetData(1)  ' 開始日時
'                  If Dir(GridSetData(5)) = "" Then
'                     .Row = ID: .Col = 2
'                     .CellBackColor = Del_Color.BackColor
'                     .TextMatrix(ID, GridCol(7)) = "X"
'                  End If
                  With db_FileCtl
                       .FE_FullPath = GridSetData(5)
                       .GetPathName
                       VsGrid1.TextMatrix(ID, GridCol(6)) = .FE_Path
                       VsGrid1.TextMatrix(ID, GridCol(2)) = .FE_File
                       If Dir(GridSetData(5)) <> "" Then
                          wSize = FileLen(GridSetData(5)) / 1024 / 1024
                          VsGrid1.TextMatrix(ID, GridCol(12)) = Format(wSize, "#,##0.0") & " MB"
                       End If
                  End With
                  .TextMatrix(ID, GridCol(5)) = GridSetData(6)  ' 復元日時
                    
                  Select Case GridSetData(4)
                      Case "A": gStr = "自動"
                      Case "D": gStr = "削除"
                      Case Else
                          gStr = "手動"
                  End Select
                  .TextMatrix(ID, GridCol(3)) = gStr            ' 処理
                  .TextMatrix(ID, GridCol(8)) = GridSetData(4)  ' 処理CD
                  If CisFun.Val2(GridSetData(3)) <> 0 Then
                      gLong = CisFun.Val2(GridSetData(3))
                      gCnt1 = gLong \ 60
                      gCnt2 = gLong Mod 60
                      
                     .TextMatrix(ID, GridCol(4)) = Format(gCnt1, "00") & ":" & _
                                                   Format(gCnt2, "00")       ' 時間
                     .TextMatrix(ID, GridCol(11)) = GridSetData(3)  ' 時間
                  End If
                  .TextMatrix(ID, GridCol(9)) = GridSetData(0)  ' 日付
                  .TextMatrix(ID, GridCol(10)) = GridSetData(1) ' 終了日時
                                    
                  ID = ID + 1
                End If
            End If
        Loop
        If sqlAgentFlg Then
           If Dir(sqlAgentFile) <> "" Then
              wDate = FileDateTime(sqlAgentFile)
              .TextMatrix(ID, GridCol(1)) = Format(wDate, "YYYY/MM/DD hh:nn:ss") ' 開始日時
                With db_FileCtl
                     .FE_FullPath = sqlAgentFile
                     .GetPathName
                     VsGrid1.TextMatrix(ID, GridCol(6)) = .FE_Path
                     VsGrid1.TextMatrix(ID, GridCol(2)) = .FE_File
                End With
                wSize = FileLen(sqlAgentFile) / 1024 / 1024
                VsGrid1.TextMatrix(ID, GridCol(12)) = Format(wSize, "#,##0.0") & " MB"
                .TextMatrix(ID, GridCol(3)) = "Agent"         ' 処理
                .TextMatrix(ID, GridCol(8)) = "S"             ' 処理CD
           End If
        End If
    End With
    With VsGrid1
        .Row = 1: .Col = 1
        .Redraw = True
    End With
    
    Call CisVsGrid.VSSort(False)
    Call CisVsGrid.ReNumber
'
    Call db_Rireki("Close")

End Sub
''******************************************'
''*****        行選択（クリック）       *****
''******************************************'
'Private Sub VSGrid1_Click()
'    If Trim(VsGrid1.TextMatrix(VsGrid1.Row, mVSObj.FixedGet(7))) = "" Then
'       Call mVSObj.vsColor(全部消去)
'       Call mVSObj.vsColor
'    End If
'End Sub

Private Sub VsGrid1_DblClick()
    Call ReturnPress
End Sub
'******************************************'
'*****        行選択（スペースｷｰ)      *****
'******************************************'
'Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> vbKeySpace Then Exit Sub
'    If Trim(VsGrid1.TextMatrix(VsGrid1.Row, mVSObj.FixedGet(7))) = "" Then
'       Call mVSObj.vsColor(全部消去)
'       Call mVSObj.vsColor
'    End If
'End Sub
'******************************************'
'*****       移動選択項目選択          *****
'******************************************'
Private Sub VSGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'' Button : 1 = 左ﾎﾞﾀﾝ ,2 = 右ﾎﾞﾀﾝ
'    With VsGrid1
'        .Tag = ""
'        If Button = 2 Then
'            If .MouseRow = 0 Then
'                .Col = 1: .ColSel = .Cols - 1
'                Exit Sub
'            End If
'            .Col = 1: .ColSel = .Cols - 1
'            Exit Sub
'        End If
'        If Button <> 1 Then Exit Sub
'    ' 固定行以外は行わない
'        If .MouseRow <> 0 Then Exit Sub
'    ' 固定列は行わない
'        If .MouseCol = 0 Then Exit Sub
'    ' 移動元列をタグへ退避
'        .Tag = Str(.MouseCol)
'        .Drag 1
'    End With
End Sub
'******************************************'
'*****        表 示 項 目 移 動        *****
'******************************************'
Private Sub VSGrid1_DragDrop(Source As Control, X As Single, Y As Single)
'    With VsGrid1
'    ' 固定列は行わない
'        If .MouseCol = 0 Then Exit Sub
'        If .Tag = "" Then Exit Sub
'' ｸﾞﾘｯﾄﾞ表示 禁止
'        .Redraw = False
'' 列移動
'        .ColPosition(Val2(.Tag)) = .MouseCol
'        .Col = 1: .ColSel = .Cols - 1
'' ｸﾞﾘｯﾄﾞ表示 解除
'        .Redraw = True
''
'        Call mVSObj.VSSort(, False)
'    End With
End Sub
'+--------------------------------+
'+   テキストファイル出力
'+--------------------------------+
Private Sub TextPut()
    If Not mRstFlg Then Exit Sub
    
    db_BackDir = RTrim(iH1_Dir)
    If Right(db_BackDir, 1) <> "\" Then db_BackDir = db_BackDir & "\"
    
    If Not db_Rireki("OutPut") Then Exit Sub

    With VsGrid1
        For gCnt1 = 1 To .Rows - 1
            If Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(1))) = "" Then Exit For
            If Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(8))) <> "S" Then
                db_LogData = Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(9)))                         ' 日付
                db_LogData = db_LogData & "," & Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(1)))      ' 開始日時
                db_LogData = db_LogData & "," & Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(10)))     ' 終了日時
                db_LogData = db_LogData & "," & Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(11)))     ' 時間
                db_LogData = db_LogData & "," & Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(8)))      ' 処理CD
                
                gStr = Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(6))) & _
                       Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(2)))
                db_LogData = db_LogData & "," & gStr      ' 退避ﾌｧｲﾙ名(ﾌﾙﾊﾟｽ)
                
                db_LogData = db_LogData & "," & Trim(.TextMatrix(gCnt1, CisVsGrid.FixedGet(5)))      ' 復元日時
                
                Call db_Rireki("Put", db_LogData)
            End If
        Next gCnt1
    End With
    
    Call db_Rireki("Close")

End Sub

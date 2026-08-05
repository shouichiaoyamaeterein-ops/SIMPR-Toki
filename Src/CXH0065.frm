VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CXH0065 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "発注予定一覧表"
   ClientHeight    =   3720
   ClientLeft      =   5220
   ClientTop       =   3450
   ClientWidth     =   5295
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5295
   Begin VB.TextBox Dummy 
      Height          =   270
      Left            =   30
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   2340
      Width           =   195
   End
   Begin VB.CommandButton PB_CAN 
      Caption         =   "取 消"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3045
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3165
      Width           =   1095
   End
   Begin VB.CommandButton PB_Ent 
      Caption         =   "開　始"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐ明朝"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1080
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3165
      Width           =   1095
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   3075
      Left            =   0
      ScaleHeight     =   3138.064
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   5205
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   5265
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   2925
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   5159
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin VB.CheckBox ChkH1_UkeUmu 
            BackColor       =   &H00C0C0C0&
            Caption         =   "無"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1620
            TabIndex        =   1
            Top             =   1365
            Width           =   810
         End
         Begin CisYMD_With_Btn.CisYMDwB iH1_HYmd 
            Height          =   375
            Left            =   1515
            TabIndex        =   0
            Top             =   855
            Width           =   2130
            _ExtentX        =   3413
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   14.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Object.Width           =   2130
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   315
            Left            =   645
            Top             =   900
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   556
            Caption         =   "発 注 日"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin VB.Frame Back_Out 
            BackColor       =   &H00C0C0C0&
            Caption         =   "　出　力　先"
            Height          =   600
            Left            =   45
            TabIndex        =   9
            Top             =   1755
            Width           =   4965
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "画面(ﾌﾟﾚﾋﾞｭｰ)"
               Height          =   240
               Index           =   1
               Left            =   2610
               TabIndex        =   3
               Top             =   270
               Width           =   1485
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0FFC0&
               Caption         =   "プリンター"
               Height          =   240
               Index           =   0
               Left            =   1020
               TabIndex        =   2
               Top             =   270
               Value           =   -1  'True
               Width           =   1110
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D28 
            Height          =   555
            Index           =   1
            Left            =   60
            Top             =   60
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   979
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderStyle     =   1
            Begin VB.PictureBox B1_Area 
               BackColor       =   &H00FFC0FF&
               Height          =   495
               Left            =   30
               ScaleHeight     =   435
               ScaleWidth      =   4845
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   30
               Width           =   4905
               Begin Cis3D_v60.CIS3D lb_msg 
                  Height          =   405
                  Left            =   60
                  Top             =   30
                  Width           =   4725
                  _ExtentX        =   8334
                  _ExtentY        =   714
                  BackColor       =   16761087
                  Caption         =   "発注予定一覧表"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cBoderColor1    =   16761087
                  cBoderColor2    =   16761087
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   7
                  cBoderStyle     =   2
               End
            End
         End
         Begin Cis3D_v60.CIS3D lb_PrtDev 
            Height          =   345
            Left            =   75
            Top             =   2505
            Width           =   4590
            _ExtentX        =   8096
            _ExtentY        =   609
            BackColor       =   16777152
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
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
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D PB_Print 
            Height          =   360
            Left            =   4680
            Top             =   2505
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   635
            Caption         =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   2
            cAlingnment     =   7
            cButton         =   -1  'True
            cPositionY      =   -60
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   315
            Left            =   660
            Top             =   1335
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            Caption         =   "受入印刷"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4740
      Top             =   3195
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "CXH0065"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  発注予定一覧表　印刷指示
'**       フォームID    :  CXH0065
'**       処理概要      :
'**
'**       作  成  日    :  2005/04/26  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ

    Dim CisFrm          As New CISFormContorl


'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Set CisFrm = New CISFormContorl

    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
'
    Op_Out(0).Value = True
'
    Dummy.Left = -1000
    Dummy.Enabled = True
'
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(False, x1024y768)

End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    lb_msg = "発注予定一覧表"
    lb_PrtDev = Printer.DeviceName
    iH1_HYmd = CXH0065_HYmd
    iH1_HYmd.SetFocus
'
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
'
    Key_Acc = False
    Select Case Key_Code
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyReturn
                               If PB_Ent.Visible Then                      '【入力】
                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
                                  Dummy.Enabled = False
                               End If
          Case Else:          Exit Function
    End Select
    Key_Acc = True
End Function
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Call ReturnPress
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    CXH0065_PrintOut = 9
    Unload CXH0065
End Sub
Private Sub ChkH1_UkeUmu_Click()
    If ChkH1_UkeUmu.Value <> 0 Then
        ChkH1_UkeUmu.Caption = "有"
    Else
        ChkH1_UkeUmu.Caption = "無"
    End If
End Sub
'+----------------------------+
'+    出力先
'+----------------------------+
Private Sub Op_Out_GotFocus(Index As Integer)
    Op_Out(0).BackColor = &HC0C0C0
    Op_Out(1).BackColor = &HC0C0C0

    Op_Out(Index).BackColor = &HC0FFC0
End Sub
Private Sub Op_Out_Click(Index As Integer)
    lb_PrtDev.Caption = Printer.DeviceName
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    If Not Head1Chk Then Exit Sub
    PB_Ent.Visible = False
    PB_CAN.Visible = False
    PB_Print.Visible = False
    Back_Out.Visible = False
'   発注日
    CXH0065_HYmd = RTrim(iH1_HYmd)
    CXH0065_UkeirUmu = ChkH1_UkeUmu.Value
'   出力先
    If Op_Out(0).Value Then
        CXH0065_PrintOut = 0
    Else
        CXH0065_PrintOut = 1
    End If
    RV_Rtn = True
    DoEvents
    Unload CXH0065
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
'発注日
    If RTrim(iH1_HYmd) = "" Then
        CisFun.MB_Lines = 4
        CisFun.MB_MSG(1) = "     　発注日　必須！     "
        CisFun.MB_MSG(3) = "     発注日を入力して下さい。     "
        CisFun.MB_Button = Error
        CisFun.MBOX
        iH1_HYmd.SetFocus
        Exit Function
    End If
    If Not iH1_HYmd.cChkResult Then
        CisFun.MB_Lines = 4
        CisFun.MB_MSG(1) = "     　発注日　異常！     "
        CisFun.MB_MSG(3) = "     発注日を再入力して下さい。     "
        CisFun.MB_Button = Error
        CisFun.MBOX
        iH1_HYmd.SetFocus
        Exit Function
    End If
    
    Head1Chk = True
End Function
'+-------------------------------+
'+     出力先 変更
'+-------------------------------+
Private Sub PB_Print_Click()
    On Error Resume Next
    CommonDialog1.Flags = 64 + cdlPDHidePrintToFile
    CommonDialog1.Max = 0
    CommonDialog1.Min = 0
    CommonDialog1.Copies = 1
    CommonDialog1.ToPage = 1
    CommonDialog1.FromPage = 1
    ' 以下のｿｰｽｺｰﾄﾞは、<ｷｬﾝｾﾙ>ﾎﾞﾀﾝがｸﾘｯｸされたときに実行されます。
    CommonDialog1.CancelError = True
    CommonDialog1.ShowPrinter
    If Err = 32755 Then
        '<ｷｬﾝｾﾙ>ﾎﾞﾀﾝが押されたときにｴﾗｰｺｰﾄﾞを返すようにする設定です。
        '必要なｿｰｽｺｰﾄﾞを入力してください。
    End If
    lb_PrtDev.Caption = Printer.DeviceName
    On Error GoTo 0
End Sub


VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CKN0110 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "発注サイクル一覧表"
   ClientHeight    =   3960
   ClientLeft      =   255
   ClientTop       =   1620
   ClientWidth     =   5295
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5295
   Begin VB.TextBox Dummy 
      Height          =   270
      Left            =   30
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3330
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3390
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3390
      Width           =   1095
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   3285
      Left            =   0
      ScaleHeight     =   3356.636
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   5205
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   5265
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   3075
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   5424
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
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            Height          =   600
            Left            =   1260
            TabIndex        =   8
            Top             =   960
            Width           =   3480
            Begin VB.OptionButton Op_PrintKB 
               BackColor       =   &H00C0C0C0&
               Caption         =   "発行時間"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   1800
               TabIndex        =   10
               Top             =   240
               Width           =   1350
            End
            Begin VB.OptionButton Op_PrintKB 
               BackColor       =   &H00C0FFC0&
               Caption         =   "仕入先"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   210
               TabIndex        =   9
               Top             =   240
               Value           =   -1  'True
               Width           =   1140
            End
         End
         Begin VB.Frame Back_Out 
            BackColor       =   &H00C0C0C0&
            Caption         =   "　出　力　先"
            Height          =   600
            Left            =   60
            TabIndex        =   5
            Top             =   1860
            Width           =   4965
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0C0C0&
               Caption         =   "画面(ﾌﾟﾚﾋﾞｭｰ)"
               Height          =   240
               Index           =   1
               Left            =   3000
               TabIndex        =   7
               Top             =   270
               Width           =   1485
            End
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0FFC0&
               Caption         =   "プリンター"
               Height          =   240
               Index           =   0
               Left            =   1410
               TabIndex        =   6
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
               TabIndex        =   2
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
                  Caption         =   "発注サイクル一覧表"
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
            Left            =   90
            Top             =   2610
            Width           =   4635
            _ExtentX        =   8176
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
            Left            =   4740
            Top             =   2610
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
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   360
            Index           =   0
            Left            =   240
            Top             =   1170
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   635
            ForeColor       =   16711680
            Caption         =   "出力順"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4740
      Top             =   3420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "CKN0110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  発注サイクル一覧表　印刷指示
'**       フォームID    :  CKN0110
'**       処理概要      :
'**
'**       作  成  日    :  2004/05/11  By CIS
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
   'Back_Out.Visible = True
    Op_Out(0).Value = True
    Op_PrintKB(1).Value = True
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
    lb_msg = "発注サイクル一覧表"
    lb_PrtDev = Printer.DeviceName
    Op_PrintKB(1).SetFocus
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
    Unload CKN0110
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    CKN0110_PrintOut = 9
    Unload CKN0110
End Sub
'+----------------------------+
'+    出力順
'+----------------------------+
Private Sub Op_PrintKB_GotFocus(Index As Integer)
    Op_PrintKB(1).BackColor = &HC0C0C0
    Op_PrintKB(2).BackColor = &HC0C0C0

    Op_PrintKB(Index).BackColor = &HC0FFC0
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
    PB_Ent.Visible = False
    PB_CAN.Visible = False
    PB_Print.Visible = False
    Back_Out.Visible = False

'   出力順
    If Op_PrintKB(1).Value Then
        CKN0110_PrintKB = 1
    Else
        CKN0110_PrintKB = 2
    End If
'   出力先
    If Op_Out(0).Value Then
        CKN0110_PrintOut = 0
    Else
        CKN0110_PrintOut = 1
    End If
    DoEvents
End Sub
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


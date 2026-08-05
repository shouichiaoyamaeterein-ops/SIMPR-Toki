VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CTP3115 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "eかんばん変換リスト発行"
   ClientHeight    =   3300
   ClientLeft      =   8205
   ClientTop       =   1995
   ClientWidth     =   5325
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5325
   Begin VB.TextBox Dummy 
      Height          =   270
      Left            =   60
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2700
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
      Left            =   3075
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2730
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
      Left            =   1110
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2730
      Width           =   1095
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   2595
      Left            =   0
      ScaleHeight     =   2638.472
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   5205
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   5265
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   2445
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   4313
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
         Begin VB.Frame Back_Prt 
            Caption         =   " 印 刷 内 容 "
            Height          =   765
            Left            =   60
            TabIndex        =   11
            Top             =   630
            Width           =   4995
            Begin VB.CheckBox H1Ch_Dbl 
               Caption         =   "重複"
               Height          =   345
               Left            =   3060
               TabIndex        =   12
               Top             =   270
               Width           =   735
            End
            Begin VB.CheckBox H1Ch_Err 
               Caption         =   "エラー内容"
               Height          =   345
               Left            =   240
               TabIndex        =   0
               Top             =   270
               Width           =   1275
            End
            Begin VB.CheckBox H1Ch_Kari 
               Caption         =   "仮登録内容"
               Height          =   345
               Left            =   1560
               TabIndex        =   1
               Top             =   270
               Width           =   1275
            End
            Begin VB.CheckBox H1Ch_Etc 
               Caption         =   "その他"
               Height          =   345
               Left            =   3990
               TabIndex        =   2
               Top             =   270
               Width           =   855
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "　出　力　先"
            Height          =   600
            Left            =   60
            TabIndex        =   10
            Top             =   1410
            Width           =   4995
            Begin VB.OptionButton Op_Out 
               BackColor       =   &H00C0FFC0&
               Caption         =   "プリンター"
               Height          =   240
               Index           =   0
               Left            =   870
               TabIndex        =   3
               Top             =   270
               Value           =   -1  'True
               Width           =   1110
            End
            Begin VB.OptionButton Op_Out 
               Caption         =   "画面(ﾌﾟﾚﾋﾞｭｰ)"
               Height          =   240
               Index           =   1
               Left            =   2850
               TabIndex        =   4
               Top             =   240
               Width           =   1485
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
               TabIndex        =   9
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
                  Caption         =   "データ抽出 処理中"
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
            Left            =   60
            Top             =   2040
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
            Height          =   345
            Left            =   4710
            Top             =   2040
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   609
            Caption         =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4650
      Top             =   2670
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "CTP3115"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  eかんばん変換リスト発行
'**       フォームID    :  CTP0115
'**       処理概要      :
'**
'**       作  成  日    :  2005/11/28  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim mCsvDir         As String       ' CSV出力先
    
    Dim CisFrm          As CISFormContorl
    
    Dim mBeforePrint    As String
    Dim mAfterPrint     As String
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   #  項 目 初 期 化   #
'   #-------------------#
    Set CisFrm = New CISFormContorl
    
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
    
    Dummy.Left = -1000
    Dummy.Enabled = True
    
    lb_msg.Caption = "eかんばん受注変換リスト発行"
    If eKanban_Proc = 1 Then
       lb_msg.Caption = "eかんばん出荷変換リスト発行"
    End If
'   #-------------------#
'   # 画面中央位置表示  #
'   #-------------------#
    Call CisFrm.Frm_Center(False, x1024y768)

End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

'ﾌﾟﾘﾝﾀ出力先
    mBeforePrint = Printer.DeviceName
    lb_PrtDev.Caption = mBeforePrint
'
    If eKanban_CvtList(eKanban_Proc) Then
        If eKanban_CvtErr(eKanban_Proc) = 0 Then H1Ch_Err.Value = 1
        If eKanban_CvtKari(eKanban_Proc) = 0 Then H1Ch_Kari.Value = 1
        If eKanban_CvtEtc(eKanban_Proc) = 0 Then H1Ch_Etc.Value = 1
        If eKanban_CvtDbl(eKanban_Proc) = 0 Then H1Ch_Dbl.Value = 1
    Else
        H1Ch_Err.Value = 1
        H1Ch_Kari.Value = 1
        H1Ch_Etc.Value = 1
        H1Ch_Dbl.Value = 1
    End If
    
    Op_Out(0).Value = True
        
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
'+      出　力　先
'+----------------------------+
Private Sub Op_Out_GotFocus(Index As Integer)
    Op_Out(0).BackColor = &H8000000F
    Op_Out(1).BackColor = &H8000000F
    
    Op_Out(Index).BackColor = &HC0FFC0
    
End Sub
Private Sub Op_Out_Click(Index As Integer)
    lb_PrtDev.Caption = Printer.DeviceName
'    gp_PrtKB = Index
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)    +
'+----------------------------+
Private Sub PB_ENT_Click()
    Call ReturnPress
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    pref_Rtn = False
    Call End_Proc
End Sub
Private Sub End_Proc()
        
    mAfterPrint = Printer.DeviceName
    If mBeforePrint <> mAfterPrint Then
       With CisFrm
            .PrintSet (mBeforePrint)
       End With
    End If
    
    Me.Hide
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    eKanban_CvtErr(eKanban_Proc) = 0
    eKanban_CvtKari(eKanban_Proc) = 0
    eKanban_CvtEtc(eKanban_Proc) = 0
    eKanban_CvtDbl(eKanban_Proc) = 0
    
    gCnt1 = H1Ch_Err.Value + H1Ch_Kari.Value + H1Ch_Etc.Value + H1Ch_Dbl.Value
    
    If gCnt1 = 0 Or gCnt1 = 4 Then
    Else
        If H1Ch_Err.Value = 0 Then eKanban_CvtErr(eKanban_Proc) = 1
        If H1Ch_Kari.Value = 0 Then eKanban_CvtKari(eKanban_Proc) = 1
        If H1Ch_Etc.Value = 0 Then eKanban_CvtEtc(eKanban_Proc) = 1
        If H1Ch_Dbl.Value = 0 Then eKanban_CvtDbl(eKanban_Proc) = 1
    End If
        
        
    Call eKanban_Cvt_List
    
    If EHTRead(gSL_Select) Then
            
       lb_msg = "【 eかんばん 】 変換リスト 印刷中": DoEvents
       With eKanban_Prt
            If eKanban_Proc = 0 Then
               '------------------------------------------- eかんばん受注変換リスト発行
               Set .iPrtForm = CTP3110L
               .dpMSG = "eかんばん受注変換リスト　印刷中"
            Else
               '------------------------------------------- eかんばん出荷変換リスト発行
               Set .iPrtForm = CTP3210L
               .dpMSG = "eかんばん出荷変換リスト　印刷中"
            End If
            If Op_Out(0).Value Then
               .iPreview = pv_Print
            Else
               .iPreview = pv_Preview
            End If
            .dpMSGScreen = モーダル
            
            .PrintStart
        End With
        Call EHTClose
    End If
    
    Call End_Proc
    
ReturnPress_Ed:
End Sub
'+-------------------------------+
'+     出力先 変更               +
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
    If ERR = 32755 Then
        '<ｷｬﾝｾﾙ>ﾎﾞﾀﾝが押されたときにｴﾗｰｺｰﾄﾞを返すようにする設定です。
        '必要なｿｰｽｺｰﾄﾞを入力してください。
    End If
    lb_PrtDev.Caption = Printer.DeviceName
    On Error GoTo 0
End Sub


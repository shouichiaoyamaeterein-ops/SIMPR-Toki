VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Begin VB.Form CXM0041 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "カレンダ自動作成"
   ClientHeight    =   6975
   ClientLeft      =   6480
   ClientTop       =   7080
   ClientWidth     =   5520
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   5520
   Begin Cis3D_v60.CIS3D B1_Area1 
      Height          =   4680
      Left            =   -4875
      Top             =   2340
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   8255
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   630
         Index           =   0
         Left            =   750
         Top             =   1725
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   1111
         ForeColor       =   65535
         Caption         =   "作成件数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   12632256
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D CIS3D10 
         Height          =   780
         Left            =   750
         Top             =   525
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   1376
         ForeColor       =   128
         Caption         =   "カレンダデータ処理件数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   5
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   630
         Index           =   1
         Left            =   750
         Top             =   2550
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   1111
         ForeColor       =   12582912
         Caption         =   "更新件数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   630
         Index           =   2
         Left            =   750
         Top             =   3375
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   1111
         ForeColor       =   255
         Caption         =   "削除件数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D B1lb_ADDCnt 
         Height          =   630
         Left            =   2610
         Top             =   1725
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   1111
         ForeColor       =   65535
         Caption         =   "999,9 件"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   12632256
         cFont3DColor2   =   8421504
         cAlingnment     =   4
         cPositionX      =   -30
      End
      Begin Cis3D_v60.CIS3D B1lb_UPDCnt 
         Height          =   630
         Left            =   2625
         Top             =   2550
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   1111
         ForeColor       =   12582912
         Caption         =   "999,9 件"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   4
         cPositionX      =   -30
      End
      Begin Cis3D_v60.CIS3D B1lb_DELCnt 
         Height          =   630
         Left            =   2610
         Top             =   3375
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   1111
         ForeColor       =   255
         Caption         =   "999,9 件"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   4
         cPositionX      =   -30
      End
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   555
      Left            =   3240
      Top             =   6135
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   979
      ForeColor       =   16711680
      Caption         =   "処理終了"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      cAlingnment     =   7
      cButton         =   -1  'True
   End
   Begin Cis3D_v60.CIS3D Lb_END 
      Height          =   195
      Left            =   3240
      Top             =   6690
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   344
      BackColor       =   8421376
      ForeColor       =   16777215
      Caption         =   "( F12 )"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   8421376
      cBoderColor2    =   8421376
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cFont3D         =   2
      cAlingnment     =   7
      cBoderStyle     =   2
   End
   Begin Cis3D_v60.CIS3D PB_ENT 
      Height          =   555
      Left            =   810
      Top             =   6135
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   979
      ForeColor       =   128
      Caption         =   "作成開始"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      cAlingnment     =   7
      cButton         =   -1  'True
   End
   Begin Cis3D_v60.CIS3D Lb_ENT 
      Height          =   195
      Left            =   810
      Top             =   6690
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   344
      BackColor       =   8421376
      ForeColor       =   16777215
      Caption         =   "( E n t e r )"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   8421376
      cBoderColor2    =   8421376
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cFont3D         =   2
      cAlingnment     =   7
      cBoderStyle     =   2
   End
   Begin Cis3D_v60.CIS3D LB_Msg 
      Height          =   585
      Left            =   165
      Top             =   540
      Visible         =   0   'False
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   1032
      BackColor       =   0
      ForeColor       =   65535
      Caption         =   "カレンダーマスタ登録中"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cAlingnment     =   7
   End
   Begin VB.PictureBox H1_Area1 
      Height          =   4740
      Left            =   150
      ScaleHeight     =   4680
      ScaleWidth      =   5115
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1260
      Width           =   5175
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   825
         Left            =   60
         Top             =   3795
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   1455
         ForeColor       =   16711680
         Caption         =   "データ削除"
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
         cAlingnment     =   6
         cPositionY      =   90
         Begin Cis3D_v60.CIS3D B1lb_DelYM 
            Height          =   360
            Left            =   45
            Top             =   420
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   635
            BackColor       =   16777152
            ForeColor       =   192
            Caption         =   "(           以前のデータを削除します )"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D8 
            Height          =   285
            Left            =   3990
            Top             =   120
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   503
            ForeColor       =   16711680
            Caption         =   "ヶ月以前"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
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
            cBoderStyle     =   2
         End
         Begin CisText_V60.CisText iH1_Del 
            Height          =   360
            Left            =   3555
            TabIndex        =   14
            Top             =   60
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
            cFormat         =   "##"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "99"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   1500
         Left            =   60
         Top             =   2280
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   2646
         ForeColor       =   16711680
         Caption         =   "参照取引先指定"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin VB.PictureBox Copy_PNLG 
            BackColor       =   &H00C0E0FF&
            Height          =   1125
            Left            =   45
            ScaleHeight     =   1065
            ScaleWidth      =   735
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   330
            Width           =   795
            Begin VB.OptionButton H1Op_Copy 
               BackColor       =   &H00C0E0FF&
               Caption         =   "有"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   90
               TabIndex        =   9
               Top             =   645
               Width           =   555
            End
            Begin VB.OptionButton H1Op_Copy 
               BackColor       =   &H00C0E0FF&
               Caption         =   "無"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   90
               TabIndex        =   8
               Top             =   255
               Value           =   -1  'True
               Width           =   555
            End
         End
         Begin VB.PictureBox Picture6 
            Height          =   1125
            Left            =   840
            ScaleHeight     =   1065
            ScaleWidth      =   4020
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   330
            Width           =   4080
            Begin Cis3D_v60.CIS3D CIS3D6 
               Height          =   360
               Left            =   3180
               Top             =   30
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   635
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
            End
            Begin VB.PictureBox PNL_Back 
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   0  'なし
               Height          =   735
               Left            =   30
               ScaleHeight     =   735
               ScaleWidth      =   3165
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   30
               Width           =   3165
               Begin CisText_V60.CisText iH1_Ukeire 
                  Height          =   375
                  Left            =   2775
                  TabIndex        =   13
                  Top             =   360
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XX"
                  MaxLength       =   2
               End
               Begin CisText_V60.CisText iH1_Torcd 
                  Height          =   375
                  Left            =   0
                  TabIndex        =   10
                  Top             =   360
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXXXXX"
                  MaxLength       =   7
                  cAutoEject      =   0   'False
               End
               Begin Cis3D_v60.CIS3D Torkb_PNLG 
                  Height          =   360
                  Left            =   1005
                  Top             =   360
                  Width           =   1755
                  _ExtentX        =   3096
                  _ExtentY        =   635
                  BackColor       =   12640511
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
                  cBoderColor1    =   8421504
                  cBoderColor2    =   -2147483643
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cBoderStyle     =   1
                  Begin VB.OptionButton H1Op_Torkb 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "受注"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Index           =   0
                     Left            =   60
                     TabIndex        =   11
                     Top             =   75
                     Value           =   -1  'True
                     Width           =   825
                  End
                  Begin VB.OptionButton H1Op_Torkb 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "発注"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Index           =   1
                     Left            =   900
                     TabIndex        =   12
                     Top             =   75
                     Width           =   825
                  End
               End
               Begin Cis3D_v60.CIS3D UC_3D2 
                  Height          =   360
                  Index           =   0
                  Left            =   0
                  Top             =   0
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  ForeColor       =   16711680
                  Caption         =   "取引先"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   12
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
               Begin Cis3D_v60.CIS3D UC_3D2 
                  Height          =   360
                  Index           =   1
                  Left            =   990
                  Top             =   0
                  Width           =   1770
                  _ExtentX        =   3122
                  _ExtentY        =   635
                  ForeColor       =   16711680
                  Caption         =   "取引先区分"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   12
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
               Begin Cis3D_v60.CIS3D UC_3D2 
                  Height          =   360
                  Index           =   2
                  Left            =   2760
                  Top             =   0
                  Width           =   390
                  _ExtentX        =   688
                  _ExtentY        =   635
                  ForeColor       =   16711680
                  Caption         =   "受"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   12
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
            Begin Cis3D_v60.CIS3D PNL_Look 
               Height          =   675
               Left            =   3180
               Top             =   390
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   1191
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               Begin Cis3D_v60.CIS3D PB_Look 
                  Height          =   405
                  Left            =   30
                  Top             =   45
                  Width           =   765
                  _ExtentX        =   1349
                  _ExtentY        =   714
                  ForeColor       =   16576
                  Caption         =   "検索"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   3
                  cAlingnment     =   7
                  cButton         =   -1  'True
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  '透明
                  Caption         =   "( F7 )"
                  Height          =   180
                  Left            =   195
                  TabIndex        =   22
                  Top             =   480
                  Width           =   435
               End
            End
            Begin Cis3D_v60.CIS3D H1lb_Tornm 
               Height          =   300
               Left            =   30
               Top             =   750
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   529
               BackColor       =   16777152
               ForeColor       =   128
               Caption         =   "XXXXXXXXX1XXXXXXXXX2"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   12
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
               cFont3D         =   2
               cAlingnment     =   1
               cBoderStyle     =   1
            End
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   735
         Left            =   75
         Top             =   1545
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "データ上書き設定"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin VB.PictureBox Add_PNLG 
            BackColor       =   &H00C0E0FF&
            Height          =   360
            Left            =   45
            ScaleHeight     =   300
            ScaleWidth      =   4815
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   330
            Width           =   4875
            Begin VB.OptionButton H1Op_Add 
               BackColor       =   &H00C0E0FF&
               Caption         =   "しない"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   255
               TabIndex        =   6
               Top             =   45
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton H1Op_Add 
               BackColor       =   &H00C0E0FF&
               Caption         =   "する"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   1
               Left            =   2565
               TabIndex        =   7
               Top             =   45
               Width           =   795
            End
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   735
         Left            =   75
         Top             =   810
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "作成基準"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin VB.PictureBox Skbn_PNLG 
            BackColor       =   &H00C0E0FF&
            Height          =   360
            Left            =   45
            ScaleHeight     =   300
            ScaleWidth      =   4815
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   330
            Width           =   4875
            Begin VB.OptionButton H1Op_Skbn 
               BackColor       =   &H00C0E0FF&
               Caption         =   "取引先マスタ"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   2565
               TabIndex        =   5
               Top             =   45
               Width           =   1680
            End
            Begin VB.OptionButton H1Op_Skbn 
               BackColor       =   &H00C0E0FF&
               Caption         =   "カレンダマスタ"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   255
               TabIndex        =   4
               Top             =   45
               Value           =   -1  'True
               Width           =   1800
            End
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   735
         Left            =   3570
         Top             =   75
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "最終登録年"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionX      =   60
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D lb_LastYear 
            Height          =   345
            Left            =   105
            Top             =   330
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   609
            BackColor       =   16777152
            ForeColor       =   12583104
            Caption         =   "9999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   735
         Left            =   75
         Top             =   75
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "作成取引先区分"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin VB.PictureBox TrnKbn_PNLG 
            BackColor       =   &H00C0E0FF&
            Height          =   360
            Left            =   45
            ScaleHeight     =   300
            ScaleWidth      =   2250
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   315
            Width           =   2310
            Begin VB.OptionButton H1Op_TrnKbn 
               BackColor       =   &H00C0E0FF&
               Caption         =   "受注"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   255
               TabIndex        =   1
               Top             =   45
               Value           =   -1  'True
               Width           =   795
            End
            Begin VB.OptionButton H1Op_TrnKbn 
               BackColor       =   &H00C0E0FF&
               Caption         =   "発注"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   1260
               TabIndex        =   2
               Top             =   45
               Width           =   795
            End
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   735
         Left            =   2460
         Top             =   75
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1296
         ForeColor       =   16711680
         Caption         =   "作成年"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_YY 
            Height          =   360
            Left            =   255
            TabIndex        =   3
            Top             =   330
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   635
            cFormat         =   "####"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "9999"
            MaxLength       =   4
            cDataType       =   1
            cDataReplace    =   1
            cFbComma        =   0
            cGFormat        =   "####"
            cILength        =   4
         End
      End
   End
   Begin VB.TextBox Dummy 
      Height          =   285
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   195
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
      Caption         =   "【 カレンダマスタ自動作成  】"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
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
      cPositionX      =   300
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   4380
         Top             =   60
         Width           =   885
         _ExtentX        =   1561
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
         Height          =   285
         Left            =   8070
         Top             =   60
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   503
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
      End
   End
End
Attribute VB_Name = "CXM0041"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  カレンダ自動作成
'**       フォームID    :  CXM0041
'**       処理概要      :
'**
'**       作  成  日    :
'**       変  更  日    :   2005.02.14  By CIS  取引先区分初期表示パラメータ追加
'**                                             ( １桁目：0/受注、1/発注 )
'**       変  更  日    :   2005.02.28  By CIS  臨時昼/夜区分、追加
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim wObj            As Object       ' ｵﾌﾞｼﾞｪｸﾄﾜｰｸｴﾘｱ
    
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim SL_Select       As String       ' SQL SELECT
    Dim w_ReTime        As Integer      ' データ保持期間
    Dim cnt_ADD         As Long
    Dim cnt_UPD         As Long
    Dim cnt_DEL         As Long
    
    Dim w_CalenDay(12)  As String       ' 参照元カレンダワーク（昼勤）
    Dim w_CalenNig(12)  As String       ' 参照元カレンダワーク（夜勤）
    Dim w_CalenRDay(12) As String       ' 参照元カレンダワーク（臨時昼勤）      '2005.02.28 Add
    Dim w_CalenRNig(12) As String       ' 参照元カレンダワーク（臨時夜勤）      '2005.02.28 Add
    Dim w_CalenSuu(12)  As Byte         ' 参照元カレンダ日数（稼動）
    Dim w_CalenSuD(12)  As Byte         ' 参照元カレンダ日数（昼勤）
    Dim w_CalenSuN(12)  As Byte         ' 参照元カレンダ日数（夜勤）
    Dim w_CalenLas(12)  As Byte         ' 参照元カレンダ末日

    Dim Idx             As Integer      ' 添字
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
    
    iH1_YY = CisFun.Val2(Format(Date, "YYYY") + 1)
    PNL_Back.Enabled = False
    PNL_Look.Visible = False
    H1Op_Add(0).Value = True           '2000/12/21

'-------------------------------------------< 2005.02.14 MNT START >
'    H1Op_TrnKbn(0).Value = True
    gInt = CisFun.Val2(Mid$(Command, 1, 1))
    H1Op_TrnKbn(gInt).Value = True
'-------------------------------------------< 2005.02.14 MNT  END  >

    H1Op_Skbn(0).Value = True
    H1Op_Copy(0).Value = True
    Dummy.Left = -1000
    Dummy.Enabled = False
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)

'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(False)
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormAct = Not FormAct

'   < 最終登録年 >
    SL_Select = "Select MAX(年) 年 From カレンダマスタ"
    If CLMRead(SL_Select, 1) Then lb_LastYear = CLM.年

'   < データ保持期間 >
    Call RetentionTime_Check

    iH1_Del = Format(w_ReTime, "00")
    Call iH1_Del_LostFocus
    
    H1Op_Add(0).Value = True           '2000/12/21
    
    CisFrm.GrpFocGotColor = gIGotSel
    CisFrm.GrpFocLostColor = gILostSel
    DoEvents

    B1_Area1.Move H1_Area1.Left + 30, H1_Area1.Top + 30
    cnt_ADD = 0:    cnt_UPD = 0:    cnt_DEL = 0
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
        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click      '【検索】
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click        '【終了】
        Case vbKeyReturn
                            If PB_ENT.Visible Then                          '【入力】
'                               Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                               Call PB_ENT_Click
'                               Dummy.Enabled = False
                            End If
        Case Else:          Exit Function
    End Select
    Key_Acc = True

End Function
'+----------------------------+
'+    検索   (F7)             +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Torihiki_Look: Exit Sub
    If PB_Look.Tag = "iH1_Ukeire" Then Call Ukeire_Look: Exit Sub
End Sub
'+----------------------------+
'+    取引先マスタ検索         +
'+----------------------------+
Private Sub Torihiki_Look()

    RV_Left = 0:    RV_Top = 0
    Select Case CisFun.Val2(H1Op_Torkb(0).Tag)
        Case 0:     RV_TorKb = 7
        Case Else:  RV_TorKb = 3
    End Select
    
    CKK0025.Show vbModal

    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RTrim$(RV_TorcdK)
        Call iH1_Torcd_LostFocus
    End If
    
    iH1_Torcd.SetFocus
End Sub
'+----------------------------+
'+    取引先受入マスタ検索     +
'+----------------------------+
Private Sub Ukeire_Look()

    RV_Left = 0:    RV_Top = 0
    Select Case CisFun.Val2(H1Op_Torkb(0).Tag)
        Case 0:     RV_TorKb = 7
        Case Else:  RV_TorKb = 3
    End Select
    
    CKK0028.Show vbModal

    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RTrim$(RV_TorcdK)
        iH1_Ukeire = RTrim$(RV_Ukeir)
        Call iH1_Torcd_LostFocus
    End If
    
    iH1_Ukeire.SetFocus
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
End Sub
'+----------------------------+
'+    終  了  キ ー(F12)      +
'+----------------------------+
Private Sub PB_END_Click()
    Me.Hide
End Sub
'*---------------------------------*
'*   H1op_TrnKbn    作成取引先区分  *
'*---------------------------------*
Private Sub H1op_TrnKbn_Click(Index As Integer)
    H1Op_TrnKbn(0).Tag = Index
End Sub
Private Sub H1op_TrnKbn_GotFocus(Index As Integer)
    CisFrm.GrpFocus H1Op_TrnKbn
End Sub
Private Sub H1op_TrnKbn_LostFocus(Index As Integer)
    CisFrm.GrpFocus H1Op_TrnKbn, Lost
End Sub
'*---------------------------------*
'*   iH1_YY         作成年          *
'*---------------------------------*
Private Sub iH1_YY_LostFocus()
    If CisFun.Val2(iH1_YY) < 1000 Then
        iH1_YY = 2000 + CisFun.Val2(iH1_YY)
    End If
End Sub
'+----------------------------------+
'+   H1Op_Skbn      作成基準区分    +
'+----------------------------------+
Private Sub H1Op_Skbn_Click(Index As Integer)
    H1Op_Skbn(0).Tag = Index
End Sub
Private Sub H1Op_Skbn_GotFocus(Index As Integer)
    CisFrm.GrpFocus H1Op_Skbn
End Sub
Private Sub H1Op_Skbn_LostFocus(Index As Integer)
    CisFrm.GrpFocus H1Op_Skbn, Lost
End Sub
'+----------------------------------+
'+   H1Op_Add      データ上書区分
'+----------------------------------+
Private Sub H1Op_Add_Click(Index As Integer)
    H1Op_Add(0).Tag = Index
End Sub
Private Sub H1Op_Add_GotFocus(Index As Integer)
    CisFrm.GrpFocus H1Op_Add
End Sub
Private Sub H1Op_Add_LostFocus(Index As Integer)
    CisFrm.GrpFocus H1Op_Add, Lost
End Sub
'+---------------------------------+
'+   H1Op_Copy  参照取引先指定      +
'+---------------------------------+
Private Sub H1Op_Copy_Click(Index As Integer)
    H1Op_Copy(0).Tag = Index
        
    If Index = 0 Then
        PNL_Back.Enabled = False
        iH1_Torcd = "": H1lb_Tornm = ""
        H1Op_Torkb(0).Value = False
        H1Op_Torkb(1).Value = False
    Else
        PNL_Back.Enabled = True
        H1lb_Tornm = ""
        H1Op_Torkb(0).Value = True
    End If
End Sub
Private Sub H1Op_Copy_GotFocus(Index As Integer)
    CisFrm.GrpFocus H1Op_Copy
End Sub
Private Sub H1Op_Copy_LostFocus(Index As Integer)
    CisFrm.GrpFocus H1Op_Copy, Lost
End Sub
'+---------------------+
'+   H1Op_Torkb     参照元取引先区分
'+---------------------+
Private Sub H1Op_Torkb_Click(Index As Integer)
    H1Op_Torkb(0).Tag = Index
End Sub
Private Sub H1Op_Torkb_GotFocus(Index As Integer)
    CisFrm.GrpFocus H1Op_Torkb
End Sub
Private Sub H1Op_Torkb_LostFocus(Index As Integer)
    CisFrm.GrpFocus H1Op_Torkb, Lost
End Sub
'+---------------------+
'+   ih1_Torcd
'+---------------------+
Private Sub iH1_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PNL_Look.Visible = True
End Sub
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PNL_Look.Visible = False
    End If
    
    Call GetTorcd
End Sub
'+---------------------+
'+   ih1_Ukeire
'+---------------------+
Private Sub iH1_Ukeire_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PNL_Look.Visible = True
End Sub
Private Sub iH1_Ukeire_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PNL_Look.Visible = False
    End If
End Sub
'+----------------------------+
'+    削除指定                 +(           以前のデータを削除します )
'+----------------------------+
Private Sub iH1_Del_LostFocus()
    gInt = CisFun.Val2(iH1_Del)
    B1lb_DelYM = "": B1lb_DelYM.Tag = ""
    If gInt = 0 Then Exit Sub
    
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add Format(Now(), "YYYYMMDD"), M, gInt * -1
            
        gStr = "( " & Mid$(.Date_Result, 1, 4) & "年" & Mid$(.Date_Result, 5, 2) & "月"
        gStr = gStr & " 以前のデータを削除します )"
        B1lb_DelYM = gStr
        B1lb_DelYM.Tag = .Date_Result
    End With
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
    Dim H1Mode      As Boolean
    Dim H1Color     As Long
    Dim B1Mode      As Boolean
    Dim PRMode      As Boolean
    
    H1Mode = False:     H1Color = gPLostSel
    B1Mode = False
    PRMode = False

    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
       Case "H1":       H1Mode = True:  H1Color = gPGotSel
       Case "B1":       B1Mode = True
       Case "PR":       PRMode = True
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Visible = B1Mode
    LB_Msg.Visible = B1Mode Or PRMode

    PB_ENT.Visible = H1Mode
    PB_END.Visible = H1Mode
    Lb_ENT.Visible = H1Mode
    Lb_END.Visible = H1Mode
    
    If B1Mode Then
        With PB_END
            .Left = Me.Width \ 2 - .Width \ 2
            .Caption = "確認"
            .Visible = B1Mode
        
            Lb_END.Left = .Left
        End With
        Lb_END.Visible = B1Mode
    End If
    
    Dummy.Enabled = False
End Function
'******************************************************
'*  ＢＯＤＹチェック処理
'******************************************************
Private Function Body_Chk()
    Static Cal_Chk              As Boolean          ' 未登録カレンダマスタフラグ
    Static Cal_Mon              As String           ' 未登録カレンダマスタ月
    
    Body_Chk = False
    Cal_Chk = True
    Erase w_CalenDay: Erase w_CalenNig
    Erase w_CalenRDay: Erase w_CalenRNig                                    '2005.05.28 Add
    Erase w_CalenSuu: Erase w_CalenSuD: Erase w_CalenSuN
    Erase w_CalenLas
    
'< 作成年 >
    If Trim(iH1_YY) = "" Then iH1_YY = CisFun.Val2(Format(Date, "YYYY") + 1)
    
'< 参照指定 >
    If CisFun.Val2(H1Op_Copy(0).Tag) = 1 Then
'   < あり >
        If RTrim(iH1_Torcd) <> "" Then
'           < 参照元　取引先 >
            If Not GetTorcd Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(1) = "    取引先マスタ未登録            "
                    .MB_Button = Error
                    .MBOX
                End With
                iH1_Torcd.SetFocus
            End If

'           < 参照元　取引先受入 >
            If RTrim$(iH1_Ukeire) <> "" Then
                SL_Select = "Select 受入 From 取引先受入マスタ "
                SL_Select = SL_Select & " Where 取引先CD = '" & RTrim$(iH1_Torcd) & "'"
                SL_Select = SL_Select & " And   取引先区分 = " & CisFun.Val2(H1Op_Torkb(0).Tag)
                SL_Select = SL_Select & " And 受入 = '" & RTrim$(iH1_Ukeire) & "'"
    
                If Not TUMRead(SL_Select, 1) Then
                    With CisFun
                        .MB_Lines = 4
                        .MB_MSG(1) = "  取引先受入マスタ未登録            "
                        .MB_Button = Error
                        .MBOX
                    End With
                    iH1_Ukeire.SetFocus
                    Exit Function
                End If
            End If
        End If
        
'       < 参照元　カレンダ >
        Cal_Mon = ""
        For Idx = 1 To 12
            SL_Select = "Select * From カレンダマスタ"
            SL_Select = SL_Select & " Where 取引先区分 = " & CisFun.Val2(H1Op_Torkb(0).Tag)
            SL_Select = SL_Select & " And   取引先 = '" & RTrim$(iH1_Torcd) & "'"
            SL_Select = SL_Select & " And   年 = '" & RTrim$(iH1_YY) & "'"
            SL_Select = SL_Select & " And   月 = '" & Format(Idx, "00") & "'"
            SL_Select = SL_Select & " And   受入 = '" & RTrim(iH1_Ukeire) & "'"
            If Not CLMRead(SL_Select) Then
                Cal_Chk = False
                Cal_Mon = Cal_Mon & " " & Format(Idx, "00")
            Else
                w_CalenDay(Idx) = CLM.稼動区分
                w_CalenNig(Idx) = CLM.夜勤区分
                w_CalenSuu(Idx) = CLM.稼動日数
                w_CalenSuD(Idx) = CLM.昼勤日数
                w_CalenSuN(Idx) = CLM.夜勤日数
                w_CalenLas(Idx) = CLM.末日
                w_CalenRDay(Idx) = CLM.臨時昼区分               '2005.02.28 Add
                w_CalenRNig(Idx) = CLM.臨時夜区分               '2005.02.28 Add
            End If
        Next Idx
       
        If Not Cal_Chk Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(1) = " 参照元カレンダマスタ未登録  "
                .MB_MSG(3) = " 　　　  年 = " & RTrim$(iH1_YY)
                .MB_MSG(4) = " 未登録月 =" & Cal_Mon
                .MB_Button = Error
                .MBOX
            End With
            GoTo Body_Chk_ED
        End If
    Else
'   < なし >
'       ( 稼動日数・区分ワーク作成 )
        For Idx = 1 To 12
            
'           ( 末日 )
            With CisFun
                .Date_RsFormat = fDD
                .Date_End RTrim$(iH1_YY) & Format(Idx, "00")
                w_CalenLas(Idx) = .Date_Result
            End With
        
            gStr = "": gCnt1 = 0
            gStr1 = ""                                                          '2005.02.28 Add
            For gInt = 1 To w_CalenLas(Idx)
                gDate = RTrim$(iH1_YY) & "/" & Format(Idx, "00") & "/" & Format(gInt, "00")
                Select Case Format(gDate, "w")
                    Case 1, 7:      gStr = gStr & "1"
                    Case Else:      gStr = gStr & "0": gCnt1 = gCnt1 + 1
                End Select
                gStr1 = gStr1 & "0"                                         '2005.02.28 Add
            Next gInt
        
'           ( 区分 )
            w_CalenDay(Idx) = gStr:     w_CalenNig(Idx) = gStr
            w_CalenRDay(Idx) = gStr1:    w_CalenRNig(Idx) = gStr1               '2005.02.28 Add
'           ( 日数 )
            w_CalenSuu(Idx) = gCnt1:    w_CalenSuD(Idx) = gCnt1:    w_CalenSuN(Idx) = gCnt1
        Next Idx
    End If
    
'( 処理確認メッセージ )
    With CisFun
        .MB_Lines = 11
    
        .MB_MSG(1) = "    【 確認メッセージ 】      "
        .MB_MSG(2) = "            作成年 ＝  " & Trim(iH1_YY)
        
        .MB_MSG(3) = "         作成方法 ＝  " & _
                        H1Op_Skbn(CisFun.Val2(H1Op_Skbn(0).Tag)).Caption & "より"
    
        .MB_MSG(4) = "         作成区分 ＝  " & H1Op_TrnKbn(0).Tag
        .MB_MSG(5) = "         削除年月 ＝  " & Trim(B1lb_DelYM) & "以前"
    
        If H1Op_Copy(0).Value = True Then
            .MB_MSG(6) = "参照取引先指定 ＝  無"
        Else
            .MB_MSG(6) = "参照取引先指定 ＝  有"
            .MB_MSG(7) = "      参照取引先 ＝  " & RTrim(H1lb_Tornm)
        End If

        If H1Op_Add(0).Value = True Then
            If H1Op_Copy(0).Value = True Then
                .MB_MSG(7) = "     データ上書き ＝  無"
            Else
                .MB_MSG(8) = "     データ上書き ＝  無"
            End If
        Else
            If H1Op_Copy(0).Value = True Then
                .MB_MSG(7) = "     データ上書き ＝  有"
            Else
                .MB_MSG(8) = "     データ上書き ＝  有"
            End If
        End If
    
        .MB_MSG(10) = "以上の条件でカレンダ自動作成を行います。よろしいですか"
    
        .MB_Title = "確認"
        .MB_Button = OK
        Body_Chk = .MBOX
    End With
    
Body_Chk_ED:
    Call CLMClose
End Function
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)

    If Not Body_Chk Then GoTo ReturnPress_Ed
    
    ProcHB = "PR"
    Call DispChange(ProcHB)
    DoEvents
    
'   ( 作成元キー作成 )
    If CisFun.Val2(H1Op_Skbn(0).Tag) = 0 Then
        Call CLKeyWork                              ' カレンダマスタ ベース
    Else
        Call TRKeyWork                              ' 取引先マスタ ベース
    End If
    
'   ( カレンダ作成 )
    Call CalenCreate

'   ( カレンダ削除 )
    Call CalenDelete
    
    B1lb_ADDCnt = Format(cnt_ADD, "#,### 件")
    B1lb_UPDCnt = Format(cnt_UPD, "#,### 件")
    B1lb_DELCnt = Format(cnt_DEL, "#,### 件")
    
    ProcHB = "B1"
    Call DispChange(ProcHB)
    LB_Msg = "【 処 理 終 了 】"
    DoEvents

ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'+-----------------------------------------+
'+          カ レ ン ダ 作 成               +
'+-----------------------------------------+
Private Function CalenCreate()
    CalenCreate = False

    On Error GoTo CalenCreate_Err
    CisDB.DBTran TransBegin

'< 更新データ(上書き) >
    If CisFun.Val2(H1Op_Add(0).Tag) = 1 Then
        For Idx = 1 To 12
    
            SL_Select = "UpDate カレンダマスタ Set" & vbCrLf
            SL_Select = SL_Select & " 末日 = " & w_CalenLas(Idx) & "," & vbCrLf
            SL_Select = SL_Select & " 稼動日数 = " & w_CalenSuu(Idx) & "," & vbCrLf
            SL_Select = SL_Select & " 昼勤日数 = " & w_CalenSuD(Idx) & "," & vbCrLf
            SL_Select = SL_Select & " 夜勤日数 = " & w_CalenSuN(Idx) & "," & vbCrLf
'2005.02.28 Upd Start
'            SL_Select = SL_Select & " 稼動区分 = " & w_CalenDay(Idx) & "," & vbCrLf
'            SL_Select = SL_Select & " 夜勤区分 = " & w_CalenNig(Idx) & "," & vbCrLf
            SL_Select = SL_Select & " 稼動区分 = '" & w_CalenDay(Idx) & "'," & vbCrLf
            SL_Select = SL_Select & " 夜勤区分 = '" & w_CalenNig(Idx) & "'," & vbCrLf
'2005.02.28 Upd End
            SL_Select = SL_Select & " 臨時昼区分 = '" & w_CalenRDay(Idx) & "'," & vbCrLf      '2005.02.28 Add
            SL_Select = SL_Select & " 臨時夜区分 = '" & w_CalenRNig(Idx) & "'," & vbCrLf      '2005.02.28 Add
            SL_Select = SL_Select & " 更新日   = GETDATE(),"
            SL_Select = SL_Select & " 更新端末 = SubString( Host_Name(), 1, 20 )"
            SL_Select = SL_Select & " FROM カレンダ作成キーワーク TR"
            SL_Select = SL_Select & " LEFT OUTER JOIN カレンダマスタ CL"
            SL_Select = SL_Select & " ON  CL.取引先区分 = TR.取引先区分"
            SL_Select = SL_Select & " AND CL.取引先 = TR.取引先"
            SL_Select = SL_Select & " AND CL.受入 = TR.受入"
            SL_Select = SL_Select & " AND CL.年 = '" & RTrim$(iH1_YY) & "'"
            SL_Select = SL_Select & " AND CL.月 = '" & Format(Idx, "00") & "'"
            SL_Select = SL_Select & " WHERE CL.取引先 IS NOT NULL"
            SL_Select = SL_Select & " AND   TR.作成端末 = SUBSTRING( HOST_NAME(), 1, 20 )"
        
            With CisDB
                .SQL = SL_Select
                .DBExec
    
                cnt_UPD = cnt_UPD + .RecordCount
            End With
        
        Next Idx
    End If
    
'< 新規データ作成 >
    For Idx = 1 To 12

        SL_Select = "Insert Into カレンダマスタ (" & vbCrLf
        SL_Select = SL_Select & " 取引先区分, 取引先, 年, 月, 受入, 末日," & vbCrLf
        SL_Select = SL_Select & " 稼動日数, 昼勤日数, 夜勤日数," & vbCrLf
        SL_Select = SL_Select & " 稼動区分, 夜勤区分," & vbCrLf
        SL_Select = SL_Select & " 臨時昼区分, 臨時夜区分," & vbCrLf                     '2005.02.28 Add
        SL_Select = SL_Select & " 作成日, 作成端末 )" & vbCrLf
        
        SL_Select = SL_Select & " Select TR.取引先区分, TR.取引先," & vbCrLf
        SL_Select = SL_Select & " '" & RTrim$(iH1_YY) & "'," & vbCrLf
        SL_Select = SL_Select & " '" & Format(Idx, "00") & "'," & vbCrLf
        SL_Select = SL_Select & " TR.受入," & vbCrLf
        SL_Select = SL_Select & "  " & w_CalenLas(Idx) & ", " & vbCrLf
        SL_Select = SL_Select & "  " & w_CalenSuu(Idx) & ", " & vbCrLf
        SL_Select = SL_Select & "  " & w_CalenSuD(Idx) & ", " & vbCrLf
        SL_Select = SL_Select & "  " & w_CalenSuN(Idx) & ", " & vbCrLf
        SL_Select = SL_Select & " '" & w_CalenDay(Idx) & "'," & vbCrLf
        SL_Select = SL_Select & " '" & w_CalenNig(Idx) & "'," & vbCrLf
        SL_Select = SL_Select & " '" & w_CalenRDay(Idx) & "'," & vbCrLf                 '2005.02.28 Add
        SL_Select = SL_Select & " '" & w_CalenRNig(Idx) & "'," & vbCrLf                 '2005.02.28 Add
        SL_Select = SL_Select & " GETDATE(), SUBSTRING( HOST_NAME(), 1, 20 )" & vbCrLf
        SL_Select = SL_Select & " FROM カレンダ作成キーワーク TR" & vbCrLf
        SL_Select = SL_Select & " LEFT OUTER JOIN カレンダマスタ CL" & vbCrLf
        SL_Select = SL_Select & " ON  CL.取引先区分 = TR.取引先区分" & vbCrLf
        SL_Select = SL_Select & " AND CL.取引先 = TR.取引先" & vbCrLf
        SL_Select = SL_Select & " AND CL.受入 = TR.受入" & vbCrLf
        SL_Select = SL_Select & " AND CL.年 = '" & RTrim$(iH1_YY) & "'" & vbCrLf
        SL_Select = SL_Select & " AND CL.月 = '" & Format(Idx, "00") & "'" & vbCrLf
        SL_Select = SL_Select & " WHERE CL.取引先 IS NULL" & vbCrLf
        SL_Select = SL_Select & " AND   TR.作成端末 = SUBSTRING( HOST_NAME(), 1, 20 )" & vbCrLf
    
        With CisDB
            .SQL = SL_Select
            .DBExec

            cnt_ADD = cnt_ADD + .RecordCount
        End With
    
    Next Idx

    CisDB.DBTran TransCommit
    On Error GoTo 0
    CalenCreate = True
    Exit Function
    
CalenCreate_Err:
    CisFun.ErrorBox
    End
End Function
'+---------------------------------------------+
'+    カレンダマスタ・カレンダ情報 削除          +
'+---------------------------------------------+
Private Function CalenDelete()
    CalenDelete = False
    
    If RTrim$(B1lb_DelYM.Tag) = "" Then GoTo CalenDelete_ED
    
    SL_Select = "Delete From カレンダマスタ "
    SL_Select = SL_Select & " Where 年 + 月 < " & "'" & RTrim$(B1lb_DelYM.Tag) & "'"
    SL_Select = SL_Select & " AND 取引先区分 = " & CisFun.Val2(H1Op_TrnKbn(0).Tag)
    
    On Error GoTo CalenDelete_ERR
    CisDB.DBTran TransBegin
    
    With CisDB
        .SQL = SL_Select
        .DBExec

        cnt_DEL = cnt_DEL + .RecordCount
    End With

    CisDB.DBTran TransCommit
    On Error GoTo 0

CalenDelete_ED:
    CalenDelete = True
    Exit Function

CalenDelete_ERR:
    CisFun.ErrorBox
    End
End Function
'*-----------------------------*
'*      取引先　略称参照        *
'*-----------------------------*
Private Function GetTorcd()
    GetTorcd = False
    
    If RTrim$(iH1_Torcd) = "" Then
        H1lb_Tornm = "☆　自社カレンダ　☆"
        GoTo GetTorcd_ED
    End If
    
    SL_Select = "Select 略称 From 取引先マスタ"
    SL_Select = SL_Select & " Where 取引先CD = '" & RTrim$(iH1_Torcd) & "'"
    SL_Select = SL_Select & " And   取引先区分 = " & CisFun.Val2(H1Op_Torkb(0).Tag)
    
    If Not TRMRead(SL_Select, 1) Then
        H1lb_Tornm = "※　未　登　録　※"
        Exit Function
    Else
        H1lb_Tornm = RTrim$(TRM.略称)
    End If

GetTorcd_ED:
    GetTorcd = True
End Function
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                     カ　レ　ン　ダ　作　成　用　ワ　ー　ク　テ　ー　ブ　ル　作　成
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'*--------------------------------------------------------------------------*
'*      取引先マスタベースの作成時の取引先区分、取引先CD、受入キーワーク        *
'*--------------------------------------------------------------------------*
Private Function TRKeyWork(Optional pKB As Byte = 0)

    SL_Select = "Delete カレンダ作成キーワーク"
    SL_Select = SL_Select & " Where 作成端末 = SubString( HOST_NAME(), 1, 20 )"
    With CisDB
        .SQL = SL_Select
        .DBExec
    End With

    If pKB <> 0 Then Exit Function

    SL_Select = "Insert Into カレンダ作成キーワーク" & vbCrLf
    
    If CisFun.Val2(H1Op_TrnKbn(0).Tag) = 0 Then
        SL_Select = SL_Select & " SELECT 0, '', '', GETDATE(), SUBSTRING( HOST_NAME(), 1, 20 )" & vbCrLf
        
        SL_Select = SL_Select & " UNION" & vbCrLf
        
        SL_Select = SL_Select & " SELECT 0, TR.取引先CD, '' 受入, GETDATE(), SUBSTRING( HOST_NAME(), 1, 20 )" & vbCrLf
        SL_Select = SL_Select & " FROM 取引先マスタ TR" & vbCrLf
        SL_Select = SL_Select & " WHERE TR.取引先区分 = 0" & vbCrLf
        
        SL_Select = SL_Select & " UNION" & vbCrLf
        
        SL_Select = SL_Select & " SELECT 0, TR.取引先CD, TU.受入, GETDATE(), SUBSTRING( HOST_NAME(), 1, 20 )" & vbCrLf
        SL_Select = SL_Select & " FROM 取引先マスタ TR" & vbCrLf
        SL_Select = SL_Select & " LEFT OUTER JOIN 取引先受入マスタ TU" & vbCrLf
        SL_Select = SL_Select & " ON  TU.取引先CD = TR.取引先CD" & vbCrLf
        SL_Select = SL_Select & " AND TU.取引先区分 = TR.取引先区分" & vbCrLf
        SL_Select = SL_Select & " WHERE TR.取引先区分 = 0" & vbCrLf
        SL_Select = SL_Select & " AND   TU.取引先CD IS NOT NULL" & vbCrLf
    Else
        SL_Select = SL_Select & " SELECT 1, '', '', GETDATE(), SUBSTRING( HOST_NAME(), 1, 20 )" & vbCrLf
        
        SL_Select = SL_Select & " UNION" & vbCrLf
        
        SL_Select = SL_Select & " SELECT 1 , TR.取引先CD, '' 受入, GETDATE(), SUBSTRING( HOST_NAME(), 1, 20 )" & vbCrLf
        SL_Select = SL_Select & " FROM 取引先マスタ TR" & vbCrLf
        SL_Select = SL_Select & " WHERE TR.取引先区分 = 1" & vbCrLf
        
        SL_Select = SL_Select & " UNION" & vbCrLf
        
        SL_Select = SL_Select & " SELECT 1, TR.取引先CD, TU.受入, GETDATE(), SUBSTRING( HOST_NAME(), 1, 20 )" & vbCrLf
        SL_Select = SL_Select & " FROM 取引先マスタ TR" & vbCrLf
        SL_Select = SL_Select & " LEFT OUTER JOIN 取引先受入マスタ TU" & vbCrLf
        SL_Select = SL_Select & " ON  TU.取引先CD = TR.取引先CD" & vbCrLf
        SL_Select = SL_Select & " AND TU.取引先区分 = TR.取引先区分" & vbCrLf
        SL_Select = SL_Select & " WHERE TR.取引先区分 = 1" & vbCrLf
        SL_Select = SL_Select & " AND   TU.取引先CD IS NOT NULL" & vbCrLf
    End If
    
    With CisDB
        .SQL = SL_Select
        .DBExec
    End With

End Function
'*--------------------------------------------------------------------------*
'*      カレンダマスタベースの作成時の取引先区分、取引先CD、受入キーワーク      *
'*--------------------------------------------------------------------------*
Private Function CLKeyWork(Optional pKB As Byte = 0)

    SL_Select = "Delete カレンダ作成キーワーク"
    SL_Select = SL_Select & " Where 作成端末 = SubString( HOST_NAME(), 1, 20 )"
    With CisDB
        .SQL = SL_Select
        .DBExec
    End With

    If pKB <> 0 Then Exit Function

    SL_Select = "Insert Into カレンダ作成キーワーク" & vbCrLf
    
    SL_Select = SL_Select & " SELECT DISTINCT"
    SL_Select = SL_Select & " 取引先区分, 取引先, 受入,"
    SL_Select = SL_Select & " GETDATE(), SUBSTRING( HOST_NAME(), 1, 20 )" & vbCrLf
    SL_Select = SL_Select & " FROM カレンダマスタ"
    SL_Select = SL_Select & " WHERE 取引先区分 = " & CisFun.Val2(H1Op_TrnKbn(0).Tag)
    SL_Select = SL_Select & " AND          年 = '" & RTrim$(iH1_YY) - 1 & "'"
    
    With CisDB
        .SQL = SL_Select
        .DBExec
    End With

End Function
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'▽
'▽                         名　称　マ　ス　タ　確　認  ( シ ス テ ム 共 通 定 数 等 )
'▽
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽▲▽
'*---------------------------------------------------------------------------*
'*      カ  レ  ン  ダ  デ　ー　タ　保　持　期　間  獲  得  ＆  セ  ッ  ト      *
'*---------------------------------------------------------------------------*
Private Sub RetentionTime_Check()
    
    Call ItemsClearSYMCreate
    
    With SYMCreate
        .区分名称 = "データ管理"
        .区分桁数 = 5
        .値 = "CLM"
        .制御区分 = "X"
        .値名称 = "保持期間(ヶ月)"
        .数字1 = 12
        .作成者 = gTanto
    End With
    
    Call SYMInitCreate(False)

    Call ItemsClearSYM
    gSL_Select = "Select 数字1 From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = 'データ管理'"
    gSL_Select = gSL_Select & " And   値 = 'CLM'"
    Call SYMRead(gSL_Select, 1)
    
    w_ReTime = SYM.数字1

End Sub

VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{8F4FF474-3E18-11D4-AEC2-00000EA57F9E}#1.0#0"; "CISMSG~1.OCX"
Begin VB.Form CNJ0450 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "内示シュミレーション"
   ClientHeight    =   11235
   ClientLeft      =   1440
   ClientTop       =   1920
   ClientWidth     =   15360
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D PNL_Wait 
      Height          =   930
      Left            =   4710
      Top             =   5310
      Visible         =   0   'False
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   1640
      BackColor       =   16761024
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      Begin Cis3D_v60.CIS3D Message 
         Height          =   525
         Left            =   240
         Top             =   210
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   926
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "【抽出中】しばらくお待ち下さい"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         cBoderWidth     =   2
         cAlingnment     =   7
         cBoderStyle     =   1
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   7590
      Top             =   420
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   150
      ScaleHeight     =   645
      ScaleWidth      =   2715
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   600
      Width           =   2775
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   540
         Left            =   60
         Top             =   60
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   953
         ForeColor       =   16711680
         Caption         =   "対象年月"
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
         cAlingnment     =   1
         cPositionX      =   90
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   375
            Left            =   1140
            TabIndex        =   0
            Top             =   90
            Width           =   1410
            _ExtentX        =   2328
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cSize           =   -1  'True
            cChkResult      =   0   'False
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   150
      ScaleHeight     =   2055
      ScaleWidth      =   14970
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1440
      Width           =   15030
      Begin vsFlexLib.vsFlexArray VsGrid1 
         Height          =   1905
         Left            =   90
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   60
         Width           =   14805
         _Version        =   196608
         _ExtentX        =   26114
         _ExtentY        =   3360
         _StockProps     =   228
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "№|XXXXXXXXX1XXXXXXXXX1XXXXX|XXXXXXX|XXXXXXXXX1XXXXXXXXX |XX|SK |展開 |999,999|99,999,999|99,999,999|99,999,999|99,999,999|"
         Rows            =   5
         Cols            =   17
         BackColor       =   12648384
         BackColorSel    =   8388608
         BackColorBkg    =   12632256
         FocusRect       =   0
         RowHeightMin    =   360
         SelectionMode   =   1
         MergeCells      =   3
         AllowBigSelection=   0   'False
         BackColorAlternate=   12648384
         Begin Cis3D_v60.CIS3D Input_Area 
            Height          =   360
            Left            =   330
            Top             =   480
            Width           =   14085
            _ExtentX        =   24844
            _ExtentY        =   635
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
            cBoderStyle     =   1
            Begin VB.CheckBox B1Chk_TenKB 
               BackColor       =   &H00C0C0C0&
               Caption         =   "有"
               Height          =   255
               Left            =   7350
               TabIndex        =   4
               Top             =   60
               Width           =   555
            End
            Begin CisText_V60.CisText iB1_Hinbn 
               Height          =   360
               Left            =   0
               TabIndex        =   1
               Top             =   0
               Width           =   3150
               _ExtentX        =   5556
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   25
            End
            Begin CisText_V60.CisText iB1_NSur1 
               Height          =   360
               Left            =   8895
               TabIndex        =   5
               Top             =   0
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   635
               cFormat         =   "##,###,###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "##,###,###"
               MaxLength       =   10
               cDataType       =   1
               cDataReplace    =   1
               cFdAutoFormat   =   1
               cGFormat        =   "########"
               cILength        =   8
            End
            Begin CisText_V60.CisText iB1_Torcd 
               Height          =   360
               Left            =   3075
               TabIndex        =   2
               Top             =   0
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXX"
               MaxLength       =   7
               cDataReplace    =   1
            End
            Begin CisText_V60.CisText iB1_Ukeir 
               Height          =   360
               Left            =   6480
               TabIndex        =   3
               Top             =   0
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XX"
               MaxLength       =   2
               cDataReplace    =   1
            End
            Begin CisText_V60.CisText iB1_NSur2 
               Height          =   360
               Left            =   10185
               TabIndex        =   6
               Top             =   0
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   635
               cFormat         =   "##,###,###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "##,###,###"
               MaxLength       =   10
               cDataType       =   1
               cDataReplace    =   1
               cFdAutoFormat   =   1
               cGFormat        =   "########"
               cILength        =   8
            End
            Begin CisText_V60.CisText iB1_NSur3 
               Height          =   360
               Left            =   11475
               TabIndex        =   7
               Top             =   0
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   635
               cFormat         =   "##,###,###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "##,###,###"
               MaxLength       =   10
               cDataType       =   1
               cDataReplace    =   1
               cFdAutoFormat   =   1
               cGFormat        =   "########"
               cILength        =   8
            End
            Begin CisText_V60.CisText iB1_ZSur 
               Height          =   360
               Left            =   12765
               TabIndex        =   8
               Top             =   0
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   635
               cFormat         =   "##,###,###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "##,###,###"
               MaxLength       =   10
               cDataType       =   1
               cDataReplace    =   1
               cFdAutoFormat   =   1
               cGFormat        =   "########"
               cILength        =   8
            End
            Begin Cis3D_v60.CIS3D B1lb_Tornm 
               Height          =   360
               Left            =   4020
               Top             =   0
               Width           =   2460
               _ExtentX        =   4339
               _ExtentY        =   635
               BackColor       =   16777152
               Caption         =   "XXXXXXXXX1XX"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
            Begin Cis3D_v60.CIS3D B1lb_SKKbn 
               Height          =   360
               Left            =   6840
               Top             =   0
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   635
               BackColor       =   16777152
               Caption         =   "XX"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   1
            End
            Begin Cis3D_v60.CIS3D B1lb_Syuyo 
               Height          =   360
               Left            =   7965
               Top             =   0
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   635
               BackColor       =   16777152
               Caption         =   "999,999"
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -30
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   360
            Left            =   330
            Top             =   0
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   635
            ForeColor       =   0
            Caption         =   "品     番"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   360
            Left            =   3405
            Top             =   0
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   635
            ForeColor       =   0
            Caption         =   "納  入  先"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   360
            Left            =   6840
            Top             =   0
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   635
            ForeColor       =   0
            Caption         =   "受"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D8 
            Height          =   360
            Left            =   7170
            Top             =   0
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   635
            ForeColor       =   0
            Caption         =   "SK"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D BLB_MM 
            Height          =   360
            Index           =   1
            Left            =   9240
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   635
            ForeColor       =   0
            Caption         =   "9999年99月"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D BLB_MM 
            Height          =   360
            Index           =   2
            Left            =   10530
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   635
            ForeColor       =   0
            Caption         =   "9999年99月"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D BLB_MM 
            Height          =   360
            Index           =   3
            Left            =   11820
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   635
            ForeColor       =   0
            Caption         =   "9999年99月"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D12 
            Height          =   360
            Left            =   13110
            Top             =   0
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   635
            ForeColor       =   0
            Caption         =   "繰越在庫"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D13 
            Height          =   360
            Left            =   8295
            Top             =   0
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   635
            ForeColor       =   0
            Caption         =   "収容数"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   360
            Left            =   7620
            Top             =   0
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   635
            ForeColor       =   0
            Caption         =   "展開"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   360
            Left            =   0
            Top             =   0
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   635
            ForeColor       =   0
            Caption         =   "№"
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
            cAlingnment     =   7
         End
      End
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
      Height          =   765
      Left            =   9060
      ScaleHeight     =   705
      ScaleWidth      =   6045
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   540
      Width           =   6105
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   3450
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   3
         BtnCaption      =   "取消"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   32768
         BtnShadowColor  =   16777215
         PFCaption       =   "( Esc )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   65535
         PFShadowColor   =   0
         PFBackColor     =   49152
      End
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   4290
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   2
         BtnCaption      =   "入力"
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
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   5130
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   1
         BtnCaption      =   "終了"
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   930
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   11
         BtnCaption      =   "検索"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
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
         PFBackColor     =   33023
      End
      Begin CisBtn_60.CisBtn PB_CHG 
         Height          =   585
         Left            =   90
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   10
         BtnCaption      =   "切替"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   8388736
         BtnShadowColor  =   16777215
         PFCaption       =   "( F6 )"
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
         PFBackColor     =   8388736
      End
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   1770
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   5
         BtnCaption      =   "前進"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   8421376
         BtnShadowColor  =   16777215
         PFCaption       =   "( Down )"
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
         PFBackColor     =   8421376
      End
      Begin CisBtn_60.CisBtn PB_BACK 
         Height          =   585
         Left            =   2610
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   6
         BtnCaption      =   "後退"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   8421376
         BtnShadowColor  =   16777215
         PFCaption       =   "( Up )"
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
         PFBackColor     =   8421376
      End
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4650
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   -2310
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Left            =   900
      Top             =   -1380
   End
   Begin VB.Timer Timer2 
      Left            =   1170
      Top             =   -1380
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 内示シュミレーション  】"
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
         Left            =   14400
         Top             =   60
         Width           =   795
         _ExtentX        =   1402
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
         Left            =   11850
         Top             =   30
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
         cPositionX      =   50
      End
   End
   Begin VB.TextBox B_TabDummy 
      Height          =   330
      Left            =   7620
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1860
      Width           =   240
   End
   Begin Cis3D_v60.CIS3D BLB_Proc 
      Height          =   690
      Left            =   3090
      Top             =   630
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   1217
      BackColor       =   4194368
      ForeColor       =   16761024
      Caption         =   "構成情報"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   20.25
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
   End
   Begin VB.PictureBox B3_Area1 
      Height          =   7245
      Left            =   4680
      ScaleHeight     =   7185
      ScaleWidth      =   10455
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3780
      Width           =   10515
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   2670
         Left            =   60
         Top             =   60
         Width           =   10350
         _ExtentX        =   18256
         _ExtentY        =   4710
         ForeColor       =   16711680
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   330
            Left            =   60
            Top             =   570
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "№"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D14 
            Height          =   330
            Left            =   540
            Top             =   570
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "品      番"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_No 
            Height          =   390
            Left            =   60
            Top             =   900
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   688
            BackColor       =   16777152
            Caption         =   "999"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
         Begin Cis3D_v60.CIS3D B3lb_Hinbn 
            Height          =   390
            Left            =   540
            Top             =   900
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   688
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX1XXXXX"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin CisMsg_v60.CISMSG CISMSG2 
            Height          =   510
            Left            =   60
            TabIndex        =   18
            Top             =   60
            Width           =   10230
            _ExtentX        =   18045
            _ExtentY        =   900
            BackColor       =   -2147483633
            ForeColor       =   12583104
            Caption         =   "※  日  程  情  報  ※"
            BackStyle       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   14.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   0
            cAlingnment     =   7
            cBoderStyle     =   1
            gGC2_Green      =   160
            gGradation      =   -1  'True
         End
         Begin Cis3D_v60.CIS3D CIS3D16 
            Height          =   330
            Left            =   3600
            Top             =   570
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "背番"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Seban 
            Height          =   390
            Left            =   3600
            Top             =   900
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   688
            BackColor       =   16777152
            Caption         =   "XXXXXXXX"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D32 
            Height          =   330
            Left            =   4635
            Top             =   570
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "手  配  先"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Torcd 
            Height          =   390
            Left            =   4635
            Top             =   900
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   688
            BackColor       =   16777152
            Caption         =   "XXXXXXX:XXXXXXXXX1XXXX"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D CIS3D20 
            Height          =   330
            Left            =   7365
            Top             =   570
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "受"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Ukeir 
            Height          =   390
            Left            =   7365
            Top             =   900
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   688
            BackColor       =   16777152
            Caption         =   "XX"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D22 
            Height          =   330
            Left            =   7785
            Top             =   570
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "納入先"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_NTorcd 
            Height          =   390
            Left            =   7785
            Top             =   900
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   688
            BackColor       =   16777152
            Caption         =   "XXXXXXX:XXXXXXXXX1"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   330
            Index           =   0
            Left            =   60
            Top             =   1290
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "SK"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_SKKbn 
            Height          =   330
            Left            =   60
            Top             =   1620
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "Ｋ"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   330
            Index           =   1
            Left            =   1470
            Top             =   1290
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "収容数1"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Syuyo1 
            Height          =   330
            Left            =   1470
            Top             =   1620
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "999,999"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   330
            Index           =   2
            Left            =   2460
            Top             =   1290
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "ロット1"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Lot1 
            Height          =   330
            Left            =   2460
            Top             =   1620
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "9,999,999"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   330
            Index           =   5
            Left            =   3660
            Top             =   1290
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "変更日 - 便"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_CDate 
            Height          =   330
            Left            =   3660
            Top             =   1620
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "9999/99/99-99"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   330
            Index           =   6
            Left            =   540
            Top             =   1290
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "ｻｲｸﾙ"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Cycle 
            Height          =   330
            Left            =   540
            Top             =   1620
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "1-01-01"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   330
            Index           =   3
            Left            =   5370
            Top             =   1290
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "収容数2"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Syuyo2 
            Height          =   330
            Left            =   5370
            Top             =   1620
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "999,999"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   330
            Index           =   4
            Left            =   6360
            Top             =   1290
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "ロット2"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Lot2 
            Height          =   330
            Left            =   6360
            Top             =   1620
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "9,999,999"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
         Begin Cis3D_v60.CIS3D B3lb_Kigen 
            Height          =   330
            Left            =   7560
            Top             =   1620
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "9999/99/99～9999/99/99"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D CIS3D29 
            Height          =   330
            Left            =   7560
            Top             =   1290
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "有効期限"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Keisu 
            Height          =   330
            Index           =   1
            Left            =   2520
            Top             =   2280
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "9.99"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
         Begin Cis3D_v60.CIS3D B3lb_Keisu 
            Height          =   330
            Index           =   2
            Left            =   3180
            Top             =   2280
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "9.99"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   330
            Index           =   9
            Left            =   2520
            Top             =   1950
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   582
            ForeColor       =   8388608
            Caption         =   "係数"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Keisu 
            Height          =   330
            Index           =   3
            Left            =   3840
            Top             =   2280
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "9.99"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   330
            Index           =   10
            Left            =   60
            Top             =   1950
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "LT"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_LT 
            Height          =   330
            Left            =   60
            Top             =   2280
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "99"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D BLB_Week 
            Height          =   330
            Left            =   600
            Top             =   1950
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "曜日指定"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Week 
            Height          =   330
            Index           =   1
            Left            =   600
            Top             =   2280
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "XX"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Week 
            Height          =   330
            Index           =   2
            Left            =   990
            Top             =   2280
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "XX"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Week 
            Height          =   330
            Index           =   3
            Left            =   1380
            Top             =   2280
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "XX"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Week 
            Height          =   330
            Index           =   4
            Left            =   1770
            Top             =   2280
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "XX"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Week 
            Height          =   330
            Index           =   5
            Left            =   2160
            Top             =   2280
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "XX"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   330
            Index           =   12
            Left            =   5970
            Top             =   1950
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "当月数"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_NSur1 
            Height          =   330
            Left            =   5970
            Top             =   2280
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "999,999,999"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   330
            Index           =   13
            Left            =   7410
            Top             =   1950
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "翌月数"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_NSur2 
            Height          =   330
            Left            =   7410
            Top             =   2280
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   582
            BackColor       =   12648384
            Caption         =   "999,999,999"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   330
            Index           =   14
            Left            =   8850
            Top             =   1950
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   8388608
            Caption         =   "翌々月数"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_NSur3 
            Height          =   330
            Left            =   8850
            Top             =   2280
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   582
            BackColor       =   12648447
            Caption         =   "999,999,999"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   330
            Index           =   15
            Left            =   4530
            Top             =   1950
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   582
            ForeColor       =   8388608
            Caption         =   "平準化"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_HJRitu 
            Height          =   330
            Left            =   5370
            Top             =   2280
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "99%"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
         Begin Cis3D_v60.CIS3D B3lb_HJKbn 
            Height          =   330
            Left            =   4530
            Top             =   2280
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
            BackColor       =   16761024
            Caption         =   "XX"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin vsFlexLib.vsFlexArray vsGrid3 
         Height          =   4365
         Left            =   90
         TabIndex        =   19
         Top             =   2760
         Width           =   10290
         _Version        =   196608
         _ExtentX        =   18150
         _ExtentY        =   7699
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "      |１日      |１日      |１日      |１日      |１日      |１日      |１日      |"
         Rows            =   36
         Cols            =   11
         BackColor       =   16777215
         BackColorFixed  =   14737632
         BackColorBkg    =   12632256
         TextStyle       =   4
         TextStyleFixed  =   3
         FocusRect       =   2
         HighLight       =   0
         RowHeightMin    =   330
         FillStyle       =   1
         GridLines       =   2
         GridColor       =   14737632
         SelectionMode   =   2
         BackColorAlternate=   16777215
         Begin Cis3D_v60.CIS3D B3lb_WeekDay 
            Height          =   330
            Index           =   1
            Left            =   795
            Top             =   0
            Width           =   1300
            _ExtentX        =   2302
            _ExtentY        =   582
            BackColor       =   12632319
            ForeColor       =   0
            Caption         =   "日曜"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D B3lb_WeekDay 
            Height          =   330
            Index           =   2
            Left            =   2115
            Top             =   0
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "月曜"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D B3lb_WeekDay 
            Height          =   330
            Index           =   3
            Left            =   3435
            Top             =   0
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "火曜"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D B3lb_WeekDay 
            Height          =   330
            Index           =   4
            Left            =   4740
            Top             =   0
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "水曜"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D B3lb_WeekDay 
            Height          =   330
            Index           =   5
            Left            =   6045
            Top             =   0
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "木曜"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D B3lb_WeekDay 
            Height          =   330
            Index           =   6
            Left            =   7365
            Top             =   0
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "金曜"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D B3lb_WeekDay 
            Height          =   330
            Index           =   7
            Left            =   8670
            Top             =   0
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   582
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "土曜"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   3
            cAlingnment     =   7
         End
      End
   End
   Begin VB.PictureBox B2_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   7260
      Left            =   150
      ScaleHeight     =   7200
      ScaleWidth      =   14970
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3780
      Width           =   15030
      Begin vsFlexLib.vsFlexArray vsGrid2 
         Height          =   7035
         Left            =   90
         TabIndex        =   16
         Top             =   90
         Width           =   14805
         _Version        =   196608
         _ExtentX        =   26114
         _ExtentY        =   12409
         _StockProps     =   228
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "№ |レベル|XXXXXXXXX1XXXXXXXXX1|XXXXXXXX|XXXXXXXXX1XXXXXXXXX1X|受|XXXXXXXX|SK|999,999|999,999,999 |999,999,999 |999,999,999 |"
         Rows            =   21
         Cols            =   17
         BackColor       =   16777152
         TextStyleFixed  =   1
         RowHeightMin    =   330
         SelectionMode   =   1
         MergeCells      =   3
         BackColorAlternate=   16777152
      End
   End
End
Attribute VB_Name = "CNJ0450"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   内示シュミレーション
'**       フォームID    :   CNJ0450
'**       処理概要      :
'**
'**       作  成  日    :   2004/04/14  By CIS SIMPR-A
'**       変  更  日    :   2008/04/11  By CIS 背番号8桁
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    Dim CisVsGridT      As New CisVsGrid3
    Dim CisVsGridM      As New CisVsGrid3
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKb         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim ActObj          As Object
    Dim ErrObj          As Object
    Dim ActRow          As Integer
    Dim wsNJYM          As String
    Dim wsTNJYM(3)      As String
    
    Dim wsLastDayX      As String
    Dim wsLastDay_TG    As Byte
    Dim wsLastDay_YG    As Byte
    Dim wsLastDay_YYG   As Byte
    Dim wsDayDisp_TG    As String
    Dim wsDayDisp_YG    As String
    Dim wsDayDisp_YYG   As String
    
    Dim mWeekStart              As Integer
    Dim mDayRow(31)             As Integer
    Dim mDayCol(31)             As Integer
    Dim wsProcKB        As String
    Dim wsCL_BackColorK     As Long     ' 稼動
    Dim wsCL_BackColorY     As Long     ' 休日

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   # 二 重 起 動 防 止 #
'   #------------------#
    If Not Dupli_Start Then End
'   #------------------#
'   # 初 期 内 容 取 得 #
'   #------------------#
    If Not IniGet Then End
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Set CisFrm = New CISFormContorl
    
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
    
    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
    
  
' 生産計画年月
    wsNJYM = ""
    Call ItemsClearSYM
    gSL_Select = "Select * From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '処理内容'"
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値 = 'PNJ0210'"
    If SYMRead(gSL_Select, 1) Then
        With SYM
            wsNJYM = SYM.英数字1
        End With
    End If
    
    If RTrim(wsNJYM) = "" Then
        iH1_YM = Format(Now(), "yyyymm")
    Else
        iH1_YM = wsNJYM
    End If
'    B2Op_MM(1).Tag = 1
    
    iH1_YM = Format(Now(), "yyyymm")
    Call SetYM
    
    ' 曜日開始を求める
    gSL_Select = "select 数字1 From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '曜日開始管理'"
    gSL_Select = gSL_Select & "   And 値 = '1'"
    If Not SYMRead(gSL_Select, 1) Then
       SYM.数字1 = 0
    End If
    
    mWeekStart = SYM.数字1 + 1
    
  ' 名称マスタ確認
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = VsGrid1
        .NoSpaceCol = 1
        .NumberSet = All
        .FixedHeight = 360
        .PatanMax = 1
        .SelectCol = 15
        .InitGet = False
        .Init
    End With
    Call GridInit
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGridT
        Set .GridObj = vsGrid2
        .NoSpaceCol = 1
        .NumberSet = All
        .FixedHeight = 360
        .PatanMax = 3
        .SelectCol = 16
        .InitGet = False
        .Init
    End With
    Call GridTInit

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    wsCL_BackColorK = B3lb_WeekDay(2).BackColor
    wsCL_BackColorY = B3lb_WeekDay(1).BackColor
    B3lb_WeekDay(1).BackColor = wsCL_BackColorK

    FormAct = Not FormAct
'
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》                 +
'+-------------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+-------------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》         +
'+-------------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    Select Case Key_Code
           Case vbKeyF6:       If PB_CHG.Visible Then Call PB_CHG_Click    '【切替】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_Can_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyDown:     If PB_NEXT.Visible Then Call PB_NEXT_Click  '【前進】
           Case vbKeyUp:       If PB_BACK.Visible Then Call PB_BACK_Click  '【後退】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      '【入力】
                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
                                  Dummy.Enabled = False
                               End If
           Case Else:          Exit Function
    End Select
    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Cancel = 1
    Call PB_END_Click
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
'+---------------------+
'+   終了処理
'+---------------------+
'    Call EndProc
'+---------------------+
'+   ﾃﾞｰﾀﾍﾞｰｽ切断
'+---------------------+
    Call CisDB.DBDISConnect
'+---------------------+
'+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
'+---------------------+
    Call CisFrm.UnLoadDisp
    End
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
   Call ReturnPress
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_Can_Click()
   Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   Call CisFrm.MousePT(11)
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      iH1_YM.SetFocus
      GoTo PB_CAN_Ed
   End If

   If ProcHB = "B2" Then
      ProcHB = "B1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B2")
      GoTo PB_CAN_Ed
   End If

   If ProcHB = "B3" Then
      ProcHB = "B2"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B3")
      GoTo PB_CAN_Ed
   End If

PB_CAN_Ed:
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iB1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iB1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iB1_Ukeir" Then Call Look_Ukeir: Exit Sub
End Sub
'+----------------------------+
'+       取引先検索            +
'+----------------------------+
Private Sub Look_Torcd()

    RV_Left = 0:    RV_Top = 0
    RV_TorKb = 1: RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
       iB1_Torcd = RTrim(RV_TorcdK)
       Call iB1_Torcd_LostFocus
    End If

End Sub
'+----------------------------+
'+    切替キー
'+----------------------------+
Private Sub PB_CHG_Click()
    If RTrim(BLB_Proc.Tag) = "" Then
        BLB_Proc.Tag = "X"
        BLB_Proc.Caption = "展開情報"
    Else
        BLB_Proc.Tag = ""
        BLB_Proc.Caption = "構成情報"
    End If
End Sub
'+----------------------------+
'+    前進  キ ー(NEXT)       +
'+----------------------------+
Private Sub PB_NEXT_Click()

    If ProcHB <> "B3" Then: Exit Sub
    
    gInt = CisFun.Val2(B3lb_No)
    If RTrim(vsGrid2.TextMatrix(gInt + 1, 1)) <> "" Then
        Call Body3Set(gInt + 1)
    End If

End Sub
'+----------------------------+
'+    前進  キ ー(BACK)       +
'+----------------------------+
Private Sub PB_BACK_Click()

    If ProcHB <> "B3" Then: Exit Sub
    
    gInt = CisFun.Val2(B3lb_No)
    
    If gInt <> 1 Then
        Call Body3Set(gInt - 1)
    End If

End Sub

'+----------------------------+
'+       受入検索            +
'+----------------------------+
Private Sub Look_Ukeir()

    RV_Left = 0: RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 1
    RV_Torcd = iB1_Torcd
    RV_Kojyo = ""
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
       iB1_Torcd = RTrim(RV_TorcdK)
       Call iB1_Torcd_LostFocus
       iB1_Ukeir = RTrim(RV_Ukeir)
    End If

End Sub
'+----------------------------+
'+       品番検索            +
'+----------------------------+
Private Sub Look_Hinbn()

    RV_Left = 0: RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 0
    RVI_SK = 9
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
       iB1_Hinbn = RTrim(RV_Hinbn)
       iB1_Torcd = RTrim(RV_TorcdK)
       Call iB1_Torcd_LostFocus
       iB1_Ukeir = RTrim(RV_Ukeir)
       
       Call GetSyuyo
    End If

End Sub
'+-----------------------------+
'+  iH1_YM
'+-----------------------------+
Private Sub iH1_YM_GotFocus()
    PB_Look.Tag = ""
    PB_Look.Visible = False
End Sub
'+-----------------------------+
'+  iB1_Hinbn
'+-----------------------------+
Private Sub iB1_Hinbn_GotFocus()
    Set ActObj = iB1_Hinbn
    B_TabDummy.Tag = "S"

    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Hinbn_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

' *品番取引先マスタ情報 参照*
    Call GetSyuyo
End Sub
'+-----------------------------+
'+  iB1_Torcd
'+-----------------------------+
Private Sub iB1_Torcd_GotFocus()
    Set ActObj = iB1_Torcd

    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Torcd_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_Tornm = ""
    If RTrim(iB1_Torcd) <> "" Then
        If TorNmGet(iB1_Torcd, 3) Then
            B1lb_Tornm = RTrim(TRM.略称)
        End If
    End If

' *品番取引先マスタ情報 参照*
    Call GetSyuyo
End Sub
'+-----------------------------+
'+  iB1_Ukeir
'+-----------------------------+
Private Sub iB1_Ukeir_GotFocus()
    Set ActObj = iB1_Ukeir

    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Ukeir_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Ukeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

' *品番取引先マスタ情報 参照*
    Call GetSyuyo
End Sub
'+-----------------------------+
'+  iB1_NSur1
'+-----------------------------+
Private Sub iB1_NSur1_GotFocus()
    Set ActObj = iB1_NSur1
End Sub
Private Sub iB1_NSur1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_NSur1_LostFocus()
    If CisFun.Val2(iB1_NSur1) <> 0 Then
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 7) = iB1_NSur1
    Else
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 7) = ""
    End If
End Sub
'+-----------------------------+
'+  iB1_NSur2
'+-----------------------------+
Private Sub iB1_NSur2_GotFocus()
    Set ActObj = iB1_NSur2
End Sub
Private Sub iB1_NSur2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_NSur2_LostFocus()
    If CisFun.Val2(iB1_NSur2) <> 0 Then
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 8) = iB1_NSur2
    Else
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 8) = ""
    End If
End Sub
'+-----------------------------+
'+  iB1_NSur3
'+-----------------------------+
Private Sub iB1_NSur3_GotFocus()
    Set ActObj = iB1_NSur3
End Sub
Private Sub iB1_NSur3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_NSur3_LostFocus()
    If CisFun.Val2(iB1_NSur3) <> 0 Then
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 9) = iB1_NSur3
    Else
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 9) = ""
    End If
End Sub
'+-----------------------------+
'+  iB1_ZSur
'+-----------------------------+
Private Sub iB1_ZSur_GotFocus()
    Set ActObj = iB1_ZSur
    B_TabDummy.Tag = "E"
End Sub
Private Sub iB1_ZSur_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_ZSur_LostFocus()
    If CisFun.Val2(iB1_ZSur) <> 0 Then
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 9) = iB1_ZSur
    Else
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 9) = ""
    End If
End Sub
'+-----------------------------+
'+  B1Chk_TenKB
'+-----------------------------+
Private Sub B1Chk_TenKB_GotFocus()
    Set ActObj = B1Chk_TenKB
End Sub
Private Sub B1Chk_TenKB_CLICK()
    If B1Chk_TenKB.Value = 0 Then
        B1Chk_TenKB.Caption = "有"
    Else
        B1Chk_TenKB.Caption = "無"
    End If
End Sub
Private Sub B1Chk_TenKB_LostFocus()
    If B1Chk_TenKB.Value = 0 Then
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 11) = ""
    Else
        VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 11) = "無"
    End If
    VsGrid1.TextMatrix(CisFun.Val2(Input_Area.Tag), 14) = B1Chk_TenKB.Value
End Sub
'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If ProcHB = "B1" Then
        If B_TabDummy.Tag = "S" Then
            Call TextToVsGrid(Val(Input_Area.Tag))
            Call Grid_RowChange(vbKeyUp)
        End If
        If B_TabDummy.Tag = "E" Then
            Call TextToVsGrid(Val(Input_Area.Tag))
            Call Grid_RowChange(vbKeyDown)
        End If
        SendKeys "{Tab}"
    End If
End Sub
'+-----------------------------+
'+  B2Op_MM
'+-----------------------------+
Private Sub B2Op_MM_Click(Index As Integer)
'    B2Op_MM(1).Tag = Index
    
'    Select Case Index
'        Case 1: Call CisVsGridT.LayOut(0)
'        Case 2: Call CisVsGridT.LayOut(1)
'        Case 3: Call CisVsGridT.LayOut(2)
'    End Select
    Call Body_Set(Index)
End Sub

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
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
    Dim B1Color     As Long
    Dim B2Mode      As Boolean
    Dim B2Color     As Long
    Dim B3Mode      As Boolean
    Dim B3Color     As Long
    
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
    B2Mode = False
    B2Color = gPLostSel
    B3Mode = False
    B3Color = gPLostSel

' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1"
          H1Mode = True
          H1Color = gPGotSel
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
       Case "B2"
          B2Mode = True
          B2Color = gPGotSel
       Case "B3"
          B3Mode = True
          B3Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B2_Area1.Enabled = B2Mode Or B3Mode
    B2_Area1.BackColor = B2Color
    B3_Area1.Enabled = B3Mode
    B3_Area1.Visible = B3Mode
    B3_Area1.BackColor = B3Color

    PB_Look.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_CHG.Visible = B1Mode
    PB_END.Visible = H1Mode
    PB_NEXT.Visible = B3Mode
    PB_BACK.Visible = B3Mode
    
    Call GridTextEnabled(B1Mode)
    
' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'**************************************
'*      Head/Body 初期化処理           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String
    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*" Else wbc_Enm = Trim(ClrType)
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then
                .Text = ""
                .Enabled = True
            End If
            If .Name Like "Chk" & wbc_Enm & "_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
        CisVsGrid.Clear
'        Back_UYmd.Visible = False
    End If
    If HeadBody = "B1" Then
        CisVsGrid.Clear
    End If
    If HeadBody = "B2" Then
        CisVsGridT.Clear
    End If


End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
            
        ProcHB = "B1"
        Call DispChange(ProcHB)
        VsGrid1.ROW = 1
        Call VsGridToText(1)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        ProcHB = "B2"
        Call DispChange(ProcHB)
        
        If Not Tenkai_Proc Then GoTo ReturnPress_Ed
'        Call B2Op_MM_Click(1)
        Call Body_Set(1)
        
        GoTo ReturnPress_Ed
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With

' * 対象年月
    If Trim(iH1_YM) = "" Then
        CisFun.MB_MSG(2) = "    対象年月を入力して下さい        "
        CisFun.MBOX
        iH1_YM.SetFocus
        Exit Function
    End If
       
' * 年月セット
    Call SetYM

    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call GridTextEnabled(False)
'ﾒｯｾｰｼﾞﾎﾞｯｸｽ設定
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
'<< 行を詰める >>
    With VsGrid1
        For gCnt1 = 1 To .Rows - 1
            If RTrim(.TextMatrix(gCnt1, 1)) = "" And _
                RTrim(.TextMatrix(gCnt1, 2)) = "" And _
                RTrim(.TextMatrix(gCnt1, 3)) = "" And _
                RTrim(.TextMatrix(gCnt1, 4)) = "" Then
                For gCnt2 = gCnt1 + 1 To .Rows - 1
                    If RTrim(.TextMatrix(gCnt2, 1)) <> "" Or _
                        RTrim(.TextMatrix(gCnt2, 2)) <> "" Or _
                        RTrim(.TextMatrix(gCnt2, 3)) <> "" Or _
                        RTrim(.TextMatrix(gCnt2, 4)) <> "" Then
                        .TextMatrix(gCnt1, 1) = .TextMatrix(gCnt2, 1)
                        .TextMatrix(gCnt1, 2) = .TextMatrix(gCnt2, 2)
                        .TextMatrix(gCnt1, 3) = .TextMatrix(gCnt2, 3)
                        .TextMatrix(gCnt1, 4) = .TextMatrix(gCnt2, 4)
                        .TextMatrix(gCnt1, 5) = .TextMatrix(gCnt2, 5)
                        .TextMatrix(gCnt1, 6) = .TextMatrix(gCnt2, 6)
                        .TextMatrix(gCnt1, 7) = .TextMatrix(gCnt2, 7)
                        .TextMatrix(gCnt1, 8) = .TextMatrix(gCnt2, 8)
                        .TextMatrix(gCnt1, 9) = .TextMatrix(gCnt2, 9)
                        .TextMatrix(gCnt1, 10) = .TextMatrix(gCnt2, 10)
                        .TextMatrix(gCnt1, 11) = .TextMatrix(gCnt2, 11)
                        .TextMatrix(gCnt1, 12) = .TextMatrix(gCnt2, 12)
                        
                        .TextMatrix(gCnt2, 1) = ""
                        .TextMatrix(gCnt2, 2) = ""
                        .TextMatrix(gCnt2, 3) = ""
                        .TextMatrix(gCnt2, 4) = ""
                        .TextMatrix(gCnt2, 5) = ""
                        .TextMatrix(gCnt2, 6) = ""
                        .TextMatrix(gCnt2, 7) = ""
                        .TextMatrix(gCnt2, 8) = ""
                        .TextMatrix(gCnt2, 9) = ""
                        .TextMatrix(gCnt2, 10) = ""
                        .TextMatrix(gCnt2, 11) = ""
                        .TextMatrix(gCnt2, 12) = ""
                        Exit For
                    End If
                Next gCnt2
            End If
        Next gCnt1
'
        For gCnt1 = 1 To .Rows - 1
            If RTrim(.TextMatrix(gCnt1, 1)) = "" And _
                RTrim(.TextMatrix(gCnt1, 2)) = "" And _
                RTrim(.TextMatrix(gCnt1, 3)) = "" And _
                RTrim(.TextMatrix(gCnt1, 4)) = "" Then
                Exit For
            End If
        
            If RTrim(.TextMatrix(gCnt1, 1)) = "" Then
                CisFun.MB_MSG(1) = "    品番を入力して下さい      "
                CisFun.MB_MSG(3) = "      【 " & gCnt1 & "行目 】"
                CisFun.MBOX
                Set ErrObj = iB1_Hinbn
                gCnt1 = gCnt2
                GoTo Body1Chk_Err
            End If

            If RTrim(.TextMatrix(gCnt1, 2)) = "" Then
                CisFun.MB_MSG(1) = "    納入先を入力して下さい      "
                CisFun.MB_MSG(3) = "      【 " & gCnt1 & "行目 】"
                CisFun.MBOX
                Set ErrObj = iB1_Torcd
                gCnt1 = gCnt2
                GoTo Body1Chk_Err
            End If

            If RTrim(.TextMatrix(gCnt1, 4)) = "" Then
                CisFun.MB_MSG(1) = "    受入を入力して下さい      "
                CisFun.MB_MSG(3) = "      【 " & gCnt1 & "行目 】"
                CisFun.MBOX
                Set ErrObj = iB1_Ukeir
                gCnt1 = gCnt2
                GoTo Body1Chk_Err
            End If

        '<< 品番取引先マスタチェック >>
            gSL_Select = "SELECT * FROM 品番取引先マスタ "
            gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(.TextMatrix(gCnt1, 1)) & "'"
            gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(.TextMatrix(gCnt1, 2)) & "'"
            gSL_Select = gSL_Select & "   AND 受入 = '" & RTrim(.TextMatrix(gCnt1, 4)) & "'"
            If Not HTMRead(gSL_Select, 1) Then
                CisFun.MB_MSG(1) = "    品番取引先マスタ 未登録      "
                CisFun.MB_MSG(3) = "      【 " & gCnt1 & "行目 】"
                CisFun.MBOX
                Set ErrObj = iB1_Hinbn
                gCnt1 = gCnt2
                GoTo Body1Chk_Err
            End If
            
        '<< 重複チェック >>
            If RTrim(.TextMatrix(gCnt1, 1)) <> "" Then
            For gCnt2 = gCnt1 + 1 To .Rows - 1
'                If RTrim(.TextMatrix(gCnt1, 1)) = RTrim(.TextMatrix(gCnt2, 1)) And
                If RTrim(.TextMatrix(gCnt1, 1)) = RTrim(.TextMatrix(gCnt2, 1)) And _
                   RTrim(.TextMatrix(gCnt1, 2)) = RTrim(.TextMatrix(gCnt2, 2)) And _
                   RTrim(.TextMatrix(gCnt1, 4)) = RTrim(.TextMatrix(gCnt2, 4)) Then
                    CisFun.MB_MSG(1) = "    品番・納入先・受入が重複しています。            "
                    CisFun.MB_MSG(3) = "    " & gCnt1 & "行目と " & gCnt2 & "行目            "
                    CisFun.MBOX
                    Set ErrObj = iB1_Hinbn
                    gCnt1 = gCnt2
                    GoTo Body1Chk_Err
                End If
            Next gCnt2
            End If
            
        Next gCnt1
    End With
    
    Body1Chk = True
    Exit Function
Body1Chk_Err:
    VsGrid1.ROW = gCnt1
    Call VsGridToText(VsGrid1.ROW)
    Call GridTextEnabled(True)
    ErrObj.SetFocus
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    Call HeadBodyClear("B1")
    
    If Not KNTRead(gSL_Select, 0, 1) Then
        Call KNTClose
        If InputKey = "Back" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     先頭レコードです            "
                .MB_Title = "【後退】"
                .MB_Button = OK
                .MBOX
            End With
            GoTo DBInput_Ed
        End If
        If InputKey = "Next" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     最終レコードです            "
                .MB_Title = "【前進】"
                .MB_Button = OK
                .MBOX
            End With
            GoTo DBInput_Ed
        End If
    Else
'        Call GridSet
        DBInput = True
    End If
    
DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
Private Function DBPut()
    DBPut = False
    Dim AddFlg      As Boolean
    Dim wsKanri     As Long
    Dim wsMaisu       As Long
    Dim wsSuryo       As Long


    On Error GoTo DBPut_Err
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    

' * ワークテーブルクリア
    With CisDB
        .SQL = "DELETE FROM S内示テーブル "
        .DBExec
    End With

'   マスタ項目移送
    For gIndex = 1 To VsGrid1.Rows - 1

        If RTrim(VsGrid1.TextMatrix(gIndex, 1)) <> "" And _
           RTrim(VsGrid1.TextMatrix(gIndex, 6)) = "" Then

            Call ItemsClearSNJ
            With SNJ
                .内示年月 = RTrim(iH1_YM)
                .品番 = RTrim(VsGrid1.TextMatrix(gIndex, 1))
                .納入先 = RTrim(VsGrid1.TextMatrix(gIndex, 2))
                .受入 = RTrim(VsGrid1.TextMatrix(gIndex, 4))
                .内示数 = CisFun.Val2(VsGrid1.TextMatrix(gIndex, 8))
                .翌月内示数 = CisFun.Val2(VsGrid1.TextMatrix(gIndex, 9))
                .翌々月内示数 = CisFun.Val2(VsGrid1.TextMatrix(gIndex, 10))
                .繰越在庫数 = CisFun.Val2(VsGrid1.TextMatrix(gIndex, 11))
            End With
            Call SNJInsert
        End If

    Next gIndex
'
' ( 追加処理 )
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    B_TabDummy.Tag = ""
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
Private Function Tenkai_Proc()
    Tenkai_Proc = False

    Message = "【 展 開 処 理 中 】"
    PNL_Wait.Visible = True
    DoEvents

    With CisDB
' +-----------------------------+
' +     生産計画処理
' +-----------------------------+
        .SQL = "S生産計画処理"
        .StoadoCount = 3
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = OutPut用
        .ParaNo = 3: .ParaIO = OutPut用
        .ParaNo = 1: .ParaValue = RTrim(iH1_YM)

        If Not .DBStored Then
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(2) = "    生産計画処理が失敗しました            "
            CisFun.MB_Title = "生産計画処理"
            CisFun.MB_Button = Error
            CisFun.MBOX
            Exit Function
        End If
    
        wsProcKB = RTrim(BLB_Proc.Tag)
        If wsProcKB <> "" Then
    ' +-----------------------------+
    ' +     展開処理_構成
    ' +-----------------------------+
            .SQL = "S展開処理"
            .StoadoCount = 3
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = Input用
            .ParaNo = 2: .ParaIO = OutPut用
            .ParaNo = 3: .ParaIO = OutPut用
            .ParaNo = 1: .ParaValue = RTrim(iH1_YM)
    
            If Not .DBStored Then
                CisFun.MB_Lines = 4
                CisFun.MB_MSG(2) = "    展開処理が失敗しました            "
                CisFun.MB_Title = "展開処理"
                CisFun.MB_Button = Error
                CisFun.MBOX
                Exit Function
            End If
        Else
    ' +-----------------------------+
    ' +     展開処理_構成
    ' +-----------------------------+
            .SQL = "S展開処理_構成"
            .StoadoCount = 2
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = Input用
            .ParaNo = 2: .ParaIO = OutPut用
            .ParaNo = 1: .ParaValue = RTrim(iH1_YM)
    
            If Not .DBStored Then
                CisFun.MB_Lines = 4
                CisFun.MB_MSG(2) = "    展開処理が失敗しました            "
                CisFun.MB_Title = "展開処理"
                CisFun.MB_Button = Error
                CisFun.MBOX
                Exit Function
            End If
        End If
    End With

    PNL_Wait.Visible = False
    DoEvents

    Tenkai_Proc = True
End Function
'*------------------------------------------*'
'*    品番取引先マスタ参照(収容数・SK区分)
'*------------------------------------------*'
Private Sub GetSyuyo()
    B1lb_SKKbn = "": B1lb_Syuyo = ""
    
    Call ItemsClearHTM
    gSL_Select = "SELECT * FROM 品番取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iB1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(iB1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND 受入 = '" & RTrim(iB1_Ukeir) & "'"
    If HTMRead(gSL_Select, 1) Then
        B1lb_Syuyo = CisFun.RSetFld(HTM.収容数1, 7, "#,###")
        If HTM.SK区分 = 0 Then
            B1lb_SKKbn = "K"
        Else
            B1lb_SKKbn = "S"
        End If
    End If
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
        Call .FixedSet(中中, 右中, なし, 2, 0, "№")
        Call .FixedSet(中中, 左中, なし, 0, 1, "品番")
        Call .FixedSet(中中, 左中, なし, 0, 2, "取引先CD")
        Call .FixedSet(中中, 左中, なし, 0, 3, "取引先名")
        Call .FixedSet(中中, 中中, なし, 0, 4, "受")
        Call .FixedSet(中中, 中中, なし, 0, 5, "SK")
        Call .FixedSet(中中, 中中, なし, 0, 6, "展開")
        Call .FixedSet(中中, 右中, なし, 0, 7, "収容数")
        Call .FixedSet(中中, 右中, なし, 0, 8, "当月数")
        Call .FixedSet(中中, 右中, なし, 0, 9, "翌月数")
        Call .FixedSet(中中, 右中, なし, 0, 10, "翌々月数")
        Call .FixedSet(中中, 右中, なし, 0, 11, "在庫数")

'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 12, "編集前品番")
        Call .FixedSet(中中, 左中, なし, 0, 13, "日量区分")
        Call .FixedSet(中中, 左中, なし, 0, 14, "展開区分")
        Call .FixedSet(中中, 左中, なし, 0, 15, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11")
        .InitDisp
    End With
    
End Sub
'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    Input_Area.Visible = EnabledFlg
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub

'テキストからグリッドに退避
    Call TextToVsGrid(Val(Input_Area.Tag))
    gInt = VsGrid1.ROW
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > VsGrid1.Rows - 1 Then
            If RTrim(VsGrid1.TextMatrix(gInt, 2)) <> "" Then
                VsGrid1.Rows = gInt + 1 + 1
                VsGrid1.TextMatrix(gInt + 1, 0) = gInt + 1
            Else
                GoTo Grid_RowChange_End
            End If
        End If
        gInt = gInt + 1
    End If
    'PageUp
    If KeyCode = vbKeyPageUp Then
        If gInt - CisVsGrid.DispMax < 1 Then
            gInt = 1
        Else
            gInt = gInt - CisVsGrid.DispMax
        End If
    End If
    'PageDown
    If KeyCode = vbKeyPageDown Then
        If gInt + CisVsGrid.DispMax > VsGrid1.Rows - 1 Then
            gInt = VsGrid1.Rows - 1
        Else
            gInt = gInt + CisVsGrid.DispMax
        End If
    End If
    VsGrid1.ROW = gInt
'グリッドからテキストに表示
    Call VsGridToText(VsGrid1.ROW)
    
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    Call iB1_Torcd_LostFocus
    Call iB1_Ukeir_LostFocus
    With VsGrid1
        .TextMatrix(tRow, 1) = RTrim(iB1_Hinbn)
        .TextMatrix(tRow, 2) = RTrim(iB1_Torcd)
        .TextMatrix(tRow, 3) = RTrim(B1lb_Tornm)
        .TextMatrix(tRow, 4) = RTrim(iB1_Ukeir)
        .TextMatrix(tRow, 5) = RTrim(B1lb_SKKbn)
        .TextMatrix(tRow, 7) = CisFun.RSetFld(B1lb_Syuyo, 7, "#,###")
        .TextMatrix(tRow, 8) = CisFun.RSetFld(iB1_NSur1, 11, "#,###")
        .TextMatrix(tRow, 9) = CisFun.RSetFld(iB1_NSur2, 11, "#,###")
        .TextMatrix(tRow, 10) = CisFun.RSetFld(iB1_NSur3, 11, "#,###")
        .TextMatrix(tRow, 11) = CisFun.RSetFld(iB1_ZSur, 11, "#,###")
        If B1Chk_TenKB.Value = 0 Then
            .TextMatrix(tRow, 6) = ""
        Else
            .TextMatrix(tRow, 6) = "無"
        End If
        
        .TextMatrix(tRow, 13) = B1Chk_TenKB.Value
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    Static wsToText     As Boolean
    
    If wsToText Then Exit Sub
    wsToText = True
    
    With VsGrid1
        Input_Area.Enabled = True
        Input_Area.Tag = tRow
        Input_Area.Move Input_Area.Left, VsGrid1.CellTop, Input_Area.Width, VsGrid1.CellHeight
        
        ' * 品番
        iB1_Hinbn = RTrim(.TextMatrix(tRow, 1))
        ' * 取引先
        iB1_Torcd = RTrim(.TextMatrix(tRow, 2))
        B1lb_Tornm = RTrim(.TextMatrix(tRow, 3))
        ' * 受入
        iB1_Ukeir = RTrim(.TextMatrix(tRow, 4))
        ' * ＳＫ
        B1lb_SKKbn = RTrim(.TextMatrix(tRow, 5))
        ' * 収容数
        B1lb_Syuyo = Format(CisFun.Val2(.TextMatrix(tRow, 7)), "###,###")
        ' * 当月数
        iB1_NSur1 = Format(CisFun.Val2(.TextMatrix(tRow, 8)), "###,###,###")
        ' * 翌月数
        iB1_NSur2 = Format(CisFun.Val2(.TextMatrix(tRow, 9)), "###,###,###")
        ' * 翌々月数
        iB1_NSur3 = Format(CisFun.Val2(.TextMatrix(tRow, 10)), "###,###,###")
        ' * 在庫数
        iB1_ZSur = Format(CisFun.Val2(.TextMatrix(tRow, 11)), "###,###,###")
        ' * 展開区分
        B1Chk_TenKB.Value = CisFun.Val2(.TextMatrix(tRow, 14))
    End With
    
    
    wsToText = False
End Sub
Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call VsGridToText(VsGrid1.ROW)
End Sub
Private Sub VSGrid1_Click()
    Call TextToVsGrid(Val(Input_Area.Tag))
    If VsGrid1.TextMatrix(VsGrid1.ROW, 0) = "" Then Exit Sub
    Call GridTextEnabled(True)
    Call VsGridToText(VsGrid1.ROW)
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
'Private Sub VSGrid_DblClick()
'    If RTrim(vsGrid.TextMatrix(vsGrid.Row, 1)) = "" And RTrim(vsGrid.TextMatrix(vsGrid.Row, 2)) = "" Then
'        Exit Sub
'    End If
'    With CisFun
'        .MB_Lines = 4
'        .MB_MSG(2) = "     " & vsGrid.Row & " 行目をクリアしますか？      "
'        .MB_Title = ""
'        .MB_Button = OK_CAN
'        If Not .MBOX Then Exit Sub
'    End With
'    iB1_Seban = "": iB1_Hinbn = "": B1lb_Syuyo = ""
'    iB1_Maisu = "": iB1_HSuryo = "": B1lb_Hasu = ""
'    B1Chk_RBKbn.Value = 0
'    Call TextToVsGrid(Val(Input_Area.Tag))
'
'End Sub
Private Sub vsGrid_Scroll()
    'スクロール時はテキストボックスを非表示にします。
'    Call GridTextEnabled(False)
End Sub
'==========================================================================
'   年  月  セ  ッ  ト
'==========================================================================
Private Function SetYM()
    SetYM = False

' * 当月稼働日数
    gSL_Select = "SELECT * FROM カレンダマスタ "
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 0 "
    gSL_Select = gSL_Select & "   AND 年 = '" & Mid(iH1_YM, 1, 4) & "'"
    gSL_Select = gSL_Select & "   AND 月 = '" & Mid(iH1_YM, 5, 2) & "'"
    If Not CLMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_Title = ""
            .MB_MSG(2) = "    カレンダマスタ 未登録        "
            .MB_Button = Error
            .MBOX
            iH1_YM.SetFocus
            Exit Function
        
        End With
    End If
    
'    B1lb_Kadsu = Format(CLM.稼動日数, "#0.0")
' 当月
    BLB_MM(1) = Mid(iH1_YM, 1, 4) & "年" & Mid(iH1_YM, 5, 2) & "月"
'    B2Op_MM(1).Caption = Mid(iH1_YM, 1, 4) & "年" & Mid(iH1_YM, 5, 2) & "月"
    wsTNJYM(1) = RTrim(iH1_YM)
' 翌月
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add iH1_YM, M, 1
        gStr = .Date_Result
    End With
    BLB_MM(2) = Mid(gStr, 1, 4) & "年" & Mid(gStr, 5, 2) & "月"
'    B2Op_MM(2).Caption = Mid(gStr, 1, 4) & "年" & Mid(gStr, 5, 2) & "月"
    wsTNJYM(2) = RTrim(gStr)
' 翌々月
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add iH1_YM, M, 2
        gStr = .Date_Result
    End With
    BLB_MM(3) = Mid(gStr, 1, 4) & "年" & Mid(gStr, 5, 2) & "月"
'    B2Op_MM(3).Caption = Mid(gStr, 1, 4) & "年" & Mid(gStr, 5, 2) & "月"
    wsTNJYM(3) = RTrim(gStr)

    SetYM = True
End Function
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridTInit()
    With CisVsGridT
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 2, 0, "№")
        Call .FixedSet(中中, 左中, なし, 0, 1, "レベル")
        Call .FixedSet(中中, 左中, あり, 0, 2, "品   番")
        Call .FixedSet(中中, 左中, なし, 0, 3, "背番号")
        Call .FixedSet(中中, 左中, なし, 0, 4, "手 配 先")
        Call .FixedSet(中中, 中中, なし, 0, 5, "受")
        Call .FixedSet(中中, 左中, なし, 0, 6, "直送先")
        Call .FixedSet(中中, 中中, なし, 0, 7, "SK")
        Call .FixedSet(中中, 右中, なし, 0, 8, "収容数")
        Call .FixedSet(中中, 右中, なし, 0, 9, "当月数")
        Call .FixedSet(中中, 右中, なし, 0, 10, "翌月数")
        Call .FixedSet(中中, 右中, なし, 0, 11, "翌々月数")
        
        Call .FixedSet(中中, 右中, なし, 0, 12, "編集前品番")
        Call .FixedSet(中中, 右中, なし, 0, 13, "編集前手配先")
        Call .FixedSet(中中, 右中, なし, 0, 14, "順位")
            
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 15, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11")
        .InitDisp
    End With
    
End Sub
'+----------------------------------------------------------+
'+   展開内容表示
'+----------------------------------------------------------+
Private Sub Body_Set(Index As Integer)
    
    Message = "【抽出中】しばらくお待ちください"
    PNL_Wait.Visible = True
    
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid2
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGridT.FixedGet(CisFun.Val2(ID))
        Next ID
        If Ttl_Kensu + 1 > .Rows Then
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
        
    gSL_Select = "SELECT SN.*,"
    gSL_Select = gSL_Select & " ISNULL(TR.略称,'') 手配先名, "
    gSL_Select = gSL_Select & " ISNULL(HT.収容数1,0) 収容数, "
    gSL_Select = gSL_Select & " ISNULL(HT.SK区分,9) SK区分,  "
    gSL_Select = gSL_Select & " ISNULL(HT.背番号,'') 背番号,  "
    gSL_Select = gSL_Select & " ISNULL(HN.表示品番,'') 表示品番  "
    gSL_Select = gSL_Select & " FROM S内示問合せワーク SN "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "   ON TR.取引先区分 = 1 "
    gSL_Select = gSL_Select & "  AND TR.取引先CD = SN.手配先 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN "
    gSL_Select = gSL_Select & "   ON HN.品番 = SN.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT "
    gSL_Select = gSL_Select & "   ON HT.品目 = 1 "
    gSL_Select = gSL_Select & "  AND HT.品番 = SN.品番 "
    gSL_Select = gSL_Select & "  AND HT.取引先 = SN.手配先 "
    gSL_Select = gSL_Select & "  AND HT.受入 = SN.受入 "
    gSL_Select = gSL_Select & " WHERE SN.年月 = '" & RTrim(wsTNJYM(Index)) & "'"
    gSL_Select = gSL_Select & " ORDER BY SN.順位 "
    If SNWRead(gSL_Select) Then
    
        Do Until Not SNW_RDSTS
            With vsGrid2
                ID = ID + 1
                .TextMatrix(ID, 0) = Format(ID, "###")
            ' * レベル
                If wsProcKB = "" Then
                    Select Case SNW.レベル
                        Case 0: .TextMatrix(ID, 1) = "0"
                        Case 1: .TextMatrix(ID, 1) = ".1"
                        Case 2: .TextMatrix(ID, 1) = "..2"
                        Case 3: .TextMatrix(ID, 1) = "...3"
                        Case 4: .TextMatrix(ID, 1) = "....4"
                        Case Else: .TextMatrix(ID, 1) = "...." & Format(SNW.レベル, "##")
                    End Select
                End If
            ' * 品番
                If RTrim(SNW.表示品番) = "" Then
                    .TextMatrix(ID, 2) = SNW.品番
                Else
                    .TextMatrix(ID, 2) = SNW.表示品番
                End If
                .TextMatrix(ID, 12) = SNW.品番
            ' * 背番号
                .TextMatrix(ID, 3) = RTrim(SNW.背番号)
            ' * 手配先
                If RTrim(SNW.手配先名) = "" Then
                    .TextMatrix(ID, 4) = SNW.手配先
                Else
                    .TextMatrix(ID, 4) = SNW.手配先 & ":" & RTrim(SNW.手配先名)
                End If
                .TextMatrix(ID, 13) = SNW.手配先
            ' * 受入
                .TextMatrix(ID, 5) = RTrim(SNW.受入)
            ' * 直送先
                .TextMatrix(ID, 6) = RTrim(SNW.直送先)
            ' * SK区分
                Select Case SNW.SK区分
                    Case 0: .TextMatrix(ID, 7) = "Ｋ"
                    Case 1: .TextMatrix(ID, 7) = "Ｓ"
                    Case Else: .TextMatrix(ID, 7) = ""
                End Select
            ' * 収容数
                .TextMatrix(ID, 8) = CisFun.RSetFld(SNW.収容数, 7, "#,###")
            ' * 月量数
                .TextMatrix(ID, 9) = CisFun.RSetFld(SNW.当月数, 10, "#,###")
                .TextMatrix(ID, 10) = CisFun.RSetFld(SNW.翌月数, 10, "#,###")
                .TextMatrix(ID, 11) = CisFun.RSetFld(SNW.翌々月数, 10, "#,###")
            ' * 順位
                .TextMatrix(ID, 14) = SNW.順位
            End With
        
            Call SNWReadNext
        Loop
    End If
    Call SNWClose
    
    PNL_Wait.Visible = False
    vsGrid2.Redraw = True

End Sub
'+-----------------------------------------------
'+  vsGrid2_DblClick
'+-----------------------------------------------
Private Sub vsGrid2_DblClick()
    
    Call Body3Set(vsGrid2.ROW)

End Sub
'*******************************************************************************************************************
'*
'*          明細グリッド処理
'*
'*******************************************************************************************************************
Private Sub Body3Set(ROW As Integer)
    With vsGrid2
        B3lb_No = CisFun.RSetFld(.TextMatrix(ROW, 0), 3, "###")
        B3lb_No.Tag = RTrim(.TextMatrix(ROW, 14))
        B3lb_Hinbn = RTrim(.TextMatrix(ROW, 2))
        B3lb_Hinbn.Tag = RTrim(.TextMatrix(ROW, 12))
        B3lb_Seban = RTrim(.TextMatrix(ROW, 3))
        B3lb_Torcd = RTrim(.TextMatrix(ROW, 4))
        B3lb_Torcd.Tag = RTrim(.TextMatrix(ROW, 13))
        B3lb_Ukeir = RTrim(.TextMatrix(ROW, 5))
        If TorNmGet(RTrim(.TextMatrix(ROW, 6)), 9) Then
            B3lb_NTorcd = RTrim(.TextMatrix(ROW, 6)) & ":" & TRM_RName
        Else
            B3lb_NTorcd = RTrim(.TextMatrix(ROW, 6))
        End If
        B3lb_NTorcd.Tag = RTrim(.TextMatrix(ROW, 6))
        B3lb_SKKbn = .TextMatrix(ROW, 7)
        
    ' * 月量数
        B3lb_NSur1 = CisFun.RSetFld(.TextMatrix(ROW, 9), 11, "#,###")
        B3lb_NSur2 = CisFun.RSetFld(.TextMatrix(ROW, 10), 11, "#,###")
        B3lb_NSur3 = CisFun.RSetFld(.TextMatrix(ROW, 11), 11, "#,###")
        
        
        gSL_Select = "SELECT HT.* FROM 品番取引先マスタ HT"
        gSL_Select = gSL_Select & " WHERE HT.品目 = 1 "
        gSL_Select = gSL_Select & "   AND HT.品番 = '" & RTrim(.TextMatrix(ROW, 12)) & "'"
        gSL_Select = gSL_Select & "   AND HT.取引先 = '" & RTrim(.TextMatrix(ROW, 13)) & "'"
        gSL_Select = gSL_Select & "   AND HT.受入 = '" & RTrim(.TextMatrix(ROW, 5)) & "'"
        If HTMRead(gSL_Select, 1) Then
            ' サイクル
            If RTrim(HTM.サイクル) <> "" Then
                B3lb_Cycle = Mid(HTM.サイクル, 1, 1) & "-" & Mid(HTM.サイクル, 2, 2) & "-" & Mid(HTM.サイクル, 4, 2)
            End If
            ' 収容数1
            If HTM.収容数1 <> 0 Then: B3lb_Syuyo1 = CisFun.RSetFld(HTM.収容数1, 7, "#,###")
            ' ロット1
            If HTM.ロット1 <> 0 Then: B3lb_Lot1 = CisFun.RSetFld(HTM.ロット1, 9, "#,###")
            ' 変更日・便
            gStr = ""
            If RTrim(HTM.変更日) <> "" Then
                gStr = Mid(HTM.変更日, 1, 4) & "/" & Mid(HTM.変更日, 5, 2) & "/" & Mid(HTM.変更日, 7, 2)
                If RTrim(HTM.変更便) <> "" Then
                    gStr = gStr & "-" & HTM.変更便
                End If
            End If
            ' 収容数2
            If HTM.収容数2 <> 0 Then: B3lb_Syuyo2 = CisFun.RSetFld(HTM.収容数2, 7, "#,###")
            ' ロット2
            If HTM.ロット2 <> 0 Then: B3lb_Lot2 = CisFun.RSetFld(HTM.ロット2, 9, "#,###")
            ' 有効期限
            gStr = ""
            If Trim(HTM.有効期限開始) <> "" Then
                gStr = Mid(HTM.有効期限開始, 1, 4) & "/" & _
                       Mid(HTM.有効期限開始, 5, 2) & "/" & _
                       Mid(HTM.有効期限開始, 7, 2)
            End If
            If Trim(HTM.有効期限終了) <> "" Then
               If gStr = "" Then gStr = Space(10)
               gStr = gStr & "～"
               gStr = gStr & Mid(HTM.有効期限終了, 1, 4) & "/" & _
                             Mid(HTM.有効期限終了, 5, 2) & "/" & _
                             Mid(HTM.有効期限終了, 7, 2)
            Else
               If Trim(HTM.有効期限開始) <> "" Then
                  gStr = gStr & "～" & Space(10)
               End If
            End If
            B3lb_Kigen = gStr
            ' LT
            If HTM.LT <> 0 Then: B3lb_LT = CisFun.RSetFld(HTM.LT, 2, "##")
            ' 曜日・日付
            If HTM.Day1 = 0 Then
                BLB_Week = "曜日指定"
                If HTM.Week1 <> 0 Then: B3lb_Week(1) = "月"
                If HTM.Week2 <> 0 Then: B3lb_Week(2) = "火"
                If HTM.Week3 <> 0 Then: B3lb_Week(3) = "水"
                If HTM.Week4 <> 0 Then: B3lb_Week(4) = "木"
                If HTM.Week5 <> 0 Then: B3lb_Week(5) = "金"
            Else
                BLB_Week = "日付指定"
                If HTM.Day1 <> 0 Then: B3lb_Week(1) = CisFun.RSetFld(HTM.Day1, 2, "99")
                If HTM.Day2 <> 0 Then: B3lb_Week(2) = CisFun.RSetFld(HTM.Day2, 2, "99")
                If HTM.Day3 <> 0 Then: B3lb_Week(3) = CisFun.RSetFld(HTM.Day3, 2, "99")
                If HTM.Day4 <> 0 Then: B3lb_Week(4) = CisFun.RSetFld(HTM.Day4, 2, "99")
                If HTM.Day5 <> 0 Then: B3lb_Week(5) = CisFun.RSetFld(HTM.Day5, 2, "99")
            End If
            ' 係数
            If HTM.係数1 <> 0 Then: B3lb_Keisu(1) = CisFun.RSetFld(HTM.係数1, 4, "9.99")
            If HTM.係数2 <> 0 Then: B3lb_Keisu(2) = CisFun.RSetFld(HTM.係数2, 4, "9.99")
            If HTM.係数3 <> 0 Then: B3lb_Keisu(3) = CisFun.RSetFld(HTM.係数3, 4, "9.99")
            ' 平準化
            Select Case HTM.平準化区分
                Case 0: B3lb_HJKbn = "無"
                Case 1: B3lb_HJKbn = "日"
                Case 2: B3lb_HJKbn = "便"
                Case 3: B3lb_HJKbn = "月"
            End Select
            If HTM.平準化係数 <> 0 Then: B3lb_HJRitu = CisFun.RSetFld(HTM.平準化係数, 2, "99") & "%"
        
        End If
    End With

    '+-----------------------+
    '+  グリッド
    '+-----------------------+
    With CisVsGridM
        Set .GridObj = vsGrid3
        .NoSpaceCol = 1
        .NumberSet = All
        .FixedHeight = 360
        .PatanMax = 3
'        .SelectCol = 45
        .InitGet = False
        .Init
    End With
    Call GridMInit
    
    ProcHB = "B3"
    Call DispChange(ProcHB)

End Sub
Private Sub GridMInit()
    Dim SetColor        As Boolean
    
'    mCellSet = True
    
    With vsGrid3
        .Redraw = False
        .Rows = 1: .Cols = 1
        .Rows = 36: .Cols = 9
        
        .ROW = 0
        For gCnt2 = 0 To .Cols - 1
            .Col = gCnt2
'            .ColAlignment(gCnt2) = flexAlignCenterCenter
            .ColAlignment(gCnt2) = flexAlignLeftCenter
            If gCnt2 > 0 Then
                If gCnt2 = 1 Then
                   .ColWidth(gCnt2) = 30
                Else
                   .ColWidth(gCnt2) = 1300
                End If
            End If
        Next gCnt2
        
    ' 明細セット
        Call DaySet
        
        .ROW = 1: .Col = 2
        .Redraw = True
    End With
'    mCellSet = False

End Sub


Private Sub DaySet()
    Dim wDayCol         As Integer
    Dim SetDay          As Integer
    Dim EndDay          As Integer
    Dim wKado           As Boolean
    
    wDayCol = 1
    SetDay = 1
    
    
    Dim wsCL_Row            As Integer
    Dim wsCL_Col            As Integer
    Dim wsCL_NJYM           As String
    Dim wsCL_TORCD          As String
    Dim wsCL_UKEIR          As String
    
    Dim wsCL_YM(3)          As String
    Dim wsCL_KadKB(3)       As String
    Dim wsCL_YakKB(3)       As String
    Dim wsCL_EndDay(3)      As Integer
    Dim wsCL_LastDay        As String
    Dim wsCL_SetDay         As String
    Dim wsCL_StrDay         As String
    
' *--------------------------------------------------------------*
' *     S内示問合せ日量ワーク作成
' *--------------------------------------------------------------*
    With CisDB
        .SQL = "S内示問合せ日量ワーク作成"
        .StoadoCount = 6
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = Input用
        .ParaNo = 3: .ParaIO = Input用
        .ParaNo = 4: .ParaIO = Input用
        .ParaNo = 5: .ParaIO = Input用
        .ParaNo = 6: .ParaIO = Input用
        
        .ParaNo = 1: .ParaValue = RTrim(wsProcKB)
        .ParaNo = 2: .ParaValue = RTrim(iH1_YM) & "01"
        .ParaNo = 3: .ParaValue = RTrim(B3lb_No.Tag)
        .ParaNo = 4: .ParaValue = RTrim(B3lb_Hinbn.Tag)
        .ParaNo = 5: .ParaValue = RTrim(B3lb_Torcd.Tag)
        .ParaNo = 6: .ParaValue = RTrim(B3lb_Ukeir)
    
        If Not .DBStored Then
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(2) = "    内示問合せ日量ワーク作成が失敗しました            "
            CisFun.MB_Title = "内示処理"
            CisFun.MB_Button = Error
            CisFun.MBOX
            Exit Sub
        End If
    End With
' *--------------------------------------------------------------*
' *     曜日セット
' *--------------------------------------------------------------*
    gCnt1 = mWeekStart
    For gInt = 1 To 7
    
        Select Case gCnt1
            Case 1: gStr = "日曜"
            Case 2: gStr = "月曜"
            Case 3: gStr = "火曜"
            Case 4: gStr = "水曜"
            Case 5: gStr = "木曜"
            Case 6: gStr = "金曜"
            Case 7: gStr = "土曜"
        End Select
        B3lb_WeekDay(gInt) = gStr
        gCnt1 = gCnt1 + 1
        If gCnt1 > 7 Then: gCnt1 = 1
    
    Next gInt
' *--------------------------------------------------------------*
' *     カレンダマスタ 退避
' *--------------------------------------------------------------*
    wsCL_NJYM = RTrim(iH1_YM)
    wsCL_TORCD = RTrim(B3lb_Torcd.Tag)
    wsCL_UKEIR = RTrim(B3lb_Ukeir)
' (当月)
    gSL_Select = "SELECT * FROM カレンダマスタ "
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1 "
    gSL_Select = gSL_Select & "   AND 年 = '" & Mid(wsCL_NJYM, 1, 4) & "'"
    gSL_Select = gSL_Select & "   AND 月 = '" & Mid(wsCL_NJYM, 5, 2) & "'"
    gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(wsCL_TORCD) & "'"
    gSL_Select = gSL_Select & "   AND 受入 = '" & RTrim(wsCL_UKEIR) & "'"
    If Not CLMRead(gSL_Select, 1) Then
        wsCL_TORCD = "": wsCL_UKEIR = ""
        
        gSL_Select = "SELECT * FROM カレンダマスタ "
        gSL_Select = gSL_Select & " WHERE 取引先区分 = 1 "
        gSL_Select = gSL_Select & "   AND 年 = '" & Mid(wsCL_NJYM, 1, 4) & "'"
        gSL_Select = gSL_Select & "   AND 月 = '" & Mid(wsCL_NJYM, 5, 2) & "'"
        gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(wsCL_TORCD) & "'"
        gSL_Select = gSL_Select & "   AND 受入 = '" & RTrim(wsCL_UKEIR) & "'"
        If Not CLMRead(gSL_Select, 1) Then: GoTo DaySet_ERR
    End If
    
    wsCL_YM(1) = wsCL_NJYM
    wsCL_EndDay(1) = CLM.末日
    wsCL_KadKB(1) = CLM.稼動区分
    wsCL_YakKB(1) = CLM.夜勤区分
    DoEvents
' (翌月)
    With CisFun
        gStr = wsCL_NJYM & "01"
        .Date_RsFormat = fYMD
        .Date_Add gStr, M, 1, LastDate
        wsCL_NJYM = Mid(.Date_Result, 1, 6)
    End With
    
    gSL_Select = "SELECT * FROM カレンダマスタ "
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1 "
    gSL_Select = gSL_Select & "   AND 年 = '" & Mid(wsCL_NJYM, 1, 4) & "'"
    gSL_Select = gSL_Select & "   AND 月 = '" & Mid(wsCL_NJYM, 5, 2) & "'"
    gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(wsCL_TORCD) & "'"
    gSL_Select = gSL_Select & "   AND 受入 = '" & RTrim(wsCL_UKEIR) & "'"
    If Not CLMRead(gSL_Select, 1) Then: GoTo DaySet_ERR
    wsCL_YM(2) = wsCL_NJYM
    wsCL_EndDay(2) = CLM.末日
    wsCL_KadKB(2) = CLM.稼動区分
    wsCL_YakKB(2) = CLM.夜勤区分
    DoEvents
' (翌々月)
    With CisFun
        gStr = wsCL_NJYM & "01"
        .Date_RsFormat = fYMD
        .Date_Add gStr, M, 1, LastDate
        wsCL_NJYM = Mid(.Date_Result, 1, 6)
        wsCL_LastDay = .Date_Result
    End With
    
    gSL_Select = "SELECT * FROM カレンダマスタ "
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1 "
    gSL_Select = gSL_Select & "   AND 年 = '" & Mid(wsCL_NJYM, 1, 4) & "'"
    gSL_Select = gSL_Select & "   AND 月 = '" & Mid(wsCL_NJYM, 5, 2) & "'"
    gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(wsCL_TORCD) & "'"
    gSL_Select = gSL_Select & "   AND 受入 = '" & RTrim(wsCL_UKEIR) & "'"
    If Not CLMRead(gSL_Select, 1) Then: GoTo DaySet_ERR
    wsCL_YM(3) = wsCL_NJYM
    wsCL_EndDay(3) = CLM.末日
    wsCL_KadKB(3) = CLM.稼動区分
    wsCL_YakKB(3) = CLM.夜勤区分
    DoEvents
    
    wsCL_SetDay = RTrim(iH1_YM) & "01"
    wsCL_StrDay = RTrim(iH1_YM) & "01"

    gSL_Select = "SELECT * FROM S内示問合せ日量ワーク "
    gSL_Select = gSL_Select & " ORDER BY 納期"
    Call SNNRead(gSL_Select)
    wsCL_Row = 1
    With vsGrid3
        Do Until Not SNN_RDSTS
        
            wsCL_SetDay = SNN.納期
            With CisFun
                .DateE_BefVal = wsCL_SetDay
                .DateEdit Slash
                gDate = DateValue(.DateE_AfVal)
            End With
            wsCL_Col = Weekday(gDate, mWeekStart)
            
            If wsCL_Col + 1 = mWeekStart And wsCL_SetDay <> wsCL_StrDay Then
                wsCL_Row = wsCL_Row + 2
            End If
            
            .ROW = wsCL_Row
            .Col = wsCL_Col + 1
            .TextMatrix(wsCL_Row, wsCL_Col + 1) = Mid(wsCL_SetDay, 7, 2) & "日"
            If SNN.数量 <> 0 Then
                .TextMatrix(wsCL_Row + 1, wsCL_Col + 1) = CisFun.RSetFld(SNN.数量, 10, "###0.0000")
            End If
            .TextMatrix(wsCL_Row, 0) = "日 付"
            .TextMatrix(wsCL_Row + 1, 0) = "数 量"
        ' * 当月 *
            If Mid(wsCL_SetDay, 1, 6) = wsCL_YM(1) Then
                If Mid(wsCL_KadKB(1), CisFun.Val2(Mid(wsCL_SetDay, 7, 2)), 1) = "0" And _
                   Mid(wsCL_YakKB(1), CisFun.Val2(Mid(wsCL_SetDay, 7, 2)), 1) = "0" Then
                    .CellBackColor = wsCL_BackColorK
                Else
                    .CellBackColor = wsCL_BackColorY
                End If
            
                .ROW = .ROW + 1
                .CellBackColor = B3lb_NSur1.BackColor
            End If
        ' * 翌月 *
            If Mid(wsCL_SetDay, 1, 6) = wsCL_YM(2) Then
                If Mid(wsCL_KadKB(2), CisFun.Val2(Mid(wsCL_SetDay, 7, 2)), 1) = "0" And _
                   Mid(wsCL_YakKB(2), CisFun.Val2(Mid(wsCL_SetDay, 7, 2)), 1) = "0" Then
                    .CellBackColor = wsCL_BackColorK
                Else
                    .CellBackColor = wsCL_BackColorY
                End If
            
                .ROW = .ROW + 1
                .CellBackColor = B3lb_NSur2.BackColor
            End If
        ' * 翌々月 *
            If Mid(wsCL_SetDay, 1, 6) = wsCL_YM(3) Then
                If Mid(wsCL_KadKB(3), CisFun.Val2(Mid(wsCL_SetDay, 7, 2)), 1) = "0" And _
                   Mid(wsCL_YakKB(3), CisFun.Val2(Mid(wsCL_SetDay, 7, 2)), 1) = "0" Then
                    .CellBackColor = wsCL_BackColorK
                Else
                    .CellBackColor = wsCL_BackColorY
                End If
            
                .ROW = .ROW + 1
                .CellBackColor = B3lb_NSur3.BackColor
            End If
            
            With CisFun
                gStr = wsCL_SetDay
                .Date_RsFormat = fYMD
                .Date_Add gStr, D, 1, ResultDate
                wsCL_SetDay = .Date_Result
            End With
        
            Call SNNReadNext
        Loop
    
        .Rows = wsCL_Row + 2
        DoEvents
    End With


DaySet_ERR:
    
End Sub

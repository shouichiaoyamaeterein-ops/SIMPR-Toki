VERSION 5.00
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CCT0045 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ﾃｷｽﾄﾌｧｲﾙ追加"
   ClientHeight    =   7425
   ClientLeft      =   2025
   ClientTop       =   3090
   ClientWidth     =   9375
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   7425
   ScaleWidth      =   9375
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2700
      Top             =   990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Height          =   795
      Left            =   5220
      ScaleHeight     =   735
      ScaleWidth      =   3765
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   540
      Width           =   3825
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   2850
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
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
         Height          =   585
         Left            =   1920
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1005
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   90
         Top             =   90
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
   End
   Begin VB.TextBox Dummy 
      Height          =   360
      Left            =   30
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   330
      Width           =   180
   End
   Begin VB.PictureBox B1_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5640
      Left            =   270
      ScaleHeight     =   5580
      ScaleWidth      =   8730
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1620
      Width           =   8790
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   5445
         Left            =   60
         Top             =   60
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   9604
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
         Begin Cis3D_v60.CIS3D Back_DeTail 
            Height          =   2625
            Index           =   1
            Left            =   30
            Top             =   2790
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   4630
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
            Begin VB.Frame Frame4 
               BackColor       =   &H00C0C0C0&
               Caption         =   "車体精工コード"
               Height          =   765
               Left            =   180
               TabIndex        =   28
               Top             =   1740
               Visible         =   0   'False
               Width           =   2025
               Begin CisText_V60.CisText iB1_TSKCode 
                  Height          =   360
                  Left            =   660
                  TabIndex        =   29
                  Top             =   300
                  Width           =   630
                  _ExtentX        =   1111
                  _ExtentY        =   635
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXXX"
                  MaxLength       =   4
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00C0C0C0&
               Caption         =   "同一データ存在"
               Height          =   765
               Left            =   4920
               TabIndex        =   25
               Top             =   840
               Width           =   3435
               Begin VB.OptionButton B1Op_ReWrite 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "数量加算"
                  Height          =   240
                  Index           =   1
                  Left            =   1830
                  TabIndex        =   27
                  Top             =   360
                  Width           =   1380
               End
               Begin VB.OptionButton B1Op_ReWrite 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "上書き"
                  Height          =   240
                  Index           =   0
                  Left            =   330
                  TabIndex        =   26
                  Top             =   360
                  Width           =   1185
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D13 
               Height          =   660
               Left            =   7140
               Top             =   30
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   1164
               BackColor       =   14737632
               Caption         =   "自社ｺｰﾄﾞﾁｪｯｸ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
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
               cAlingnment     =   6
               cBoderStyle     =   1
               cPositionY      =   30
               Begin VB.CheckBox B1Ck_Jisya 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "無"
                  Height          =   240
                  Left            =   450
                  TabIndex        =   13
                  Top             =   330
                  Width           =   690
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00C0C0C0&
               Caption         =   "内示ﾃﾞｰﾀ保持期間 "
               Height          =   765
               Left            =   2505
               TabIndex        =   23
               Top             =   840
               Width           =   2355
               Begin CisText_V60.CisText iB1_NaijiDel 
                  Height          =   375
                  Left            =   720
                  TabIndex        =   16
                  Top             =   300
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   661
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
                  Text            =   "##"
                  MaxLength       =   2
                  cDataType       =   1
                  cDataReplace    =   1
                  cFbComma        =   0
                  cFdAutoFormat   =   1
                  cGFormat        =   "##"
                  cILength        =   2
               End
               Begin VB.Label Label1 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "ヶ月"
                  Height          =   255
                  Left            =   1200
                  TabIndex        =   24
                  Top             =   420
                  Width           =   885
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00C0C0C0&
               Caption         =   "内示処理区分 "
               Height          =   765
               Left            =   180
               TabIndex        =   22
               Top             =   840
               Width           =   2265
               Begin VB.OptionButton B1Op_NaijiProc 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "計変"
                  Height          =   435
                  Index           =   1
                  Left            =   1230
                  TabIndex        =   15
                  Top             =   270
                  Width           =   945
               End
               Begin VB.OptionButton B1Op_NaijiProc 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "定常"
                  Height          =   435
                  Index           =   0
                  Left            =   210
                  TabIndex        =   14
                  Top             =   270
                  Width           =   945
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D10 
               Height          =   660
               Left            =   30
               Top             =   30
               Width           =   7095
               _ExtentX        =   12515
               _ExtentY        =   1164
               BackColor       =   14737632
               Caption         =   "TEXT変換元格納先"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
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
               cAlingnment     =   6
               cBoderStyle     =   1
               cPositionY      =   30
               Begin CisText_V60.CisText iB1_TextDir 
                  Height          =   345
                  Left            =   60
                  TabIndex        =   12
                  Top             =   270
                  Width           =   6750
                  _ExtentX        =   11906
                  _ExtentY        =   609
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   $"CCT0045.frx":0000
                  MaxLength       =   250
                  IMEMode         =   2
                  cCaps           =   2
               End
               Begin Cis3D_v60.CIS3D PB_TextDir 
                  Height          =   315
                  Left            =   6690
                  Top             =   270
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   556
                  ForeColor       =   0
                  Caption         =   "…"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   3
                  cAlingnment     =   8
                  cButton         =   -1  'True
               End
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   345
            Left            =   30
            Top             =   2430
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   609
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
            Begin Cis3D_v60.CIS3D Blb_Head 
               Height          =   345
               Left            =   720
               Top             =   0
               Width           =   3675
               _ExtentX        =   6482
               _ExtentY        =   609
               Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
               cFont3D         =   2
               cAlingnment     =   1
               cPositionX      =   60
            End
            Begin Cis3D_v60.CIS3D UC_3D5 
               Height          =   345
               Index           =   0
               Left            =   0
               Top             =   0
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   609
               BackColor       =   16761024
               Caption         =   "ﾍｯﾀﾞｰ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
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
            Begin Cis3D_v60.CIS3D UC_3D5 
               Height          =   345
               Index           =   1
               Left            =   4350
               Top             =   0
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   609
               BackColor       =   16761024
               Caption         =   "日時"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
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
            Begin Cis3D_v60.CIS3D Blb_DateTime 
               Height          =   345
               Left            =   4920
               Top             =   0
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   609
               Caption         =   "XXXXXXXXX1XXXXXXX"
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
               cFont3D         =   2
               cAlingnment     =   7
            End
            Begin Cis3D_v60.CIS3D UC_3D5 
               Height          =   345
               Index           =   2
               Left            =   7110
               Top             =   0
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   609
               BackColor       =   16761024
               Caption         =   "件数"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
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
            Begin Cis3D_v60.CIS3D Blb_Kensu 
               Height          =   345
               Left            =   7710
               Top             =   0
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   609
               Caption         =   "ZZ,ZZZ"
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
               cFont3D         =   2
               cAlingnment     =   4
               cPositionX      =   -30
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D8 
            Height          =   810
            Left            =   5310
            Top             =   840
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   1429
            ForeColor       =   16711680
            Caption         =   "ファイル形式"
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
            cPositionY      =   30
            Begin VB.OptionButton B1Op_Keishiki 
               BackColor       =   &H00C0C0C0&
               Caption         =   "改行あり"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   1
               Left            =   1680
               TabIndex        =   7
               Top             =   360
               Width           =   1275
            End
            Begin VB.OptionButton B1Op_Keishiki 
               BackColor       =   &H00C0C0C0&
               Caption         =   "改行なし"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   0
               Left            =   210
               TabIndex        =   6
               Top             =   360
               Width           =   1275
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   810
            Left            =   4290
            Top             =   840
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   1429
            ForeColor       =   16711680
            Caption         =   "ﾚｺｰﾄﾞ長"
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
            cPositionY      =   30
            Begin CisText_V60.CisText iB1_Leng 
               Height          =   375
               Left            =   180
               TabIndex        =   5
               Top             =   360
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   661
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
               Text            =   "####"
               MaxLength       =   4
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "####"
               cILength        =   4
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   810
            Left            =   30
            Top             =   840
            Width           =   4245
            _ExtentX        =   7488
            _ExtentY        =   1429
            ForeColor       =   16711680
            Caption         =   "変換元ファイル名"
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
            cPositionY      =   30
            Begin VB.CommandButton PB_FileGet 
               Caption         =   "‥"
               Height          =   375
               Left            =   3810
               TabIndex        =   4
               Top             =   360
               Width           =   405
            End
            Begin CisText_V60.CisText iB1_HFileName 
               Height          =   375
               Left            =   120
               TabIndex        =   3
               Top             =   360
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   661
               cWildeCard      =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   30
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   810
            Left            =   5310
            Top             =   30
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   1429
            ForeColor       =   16711680
            Caption         =   "レイアウト種類"
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
            cPositionY      =   30
            Begin VB.ComboBox B1Comb_Layout 
               BackColor       =   &H00C0E0FF&
               Height          =   360
               Left            =   120
               Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
               TabIndex        =   2
               Top             =   330
               Width           =   3015
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   810
            Left            =   30
            Top             =   30
            Width           =   5265
            _ExtentX        =   9287
            _ExtentY        =   1429
            ForeColor       =   16711680
            Caption         =   "業　務　名　（テキストファイル名称)"
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
            cPositionY      =   30
            Begin CisText_V60.CisText iB1_FileName 
               Height          =   375
               Left            =   120
               TabIndex        =   1
               Top             =   360
               Width           =   5085
               _ExtentX        =   8969
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
               Text            =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3XXXXXXXXX4"
               MaxLength       =   40
               IMEMode         =   4
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   780
            Left            =   4290
            Top             =   1650
            Width           =   4275
            _ExtentX        =   7541
            _ExtentY        =   1376
            ForeColor       =   16711680
            Caption         =   "連 続 実 行"
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
            cPositionY      =   30
            Begin VB.OptionButton B1Op_Exec 
               BackColor       =   &H00C0FFC0&
               Caption         =   "無し"
               Height          =   315
               Index           =   0
               Left            =   390
               TabIndex        =   9
               Top             =   360
               Width           =   885
            End
            Begin VB.OptionButton B1Op_Exec 
               BackColor       =   &H00C0C0C0&
               Caption         =   "有り"
               Height          =   315
               Index           =   1
               Left            =   1335
               TabIndex        =   10
               Top             =   360
               Width           =   885
            End
            Begin CisText_V60.CisText iB1_ExecGrp 
               Height          =   375
               Left            =   3510
               TabIndex        =   11
               Top             =   300
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
            Begin VB.Label ExeMei 
               BackColor       =   &H00C0C0C0&
               Caption         =   "実行GRP"
               ForeColor       =   &H00FF0000&
               Height          =   225
               Left            =   2610
               TabIndex        =   21
               Top             =   360
               Width           =   1005
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D11 
            Height          =   780
            Left            =   30
            Top             =   1650
            Width           =   4260
            _ExtentX        =   7514
            _ExtentY        =   1376
            ForeColor       =   16711680
            Caption         =   "契　約　先"
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
            Begin CisText_V60.CisText iB1_KTorcd 
               Height          =   360
               Left            =   60
               TabIndex        =   8
               Top             =   330
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   635
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
            End
            Begin Cis3D_v60.CIS3D B1lb_KTorNm 
               Height          =   360
               Left            =   1050
               Top             =   330
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   635
               BackColor       =   16777152
               Caption         =   "XXXXXXXXX1XXXX"
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
               cFont3DColor1   =   8421504
               cFont3DColor2   =   -2147483643
               cFont3D         =   2
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   300
      ScaleHeight     =   900
      ScaleWidth      =   1980
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   540
      Width           =   2040
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   810
         Left            =   90
         Top             =   60
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1429
         ForeColor       =   16711680
         Caption         =   "ﾌｧｲﾙｺｰﾄﾞ"
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
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_FileCd 
            Height          =   375
            Left            =   210
            TabIndex        =   0
            Top             =   375
            Width           =   1380
            _ExtentX        =   2434
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
            Text            =   "XXXXXXXXXX"
            MaxLength       =   10
         End
      End
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 テキストファイル詳細設定  】"
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
         Left            =   8340
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
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   16711935
         Caption         =   "XXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
      End
   End
End
Attribute VB_Name = "CCT0045"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   ﾃｷｽﾄﾌｧｲﾙ詳細設定
'**       フォームID    :   CCT0045
'**       処理概要      :
'**
'**       作  成  日    :   2003/12/13  By CIS
'**       変  更  日    :   2004/11/08  By CIS 財務ﾃﾞｰﾀ対応(売上照合)
'**       変  更  日    :   2005/06/20  By CIS 不具合修正(自社ｺｰﾄﾞﾁｪｯｸの画面表示)
'**                         2005/12/12  By CIS FCODE
'**       変  更  日    :   2008/04/16  By CIS 同一データ存在時の処理区分追加(ｵﾌﾟｼｮﾝ4使用)
'**       変  更  日    :   2008/05/08  By CIS 車体精工対応
'**********************************************************'

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替

    Dim CisFrm          As CISFormContorl
    Dim CisComb         As CisCombo
    
    Dim mGotColor       As Long
    Dim mLostColor      As Long
   
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
    
    Dummy.Enabled = False
    Dummy.Left = -1000
    
    mGotColor = B1Op_Exec(0).BackColor
    mLostColor = B1Op_Exec(1).BackColor
   
    B1Op_NaijiProc(0).Caption = CT0040_Naiji
    CT0040_RTcd = False
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'    '+---------------------+
'    '+ ｺﾝﾎﾞﾎﾞｯｸｽ 初期設定  +
'    '+---------------------+
    Set CisComb = New CisCombo
    With CisComb
        Set .Connect = CisDB
        Set .CombObj = B1Comb_Layout
        .Control = "名称マスタ"
        .ComboName = "レイアウト情報内示"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .Widht = 250
        .ComboInit
    End With
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(False, x1024y768)

End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    FormAct = Not FormAct
    With CisFun
        '個別
        .INI_FileID = gIniExe & gwsIniName
    End With
    
    If CT0040_Proc = "ADD" Then
        K_Sykbnm = "追加"
        K_Sykbnm.ForeColor = &HFFFF&   '(黄色)
        iH1_FileCd = CT0040_FileCd
        iB1_FileName = CT0040_GyoName
        iB1_HFileName = iH1_FileCd & ".txt"
        Call ReturnPress
    Else
        K_Sykbnm = "修正"
        K_Sykbnm.ForeColor = &HFF0000  '(青色)
        iH1_FileCd = CT0040_FileCd
        Call ReturnPress
    End If
        
        
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
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
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
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    Me.Hide
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    Call PB_END_Click
'    If CT0040_Proc = "MNT" Then
'       Call PB_END_Click
'       Exit Sub
'    End If
'
'    If ProcHB = "B1" Then
'       ProcHB = "H1"
'       Call DispChange(ProcHB)
'       Call HeadBodyClear("B")
'    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iB1_KTorcd" Then Call Look_KTorcd: Exit Sub
End Sub
'+----------------------------+
'+    変換元ファイル名検索     +
'+----------------------------+
Private Sub PB_FileGet_Click()
     On Error Resume Next
     CommonDialog1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly
     CommonDialog1.Filter = "テキスト ファイル (*.txt)|*.txt|ＡＰＬ ファイル (*.apl)|*.apl|全てのファイル (*.*)|*.*"
     With CisFun
        .INI_Section = "TextConvert"
        .INI_Default = gIniExe & "Text\"
        .INI_Key = "InputDir"
        .GetIni
        ' 変換元フォルダ
        If Trim(iB1_HFileName) = "" Then
           CommonDialog1.FileName = Trim(.INI_String) & iH1_FileCd & ".txt"
        Else
           CommonDialog1.FileName = Trim(.INI_String) & iB1_HFileName
        End If
     End With
    
    ' 以下のｿｰｽｺｰﾄﾞは、<ｷｬﾝｾﾙ>ﾎﾞﾀﾝがｸﾘｯｸされたときに実行されます。
     CommonDialog1.CancelError = True
     CommonDialog1.ShowOpen
     If Err = 32755 Then
    '<ｷｬﾝｾﾙ>ﾎﾞﾀﾝが押されたときにｴﾗｰｺｰﾄﾞを返すようにする設定です。
    '必要なｿｰｽｺｰﾄﾞを入力してください。
     End If
     If Trim(CommonDialog1.FileName) <> "" Then
        With CisFun
            .FE_FullPath = CommonDialog1.FileName
            .GetPathName
            iB1_HFileName = .FE_File
        End With
     End If
     On Error GoTo 0
End Sub

'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
   Call ReturnPress
End Sub
'+----------------------------+
'+       契約先検索(得意先)    +
'+----------------------------+
Private Sub Look_KTorcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 2
    RV_Call = "Z"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iB1_KTorcd = RV_TorcdK
        B1lb_KTorNm = RV_TorRName
    End If
End Sub
'+-----------------------------+
'+  B1Comb_Layout　レイアウト種類
'+-----------------------------+
Private Sub B1Comb_Layout_GotFocus()
    B1Comb_Layout.BackColor = gIGotSel
End Sub
Private Sub B1Comb_Layout_LostFocus()
    B1Comb_Layout.BackColor = gILostSel
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT START
'+------------------------------------+
'+     レイアウト種類変更時の処理     +
'+------------------------------------+
    Call CisComb.CodeGet
    If CisComb.Code = 6 Then
       Frame4.Visible = True
    Else
       Frame4.Visible = False
       iB1_TSKCode = ""
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT END
End Sub
Private Sub B1Comb_Layout_Click()
    Call LayDisp
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT START
'+------------------------------------+
'+     レイアウト種類変更時の処理     +
'+------------------------------------+
    Call CisComb.CodeGet
    If CisComb.Code = 6 Then
       Frame4.Visible = True
    Else
       Frame4.Visible = False
       iB1_TSKCode = ""
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT END
End Sub
'+-----------------------------+
'+  B1Op_Keishiki ファイル形式
'+-----------------------------+
Private Sub B1Op_Keishiki_Click(Index As Integer)
    
    B1Op_Keishiki(0).Tag = Index
    For Each gObj In B1Op_Keishiki
        If gObj.Index = Index Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
'+-----------------------------+
'+  契約先
'+-----------------------------+
Private Sub iB1_KTorcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_KTorcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    B1lb_KTorNm = ""
    If TorNmGet(iB1_KTorcd, 2) Then
        B1lb_KTorNm = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  B1Op_Exec 連続実行
'+-----------------------------+
Private Sub B1Op_Exec_Click(Index As Integer)
    
    B1Op_Exec(0).Tag = Index
    For Each gObj In B1Op_Exec
        If gObj.Index = Index Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
    If Index = 0 Then
       iB1_ExecGrp.Visible = False
       ExeMei.Visible = False
    Else
       iB1_ExecGrp.Visible = True
       ExeMei.Visible = True
    End If
End Sub

'+-----------------------------+
'+  B1Op_NaijiProc 内示処理区分
'+-----------------------------+
Private Sub B1Op_NaijiProc_Click(Index As Integer)
    
    B1Op_NaijiProc(0).Tag = Index
    For Each gObj In B1Op_NaijiProc
        If gObj.Index = Index Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
'[][][][][][][][][][] 2008/04/16 Add Start [][][][][][][][][][]
'+-----------------------------+
'+  B1Op_ReWrite 同一ﾃﾞｰﾀ存在時処理
'+-----------------------------+
Private Sub B1Op_ReWrite_Click(Index As Integer)
    
    B1Op_ReWrite(0).Tag = Index
    For Each gObj In B1Op_ReWrite
        If gObj.Index = Index Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
'[][][][][][][][][][] 2008/04/16 Add End [][][][][][][][][][]

'+-------------------------------+
'+   TEXT変換元格納先　参照　    +
'+-------------------------------+
Private Sub PB_TextDir_Click()
    With CisFun
        .Folder_Title = "TEXT変換元格納先参照"
        .Folder_DefDir = iB1_TextDir
        If Trim(iB1_TextDir) = "" Then
           .INI_Section = "TextConvert"
           .INI_Default = gIniExe & "Text\"
           .INI_Key = "InputDir"
           .GetIni
           .Folder_DefDir = Trim(.INI_String)
        End If
        .FolderSel
        
        If RTrim$(.Folder_Result) <> "" Then iB1_TextDir = RTrim$(.Folder_Result)
    End With
    iB1_TextDir.SetFocus
End Sub
'+-----------------------------------+
'+  自社ｺｰﾄﾞﾁｪｯｸ
'+-----------------------------------+
Private Sub B1Ck_Jisya_Click()
    If B1Ck_Jisya.Value = 0 Then
       B1Ck_Jisya.Caption = "無"
    Else
       B1Ck_Jisya.Caption = "有"
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
    Dim H1Mode      As Boolean
    Dim H1Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
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
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    
    PB_CAN.Visible = B1Mode
    If CT0040_Proc = "MNT" Then
       PB_CAN.Visible = False
    End If
    PB_CAN.Visible = False
    
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
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
                .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
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
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call PB_END_Click
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
    
    If Trim(iH1_FileCd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ファイルコードを入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       Exit Function
    End If

' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = " SELECT * "
    gSL_Select = gSL_Select & " FROM TEXT管理テーブル "
    gSL_Select = gSL_Select & " WHERE FILECD = '" & RTrim(iH1_FileCd) & "'"
    If TKTRead(gSL_Select, 1) Then
       If CT0040_Proc = "ADD" Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    ファイルコード登録済            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          Exit Function
       End If
    Else
       If CT0040_Proc <> "ADD" Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    ファイルコード未登録            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          Call PB_END_Click
          Exit Function
       End If
    End If
    
    Call BodySet
    
    Head1Chk = True
    
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
    If Trim(iB1_FileName) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    業務名(テキストファイル名称)を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_FileName.SetFocus
       Exit Function
    End If
    If Trim(iB1_HFileName) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    変換元ファイル名を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_HFileName.SetFocus
       Exit Function
    End If
    If B1Op_Keishiki(0).Value And CisFun.Val2(iB1_Leng) = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    レコード長を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_Leng.SetFocus
       Exit Function
    End If
    B1lb_KTorNm = ""
    If iB1_KTorcd = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    契約先を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_KTorcd.SetFocus
       Exit Function
    End If
    
    If TorNmGet(iB1_KTorcd, 2) Then
        B1lb_KTorNm = TRM_RName
    Else
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    契約先未登録（得意先）            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_KTorcd.SetFocus
       Exit Function
    End If
    
    Body1Chk = True
    
End Function
'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Sub BodySet()
    
    Call CisComb.ComboSet
    
    Call LayDisp
    
    If CT0040_Proc = "ADD" Then
       B1Op_Keishiki(0).Value = True
       B1Op_Exec(0).Value = True
       Exit Sub
    End If
    
    With TKT
        iB1_FileName = .F名称
        B1Op_Keishiki(.オプション1).Value = True
        B1Op_Exec(.連続実行).Value = True
        iB1_ExecGrp = .実行GRP
        iB1_HFileName = .変換元ファイル名
        iB1_Leng = CisFun.RSetFld(.有効長, 0, iB1_Leng.cFormat)
        With CisComb
            .Code = TKT.TEXTレイアウト
            .TextGet
        End With
        Blb_Head = .最終処理
        If Format(.最終処理日時, "YYYYMMDD") > "20000101" Then
            Blb_DateTime = Format(.最終処理日時, "YY/MM/DD hh:nn:ss")
            Blb_Kensu = Format(.処理件数, "##,###")
        End If
        iB1_KTorcd = .契約先
        Call iB1_KTorcd_LostFocus
        
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT START
        If .TEXTレイアウト = 6 Then
           iB1_TSKCode = .車体精工コード
        End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT END
        
    End With
    With CisFun
        '個別
        .INI_Section = "TextConvert"
        .INI_Key = "InputDir_" & Trim(iH1_FileCd)
        .GetIni
        If Trim(.INI_String) <> "" Then
           iB1_TextDir = RTrim(.INI_String)
        End If                                                                  '05.06.20 Add
        .INI_Key = "Jisya_" & Trim(iH1_FileCd)
        .GetIni
        If UCase(.INI_String) = "Y" Or UCase(.INI_String) = "YES" Then
           B1Ck_Jisya.Value = 1
        Else
           B1Ck_Jisya.Value = 0
        End If
'        End If                                                                 '05.06.20 Del
    End With
End Sub
'**************************************
'*       レイアウト別内容表示          *
'**************************************
Private Sub LayDisp()
    Dim wType       As Long
' レイアウト種類より　編集タイプを求める
    CisComb.CodeGet
    
    gSL_Select = "Select 数字1 From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = 'レイアウト情報内示'"
    gSL_Select = gSL_Select & "   And 値 = '" & CisComb.Code & "'"
    If Not SYMRead(gSL_Select, 1) Then
       Call ItemsClearSYM
    End If
    
    wType = SYM.数字1
    If wType = 0 Then
        For Each gObj In Back_DeTail
            gObj.Visible = False
        Next gObj
        Exit Sub
    End If
    
    For Each gObj In Back_DeTail
        If wType = gObj.Index Then
           gObj.Visible = True
        Else
           gObj.Visible = False
        End If
    Next gObj
    ' 内示
    If wType = 1 Then
       If CT0040_Proc = "ADD" Then
          B1Op_NaijiProc(0).Value = True
          iB1_NaijiDel = " 3"
          B1Op_ReWrite(0).Value = True                      '2008/04/16
       Else
          B1Op_NaijiProc(TKT.オプション3).Value = True
          iB1_NaijiDel = CisFun.RSetFld(TKT.オプション2, 0, iB1_NaijiDel.cFormat)
          B1Op_ReWrite(TKT.オプション4).Value = True        '2008/04/16
       End If
       Exit Sub
    End If
End Sub

'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Dim AddFlg  As Boolean

    With CisFun
        .MB_Lines = 4
        If CT0040_Proc = "ADD" Then
            .MB_MSG(2) = "     追加処理を行います      "
            .MB_Title = "追加処理"
        Else
            .MB_MSG(2) = "     修正処理を行います      "
            .MB_Title = "修正処理"
        End If
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err
'   マスタ項目移送
    gSL_Select = " SELECT * "
    gSL_Select = gSL_Select & " FROM TEXT管理テーブル "
    gSL_Select = gSL_Select & " WHERE FILECD = '" & RTrim(iH1_FileCd) & "'"
    If TKTRead(gSL_Select, 1) Then
       AddFlg = False
    Else
       Call ItemsClearTKT
       AddFlg = True
       TKT.FILECD = RTrim(iH1_FileCd)
    End If
    With TKT
        .F名称 = RTrim(iB1_FileName)
        CisComb.CodeGet
        .TEXTレイアウト = CisFun.Val2(CisComb.Code)
        .レイアウト名 = RTrim(CisComb.CombText)
        .変換元ファイル名 = RTrim(iB1_HFileName)
        .有効長 = CisFun.Val2(iB1_Leng)
        .オプション1 = CisFun.Val2(B1Op_Keishiki(0).Tag)
        .契約先 = iB1_KTorcd
        .連続実行 = CisFun.Val2(B1Op_Exec(0).Tag)
        If .連続実行 = 0 Then
           .実行GRP = ""
        Else
           .実行GRP = iB1_ExecGrp
        End If
        ' 内示
        If Back_DeTail(1).Visible Then
           .オプション2 = CisFun.Val2(iB1_NaijiDel)
           .オプション3 = CisFun.Val2(B1Op_NaijiProc(0).Tag)
           .オプション4 = CisFun.Val2(B1Op_ReWrite(0).Tag)          '2008/04/16
        End If

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT START
        .車体精工コード = RTrim(iB1_TSKCode)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/08 INSERT END
        
'        .FCODE = Right(.FILECD, 3)     '2005/12/12
        If Len(.FILECD) < 8 Then
           .FCODE = Right(.FILECD, 3)
        Else
           .FCODE = Right(.FILECD, 4)
        End If
        
        .TNSFMT = 1
    End With

    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 追加処理 )
    If AddFlg Then
        TKT.作成者 = gTanto
        Call TKTInsert
    Else
' ( 更新処理 )
        TKT.更新者 = gTanto
        Call TKTUpdate
    End If

    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    With CisFun
        .INI_Section = "TextConvert"
        .INI_Key = "InputDir_" & Trim(iH1_FileCd)
        .INI_String = Trim(iB1_TextDir)
        .PutIni
        
        .INI_Key = "Jisya_" & Trim(iH1_FileCd)
        If B1Ck_Jisya.Value = 1 Then
           .INI_String = "Yes"
        Else
           .INI_String = "No"
        End If
        .PutIni
    End With
    DBPut = True
    CT0040_RTcd = True
    
    Exit Function
DBPut_Err:
    CisDB.DBTran (TransRollback) ' ﾄﾗﾝｻﾞｸｼｮﾝ 破棄
    
    CisFun.ErrorBox
    End
End Function


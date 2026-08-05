VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "Cis3D_6.0.ocx"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CisText_6.0.ocx"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CisYM_6.0.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CisYMDwB_6.0.ocx"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CisBtn_6.0.ocx"
Begin VB.Form CTA0550 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "売上年月変更処理"
   ClientHeight    =   11235
   ClientLeft      =   7755
   ClientTop       =   1665
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   Begin VB.PictureBox B2_Area1 
      BackColor       =   &H00FF0000&
      Height          =   4500
      Left            =   1860
      ScaleHeight     =   4440
      ScaleWidth      =   11385
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4770
      Width           =   11445
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   4305
         Left            =   90
         Top             =   60
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   7594
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
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   4170
            Index           =   1
            Left            =   105
            Top             =   75
            Width           =   10965
            _ExtentX        =   19341
            _ExtentY        =   7355
            BackColor       =   8421504
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cBoderStyle     =   1
            Begin vsFlexLib.vsFlexArray vsGrid1 
               Height          =   4095
               Left            =   30
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   30
               Width           =   10960
               _Version        =   196608
               _ExtentX        =   19341
               _ExtentY        =   7223
               _StockProps     =   228
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ConvInfo        =   1418783674
               FormatString    =   "No|処理日       |担当者　　　　 　　|置場           |連番　　|品番　　　　 　　　　|実棚数　|  |"
               Rows            =   14
               BackColor       =   12648447
               BackColorBkg    =   12632256
               FocusRect       =   0
               HighLight       =   2
               ScrollBars      =   2
               RowHeightMin    =   300
               SelectionMode   =   1
               AllowUserResizing=   1
               BackColorAlternate=   12648447
            End
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   405
      Left            =   1860
      ScaleHeight     =   345
      ScaleWidth      =   2010
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2670
      Width           =   2070
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   345
         Index           =   1
         Left            =   15
         Top             =   0
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   609
         Caption         =   " 表示件数"
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
         cAlingnment     =   1
         Begin Cis3D_v60.CIS3D Dsp_Kensu 
            Height          =   270
            Left            =   1095
            Top             =   30
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   476
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -60
         End
      End
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   825
      Left            =   7110
      ScaleHeight     =   765
      ScaleWidth      =   6060
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1560
      Width           =   6120
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   3465
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
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   4320
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
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   5175
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
      Begin CisBtn_60.CisBtn PB_LOOK 
         Height          =   585
         Left            =   2610
         Top             =   90
         Visible         =   0   'False
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1755
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
         BtnType         =   4
         BtnCaption      =   "削除"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   255
         BtnShadowColor  =   16777215
         PFCaption       =   "( F3 )"
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
         PFBackColor     =   255
      End
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   45
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
         BtnCaption      =   "追加"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   65535
         BtnShadowColor  =   0
         PFCaption       =   "( F1 )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   0
         PFShadowColor   =   16777215
         PFBackColor     =   65535
      End
      Begin CisBtn_60.CisBtn PB_MNT 
         Height          =   585
         Left            =   900
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
         BtnType         =   8
         BtnCaption      =   "修正"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   16776960
         BtnShadowColor  =   0
         PFCaption       =   "( F2 )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   0
         PFShadowColor   =   16777215
         PFBackColor     =   16776960
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   1455
      Left            =   1860
      ScaleHeight     =   1395
      ScaleWidth      =   11385
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3240
      Width           =   11445
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   1260
         Left            =   90
         Top             =   60
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   2223
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
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   990
            Index           =   0
            Left            =   6900
            Top             =   120
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   1746
            ForeColor       =   16711680
            Caption         =   "連番"
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
            cPositionY      =   60
            Begin CisText_V60.CisText iB1_Renbn 
               Height          =   360
               Left            =   60
               TabIndex        =   4
               Top             =   465
               Width           =   810
               _ExtentX        =   1429
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
               Text            =   "XXXXX"
               MaxLength       =   5
            End
         End
         Begin Cis3D_v60.CIS3D Back_Torcd 
            Height          =   990
            Index           =   0
            Left            =   4620
            Top             =   120
            Width           =   2280
            _ExtentX        =   4022
            _ExtentY        =   1746
            ForeColor       =   16711680
            Caption         =   "置　場"
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
            cPositionY      =   60
            Begin CisText_V60.CisText iB1_Okiba 
               Height          =   360
               Left            =   90
               TabIndex        =   3
               Top             =   480
               Width           =   2055
               _ExtentX        =   3625
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
               Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   30
               IMEMode         =   1
            End
         End
         Begin Cis3D_v60.CIS3D Back_Bin 
            Height          =   990
            Left            =   90
            Top             =   120
            Width           =   2250
            _ExtentX        =   3969
            _ExtentY        =   1746
            ForeColor       =   128
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   3
            cPositionX      =   60
            cPositionY      =   60
            Begin CisYMD_With_Btn.CisYMDwB iB1_Ymd 
               Height          =   345
               Left            =   135
               TabIndex        =   1
               Top             =   480
               Width           =   1950
               _ExtentX        =   2990
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
               Object.Height          =   345
               Object.Width           =   1950
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
               Index           =   0
               Left            =   690
               Top             =   60
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               ForeColor       =   128
               Caption         =   "処理日"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   12632256
               cBoderColor2    =   12632256
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderStyle     =   2
               cPositionY      =   10
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   990
            Index           =   3
            Left            =   2355
            Top             =   120
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   1746
            ForeColor       =   128
            Caption         =   "担当者"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   60
            Begin CisText_V60.CisText iB1_Tanto 
               Height          =   360
               Left            =   75
               TabIndex        =   2
               Top             =   480
               Width           =   510
               _ExtentX        =   900
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
            Begin Cis3D_v60.CIS3D B1lb_Tanto 
               Height          =   360
               Left            =   570
               Top             =   480
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   635
               BackColor       =   16777152
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
               cAlingnment     =   1
               cBoderStyle     =   1
               cPositionX      =   30
            End
         End
         Begin Cis3D_v60.CIS3D Back_Torcd 
            Height          =   990
            Index           =   1
            Left            =   7830
            Top             =   120
            Width           =   2280
            _ExtentX        =   4022
            _ExtentY        =   1746
            ForeColor       =   16711680
            Caption         =   "品　番"
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
            cPositionY      =   60
            Begin CisText_V60.CisText iB1_Hinbn 
               Height          =   360
               Left            =   90
               TabIndex        =   5
               Top             =   480
               Width           =   2055
               _ExtentX        =   3625
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
               Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   30
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   990
            Index           =   1
            Left            =   10110
            Top             =   120
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   1746
            ForeColor       =   16711680
            Caption         =   "実棚数"
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
            cPositionY      =   60
            Begin CisText_V60.CisText iB1_Suryo 
               Height          =   360
               Left            =   60
               TabIndex        =   6
               Top             =   465
               Width           =   810
               _ExtentX        =   1429
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
               Text            =   "XXXXX"
               MaxLength       =   5
               cAlignment      =   1
            End
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   915
      Left            =   1830
      ScaleHeight     =   889.899
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   4260
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1530
      Width           =   4320
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   645
         Index           =   1
         Left            =   90
         Top             =   86
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   1138
         ForeColor       =   128
         Caption         =   "棚卸年月"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   240
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   390
            Left            =   2400
            TabIndex        =   0
            Top             =   135
            Width           =   1440
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
            cAutoEject      =   0   'False
         End
      End
   End
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 棚卸年月入力 】"
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
         Width           =   825
         _ExtentX        =   1455
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
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   345
         Left            =   0
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
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
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   705
      Top             =   -120
   End
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   -150
   End
   Begin VB.TextBox Dummy 
      Height          =   375
      Left            =   -15
      TabIndex        =   7
      Text            =   "Dummy"
      Top             =   -105
      Width           =   345
   End
   Begin VB.Label BG_Logo 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  '透明
      Caption         =   "SIMPR-A"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   390
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "CTA0550"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  棚卸データメンテナンス
'**       フォームID    :  CTA0550
'**       処理概要      :
'**
'**       作  成  日    :  2019/09/10   By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl

    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim CisVsGrid       As New CisVsGrid3

''''Dim mJiseki         As String

    '===================================================================' 2019/09/06 ADD START
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除)
    Dim RowColFlg       As Boolean
    Dim SaveRow         As Integer
    Dim SelectedRow     As Long
    '===================================================================' 2019/09/06 ADD END

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

    '+--------------------------------+
    '+  グリッド初期設定
    '+--------------------------------+
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .SelectCol = 8
        .InitGet = False
        .Init
    End With
    Call GridInit

'   < シングルプロセスチェック >
    'If Not SingleProcess_Check Then PB_END_Click

'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = ""
    Call SyoriName(SyoriKB)
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    If gLogoDisp Then
       BG_Logo(0).Caption = gLogo
       CisFrm.BGLDisp BG_Logo, 10
    End If
    iH1_YM = Format(Now(), "yyyyMM")        ' 2019/09/09 ADD
    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 1000   ' 時刻表示間隔
    Timer2.Enabled = True    ' 時刻表示開始

    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Cancel = 1
    Call PB_END_Click
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
    Key_Acc = False
    If Not FormAct Then Exit Function
    If UCase(IVENT) <> "DOWN" Then Exit Function

    Select Case Key_Code
        Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click        '【追加】       ' 2019/09/09 ADD
        Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click        '【修正】       ' 2019/09/09 ADD
        Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click        '【削除】
        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click      '【検索】
        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click        '【取消】
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click        '【終了】
        Case vbKeyReturn
                            If PB_ENT.Visible Then                          '【入力】
                                Call PB_ENT_Click
                            End If
        Case Else:          Exit Function
    End Select

    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      フォーム    《Ｐａｉｎｔ イベント》                 +
'+----------------------------------------------------------+
Private Sub Form_Paint()
'ｸﾞﾗﾃﾞｰｼｮﾝ表示
    With CisFrm
        .Gradation_Direction = Portrait
        .Gradation_Shading = Lighten
        .Gradation 250, 100, 100, 2, 5, 10
    End With
End Sub

'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    
    Call SingleProcess_Release
    
    Call CisDB.DBDISConnect
    Call CisFrm.UnLoadDisp
    End
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    Call CisFrm.MousePT(11)
   
    If ProcHB = "B2" Then
       Call HeadBodyClear("B1")
       Call CisVsGrid.Clear
       SyoriKB = ""
       Call SyoriName(SyoriKB)
       ProcHB = "H1"
       Call DispChange(ProcHB)
       GoTo PB_CAN_Ed
    End If
    
    If ProcHB = "B1" Then
       Call HeadBodyClear("B1")
       
        If SyoriKB = "ADD" Then
'            iB1_Ymd.SetFocus
'            If CisFun.Val2(Dsp_Kensu) = 0 Then
                Call CisVsGrid.Clear
                SyoriKB = ""
                Call SyoriName(SyoriKB)
                ProcHB = "H1"
                Call DispChange(ProcHB)
'            End If
            GoTo PB_CAN_Ed
        End If
       
       ProcHB = "B2"
       Call DispChange(ProcHB)
       Call SyoriName(SyoriKB)
       
       
Call CisVsGrid.vsColor(消去個別)
vsGrid1.SetFocus
       
       GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)    +
'+----------------------------+
Private Sub PB_ENT_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iB1_Ymd" Then iB1_Ymd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iB1_Hinbn" Then Call Look_Hinbn: Exit Sub
End Sub
'===========================================' 2019/09/06 ADD START
'+----------------------------+
'+    新　規  キ ー(F1)        +
'+----------------------------+
Private Sub PB_ADD_Click()
    If SyoriKB <> "ADD" Then
        ProcHB = "B1"
        SyoriKB = "ADD"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        Call HeadBodyClear("B", "B1")
    End If
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)        +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKB <> "MNT" Then
'        If SyoriKB = "ADD" Or SyoriKB = "" Then
'            vsGrid1.Row = 1
'        End If
        Call CisVsGrid.vsColor(消去個別)
        ProcHB = "B2"
        SyoriKB = "MNT"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        Call GridToText
        vsGrid1.SetFocus
       'SendKeys "{TAB}"
    End If
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)        +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
Call CisVsGrid.vsColor(消去個別)
        ProcHB = "B2"
        SyoriKB = "DEL"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
       'SendKeys "{TAB}"
Call GridToText
vsGrid1.SetFocus
    End If
End Sub
'===========================================' 2019/09/06 ADD END



'+----------------------------+
'+       品番検索            +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 0
    RVI_SK = 9
    RVI_Torcd = ""
    RVI_Ukeir = ""
    RVI_Seban = ""
    RVI_Tehai = ""

    CKK0030.Show vbModal
    Unload CKK0030
    Set CKK0030 = Nothing

    If RV_Rtn Then
       iB1_Hinbn = RTrim(RV_Hinbn)
    End If
End Sub
'+-----------------------------+
'+  iB1_Ymd 処理日
'+-----------------------------+
Private Sub iB1_Ymd_GotFocus()
    PB_Look.Visible = True                  ' 2019/09/09 UPD：順序入れ替え
    PB_Look.Tag = ActiveControl.Name
End Sub
Private Sub iB1_Ymd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Tanto   担当者
'+-----------------------------+
Private Sub iB1_Tanto_GotFocus()
    PB_Look.Visible = True                  ' 2019/09/09 ADD
    PB_Look.Tag = ""                        ' 2019/09/09 UPD：ActiveControl.Name
    PB_Look.Visible = False                 ' 2019/09/09 UPD：True
End Sub
Private Sub iB1_Tanto_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    If SyoriKB = "MNT" And Trim(iB1_Tanto.Text) = "" And _
       vsGrid1.TextMatrix(SelectedRow, CisVsGrid.FixedGet(7)) = "" Then
        B1lb_Tanto = vsGrid1.TextMatrix(SelectedRow, CisVsGrid.FixedGet(2))
    Else
        B1lb_Tanto = GetTanto(iB1_Tanto)
    End If
End Sub
'+-----------------------------+
'+  iB1_Okiba   置場
'+-----------------------------+
Private Sub iB1_Okiba_GotFocus()
    PB_Look.Visible = True                  ' 2019/09/09 ADD
    PB_Look.Tag = ""                        ' 2019/09/09 UPD：ActiveControl.Name
    PB_Look.Visible = False                 ' 2019/09/09 UPD：True
End Sub
Private Sub iB1_Okiba_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Hinbn   品番
'+-----------------------------+
Private Sub iB1_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Renbn   連番
'+-----------------------------+
Private Sub iB1_Renbn_GotFocus()
    PB_Look.Visible = True                  ' 2019/09/09 ADD
    PB_Look.Tag = ""                        ' 2019/09/09 UPD：ActiveControl.Name
    PB_Look.Visible = False                 ' 2019/09/09 UPD：True
End Sub
Private Sub iB1_Renbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    iB1_Renbn = Format(CisFun.Val2(iB1_Renbn), "#")
End Sub
'+-----------------------------+
'+  iB1_Suryo   実棚数
'+-----------------------------+
Private Sub iB1_Suryo_GotFocus()
    PB_Look.Visible = True                  ' 2019/09/09 ADD
    PB_Look.Tag = ""                        ' 2019/09/09 UPD：ActiveControl.Name
    PB_Look.Visible = False                 ' 2019/09/09 UPD：True
End Sub
Private Sub iB1_Suryo_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    If iB1_Suryo <> "" Then iB1_Suryo = Format(CisFun.Val2(iB1_Suryo), "#0")
End Sub

'===============================================================================' 2019/09/06 ADD START
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
    Dim w_Color     As Long
    w_Color = &H80000008 '初期色ｾｯﾄ(黒)
'
    Select Case SYKB
        Case "ADD"
            SyoriNM = "追加"
            K_Sykbnm.ForeColor = &HFFFF&   '(黄色)
        Case "MNT"
            SyoriNM = "修正"
            K_Sykbnm.ForeColor = &HFF0000  '(青色)
        Case "DEL"
            SyoriNM = "削除"
            K_Sykbnm.ForeColor = &HFF&     '(赤色)
            w_Color = &HFF&
        Case Else
            SyoriNM = ""
    End Select
    For Each gObj In Me.Controls
        If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
        If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
    '>> * PF-Key 表示設定 *
        If SYKB <> "LOK" Then
            If gObj.Name = "PB_ADD" Or _
               gObj.Name = "PB_MNT" Or _
               gObj.Name = "PB_DEL" Then
                gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SYKB)
            End If
        End If
    Next gObj
'
'=======================================' 2019/09/09 ADD START
    ' PFキー　再設定
    Select Case SYKB
        Case "ADD"
            PB_ADD.Visible = False
            If CisFun.Val2(Dsp_Kensu) = 0 Then
                PB_MNT.Visible = False
                PB_DEL.Visible = False
            End If
        Case "MNT"
            PB_MNT.Visible = False
        Case "DEL"
            PB_DEL.Visible = False
    End Select
'=======================================' 2019/09/09 ADDEND
    
    Timer2.Enabled = False
    If SyoriNM = "" Then
       K_Sykbnm.Visible = False
       K_Sykbnm.Caption = ""
       Exit Sub
    End If
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    Timer2.Enabled = True
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*******************************
'*      処理区分名点滅制御      *
'*******************************
Private Sub Timer2_Timer()
    If K_Sykbnm.Caption <> "" Then
        K_Sykbnm.Caption = ""
    Else
        K_Sykbnm.Caption = SyoriNM
    End If
End Sub
'===============================================================================' 2019/09/06 ADD END

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
    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
    B2Mode = False: B2Color = gPLostSel
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
          Dsp_Kensu = ""
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
       Case "B2"
          B2Mode = True
          B2Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B2_Area1.Enabled = B2Mode
    B2_Area1.BackColor = B2Color

    PB_CAN.Visible = Not H1Mode
    PB_ADD.Visible = Not H1Mode
    PB_MNT.Visible = Not H1Mode
    PB_DEL.Visible = Not H1Mode
    PB_END.Visible = H1Mode

    iB1_Renbn.Enabled = True
    If SyoriKB = "MNT" Then
        iB1_Renbn.Enabled = False
    End If

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
    
    On Error GoTo HeadBodyClear_ER
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then
                .Text = ""
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    
    Exit Sub
HeadBodyClear_ER:
End Sub
'*********************************
'*          入力処理(Enter)      *
'*********************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)                     ' 2019/09/06 ADD

        If SyoriKB = "MNT" Then
            Call VSGrid1_Click
            vsGrid1.SetFocus
        End If
       'If SyoriKB = "MNT" Then SendKeys "{TAB}"    ' 2019/09/06 ADD                ''マシンによってエラーになる
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B2" Then
        If Not Body2Chk Then GoTo ReturnPress_Ed
        
    ' << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then
                RowColFlg = True    '### vsGrid1_RowColChange : 無効にする ###
                Call CisVsGrid.vsColor(消去個別)
                RowColFlg = False   '### vsGrid1_RowColChange : 有効に戻す ###
                vsGrid1.SetFocus                                                    ' 2019/09/09 ADD
            Else
                If CisFun.Val2(Dsp_Kensu) <> 0 Then
                    ProcHB = "B2"
                Else
                    ProcHB = "B1"
                    SyoriKB = "ADD"
                End If
                Call DispChange(ProcHB)
                Call SyoriName(SyoriKB)
                If vsGrid1.Row > CisFun.Val2(Dsp_Kensu.Caption) Then                ' 2019/09/09 ADD
                    vsGrid1.Row = CisFun.Val2(Dsp_Kensu.Caption)                    ' 2019/09/09 ADD
                End If                                                              ' 2019/09/09 ADD
                vsGrid1.SetFocus                                                    ' 2019/09/09 ADD
            ''''If SyoriKB = "DEL" Then SendKeys "{TAB}"
            End If
        Else
            ProcHB = "B1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)                 ' 2019/09/10 ADD
            iB1_Ymd.SetFocus                        ' 2019/09/09 ADD
        End If
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Select Case SyoriKB
            Case "ADD": ProcHB = "B1"
            Case "MNT": ProcHB = "B2"
        End Select
    ''''Call HeadBodyClear("B1")
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        If ProcHB = "B2" Then
            vsGrid1.SetFocus
            Call GridToText
        Else
            iB1_Ymd.SetFocus
        End If
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

    If Trim(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　棚卸年月を入力して下さい           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If
    
    If Not iH1_YM.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　正しい年月にて入力して下さい           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If

'===============================' 2019/09/09 ADD START
    If DBInput Then
        SyoriKB = "MNT"
        ProcHB = "B2"
    Else
        SyoriKB = "ADD"
        ProcHB = "B1"
    End If
'===============================' 2019/09/09 ADD END
    
    Head1Chk = True
End Function
'****************************
'*      BODY1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False

    Dim bFlg As Boolean

    If Trim(iB1_Ymd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　処理日を入力して下さい           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_Ymd.SetFocus
        Exit Function
    End If
    If Not iB1_Ymd.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　処理日を正しい日付にて入力して下さい           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_Ymd.SetFocus
        Exit Function
    End If
    
    bFlg = True
    If Trim(B1lb_Tanto.Caption) = "" Then
        bFlg = False
        If Trim(iB1_Tanto) <> "" Then
            B1lb_Tanto = GetTanto(Trim(iB1_Tanto))
            If B1lb_Tanto = "" Then
                bFlg = False
            End If
        End If
    End If
    If Not bFlg Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　担当者を入力して下さい           "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_Tanto.SetFocus
        Exit Function
    End If
    
    If Trim(iB1_Okiba) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　置場を入力して下さい      "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_Okiba.SetFocus
        Exit Function
    End If

    If CisFun.Val2(iB1_Renbn) = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　連番を入力してください    "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_Renbn.SetFocus
        Exit Function
    End If

    If Trim(iB1_Hinbn) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　品番を入力してください    "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_Hinbn.SetFocus
        Exit Function
    Else
        gSL_Select = "SELECT * FROM 棚卸読取結果テーブル"
        gSL_Select = gSL_Select & vbCrLf & " WHERE 品番     = '" + Trim(iB1_Hinbn) + "'"
        Call Str_Monitor(gSL_Select)
        If Not HNMRead(gSL_Select) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "   　品番マスタ未登録    "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB1_Hinbn.SetFocus
            Exit Function
        End If
    End If

'    ' 実棚数はゼロを許可するか？
'    If CisFun.Val2(iB1_Suryo) = 0 Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "   　実棚数を入力して下さい           "
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'        End With
'        iB1_Suryo.SetFocus
'        Exit Function
'    End If
    
If SyoriKB = "ADD" Then
    gSL_Select = "SELECT * FROM 棚卸読取結果テーブル"
    gSL_Select = gSL_Select & vbCrLf & " WHERE 読取連番 =  " + Trim(iB1_Renbn)
'    gSL_Select = gSL_Select & vbCrLf & "   AND 対象年月 = '" + Trim(iH1_YM) + "'"
'    gSL_Select = gSL_Select & vbCrLf & "   AND 処理日   = '" + Mid(iB1_Ymd, 1, 4) + "/" + Mid(iB1_Ymd, 5, 2) + "/" + Mid(iB1_Ymd, 7, 2) + "/" + "'"
'    gSL_Select = gSL_Select & vbCrLf & "   AND REPLACE(担当者, ' ', '') = '" + Replace(iB1_Tanto) + "'"
'    gSL_Select = gSL_Select & vbCrLf & "   AND 置場     = '" + Trim(iB1_Okiba) + "'"
'    gSL_Select = gSL_Select & vbCrLf & "   AND 品番     = '" + Trim(iB1_Hinbn) + "'"
'    gSL_Select = gSL_Select & vbCrLf & "   AND 実棚数   =  " + CStr(CisFun.Val2(iB1_Renbn))
'    gSL_Select = gSL_Select & vbCrLf & " /*AND 作成種別 = 'G'*/"
    Call Str_Monitor(gSL_Select)
    If HHRRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "   　使用されている連番です。再入力してください    "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_Renbn.SetFocus
        Exit Function
    End If
End If

    Body1Chk = True
End Function

'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 4, 0, "No.")
        Call .FixedSet(中中, 中中, あり, 11, 1, "処理日")
        Call .FixedSet(中中, 左中, あり, 20, 2, "担当者")
        Call .FixedSet(中中, 左中, あり, 10, 3, "置場")
        Call .FixedSet(中中, 右中, あり, 6, 4, "連番")
        Call .FixedSet(中中, 左中, あり, 25, 5, "品番")
        Call .FixedSet(中中, 右中, あり, 8, 6, "実棚数")
        Call .FixedSet(中中, 右中, あり, 4, 7, "担当者コード")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 2, 8, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6")
      ''Call .PatanSet(0, "0,1,7,2,3,4,5,6,8"): vsGrid1.ScrollBars = flexScrollBarBoth    ' テスト時のみ
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
''''Private Sub GridSet()
Private Function GridSet() As Boolean
    GridSet = False
    
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

' 抽出処理 初期ｾｯﾄ
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
    Dsp_Kensu = Format(Ttl_Kensu, "#,##0")                      ' 2019/09/06 ADD
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
        ID = 0
        Do Until Not HHR_RDSTS
            ID = ID + 1
            vsGrid1.TextMatrix(ID, CisVsGrid.FixedGet(0)) = CStr(ID)
            ' 処理日
            gStr = HHR.処理日
            vsGrid1.TextMatrix(ID, CisVsGrid.FixedGet(1)) = gStr
            ' 担当者
            .TextMatrix(ID, CisVsGrid.FixedGet(2)) = HHR.担当者
            ' 置場
            gStr = HHR.置場
            .TextMatrix(ID, CisVsGrid.FixedGet(3)) = gStr
            ' 連番
            gStr = HHR.読取連番
            .TextMatrix(ID, CisVsGrid.FixedGet(4)) = gStr
            ' 品番
            .TextMatrix(ID, CisVsGrid.FixedGet(5)) = HHR.品番
            '実棚数
            .TextMatrix(ID, CisVsGrid.FixedGet(6)) = Format(HHR.実棚数, "#,##0")
            ' 担当者コード
            .TextMatrix(ID, CisVsGrid.FixedGet(7)) = HHR.担当者コード

            Call HHRReadNext
        Loop
    ''''Call HHRClose                                           ' 2019/09/09 DEL
        
        .Row = 1: .Col = 1: .ColSel = .Cols - 1
        .Redraw = True
    End With
    
    GridSet = True
End Function
'**************************************
'*                                    *
'*          データベース更新          *
'*                                    *
'**************************************
'+----------------------------+
'+      削 除 処 理            +
'+----------------------------+
Private Function DBDelete()
    DBDelete = False
'  ( 確認メッセージ )
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     削除処理を行います      "
        .MB_Title = "削除処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With

'===========================================================================================' 2019/09/09 ADD START
    With CisDB
        On Error GoTo DBDelete_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 削除処理 )
        .SQL = "DELETE FROM 棚卸読取結果テーブル "
        .SQL = .SQL & vbCrLf & " WHERE 対象年月 = '" + iH1_YM.Text + "'"
        .SQL = .SQL & vbCrLf & "   AND 処理日   = '" + vsGrid1.TextMatrix(SelectedRow, CisVsGrid.FixedGet(1)) + "'"
        .SQL = .SQL & vbCrLf & "   AND REPLACE(担当者, ' ', '') = '" + vsGrid1.TextMatrix(SelectedRow, CisVsGrid.FixedGet(2)) + "'"
        .SQL = .SQL & vbCrLf & "   AND 置場     = '" + vsGrid1.TextMatrix(SelectedRow, CisVsGrid.FixedGet(3)) + "'"
        .SQL = .SQL & vbCrLf & "   AND 読取連番 =  " + vsGrid1.TextMatrix(SelectedRow, CisVsGrid.FixedGet(4))
        .SQL = .SQL & vbCrLf & "   AND 品番     = '" + vsGrid1.TextMatrix(SelectedRow, CisVsGrid.FixedGet(5)) + "'"
        .SQL = .SQL & vbCrLf & "   AND 作成種別 = 'G'"
        Call Str_Monitor(.SQL)
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        On Error GoTo 0
    End With
    ' Grid再表示
    Call vsGrid1_ReSetDisp
'===========================================================================================' 2019/09/09 ADD START
    
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了         ' 2019/09/09 ADD
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加･更新 処理    +
'+----------------------+
Private Function DBPut() As Boolean
    DBPut = False
   'Dim wWhere  As String      ' 2019/09/09 DEL
    Dim wStr    As String      ' 2019/09/09 ADD
    
    On Error GoTo DBPut_Err
'
    If SyoriKB = "ADD" Then
        wStr = "追加"
    Else
        wStr = "修正"
    End If
    
' DB更新確認
    With CisFun
        .MB_Lines = 4
        .MB_Title = wStr + "処理"
        .MB_MSG(2) = "     棚卸年月　" + wStr + "処理を行います      "
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    If SyoriKB = "ADD" Then
        Call ItemsClearHHR
        With HHR
            .対象年月 = iH1_YM.Text
            .処理日 = Mid(iB1_Ymd.Text, 1, 4) + "/" + Mid(iB1_Ymd.Text, 5, 2) + "/" + Mid(iB1_Ymd.Text, 7, 2)
            .担当者 = B1lb_Tanto.Caption
            .置場 = iB1_Okiba.Text
            .読取連番 = CisFun.Val2(iB1_Renbn.Text)
            .品番 = iB1_Hinbn.Text
            .実棚数 = CisFun.Val2(iB1_Suryo.Text)
            .作成種別 = "G"
            .作成者 = gTanto
        End With
        If Not HHRInsert() Then GoTo DBPut_Err
    Else
        gSL_Select = gSL_Select & vbCrLf & "UPDATE 棚卸読取結果テーブル"
        gSL_Select = gSL_Select & vbCrLf & "   SET 処理日   = '" + Mid(iB1_Ymd.Text, 1, 4) + "/" + Mid(iB1_Ymd.Text, 5, 2) + "/" + Mid(iB1_Ymd.Text, 7, 2) + "',"
        gSL_Select = gSL_Select & vbCrLf & "       担当者   = '" + B1lb_Tanto.Caption + "',"
        gSL_Select = gSL_Select & vbCrLf & "       置場     = '" + iB1_Okiba.Text + "',"
        gSL_Select = gSL_Select & vbCrLf & "     /*読取連番 =  " + Trim(iB1_Renbn.Text) + ",*/"
        gSL_Select = gSL_Select & vbCrLf & "       品番     = '" + iB1_Hinbn.Text + "',"
        gSL_Select = gSL_Select & vbCrLf & "       実棚数   =  " + CStr(CisFun.Val2(iB1_Suryo.Text))
        gSL_Select = gSL_Select & vbCrLf & " WHERE 対象年月 = '" + iH1_YM.Text + "'"
        gSL_Select = gSL_Select & vbCrLf & "   AND 処理日   = '" + vsGrid1.TextMatrix(SelectedRow, CisVsGrid.FixedGet(1)) + "'"
        gSL_Select = gSL_Select & vbCrLf & "   AND REPLACE(担当者, ' ', '') = '" + vsGrid1.TextMatrix(SelectedRow, CisVsGrid.FixedGet(2)) + "'"
        gSL_Select = gSL_Select & vbCrLf & "   AND 置場     = '" + vsGrid1.TextMatrix(SelectedRow, CisVsGrid.FixedGet(3)) + "'"
        gSL_Select = gSL_Select & vbCrLf & "   AND 読取連番 =  " + vsGrid1.TextMatrix(SelectedRow, CisVsGrid.FixedGet(4))
        gSL_Select = gSL_Select & vbCrLf & "   AND 品番     = '" + vsGrid1.TextMatrix(SelectedRow, CisVsGrid.FixedGet(5)) + "'"
        gSL_Select = gSL_Select & vbCrLf & "   AND 作成種別 = 'G'"
        Call Str_Monitor(gSL_Select)
        With CisDB
            .SQL = gSL_Select
            If Not .DBExec Then GoTo DBPut_Err
        End With
    End If

    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

    ' Grid再表示
    Call vsGrid1_ReSetDisp
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了         ' 2019/09/09 ADD
    CisFun.ErrorBox
    End
End Function
'*---------------------------------------------------------*
'*          シ ン グ ル プ ロ セ ス チ ェ ッ ク             *
'*---------------------------------------------------------*
Private Function SingleProcess_Check()
    SingleProcess_Check = False
    
    Call ItemsClearSYMCreate
    
    With SYMCreate
        .区分名称 = "シングルプロセス"
        .区分桁数 = 10
        .値 = App.EXEName
        .制御区分 = "X"
        .値名称 = "売上変更処理"
        .作成者 = gTanto
    End With
    
    Call SYMInitCreate(False)
    
    Call ItemsClearSYM
    gSL_Select = "Select 英数字1 From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称   = 'シングルプロセス'"
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値        = '" & App.EXEName & "'"
    Call SYMRead(gSL_Select, 1)
    
    With CisFun
        If RTrim$(SYM.英数字1) <> "" And RTrim$(SYM.英数字1) <> .WSID Then
            .MB_Lines = 3
            .MB_MSG(1) = "    他のパソコンで起動中です         "
            .MB_MSG(2) = "        ( " & RTrim$(SYM.英数字1) & " )"
            .MB_Button = Error
            .MBOX
            Exit Function
        End If
    End With

'   < プロセスのロック >
    gSL_Select = "UpDate 名称マスタ Set"
    gSL_Select = gSL_Select & " 英数字1 = SubString( Host_Name(), 1, 20 ),"
    gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
    gSL_Select = gSL_Select & " 更新者 = '" & RTrim(gTanto) & "',"
    gSL_Select = gSL_Select & " 更新端末 = SubString( Host_Name(), 1, 20 )"
    gSL_Select = gSL_Select & " Where 区分名称 = 'シングルプロセス'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 = '" & App.EXEName & "'"

    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With

    SingleProcess_Check = True
End Function
'*-------------------------------------------------*
'*          シ ン グ ル プ ロ セ ス 解 除           *
'*-------------------------------------------------*
Private Function SingleProcess_Release()
    SingleProcess_Release = False
    
    gSL_Select = "UpDate 名称マスタ Set"
    gSL_Select = gSL_Select & " 英数字1 = '',"
    gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
    gSL_Select = gSL_Select & " 更新者 = '" & RTrim(gTanto) & "',"
    gSL_Select = gSL_Select & " 更新端末 = SubString( Host_Name(), 1, 20 )"
    gSL_Select = gSL_Select & " Where 区分名称 = 'シングルプロセス'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 = '" & App.EXEName & "'"
    gSL_Select = gSL_Select & "   And 英数字1 = SubString( Host_Name(), 1, 20 )"

    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With
    
    SingleProcess_Release = True
End Function

'=======================================================================================================' 2019/09/06 ADD START
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                 +
'+  担当者コードを取得する                         +
'+                                                 +
'+++++++++++++++++++++++++++++++++++++++++++++++++++
Private Function GetTantoCode(Optional ByVal wName As String = "") As String
    GetTantoCode = ""
    
    If wName = "" Then wName = B1lb_Tanto
    wName = Replace(wName, " ", "")
    gSL_Select = "SELECT * FROM 名称マスタ"
    gSL_Select = gSL_Select & vbCrLf & " WHERE 区分名称 = '担当管理'"
    gSL_Select = gSL_Select & vbCrLf & "   AND REPLACE(値名称, ' ', '') = '" + Trim(wName) + "'"
    Call Str_Monitor(gSL_Select)
    If SYMRead(gSL_Select) Then
        GetTantoCode = SYM.値
    End If
End Function
'+++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                 +
'+  担当者名を取得する                             +
'+                                                 +
'+++++++++++++++++++++++++++++++++++++++++++++++++++
Private Function GetTanto(Optional ByVal wName As String = "") As String
    GetTanto = ""
    
    If wName = "" Then wName = iB1_Tanto
    gSL_Select = "SELECT * FROM 名称マスタ"
    gSL_Select = gSL_Select & vbCrLf & " WHERE 区分名称 = '担当管理'"
    gSL_Select = gSL_Select & vbCrLf & "   AND 値       = '" + Trim(wName) + "'"
    Call Str_Monitor(gSL_Select)
    If SYMRead(gSL_Select) Then
        GetTanto = SYM.値名称
    End If
End Function

'****************************
'*      BODY2ﾁｪｯｸ処理
'****************************
Private Function Body2Chk() As Boolean
    Body2Chk = False
'
    With vsGrid1
        If Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(1))) = "" Then GoTo Body2Chk_ED    ' 2019/09/09 UPD：Exit Function
        If Trim(.TextMatrix(.Row, 0)) = "" Then GoTo Body2Chk_ED                        ' 2019/09/09 UPD：Exit Function
'
        RowColFlg = True    '### vsGrid1_RowColChange : 無効にする ###
        Select Case SyoriKB
            Case "MNT": Call CisVsGrid.vsColor(個別, &HC0FFC0)
            Case "DEL": Call CisVsGrid.vsColor(個別, &HFFC0FF)
        End Select
        RowColFlg = False   '### vsGrid1_RowColChange : 有効に戻す ###
    End With
'
    Body2Chk = True
Body2Chk_ED:                ' 2019/09/09 ADD
    vsGrid1.SetFocus        ' 2019/09/09 ADD
End Function

'+---------------------------------------------------------------------------------------+
'+  DB読込
'+---------------------------------------------------------------------------------------+
Private Function DBInput() As Boolean
    DBInput = False
    
    On Error GoTo DBInput_Err
    
    ' 抽出処理 初期ｾｯﾄ
''''    gSL_Select = "SELECT *"
''''    gSL_Select = gSL_Select & vbCrLf & "  FROM 棚卸読取結果テーブル"
''''    gSL_Select = gSL_Select & vbCrLf & " WHERE 対象年月 = '" + iH1_YM.Text + "'"
''''    gSL_Select = gSL_Select & vbCrLf & "   AND 作成種別 = 'G'"
    gSL_Select = ""
    gSL_Select = gSL_Select & vbCrLf & "SELECT TT.対象年月, TT.処理日,"
    gSL_Select = gSL_Select & vbCrLf & "       ISNULL(MM.値, '') 担当者コード,"
    gSL_Select = gSL_Select & vbCrLf & "       CASE WHEN ISNULL(MM.値, '') = ''"
    gSL_Select = gSL_Select & vbCrLf & "            THEN TT.担当者"
    gSL_Select = gSL_Select & vbCrLf & "            ELSE MM.値名称"
    gSL_Select = gSL_Select & vbCrLf & "       END  担当者,"
    gSL_Select = gSL_Select & vbCrLf & "       TT.置場, TT.読取連番, TT.品番, TT.実棚数"
    gSL_Select = gSL_Select & vbCrLf & "    --,TT.*"
    gSL_Select = gSL_Select & vbCrLf & "  FROM 棚卸読取結果テーブル TT"
    gSL_Select = gSL_Select & vbCrLf & "  LEFT OUTER JOIN 名称マスタ MM"
    gSL_Select = gSL_Select & vbCrLf & "    ON MM.区分名称   = '担当管理'"
    gSL_Select = gSL_Select & vbCrLf & "   AND MM.区分タイプ = ''"
    gSL_Select = gSL_Select & vbCrLf & "   AND REPLACE(MM.値名称, ' ', '') = REPLACE(TT.担当者, ' ', '')"
    gSL_Select = gSL_Select & vbCrLf & " WHERE TT.対象年月 = '" + iH1_YM.Text + "'"
    gSL_Select = gSL_Select & vbCrLf & "   AND TT.作成種別 = 'G'"
    Call Str_Monitor(gSL_Select)
    If HHRRead(gSL_Select) Then
        Call GridSet
        DBInput = True
    End If
    Call HHRClose
    
    Exit Function
DBInput_Err:
    CisFun.ErrorBox
End Function

'+---------------------------------------------------------------------------------------+
'+  グリッドより入力域へセット
'+---------------------------------------------------------------------------------------+
Private Sub GridToText()
    With vsGrid1
        SelectedRow = .Row
        iB1_Ymd = Replace(.TextMatrix(SelectedRow, CisVsGrid.FixedGet(1)), "/", "")
        B1lb_Tanto = .TextMatrix(SelectedRow, CisVsGrid.FixedGet(2))
        iB1_Tanto = GetTantoCode(B1lb_Tanto)
        iB1_Okiba = .TextMatrix(SelectedRow, CisVsGrid.FixedGet(3))
        iB1_Renbn = .TextMatrix(SelectedRow, CisVsGrid.FixedGet(4))
        iB1_Hinbn = .TextMatrix(SelectedRow, CisVsGrid.FixedGet(5))
        iB1_Suryo = .TextMatrix(SelectedRow, CisVsGrid.FixedGet(6))
    End With
End Sub
'+---------------------------------------+
'+      グリッド    (Click)
'+---------------------------------------+
Private Sub VSGrid1_Click()
    If ProcHB = "B2" Then
        Call GridToText
    End If
End Sub
'+---------------------------------------+
'+      グリッド    (DblClick)
'+---------------------------------------+
Private Sub vsGrid1_DblClick()
    If ProcHB = "B2" Then
        Call ReturnPress
    End If
End Sub
'+---------------------------------------+
'+      グリッド    (KeyPress)
'+---------------------------------------+
Private Sub vsGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
       Call vsGrid1_DblClick
    End If
End Sub
'+---------------------------------------+
'+      グリッド    (RowColChange)
'+---------------------------------------+
Private Sub vsGrid1_RowColChange()
    If FormAct Then
    '>> 実行判定 .. False:実行
        Select Case RowColFlg
            Case False: Call VSGrid1_Click
        End Select
    End If
End Sub
'+---------------------------------------+
'+      グリッド 再表示
'+---------------------------------------+
Private Sub vsGrid1_ReSetDisp()
    SaveRow = vsGrid1.Row
'
    RowColFlg = True    '### vsGrid1_RowColChange : 無効にする ###
    Call HeadBodyClear("B")
    If SyoriKB = "DEL" Then
        Call CisVsGrid.Clear
    End If
    Call DBInput
    RowColFlg = False   '### vsGrid1_RowColChange : 有効に戻す ###
'
    If SaveRow > vsGrid1.Rows - 1 Then
        SaveRow = SaveRow - 1
    End If
    If SyoriKB <> "ADD" Then
        vsGrid1.Row = SaveRow
    End If
    If vsGrid1.RowIsVisible(SaveRow) = False Then
        vsGrid1.TopRow = SaveRow
    End If

    '===========================================' 2019/09/09 ADD START
    Call CisVsGrid.vsColor(消去個別)
    Call SyoriName(SyoriKB)
    Call DispChange(ProcHB)
    If ProcHB = "B2" Then vsGrid1.SetFocus
    '===========================================' 2019/09/09 ADD END
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++    '++ デバッグ用です +++
'+                                                 +
'+  文字列をイミディエイトウィンドウに表示する     +
'+                                                 +
'+++++++++++++++++++++++++++++++++++++++++++++++++++
Public Sub Str_Monitor(ByVal wStr As String)
    Debug.Print vbCrLf
    Debug.Print Format(Now, "yyyy/MM/dd HH:mm:ss") + _
                " .................................................." + _
                vbCrLf + wStr
End Sub
'=======================================================================================================' 2019/09/06 ADD END

''+---------------------+
''+    契約先検索        +
''+---------------------+
'Private Sub Look_KTorcd()
'
'    RV_Left = 0
'    RV_Top = 0
'    RV_Call = "C"
'    RV_TorKb = 2
'
'    CKK0025.Show vbModal
'    Unload CKK0025
'    Set CKK0025 = Nothing
'
'    iB1_KTorcd.SetFocus
'
'    If RV_Rtn Then
'       iB1_KTorcd = RV_TorcdK
'       Call iB1_KTorcd_LostFocus
'    End If
'
'End Sub
''+----------------------------+
''+       納入先検索            +
''+----------------------------+
'Private Sub Look_Torcd()
'
'    RV_Left = 0
'    RV_Top = 0
'    RV_Call = "C"
'    RV_TorKb = 3
'
'    CKK0025.Show vbModal
'    Unload CKK0025
'    Set CKK0025 = Nothing
'
'    If RV_Rtn Then
'        iB1_Torcd = RTrim(RV_TorcdK)
'        Call iB1_Torcd_LostFocus
'    End If
'End Sub
''+----------------------------+
''+     納入先受入検索          +
''+----------------------------+
'Private Sub Look_Ukeir()
'
'    RV_Left = 0
'    RV_Top = 0
'    RV_Call = "C"
'    RV_TorKb = 3
'    RV_Torcd = RTrim(iB1_Torcd)
'
'    CKK0028.Show vbModal
'    Unload CKK0028
'    Set CKK0028 = Nothing
'
'    If RV_Rtn Then
'        iB1_Torcd = RTrim(RV_TorcdK)
'        Call iB1_Torcd_LostFocus
'        iB1_Ukeir = RTrim(RV_Ukeir)
'    End If
'End Sub


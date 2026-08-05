VERSION 5.00
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXT0210 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ＴＮＳ送信データ作成設定"
   ClientHeight    =   11235
   ClientLeft      =   1635
   ClientTop       =   1380
   ClientWidth     =   15360
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
   Picture         =   "CXT0210.frx":0000
   ScaleHeight     =   11235
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   7350
      ScaleHeight     =   750
      ScaleWidth      =   7740
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   720
      Width           =   7800
      Begin CisBtn_60.CisBtn PB_LOOK 
         Height          =   585
         Left            =   30
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   5160
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
         Left            =   6015
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
         Left            =   6870
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   2595
         Top             =   90
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
         Left            =   885
         Top             =   90
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
         BtnType         =   7
         BtnCaption      =   "新規"
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
         Left            =   1740
         Top             =   90
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
      Begin CisBtn_60.CisBtn PB_BACK 
         Height          =   585
         Left            =   3450
         Top             =   90
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
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   4305
         Top             =   90
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
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'ｵﾌ固定
      Left            =   -30
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   420
      Width           =   855
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
      Height          =   7275
      Left            =   4380
      ScaleHeight     =   7215
      ScaleWidth      =   6525
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3510
      Width           =   6585
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   7020
         Left            =   120
         Top             =   90
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   12383
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
         Begin Cis3D_v60.CIS3D Back_Hinbn 
            Height          =   810
            Left            =   30
            Top             =   3930
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   1429
            ForeColor       =   16711680
            Caption         =   "  品番ﾁｪｯｸ"
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
            cPositionX      =   90
            cPositionY      =   30
            Begin VB.CheckBox B1Chk_UkeChk 
               BackColor       =   &H00C0C0C0&
               Caption         =   "確認なし"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1980
               TabIndex        =   19
               Top             =   405
               Width           =   1185
            End
            Begin VB.CheckBox B1Chk_HINList 
               BackColor       =   &H00C0C0C0&
               Caption         =   "発行なし"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3630
               TabIndex        =   20
               Top             =   390
               Width           =   1185
            End
            Begin VB.CheckBox B1Chk_HINChk 
               BackColor       =   &H00C0C0C0&
               Caption         =   "確認なし"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   300
               TabIndex        =   18
               Top             =   420
               Width           =   1185
            End
            Begin Cis3D_v60.CIS3D CIS3D25 
               Height          =   285
               Left            =   3810
               Top             =   30
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               ForeColor       =   16711680
               Caption         =   "ｴﾗｰﾘｽﾄ"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cBoderStyle     =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D26 
               Height          =   285
               Left            =   2040
               Top             =   45
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               ForeColor       =   16711680
               Caption         =   "受入ﾁｪｯｸ"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cBoderStyle     =   2
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D20 
            Height          =   435
            Left            =   30
            Top             =   4755
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   767
            BackColor       =   0
            ForeColor       =   65535
            Caption         =   "《 クライアント設定 》"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   18
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D18 
            Height          =   810
            Left            =   30
            Top             =   1650
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   1429
            ForeColor       =   16711680
            Caption         =   "ｾﾞﾛ件ﾃﾞｰﾀ"
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
            Begin VB.CheckBox B1Chk_ZeroKB 
               BackColor       =   &H00C0C0C0&
               Caption         =   "作成あり"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   90
               TabIndex        =   11
               Top             =   390
               Width           =   1155
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D13 
            Height          =   810
            Left            =   3045
            Top             =   1650
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   1429
            ForeColor       =   16711680
            Caption         =   "TNSフォーマット"
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
            Begin VB.OptionButton B1Op_TNSFormat 
               BackColor       =   &H00C0C0C0&
               Caption         =   "TNS拡張形式"
               Height          =   255
               Index           =   1
               Left            =   1410
               TabIndex        =   14
               Top             =   420
               Width           =   1635
            End
            Begin VB.OptionButton B1Op_TNSFormat 
               BackColor       =   &H00C0C0C0&
               Caption         =   "TNS形式"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   13
               Top             =   420
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D14 
            Height          =   810
            Left            =   1005
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1429
            ForeColor       =   16711680
            Caption         =   "発信元"
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
            Begin CisText_V60.CisText iB1_SMoto 
               Height          =   375
               Left            =   90
               TabIndex        =   8
               Top             =   360
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "XXXXX"
               MaxLength       =   5
               cDataReplace    =   1
               cFbComma        =   0
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D12 
            Height          =   810
            Left            =   30
            Top             =   840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1429
            ForeColor       =   16711680
            Caption         =   "宛先"
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
            Begin CisText_V60.CisText iB1_Atesaki 
               Height          =   375
               Left            =   90
               TabIndex        =   7
               Top             =   360
               Width           =   735
               _ExtentX        =   1296
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
               Text            =   "XXXXX"
               MaxLength       =   5
               cDataReplace    =   1
               cFbComma        =   0
            End
         End
         Begin Cis3D_v60.CIS3D Back_Detail 
            Height          =   915
            Index           =   1
            Left            =   30
            Top             =   5190
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   1614
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
            Begin Cis3D_v60.CIS3D CIS3D10 
               Height          =   840
               Left            =   30
               Top             =   30
               Width           =   6165
               _ExtentX        =   10874
               _ExtentY        =   1482
               BackColor       =   14737632
               Caption         =   "TNS送信ﾃﾞｰﾀ作成先"
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
               cAlingnment     =   6
               cBoderStyle     =   1
               cPositionY      =   30
               Begin VB.CheckBox B1Chk_CrtKB1 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "共通"
                  Height          =   240
                  Left            =   120
                  TabIndex        =   21
                  Top             =   60
                  Width           =   1005
               End
               Begin CisText_V60.CisText iB1_TextDir 
                  Height          =   375
                  Left            =   60
                  TabIndex        =   22
                  Top             =   360
                  Width           =   5790
                  _ExtentX        =   10213
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
                  Text            =   $"CXT0210.frx":0342
                  MaxLength       =   250
                  IMEMode         =   2
                  cCaps           =   2
               End
               Begin Cis3D_v60.CIS3D PB_TextDir 
                  Height          =   315
                  Left            =   5760
                  Top             =   360
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
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   810
            Left            =   1980
            Top             =   840
            Width           =   1065
            _ExtentX        =   1879
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
               TabIndex        =   9
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
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   810
            Left            =   30
            Top             =   30
            Width           =   6250
            _ExtentX        =   11033
            _ExtentY        =   1429
            ForeColor       =   16711680
            Caption         =   "業務名（ＴＮＳファイル名称)"
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
               TabIndex        =   6
               Top             =   360
               Width           =   6135
               _ExtentX        =   10821
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
               Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   48
               IMEMode         =   4
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   870
            Left            =   30
            Top             =   6105
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   1535
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cPositionX      =   60
            Begin Cis3D_v60.CIS3D CIS3D2 
               Height          =   810
               Left            =   30
               Top             =   30
               Width           =   6180
               _ExtentX        =   10901
               _ExtentY        =   1429
               BackColor       =   14737632
               Caption         =   "TNS送信ﾃﾞｰﾀﾊﾞｯｸｱｯﾌﾟ作成先"
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
               cAlingnment     =   6
               cBoderStyle     =   1
               cPositionY      =   30
               Begin VB.CheckBox B1Chk_CrtKB2 
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "共通"
                  Height          =   240
                  Left            =   75
                  TabIndex        =   23
                  Top             =   60
                  Width           =   1005
               End
               Begin CisText_V60.CisText iB1_BackDir 
                  Height          =   375
                  Left            =   60
                  TabIndex        =   24
                  Top             =   360
                  Width           =   5790
                  _ExtentX        =   10213
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
                  Text            =   $"CXT0210.frx":0440
                  MaxLength       =   250
                  IMEMode         =   2
                  cCaps           =   2
               End
               Begin Cis3D_v60.CIS3D PB_BackDir 
                  Height          =   315
                  Left            =   5760
                  Top             =   360
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
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   810
            Left            =   2415
            Top             =   3120
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   1429
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cPositionX      =   30
            Begin Cis3D_v60.CIS3D CIS3D24 
               Height          =   345
               Left            =   645
               Top             =   45
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   609
               ForeColor       =   16711680
               Caption         =   "ﾊﾞｯｸｱｯﾌﾟ 削除日数"
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
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cBoderStyle     =   2
            End
            Begin CisText_V60.CisText iB1_Del 
               Height          =   375
               Left            =   180
               TabIndex        =   17
               Top             =   360
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   661
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "###"
               cILength        =   3
            End
            Begin VB.Label Label1 
               BackColor       =   &H00C0C0C0&
               Caption         =   "日以前を削除(未入力は削除なし)"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   780
               TabIndex        =   28
               Top             =   480
               Width           =   2745
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D8 
            Height          =   660
            Left            =   30
            Top             =   2460
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   1164
            ForeColor       =   16711680
            Caption         =   "TNS送信ﾌｧｲﾙ名"
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
            cAlingnment     =   1
            cPositionX      =   30
            Begin Cis3D_v60.CIS3D CIS3D22 
               Height          =   240
               Left            =   3600
               Top             =   330
               Width           =   2640
               _ExtentX        =   4657
               _ExtentY        =   423
               Caption         =   "【ﾌｧｲﾙｺｰﾄﾞ+年月日時分秒.TXT】"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
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
               cFont3D         =   2
               cBoderStyle     =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D19 
               Height          =   225
               Left            =   4230
               Top             =   90
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   397
               Caption         =   "ﾌｧｲﾙ名空白時"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
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
               cFont3D         =   2
               cBoderStyle     =   2
            End
            Begin CisText_V60.CisText iB1_TextID 
               Height          =   375
               Left            =   1650
               TabIndex        =   15
               Top             =   150
               Width           =   1905
               _ExtentX        =   3360
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
               Text            =   "XXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   20
               cDataReplace    =   1
               cFbComma        =   0
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   810
            Left            =   3045
            Top             =   840
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   1429
            ForeColor       =   16711680
            Caption         =   "TNSデータ作成タイプ"
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
            Begin VB.ComboBox B1Comb_CrtType 
               BackColor       =   &H00C0E0FF&
               Height          =   360
               Left            =   90
               Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
               TabIndex        =   10
               Top             =   360
               Width           =   3075
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D23 
            Height          =   810
            Left            =   1380
            Top             =   1650
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   1429
            ForeColor       =   16711680
            Caption         =   "TNSｼｽﾃﾑ起動"
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
            Begin VB.CheckBox B1Chk_TNSSys 
               BackColor       =   &H00C0C0C0&
               Caption         =   "確認なし"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   210
               TabIndex        =   12
               Top             =   420
               Width           =   1395
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D30 
            Height          =   810
            Left            =   30
            Top             =   3120
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   1429
            ForeColor       =   16711680
            Caption         =   "送信開始日"
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
            Begin CisYMD_With_Btn.CisYMDwB iB1_SSYmd 
               Height          =   315
               Left            =   285
               TabIndex        =   16
               Top             =   375
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
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
      Height          =   1710
      Left            =   4395
      ScaleHeight     =   1650
      ScaleWidth      =   6510
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1665
      Width           =   6570
      Begin Cis3D_v60.CIS3D CIS3D16 
         Height          =   1560
         Left            =   5370
         Top             =   60
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   2752
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
            Left            =   240
            TabIndex        =   1
            Top             =   330
            Width           =   540
            _ExtentX        =   953
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
            Text            =   "XXX"
            MaxLength       =   3
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D17 
         Height          =   780
         Left            =   1710
         Top             =   60
         Width           =   3660
         _ExtentX        =   6456
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
         Begin CisText_V60.CisText iH1_KTorcd 
            Height          =   360
            Left            =   90
            TabIndex        =   0
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
         Begin Cis3D_v60.CIS3D H1lb_KTorNm 
            Height          =   360
            Left            =   1080
            Top             =   330
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXXX"
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
      Begin Cis3D_v60.CIS3D CIS3D27 
         Height          =   780
         Left            =   1710
         Top             =   840
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "代行契約先"
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
         Begin CisText_V60.CisText iH1_DaikoK 
            Height          =   360
            Left            =   90
            TabIndex        =   2
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
         Begin Cis3D_v60.CIS3D H1lb_DaikoK 
            Height          =   360
            Left            =   1080
            Top             =   330
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXXX"
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
      Begin Cis3D_v60.CIS3D CIS3D29 
         Height          =   1560
         Left            =   45
         Top             =   60
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   2752
         ForeColor       =   16711680
         Caption         =   "送信先指定"
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
         Begin VB.OptionButton OpH1_SendKb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "代行契約先"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   90
            TabIndex        =   4
            Top             =   945
            Width           =   1515
         End
         Begin VB.OptionButton OpH1_SendKb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "契約先"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   90
            TabIndex        =   3
            Top             =   480
            Value           =   -1  'True
            Width           =   1215
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   420
   End
   Begin VB.Timer Timer2 
      Left            =   690
      Top             =   420
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 ＴＮＳ送信データ作成設定  】"
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
         Left            =   14220
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   11730
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
         cPositionX      =   50
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
            Size            =   12
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Cis3D_v60.CIS3D CIS3D11 
      Height          =   810
      Left            =   2850
      Top             =   5520
      Visible         =   0   'False
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1429
      ForeColor       =   16711680
      Caption         =   "送信"
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
      Begin VB.CheckBox B1Chk_SendKB 
         BackColor       =   &H00C0C0C0&
         Caption         =   "手動"
         Height          =   315
         Left            =   60
         TabIndex        =   29
         Top             =   390
         Width           =   825
      End
   End
   Begin Cis3D_v60.CIS3D CIS3D5 
      Height          =   810
      Left            =   1800
      Top             =   4680
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1429
      ForeColor       =   16711680
      Caption         =   "TNS宛先"
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
      Begin CisText_V60.CisText iB1_TNSate 
         Height          =   375
         Left            =   60
         TabIndex        =   30
         Top             =   360
         Width           =   2145
         _ExtentX        =   3784
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
         Text            =   "XXXXXXXXXXXXXXXX"
         MaxLength       =   16
         cDataReplace    =   1
         cFbComma        =   0
      End
   End
   Begin Cis3D_v60.CIS3D CIS3D15 
      Height          =   810
      Left            =   2730
      Top             =   3810
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1429
      ForeColor       =   16711680
      Caption         =   "識別子"
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
      Begin CisText_V60.CisText iB1_Shiki 
         Height          =   375
         Left            =   90
         TabIndex        =   31
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
         cDataReplace    =   1
         cFbComma        =   0
      End
   End
   Begin Cis3D_v60.CIS3D CIS3D21 
      Height          =   810
      Left            =   2160
      Top             =   6420
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   1429
      ForeColor       =   16711680
      Caption         =   "出力区分"
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
      Begin VB.OptionButton B1Op_OutKB 
         BackColor       =   &H00C0C0C0&
         Caption         =   "置換"
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
         Left            =   60
         TabIndex        =   33
         Top             =   360
         Width           =   795
      End
      Begin VB.OptionButton B1Op_OutKB 
         BackColor       =   &H00C0C0C0&
         Caption         =   "追加"
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
         Left            =   840
         TabIndex        =   32
         Top             =   360
         Width           =   795
      End
   End
End
Attribute VB_Name = "CXT0210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   ＴＮＳ送信データ作成設定
'**       フォームID    :   CXT0210
'**       処理概要      :
'**
'**       作  成  日    :   2006/02/09  By CIS
'**       変  更  日    :   2006/03/09  By CIS ファミックカスタマイズ　出力ファイル名空白をOKとする
'**                                             PCT0010にて作成時、ファイル名空白は【FCODE+年月日字分秒.TXT】
'**       変  更  日    :   2007/01/26  By CIS
'**       変  更  日    :   2007/02/08  By CIS  車体支給追加(TNS区分)･TNSシステム起動項目追加
'**       変  更  日    :   2007/03/08  By CIS  品番チェック項目追加(C04･R07対応)
'**       変  更  日    :   2007/12/27  By CIS  代行契約先への送信設定追加
'**                                             送信開始日･受入チェック項目追加
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim SyoriKB                 As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM                 As String       ' 処理名称(新規･修正･削除 等)
    
    Dim CisComb                 As CisCombo     '

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
    If Not IniGet Then End
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
    
  ' 名称マスタ確認
    Call MeisyoInit
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Call CisFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
    
'    '+---------------------+
'    '+ ｺﾝﾎﾞﾎﾞｯｸｽ 初期設定  +
'    '+---------------------+
    Set CisComb = New CisCombo
    With CisComb
        Set .Connect = CisDB
        Set .CombObj = B1Comb_CrtType
        .Control = "名称マスタ"
        .ComboName = "TNS作成タイプ"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
    End With
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = "ADD"
    Call SyoriName(SyoriKB)
    
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
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
    Key_Acc = False
    
    Select Case Key_Code
           Case vbKeyF1:        If PB_ADD.Visible Then Call PB_ADD_Click    '【新規】
           Case vbKeyF2:        If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF3:        If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF7:        If PB_LOOK.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyPageUp:    If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown:  If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_Can_Click    '【取消】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                      '【入力】
                                    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT_Click
                                    Dummy.Enabled = False
                                End If
           Case Else:           Exit Function
    End Select
    
    Key_Acc = True
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
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
    Call CisFrm.MousePT(11)
    
    If ProcHB = "B1" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        Call SyoriName(SyoriKB)
        iH1_KTorcd.SetFocus
        GoTo PB_CAN_Ed
    End If
    
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新  規  キ ー(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    If SyoriKB <> "ADD" Then
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
    End If
    iH1_KTorcd.SetFocus
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKB <> "MNT" Then
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
    End If
    iH1_KTorcd.SetFocus
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        SyoriKB = "DEL"
    Else
        SyoriKB = ""
    End If
    Call SyoriName(SyoriKB)
    iH1_KTorcd.SetFocus
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = "SELECT"
    gSL_Select = gSL_Select & " * From TNS送信設定テーブル"
    gSL_Select = gSL_Select & " WHERE 枝番 = 0"
    gSL_Select = gSL_Select & " AND 送信先 + SPACE(5 - LEN(送信先)) + 代行送信先 + SPACE(5 - LEN(代行送信先)) + FCODE "
    gSL_Select = gSL_Select & "     < '" & RTrim$(iH1_KTorcd) & Space(5 - Len(iH1_KTorcd)) & RTrim$(iH1_DaikoK) & Space(5 - Len(iH1_DaikoK)) & RTrim$(iH1_FileCd) & "'"
    gSL_Select = gSL_Select & " ORDER BY 送信先 DESC,FCODE DESC "

    Call DBInput("Back")
    
    iH1_KTorcd.SetFocus
    
    Call iH1_KTorcd_GotFocus '------------------------------ 2007/01/26 Insert
    
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = "SELECT"
    gSL_Select = gSL_Select & " * From TNS送信設定テーブル"
    gSL_Select = gSL_Select & " WHERE 枝番 = 0"
    gSL_Select = gSL_Select & " AND 送信先 + SPACE(5 - LEN(送信先)) + 代行送信先 + SPACE(5 - LEN(代行送信先)) + FCODE "
    gSL_Select = gSL_Select & "     > '" & RTrim$(iH1_KTorcd) & Space(5 - Len(iH1_KTorcd)) & RTrim$(iH1_DaikoK) & Space(5 - Len(iH1_DaikoK)) & RTrim$(iH1_FileCd) & "'"
    gSL_Select = gSL_Select & " ORDER BY 送信先 ,FCODE "
    
    Call DBInput("Next")
    
    Call iH1_KTorcd_GotFocus '------------------------------ 2007/01/26 Insert
    
    iH1_KTorcd.SetFocus
    
    Call CisFrm.MousePT(1)
End Sub
Private Sub PB_TextDir_Click()
    With CisFun
        .Folder_Title = "TNS送信ﾃﾞｰﾀ作成先参照"
        If Trim(iB1_TextDir) = "" Then
           .Folder_DefDir = gIniExe
        Else
           .Folder_DefDir = iB1_TextDir
        End If
        .FolderSel
        
        If RTrim$(.Folder_Result) <> "" Then iB1_TextDir = RTrim$(.Folder_Result)
    End With
    iB1_TextDir.SetFocus

End Sub
Private Sub PB_BackDir_Click()
    With CisFun
        .Folder_Title = "TNS送信ﾃﾞｰﾀ作成先参照"
        If Trim(iB1_BackDir) = "" Then
           .Folder_DefDir = gIniExe
        Else
           .Folder_DefDir = iB1_BackDir
        End If
        .FolderSel
        
        If RTrim$(.Folder_Result) <> "" Then iB1_BackDir = RTrim$(.Folder_Result)
    End With
    iB1_BackDir.SetFocus

End Sub

'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_LOOK.Tag = "iH1_KTorcd" Then Call Look_Torcd(0): Exit Sub
    If PB_LOOK.Tag = "iH1_DaikoK" Then Call Look_Torcd(1): Exit Sub                     '2007/12/26 ADD
    If PB_LOOK.Tag = "iB1_TextDir" Then Call PB_TextDir_Click: Exit Sub
    If PB_LOOK.Tag = "iB1_BackDir" Then Call PB_BackDir_Click: Exit Sub
End Sub
'+----------------------------+
'+               検索         +
'+----------------------------+
Private Sub Look_Torcd(TKb As Byte)
    RV_Left = 0
    RV_Top = 0
    If TKb = 0 Then
        RV_TorKb = 5
    Else
        RV_TorKb = 8                                                                    '2007/12/26 ADD
    End If
    RV_Call = ""
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        If TKb = 0 Then
            iH1_KTorcd = RV_TorcdK
            H1lb_KTorNm = RV_TorRName
        Else
            iH1_DaikoK = RV_TorcdK                                                      '2007/12/26 ADD
            H1lb_DaikoK = RV_TorRName
        End If
    End If
End Sub
'+-----------------------------+
'+  iH1_KTorcd
'+-----------------------------+
Private Sub iH1_KTorcd_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name:       PB_LOOK.Visible = True
End Sub
Private Sub iH1_KTorcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
    
    H1lb_KTorNm = ""
    If TorNmGet(iH1_KTorcd, 5) Then
       H1lb_KTorNm = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iH1_DaikoK                      '2007/12/26 ADD
'+-----------------------------+
Private Sub iH1_DaikoK_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name:       PB_LOOK.Visible = True
End Sub
Private Sub iH1_DaikoK_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
    
    H1lb_DaikoK = ""
    If TorNmGet(iH1_DaikoK, 8) Then
       H1lb_DaikoK = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  OpH1_SendKb
'+-----------------------------+
Private Sub OpH1_SendKb_Click(Index As Integer)
    If OpH1_SendKb(0).Value Then
        OpH1_SendKb(0).Tag = "0"
    Else
        OpH1_SendKb(0).Tag = "1"
    End If
End Sub
'+-----------------------------+
'+  B1Chk_ZeroKB
'+-----------------------------+
Private Sub B1Chk_ZeroKB_Click()
    If B1Chk_ZeroKB.Value = 0 Then
        B1Chk_ZeroKB.Caption = "作成あり"
    Else
        B1Chk_ZeroKB.Caption = "作成なし"
    End If
End Sub
'+--------------------------------------------------------- 2007/02/08 Insert
'+-----------------------------+
'+  B1Chk_TNSSys
'+-----------------------------+
Private Sub B1Chk_TNSSys_Click()
    If B1Chk_TNSSys.Value = 0 Then
        B1Chk_TNSSys.Caption = "確認なし"
    Else
        B1Chk_TNSSys.Caption = "確認あり"
    End If
End Sub
'+--------------------------------------------------------- 2007/02/08 Insert End
'+--------------------------------------------------------- 2007/03/08 Insert
'+-----------------------------+
'+  B1Chk_HINChk
'+-----------------------------+
Private Sub B1Chk_HINChk_Click()
    If B1Chk_HINChk.Value = 0 Then
        B1Chk_HINChk.Caption = "確認なし"
    Else
        B1Chk_HINChk.Caption = "確認あり"
    End If
End Sub
'+-----------------------------+
'+  B1Chk_UkeChk                        '2007/12/26 ADD
'+-----------------------------+
Private Sub B1Chk_UkeChk_Click()
    If B1Chk_UkeChk.Value = 0 Then
        B1Chk_UkeChk.Caption = "確認なし"
    Else
        B1Chk_UkeChk.Caption = "確認あり"
    End If
End Sub
'+-----------------------------+
'+  B1Chk_HINList
'+-----------------------------+
Private Sub B1Chk_HINList_Click()
    If B1Chk_HINList.Value = 0 Then
        B1Chk_HINList.Caption = "発行なし"
    Else
        B1Chk_HINList.Caption = "発行あり"
    End If
End Sub
'+--------------------------------------------------------- 2007/03/08 Insert End
'+-----------------------------+
'+  B1Chk_SendKB
'+-----------------------------+
Private Sub B1Chk_SendKB_Click()
    If B1Chk_SendKB.Value = 0 Then
        B1Chk_SendKB.Caption = "手動"
    Else
        B1Chk_SendKB.Caption = "自動"
    End If
End Sub
'+-----------------------------+
'+  iB1_TextDir
'+-----------------------------+
Private Sub iB1_TextDir_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name:       PB_LOOK.Visible = True
End Sub
Private Sub iB1_TextDir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
End Sub
'+-----------------------------+
'+  B1Chk_CrtKB1
'+-----------------------------+
Private Sub B1Chk_CrtKB1_Click()
    With CisFun
        .INI_FileID = gIniExe & gwsIniName
        .INI_Section = "TNSSENDDIR"
        If B1Chk_CrtKB1.Value = 0 Then
            B1Chk_CrtKB1.Caption = "共通"
            .INI_Key = "COMMON"
        Else
            B1Chk_CrtKB1.Caption = "個別"
            .INI_Key = RTrim(iH1_KTorcd) & "_" & RTrim(iH1_FileCd)
        End If
        .GetIni
        iB1_TextDir = RTrim(.INI_String)
        
        .INI_FileID = gIniExe & gIniName
    End With
End Sub
'+-----------------------------+
'+  iB1_BackDir
'+-----------------------------+
Private Sub iB1_BackDir_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name:       PB_LOOK.Visible = True
End Sub
Private Sub iB1_BackDir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
End Sub
'+-----------------------------+
'+  B1Chk_CrtKB2
'+-----------------------------+
Private Sub B1Chk_CrtKB2_Click()
    With CisFun
        .INI_FileID = gIniExe & gwsIniName
        .INI_Section = "TNSBACKDIR"
        If B1Chk_CrtKB2.Value = 0 Then
            B1Chk_CrtKB2.Caption = "共通"
            .INI_Key = "COMMON"
        Else
            B1Chk_CrtKB2.Caption = "個別"
            .INI_Key = RTrim(iH1_KTorcd) & "_" & RTrim(iH1_FileCd)
        End If
        CisFun.GetIni
        iB1_BackDir = RTrim(.INI_String)
        .INI_FileID = gIniExe & gIniName
    End With
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
    Static w_Color      As Long
    w_Color = &H80000008 '初期色ｾｯﾄ(黒)
    
    Select Case SyoriKB
        Case "ADD"
                            SyoriNM = "追加"
                            K_Sykbnm.ForeColor = &HFFFF&            '(黄色)
        Case "MNT"
                            SyoriNM = "修正"
                            K_Sykbnm.ForeColor = &HFF0000           '(青色)
        Case "DEL"
                            SyoriNM = "削除"
                            K_Sykbnm.ForeColor = &HFF&              '(赤色)
                            w_Color = &HFF&
        Case Else:          SyoriNM = ""
    End Select
    
    For Each gObj In Me.Controls
        If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
        If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
    
'       * PF-Key 表示設定 *
        If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
           gObj.Name = "PB_DEL" Then
                gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
        End If
    Next gObj
    
    Timer2.Enabled = False
    K_Sykbnm = SyoriNM
    K_Sykbnm.Visible = Not (SyoriNM = "")
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
    If RTrim$(K_Sykbnm) <> "" Then K_Sykbnm = "" Else K_Sykbnm = SyoriNM
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
    
    H1Mode = False:     H1Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "B1":      B1Mode = True:  B1Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    
    PB_LOOK.Visible = H1Mode
    PB_ADD.Visible = H1Mode
    PB_MNT.Visible = H1Mode
    PB_DEL.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    
'   ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
End Function
'**************************************
'*      Head/Body 初期化処理           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String
    
    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*" Else wbc_Enm = Trim(ClrType)
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = "": .Enabled = True
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "BC" Then
       B1Op_TNSFormat(0).Value = True
       B1Op_OutKB(0).Value = True
       CisComb.ComboSet
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)

    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed

'       << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        Else
            ProcHB = "B1"
            Call DispChange(ProcHB)
        End If
        GoTo ReturnPress_Ed
    End If

    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
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
    
'(契約先)
   If RTrim$(iH1_KTorcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    契約先を入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_KTorcd.SetFocus
        Exit Function
    Else
        If Not TorNmGet(iH1_KTorcd, 5) Then
           With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    契約先未登録         "
                .MB_Button = Error
                .MBOX
           End With
           iH1_KTorcd.SetFocus
           Exit Function
        End If
    End If

'(ファイルコード)
   If RTrim$(iH1_FileCd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    ファイルコードを入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_FileCd.SetFocus
        Exit Function
    End If
'
'(代行契約先)                               '2007/12/26 ADD
    If OpH1_SendKb(1).Value Then
        If RTrim$(iH1_DaikoK) = "" Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    代行契約先を入力して下さい         "
                 .MB_Button = Error
                 .MBOX
             End With
             iH1_DaikoK.SetFocus
             Exit Function
         Else
             If Not TorNmGet(iH1_DaikoK, 8) Then
                With CisFun
                     .MB_Lines = 4
                     .MB_MSG(2) = "    契約先未登録         "
                     .MB_Button = Error
                     .MBOX
                End With
                iH1_DaikoK.SetFocus
                Exit Function
             End If
         End If
    End If
'
    gSL_Select = "SELECT"
    gSL_Select = gSL_Select & " * From TNS送信設定テーブル"
    gSL_Select = gSL_Select & " Where 送信先 = '" & RTrim$(iH1_KTorcd) & "'"
    gSL_Select = gSL_Select & " And   FCODE ='" & RTrim$(iH1_FileCd) & "'"
    gSL_Select = gSL_Select & " And   代行送信先 = '" & RTrim$(iH1_DaikoK) & "'"                    '2007/12/26 ADD
    gSL_Select = gSL_Select & " And   送信先区分 = " & CisFun.Val2(OpH1_SendKb(0).Tag) & ""         '2007/12/26 ADD
    gSL_Select = gSL_Select & " And   枝番 = 0"
    
    If DBInput("INV") Then
        If SyoriKB = "ADD" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    TNS送信データ設定情報登録済み         "
                .MB_Button = Error
                .MBOX
            End With
            iH1_KTorcd.SetFocus
            Exit Function
        End If
    Else
        If SyoriKB <> "ADD" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    TNS送信データ設定情報未登録         "
                .MB_Button = Error
                .MBOX
            End With
            iH1_KTorcd.SetFocus
            Exit Function
        End If
        Call B1Chk_CrtKB1_Click
        Call B1Chk_CrtKB2_Click
        
        gStr = Mid(RTrim(gCompany), 1, 4)
        iB1_SMoto = Format(gStr, "@@@@@")
        
    End If
    
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
    With CisFun
        .MB_Lines = 4
        .MB_Button = Error
    End With
    '業務名
    If RTrim(iB1_FileName) = "" Then
        CisFun.MB_MSG(2) = "    業務名（ＴＮＳファイル名称）を入力して下さい         "
        CisFun.MBOX
        iB1_FileName.SetFocus
        Exit Function
    End If
    '宛先
    If RTrim(iB1_Atesaki) = "" Then
        CisFun.MB_MSG(2) = "    　　宛先を入力して下さい         "
        CisFun.MBOX
        iB1_Atesaki.SetFocus
        Exit Function
    End If
    '発信元
    If RTrim(iB1_SMoto) = "" Then
        CisFun.MB_MSG(2) = "    　　発信元を入力して下さい         "
        CisFun.MBOX
        iB1_SMoto.SetFocus
        Exit Function
    End If
    '<< 送信 = 自動　の時 >>
    If B1Chk_SendKB.Value = 1 Then
        '識別子
        If RTrim(iB1_Shiki) = "" Then
            CisFun.MB_MSG(2) = "    　　識別子を入力して下さい         "
            CisFun.MBOX
            iB1_Shiki.SetFocus
            Exit Function
        End If
        'TNS宛先
        If RTrim(iB1_TNSate) = "" Then
            CisFun.MB_MSG(2) = "    　　TNS宛先を入力して下さい         "
            CisFun.MBOX
            iB1_TNSate.SetFocus
            Exit Function
        End If
    End If
    'ﾚｺｰﾄﾞ長
    If CisFun.Val2(iB1_Leng) = 0 Then
        CisFun.MB_MSG(2) = "    　　ﾚｺｰﾄﾞ長を入力して下さい         "
        CisFun.MBOX
        iB1_Leng.SetFocus
        Exit Function
    End If
    '送信開始日                                                         '2007/12/26 ADD
    If RTrim(iB1_SSYmd) <> "" Then
        If Not iB1_SSYmd.cChkResult Then
            CisFun.MB_MSG(2) = "    　　正しい送信開始日を入力して下さい         "
            CisFun.MBOX
            iB1_SSYmd.SetFocus
            Exit Function
        End If
    End If
    'TNS送信ﾃﾞｰﾀ作成先
    If RTrim(iB1_TextDir) = "" Then
        CisFun.MB_MSG(2) = "    TNS送信ﾃﾞｰﾀ作成先を入力して下さい         "
        CisFun.MBOX
        iB1_TextDir.SetFocus
        Exit Function
    End If
' ディレクトリ存在確認
    With CisFun
        .FE_Path = RTrim(iB1_TextDir)
        .DirCheck Msg_Crt
        If Not .DirCheck Then
            MsgBox "　(( TNS送信ﾃﾞｰﾀ作成先 )) 存在する作成先を入力して下さい。　", vbExclamation, "必須入力ｴﾗｰ"
            iB1_TextDir.SetFocus
            Exit Function
        End If
    End With
    'TNS送信ﾌｧｲﾙ名
'   2006/03/09
'    If RTrim(iB1_TextID) = "" Then
'        CisFun.MB_MSG(2) = "    TNS送信ﾌｧｲﾙ名を入力して下さい         "
'        CisFun.MBOX
'        iB1_TextID.SetFocus
'        Exit Function
'    End If
'
    'TNS送信ﾃﾞｰﾀﾊﾞｯｸｱｯﾌﾟ作成先
    If RTrim(iB1_BackDir) = "" Then
        CisFun.MB_MSG(2) = "    TNS送信ﾃﾞｰﾀﾊﾞｯｸｱｯﾌﾟ作成先を入力して下さい         "
        CisFun.MBOX
        iB1_BackDir.SetFocus
        Exit Function
    End If
' ディレクトリ存在確認
    With CisFun
        .FE_Path = RTrim(iB1_BackDir)
        .DirCheck Msg_Crt
        If Not .DirCheck Then
            MsgBox "　(( TNS送信ﾃﾞｰﾀﾊﾞｯｸｱｯﾌﾟ作成先 )) 存在する作成先を入力して下さい。　", vbExclamation, "必須入力ｴﾗｰ"
            iB1_BackDir.SetFocus
            Exit Function
        End If
    End With
    
    Body1Chk = True
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    
    Call HeadBodyClear("BC")
'
    CisComb.ComboSet
    
    '-------------------------------------------------- 2007/03/08 Insert
    Back_Hinbn.Enabled = False
    If iH1_FileCd = "C04" Or iH1_FileCd = "R07" Then
       Back_Hinbn.Enabled = True
    End If
    '-------------------------------------------------- 2007/03/08 Insert End
    
    On Error GoTo DBInput_Err:
    
    If Not TSSRead(gSL_Select, 1) Then
        If InputKey = "Next" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    最終レコードです            "
                .MB_Title = "【前進】"
                .MB_Button = OK
                .MBOX
            End With
            GoTo DBInput_Ed
        End If
        If InputKey = "Back" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    先頭レコードです            "
                .MB_Title = "【後退】"
                .MB_Button = OK
                .MBOX
            End With
        End If
        GoTo DBInput_Ed
    Else
        Call BodySet("")
        DBInput = True
    End If
    
DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Sub BodySet(SetType As String)
'
'   ( 項目移送 )
    With TSS
        iH1_KTorcd = RTrim(.送信先)
        Call iH1_KTorcd_LostFocus
        iH1_FileCd = RTrim(.FCODE)
    '+------------------------------------------------ 2007/12/26 Insert
        iH1_DaikoK = RTrim(.代行送信先)
        Call iH1_DaikoK_LostFocus
        If .送信先区分 = 0 Then
            OpH1_SendKb(0).Value = True
        Else
            OpH1_SendKb(1).Value = True
        End If
    '+------------------------------------------------ 2007/12/26 Insert End
        
        iB1_FileName = RTrim(.ファイル名称)
        B1Chk_SendKB.Value = .送信
        B1Chk_ZeroKB.Value = .ゼロ件データ
        iB1_Atesaki = RTrim(.宛先)
        iB1_SMoto = RTrim(.発信元)
        iB1_Shiki = RTrim(.識別子)
        iB1_TNSate = RTrim(.TNS宛先)
        iB1_Leng = .レコード長
        CisComb.Code = .処理区分
        CisComb.TextGet
        If .TNSフォーマット = 0 Then
            B1Op_TNSFormat(0).Value = True
        Else
            B1Op_TNSFormat(1).Value = True
        End If
        If .出力区分 = 0 Then
            B1Op_OutKB(0).Value = True
        Else
            B1Op_OutKB(1).Value = True
        End If
        
        iB1_TextID = RTrim(.TNS送信ファイル名)
        iB1_Del = .削除日数
        B1Chk_CrtKB1.Value = .作成先区分1
        Call B1Chk_CrtKB1_Click
        B1Chk_CrtKB2.Value = .作成先区分2
        Call B1Chk_CrtKB2_Click
        
    '+------------------------------------------------ 2007/02/08 Insert
        B1Chk_TNSSys.Value = .TNSシステム起動
    '+------------------------------------------------ 2007/02/08 Insert End
    '+------------------------------------------------ 2007/03/08 Insert
        B1Chk_HINChk.Value = .品番CHK
        B1Chk_HINList.Value = .品番ERRLIST
    '+------------------------------------------------ 2007/03/08 Insert
        iB1_SSYmd = .送信開始日                         '2007/12/26 ADD
        B1Chk_UkeChk.Value = .受入CHK                   '2007/12/26 ADD
    
    End With
    
End Sub
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'*-------------------------------------------------*
'*      デ ー タ 更 新  確 認 メ ッ セ ー ジ
'*-------------------------------------------------*
Private Function CheckMessage() As Boolean
    CheckMessage = False
    
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     " & SyoriNM & "処理を行います      "
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Title = SyoriNM & "処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    CheckMessage = True
End Function
'+----------------------+
'+      削 除 処 理      +
'+----------------------+
Private Function DBDelete()
    DBDelete = False
    
'   ( 確認メッセージ )
    If Not CheckMessage Then Exit Function
    
'   ( 削除処理 )
    gSL_Select = " DELETE From TNS送信設定テーブル"
    gSL_Select = gSL_Select & " Where 送信先 = '" & RTrim$(iH1_KTorcd) & "'"
    gSL_Select = gSL_Select & " And   FCODE ='" & RTrim$(iH1_FileCd) & "'"
    gSL_Select = gSL_Select & " And   代行送信先 ='" & RTrim$(iH1_DaikoK) & "'"             '2007/12/26 ADD
    gSL_Select = gSL_Select & " And   送信先区分 = " & CisFun.Val2(OpH1_SendKb(0).Tag)      '2007/12/26 ADD
    gSL_Select = gSL_Select & " And   枝番 = 0"
    
    On Error GoTo DBDelete_Err
    CisDB.DBTran TransBegin                                     ' ﾄﾗﾝｻﾞｸｼｮﾝ 開始
    
    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With
    
    CisDB.DBTran TransCommit                                    ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    On Error GoTo 0
    
    DBDelete = True
    Exit Function
    
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Static UpdFlg               As Boolean
    
'   ( 確認メッセージ )
    If Not CheckMessage Then Exit Function
    
    'TNS作成タイプ
    CisComb.CodeGet
    
    gSL_Select = "Select * From 名称マスタ"
    gSL_Select = gSL_Select & " Where 区分名称 = 'TNS作成タイプ'"
    gSL_Select = gSL_Select & "   And 値   = '" & CisComb.Code & "'"
    gSL_Select = gSL_Select & "   And 区分ﾀｲﾌﾟ = ''"
    If Not SYMRead(gSL_Select, 1) Then
       Call ItemsClearSYM
    End If
'   ( データ存在確認 )
    gSL_Select = "Select * FROM TNS送信設定テーブル"
    gSL_Select = gSL_Select & " Where 送信先 = '" & RTrim$(iH1_KTorcd) & "'"
    gSL_Select = gSL_Select & " And   FCODE ='" & RTrim$(iH1_FileCd) & "'"
    gSL_Select = gSL_Select & " And   代行送信先 ='" & RTrim$(iH1_DaikoK) & "'"             '2007/12/26 ADD
    gSL_Select = gSL_Select & " And   送信先区分 = " & CisFun.Val2(OpH1_SendKb(0).Tag)      '2007/12/26 ADD
    gSL_Select = gSL_Select & " And   枝番 = 0"
    If Not TSSRead(gSL_Select, 1) Then
        Call ItemsClearTSS
        UpdFlg = True
    Else
        UpdFlg = False
    End If
        
'   ( 項目移送 )
    With TSS
        .送信先 = RTrim(iH1_KTorcd)
        .FCODE = RTrim(iH1_FileCd)
        .代行送信先 = RTrim(iH1_DaikoK)                                                     '2007/12/26 ADD
        .送信先区分 = CisFun.Val2(OpH1_SendKb(0).Tag)                                       '2007/12/26 ADD
        .枝番 = 0
        .ファイル名称 = iB1_FileName
        .送信 = B1Chk_SendKB.Value
        .ゼロ件データ = B1Chk_ZeroKB.Value
        .宛先 = iB1_Atesaki
        .発信元 = iB1_SMoto
        .識別子 = iB1_Shiki
        .TNS宛先 = iB1_TNSate
        .レコード長 = CisFun.Val2(iB1_Leng)
        .処理区分 = SYM.値
        .TNSタイプ = SYM.数字3
        If B1Op_TNSFormat(0).Value Then
            .TNSフォーマット = 0
        Else
            .TNSフォーマット = 1
        End If
        If B1Op_OutKB(0).Value Then
            .出力区分 = 0
        Else
            .出力区分 = 1
        End If
        .TNS送信ファイル名 = iB1_TextID
        .削除日数 = CisFun.Val2(iB1_Del)
        .作成先区分1 = B1Chk_CrtKB1.Value
        .作成先区分2 = B1Chk_CrtKB2.Value
    
    '+------------------------------------------------ 2007/02/08 Insert
        .TNSシステム起動 = B1Chk_TNSSys.Value
    '+------------------------------------------------ 2007/02/08 Insert End
    '+------------------------------------------------ 2007/03/08 Insert
        .品番CHK = B1Chk_HINChk.Value
        .品番ERRLIST = B1Chk_HINList.Value
    '+------------------------------------------------ 2007/03/08 Insert End
        .受入CHK = B1Chk_UkeChk.Value                   '2007/12/26 ADD
        .送信開始日 = RTrim(iB1_SSYmd)                  '2007/12/26 ADD
    End With
    
    On Error GoTo DBPut_Err
    CisDB.DBTran TransBegin                                     ' ﾄﾗﾝｻﾞｸｼｮﾝ 開始
    
' ( 追加処理 )
    If UpdFlg Then
        Call TSSInsert
    Else
' ( 更新処理 )
        Call TSSUpdate
    End If
    
'データ作成先 INIファイル更新
    With CisFun
        .INI_FileID = gIniExe & gwsIniName
        
        .INI_Section = "TNSSENDDIR"
        If TSS.作成先区分1 = 0 Then
            .INI_Key = "COMMON"
        Else
            .INI_Key = RTrim(iH1_KTorcd) & "_" & RTrim(iH1_FileCd)
        End If
        .INI_String = iB1_TextDir
        .PutIni
        
        .INI_Section = "TNSBACKDIR"
        If TSS.作成先区分2 = 0 Then
            .INI_Key = "COMMON"
        Else
            .INI_Key = RTrim(iH1_KTorcd) & "_" & RTrim(iH1_FileCd)
        End If
        .INI_String = iB1_BackDir
        .PutIni
    
    
        .INI_FileID = gIniExe & gIniName
    End With
    
    CisDB.DBTran TransCommit                                    ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    On Error GoTo 0
    
    DBPut = True
    Exit Function
    
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+---------------------------------------+
'+      名称マスタ初期設定
'+---------------------------------------+
Private Sub MeisyoInit()
' TNSデータ作成ﾀｲﾌﾟ
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "TNS作成タイプ"
        .区分桁数 = 2
        .値 = "00"
        .制御区分 = "X"
        .値名称 = "支給品入荷データ(紡織)"
        .英数字1 = "直直支給#": .数字1 = 7
        .英数字2 = "自給#": .数字2 = 5
        .英数字3 = "TNSタイプ": .数字3 = 1
        .作成者 = gTanto
    End With
    If Not SYMInitCreate(True) Then
    '+-------------------------------------------- 2007/02/08 Insert
        For gInt = 1 To 1
            With SYMCreate
                .値 = Format(gInt, "00")
                Select Case gInt
                     Case 1
                            .値名称 = "支給品入荷データ(車体)"
                            .英数字1 = "直直支給#": .数字1 = 1
                            .英数字2 = "自給#": .数字2 = 5
                            .英数字3 = "TNSタイプ": .数字3 = 2
                            .作成者 = gTanto
                     Case 2
                            .値名称 = "支給品入荷データ(トヨタ)"
                            .英数字1 = "直直支給#": .数字1 = 7
                            .英数字2 = "自給#": .数字2 = 7
                            .英数字3 = "TNSタイプ": .数字3 = 2
                            .作成者 = gTanto
                End Select
            End With
            Call SYMInitCreate
        Next gInt
    '+-------------------------------------------- 2007/02/08 Insert End
    End If
End Sub







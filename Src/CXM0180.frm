VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "Cis3D_6.0.ocx"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CisBtn_6.0.ocx"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CisText_6.0.ocx"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdodc.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CisYMDwB_6.0.ocx"
Begin VB.Form CXM0180 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "購入単価マスタ"
   ClientHeight    =   11190
   ClientLeft      =   540
   ClientTop       =   4500
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   Picture         =   "CXM0180.frx":0000
   ScaleHeight     =   11190
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin VB.PictureBox H0_Area1 
      Height          =   1680
      Left            =   870
      ScaleHeight     =   1620
      ScaleWidth      =   1275
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1650
      Width           =   1335
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   1500
         Left            =   60
         Top             =   60
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   2646
         ForeColor       =   16711680
         Caption         =   "単価区分"
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
         Begin VB.OptionButton H0Op_Tanku 
            BackColor       =   &H00C0C0C0&
            Caption         =   "部品"
            Height          =   345
            Index           =   0
            Left            =   180
            TabIndex        =   36
            Top             =   540
            Width           =   735
         End
         Begin VB.OptionButton H0Op_Tanku 
            BackColor       =   &H00C0FFC0&
            Caption         =   "材料"
            Height          =   345
            Index           =   1
            Left            =   180
            TabIndex        =   35
            Top             =   900
            Value           =   -1  'True
            Width           =   735
         End
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   165
      Top             =   945
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   635
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   810
      Left            =   7260
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
      Height          =   270
      Left            =   -30
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   420
      Width           =   855
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   7185
      Left            =   870
      ScaleHeight     =   7125
      ScaleWidth      =   14130
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3405
      Width           =   14190
      Begin Cis3D_v60.CIS3D Panel_SonoTa 
         Height          =   1020
         Left            =   120
         Top             =   1170
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   1799
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
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   435
            Index           =   12
            Left            =   2985
            Top             =   15
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   767
            ForeColor       =   16711680
            Caption         =   " "
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
            Begin CisText_V60.CisText iB1_EtcMsg 
               Height          =   360
               Index           =   0
               Left            =   90
               TabIndex        =   12
               Top             =   45
               Width           =   1590
               _ExtentX        =   2805
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
               Text            =   "XXXXXXXXXXXX"
               MaxLength       =   12
               IMEMode         =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   540
            Index           =   13
            Left            =   2985
            Top             =   465
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   953
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   60
            Begin CisText_V60.CisText iB1_EtcTan 
               Height          =   375
               Index           =   0
               Left            =   90
               TabIndex        =   13
               Top             =   90
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   661
               cFormat         =   "#,###,##0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#,###,##0.00"
               MaxLength       =   12
               cDataType       =   2
               cDataReplace    =   1
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "######0.00"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   435
            Index           =   18
            Left            =   6495
            Top             =   15
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   767
            ForeColor       =   16711680
            Caption         =   " "
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
            Begin CisText_V60.CisText iB1_EtcMsg 
               Height          =   375
               Index           =   2
               Left            =   75
               TabIndex        =   16
               Top             =   45
               Width           =   1590
               _ExtentX        =   2805
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
               Text            =   "XXXXXXXXXXXX"
               MaxLength       =   12
               IMEMode         =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   540
            Index           =   19
            Left            =   6495
            Top             =   465
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   953
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   60
            Begin CisText_V60.CisText iB1_EtcTan 
               Height          =   375
               Index           =   2
               Left            =   90
               TabIndex        =   17
               Top             =   90
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   661
               cFormat         =   "#,###,##0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#,###,##0.00"
               MaxLength       =   12
               cDataType       =   2
               cDataReplace    =   1
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "######0.00"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   435
            Index           =   20
            Left            =   8250
            Top             =   15
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   767
            ForeColor       =   16711680
            Caption         =   " "
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
            Begin CisText_V60.CisText iB1_EtcMsg 
               Height          =   375
               Index           =   3
               Left            =   75
               TabIndex        =   18
               Top             =   45
               Width           =   1590
               _ExtentX        =   2805
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
               Text            =   "XXXXXXXXXXXX"
               MaxLength       =   12
               IMEMode         =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   540
            Index           =   21
            Left            =   8250
            Top             =   465
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   953
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   60
            Begin CisText_V60.CisText iB1_EtcTan 
               Height          =   375
               Index           =   3
               Left            =   75
               TabIndex        =   19
               Top             =   90
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   661
               cFormat         =   "#,###,##0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#,###,##0.00"
               MaxLength       =   12
               cDataType       =   2
               cDataReplace    =   1
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "######0.00"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   990
            Index           =   1
            Left            =   15
            Top             =   15
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   1746
            BackColor       =   16777152
            ForeColor       =   0
            Caption         =   "その他単価内訳"
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
            cPositionY      =   30
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   435
            Index           =   22
            Left            =   10005
            Top             =   15
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   767
            ForeColor       =   16711680
            Caption         =   " "
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
            Begin CisText_V60.CisText iB1_EtcMsg 
               Height          =   375
               Index           =   4
               Left            =   75
               TabIndex        =   20
               Top             =   45
               Width           =   1590
               _ExtentX        =   2805
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
               Text            =   "XXXXXXXXXXXX"
               MaxLength       =   12
               IMEMode         =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   540
            Index           =   23
            Left            =   10005
            Top             =   465
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   953
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   60
            Begin CisText_V60.CisText iB1_EtcTan 
               Height          =   375
               Index           =   4
               Left            =   75
               TabIndex        =   21
               Top             =   90
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   661
               cFormat         =   "#,###,##0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#,###,##0.00"
               MaxLength       =   12
               cDataType       =   2
               cDataReplace    =   1
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "######0.00"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   435
            Index           =   14
            Left            =   4740
            Top             =   15
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   767
            ForeColor       =   16711680
            Caption         =   " "
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
            Begin CisText_V60.CisText iB1_EtcMsg 
               Height          =   375
               Index           =   1
               Left            =   90
               TabIndex        =   14
               Top             =   45
               Width           =   1590
               _ExtentX        =   2805
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
               Text            =   "XXXXXXXXXXXX"
               MaxLength       =   12
               IMEMode         =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   540
            Index           =   15
            Left            =   4740
            Top             =   465
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   953
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   60
            Begin CisText_V60.CisText iB1_EtcTan 
               Height          =   375
               Index           =   1
               Left            =   90
               TabIndex        =   15
               Top             =   90
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   661
               cFormat         =   "#,###,##0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#,###,##0.00"
               MaxLength       =   12
               cDataType       =   2
               cDataReplace    =   1
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "######0.00"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   435
            Index           =   16
            Left            =   11760
            Top             =   15
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   767
            ForeColor       =   16711680
            Caption         =   " "
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
            Begin CisText_V60.CisText iB1_EtcMsg 
               Height          =   375
               Index           =   5
               Left            =   75
               TabIndex        =   22
               Top             =   45
               Width           =   1590
               _ExtentX        =   2805
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
               Text            =   "XXXXXXXXXXXX"
               MaxLength       =   12
               IMEMode         =   4
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   540
            Index           =   17
            Left            =   11760
            Top             =   465
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   953
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   60
            Begin CisText_V60.CisText iB1_EtcTan 
               Height          =   375
               Index           =   5
               Left            =   75
               TabIndex        =   23
               Top             =   90
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   661
               cFormat         =   "#,###,##0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#,###,##0.00"
               MaxLength       =   12
               cDataType       =   2
               cDataReplace    =   1
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "######0.00"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   435
            Index           =   24
            Left            =   1950
            Top             =   15
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   767
            ForeColor       =   0
            Caption         =   "コメント"
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
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   555
            Index           =   25
            Left            =   1950
            Top             =   450
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   979
            ForeColor       =   0
            Caption         =   "単価"
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
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   990
            Index           =   2
            Left            =   13515
            Top             =   15
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   1746
            ForeColor       =   0
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
            cPositionX      =   80
            cPositionY      =   -170
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   4890
         Left            =   15
         Top             =   2235
         Width           =   14130
         _ExtentX        =   24924
         _ExtentY        =   8625
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
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   450
            Left            =   15
            Top             =   30
            Width           =   14100
            _ExtentX        =   24871
            _ExtentY        =   794
            BackColor       =   12648384
            ForeColor       =   16711680
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   20.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cPositionY      =   -120
            Begin Cis3D_v60.CIS3D CIS3D10 
               Height          =   345
               Left            =   12240
               Top             =   45
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   609
               Caption         =   "件 数"
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
            End
            Begin Cis3D_v60.CIS3D CIS3D8 
               Height          =   330
               Left            =   4290
               Top             =   45
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   582
               BackColor       =   12648384
               ForeColor       =   16711680
               Caption         =   "単  価  履  歴"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   12648384
               cBoderColor2    =   12648384
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cBoderStyle     =   2
            End
            Begin Cis3D_v60.CIS3D B1lb_DispSu 
               Height          =   345
               Left            =   13110
               Top             =   45
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   609
               Caption         =   "#,###"
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
               cAlingnment     =   4
               cBoderStyle     =   1
               cPositionX      =   -60
            End
         End
         Begin vsFlexLib.vsFlexArray vsGrid1 
            Height          =   4305
            Left            =   120
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   480
            Width           =   13920
            _Version        =   196608
            _ExtentX        =   24553
            _ExtentY        =   7594
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
            FormatString    =   $"CXM0180.frx":0342
            Rows            =   11
            Cols            =   24
            BackColor       =   12648447
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   0
            RowHeightMin    =   360
            GridLines       =   2
            SelectionMode   =   1
            BackColorAlternate=   12648447
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1035
         Index           =   0
         Left            =   120
         Top             =   90
         Width           =   13965
         _ExtentX        =   24633
         _ExtentY        =   1826
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
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   435
            Index           =   0
            Left            =   735
            Top             =   15
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   767
            ForeColor       =   16711680
            Caption         =   "単  価"
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
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   540
            Index           =   0
            Left            =   2430
            Top             =   450
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   953
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            Begin VB.CheckBox B1Ch_Kari 
               BackColor       =   &H00C0C0C0&
               Caption         =   "正"
               Height          =   270
               Left            =   45
               TabIndex        =   8
               Top             =   150
               Width           =   480
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   540
            Index           =   1
            Left            =   735
            Top             =   450
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   953
            BackColor       =   14737632
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   60
            Begin CisText_V60.CisText iB1_Tanka 
               Height          =   375
               Left            =   60
               TabIndex        =   7
               Top             =   90
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   661
               cFormat         =   "#,###,##0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "#,###,##0.00"
               MaxLength       =   12
               cDataType       =   2
               cDataReplace    =   1
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "######0.00"
               cILength        =   7
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   435
            Left            =   2430
            Top             =   15
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   767
            ForeColor       =   16711680
            Caption         =   "仮単"
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
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   435
            Index           =   2
            Left            =   3000
            Top             =   15
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   767
            ForeColor       =   16711680
            Caption         =   "適用開始日"
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
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   540
            Index           =   3
            Left            =   3000
            Top             =   450
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   953
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   60
            Begin CisYMD_With_Btn.CisYMDwB iB1_TSYmd 
               Height          =   360
               Left            =   90
               TabIndex        =   9
               Top             =   90
               Width           =   1800
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
               Object.Height          =   360
               Object.Width           =   1800
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
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   435
            Index           =   4
            Left            =   4980
            Top             =   15
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   767
            ForeColor       =   16711680
            Caption         =   "適用終了日"
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
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   540
            Index           =   5
            Left            =   4980
            Top             =   450
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   953
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   60
            Begin CisYMD_With_Btn.CisYMDwB iB1_TEYmd 
               Height          =   360
               Left            =   90
               TabIndex        =   10
               Top             =   90
               Width           =   1800
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
               Object.Height          =   360
               Object.Width           =   1800
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
         End
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   975
            Index           =   0
            Left            =   15
            Top             =   15
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   1720
            BackColor       =   16761024
            ForeColor       =   0
            Caption         =   "最新単価"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   14.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cPositionY      =   -120
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   540
            Index           =   1
            Left            =   6960
            Top             =   450
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   953
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            Begin VB.CheckBox B1Ch_UMu 
               BackColor       =   &H00C0C0C0&
               Caption         =   "有償"
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
               Left            =   45
               TabIndex        =   11
               Top             =   150
               Width           =   765
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D11 
            Height          =   435
            Left            =   6960
            Top             =   15
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   767
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
            cAlingnment     =   6
            cPositionY      =   60
         End
         Begin Cis3D_v60.CIS3D Panel_Uchi 
            Height          =   975
            Left            =   7815
            Top             =   15
            Width           =   6120
            _ExtentX        =   10795
            _ExtentY        =   1720
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
            cAlingnment     =   6
            cPositionY      =   60
            Begin Cis3D_v60.CIS3D CIS3D1 
               Height          =   975
               Index           =   0
               Left            =   0
               Top             =   0
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   1720
               BackColor       =   14737632
               ForeColor       =   0
               Caption         =   "単価内訳"
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
               cPositionX      =   80
               cPositionY      =   -170
            End
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   435
               Index           =   6
               Left            =   615
               Top             =   0
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   767
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "部 品 単 価"
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
            End
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   540
               Index           =   7
               Left            =   615
               Top             =   435
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   953
               BackColor       =   14737632
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cPositionY      =   60
               Begin CisText_V60.CisText iB1_BTanka 
                  Height          =   375
                  Left            =   75
                  TabIndex        =   41
                  Top             =   90
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   661
                  cFormat         =   "#,###,##0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#,###,##0.00"
                  MaxLength       =   12
                  cDataType       =   2
                  cDataReplace    =   1
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "######0.00"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   435
               Index           =   8
               Left            =   2310
               Top             =   0
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   767
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "加 工 単 価"
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
            End
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   540
               Index           =   9
               Left            =   2310
               Top             =   435
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   953
               BackColor       =   14737632
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cPositionY      =   60
               Begin CisText_V60.CisText iB1_KTanka 
                  Height          =   375
                  Left            =   60
                  TabIndex        =   42
                  Top             =   90
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   661
                  cFormat         =   "#,###,##0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#,###,##0.00"
                  MaxLength       =   12
                  cDataType       =   2
                  cDataReplace    =   1
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "######0.00"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   435
               Index           =   10
               Left            =   4005
               Top             =   0
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   767
               BackColor       =   14737632
               ForeColor       =   16711680
               Caption         =   "その他単価"
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
            End
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   540
               Index           =   11
               Left            =   4005
               Top             =   435
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   953
               BackColor       =   14737632
               ForeColor       =   16711680
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
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cPositionY      =   60
               Begin CisText_V60.CisText iB1_STanka 
                  Height          =   375
                  Left            =   60
                  TabIndex        =   43
                  Top             =   90
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   661
                  cFormat         =   "#,###,##0.00"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#,###,##0.00"
                  MaxLength       =   12
                  cDataType       =   2
                  cDataReplace    =   1
                  cFcDecimal      =   2
                  cFdAutoFormat   =   1
                  cGFormat        =   "######0.00"
                  cILength        =   7
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D1 
               Height          =   975
               Index           =   1
               Left            =   5700
               Top             =   0
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   1720
               BackColor       =   14737632
               ForeColor       =   0
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
               cPositionX      =   80
               cPositionY      =   -170
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   975
            Left            =   7815
            Top             =   15
            Width           =   6120
            _ExtentX        =   10795
            _ExtentY        =   1720
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
            cAlingnment     =   6
            cPositionY      =   60
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      Height          =   1680
      Left            =   2250
      ScaleHeight     =   1620
      ScaleWidth      =   12750
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1650
      Width           =   12810
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   1500
         Left            =   90
         Top             =   60
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   2646
         ForeColor       =   16711680
         Caption         =   "　品　　番"
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
         cPositionY      =   60
         Begin Cis3D_v60.CIS3D H1_ZaiBox 
            Height          =   375
            Left            =   210
            Top             =   480
            Visible         =   0   'False
            Width           =   6150
            _ExtentX        =   10848
            _ExtentY        =   661
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
            cFont3D         =   2
            cBoderStyle     =   1
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   345
               Index           =   27
               Left            =   5130
               Top             =   15
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   609
               ForeColor       =   16711680
               Caption         =   "長"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cPositionY      =   60
            End
            Begin Cis3D_v60.CIS3D DH1_Long 
               Height          =   90
               Left            =   5310
               Top             =   15
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   159
               BackColor       =   12640511
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
               cBoderColor1    =   12632256
               cBoderColor2    =   12640511
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   1
               cBoderStyle     =   2
               cPositionX      =   90
            End
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   345
               Index           =   28
               Left            =   4140
               Top             =   15
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   609
               ForeColor       =   16711680
               Caption         =   "幅"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cPositionY      =   60
            End
            Begin Cis3D_v60.CIS3D DH1_Width 
               Height          =   90
               Left            =   4320
               Top             =   15
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   159
               BackColor       =   12640511
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
               cBoderColor1    =   12632256
               cBoderColor2    =   12640511
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   1
               cBoderStyle     =   2
               cPositionX      =   90
            End
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   345
               Index           =   29
               Left            =   3015
               Top             =   15
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   609
               ForeColor       =   16711680
               Caption         =   "板厚"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   3
               cPositionY      =   60
            End
            Begin Cis3D_v60.CIS3D DH1_Itatu 
               Height          =   90
               Left            =   3330
               Top             =   15
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   159
               BackColor       =   12640511
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
               cBoderColor1    =   12632256
               cBoderColor2    =   12640511
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   1
               cBoderStyle     =   2
               cPositionX      =   90
            End
            Begin Cis3D_v60.CIS3D UC_3D15 
               Height          =   345
               Index           =   26
               Left            =   15
               Top             =   15
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   609
               ForeColor       =   16711680
               Caption         =   "材質"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   3
               cPositionY      =   60
            End
            Begin Cis3D_v60.CIS3D DH1_Zaist 
               Height          =   90
               Left            =   330
               Top             =   15
               Width           =   2820
               _ExtentX        =   4974
               _ExtentY        =   159
               BackColor       =   12640511
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
               cBoderColor1    =   12632256
               cBoderColor2    =   12640511
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cAlingnment     =   1
               cBoderStyle     =   2
               cPositionX      =   90
            End
            Begin CisText_V60.CisText iH1_Long 
               Height          =   330
               Left            =   5340
               TabIndex        =   4
               Top             =   60
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
               cFormat         =   "###0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###0.00"
               MaxLength       =   7
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "###0.00"
               cILength        =   4
            End
            Begin CisText_V60.CisText iH1_Width 
               Height          =   300
               Left            =   4350
               TabIndex        =   3
               Top             =   60
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   529
               cFormat         =   "###0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###0.00"
               MaxLength       =   7
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "###0.00"
               cILength        =   4
            End
            Begin CisText_V60.CisText iH1_Itatu 
               Height          =   300
               Left            =   3360
               TabIndex        =   2
               Top             =   60
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   529
               cFormat         =   "###0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###0.00"
               MaxLength       =   7
               cDataType       =   2
               cDataReplace    =   1
               cFbComma        =   0
               cFcDecimal      =   2
               cFdAutoFormat   =   1
               cGFormat        =   "###0.00"
               cILength        =   4
            End
            Begin CisText_V60.CisText iH1_Zaist 
               Height          =   300
               Left            =   360
               TabIndex        =   1
               Top             =   60
               Width           =   2730
               _ExtentX        =   4815
               _ExtentY        =   529
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   25
               IMEMode         =   2
            End
         End
         Begin VB.OptionButton H1Op_Hinku 
            BackColor       =   &H00C0C0C0&
            Caption         =   "材料/寸法"
            Height          =   180
            Index           =   3
            Left            =   5160
            TabIndex        =   40
            Top             =   120
            Width           =   1200
         End
         Begin VB.OptionButton H1Op_Hinku 
            BackColor       =   &H00C0C0C0&
            Caption         =   "発注品番"
            Height          =   180
            Index           =   0
            Left            =   1440
            TabIndex        =   39
            Top             =   120
            Width           =   1200
         End
         Begin VB.OptionButton H1Op_Hinku 
            BackColor       =   &H00C0C0C0&
            Caption         =   "自社品番"
            Height          =   180
            Index           =   2
            Left            =   3960
            TabIndex        =   38
            Top             =   120
            Width           =   1200
         End
         Begin VB.OptionButton H1Op_Hinku 
            BackColor       =   &H00C0C0C0&
            Caption         =   "受給品番"
            Height          =   180
            Index           =   1
            Left            =   2760
            TabIndex        =   37
            Top             =   120
            Width           =   1200
         End
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   210
            TabIndex        =   0
            Top             =   360
            Width           =   6255
            _ExtentX        =   11033
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   50
            IMEMode         =   2
            cAutoEject      =   0   'False
         End
         Begin Cis3D_v60.CIS3D H1lb_HinBn 
            Height          =   225
            Index           =   3
            Left            =   930
            Top             =   1200
            Width           =   5400
            _ExtentX        =   9525
            _ExtentY        =   397
            Caption         =   "SPH440-0D2.9Ｘ238Ｘ1219 5XXXX3XXXXXXXXX4XXXXXXXXX5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
            cPositionX      =   60
         End
         Begin Cis3D_v60.CIS3D H1lb_HinBn 
            Height          =   225
            Index           =   1
            Left            =   930
            Top             =   780
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   397
            Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXX5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
            cPositionX      =   60
         End
         Begin Cis3D_v60.CIS3D b_HinNM 
            Height          =   225
            Index           =   1
            Left            =   210
            Top             =   780
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            Caption         =   "受給品番"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D b_HinNM 
            Height          =   225
            Index           =   3
            Left            =   210
            Top             =   1200
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            Caption         =   "材　　　料"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D H1lb_HinBn 
            Height          =   225
            Index           =   2
            Left            =   930
            Top             =   990
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   397
            Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXX5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
            cPositionX      =   60
         End
         Begin Cis3D_v60.CIS3D b_HinNM 
            Height          =   225
            Index           =   2
            Left            =   210
            Top             =   990
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            Caption         =   "自社品番"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D H1lb_HinBn 
            Height          =   225
            Index           =   0
            Left            =   930
            Top             =   660
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   397
            Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXX5"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
            cPositionX      =   60
         End
         Begin Cis3D_v60.CIS3D b_HinNM 
            Height          =   225
            Index           =   0
            Left            =   210
            Top             =   660
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
            Caption         =   "発注品番"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cBoderStyle     =   2
         End
         Begin Cis3D_v60.CIS3D H1lb_ZaiKan 
            Height          =   255
            Left            =   4620
            Top             =   720
            Visible         =   0   'False
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   450
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
            cFont3D         =   2
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D H1_TorMsg 
         Height          =   1500
         Left            =   6825
         Top             =   60
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   2646
         ForeColor       =   16711680
         Caption         =   "仕 入 先"
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
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   990
            _ExtentX        =   1746
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
         End
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   240
            Top             =   720
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
      Begin Cis3D_v60.CIS3D H1_KTorMsg 
         Height          =   1500
         Left            =   9750
         Top             =   60
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   2646
         ForeColor       =   16711680
         Caption         =   "契 約 先"
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
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   990
            _ExtentX        =   1746
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
         End
         Begin Cis3D_v60.CIS3D H1lb_KTornm 
            Height          =   360
            Left            =   240
            Top             =   720
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
      Left            =   120
      Top             =   120
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 購入単価マスタ メンテ  】"
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
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'なし
         Height          =   345
         Index           =   1
         Left            =   10110
         ScaleHeight     =   345
         ScaleWidth      =   3675
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   30
         Width           =   3675
         Begin Cis3D_v60.CIS3D CIS3D15 
            Height          =   345
            Left            =   2460
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   8421631
            ForeColor       =   0
            Caption         =   "削除"
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
            cAlingnment     =   1
            cPositionX      =   20
            Begin Cis3D_v60.CIS3D iT1_Delsu 
               Height          =   255
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   8421631
               ForeColor       =   0
               Caption         =   "ZZ9"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
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
         Begin Cis3D_v60.CIS3D UC_3D6 
            Height          =   345
            Left            =   0
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   8454143
            ForeColor       =   0
            Caption         =   "追加"
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
            cAlingnment     =   1
            cPositionX      =   20
            Begin Cis3D_v60.CIS3D iT1_Addsu 
               Height          =   255
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   8454143
               ForeColor       =   0
               Caption         =   "ZZ9"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
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
         Begin Cis3D_v60.CIS3D UC_3D7 
            Height          =   345
            Left            =   1230
            Top             =   0
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   609
            BackColor       =   16777088
            ForeColor       =   0
            Caption         =   "修正"
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
            cAlingnment     =   1
            cPositionX      =   20
            Begin Cis3D_v60.CIS3D iT1_Mntsu 
               Height          =   255
               Left            =   540
               Top             =   30
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   450
               BackColor       =   16777088
               ForeColor       =   0
               Caption         =   "ZZ9"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   11.25
                  Charset         =   0
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
      End
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
         Left            =   7560
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
   Begin VB.PictureBox Back_Call 
      Height          =   1230
      Left            =   11565
      ScaleHeight     =   1170
      ScaleWidth      =   3435
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1815
      Visible         =   0   'False
      Width           =   3495
      Begin Cis3D_v60.CIS3D Call_YMD 
         Height          =   315
         Left            =   150
         Top             =   330
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         BackColor       =   16777152
         Caption         =   "9999/99/99 ～ 9999/99/99"
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
         cAlingnment     =   7
         cBoderStyle     =   1
      End
      Begin Cis3D_v60.CIS3D CALL_Su 
         Height          =   315
         Left            =   2145
         Top             =   780
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         BackColor       =   16777152
         Caption         =   "99,999"
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
         cAlingnment     =   5
         cBoderStyle     =   1
         cPositionX      =   -60
      End
      Begin Cis3D_v60.CIS3D CALL_KTorcd 
         Height          =   315
         Left            =   735
         Top             =   780
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         BackColor       =   16777152
         Caption         =   "99,999"
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
         cAlingnment     =   7
         cBoderStyle     =   1
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "契約先"
         Height          =   180
         Left            =   165
         TabIndex        =   33
         Top             =   855
         Width           =   540
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "単 価 対 象 日 付"
         Height          =   225
         Left            =   1080
         TabIndex        =   32
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "数量"
         Height          =   180
         Left            =   1755
         TabIndex        =   31
         Top             =   840
         Width           =   360
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  '不透明
         Height          =   435
         Left            =   1680
         Shape           =   4  '丸みのある長方形
         Top             =   720
         Width           =   1725
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  '不透明
         Height          =   705
         Left            =   60
         Shape           =   4  '丸みのある長方形
         Top             =   30
         Width           =   3345
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   1  '不透明
         Height          =   435
         Left            =   60
         Shape           =   4  '丸みのある長方形
         Top             =   720
         Width           =   1635
      End
   End
   Begin Cis3D_v60.CIS3D H1_ZaiDsp 
      Height          =   300
      Left            =   870
      Top             =   10680
      Visible         =   0   'False
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   529
      BackColor       =   16744576
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
      cBoderColor1    =   16744576
      cBoderColor2    =   16744576
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cFont3D         =   2
      cBoderStyle     =   2
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   330
         Index           =   30
         Left            =   0
         Top             =   -30
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   582
         BackColor       =   4210752
         ForeColor       =   16777215
         Caption         =   "材料/寸法"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   4210752
         cBoderColor2    =   4210752
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cFont3D         =   2
         cAlingnment     =   6
         cBoderStyle     =   2
         cPositionY      =   60
      End
      Begin Cis3D_v60.CIS3D H1_ZaiRyo 
         Height          =   300
         Left            =   960
         Top             =   0
         Width           =   5910
         _ExtentX        =   10425
         _ExtentY        =   529
         BackColor       =   16744576
         ForeColor       =   16777215
         Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXX5/ZZZZ.ZZx0000.00xZZZZ.ZZ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   16744576
         cBoderColor2    =   16744576
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cAlingnment     =   1
         cBoderStyle     =   2
         cPositionX      =   30
      End
   End
End
Attribute VB_Name = "CXM0180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   購入単価マスタメンテナンス
'**       フォームID    :   CXM0180
'**       処理概要      :
'**
'**       作  成  日    :   2009/09/30  By CIS SIMPR-A
'**       変  更  日    :   2019/07/19  単価0円入力可
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    
    Dim EtcMsg(6)       As String
    Dim EtcTan(6)       As Currency

    Dim Hinku_top(3)   As Integer       ' 品番表示位置      ' 2009.09.30 add
    Dim gMaeIndex      As Integer       ' １つ前の選択品番  ' 2009.09.30 add
    Dim cnt            As Integer       ' ワークカウンタ    ' 2009.10.01 add
    Dim wsl            As String        ' ワーク文字列      ' 2009.10.01 add

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    '------------------------------ 2005/06/14  Insert
    If Rv_Call3 = "" Then
    '------------------------------ 2005/06/14  Insert End
    '   #------------------#
    '   # 二 重 起 動 防 止 #
    '   #------------------#
        If Not Dupli_Start Then End
    '   #------------------#
    '   # 初 期 内 容 取 得 #
    '   #------------------#
        If Not IniGet Then End
    '------------------------------ 2005/06/14  Insert
    End If
    '------------------------------ 2005/06/14  Insert End
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
'
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 360
        .PatanMax = 1
        .SelectCol = 21
        .InitGet = False
        .Init
    End With
    Call GridInit
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H0"                               ' 2009.09.30 upd : "H1"-->"H0"
    Call DispChange(ProcHB)
'-----------------------------------------------' 2009.09.30 add start
    mGotColor = H0Op_Tanku(1).BackColor
    mLostColor = H0Op_Tanku(0).BackColor
    Hinku_top(0) = b_HinNM(1).Top
    Hinku_top(1) = b_HinNM(1).Top
    Hinku_top(2) = b_HinNM(2).Top
    Hinku_top(3) = b_HinNM(3).Top
    H0Op_Tanku(0).Value = True
    Call H0Op_Tanku_Click(1)
    H1_ZaiRyo.Caption = ""
   'H1Op_Hinku(0).Value = True
   'Call H1Op_Hinku_Click(0)
'-----------------------------------------------' 2009.09.30 add end
'-------------------------------------------------------------------------------' 2009.09.30 試験用 start
'H1lb_HinBn(0).Caption = "HC-HINBAN-発注品番"
'H1lb_HinBn(1).Caption = "JK-HINBAN-受給品番"
'H1lb_HinBn(2).Caption = "JS-HINBAN-自社品番"
'H1lb_HinBn(3).Caption = "SPH440-0D2.9Ｘ238Ｘ1219 5XXXX3XXXXXXXXX4XXXXXXXXX5"
'-------------------------------------------------------------------------------' 2009.09.30 試験用 end
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
'    SyoriKB = "ADD"                            ' 2009.09.30 del
'    Call SyoriName(SyoriKB)                    ' 2009.09.30 del
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    '------------------------------ 2005/06/14  Insert
    If Rv_Call3 = "" Then
    '------------------------------ 2005/06/14  Insert End
       Me.Show
    '------------------------------ 2005/06/14  Insert
    End If
    '------------------------------ 2005/06/14  Insert End
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    FormAct = Not FormAct
'初期設定
    Call EtcTanka_Naiyo
    
    '------------------------------ 2005/06/14  Insert
    If Rv_Call3 = "X" Then
       iH1_Hinbn = RV_Hinbn
       iH1_Torcd = RV_Torcd
       Call iH1_Torcd_LostFocus
       
'       iH1_KTorcd = RV_KTorcd
'       Call iH1_KTorcd_LostFocus
       
       Call_YMD = RV_SYmd
       CALL_Su = RV_EYmd
       CALL_KTorcd = RV_KTorcd
       
       Back_Call.Visible = True
       
       Call ReturnPress
    End If
    '------------------------------ 2005/06/14  Insert End

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
           Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click    '【新規】
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
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
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Cancel = 1
    Call PB_END_Click
End Sub

'---------------------------------------------------------------' 2009.09.30 add start
'+---------------------------+
'+      H1Op_Hinku           +
'+---------------------------+
Private Sub H1Op_Hinku_Click(INDEX As Integer)
    H1Op_Hinku(0).BackColor = mLostColor
    H1Op_Hinku(1).BackColor = mLostColor
    H1Op_Hinku(2).BackColor = mLostColor
    H1Op_Hinku(3).BackColor = mLostColor
    H1Op_Hinku(INDEX).BackColor = mGotColor
    
    For cnt = 0 To 3
        b_HinNM(cnt).Visible = True
        H1lb_HinBn(cnt).Visible = True
    Next
    b_HinNM(INDEX).Visible = False
    H1lb_HinBn(INDEX).Visible = False

    gStr = RTrim(iH1_Hinbn.Text)
    iH1_Hinbn.Text = RTrim(H1lb_HinBn(INDEX))
    H1lb_HinBn(gMaeIndex) = gStr
   'Call HinbnSet(0)

    If H0Op_Tanku(1).Value Then
       Call HinbnToLabel
       Select Case INDEX
          Case 0:
            iH1_Hinbn.Text = H1lb_HinBn(0).Caption
            b_HinNM(1).Top = Hinku_top(1)
            H1lb_HinBn(1).Top = Hinku_top(1)
            b_HinNM(2).Top = Hinku_top(2)
            H1lb_HinBn(2).Top = Hinku_top(2)
            b_HinNM(3).Top = Hinku_top(3)
            H1lb_HinBn(3).Top = Hinku_top(3)
         
         Case 1:
            iH1_Hinbn.Text = H1lb_HinBn(1).Caption
            b_HinNM(0).Top = Hinku_top(1)
            H1lb_HinBn(0).Top = Hinku_top(1)
            b_HinNM(2).Top = Hinku_top(2)
            H1lb_HinBn(2).Top = Hinku_top(2)
            b_HinNM(3).Top = Hinku_top(3)
            H1lb_HinBn(3).Top = Hinku_top(3)
          
          Case 2:
            iH1_Hinbn.Text = H1lb_HinBn(2).Caption
            b_HinNM(0).Top = Hinku_top(1)
            H1lb_HinBn(0).Top = Hinku_top(1)
            b_HinNM(1).Top = Hinku_top(2)
            H1lb_HinBn(1).Top = Hinku_top(2)
            b_HinNM(3).Top = Hinku_top(3)
            H1lb_HinBn(3).Top = Hinku_top(3)

          Case 3:
            iH1_Hinbn.Text = H1lb_HinBn(3).Caption
            b_HinNM(0).Top = Hinku_top(1)
            H1lb_HinBn(0).Top = Hinku_top(1)
            b_HinNM(1).Top = Hinku_top(2)
            H1lb_HinBn(1).Top = Hinku_top(2)
            b_HinNM(2).Top = Hinku_top(3)
            H1lb_HinBn(2).Top = Hinku_top(3)
            Call H1_ZaiBox_Set
          Case Else:
       End Select
    End If
    gMaeIndex = INDEX
    
    If H1Op_Hinku(3).Value = True Then
      'iH1_Hinbn.Visible = False
       H1_ZaiBox.Top = 360 '360 '750 : 750:試験用 / 360:正しい位置
       H1_ZaiBox.Visible = True
    Else
       H1_ZaiBox.Visible = False
       iH1_Hinbn.Visible = True
    End If
    
    If iH1_Hinbn.Tag = "1" Then
       If H1Op_Hinku(3).Value = True Then
          Call iH1_Zaist.SetFocus
       Else
          Call iH1_Hinbn.SetFocus
       End If
    ElseIf iH1_Torcd.Tag = "1" Then
       Call iH1_Torcd.SetFocus
    End If
End Sub
'+-------------------------+
'+  H0Op_Tanku             +
'+-------------------------+
Private Sub H0Op_Tanku_Click(INDEX As Integer)
    Dim cnt As Integer
    
    H0Op_Tanku(0).BackColor = mLostColor
    H0Op_Tanku(1).BackColor = mLostColor
    H0Op_Tanku(INDEX).BackColor = mGotColor
    H0Op_Tanku(INDEX).Value = True
    
    H1Op_Hinku(0).Visible = H0Op_Tanku(1).Value
    b_HinNM(0).Visible = False
    H1lb_HinBn(0).Visible = False
    
    For cnt = 1 To 3
        H1Op_Hinku(cnt).Visible = H0Op_Tanku(1).Value
        b_HinNM(cnt).Visible = H0Op_Tanku(1).Value
        H1lb_HinBn(cnt).Visible = H0Op_Tanku(1).Value
        b_HinNM(cnt).Top = Hinku_top(cnt)
        H1lb_HinBn(cnt).Top = Hinku_top(cnt)
    Next

    If INDEX = 0 Then               ' 単価内訳は「部品」のときのみ有効
      'Panel_Uchi.Visible = True: Panel_SonoTa.Visible = True
       iB1_BTanka.Visible = True: iB1_KTanka.Visible = True: iB1_STanka.Visible = True
       For cnt = 0 To 5
           iB1_EtcMsg(cnt).Visible = True: iB1_EtcTan(cnt).Visible = True
       Next
    Else
      'Panel_Uchi.Visible = False: Panel_SonoTa.Visible = False
       iB1_BTanka.Visible = False: iB1_KTanka.Visible = False: iB1_STanka.Visible = False
       For cnt = 0 To 5
           iB1_EtcMsg(cnt).Visible = False: iB1_EtcTan(cnt).Visible = False
       Next
    End If
End Sub
'---------------------------------------------------------------' 2009.09.30 add end

'---------------------------------------------------------------' 2009.10.06 add start
'+---------------------+
'+    材質入力エリア   +
'+---------------------+
Private Sub iH1_Zaist_GotFocus()
    DH1_Zaist.BackColor = &HC0FFFF
    DH1_Zaist.cBoderColor2 = &HC0FFFF
    PB_Look.Tag = "iH1_Hinbn"   ' ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Zaist_LostFocus()
    Call HinbnToLabel
    DH1_Zaist.BackColor = &HC0E0FF
    DH1_Zaist.cBoderColor2 = &HC0E0FF
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    
    Call Zai_Chk

End Sub
Private Sub iH1_Itatu_GotFocus()
    DH1_Itatu.BackColor = &HC0FFFF
    DH1_Itatu.cBoderColor2 = &HC0FFFF
    PB_Look.Tag = "iH1_Hinbn"   ' ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Itatu_LostFocus()
    Call HinbnToLabel
    DH1_Itatu.BackColor = &HC0E0FF
    DH1_Itatu.cBoderColor2 = &HC0E0FF
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    
    Call Zai_Chk

End Sub
Private Sub iH1_Width_GotFocus()
    DH1_Width.BackColor = &HC0FFFF
    DH1_Width.cBoderColor2 = &HC0FFFF
    PB_Look.Tag = "iH1_Hinbn"   ' ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Width_LostFocus()
    Call HinbnToLabel
    DH1_Width.BackColor = &HC0E0FF
    DH1_Width.cBoderColor2 = &HC0E0FF
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    Call Zai_Chk

End Sub
Private Sub iH1_Long_GotFocus()
    DH1_Long.BackColor = &HC0FFFF
    DH1_Long.cBoderColor2 = &HC0FFFF
    PB_Look.Tag = "iH1_Hinbn"   ' ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Long_LostFocus()
    Call HinbnToLabel
    DH1_Long.BackColor = &HC0E0FF
    DH1_Long.cBoderColor2 = &HC0E0FF
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    
    Call Zai_Chk

End Sub
Private Sub H1_ZaiBox_Set()
    Dim Atai As String
    iH1_Zaist.Text = RTrim(Mid(H1_ZaiRyo.Caption, 1, 25))
    Atai = RTrim(Mid(H1_ZaiRyo.Caption, 27, 7))
    If Atai <> "" Then iH1_Itatu.Text = Format(CisFun.Val2(Atai), "#,##0.00")
    Atai = RTrim(Mid(H1_ZaiRyo.Caption, 35, 7))
    If Atai <> "" Then iH1_Width.Text = Format(CisFun.Val2(Atai), "#,##0.00")
    Atai = RTrim(Mid(H1_ZaiRyo.Caption, 43, 7))
    If Atai <> "" Then iH1_Long.Text = Format(CisFun.Val2(Atai), "#,##0.00")
End Sub
'---------------------------------------------------------------' 2009.10.06 add end
Private Sub Zai_Chk()
    
    If H1Op_Hinku(0) Then
        gSL_Select = ""
        gSL_Select = gSL_Select & " SELECT 材料管理番号,COUNT(*) 件数 FROM 材料取引先マスタ "
        gSL_Select = gSL_Select & " WHERE 発注品番 = '" & RTrim(iH1_Hinbn) & "'"
        If RTrim(iH1_Torcd) <> "" Then
           gSL_Select = gSL_Select & " AND   取引先 = '" & RTrim(iH1_Torcd) & "'"
        End If
        gSL_Select = gSL_Select & " GROUP BY 材料管理番号 "
        If ZTMRead(gSL_Select, 1) Then
            If ZTM.件数 = 1 Then
               H1lb_ZaiKan = ZRM.材料管理番号
            Else
               H1lb_ZaiKan = ""
            End If
        End If
    End If
    If H1Op_Hinku(1) Then
        gSL_Select = ""
        gSL_Select = gSL_Select & " SELECT 材料管理番号,COUNT(*) 件数 FROM 材料取引先マスタ "
        gSL_Select = gSL_Select & " WHERE 受給品番 = '" & RTrim(iH1_Hinbn) & "'"
        If RTrim(iH1_Torcd) <> "" Then
           gSL_Select = gSL_Select & " AND   取引先 = '" & RTrim(iH1_Torcd) & "'"
        End If
        gSL_Select = gSL_Select & " GROUP BY 材料管理番号 "
        If ZTMRead(gSL_Select, 1) Then
            If ZTM.件数 = 1 Then
               H1lb_ZaiKan = ZRM.材料管理番号
            Else
               H1lb_ZaiKan = ""
            End If
        End If
    End If
    If H1Op_Hinku(2) Then
        gSL_Select = ""
        gSL_Select = gSL_Select & " SELECT 材料管理番号,COUNT(*) 件数 FROM 材料取引先マスタ "
        gSL_Select = gSL_Select & " WHERE 自社品番 = '" & RTrim(iH1_Hinbn) & "'"
        If RTrim(iH1_Torcd) <> "" Then
           gSL_Select = gSL_Select & " AND   取引先 = '" & RTrim(iH1_Torcd) & "'"
        End If
        gSL_Select = gSL_Select & " GROUP BY 材料管理番号 "
        If ZTMRead(gSL_Select, 1) Then
            If ZTM.件数 = 1 Then
               H1lb_ZaiKan = ZRM.材料管理番号
            Else
               H1lb_ZaiKan = ""
            End If
        End If
    End If
    
    
    If H1Op_Hinku(3) Then
        gSL_Select = ""
        gSL_Select = gSL_Select & " SELECT 材料管理番号 FROM 材料マスタ "
        gSL_Select = gSL_Select & " WHERE 材質 = '" & RTrim(iH1_Zaist) & "'"
        gSL_Select = gSL_Select & " AND   板厚 = " & CisFun.Val2(iH1_Itatu)
        gSL_Select = gSL_Select & " AND   幅 = " & CisFun.Val2(iH1_Width)
        gSL_Select = gSL_Select & " AND   長さ = " & CisFun.Val2(iH1_Long)
        If ZRMRead(gSL_Select, 1) Then
            H1lb_ZaiKan = ZRM.材料管理番号
        End If
    End If

End Sub

'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    '------------------------------ 2005/06/14  Insert
    If Rv_Call3 = "" Then
    '------------------------------ 2005/06/14  Insert End
    '+---------------------+
    '+   ﾃﾞｰﾀﾍﾞｰｽ切断
    '+---------------------+
        Call CisDB.DBDISConnect
    '+---------------------+
    '+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
    '+---------------------+
        Call CisFrm.UnLoadDisp
        End
    '------------------------------ 2005/06/14  Insert
    Else
        Me.Hide
    End If
    '------------------------------ 2005/06/14  Insert End
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
Private Sub PB_CAN_Click()
   Call CisFrm.MousePT(11)
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKB)
      If H1Op_Hinku(3).Value = True Then
         iH1_Zaist.SetFocus
      Else
         iH1_Hinbn.SetFocus
      End If
      GoTo PB_CAN_Ed
   End If

'-----------------------------------------------' 2009.09.30 add start
   If ProcHB = "H1" Then
      ProcHB = "H0"
      Call DispChange(ProcHB)
      Call HeadBodyClear("H1")
      Call HeadBodyClear("B")
''''  Call SyoriName(SyoriKB)
''''  iH1_Hinbn.SetFocus
      H1_ZaiBox.Visible = False
      H1_ZaiRyo.Caption = ""
      iH1_Hinbn.Text = ""
      iH1_Hinbn.Visible = True
     'h1op_hinku(0).Value=true
      If H0Op_Tanku(0).Value Then
         H0Op_Tanku(0).SetFocus
      Else
         H0Op_Tanku(1).SetFocus
      End If
      GoTo PB_CAN_Ed
   End If
'-----------------------------------------------' 2009.09.30 add end

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
    If H1Op_Hinku(3).Value = True Then
       iH1_Zaist.SetFocus
    Else
       iH1_Hinbn.SetFocus
    End If
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKB <> "MNT" Then
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
    End If
    If H1Op_Hinku(3).Value = True Then
       iH1_Zaist.SetFocus
    Else
       iH1_Hinbn.SetFocus
    End If
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
    If H1Op_Hinku(3).Value = True Then
       iH1_Zaist.SetFocus
    Else
       iH1_Hinbn.SetFocus
    End If
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) + ' 部品側
'+----------------------------+
Private Sub PB_Back_Click()
  If H0Op_Tanku(0).Value Then   ' 2009.10.01 add
    Call CisFrm.MousePT(11)
    gSL_Select = " SELECT *"
    gSL_Select = gSL_Select & " FROM 購入単価マスタ "
    gSL_Select = gSL_Select & " WHERE 品番 + 取引先 + 契約先 < '" & RTrim(iH1_Hinbn) & iH1_Torcd & iH1_KTorcd & "'"
    gSL_Select = gSL_Select & " ORDER BY 品番 + 取引先 + 契約先 DESC "
    If Not KTMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     先頭レコードです            "
            .MB_Title = "【後退】"
            .MB_Button = OK
            .MBOX
        End With
        Call CisFrm.MousePT(1)
        Exit Sub
    End If
    
    iH1_Hinbn = KTM.品番
    iH1_Torcd = KTM.取引先
    Call iH1_Torcd_LostFocus
    iH1_KTorcd = KTM.契約先
    Call iH1_KTorcd_LostFocus
    
    gSL_Select = "Select * FROM 購入単価マスタ "
    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   And 契約先 = '" & RTrim(iH1_KTorcd) & "'"
    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Hinbn.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
  Else                      ' 2009.10.01 add
    Call PB_Back_Click_Z    ' 2009.10.01 add
  End If                    ' 2009.10.01 add
End Sub
'-----------------------------------------------------------------------------------' 2009.10.01 add start
'+----------------------------+
'+    後  退  キ ー(PageUp  ) + ' 材料側
'+----------------------------+
Private Sub PB_Back_Click_Z()
    Call CisFrm.MousePT(11)
    
    '================= 入力情報より後退時のキー作成 ===================
    gStr = ""
    If H1Op_Hinku(3).Value = True Then
'       gStr = gStr & RTrim(iH1_Hinbn) & Space(50 - CisFun.Len2(RTrim(iH1_Hinbn)))
       gStr = gStr & RTrim(iH1_Hinbn)
    Else
       gStr = gStr & RTrim(iH1_Hinbn) & Space(25 - CisFun.Len2(RTrim(iH1_Hinbn)))
    End If
    gStr = gStr & RTrim(iH1_Torcd) & Space(7 - CisFun.Len2(RTrim(iH1_Torcd)))
    gStr = gStr & RTrim(iH1_KTorcd) & Space(7 - CisFun.Len2(RTrim(iH1_KTorcd)))

    '============================================
    '============= 部品区分：材料 ===============
    '============================================
    gSL_Select = ""
    gSL_Select = gSL_Select & " SELECT KM.*, ISNULL(HT.発注品番,'') 発注品番, ISNULL(HT.受給品番,'') 受給品番, ISNULL(HT.自社品番,'') 自社品番," ' ISNULL(HT.背番号, '') 背番号, を外して自社品番を追加
    gSL_Select = gSL_Select & "        ZM.材質, ZM.板厚, ZM.幅, ZM.長さ, 表示寸法"
    gSL_Select = gSL_Select & "   FROM 材料購入単価マスタ KM"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料マスタ ZM"
    gSL_Select = gSL_Select & "             ON ZM.材料管理番号 = KM.材料管理番号"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料取引先マスタ HT"
    gSL_Select = gSL_Select & "             ON HT.材料管理番号 = KM.材料管理番号"
    gSL_Select = gSL_Select & "            AND HT.取引先 = KM.取引先"
'    gSL_Select = gSL_Select & "            AND HT.契約先 = KM.契約先"
'   gSL_Select = gSL_Select & "            AND HT.品目 = 1"
    
    If H1Op_Hinku(0).Value Then          ' 発注品番
'       gSL_Select = gSL_Select & " WHERE KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
'       gSL_Select = gSL_Select & "       KM.契約先 + SPACE(7 - LEN(KM.契約先)) + "
'       gSL_Select = gSL_Select & "       HT.発注品番 + SPACE(25 - LEN(HT.発注品番)) < '" & gStr & "'"
'       gSL_Select = gSL_Select & " ORDER BY KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
'       gSL_Select = gSL_Select & "          KM.契約先 + SPACE(7 - LEN(KM.契約先)) + "
'       gSL_Select = gSL_Select & "          HT.発注品番 + SPACE(25 - LEN(HT.発注品番))"
       gSL_Select = gSL_Select & " WHERE HT.発注品番 + SPACE(25 - LEN(HT.発注品番)) + "
       gSL_Select = gSL_Select & "       KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
       gSL_Select = gSL_Select & "       KM.契約先 + SPACE(7 - LEN(KM.契約先))  < '" & gStr & "'"
       gSL_Select = gSL_Select & " ORDER BY HT.発注品番 + SPACE(25 - LEN(HT.発注品番)) + "
       gSL_Select = gSL_Select & "          KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
       gSL_Select = gSL_Select & "          KM.契約先 + SPACE(7 - LEN(KM.契約先))"
'
    ElseIf H1Op_Hinku(1).Value Then      ' 受給品番
'       gSL_Select = gSL_Select & " WHERE KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
'       gSL_Select = gSL_Select & "       KM.契約先 + SPACE(7 - LEN(KM.契約先)) + "
'       gSL_Select = gSL_Select & "       HT.受給品番 + SPACE(25 - LEN(HT.受給品番)) < '" & gStr & "'"
'       gSL_Select = gSL_Select & " ORDER BY KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
'       gSL_Select = gSL_Select & "          KM.契約先 + SPACE(7 - LEN(KM.契約先)) + "
'       gSL_Select = gSL_Select & "          HT.受給品番 + SPACE(25 - LEN(HT.受給品番))"
       gSL_Select = gSL_Select & " WHERE HT.受給品番 + SPACE(25 - LEN(HT.受給品番)) + "
       gSL_Select = gSL_Select & "       KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
       gSL_Select = gSL_Select & "       KM.契約先 + SPACE(7 - LEN(KM.契約先)) < '" & gStr & "'"
       gSL_Select = gSL_Select & " ORDER BY HT.受給品番 + SPACE(25 - LEN(HT.受給品番)) + "
       gSL_Select = gSL_Select & "          KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
       gSL_Select = gSL_Select & "          KM.契約先 + SPACE(7 - LEN(KM.契約先)) "
'
    ElseIf H1Op_Hinku(2).Value Then      ' 自社品番
'       gSL_Select = gSL_Select & " WHERE KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
'       gSL_Select = gSL_Select & "       KM.契約先 + SPACE(7 - LEN(KM.契約先)) + "
'       gSL_Select = gSL_Select & "       HT.自社品番 + SPACE(25 - LEN(HT.自社品番)) < '" & gStr & "'"
'       gSL_Select = gSL_Select & " ORDER BY KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
'       gSL_Select = gSL_Select & "          KM.契約先 + SPACE(7 - LEN(KM.契約先)) + "
'       gSL_Select = gSL_Select & "          HT.自社品番 + SPACE(25 - LEN(HT.自社品番))"
       gSL_Select = gSL_Select & " WHERE HT.自社品番 + SPACE(25 - LEN(HT.自社品番)) + "
       gSL_Select = gSL_Select & "       KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
       gSL_Select = gSL_Select & "       KM.契約先 + SPACE(7 - LEN(KM.契約先)) < '" & gStr & "'"
       gSL_Select = gSL_Select & " ORDER BY HT.自社品番 + SPACE(25 - LEN(HT.自社品番)) + "
       gSL_Select = gSL_Select & "          KM.取引先 + SPACE(7 - LEN(KM.取引先))  + "
       gSL_Select = gSL_Select & "          KM.契約先 + SPACE(7 - LEN(KM.契約先)) "
'
    Else                                 ' 材質/寸法
       wsl = " RTrim(ZM.材質) + ' ' + cast(ZM.板厚 as varchar) + 'X' + cast(ZM.幅   as varchar) + 'X' + cast(ZM.長さ as varchar) "
'       gSL_Select = gSL_Select & " WHERE KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
'       gSL_Select = gSL_Select & "       KM.契約先 + SPACE(7 - LEN(KM.契約先)) + "
'       gSL_Select = gSL_Select & wsl & " < '" & gStr & "'"
'       gSL_Select = gSL_Select & " ORDER BY KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
'       gSL_Select = gSL_Select & "          KM.契約先 + SPACE(7 - LEN(KM.契約先)) + " & wsl
       gSL_Select = gSL_Select & " WHERE " & wsl & " + "
       gSL_Select = gSL_Select & " KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
       gSL_Select = gSL_Select & " KM.契約先 + SPACE(7 - LEN(KM.契約先))"
       gSL_Select = gSL_Select & " < '" & gStr & "'"
       gSL_Select = gSL_Select & " ORDER BY " & wsl & " + "
       gSL_Select = gSL_Select & " KM.取引先 + Space(7 - Len(KM.取引先)) + "
       gSL_Select = gSL_Select & " KM.契約先 + SPACE(7 - LEN(KM.契約先))"
    End If
    gSL_Select = gSL_Select & "       DESC"
''''Call Test_LogOutput(gSL_Select, 1)       ' SQL確認用
    
    If Not ZKTRead(gSL_Select, 1) Then
       With CisFun
           .MB_Lines = 4
           .MB_MSG(2) = "     先頭レコードです            "
           .MB_Title = "【後退】"
           .MB_Button = OK
           .MBOX
       End With
       Call CisFrm.MousePT(1)
       Exit Sub
    End If
    
    H1lb_HinBn(0).Caption = ZKT.発注品番
    H1lb_HinBn(1).Caption = ZKT.受給品番
    H1lb_HinBn(2).Caption = ZKT.自社品番
    H1lb_HinBn(3).Caption = RTrim(ZKT.材質) & " " & _
                            Format(ZKT.板厚, "#,##0.00") & "X" & _
                            Format(ZKT.幅, "#,##0.00") & "X" & _
                            Format(ZKT.長さ, "#,##0.00")
    H1_ZaiRyo.Caption = RTrim(ZKT.材質) & Space(25 - Len(RTrim(ZKT.材質))) & "/" & _
                        Format(ZKT.板厚, "0000.00") & "x" & _
                        Format(ZKT.幅, "0000.00") & "x" & _
                        Format(ZKT.長さ, "0000.00")
    iH1_Hinbn.Text = ""
    For cnt = 0 To 3
        If H1Op_Hinku(cnt).Value = True Then
           iH1_Hinbn.Text = H1lb_HinBn(cnt).Caption
           cnt = 4
       End If
    Next
    Call H1_ZaiBox_Set
    iH1_Torcd = ZKT.取引先
    iH1_KTorcd = ZKT.契約先
    H1lb_ZaiKan = RTrim(ZKT.材料管理番号)
    
    Call iH1_Torcd_LostFocus
    Call iH1_KTorcd_LostFocus
   'Call HinbnSet
    
    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Torcd.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'-----------------------------------------------------------------------------------' 2009.10.01 add end
'+----------------------------+
'+    前  進  キ ー(PageDown) + ' 部品側
'+----------------------------+
Private Sub PB_Next_Click()
  If H0Op_Tanku(0).Value Then   ' 2009.10.01 add
    Call CisFrm.MousePT(11)
    gSL_Select = " SELECT *"
    gSL_Select = gSL_Select & " FROM 購入単価マスタ "
    gSL_Select = gSL_Select & " WHERE 品番 + 取引先 + 契約先 > '" & RTrim(iH1_Hinbn) & iH1_Torcd & iH1_KTorcd & "'"
    gSL_Select = gSL_Select & " ORDER BY 品番 + 取引先 + 契約先 "
    If Not KTMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     最終レコードです            "
            .MB_Title = "【前進】"
            .MB_Button = OK
            .MBOX
        End With
        Call CisFrm.MousePT(1)
        Exit Sub
    End If
    
    iH1_Hinbn = KTM.品番
    iH1_Torcd = KTM.取引先
    Call iH1_Torcd_LostFocus
    iH1_KTorcd = KTM.契約先
    Call iH1_KTorcd_LostFocus
    
    gSL_Select = "Select * FROM 購入単価マスタ "
    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   And 契約先 = '" & RTrim(iH1_KTorcd) & "'"
    Call DBInput("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Hinbn.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
  Else                      ' 2009.10.01 add
    Call PB_Next_Click_Z    ' 2009.10.01 add
  End If                    ' 2009.10.01 add
End Sub
'-----------------------------------------------------------------------------------' 2009.10.01 add start
'+----------------------------+
'+    前  進  キ ー(PageDown) + ' 材料側
'+----------------------------+
Private Sub PB_Next_Click_Z()
    Call CisFrm.MousePT(11)
    
    '================= 入力情報より後退時のキー作成 ===================
    gStr = ""
    If H1Op_Hinku(3).Value = True Then
'       gStr = gStr & RTrim(iH1_Hinbn) & Space(50 - CisFun.Len2(RTrim(iH1_Hinbn)))
       gStr = gStr & RTrim(iH1_Hinbn)
    Else
       gStr = gStr & RTrim(iH1_Hinbn) & Space(25 - CisFun.Len2(RTrim(iH1_Hinbn)))
    End If
    gStr = gStr & RTrim(iH1_Torcd) & Space(7 - CisFun.Len2(RTrim(iH1_Torcd)))
    gStr = gStr & RTrim(iH1_KTorcd) & Space(7 - CisFun.Len2(RTrim(iH1_KTorcd)))
    
    '============================================
    '============= 部品区分：材料 ===============
    '============================================
    gSL_Select = ""
    gSL_Select = gSL_Select & " SELECT KM.*, ISNULL(HT.発注品番,'') 発注品番, ISNULL(HT.受給品番,'') 受給品番, ISNULL(HT.自社品番,'') 自社品番," ' ISNULL(HT.背番号, '') 背番号, を外して自社品番を追加
    gSL_Select = gSL_Select & "        ZM.材質, ZM.板厚, ZM.幅, ZM.長さ"
    gSL_Select = gSL_Select & "   FROM 材料購入単価マスタ KM"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料マスタ ZM"
    gSL_Select = gSL_Select & "             ON ZM.材料管理番号 = KM.材料管理番号"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料取引先マスタ HT"
    gSL_Select = gSL_Select & "             ON HT.材料管理番号 = KM.材料管理番号"
    gSL_Select = gSL_Select & "            AND HT.取引先 = KM.取引先"
'    gSL_Select = gSL_Select & "            AND HT.契約先 = KM.契約先"
'   gSL_Select = gSL_Select & "            AND HT.品目 = 1"
       
    If H1Op_Hinku(0).Value Then          ' 発注品番
'       gSL_Select = gSL_Select & " WHERE KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
'       gSL_Select = gSL_Select & "       KM.契約先 + SPACE(7 - LEN(KM.契約先)) + "
'       gSL_Select = gSL_Select & "       HT.発注品番 + SPACE(25 - LEN(HT.発注品番)) > '" & gStr & "'"
'       gSL_Select = gSL_Select & " ORDER BY KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
'       gSL_Select = gSL_Select & "          KM.契約先 + SPACE(7 - LEN(KM.契約先)) + "
'       gSL_Select = gSL_Select & "          HT.発注品番 + SPACE(25 - LEN(HT.発注品番))"
       gSL_Select = gSL_Select & " WHERE HT.発注品番 + SPACE(25 - LEN(HT.発注品番)) + "
       gSL_Select = gSL_Select & "       KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
       gSL_Select = gSL_Select & "       KM.契約先 + SPACE(7 - LEN(KM.契約先)) > '" & gStr & "'"
       gSL_Select = gSL_Select & " ORDER BY HT.発注品番 + SPACE(25 - LEN(HT.発注品番)) + "
       gSL_Select = gSL_Select & "          KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
       gSL_Select = gSL_Select & "          KM.契約先 + SPACE(7 - LEN(KM.契約先))"
'
    ElseIf H1Op_Hinku(1).Value Then      ' 受給品番
'       gSL_Select = gSL_Select & " WHERE KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
'       gSL_Select = gSL_Select & "       KM.契約先 + SPACE(7 - LEN(KM.契約先)) + "
'       gSL_Select = gSL_Select & "       HT.受給品番 + SPACE(25 - LEN(HT.受給品番)) > '" & gStr & "'"
'       gSL_Select = gSL_Select & " ORDER BY KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
'       gSL_Select = gSL_Select & "          KM.契約先 + SPACE(7 - LEN(KM.契約先)) + "
'       gSL_Select = gSL_Select & "          HT.受給品番 + SPACE(25 - LEN(HT.受給品番))"
       gSL_Select = gSL_Select & " WHERE HT.受給品番 + SPACE(25 - LEN(HT.受給品番)) + "
       gSL_Select = gSL_Select & "       KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
       gSL_Select = gSL_Select & "       KM.契約先 + SPACE(7 - LEN(KM.契約先)) > '" & gStr & "'"
       gSL_Select = gSL_Select & " ORDER BY HT.受給品番 + SPACE(25 - LEN(HT.受給品番)) + "
       gSL_Select = gSL_Select & "          KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
       gSL_Select = gSL_Select & "          KM.契約先 + SPACE(7 - LEN(KM.契約先)) "
'
    ElseIf H1Op_Hinku(2).Value Then      ' 自社品番
'       gSL_Select = gSL_Select & " WHERE KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
'       gSL_Select = gSL_Select & "       KM.契約先 + SPACE(7 - LEN(KM.契約先)) + "
'       gSL_Select = gSL_Select & "       HT.自社品番 + SPACE(25 - LEN(HT.自社品番)) > '" & gStr & "'"
'       gSL_Select = gSL_Select & " ORDER BY KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
'       gSL_Select = gSL_Select & "          KM.契約先 + SPACE(7 - LEN(KM.契約先)) + "
'       gSL_Select = gSL_Select & "          HT.自社品番 + SPACE(25 - LEN(HT.自社品番))"
       gSL_Select = gSL_Select & " WHERE HT.自社品番 + SPACE(25 - LEN(HT.自社品番))  + "
       gSL_Select = gSL_Select & "       KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
       gSL_Select = gSL_Select & "       KM.契約先 + SPACE(7 - LEN(KM.契約先))  > '" & gStr & "'"
       gSL_Select = gSL_Select & " ORDER BY HT.自社品番 + SPACE(25 - LEN(HT.自社品番)) + "
       gSL_Select = gSL_Select & "          KM.取引先 + SPACE(7 - LEN(KM.取引先))  + "
       gSL_Select = gSL_Select & "          KM.契約先 + SPACE(7 - LEN(KM.契約先)) "
'
    Else                                 ' 材質/寸法
       wsl = " RTrim(ZM.材質) + ' ' + cast(ZM.板厚 as varchar) + 'X' + cast(ZM.幅   as varchar) + 'X' + cast(ZM.長さ as varchar) "
'       gSL_Select = gSL_Select & " WHERE KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
'       gSL_Select = gSL_Select & "       KM.契約先 + SPACE(7 - LEN(KM.契約先)) + "
'       gSL_Select = gSL_Select & wsl & " > '" & gStr & "'"
'       gSL_Select = gSL_Select & " ORDER BY KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
'       gSL_Select = gSL_Select & "          KM.契約先 + SPACE(7 - LEN(KM.契約先)) + " & wsl
       gSL_Select = gSL_Select & " WHERE " & wsl & " + "
       gSL_Select = gSL_Select & " KM.取引先 + SPACE(7 - LEN(KM.取引先)) + "
       gSL_Select = gSL_Select & " KM.契約先 + SPACE(7 - LEN(KM.契約先))"
       gSL_Select = gSL_Select & " > '" & gStr & "'"
       gSL_Select = gSL_Select & " ORDER BY " & wsl & " + "
       gSL_Select = gSL_Select & " KM.取引先 + Space(7 - Len(KM.取引先)) + "
       gSL_Select = gSL_Select & " KM.契約先 + SPACE(7 - LEN(KM.契約先))"
    End If
''''Call Test_LogOutput(gSL_Select, 1)       ' SQL確認用
       
    If Not ZKTRead(gSL_Select, 1) Then
       With CisFun
           .MB_Lines = 4
           .MB_MSG(2) = "     最終レコードです            "
           .MB_Title = "【前進】"
           .MB_Button = OK
           .MBOX
       End With
       Call CisFrm.MousePT(1)
       Exit Sub
    End If
    
    H1lb_HinBn(0).Caption = ZKT.発注品番
    H1lb_HinBn(1).Caption = ZKT.受給品番
    H1lb_HinBn(2).Caption = ZKT.自社品番
    H1lb_HinBn(3).Caption = RTrim(ZKT.材質) & " " & _
                            Format(ZKT.板厚, "#,##0.00") & "X" & _
                            Format(ZKT.幅, "#,##0.00") & "X" & _
                            Format(ZKT.長さ, "#,##0.00")
    H1_ZaiRyo.Caption = RTrim(ZKT.材質) & Space(25 - Len(RTrim(ZKT.材質))) & "/" & _
                        Format(ZKT.板厚, "0000.00") & "x" & _
                        Format(ZKT.幅, "0000.00") & "x" & _
                        Format(ZKT.長さ, "0000.00")
    iH1_Hinbn.Text = ""
    For cnt = 0 To 3
        If H1Op_Hinku(cnt).Value = True Then
           iH1_Hinbn.Text = H1lb_HinBn(cnt).Caption
           cnt = 4
       End If
    Next
    Call H1_ZaiBox_Set
    iH1_Torcd = ZKT.取引先
    iH1_KTorcd = ZKT.契約先
    H1lb_ZaiKan = RTrim(ZKT.材料管理番号)
    
    Call iH1_Torcd_LostFocus
    Call iH1_KTorcd_LostFocus
   'Call HinbnSet
    
    Call DBInput("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Torcd.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'-----------------------------------------------------------------------------------' 2009.10.01 add end

'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If H0Op_Tanku(0).Value = True Then                                      ' 2009.10.02 add
        If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
        If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd(0): Exit Sub
        If PB_Look.Tag = "iH1_KTorcd" Then Call Look_Torcd(1): Exit Sub
    Else                                                                    ' 2009.10.02 add start
        If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Zairyo: Exit Sub
        If PB_Look.Tag = "iH1_Torcd" Then Call Look_ZaiTor(0):  Exit Sub
        If PB_Look.Tag = "iH1_KTorcd" Then Call Look_ZaiTor(1): Exit Sub
    End If                                                                  ' 2009.10.02 add end
    
    If PB_Look.Tag = "iB1_TSYmd" Then iB1_TSYmd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iB1_TEYmd" Then iB1_TEYmd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iB1_EtcMsg" Then Call Look_Naiyo: Exit Sub
End Sub
'+----------------------------+
'+       検　　　　 索         +
'+----------------------------+
Private Sub Look_Hinbn()

    RV_Left = 0
    RV_Top = 0
    RVI_Hinmoku = 1
    RVI_SK = 9
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing
    
    If RV_Rtn Then
       iH1_Hinbn = RV_Hinbn
       iH1_Torcd = RV_TorcdK
       Call iH1_Torcd_LostFocus
    End If
End Sub
'+----------------------------+
'+       手配先検索            +
'+----------------------------+
Private Sub Look_Torcd(TKbn As Byte)
    
    RV_Left = 0
    RV_Top = 0
    If TKbn = 0 Then
        RV_TorKb = 8
    Else
        RV_TorKb = 5
    End If
    RV_Call = ""
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        If TKbn = 0 Then
            iH1_Torcd = RV_TorcdK
            H1lb_Tornm = RV_TorRName
        Else
            iH1_KTorcd = RV_TorcdK
            H1lb_KTornm = RV_TorRName
        End If
    End If
End Sub
'-----------------------------------------------' 2009.10.02 add start
'+----------------------------+
'+       材料検索             +
'+----------------------------+
Private Sub Look_Zairyo()
    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RV_Hinmoku = 1
    RV_Zaist = ""

    CKK0070.Show vbModal
    Unload CKK0070
    Set CKK0070 = Nothing

    If RV_Rtn Then
        H1lb_HinBn(0) = RV_HHinbn
        H1lb_HinBn(1) = RV_JHinbn
        H1lb_HinBn(3) = RV_Zaist & " " & _
                        Format(RV_Itatu, "#,##0.00") & "X" & _
                        Format(RV_Width, "#,##0.00") & "X" & _
                        Format(RV_Long, "#,##0.00")
        H1_ZaiRyo.Caption = RTrim(RV_Zaist) & Space(25 - Len(RTrim(RV_Zaist))) & "/" & _
                        Format(RV_Itatu, "0000.00") & "x" & _
                        Format(RV_Width, "0000.00") & "x" & _
                        Format(RV_Long, "0000.00")
                        Call H1_ZaiBox_Set
        H1lb_ZaiKan = Rv_KanriNo
        
        gSL_Select = "SELECT * FROM 材料取引先マスタ WHERE 材料管理番号 = " & Rv_KanriNo
        Call ZTMRead(gSL_Select)
        H1lb_HinBn(2) = ZTM.自社品番
    
        For cnt = 0 To 3
            If H1Op_Hinku(cnt).Value = True Then
               iH1_Hinbn = H1lb_HinBn(cnt)
               cnt = 4
            End If
        Next
    End If
End Sub
''+------------------------+
''+     材料取引先検索     +
''+------------------------+
Private Sub Look_ZaiTor(Kei As Integer)
    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 1
    RVI_SK = 9
    RVI_Tehai = ""      ' 2009.10.29 upd : 0-->"" 手配区分初期値=空白

    CKK0077.Show vbModal
    Unload CKK0077
    Set CKK0077 = Nothing

    If RV_Rtn Then
       If Kei = 0 Then
          iH1_Torcd = RV_TorcdK
       Else
          iH1_KTorcd = RV_TorcdK
       End If
'      iH1_Torcd = RV_Torcd
       H1lb_Tornm = RV_TorRName
       
       H1lb_HinBn(0) = RV_HHinbn
       H1lb_HinBn(1) = RV_JHinbn
       H1lb_HinBn(3) = RV_Zaist & " " & _
                       Format(RV_Itatu, "#,##0.00") & "X" & _
                       Format(RV_Width, "#,##0.00") & "X" & _
                       Format(RV_Long, "#,##0.00")
        H1_ZaiRyo.Caption = RTrim(RV_Zaist) & Space(25 - Len(RTrim(RV_Zaist))) & "/" & _
                            Format(RV_Itatu, "0000.00") & "x" & _
                            Format(RV_Width, "0000.00") & "x" & _
                            Format(RV_Long, "0000.00")
       Call H1_ZaiBox_Set
       H1lb_ZaiKan = Rv_KanriNo
       
       gSL_Select = "SELECT * FROM 材料取引先マスタ WHERE 材料管理番号 = " & Rv_KanriNo
       Call ZTMRead(gSL_Select)
       H1lb_HinBn(2) = ZTM.自社品番
    
       For cnt = 0 To 3
           If H1Op_Hinku(cnt).Value = True Then
              iH1_Hinbn = H1lb_HinBn(cnt)
              cnt = 4
           End If
       Next
       
       Call iH1_Torcd_LostFocus
       Call iH1_KTorcd_LostFocus
    End If
End Sub
'-----------------------------------------------' 2009.10.02 add end

'+----------------------------+
'+       名称マスタ検索        +
'+----------------------------+
Private Sub Look_Naiyo()

    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "単価内訳"
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        iB1_EtcMsg(CisFun.Val2(iB1_EtcMsg(0).Tag)) = RV_Name
    End If

End Sub
'+-----------------------------+
'+  iH1_Hinbn
'+-----------------------------+
Private Sub iH1_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    Call HinbnToLabel   ' 2009.10.06 add
    
    Call Zai_Chk

    End Sub
'+-----------------------------+
'+  iH1_Torcd
'+-----------------------------+
Private Sub iH1_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H1lb_Tornm = ""
    If TorNmGet(iH1_Torcd, 8, 1) Then
       H1lb_Tornm = TRM_RName
    End If

    Call Zai_Chk
    
End Sub
'+-----------------------------+
'+  iH1_KTorcd
'+-----------------------------+
Private Sub iH1_KTorcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_KTorcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H1lb_KTornm = ""
    If TorNmGet(iH1_KTorcd, 5, 1) Then
       H1lb_KTornm = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iB1_Tanka
'+-----------------------------+
Private Sub iB1_Tanka_LostFocus()
    If CisFun.Val2(iB1_BTanka) = 0 Then
        iB1_BTanka = iB1_Tanka
    End If
End Sub
'+-----------------------------+
'+  B1Ch_Kari
'+-----------------------------+
Private Sub B1Ch_Kari_Click()
    If B1Ch_Kari.Value = 1 Then
       B1Ch_Kari.Caption = "仮"
    Else
       B1Ch_Kari.Caption = "正"
    End If
End Sub
'+-----------------------------+
'+  iB1_TSYmd
'+-----------------------------+
Private Sub iB1_TSYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_TSYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_TEYmd
'+-----------------------------+
Private Sub iB1_TEYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_TEYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  B1Ch_UMu
'+-----------------------------+
Private Sub B1Ch_UMu_Click()
    If B1Ch_UMu.Value = 1 Then
       B1Ch_UMu.Caption = "無償"
    Else
       B1Ch_UMu.Caption = "有償"
    End If
End Sub
'+-----------------------------+
'+  iB1_EtcMsg
'+-----------------------------+
Private Sub iB1_EtcMsg_GotFocus(INDEX As Integer)
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    iB1_EtcMsg(0).Tag = INDEX
End Sub
Private Sub iB1_EtcMsg_LostFocus(INDEX As Integer)
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
        iB1_EtcMsg(0).Tag = ""
    End If
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim W_Color     As Long
   W_Color = &H80000008 '初期色ｾｯﾄ(黒)
   Select Case SyoriKB
     Case "ADD"
          SyoriNM = "追加"
          K_Sykbnm.ForeColor = &HFFFF&   '(黄色)
          PB_ADD.Visible = False
     Case "MNT"
          SyoriNM = "修正"
          K_Sykbnm.ForeColor = &HFF0000  '(青色)
     Case "DEL"
          SyoriNM = "削除"
          K_Sykbnm.ForeColor = &HFF&     '(赤色)
          W_Color = &HFF&
     Case "LOK"
          SyoriNM = "照会"
          K_Sykbnm.ForeColor = &HFF00FF   '(ピンク)
          PB_ADD.Visible = False
          PB_MNT.Visible = False
          PB_DEL.Visible = False
     Case Else
          SyoriNM = ""
   End Select
   For Each gObj In Me.Controls
       If gObj.Name Like "i*" Then gObj.ForeColor = W_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = W_Color
'      * PF-Key 表示設定 *
       If SyoriKB <> "LOC" Then
          If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
             gObj.Name = "PB_DEL" Then
             gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
          End If
       End If
   Next gObj
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
'****************************
'*                          *
'*      HEAD/BODY 処理      *
'*                          *
'****************************
'****************************
'*      入力切替 処理        *
'****************************
Private Function DispChange(Pro As String)
    Dim H0Mode      As Boolean          ' 2009.09.30 add
    Dim H0Color     As Long             ' 2009.09.30 add
    Dim H1Mode      As Boolean
    Dim H1Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    H0Mode = False                      ' 2009.09.30 add
    H0Color = gPLostSel                 ' 2009.09.30 add
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
       Case "H0"                        ' 2009.09.30 add
          H0Mode = True                 ' 2009.09.30 add
          H0Color = gPGotSel            ' 2009.09.30 add
       Case "H1"
          H1Mode = True
          H1Color = gPGotSel
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
    End Select
    
    H0_Area1.Enabled = H0Mode           ' 2009.09.30 add
    H0_Area1.BackColor = H0Color        ' 2009.09.30 add
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    PB_Look.Visible = H1Mode
    PB_ADD.Visible = H1Mode
    PB_MNT.Visible = H1Mode
    PB_DEL.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = Not H0Mode         ' 2009.09.30 upd : H1Mode --> H0Mode
    K_Sykbnm.Visible = Not H0Mode       ' 2009.09.30 add
    PB_END.Visible = H0Mode             ' 2009.09.30 add
'    PB_END.Visible = H1Mode
    '------------------------------ 2005/06/14  Insert
    If Rv_Call3 <> "" Then
       If Rv_Call3 = "X" Then
'          PB_CAN.Visible = False
       End If
       PB_END.BtnCaption = "戻る"
    End If
    '------------------------------ 2005/06/14  Insert End
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
            If .Name Like wbc_Enm & "*Ch_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
                .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
       On Error Resume Next
       Call CisVsGrid.Clear
       On Error GoTo 0
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
'-----------------------------------------------' 2009.10.01 add start
    If ProcHB = "H0" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
        iH1_Hinbn.Tag = "1"
        iH1_Torcd.Tag = ""
        If H0Op_Tanku(1).Value = True Then
        ''''H1lb_HinBn(0).Caption = "発注品番です": H1lb_HinBn(1).Caption = "受給品番です": H1lb_HinBn(2).Caption = "自社品番です": H1lb_HinBn(3).Caption = "材料/寸法です" ' テスト表示
            H1Op_Hinku(0).Value = True
        Else
            For cnt = 0 To 3
                H1Op_Hinku(cnt).Visible = False
                H1lb_HinBn(cnt).Visible = False
                b_HinNM(cnt).Visible = False
            Next
        End If
        Call iH1_Hinbn.SetFocus
        gMaeIndex = 0
        GoTo ReturnPress_Ed
    End If
'-----------------------------------------------' 2009.10.01 add end
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
' << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        Else
            ProcHB = "B1"
            Call DispChange(ProcHB)
            iB1_Tanka.SetFocus          ' 2009.09.30 add
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        '------------------------------ 2005/06/14  Insert
        If Rv_Call3 = "X" Then
           Rv_Call3Rtn = 1
           Call PB_END_Click
        End If
        '------------------------------ 2005/06/14  Insert End
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        Call DBInput("INV")             ' 2009.09.30 add
        Call iH1_Hinbn.SetFocus         ' 2009.10.01 add
        GoTo ReturnPress_Ed
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'---------------------------------------------------------------' 2009.09.30 del start
''****************************
''*      HEAD1ﾁｪｯｸ処理
''****************************
'Private Function Head1Chk() As Boolean
'    Head1Chk = False
'
'    If Trim(iH1_Hinbn) = "" Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    品番を入力して下さい            "
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'        End With
'       iH1_Hinbn.SetFocus
'       Exit Function
'    End If
'
'    gSL_Select = " SELECT * "
'    gSL_Select = gSL_Select & " FROM 品番マスタ "
'    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
'    If Not HNMRead(gSL_Select, 1) Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    品番マスタ未登録            "
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'        End With
'       iH1_Hinbn.SetFocus
'       Exit Function
'    End If
'
'    If Trim(iH1_Torcd) = "" Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    仕入先を入力して下さい            "
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'        End With
'       iH1_Torcd.SetFocus
'       Exit Function
'    End If
'    If Not TorNmGet(iH1_Torcd, 8) Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    仕入先　未登録！            "
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'        End With
'       iH1_Torcd.SetFocus
'       Exit Function
'    End If
''契約先                                         2008/01/21追加
'    If Trim(iH1_KTorcd) <> "" Then
'        If Not TorNmGet(iH1_KTorcd, 5) Then
'            With CisFun
'                .MB_Lines = 4
'                .MB_MSG(2) = "    契約先　未登録！            "
'                .MB_Title = ""
'                .MB_Button = Error
'                .MBOX
'            End With
'           iH1_KTorcd.SetFocus
'           Exit Function
'        End If
'    End If
'
'' マスタ入力 SELECT文 ｾｯﾄ
'    If DBInput("INV") Then
'       If SyoriKB = "ADD" Then
'          With CisFun
'              .MB_Lines = 4
'              .MB_MSG(2) = "    購入単価マスタ 登録済！            "
'              .MB_Title = ""
'              .MB_Button = Error
'              .MBOX
'          End With
'          iH1_Hinbn.SetFocus
'          Exit Function
'       End If
'    Else
'       If SyoriKB <> "ADD" Then
'          With CisFun
'              .MB_Lines = 4
'              .MB_MSG(2) = "    購入単価マスタ 未登録！            "
'              .MB_Title = ""
'              .MB_Button = Error
'              .MBOX
'          End With
'          iH1_Hinbn.SetFocus
'          Exit Function
'       End If
'    End If
'    For gInt = 1 To 6
'        iB1_EtcMsg(gInt - 1) = EtcMsg(gInt)
'        If EtcTan(gInt) <> 0 Then
'            iB1_EtcTan(gInt - 1) = Format(EtcTan(gInt), "#,##0.00")
'        End If
'    Next
'
'    Head1Chk = True
'End Function
'---------------------------------------------------------------' 2009.09.30 del end
        
'---------------------------------------------------------------' 2009.10.06 add start
Private Sub HinbnToLabel()
    If H0Op_Tanku(1).Value = False Then Exit Sub
    
    If H1Op_Hinku(0).Value = True Then
       H1lb_HinBn(0).Caption = Trim(iH1_Hinbn.Text)
    ElseIf H1Op_Hinku(1).Value = True Then
       H1lb_HinBn(1).Caption = Trim(iH1_Hinbn.Text)
    ElseIf H1Op_Hinku(2).Value = True Then
       H1lb_HinBn(2).Caption = Trim(iH1_Hinbn.Text)
    Else
       H1lb_HinBn(3).Caption = "": H1_ZaiRyo.Caption = "": iH1_Hinbn.Text = ""
       If Trim(iH1_Zaist.Text) <> "" Then
          H1lb_HinBn(3).Caption = Trim(iH1_Zaist.Text) & " " & _
                                  Format(CisFun.Val2(iH1_Itatu.Text), "###0.00") & "X" & _
                                  Format(CisFun.Val2(iH1_Width.Text), "###0.00") & "X" & _
                                  Format(CisFun.Val2(iH1_Long.Text), "###0.00")
          H1_ZaiRyo.Caption = Trim(iH1_Zaist.Text) & Space(25 - Len(Trim(iH1_Zaist.Text))) & "/" & _
                              Format(CisFun.Val2(iH1_Itatu.Text), "0000.00") & "x" & _
                              Format(CisFun.Val2(iH1_Width.Text), "0000.00") & "x" & _
                              Format(CisFun.Val2(iH1_Long.Text), "0000.00")
          iH1_Hinbn.Text = H1lb_HinBn(3).Caption
       End If
    End If

    For cnt = 0 To 3
        If Trim(H1lb_HinBn(cnt)) <> "" Then Exit For
    Next
    If cnt > 3 Then
        H1lb_ZaiKan.Caption = ""        ' 全ての品番が空白ならば材料管理番号を無しにする
    End If
End Sub
'---------------------------------------------------------------' 2009.10.06 add end

'---------------------------------------------------------------' 2009.09.30 add start
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
    Call HinbnToLabel       ' 2009.10.06 add
    
    '============== 仕入先必須 =================
    If Trim(iH1_Torcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    仕入先を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
    '============== 仕入先登録 =================
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/19 UPDATE START
    'If Not TorNmGet(iH1_Torcd, 8) Then
    If Not TorNmGet(iH1_Torcd, 9) Then
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/19 UPDATE END
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    仕入先未登録            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If

    If RTrim(iH1_KTorcd) <> "" Then
       If Not TorNmGet(iH1_KTorcd, 5) Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    契約先未登録            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          iH1_KTorcd.SetFocus
          Exit Function
       End If
    End If
    
    
    If H0Op_Tanku(0).Value Then
    '============================================
    '============= 部品区分：品番 ===============
    '============================================
       
       If RTrim(iH1_Hinbn) = "" Then

          '============== 品番取引先登録 =================
          gSL_Select = ""
          gSL_Select = gSL_Select & " SELECT 品番, 背番号" & vbCrLf
          gSL_Select = gSL_Select & "   FROM 品番取引先マスタ" & vbCrLf
          gSL_Select = gSL_Select & "  WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
          ' 契約先が空白でなければ契約先も含む                                                  ' 2009.10.01 add
          If RTrim(iH1_KTorcd) <> "" Then                                                       ' 2009.10.01 add
             gSL_Select = gSL_Select & "    AND 契約先 = '" & RTrim(iH1_KTorcd) & "'" & vbCrLf  ' 2009.10.01 add
          End If                                                                                ' 2009.10.01 add
'          gSL_Select = gSL_Select & "    AND 背番号 = '" & RTrim(iH1_Seban) & "'" & vbCrLf
          gSL_Select = gSL_Select & "    AND 品目 = 1"
          gSL_Select = gSL_Select & "  GROUP BY 品番, 背番号"
          If KTMRead(gSL_Select, , 1) Then
             If CisDB.RecordCount = 1 Then
                iH1_Hinbn = RTrim(KTM.品番)
'                iH1_Seban = RTrim(KTM.背番号)
             Else
                RV_Torcd = RTrim(iH1_Torcd)
                RV_TorRName = RTrim(H1lb_Tornm)
                RV_Seban = ""       ' RV_Seban = RTrim(iH1_Seban)
                 
                RV_Top = Me.Top + 3200
                RV_Left = Me.Left + 2300
                  
                RVI_Hinmoku = 1
                  
                CHinbnSelect.Show vbModal
                Unload CHinbnSelect
                Set CHinbnSelect = Nothing
                If RV_Rtn Then
                   iH1_Hinbn = RV_Hinbn
'                   iH1_Hinbn.SetFocus
'                   Exit Function
                Else
                   iH1_Torcd.SetFocus
                   Exit Function
                End If
             End If
          Else
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    品番を入力して下さい            "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
             iH1_Hinbn.SetFocus
             Exit Function
          End If
       End If
       
       '=============== 品番マスタ登録 ================
       gSL_Select = " SELECT * "
       gSL_Select = gSL_Select & " FROM 品番マスタ "
       gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
       If Not HNMRead(gSL_Select, 1) Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    品番マスタ未登録            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          iH1_Hinbn.SetFocus
          Exit Function
       End If
       
'---------------------------------------------------------------------------------------------------' 2009.10.02 del start
       '=============== 品番取引先マスタ登録 ================
'       gSL_Select = ""                                                     ' 背番号抽出なので削除（背番号未使用のため）
'       gSL_Select = gSL_Select & " SELECT * FROM 品番取引先マスタ"
'       gSL_Select = gSL_Select & "  WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
'       gSL_Select = gSL_Select & "    AND 取引先 = '" & RTrim(iH1_Torcd) & "'"
'       gSL_Select = gSL_Select & "    AND 品目 = 1"
'       If KTMRead(gSL_Select, , 1) Then
'          iH1_Seban = KTM.背番号
'       Else
'          iH1_Seban = ""
'       End If
'---------------------------------------------------------------------------------------------------' 2009.10.02 del end

'---------------------------------------------------------------------------------------------------' 2009.10.29 add start
        '=============== 手配区分確認 ==================
        gSL_Select = ""
        gSL_Select = gSL_Select & "SELECT * FROM 品番取引先マスタ "
        gSL_Select = gSL_Select & " WHERE 品番 = '" & Trim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & "   AND 取引先 = '" & Trim(iH1_Torcd) & "'"
        If Not HTMRead(gSL_Select) Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    品番取引先マスタ　未登録        "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
             iH1_Hinbn.SetFocus
             Exit Function
        End If
        If HTM.手配区分 <> "0" And HTM.手配区分 <> "1" And HTM.手配区分 <> "3" Then
        With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    手配区分が違います          "
              .MB_MSG(4) = "     (手配区分：なし)           "
              If HTM.手配区分 = "2" Then .MB_MSG(4) = "     (手配区分：社内)           "
              If HTM.手配区分 = "4" Then .MB_MSG(4) = "     (手配区分：支給)           "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          iH1_Hinbn.SetFocus
          Exit Function
        End If
'---------------------------------------------------------------------------------------------------' 2009.10.29 add end
   
   'マスタ入力 SELECT文 ｾｯﾄ
       '=============== 購入単価マスタ登録 ================
       If DBInput("INV") Then
          If SyoriKB = "ADD" Then
             If Rv_Call3 <> "" Then
                SyoriKB = "MNT"
                Call SyoriName(SyoriKB)
                Head1Chk = True
                Exit Function
             End If
         
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    購入単価マスタ登録済            "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
             iH1_Torcd.SetFocus
             Exit Function
          End If
       Else
          If SyoriKB <> "ADD" Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    購入単価マスタ未登録            "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
             iH1_Torcd.SetFocus
             Exit Function
          End If
       End If
    Else
    '============================================
    '============= 部品区分：材料 ===============
    '============================================
       If RTrim(iH1_Hinbn) = "" And RTrim(H1lb_ZaiKan) = "" Then
          '=================== 材料取引先登録 =======================
          gSL_Select = ""
          gSL_Select = gSL_Select & " SELECT ISNULL(ZT.受給品番,'') 受給品番, ISNULL(ZT.発注品番,'') 発注品番, ISNULL(ZT.自社品番,'') 自社品番," 'ZT.背番号 削除
          gSL_Select = gSL_Select & "        ZM.材質, ZM.材料管理番号, ZT.板厚, ZT.幅, ZT.長さ, ZT.契約先" & vbCrLf     ' ZT.契約先追加
          gSL_Select = gSL_Select & "   FROM 材料取引先マスタ ZT" & vbCrLf
          gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料マスタ ZM"
          gSL_Select = gSL_Select & "             ON ZM.材料管理番号 = ZT.材料管理番号"
          gSL_Select = gSL_Select & "  WHERE ZT.取引先 = '" & RTrim(iH1_Torcd) & "' " & vbCrLf
          If RTrim(iH1_KTorcd) <> "" Then
             gSL_Select = gSL_Select & "    AND ZT.契約先 = '" & RTrim(iH1_KTorcd) & "'" & vbCrLf
          End If
'          gSL_Select = gSL_Select & "    AND ZT.背番号 = '" & RTrim(iH1_Seban) & "'" & vbCrLf
'          gSL_Select = gSL_Select & "    AND ZT.品目 = 1"

          If ZTMRead(gSL_Select, , 1) Then
             If CisDB.RecordCount = 1 Then
                H1lb_HinBn(0) = RTrim(ZTM.発注品番)
                H1lb_HinBn(1) = RTrim(ZTM.受給品番)
                H1lb_HinBn(2) = RTrim(ZTM.自社品番)
                H1lb_HinBn(3) = RTrim(ZTM.材質) & " " & _
                                Format(ZTM.板厚, "#,##0.00") & "X" & _
                                Format(ZTM.幅, "#,##0.00") & "X" & _
                                Format(ZTM.長さ, "#,##0.00")
                H1_ZaiRyo.Caption = RTrim(ZKT.材質) & Space(25 - Len(RTrim(ZKT.材質))) & "/" & _
                                Format(ZKT.板厚, "0000.00") & "x" & _
                                Format(ZKT.幅, "0000.00") & "x" & _
                                Format(ZKT.長さ, "0000.00")
                Call H1_ZaiBox_Set
                For cnt = 0 To 3
                    If H1Op_Hinku(cnt).Value = True Then
                       Exit For
                    End If
                Next
                If cnt > 3 Then cnt = 0
                iH1_Hinbn = H1lb_HinBn(cnt)
                H1lb_ZaiKan = RTrim(ZTM.材料管理番号)
                iH1_KTorcd = RTrim(ZTM.契約先)
                Call iH1_KTorcd_LostFocus

             Else   '------ 複数件ある場合は選択 -------
                RV_Torcd = RTrim(iH1_Torcd)
                RV_TorRName = RTrim(H1lb_Tornm)
                RV_Seban = ""                   ' RV_Seban = ""RTrim(iH1_Seban)
                RVI_Hinmoku = 1
                RV_KTorcd = ""
                
                RV_Top = Me.Top + 3200
                RV_Left = Me.Left + 2300

                CZairyoSelectZSM.Show vbModal
                Unload CZairyoSelectZSM
                Set CZairyoSelectZSM = Nothing
                
                If RV_Rtn Then
                   H1lb_HinBn(0) = RV_HHinbn
                   H1lb_HinBn(1) = RV_JHinbn
                   H1lb_HinBn(2) = RV_Zainm
                   H1lb_HinBn(3) = RV_Zaist & " " & _
                                   Format(RV_Itatu, "#,##0.00") & "X" & _
                                   Format(RV_Width, "#,##0.00") & "X" & _
                                   Format(RV_Long, "#,##0.00")
                   H1_ZaiRyo.Caption = RTrim(RV_Zaist) & Space(25 - Len(RTrim(RV_Zaist))) & "/" & _
                                       Format(RV_Itatu, "0000.00") & "x" & _
                                       Format(RV_Width, "0000.00") & "x" & _
                                       Format(RV_Long, "0000.00")
                   Call H1_ZaiBox_Set
                   For cnt = 0 To 3
                       If H1Op_Hinku(cnt).Value = True Then
                          Exit For
                       End If
                   Next
                   If cnt > 3 Then cnt = 0
                   iH1_Hinbn = H1lb_HinBn(cnt)
                   H1lb_ZaiKan = Rv_KanriNo
                   iH1_KTorcd = RV_KTorcd
                   Call iH1_KTorcd_LostFocus
                Else
                   iH1_Torcd.SetFocus
                   Exit Function
                End If
             End If
          Else
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    材料マスタ未登録            "
'                 If H1Op_Hinku(0).Value Then
'                    .MB_MSG(3) = "    発注品番を入力して下さい            "
'                 ElseIf H1Op_Hinku(1).Value Then
'                    .MB_MSG(3) = "    受給品番を入力して下さい            "
'                 ElseIf H1Op_Hinku(2).Value Then
'                    .MB_MSG(3) = "    自社品番を入力して下さい            "
'                 Else
'                    .MB_MSG(3) = "    材質/寸法を入力して下さい           "
'                 End If
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
             iH1_Hinbn.SetFocus
             Exit Function
          End If
       End If
       
'-------------------------------------------------------' 2009.10.02 del start
'       '========= 契約先確認 =================
'       If RTrim(iH1_KTorcd) = "" Then
'             With CisFun
'                 .MB_Lines = 4
'                 .MB_MSG(2) = "    契約先を入力してください            "
'                 .MB_Title = ""
'                 .MB_Button = Error
'                 .MBOX
'             End With
'             iH1_KTorcd.SetFocus
'             Exit Function
'       End If
'-------------------------------------------------------' 2009.10.02 del end
       
       '========= 材料マスタ／材料取引先マスタ登録確認 =========
       gSL_Select = " SELECT * "
'       gSL_Select = gSL_Select & " FROM  材料マスタ "
'       gSL_Select = gSL_Select & " WHERE 材料管理番号 = " & CisFun.Val2(H1lb_ZaiKan)
       gSL_Select = gSL_Select & " FROM  材料マスタ ZM"
       gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料取引先マスタ ZT"
       gSL_Select = gSL_Select & "             ON ZT.材料管理番号 = ZM.材料管理番号"
       gSL_Select = gSL_Select & "  WHERE ZT.取引先 = '" & RTrim(iH1_Torcd) & "' " & vbCrLf
       If RTrim(iH1_KTorcd) <> "" Then  ' 2009.10.02 add
          gSL_Select = gSL_Select & "    AND ZT.契約先 = '" & RTrim(iH1_KTorcd) & "'" & vbCrLf
       End If                           ' 2009.10.02 add
       If H1Op_Hinku(0) Then
          gSL_Select = gSL_Select & "    AND ZT.発注品番 = '" & Trim(H1lb_HinBn(0).Caption) & "'" & vbCrLf
       End If
       If H1Op_Hinku(1) Then
          gSL_Select = gSL_Select & "    AND ZT.受給品番 = '" & Trim(H1lb_HinBn(1).Caption) & "'" & vbCrLf
       End If
       If H1Op_Hinku(2) Then
          gSL_Select = gSL_Select & "    AND ZT.自社品番 = '" & Trim(H1lb_HinBn(2).Caption) & "'" & vbCrLf
       End If
       If H1Op_Hinku(3) Then
          gSL_Select = gSL_Select & "    AND ZM.材質     = '" & Trim(Mid(H1_ZaiRyo.Caption, 1, 25)) & "'" & vbCrLf
          gSL_Select = gSL_Select & "    AND ZM.板厚     =  " & Format(CisFun.Val2(Mid(H1_ZaiRyo.Caption, 27, 7)), "###0.00") & vbCrLf
          gSL_Select = gSL_Select & "    AND ZM.幅       =  " & Format(CisFun.Val2(Mid(H1_ZaiRyo.Caption, 35, 7)), "###0.00") & vbCrLf
          gSL_Select = gSL_Select & "    AND ZM.長さ     =  " & Format(CisFun.Val2(Mid(H1_ZaiRyo.Caption, 43, 7)), "###0.00") & vbCrLf
       End If
       
       If Not ZRMRead(gSL_Select, , 1) Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    材料マスタ未登録            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          iH1_Hinbn.SetFocus
          Exit Function
       End If
       
       H1lb_ZaiKan = ZRM.材料管理番号
       
'-----------------------------------------------------------' 2009.10.01 del start
'       '======== 材料取引先マスタ登録 ========
'       gSL_Select = ""
'       gSL_Select = gSL_Select & " SELECT * FROM 材料取引先マスタ ZT"
'       gSL_Select = gSL_Select & "          LEFT OUTER JOIN 材料マスタ ZM"
'       gSL_Select = gSL_Select & "               ON ZM.材料管理番号 = ZT.材料管理番号"
'       If H1Op_Hinku(0).Value Then
'          If RTrim(iH1_Hinbn) <> "" Then
'             gSL_Select = gSL_Select & "  WHERE ISNULL(ZT.発注品番,'') = '" & RTrim(iH1_Hinbn) & "'"
'          Else
'             gSL_Select = gSL_Select & "  WHERE ISNULL(ZT.材料管理番号,0) = " & CisFun.Val2(H1lb_ZaiKan)
'          End If
'       Else
'          If RTrim(iH1_Hinbn) <> "" Then
'             gSL_Select = gSL_Select & "  WHERE ISNULL(ZT.受給品番,'') = '" & RTrim(iH1_Hinbn) & "'"
'          Else
'             gSL_Select = gSL_Select & "  WHERE ISNULL(ZT.材料管理番号,0) = " & CisFun.Val2(H1lb_ZaiKan)
'          End If
'       End If
'       gSL_Select = gSL_Select & "    AND ZT.取引先 = '" & RTrim(iH1_Torcd) & "'"
'       gSL_Select = gSL_Select & "    AND ZT.品目 = 1"
'       If ZTMRead(gSL_Select, , 1) Then
'          iH1_Seban = ZTM.背番号
'       Else
'          iH1_Seban = ""
'       End If
'-----------------------------------------------------------' 2009.10.01 del end
'---------------------------------------------------------------------------------------------------' 2009.10.29 add start
        '=============== 手配区分確認 ==================
        gSL_Select = ""
        gSL_Select = gSL_Select & "SELECT ZT.* "
        gSL_Select = gSL_Select & " FROM  材料取引先マスタ ZT"
        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料マスタ ZM"
        gSL_Select = gSL_Select & "             ON ZM.材料管理番号 = ZT.材料管理番号"
        gSL_Select = gSL_Select & "  WHERE ZT.取引先 = '" & RTrim(iH1_Torcd) & "' " & vbCrLf
       'gSL_Select = gSL_Select & "    AND ZT.契約先 = '" & RTrim(iH1_KTorcd) & "'" & vbCrLf
        If H1Op_Hinku(0) Then
           gSL_Select = gSL_Select & "    AND ZT.発注品番 = '" & Trim(H1lb_HinBn(0).Caption) & "'" & vbCrLf
        End If
        If H1Op_Hinku(1) Then
           gSL_Select = gSL_Select & "    AND ZT.受給品番 = '" & Trim(H1lb_HinBn(1).Caption) & "'" & vbCrLf
        End If
        If H1Op_Hinku(2) Then
           gSL_Select = gSL_Select & "    AND ZT.自社品番 = '" & Trim(H1lb_HinBn(2).Caption) & "'" & vbCrLf
        End If
        If H1Op_Hinku(3) Then
           gSL_Select = gSL_Select & "    AND ZM.材質     = '" & Trim(Mid(H1_ZaiRyo.Caption, 1, 25)) & "'" & vbCrLf
           gSL_Select = gSL_Select & "    AND ZM.板厚     =  " & Format(CisFun.Val2(Mid(H1_ZaiRyo.Caption, 27, 7)), "###0.00") & vbCrLf
           gSL_Select = gSL_Select & "    AND ZM.幅       =  " & Format(CisFun.Val2(Mid(H1_ZaiRyo.Caption, 35, 7)), "###0.00") & vbCrLf
           gSL_Select = gSL_Select & "    AND ZM.長さ     =  " & Format(CisFun.Val2(Mid(H1_ZaiRyo.Caption, 43, 7)), "###0.00") & vbCrLf
        End If
        
        If Not ZTMRead(gSL_Select) Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    材料取引先マスタ　未登録      "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
             iH1_Torcd.SetFocus
             Exit Function
        End If
        
        If ZTM.手配区分 <> 0 And ZTM.手配区分 <> 1 And ZTM.手配区分 <> 3 Then
           With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    手配区分が違います          "
              .MB_MSG(4) = "     (手配区分：なし)           "
              If ZTM.手配区分 = 2 Then .MB_MSG(4) = "     (手配区分：社内)           "
              If ZTM.手配区分 = 4 Then .MB_MSG(4) = "     (手配区分：支給)           "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
           End With
           iH1_Torcd.SetFocus
           Exit Function
        End If
'---------------------------------------------------------------------------------------------------' 2009.10.29 add end
   
   'マスタ入力 SELECT文 ｾｯﾄ
       '======== 材料購入単価マスタ登録 =========
       If DBInput("INV") Then
          If SyoriKB = "ADD" Then
             If Rv_Call3 <> "" Then
                SyoriKB = "MNT"
                Call SyoriName(SyoriKB)
                Head1Chk = True
                Exit Function
             End If
         
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    材料購入単価マスタ登録済            "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
             iH1_Torcd.SetFocus
             Exit Function
          End If
       Else
          If SyoriKB <> "ADD" Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    材料購入単価マスタ未登録            "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
             iH1_Torcd.SetFocus
             Exit Function
          End If
       End If
    End If
    
    Head1Chk = True
End Function
'---------------------------------------------------------------' 2009.09.30 add end
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    Dim wEtcTan     As Currency
'入力確認
'    If CisFun.Val2(iB1_Tanka) = 0 Then                         '2019/07/19 単価0円入力可
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    単価を入力して下さい            "
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'        End With
'       iB1_Tanka.SetFocus
'       Exit Function
'    End If
    If Trim(iB1_TSYmd) <> "" Then
       If Not iB1_TSYmd.cChkResult Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    適用開始日を正しい日付にて入力して下さい            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          iB1_TSYmd.SetFocus
       Exit Function
       End If
    End If
    If Trim(iB1_TEYmd) <> "" Then
       If Not iB1_TEYmd.cChkResult Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    適用終了日を正しい日付にて入力して下さい            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          iB1_TEYmd.SetFocus
       Exit Function
       End If
    End If
    If Trim(iB1_TSYmd) <> "" And Trim(iB1_TEYmd) <> "" Then
       If Trim(iB1_TSYmd) > Trim(iB1_TEYmd) Then
          With CisFun
              .MB_Lines = 4
              .MB_MSG(2) = "    適用期間を正しくにて入力して下さい            "
              .MB_MSG(4) = "    　　　(開始＞終了）                            "
              .MB_Title = ""
              .MB_Button = Error
              .MBOX
          End With
          iB1_TSYmd.SetFocus
       Exit Function
       End If
    End If

    If H0Op_Tanku(0).Value = True Then        ' 2009.10.01 add
'単価 = 部品単価+加工単価+その他単価 確認
        If CisFun.Val2(iB1_Tanka) <> CisFun.Val2(iB1_BTanka) + CisFun.Val2(iB1_KTanka) + CisFun.Val2(iB1_STanka) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    単価 <> 部品単価+加工単価+その他単価です。            "
                .MB_MSG(4) = "    　　よろしいですか？                      "
                .MB_Title = ""
                .MB_Button = OK_CAN
                If Not .MBOX Then
                   iB1_Tanka.SetFocus
                   Exit Function
                End If
            End With
        End If
'その他単価 = その他単価内訳合計　確認
        wEtcTan = 0
        For gInt = 0 To 5
            If CisFun.Val2(iB1_EtcTan(gInt)) <> 0 Then
                wEtcTan = wEtcTan + CisFun.Val2(iB1_EtcTan(gInt))
            End If
        Next
        If CisFun.Val2(iB1_STanka) <> wEtcTan Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    その他単価 <> その他単価内訳合計です。            "
                .MB_MSG(4) = "    　　よろしいですか？                      "
                .MB_Title = ""
                .MB_Button = OK_CAN
                If Not .MBOX Then
                   iB1_STanka.SetFocus
                   Exit Function
                End If
            End With
        End If
    End If                                    ' 2009.10.01 add
    
    Body1Chk = True
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    
    Call HeadBodyClear("B")
    
    gSL_Select = "Select 履歴NO,単価,有無償区分,"
    gSL_Select = gSL_Select & "    Case When 仮単価区分 = '' Then '' Else '仮' End 仮単価区分,"
    gSL_Select = gSL_Select & "    CASE WHEN ISNULL(適用開始日付,'') <> '' THEN"
    gSL_Select = gSL_Select & "        SUBSTRING(適用開始日付,1,4) + '/' + SUBSTRING(適用開始日付,5,2) + '/' +"
    gSL_Select = gSL_Select & "        SUBSTRING(適用開始日付,7,2)"
    gSL_Select = gSL_Select & "    END 適用開始日付,"
    gSL_Select = gSL_Select & "    CASE WHEN ISNULL(適用終了日付,'') <> '' THEN"
    gSL_Select = gSL_Select & "        SUBSTRING(適用終了日付,1,4) + '/' + SUBSTRING(適用終了日付,5,2) + '/' +"
    gSL_Select = gSL_Select & "        SUBSTRING(適用終了日付,7,2)"
    gSL_Select = gSL_Select & "    END 適用終了日付,"
    gSL_Select = gSL_Select & "    部品単価,加工単価,その他単価"
    If H0Op_Tanku(0).Value = True Then                                                  ' 2009.10.01 add
       gSL_Select = gSL_Select & "    From 購入単価マスタ"
       gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
       gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(iH1_Torcd) & "'"
       gSL_Select = gSL_Select & "   And 契約先 = '" & RTrim(iH1_KTorcd) & "'"
    Else                                                                                ' 2009.10.01 add
       gSL_Select = gSL_Select & "    From 材料購入単価マスタ"                          ' 2009.10.01 add
       gSL_Select = gSL_Select & " WHERE 材料管理番号 = " & CisFun.Val2(H1lb_ZaiKan)    ' 2009.10.01 add
    End If                                                                              ' 2009.10.01 add
    
    gSL_Select = gSL_Select & " Order By 履歴NO Desc "
    If Not KTMRead(gSL_Select, 0) Then
        Call KTMClose(0)
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
        Call GridSet
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
    
End Sub
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+      削 除 処 理      +
'+----------------------+
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
    
    With CisDB
        On Error GoTo DBDelete_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
        If H0Op_Tanku(0).Value = True Then                              ' 2009.10.01 add
            ' ( 削除処理 )
            .SQL = "DELETE FROM 購入単価マスタ "
            .SQL = .SQL & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
            .SQL = .SQL & "   And 取引先 = '" & RTrim(iH1_Torcd) & "'"
            .SQL = .SQL & "   And 契約先 = '" & RTrim(iH1_KTorcd) & "'"
            On Error GoTo DBDelete_Err
            .DBExec
            ' ( 削除処理 )
            .SQL = "DELETE FROM 購入単価内訳マスタ "
            .SQL = .SQL & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
            .SQL = .SQL & "   And 取引先 = '" & RTrim(iH1_Torcd) & "'"
            .SQL = .SQL & "   And 契約先 = '" & RTrim(iH1_KTorcd) & "'"
            On Error GoTo DBDelete_Err
            .DBExec
'-----------------------------------------------------------------------' 2009.10.01 add start
        Else
            ' ( 削除処理 )
            .SQL = "DELETE FROM 材料購入単価マスタ "
            .SQL = .SQL & " WHERE 材料管理番号 = '" & RTrim(H1lb_ZaiKan) & "'"
            .SQL = .SQL & "   And 取引先 = '" & RTrim(iH1_Torcd) & "'"
            .SQL = .SQL & "   And 契約先 = '" & RTrim(iH1_KTorcd) & "'"
            On Error GoTo DBDelete_Err
            .DBExec
            ' 材料購入単価内訳マスタは存在しないので削除処理無し
        End If
'-----------------------------------------------------------------------' 2009.10.01 add end
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
        iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")
        
        On Error GoTo 0
    End With
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
    Dim AddFlg  As Boolean

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        If SyoriKB = "ADD" Then
            .MB_MSG(2) = "     追加処理を行います      "
            .MB_Title = "追加処理"
        Else
            .MB_MSG(2) = "     修正処理を行います      "
            .MB_Title = "修正処理"
        End If
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
' ( 削除処理 )
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    On Error GoTo DBPut_Err

    If H0Op_Tanku(0).Value = True Then      ' 2009.10.01 add
        gSL_Select = "Select max(履歴NO) 履歴NO From 購入単価マスタ "
        gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   And 契約先 = '" & RTrim(iH1_KTorcd) & "'"
        Call KTMRead(gSL_Select, 1)
        gLong = KTM.履歴NO + 1
    '-----------------------------'
    '   購入単価マスタ出力
    '-----------------------------'
        Call ItemsClearKTM
        With KTM
            .品番 = RTrim(iH1_Hinbn)
            .取引先 = RTrim(iH1_Torcd)
            .契約先 = RTrim(iH1_KTorcd)
            .履歴NO = gLong
            If B1Ch_Kari.Value = 1 Then
               .仮単価区分 = "1"
            End If
            .単価 = CisFun.Val2(iB1_Tanka)
            .適用開始日付 = Trim(iB1_TSYmd)
            .適用終了日付 = Trim(iB1_TEYmd)
            .有無償区分 = B1Ch_UMu.Value
            .部品単価 = CisFun.Val2(iB1_BTanka)
            .加工単価 = CisFun.Val2(iB1_KTanka)
            .その他単価 = CisFun.Val2(iB1_STanka)
            .作成者 = gTanto
        End With
        ' ( 追加処理 )
        Call KTMInsert
    '-----------------------------'
    '   購入単価内訳マスタ出力
    '-----------------------------'
        Call ItemsClearKTU
        gCnt1 = 0
        For gInt = 0 To 5
            If RTrim(iB1_EtcMsg(gInt)) <> "" Then
                gCnt1 = gCnt1 + 1
                With KTU
                    .品番 = RTrim(iH1_Hinbn)
                    .取引先 = RTrim(iH1_Torcd)
                    .契約先 = RTrim(iH1_KTorcd)
                    .履歴NO = KTM.履歴NO
                    .連番 = gCnt1
                    .内訳内容 = RTrim(iB1_EtcMsg(gInt))
                    .単価 = CisFun.Val2(iB1_EtcTan(gInt))
                    .作成者 = gTanto
                End With
                Call KTUInsert
            End If
        Next
'-----------------------------------------------------------------------------------' 2009.10.01 add start
    Else
    '============================================
    '============= 部品区分：材料 ===============
    '============================================
        gSL_Select = "Select max(履歴NO) 履歴NO From 材料購入単価マスタ"
        gSL_Select = gSL_Select & " WHERE 材料管理番号 = " & CisFun.Val2(H1lb_ZaiKan)
        gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   And 契約先 = '" & RTrim(iH1_KTorcd) & "'"
        
        Call ZKTRead(gSL_Select, 1)
        gLong = ZKT.履歴NO + 1
        '-----------------------------'
        '   材料購入単価マスタ出力
        '-----------------------------'
        Call ItemsClearZKT
        With ZKT
            .材料管理番号 = CisFun.Val2(H1lb_ZaiKan)
            .取引先 = RTrim(iH1_Torcd)
            .契約先 = RTrim(iH1_KTorcd)
            .履歴NO = gLong
            .単価 = CisFun.Val2(iB1_Tanka)
            .仮単価区分 = "": If B1Ch_Kari.Value Then .仮単価区分 = "1"
            .適用開始日付 = Trim(iB1_TSYmd)
            .適用終了日付 = Trim(iB1_TEYmd)
            .有無償区分 = 0: If B1Ch_UMu.Value Then .有無償区分 = 1
            .部品単価 = 0
            .加工単価 = 0
            .その他単価 = 0
            .作成者 = gTanto
        End With
        ' ( 追加処理 )
        Call ZKTInsert
    End If
'-----------------------------------------------------------------------------------' 2009.10.01 add end

    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    If SyoriKB = "ADD" Then
        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
    Else
        iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
    End If
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------------------------------------------+
'+                                                          +
'+                                                          +
'+      グリッド処理(vsGrid)                                 +
'+                                                          +
'+                                                          +
'+----------------------------------------------------------+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設
'定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 6, 0, "履歴№")
        Call .FixedSet(中中, 右中, あり, 14, 1, "単　価")
        Call .FixedSet(中中, 中中, あり, 4, 2, "仮単")
        Call .FixedSet(中中, 中中, あり, 15, 3, "適用開始日")
        Call .FixedSet(中中, 中中, あり, 15, 4, "適用終了日")
        Call .FixedSet(中中, 中中, あり, 6, 5, "有無償")
        Call .FixedSet(中中, 中中, あり, 2, 6, "訳")
        Call .FixedSet(中中, 右中, あり, 13, 7, "部品単価")
        Call .FixedSet(中中, 右中, あり, 13, 8, "加工単価")
        Call .FixedSet(中中, 右中, あり, 13, 9, "その他単価")
        Call .FixedSet(中中, 中中, あり, 1, 10, "")
        Call .FixedSet(中中, 中中, あり, 12, 11, "その他内訳1")
        Call .FixedSet(中中, 右中, あり, 12, 12, "その他単価1")
        Call .FixedSet(中中, 中中, あり, 12, 13, "その他内訳2")
        Call .FixedSet(中中, 右中, あり, 12, 14, "その他単価2")
        Call .FixedSet(中中, 中中, あり, 12, 15, "その他内訳3")
        Call .FixedSet(中中, 右中, あり, 12, 16, "その他単価3")
        Call .FixedSet(中中, 中中, あり, 12, 17, "その他内訳4")
        Call .FixedSet(中中, 右中, あり, 12, 18, "その他単価4")
        Call .FixedSet(中中, 中中, あり, 12, 19, "その他内訳5")
        Call .FixedSet(中中, 右中, あり, 12, 20, "その他単価5")
        Call .FixedSet(中中, 中中, あり, 12, 21, "その他内訳6")
        Call .FixedSet(中中, 右中, あり, 12, 22, "その他単価6")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 23, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22")
        .InitDisp
    End With
    
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
    
    If Not KTMRead(gSL_Select) Then Exit Sub
' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
    B1lb_DispSu = Format(Ttl_Kensu, "#,###")
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(CisFun.Val2(ID))
        Next ID
        If Ttl_Kensu + 1 > .Rows Then
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not KTM_RDSTS
        'Grid Row,Col
        ID = ID + 1
        With vsGrid1
            .TextMatrix(ID, GridCol(0)) = KTM.履歴NO
            .TextMatrix(ID, GridCol(1)) = Format(KTM.単価, "#,##0.00")
            .TextMatrix(ID, GridCol(2)) = RTrim(KTM.仮単価区分)
            .TextMatrix(ID, GridCol(3)) = RTrim(KTM.適用開始日付)
            .TextMatrix(ID, GridCol(4)) = RTrim(KTM.適用終了日付)
'-------------------------------------------------------------------------------' 2009.10.01 upd start
            If KTM.有無償区分 = 0 Then
                .TextMatrix(ID, GridCol(5)) = "有償"
            Else
                .TextMatrix(ID, GridCol(5)) = "無償"
            End If
            If H0Op_Tanku(0).Value = True Then
                If KTM.部品単価 <> 0 Then
                    .TextMatrix(ID, GridCol(7)) = Format(KTM.部品単価, "#,##0.00")
                End If
                If KTM.加工単価 <> 0 Then
                    .TextMatrix(ID, GridCol(8)) = Format(KTM.加工単価, "#,##0.00")
                End If
                If KTM.その他単価 <> 0 Then
                    .TextMatrix(ID, GridCol(9)) = Format(KTM.その他単価, "#,##0.00")
                End If
            'その他単価内訳
                gInt = 0
                gSL_Select = "SELECT * FROM 購入単価内訳マスタ"
                gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
                gSL_Select = gSL_Select & "   And 取引先 = '" & RTrim(iH1_Torcd) & "'"
                gSL_Select = gSL_Select & "   And 契約先 = '" & RTrim(iH1_KTorcd) & "'"
                gSL_Select = gSL_Select & "   And 履歴NO = " & KTM.履歴NO & ""
                If KTURead(gSL_Select, 0, 1) Then
                    Do Until Not KTU_RDSTS
                        .TextMatrix(ID, GridCol(gInt + 11)) = RTrim(KTU.内訳内容)
                        If KTM.部品単価 <> 0 Then
                            .TextMatrix(ID, GridCol(gInt + 12)) = Format(KTU.単価, "#,##0.00")
                        End If
                    
                        gInt = gInt + 2
                        Call KTUReadNext(1)
                    Loop
                End If
                .TextMatrix(ID, GridCol(10)) = Format(gInt / 2, "##")
            End If
'-------------------------------------------------------------------------------' 2009.10.01 upd end
            Call KTUClose(1)
        End With
        Call KTMReadNext(0)
    Loop

    With vsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    Call KTMClose(0)
End Sub

Private Sub vsGrid1_DblClick()
    If ProcHB <> "B1" Then Exit Sub
    If RTrim(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid.FixedGet(0))) = "" Then Exit Sub
    
    'vsGridから複写
    With vsGrid1
        iB1_Tanka = .TextMatrix(.Row, CisVsGrid.FixedGet(1))
        If .TextMatrix(.Row, CisVsGrid.FixedGet(2)) = "" Then
            B1Ch_Kari.Value = 0
        Else
            B1Ch_Kari.Value = 1
        End If
        Call B1Ch_Kari_Click
        CisFun.DateE_AfVal = .TextMatrix(.Row, CisVsGrid.FixedGet(3))
        CisFun.DateEditUn
        iB1_TSYmd = CisFun.DateE_BefVal
        CisFun.DateE_AfVal = .TextMatrix(.Row, CisVsGrid.FixedGet(4))
        CisFun.DateEditUn
        iB1_TEYmd = CisFun.DateE_BefVal
'---------------------------------------------------------------------------' 2009.10.01 upd start
        If RTrim(.TextMatrix(.Row, CisVsGrid.FixedGet(5))) = "有償" Then
            B1Ch_UMu.Value = 0
        Else
            B1Ch_UMu.Value = 1
        End If
        Call B1Ch_UMu_Click
        If H0Op_Tanku(0).Value = True Then                                  ' 2009.10.01 add
            iB1_BTanka = .TextMatrix(.Row, CisVsGrid.FixedGet(7))
            iB1_KTanka = .TextMatrix(.Row, CisVsGrid.FixedGet(8))
            iB1_STanka = .TextMatrix(.Row, CisVsGrid.FixedGet(9))
            For gInt = 0 To 5
                iB1_EtcMsg(gInt) = .TextMatrix(.Row, CisVsGrid.FixedGet(gInt * 2 + 11))
                iB1_EtcTan(gInt) = .TextMatrix(.Row, CisVsGrid.FixedGet(gInt * 2 + 12))
            Next
        End If
'---------------------------------------------------------------------------' 2009.10.01 upd end
    End With
    iB1_Tanka.SetFocus
End Sub
'+----------------------+
'+  単価内訳コード獲得   +
'+----------------------+
Private Function EtcTanka_Naiyo() As Boolean
    EtcTanka_Naiyo = False
'
    Erase EtcMsg: Erase EtcTan: gCnt1 = 0
    gSL_Select = "SELECT 値,値名称,英数字1,数字1 FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = 'その他単価内訳'"
    gSL_Select = gSL_Select & " AND 値 <> ''"
    gSL_Select = gSL_Select & " ORDER BY 値"
    If SYMRead(gSL_Select, 0) Then
        Do Until Not SYM_RDSTS
            gCnt1 = gCnt1 + 1
            If CisFun.Val2(SYM.値) <= 6 Then
                EtcMsg(gCnt1) = SYM.値名称
                EtcTan(gCnt1) = SYM.数字1
            End If
            Call SYMReadNext
        Loop
        Call SYMClose
    Else
        Call SYMClose
        Call ItemsClearSYMCreate
        With SYMCreate
            .区分名称 = "その他単価内訳"
            .区分桁数 = 1
            .値 = "1"
            .制御区分 = ""
            .値名称 = "内訳１"
            .英数字1 = "数字(1)に単価入力"
            .作成者 = gTanto
            .検索名 = "その他単価内訳"
            .コード名 = "CD"
            .名称名 = "内 訳 内 容"
        End With
        Call SYMInitCreate(True)
        With SYMCreate
            .値 = "2"
            .値名称 = "内訳２"
            .検索名 = ""
        End With
        Call SYMInitCreate(False)
        With SYMCreate
            .値 = "3"
            .値名称 = "内訳３"
        End With
        Call SYMInitCreate(False)
        With SYMCreate
            .値 = "4"
            .値名称 = "内訳４"
        End With
        Call SYMInitCreate(False)
        With SYMCreate
            .値 = "5"
            .値名称 = "内訳５"
        End With
        Call SYMInitCreate(False)
        With SYMCreate
            .値 = "6"
            .値名称 = "内訳６"
        End With
        Call SYMInitCreate(False)
    
    End If
        
    EtcTanka_Naiyo = True
    Exit Function
EtcTanka_Naiyo_Err:
    CisFun.ErrorBox
    End
End Function

''+-------------------------------------+
''+      Debug用 SQL記録                +
''+-------------------------------------+
Private Sub Test_LogOutput(str As String, Msg As Integer)
'Exit Sub   '=========================================== 実行しないときはこの行を生かすこと =========================
    Dim dt01 As String
    Dim dt02 As String
    Dim f_No As Integer
    
    If Msg <> 0 Then MsgBox (str)
    dt01 = CStr(DateTime.Now)
    dt02 = Mid(dt01, 1, 4) & "." & _
           Mid(dt01, 6, 2) & "." & _
           Mid(dt01, 9, 2)

    f_No = FreeFile                             'ファイル番号の取得
    Open "D:\TestLog\SQL_" & dt02 & ".log" _
          For Append As #f_No                   'ファイルを追加モードで開く
    Print #f_No, "[" & Me.Name & "] " & dt01 & _
                 " ----------------------------------------" & Chr(13) & Chr(10) & _
                 str & Chr(13) & Chr(10)        'ファイルへ書き込む"
    Close #f_No                                 'ファイルを閉じる
End Sub

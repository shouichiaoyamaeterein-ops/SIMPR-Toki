VERSION 5.00
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CCT0040 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "テキストファイル設定"
   ClientHeight    =   11175
   ClientLeft      =   15
   ClientTop       =   375
   ClientWidth     =   15255
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
   ScaleHeight     =   11175
   ScaleWidth      =   15255
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D Grp_Back 
      Height          =   1980
      Left            =   330
      Top             =   4680
      Visible         =   0   'False
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   3493
      BackColor       =   16761024
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
      cBoderWidth     =   3
      Begin VB.ListBox Grp_List 
         BackColor       =   &H00C0FFFF&
         Height          =   1260
         Left            =   270
         Sorted          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   570
         Width           =   5700
      End
      Begin VB.PictureBox Picture1 
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
         Height          =   720
         Index           =   0
         Left            =   12000
         ScaleHeight     =   660
         ScaleWidth      =   1965
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   120
         Width           =   2025
         Begin CisBtn_60.CisBtn PB_Can2 
            Height          =   585
            Left            =   90
            Top             =   30
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
         Begin CisBtn_60.CisBtn PB_ENT2 
            Height          =   585
            Left            =   1020
            Top             =   15
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
      End
      Begin Cis3D_v60.CIS3D UC_3D7 
         Height          =   405
         Left            =   270
         Top             =   90
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   714
         BackColor       =   8438015
         Caption         =   "※  グループ設定  ※"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   15.75
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
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   825
         Left            =   6030
         Top             =   570
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   1455
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
         Begin Cis3D_v60.CIS3D UC_3D5 
            Height          =   375
            Left            =   600
            Top             =   30
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   661
            Caption         =   "グ   ル   ー   プ   名"
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
         Begin Cis3D_v60.CIS3D UC_3D4 
            Height          =   375
            Left            =   30
            Top             =   30
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   661
            Caption         =   "GRP"
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
         Begin CisText_V60.CisText iB2_GrpNm 
            Height          =   375
            Left            =   600
            TabIndex        =   3
            Top             =   420
            Width           =   5025
            _ExtentX        =   8864
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
         Begin CisText_V60.CisText iB2_Grp 
            Height          =   375
            Left            =   30
            TabIndex        =   2
            Top             =   420
            Width           =   555
            _ExtentX        =   979
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
      End
   End
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
      Height          =   795
      Left            =   9540
      ScaleHeight     =   735
      ScaleWidth      =   5235
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1335
      Width           =   5295
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   915
         Top             =   60
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   2625
         Top             =   60
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
         Left            =   3480
         Top             =   60
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
         Left            =   4335
         Top             =   60
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
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   60
         Top             =   60
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
      Begin CisBtn_60.CisBtn PB_Set 
         Height          =   585
         Left            =   1770
         Top             =   60
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
         BtnType         =   15
         BtnCaption      =   "設定"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   0
         BtnShadowColor  =   16777215
         PFCaption       =   "( F11 )"
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
         PFBackColor     =   0
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
      Height          =   1035
      Left            =   420
      ScaleHeight     =   975
      ScaleWidth      =   1815
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1230
      Width           =   1875
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   870
         Left            =   105
         Top             =   75
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1535
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
         Begin CisText_V60.CisText iH1_Gyomu 
            Height          =   375
            Left            =   150
            TabIndex        =   1
            Top             =   390
            Width           =   1380
            _ExtentX        =   2434
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
            Text            =   "XXXXXXXXXX"
            MaxLength       =   10
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   870
         Left            =   105
         Top             =   75
         Visible         =   0   'False
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1535
         ForeColor       =   16711680
         Caption         =   "Fｺｰﾄﾞ"
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
         Begin CisText_V60.CisText iH1_FCode 
            Height          =   375
            Left            =   165
            TabIndex        =   0
            Top             =   390
            Width           =   600
            _ExtentX        =   1058
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
            Text            =   "XXXX"
            MaxLength       =   4
         End
      End
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
      Height          =   7635
      Left            =   405
      ScaleHeight     =   7575
      ScaleWidth      =   14355
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2580
      Width           =   14415
      Begin vsFlexLib.vsFlexArray VSGrid1 
         DragIcon        =   "CCT0040.frx":0000
         Height          =   7365
         Left            =   90
         TabIndex        =   5
         Top             =   120
         Width           =   14160
         _Version        =   196608
         _ExtentX        =   24977
         _ExtentY        =   12991
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   $"CCT0040.frx":0442
         Rows            =   21
         Cols            =   14
         BackColor       =   12648447
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   330
         FillStyle       =   1
         SelectionMode   =   1
         BackColorAlternate=   12648447
      End
   End
   Begin VB.Timer Timer1 
      Left            =   2640
      Top             =   270
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
      Height          =   285
      IMEMode         =   3  'ｵﾌ固定
      Left            =   0
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   480
      Width           =   225
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
      Caption         =   "【 テキストファイル設定  】"
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
         Top             =   90
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
         Left            =   11880
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
   End
   Begin Cis3D_v60.CIS3D Del_Msg 
      Height          =   615
      Left            =   4860
      Top             =   1365
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   1085
      BackColor       =   255
      ForeColor       =   16777215
      Caption         =   "削 除"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   18
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cBoderWidth     =   3
      cAlingnment     =   7
   End
End
Attribute VB_Name = "CCT0040"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   TEXTﾌｧｲﾙ変換設定
'**       フォームID    :   CCT0040
'**       処理概要      :
'**
'**       作  成  日    :   2003/12/12  By CIS
'**       変  更  日    :   2004/11/08  By CIS 財務ﾃﾞｰﾀ対応(売上照合)
'**       変  更  日    :   2007/03/05  By CIS 手配データ追加
'**       変  更  日    :   2007/03/12  By CIS 管理データ未登録時、エラー表示なしに変更
'**                                            初期管理データ登録なしに変更
'**       変  更  日    :   2008/05/12  By CIS レイアウト情報初期登録処理追加
'**
'**********************************************************'

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
   
    Dim mDelColor       As Long         ' 削除時、ｸﾞﾘｯﾄﾞ色
    Dim mUnDelColor     As Long         ' 削除、設定以外
    Dim mSetColor       As Long         ' ｸﾞﾙｰﾌﾟ設定時、ｸﾞﾘｯﾄﾞ色(BackColor)
    Dim mGrpColor       As Long         ' ｸﾞﾙｰﾌﾟ設定時、ｸﾞﾘｯﾄﾞ色

    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   # 二 重 起 動 防 止  #
'   #-------------------#
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
    
    Grp_Back.Move 390, 570
' グリッド色設定
    mDelColor = &HC0C0FF
    mSetColor = &HC0FFC0
    mGrpColor = &HC0FFC0    ' &HFFC0FF
    mUnDelColor = VSGrid1.BackColor
    
    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
'   #------------------#
'   #  ｸﾞﾘｯﾄﾞ初期処理   #
'   #------------------#
    With CisVsGrid
        Set .GridObj = VSGrid1              ' ｸﾞﾘｯﾄﾞｵﾌﾞｼﾞｪｸﾄ名 ｾｯﾄ
        .FixedHeight = 380                  ' 固定行(見出し)高さ設定(省略時は300)
        .PatanMax = 1                       ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ最大数
        .PatanInit = 0                      ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ初期値
        .SelectCol = 12                     ' セル選択処理ﾜｰｸ列
        .NoSpaceCol = 1
        .NumberSet = NoSpace
'        Set .SelectDispObj = B1lb_SKensu    ' セル選択件数表示ｵﾌﾞｼﾞｪｸﾄｾｯﾄ
' ﾌﾟﾛｸﾞﾗﾑ作成時、以下のﾌﾗｸﾞをｾｯﾄし、ｸﾘｯﾌﾟﾎﾞｰﾄﾞよりｺｰﾄﾞの貼り付けを行う。(GridInit)
        .InitGet = False                    ' ｸﾞﾘｯﾄﾞ内容初期設定 ｸﾘｯﾌﾟﾎﾞｰﾄﾞ作成ﾌﾗｸﾞ

        .Init                               ' ｸﾞﾘｯﾄﾞ初期処理 実行
    End With
    Call GridInit                           ' ｸﾞﾘｯﾄﾞ初期表示
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/12 INSERT START
    Call MeisyoInit
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/12 INSERT END
    
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
'
    Call InitAdd
    
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
           Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click    '【新規】
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF11:      If PB_Set.Visible Then Call PB_Set_Click    '【設定】
           Case vbKeyEscape
                If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
                If PB_Can2.Visible Then Call PB_CAN2_Click  '【取消(GRP)】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      '【入力】
                                  Call PB_ENT_Click
                               End If
                               If PB_ENT2.Visible Then                     '【入力(GRP)】
                                  Call PB_ENT2_Click
                               End If
           Case Else:          Exit Function
    End Select
    Key_Acc = True
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    Call CisDB.DBDISConnect
    Call CisFrm.UnLoadDisp
    End
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
   Call CisFrm.MousePT(11)
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DelColor(True)
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      GoTo PB_CAN_Ed
   End If
   If ProcHB = "S1" Then
      ProcHB = "B1"
      Call DispChange(ProcHB)
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
End Sub
'+------------------------------------+
'+    入  力  キ ー(Enter)  設定更新   +
'+------------------------------------+
Private Sub PB_ENT2_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call SetDBPut
    Dummy.Enabled = False
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN2_Click()
    ProcHB = "B1"
    Call DispChange(ProcHB)
    DoEvents
    Call SetColor(True)
End Sub
'+----------------------------+
'+    新  規  キ ー(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    
    CT0040_Proc = "ADD"
    
    CCT0041.Show vbModal
    Unload CCT0041
    Set CCT0041 = Nothing
'    CCT0045.Show vbModal
'    Unload CCT0045
'    Set CCT0045 = Nothing
'
    If ProcHB = "B1" And CT0040_RTcd Then
       Call GridSet
    End If
End Sub
'+----------------------------+
'+    削　除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    Call DelColor(Del_Msg.Visible)
End Sub
'+----------------------------+
'+      削除色設定             +
'+----------------------------+
Private Sub DelColor(DelSet As Boolean)
    If DelSet Then
       Del_Msg.Visible = False
       VSGrid1.BackColor = mUnDelColor
       Call CisVsGrid.vsColor(消去全部)
    Else
       Del_Msg.Visible = True
       Call CisVsGrid.vsColor(消去全部, mDelColor)
    End If
End Sub
'+----------------------------+
'+    連続実行GRP設定(F11)     +
'+----------------------------+
Private Sub PB_Set_Click()
    If Del_Msg.Visible Then
       Call DelColor(True)
    End If
    ProcHB = "S1"
    Call DispChange(ProcHB)
    If Grp_Back.Visible Then
       Call SetColor(False)
    End If
' ｸﾞﾙｰﾌﾟ表示
    Call GrpGet
    
    iB2_Grp = ""
    iB2_GrpNm = ""
    
End Sub
'+----------------------------+
'+      設定色設定             +
'+----------------------------+
Private Sub SetColor(SetSet As Boolean)
    If SetSet Then
       VSGrid1.BackColor = mUnDelColor
       Call CisVsGrid.vsColor(消去全部, mUnDelColor)
    Else
       Call CisVsGrid.vsColor(消去全部, mSetColor)
    End If
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
       Case "B1", "S1"
          B1Mode = True
          B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    PB_CAN.Visible = Not H1Mode
    
    PB_ADD.Visible = Not H1Mode
    PB_DEL.Visible = Not H1Mode
    PB_Set.Visible = Not H1Mode
    
    Grp_Back.Visible = False
    PNL_PF.Visible = True
    If Pro = "S1" Then
       Grp_Back.Visible = True
       PNL_PF.Visible = False
    End If
' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        If Not GridSet Then
'-------------------------------------------------------------------- 2007/03/12 Delete
'           With CisFun
'                .MB_Lines = 3
'                .MB_MSG(2) = "  表示データが存在しません。  "
'                .MB_Button = Error
'                .MBOX
'           End With
'           Exit Sub
'-------------------------------------------------------------------- 2007/03/12 Delete End
        End If
            
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
       Call CallRtn
    End If
    If ProcHB = "B1" And Del_Msg.Visible Then
       If Not DBDelete Then GoTo ReturnPress_Ed
       If Not GridSet Then
          Call PB_CAN_Click
          Exit Sub
       End If
       Call PB_DEL_Click
    End If
ReturnPress_Ed:
End Sub
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
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj
    If HeadBody = "B" Then
        Call CisVsGrid.Clear
    End If
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    Body1Chk = True
End Function
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+      削 除 処 理      +
'+----------------------+
Private Function DBDelete()
    Dim wDelChk As Boolean
    
    DBDelete = False

' 削除選択確認
    If CisVsGrid.SelectCount = 0 Then
        With CisFun
            .MB_Lines = 3
            .MB_MSG(2) = "    削除する行を選択して下さい。    "
            .MB_Title = "削除確認"
            .MB_Button = OK
            .MBOX
        End With
        Exit Function
    End If

'  ( 確認メッセージ )
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "    選択されている内容を削除します。    "
        .MB_MSG(4) = "         よろしいですか？              "
        .MB_Title = "削除確認"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    With VSGrid1
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(1))) = "" Then Exit For
            
            If Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(CisVsGrid.SelectCol))) = "X" Then
        ' ( 削除処理 )
                gStr = Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(1)))
                With CisDB
                    .SQL = "DELETE FROM TEXT管理テーブル "
                    .SQL = .SQL & " Where FILECD = '" & gStr & "'"
                    On Error GoTo DBDelete_Err
                    Call .DBExec
                End With
            End If
        Next gLong
    End With
    On Error GoTo 0
'--------------------------------------------------------------------- 2007/03/12 Delete
'    gSL_Select = "SELECT COUNT(*) AS 件数 FROM TEXT管理テーブル "
'    Call TKTRead(gSL_Select, 1)
'    If TKT.件数 = 0 Then
''  ( 確認メッセージ )
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(1) = "    TNS管理情報がすべて削除されました。    "
'            .MB_MSG(3) = "        初期設定内容を登録します。         "
'            .MB_Title = "初期設定確認"
'            .MB_Button = OK
'            .MBOX
'        End With
'        Call TEXTInitAdd
'    End If
'--------------------------------------------------------------------- 2007/03/12 Delete End
    
    Call CisDB.DBTran(TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    DBDelete = True
    Exit Function
DBDelete_Err:
    Call CisDB.DBTran(TransRollback) ' ﾄﾗﾝｻﾞｸｼｮﾝ ﾛｰﾙﾊﾞｯｸ
    CisFun.ErrorBox
    End
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
        Call .FixedSet(中中, 右中, あり, 0, 0, "№")
        Call .FixedSet(中中, 左中, あり, 10, 1, "ﾌｧｲﾙｺｰﾄﾞ")
        Call .FixedSet(中中, 左中, あり, 30, 2, "ファイル名称")
        Call .FixedSet(中中, 右中, あり, 0, 3, "有効長")
        Call .FixedSet(中中, 左中, あり, 16, 4, "変換元ﾌｧｲﾙ名")
        Call .FixedSet(中中, 左中, あり, 30, 5, "TEXTレイアウト")
        Call .FixedSet(中中, 中中, あり, 0, 6, "改行")
        Call .FixedSet(中中, 中中, あり, 0, 7, "連続")
        Call .FixedSet(中中, 中中, あり, 0, 8, "GRP")
        Call .FixedSet(中中, 中中, あり, 16, 9, "処 理 日 時")
        Call .FixedSet(中中, 右中, あり, 0, 10, "処理件数")
        Call .FixedSet(中中, 右中, あり, 0, 11, "TEXTタイプ")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 12, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Function GridSet() As Boolean
    Dim ID          As Long
    Dim Ttl_Kensu   As Integer
    Dim wSl_Where   As String

    GridSet = False
    
    Call CisVsGrid.Clear
' 抽出処理 初期ｾｯﾄ
    Call CisFrm.MousePT(11)   ' ﾏｳｽﾎﾟｲﾝﾀ (砂時計)
'
    PNL_PF.Enabled = False      ' ﾌｧﾝｸｼｮﾝ使用禁止
    
    gSL_Select = "Select TK.*,"
    '-------------------------------------- 2007/03/05 Delete
'    gSL_Select = gSL_Select & " Case WHEN TK.TNSタイプ = 1 THEN RTRIM(SY2.値名称) "
'    gSL_Select = gSL_Select & "      ELSE RTRIM(SY.値名称) END レイアウト名   "
    '-------------------------------------- 2007/03/05 Delete
    
    '-------------------------------------- 2007/03/05 Insert
    gSL_Select = gSL_Select & " Case WHEN TK.TNSタイプ = 1 THEN RTRIM(SY2.値名称) "
    gSL_Select = gSL_Select & "      WHEN TK.TNSタイプ = 2 THEN RTRIM(SY3.値名称) "
    gSL_Select = gSL_Select & "      ELSE RTRIM(SY.値名称) END レイアウト名   "
    '-------------------------------------- 2007/03/05 Insert End
    gSL_Select = gSL_Select & " from TEXT管理テーブル TK"
    gSL_Select = gSL_Select & " Left Outer Join 名称マスタ SY"
    gSL_Select = gSL_Select & "   On  SY.区分名称 = 'レイアウト情報内示'"
    gSL_Select = gSL_Select & "   And SY.値 = TK.TEXTレイアウト"
    gSL_Select = gSL_Select & "   And SY.区分タイプ = ''"
    gSL_Select = gSL_Select & " Left Outer Join 名称マスタ SY2"
    gSL_Select = gSL_Select & "   On  SY2.区分名称 = 'レイアウト情報財務'"
    gSL_Select = gSL_Select & "   And SY2.値 = TK.TEXTレイアウト"
    gSL_Select = gSL_Select & "   And SY2.区分タイプ = ''"
    '--------------------------------------- 2007/03/05 Insert
    gSL_Select = gSL_Select & " Left Outer Join 名称マスタ SY3"
    gSL_Select = gSL_Select & "   On  SY3.区分名称 = 'レイアウト情報手配'"
    gSL_Select = gSL_Select & "   And SY3.値 = TK.TEXTレイアウト"
    gSL_Select = gSL_Select & "   And SY3.区分タイプ = ''"
    '--------------------------------------- 2007/03/05 Insert End
    
    wSl_Where = ""
    With CisFun
        .Where_CreateKB = NewDocument
        .Where_Create AlphaMe, "FCODE", iH1_FCode
        .Where_Create AlphaMe, "FILECD", iH1_Gyomu
'        .Where_Create AlphaMe, "発信元", "", 等しくない, Create
        wSl_Where = .Where_Phrase
    End With
    
    
    wSl_Where = wSl_Where & " ORDER BY FILECD"
    
    If Not TKTRead(gSL_Select & wSl_Where) Then GoTo GridSet_Ed
    
    Ttl_Kensu = CisDB.RecordCount      ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VSGrid1
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 1
    Do Until Not TKT_RDSTS
        Call ToGrid(ID)
        
        ID = ID + 1
        
        Call TKTReadNext
    Loop
    With VSGrid1
        .Row = 1: .Col = 1
        .Redraw = True
    End With
    Call CisVsGrid.VSSort
    GridSet = True
GridSet_Ed:
    Call TKTClose
    
    PNL_PF.Enabled = True      ' ﾌｧﾝｸｼｮﾝ使用禁止 解除
    Call CisFrm.MousePT(1)         ' ﾏｳｽﾎﾟｲﾝﾀ (砂時計)解除
'
End Function
'+-----------------------------------------+
'+    WIPS管理テーブルよりｸﾞﾘｯﾄﾞへ内容ｾｯﾄ     +
'+-----------------------------------------+
Private Sub ToGrid(lRow As Long)
    Dim ColId       As Long
    Dim GridCol()   As Long
    
    With VSGrid1
        ReDim GridCol(.Cols)
        For ColId = 0 To .Cols - 1
            GridCol(ColId) = CisVsGrid.FixedGet(ColId)
        Next ColId
        
        .TextMatrix(lRow, GridCol(0)) = Format(lRow, "###")
                
        .TextMatrix(lRow, GridCol(1)) = TKT.FILECD
        .TextMatrix(lRow, GridCol(2)) = TKT.F名称
        .TextMatrix(lRow, GridCol(3)) = Format(TKT.有効長, "###")
        .TextMatrix(lRow, GridCol(4)) = TKT.変換元ファイル名
        .TextMatrix(lRow, GridCol(5)) = TKT.レイアウト名
        If TKT.オプション1 = 1 Then
           .TextMatrix(lRow, GridCol(6)) = "○"
        Else
           .TextMatrix(lRow, GridCol(6)) = ""
        End If
        
        If TKT.連続実行 = 1 Then
           .TextMatrix(lRow, GridCol(7)) = "○"
        Else
           .TextMatrix(lRow, GridCol(7)) = ""
        End If
        .TextMatrix(lRow, GridCol(8)) = TKT.実行GRP
        
        If Format(TKT.最終処理日時, "YYYYMMDD") > "20000101" Then
           .TextMatrix(lRow, GridCol(9)) = Format(TKT.最終処理日時, "YYYY/MM/DD hh:ss")
           .TextMatrix(lRow, GridCol(10)) = Format(TKT.処理件数, "#,###")
        End If
        .TextMatrix(lRow, GridCol(11)) = TKT.TNSタイプ
    
    End With
End Sub
'******************************************'
'*****        行選択（クリック）       *****
'******************************************'
Private Sub VSGrid1_Click()
    If ProcHB = "S1" Then
'       If Trim(iB2_Grp) <> "" Then
          Call CisVsGrid.vsColor(個別, mGrpColor)
'       End If
       Exit Sub
    End If

    If Not Del_Msg.Visible Then Exit Sub
    Call CisVsGrid.vsColor(個別, mDelColor)
End Sub
'******************************************'
'*****     行選択（ダブルクリック）     *****
'******************************************'
Private Sub VSGrid1_DblClick()
    Call CallRtn
End Sub
'******************************************'
'*****        行選択（スペースｷｰ)      *****
'******************************************'
Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    Call VSGrid1_Click
'    If Not Del_Msg.Visible Then Exit Sub
'    Call CisVsGrid.vsColor(消去全部, mDelColor)
End Sub
'******************************************'
'*****       移動選択項目選択          *****
'******************************************'
Private Sub VSGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Button : 1 = 左ﾎﾞﾀﾝ ,2 = 右ﾎﾞﾀﾝ
    With VSGrid1
        .Tag = ""
        If Button = 2 Then
            .Row = .MouseRow
            If .MouseRow = 0 Then
                .Col = 1: .ColSel = .Cols - 1
                Exit Sub
            End If
            .Col = 1: .ColSel = .Cols - 1
            Exit Sub
        End If
        If Button <> 1 Then Exit Sub
    ' 固定行以外は行わない
        If .MouseRow <> 0 Then Exit Sub
    ' 固定列は行わない
        If .MouseCol = 0 Then Exit Sub
    ' 移動元列をタグへ退避
        .Tag = Str(.MouseCol)
        .Drag 1
    End With
End Sub
'******************************************'
'*****        表 示 項 目 移 動        *****
'******************************************'
Private Sub VSGrid1_DragDrop(Source As Control, X As Single, Y As Single)
    With VSGrid1
    ' 固定列は行わない
        If .MouseCol = 0 Then Exit Sub
        If .Tag = "" Then Exit Sub
' ｸﾞﾘｯﾄﾞ表示 禁止
        .Redraw = False
' 列移動
        .ColPosition(CisFun.Val2(.Tag)) = .MouseCol
        .Col = 1: .ColSel = .Cols - 1
' ｸﾞﾘｯﾄﾞ表示 解除
        .Redraw = True
'
        Call CisVsGrid.VSSort
    End With
End Sub
'+--------------------------------+
'+      詳細設定画面呼出           +
'+--------------------------------+
Private Sub CallRtn()

    If Del_Msg.Visible Then Exit Sub
    If ProcHB = "S1" Then Exit Sub
    
    CT0040_FileCd = Trim(VSGrid1.TextMatrix(VSGrid1.Row, CisVsGrid.FixedGet(1)))
    CT0040_TEXTTYPE = CisFun.Val2(VSGrid1.TextMatrix(VSGrid1.Row, CisVsGrid.FixedGet(11)))
'
    If CT0040_FileCd = "" Then
       Call PB_ADD_Click
       Exit Sub
    End If
    CT0040_Proc = "MNT"
    
    If CT0040_TEXTTYPE = 0 Then
        CCT0045.Show vbModal
        Unload CCT0045
        Set CCT0045 = Nothing
    '
        Call CisVsGrid.vsColor(消去個別)
        If CT0040_RTcd Then
           Call ToGrid(VSGrid1.Row)
        End If
    End If
'+---------------------------------- 2007/03/05 Delete
    'If CT0040_TEXTTYPE = 1 Then
'+---------------------------------- 2007/03/05 Delete End
'+---------------------------------- 2007/03/05 Insert
    If CT0040_TEXTTYPE = 1 Or CT0040_TEXTTYPE = 2 Then
'+---------------------------------- 2007/03/05 Insert End
        CCT0046.Show vbModal
        Unload CCT0046
        Set CCT0046 = Nothing
    '
        Call CisVsGrid.vsColor(消去個別)
        If CT0040_RTcd Then
           Call ToGrid(VSGrid1.Row)
        End If
    End If
    
'    CT0040_FileCd = Trim(VSGrid1.TextMatrix(VSGrid1.Row, CisVsGrid.FixedGet(1)))
''
'    CT0040_Proc = "MNT"
'    CCT0045.Show vbModal
'    Unload CCT0045
'    Set CCT0045 = Nothing
''
'    Call CisVsGrid.vsColor(消去個別)
'    If CT0040_RTcd Then
'       Call ToGrid(VSGrid1.Row)
'    End If
End Sub
'+--------------------------------+
'+  データなし時、基本情報出力      +
'+--------------------------------+
Private Sub InitAdd()
    Call TEXTInitAdd

    ' 対象年月初期表示値を求める
    gSL_Select = "Select 英数字2 From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = '内示年月'"
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値 = '1'"
    If SYMRead(gSL_Select, 1) Then
        CT0040_Naiji = Mid(SYM.英数字2, 1, 2)
    Else
        CT0040_Naiji = "定常"
    End If
End Sub
'+-----------------------------------------------------+
'+-----------------------------------------------------+
'+          グループ設定
'+-----------------------------------------------------+
'+-----------------------------------------------------+
'+---------------------+
'+    GRP(iB2_Grp)     +
'+---------------------+
Private Sub iB2_Grp_LostFocus()
    Call GrpItemGet(iB2_Grp, gStr)
    iB2_GrpNm = gStr
    Call GrpGridSelect(iB2_Grp)
End Sub
'+--------------------------+
'+    GRP_List(Grp_List)    +
'+--------------------------+
Private Sub Grp_List_LostFocus()
    Call GrpGet
End Sub
Private Sub Grp_List_DblClick()
    gStr = Mid(Grp_List.List(Grp_List.ListIndex), 1, 2)
    If Trim(gStr) <> "" Then
        iB2_Grp = gStr
        Call GrpItemGet(iB2_Grp, gStr)
        iB2_GrpNm = gStr
        iB2_Grp.SetFocus
        
        Call GrpGridSelect(iB2_Grp)
    End If
End Sub
'+-------------------------------------+
'+    GRP(2桁)よりグループ名を求める     +
'+-------------------------------------+
Private Sub GrpItemGet(cItem As String, cData As String)
    Static GrpInt   As Integer
    cData = ""
    
    For GrpInt = 0 To Grp_List.ListCount - 1
        If Trim(cItem) = Trim(Mid(Grp_List.List(GrpInt), 1, 2)) Then
           cData = Trim(Mid(Grp_List.List(GrpInt), 3, Len(Grp_List.List(GrpInt)) - 2))
           Exit For
        End If
    Next GrpInt
End Sub
'******************************************'
'*****      対象グループ色設定         *****
'******************************************'
Private Sub GrpGridSelect(LGrp As String)
    Static GrpInt   As Integer
    Static sRow     As Long
    
    With VSGrid1
        .Redraw = False
        sRow = .Row
        For GrpInt = 1 To .Rows - 1
            .Row = GrpInt
            If Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(1))) <> "" Then
                .Col = CisVsGrid.FixedGet(8)    ' GRP
                If Trim(LGrp) = Trim(.Text) Then
                   Call CisVsGrid.vsColor(個別, mGrpColor)
                Else
                   Call CisVsGrid.vsColor(消去個別, mSetColor)
                End If
            End If
        Next GrpInt
        .Row = sRow: .Col = 1
        .Redraw = True
    End With
End Sub
'******************************************'
'*****      グループ表示(設定済み)      *****
'******************************************'
Private Sub GrpGet()
    Static GrpInt   As Integer
    Static Grp      As String
    Static GrpName  As String
    
    Grp_List.Clear
  
    gSL_Select = "Select * From Irnaksic"
    gSL_Select = gSL_Select & " Where bk = 'Group'"
    gSL_Select = gSL_Select & " Order By rno "

    If Not IRNRead(gSL_Select) Then
       Call IRNClose
       Exit Sub
    End If
    
    Do Until Not IRN_RDSTS
        With IRN
            Grp_List.AddItem RTrim(.oyian) & " " & RTrim(.oyian2)
        End With
            
        Call IRNReadNext
    Loop
    
    Call IRNClose
    
End Sub
'+----------------------------+
'+    設定更新(GRP)            +
'+----------------------------+
Private Sub SetDBPut()
    Static GrpInt   As Integer
    
    VSGrid1.SetFocus
    ''
    With CisFun
        .MB_Lines = 3
        .MB_MSG(2) = "    グループ設定を更新します       "
        .MB_Title = "ｸﾞﾙｰﾌﾟ設定"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Sub
    End With
    Static wGrpSet As Boolean
    
' ｸﾞﾙｰﾌﾟ名更新
    wGrpSet = False
    If Grp_List.ListCount > 0 Then
       For GrpInt = 0 To Grp_List.ListCount - 1
           If Trim(iB2_Grp) = Trim(Mid(Grp_List.List(GrpInt), 1, 2)) Then
              If Trim(iB2_GrpNm) <> "" Then
                 Grp_List.List(GrpInt) = iB2_Grp & " " & Trim(iB2_GrpNm)
              Else
                 Grp_List.RemoveItem (GrpInt)
              End If
              wGrpSet = True
              Exit For
           End If
       Next GrpInt
    End If
    If Not wGrpSet Then
       If Trim(iB2_Grp) <> "" And Trim(iB2_GrpNm) <> "" Then
          Grp_List.AddItem iB2_Grp & " " & Trim(iB2_GrpNm)
       End If
    End If
 'ファイル更新
    GrpInt = 0
    With CisDB
        .SQL = "DELETE From Irnaksic"
        .SQL = .SQL & " Where bk = 'Group'"
        Call .DBExec
    End With
    If Grp_List.ListCount > 0 Then

       For GrpInt = 0 To Grp_List.ListCount - 1
           With IRN
               Call ItemsClearIRN
               .bk = "Group"
               .rno = GrpInt
               .oyian = Trim(Mid(Grp_List.List(GrpInt), 1, 2))
               .oyian2 = Trim(Mid(Grp_List.List(GrpInt), 4, 40))
           End With
           Call IRNInsert

       Next GrpInt
    End If

    wGrpSet = False

    With VSGrid1
        .Redraw = False
        For GrpInt = 1 To .Rows - 1
            .Row = GrpInt
            .Col = CisVsGrid.FixedGet(1)
            If Trim(.Text) = "" Then Exit For
            .Col = CisVsGrid.FixedGet(CisVsGrid.SelectCol)
            If Trim(.Text) = "X" Then
               .TextMatrix(.Row, CisVsGrid.FixedGet(8)) = iB2_Grp
               .TextMatrix(.Row, CisVsGrid.FixedGet(7)) = "○"
               wGrpSet = True
            Else
                If Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(8))) = iB2_Grp Then
                  .TextMatrix(.Row, CisVsGrid.FixedGet(8)) = ""
                  .TextMatrix(.Row, CisVsGrid.FixedGet(7)) = ""
                  wGrpSet = True
               End If
            End If

            If wGrpSet Then
                gStr = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(1)))
                CisDB.SQL = " UPDATE TEXT管理テーブル SET"
                CisDB.SQL = CisDB.SQL & " 実行GRP = '" & Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(8))) & "',"
                If Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(7))) = "○" Then
                   CisDB.SQL = CisDB.SQL & " 連続実行 = 1"
                Else
                   CisDB.SQL = CisDB.SQL & " 連続実行 = 0"
                End If

                CisDB.SQL = CisDB.SQL & " WHERE FILECD = '" & gStr & "'"
                Call CisDB.DBExec
            End If

        Next GrpInt
        .Row = 1
        .Redraw = True
    End With
End Sub

'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/12 INSERT START
'+---------------------------------------+
'+      名称マスタ初期設定
'+---------------------------------------+
Private Sub MeisyoInit()
' レイアウト情報内示
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "レイアウト情報内示"
        .区分桁数 = 1
        .値 = "0"
        .英数字1 = ""
        .値名称 = "部品納入内示表(トヨタ形式)"
        .制御区分 = "X"
        .作成者 = gTanto
        If Not SYMInitCreate(True) Then
            For gInt = 1 To 4
                .値 = CStr(gInt)
                Select Case gInt
                       Case 1
                             .値名称 = "納入内示ﾃﾞｰﾀB(トヨタ紡織形式)"
                       Case 2
                             .値名称 = "補給内示ﾃﾞｰﾀ(トヨタ紡織形式)"
                       Case 3
                             .値名称 = "納入内示ﾃﾞｰﾀ(新ﾚｲｱｳﾄ形式)"
                       Case 4
                             .値 = "6"
                             .値名称 = "内示データ(車体精工ﾚｲｱｳﾄ形式)"
                End Select
                Call SYMInitCreate
            Next gInt
        End If
    End With
' レイアウト情報買掛
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "レイアウト情報財務"
        .区分桁数 = 1
        .値 = "0"
        .値名称 = "財務データ(トヨタ形式)"
        .英数字1 = "G530"
        .制御区分 = "X"
        .作成者 = gTanto
        If Not SYMInitCreate(True) Then
            For gInt = 1 To 7
                .値 = CStr(gInt)
                Select Case gInt
                       Case 1
                            .値名称 = "検収データ(ﾄﾖﾀ紡織:ｶｰﾄﾞ)"
                            .英数字1 = "060"
                       Case 2
                            .値名称 = "検収データ(ﾄﾖﾀ紡織:伝票)"
                            .英数字1 = "070"
                       Case 3
                            .値名称 = "検収データ(ﾄﾖﾀ紡織:有償支給品)"
                            .英数字1 = "080"
                       Case 4
                            .値名称 = "検収データ(ﾄﾖﾀ紡織:型支払分)"
                            .英数字1 = "090"
                       Case 5
                            .値名称 = "月間検収ﾃﾞｰﾀ(トヨタ車体形式)"
                            .英数字1 = "701"
                       Case 6
                            .値名称 = "検収データ(車体精工:購入)"
                            .英数字1 = "100"
                       Case 7
                            .値名称 = "検収データ(車体精工:支給)"
                            .英数字1 = "110"
                End Select
                Call SYMInitCreate
            Next gInt
        End If
    End With
End Sub
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/05/12 INSERT END


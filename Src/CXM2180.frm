VERSION 5.00
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXM2180 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "購入単価マスタ 一括入力"
   ClientHeight    =   11310
   ClientLeft      =   1125
   ClientTop       =   1545
   ClientWidth     =   15510
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11310
   ScaleWidth      =   15510
   Begin Cis3D_v60.CIS3D Back_ZTanka 
      Height          =   255
      Left            =   11040
      Top             =   3630
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   450
      BackColor       =   14737632
      Caption         =   "材料単価"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   14737632
      cBoderColor2    =   14737632
      cFont3DColor1   =   8421504
      cFont3DColor2   =   -2147483643
      cFont3D         =   2
      cAlingnment     =   7
      cBoderStyle     =   2
   End
   Begin VB.PictureBox H0_Area1 
      Height          =   930
      Left            =   15
      ScaleHeight     =   870
      ScaleWidth      =   1965
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   630
      Width           =   2025
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   765
         Left            =   60
         Top             =   60
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   1349
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
         cPositionY      =   30
         Begin VB.OptionButton H0Op_Tanku 
            BackColor       =   &H00C0FFC0&
            Caption         =   "材料"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   930
            TabIndex        =   0
            Top             =   390
            Width           =   750
         End
         Begin VB.OptionButton H0Op_Tanku 
            BackColor       =   &H00C0C0C0&
            Caption         =   "部品"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   150
            TabIndex        =   1
            Top             =   390
            Width           =   750
         End
      End
   End
   Begin Cis3D_v60.CIS3D Back_ZHinbn 
      Height          =   255
      Left            =   540
      Top             =   3630
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   450
      BackColor       =   14737632
      Caption         =   "材料/寸法"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBoderColor1    =   14737632
      cBoderColor2    =   14737632
      cFont3DColor1   =   8421504
      cFont3DColor2   =   -2147483643
      cFont3D         =   2
      cAlingnment     =   7
      cBoderStyle     =   2
   End
   Begin Cis3D_v60.CIS3D Message 
      Height          =   870
      Left            =   4065
      Top             =   5145
      Visible         =   0   'False
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   1535
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   720
         Left            =   75
         Top             =   75
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   1270
         BackColor       =   16777152
         Caption         =   "【 処 理 中 】 しばらくお待ち下さい"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "HG創英角ﾎﾟｯﾌﾟ体"
            Size            =   18
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
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      Height          =   345
      Index           =   1
      Left            =   7695
      ScaleHeight     =   345
      ScaleWidth      =   3900
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   30
      Width           =   3900
      Begin Cis3D_v60.CIS3D UC_3D6 
         Height          =   345
         Left            =   0
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         BackColor       =   12648384
         ForeColor       =   0
         Caption         =   "表示件数"
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
         Begin Cis3D_v60.CIS3D B1lb_Kensu 
            Height          =   285
            Left            =   990
            Top             =   30
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            BackColor       =   12648384
            ForeColor       =   0
            Caption         =   "##,###"
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
         Left            =   1935
         Top             =   0
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   609
         BackColor       =   16777152
         ForeColor       =   0
         Caption         =   "変更件数"
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
            Height          =   285
            Left            =   990
            Top             =   45
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            BackColor       =   16777152
            ForeColor       =   0
            Caption         =   "##,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
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
   Begin VB.Timer Timer1 
      Left            =   5985
      Top             =   15
   End
   Begin VB.TextBox Dummy 
      Height          =   270
      IMEMode         =   3  'ｵﾌ固定
      Left            =   0
      TabIndex        =   24
      Text            =   "Dummy"
      Top             =   0
      Width           =   210
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   1800
      Left            =   15
      ScaleHeight     =   1740
      ScaleWidth      =   15300
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1650
      Width           =   15360
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   1620
         Left            =   60
         Top             =   60
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   2858
         BackColor       =   16761087
         ForeColor       =   16711680
         Caption         =   "対象期間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   315
            Left            =   150
            Top             =   510
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            Caption         =   "開始"
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
            cAlingnment     =   7
            cPositionX      =   30
         End
         Begin CisYMD_With_Btn.CisYMDwB iH1_EYmd 
            Height          =   330
            Left            =   750
            TabIndex        =   3
            Top             =   840
            Width           =   1695
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_SYmd 
            Height          =   330
            Left            =   750
            TabIndex        =   2
            Top             =   510
            Width           =   1695
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
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   315
            Left            =   150
            Top             =   840
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            Caption         =   "終了"
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
            cAlingnment     =   7
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D CIS3D12 
            Height          =   315
            Left            =   150
            Top             =   720
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
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
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D16 
         Height          =   810
         Left            =   2670
         Top             =   60
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   1429
         BackColor       =   16761087
         ForeColor       =   16711680
         Caption         =   "仕 入 先"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   135
            TabIndex        =   4
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
            Left            =   1110
            Top             =   360
            Width           =   2205
            _ExtentX        =   3889
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
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1620
         Left            =   12720
         Top             =   60
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   2858
         ForeColor       =   16711680
         Caption         =   "品番表示"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         Begin VB.OptionButton OpH1_HinKb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "品番マスタ"
            Height          =   225
            Index           =   1
            Left            =   105
            TabIndex        =   16
            Top             =   735
            Width           =   1200
         End
         Begin VB.OptionButton OpH1_HinKb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "単価マスタ"
            Height          =   225
            Index           =   0
            Left            =   105
            TabIndex        =   15
            Top             =   465
            Value           =   -1  'True
            Width           =   1200
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   1620
         Left            =   14070
         Top             =   60
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   2858
         ForeColor       =   16711680
         Caption         =   "単価表示"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         Begin VB.OptionButton OpH1_TanKb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "無"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   17
            Top             =   465
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton OpH1_TanKb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "前回単価"
            Height          =   225
            Index           =   1
            Left            =   60
            TabIndex        =   18
            Top             =   735
            Width           =   1065
         End
      End
      Begin Cis3D_v60.CIS3D Back_Hinbn 
         Height          =   1620
         Left            =   6210
         Top             =   60
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   2858
         ForeColor       =   16711680
         Caption         =   " 品　　　番            "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         Begin VB.OptionButton H1Op_Hinku 
            BackColor       =   &H00C0C0C0&
            Caption         =   "材料/寸法"
            Height          =   240
            Index           =   3
            Left            =   5160
            TabIndex        =   9
            Top             =   120
            Width           =   1140
         End
         Begin VB.OptionButton H1Op_Hinku 
            BackColor       =   &H00C0C0C0&
            Caption         =   "自社品番"
            Height          =   240
            Index           =   2
            Left            =   3960
            TabIndex        =   8
            Top             =   120
            Width           =   1140
         End
         Begin Cis3D_v60.CIS3D H1_ZaiBox 
            Height          =   375
            Left            =   180
            Top             =   240
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
            cBoderColor2    =   16777215
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
               cBoderStyle     =   1
               cPositionX      =   90
            End
            Begin CisText_V60.CisText iH1_Long 
               Height          =   330
               Left            =   5340
               TabIndex        =   14
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
               TabIndex        =   13
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
               TabIndex        =   12
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
               TabIndex        =   11
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
               Text            =   "XXXXXXXXX1XXXXXXXXX2XXXX5"
               MaxLength       =   25
            End
         End
         Begin VB.OptionButton H1Op_Hinku 
            BackColor       =   &H00C0C0C0&
            Caption         =   "受給品番"
            Height          =   240
            Index           =   1
            Left            =   2760
            TabIndex        =   7
            Top             =   120
            Width           =   1035
         End
         Begin VB.OptionButton H1Op_Hinku 
            BackColor       =   &H00C0C0C0&
            Caption         =   "発注品番"
            Height          =   240
            Index           =   0
            Left            =   1560
            TabIndex        =   6
            Top             =   120
            Value           =   -1  'True
            Width           =   1035
         End
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   180
            TabIndex        =   10
            Top             =   420
            Width           =   6240
            _ExtentX        =   11007
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
            IMEMode         =   2
         End
         Begin Cis3D_v60.CIS3D H1lb_HinBn 
            Height          =   225
            Index           =   3
            Left            =   1080
            Top             =   1320
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   397
            Caption         =   "SPH440-0D2.9x238x1219   5XXXX3XXXXXXXXX4XXXXXXXXX5"
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
            Left            =   1080
            Top             =   840
            Width           =   2730
            _ExtentX        =   4815
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
            Left            =   360
            Top             =   840
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
            Left            =   360
            Top             =   1320
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
            Left            =   1080
            Top             =   1080
            Width           =   2730
            _ExtentX        =   4815
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
            Left            =   360
            Top             =   1080
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
            Left            =   1080
            Top             =   690
            Width           =   2250
            _ExtentX        =   3969
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
            Left            =   360
            Top             =   690
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
      End
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   810
         Left            =   2670
         Top             =   870
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   1429
         BackColor       =   8438015
         ForeColor       =   16711680
         Caption         =   "契 約 先"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         Begin CisText_V60.CisText iH1_KTorcd 
            Height          =   375
            Left            =   135
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
         Begin Cis3D_v60.CIS3D H1lb_KTorNM 
            Height          =   360
            Left            =   1110
            Top             =   360
            Width           =   2205
            _ExtentX        =   3889
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
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 購入単価マスタ　一括入力  】"
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
         Left            =   11640
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
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   16711935
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
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   7560
      Left            =   15
      ScaleHeight     =   7500
      ScaleWidth      =   15315
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3510
      Width           =   15375
      Begin vsFlexLib.vsFlexArray VsGrid 
         Height          =   7410
         Left            =   45
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   30
         Width           =   15225
         _Version        =   196608
         _ExtentX        =   26855
         _ExtentY        =   13070
         _StockProps     =   228
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   $"CXM2180.frx":0000
         Rows            =   24
         Cols            =   44
         BackColor       =   12648384
         BackColorFixed  =   14737632
         BackColorBkg    =   14737632
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   330
         GridLines       =   2
         SelectionMode   =   1
         Begin Cis3D_v60.CIS3D Input_Area 
            Height          =   345
            Left            =   9120
            Top             =   990
            Width           =   5850
            _ExtentX        =   10319
            _ExtentY        =   609
            BackColor       =   4210752
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
            Begin Cis3D_v60.CIS3D B1Btn_Kari 
               Height          =   345
               Left            =   1365
               Top             =   0
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   609
               BackColor       =   14737632
               Caption         =   "正"
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
               cBoderWidth     =   5
               cAlingnment     =   7
               cButton         =   -1  'True
            End
            Begin CisText_V60.CisText iB1_Tan 
               Height          =   375
               Left            =   0
               TabIndex        =   19
               Top             =   -15
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   661
               cFormat         =   "#,###,##0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   0
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
            Begin CisText_V60.CisText iB1_ZTan 
               Height          =   375
               Left            =   1770
               TabIndex        =   20
               Top             =   0
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   661
               cFormat         =   "#,###,##0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   0
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
            Begin CisText_V60.CisText iB1_CTan 
               Height          =   375
               Left            =   3120
               TabIndex        =   21
               Top             =   0
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   661
               cFormat         =   "#,###,##0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   0
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
            Begin CisText_V60.CisText iB1_STan 
               Height          =   375
               Left            =   4470
               TabIndex        =   22
               Top             =   0
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   661
               cFormat         =   "#,###,##0.00"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   0
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
      End
      Begin VB.TextBox B_TabDummy 
         Height          =   330
         Left            =   90
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   90
         Width           =   240
      End
   End
   Begin Cis3D_v60.CIS3D H1_ZaiDsp 
      Height          =   300
      Left            =   3360
      Top             =   1230
      Visible         =   0   'False
      Width           =   8280
      _ExtentX        =   14605
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
      Begin Cis3D_v60.CIS3D H1lb_ZaiKan 
         Height          =   300
         Left            =   7050
         Top             =   0
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "9999999999"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         cAlingnment     =   1
         cBoderStyle     =   2
      End
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   825
      Left            =   11820
      ScaleHeight     =   765
      ScaleWidth      =   3450
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   720
      Width           =   3510
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
         Left            =   885
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
         Left            =   1740
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
         Left            =   2595
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
   End
   Begin VB.PictureBox del_Picture 
      BorderStyle     =   0  'なし
      Height          =   690
      Left            =   11880
      ScaleHeight     =   690
      ScaleWidth      =   4065
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   4065
      Begin CisText_V60.CisText del_iB1_BTan 
         Height          =   375
         Left            =   810
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   661
         cFormat         =   "#,###,##0.00"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   0
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
      Begin Cis3D_v60.CIS3D del_B1Btn_Bum 
         Height          =   345
         Left            =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
         BackColor       =   14737632
         Caption         =   "有"
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
         cBoderWidth     =   5
         cAlingnment     =   7
         cButton         =   -1  'True
      End
      Begin Cis3D_v60.CIS3D del_B1Btn_Bjs 
         Height          =   345
         Left            =   405
         Top             =   240
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
         BackColor       =   14737632
         Caption         =   "自"
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
         cBoderWidth     =   5
         cAlingnment     =   7
         cButton         =   -1  'True
      End
      Begin Cis3D_v60.CIS3D del_B1Btn_Zum 
         Height          =   345
         Left            =   2160
         Top             =   240
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
         BackColor       =   14737632
         Caption         =   "有"
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
         cBoderWidth     =   5
         cAlingnment     =   7
         cButton         =   -1  'True
      End
      Begin Cis3D_v60.CIS3D del_B1Btn_Zjs 
         Height          =   345
         Left            =   2565
         Top             =   240
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
         BackColor       =   14737632
         Caption         =   "自"
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
         cBoderWidth     =   5
         cAlingnment     =   7
         cButton         =   -1  'True
      End
      Begin Cis3D_v60.CIS3D del_CIS3D10 
         Height          =   285
         Left            =   2160
         Top             =   0
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BackColor       =   14737632
         Caption         =   "有無"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   14737632
         cBoderColor2    =   14737632
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cAlingnment     =   7
         cBoderStyle     =   2
      End
      Begin Cis3D_v60.CIS3D del_CIS3D11 
         Height          =   285
         Left            =   2565
         Top             =   0
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BackColor       =   14737632
         Caption         =   "自支"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   14737632
         cBoderColor2    =   14737632
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cAlingnment     =   7
         cBoderStyle     =   2
      End
      Begin Cis3D_v60.CIS3D del_CIS3D8 
         Height          =   285
         Left            =   405
         Top             =   0
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BackColor       =   14737632
         Caption         =   "自支"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   14737632
         cBoderColor2    =   14737632
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cAlingnment     =   7
         cBoderStyle     =   2
      End
      Begin Cis3D_v60.CIS3D del_CIS3D1 
         Height          =   285
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BackColor       =   14737632
         Caption         =   "有無"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   14737632
         cBoderColor2    =   14737632
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cAlingnment     =   7
         cBoderStyle     =   2
      End
      Begin CisBtn_60.CisBtn del_PB_SET 
         Height          =   585
         Left            =   3120
         Top             =   0
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
         BtnCaption      =   "詳細"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   16711935
         BtnShadowColor  =   16777215
         PFCaption       =   "( F5 )"
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
         PFBackColor     =   16711935
      End
   End
End
Attribute VB_Name = "CXM2180"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   購入単価マスタ　一括入力
'**       フォームID    :   CXM2180
'**       処理概要      :
'**
'**       作  成  日    :   2009/10/08  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    
    
' 置換　材料購入単価ワーク　→　材料支給単価ワーク

    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim SyoriKB                 As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM                 As String       ' 処理名称(新規･修正･削除 等)
    Dim CisVsGrid               As New CisVsGrid3
    
    Dim ActObj                  As Object
    Dim ErrObj                  As Object
    Dim ActRow                  As Integer
    Dim SaveSYmd                As String
    Dim SaveEYmd                As String
    Dim SaveRow                 As Integer
    
    Dim KamiKI                  As Integer
    Dim SimoKI                  As Integer
    Dim wEtcTan                 As Currency
    
    Dim mGotColor               As Long
    Dim mLostColor              As Long

    Dim Hinku_top(3)            As Integer      ' 品番表示位置      ' 2009.10.09 add
    Dim gMaeIndex               As Integer      ' １つ前の選択品番  ' 2009.10.09 add
    Dim cnt                     As Integer      ' ワークカウンタ    ' 2009.10.09 add

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub form_load()
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
    
    mGotColor = H0Op_Tanku(1).BackColor
    mLostColor = H0Op_Tanku(0).BackColor
    
    mGotColor = H0Op_Tanku(1).BackColor
    mLostColor = H0Op_Tanku(0).BackColor
    
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = VsGrid
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 33
        
        .InitGet = False
        .Init
    End With
    Call GridInit
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
'-----------------------------------------------' 2009.10.07 add start
    Hinku_top(0) = b_HinNM(1).Top
    Hinku_top(1) = b_HinNM(1).Top
    Hinku_top(2) = b_HinNM(2).Top
    Hinku_top(3) = b_HinNM(3).Top
    H0Op_Tanku(1).Value = True
    Call H0Op_Tanku_Click(1)
   'H1_ZaiRyo.Caption = ""
   'H1_ZaiDsp.Visible = True
'-----------------------------------------------' 2009.10.07 add end
    
'   #------------------#
'   #    初期入力設定  #
'   #------------------#
    ProcHB = "H0"
    Call DispChange(ProcHB)
    
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
    
    H0Op_Tanku(1).Value = True
   
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
'          Case vbKeyF5:        If PB_SET.Visible Then Call PB_SET_Click    '【詳細】   ' 2009.10.09 del
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF11: MsgBox (ActiveControl.Name)
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

'---------------------------------------------------------------' 2009.10.09 add start
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
'---------------------------------------------------------------' 2009.10.09 add end

'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
'    Call CisDB.DBDISConnect                            ' ﾃﾞｰﾀﾍﾞｰｽ切断
    Call CisFrm.UnLoadDisp                          ' ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
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
Private Sub PB_CAN_Click()
    Call CisFrm.MousePT(11)
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   
    If ProcHB = "H1" Then
       ProcHB = "H0"
       Call DispChange(ProcHB)
       Call HeadBodyClear("H1")
'-----------------------------------------------' 2009.10.09 add start
      Call HeadBodyClear("B")
      H1_ZaiBox.Visible = False
      H1_ZaiRyo.Caption = ""
      iH1_Hinbn.Text = ""
      iH1_Hinbn.Visible = True
      If H0Op_Tanku(0).Value Then
         H0Op_Tanku(0).SetFocus
      Else
         H0Op_Tanku(1).SetFocus
      End If
'-----------------------------------------------' 2009.10.09 add end
       GoTo PB_CAN_Ed
    End If
    
    If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      iH1_Torcd.SetFocus
      iT1_Mntsu = ""
'-----------------------------------------------' 2009.10.09 add start
      For cnt = 0 To 3: H1lb_Hinbn(cnt) = "": Next
      iH1_Hinbn.Text = "": iH1_Zaist.Text = ""
      iH1_Itatu.Text = "": iH1_Width = "": iH1_Long.Text = ""
'-----------------------------------------------' 2009.10.09 add end
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
'---------------------------------------------------------------------------' 2009.10.28 upd start
    If H0Op_Tanku(0).Value = True Then
        If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
        If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd(0): Exit Sub
        If PB_Look.Tag = "iH1_KTorcd" Then Call Look_Torcd(1): Exit Sub
    Else
        If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Zairyo: Exit Sub
        If PB_Look.Tag = "iH1_Torcd" Then Call Look_ZaiTor(0):  Exit Sub
        If PB_Look.Tag = "iH1_KTorcd" Then Call Look_ZaiTor(1): Exit Sub
    End If
'---------------------------------------------------------------------------' 2009.10.28 upd end
   'If PB_LOOK.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub             ' 2009.10.28 del
   'If PB_LOOK.Tag = "iH1_KTorcd" Then Call Look_KTorcd: Exit Sub           ' 2009.10.28 del
    If PB_Look.Tag = "iH1_SYmd" Then iH1_SYmd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_EYmd" Then iH1_EYmd.ShowCalender: Exit Sub
End Sub
'-----------------------------------------------' 2009.10.09 del start
''+----------------------------+
''+    詳  細  キ ー(F5)       +
''+----------------------------+
'Private Sub PB_SET_Click()
'
'    Rv_Call3 = "K"
'    If H0Op_Tanku(0).Value Then
'       RV_HinZai = 0                   ' 品番,材料区分
'    Else
'       RV_HinZai = 1                   ' 品番,材料区分
'    End If
'    RV_Hinbn = VsGrid.TextMatrix(VsGrid.Row, 16)
'
'    RV_Torcd = RTrim(iH1_Torcd)
'    RV_KTorcd = RTrim(iH1_KTorcd)
'    RV_SYmd = ""
'    RV_EYmd = ""
'    Rv_Call3Rtn = 0
'    CXM0180.Show vbModal
'    Unload CXM0180
'    Set CXM0180 = Nothing
'
'    If RV_Rtn Then
'    End If
'    Rv_Call3 = ""
'
'    ProcHB = "H1"
'    Call ReturnPress
'
'End Sub
'-----------------------------------------------' 2009.10.09 del end

'-----------------------------------------------' 2009.10.28 add start
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
        RV_TorKb = 8    ' 9 ??
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
            H1lb_KTorNM = RV_TorRName
        End If
    End If
End Sub
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
        H1lb_Hinbn(0) = RV_HHinbn
        H1lb_Hinbn(1) = RV_JHinbn
        H1lb_Hinbn(3) = RV_Zaist & " " & _
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
        H1lb_Hinbn(2) = ZTM.自社品番
    
        For cnt = 0 To 3
            If H1Op_Hinku(cnt).Value = True Then
               iH1_Hinbn = H1lb_Hinbn(cnt)
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
       
       H1lb_Hinbn(0) = RV_HHinbn
       H1lb_Hinbn(1) = RV_JHinbn
       H1lb_Hinbn(3) = RV_Zaist & " " & _
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
       H1lb_Hinbn(2) = ZTM.自社品番
    
       For cnt = 0 To 3
           If H1Op_Hinku(cnt).Value = True Then
              iH1_Hinbn = H1lb_Hinbn(cnt)
              cnt = 4
           End If
       Next
       
       Call iH1_Torcd_LostFocus
       Call iH1_KTorcd_LostFocus
    End If
End Sub
'-----------------------------------------------' 2009.10.28 add end
'-----------------------------------------------' 2009.10.28 del start
''+----------------------------+
''+       仕入先検索            +
''+----------------------------+
'Private Sub Look_Torcd()
'
'    RV_Left = 0
'    RV_Top = 0
'    RV_TorKb = 9
'    RV_Call = ""
'
'    CKK0025.Show vbModal
'    Unload CKK0025
'    Set CKK0025 = Nothing
'
'    If RV_Rtn Then
'        iH1_Torcd = RV_TorcdK
'        H1lb_Tornm = RV_TorRName
'    End If
'End Sub
'
''+----------------------------+
''+       契約先検索           +
''+----------------------------+
'Private Sub Look_KTorcd()
'
'    RV_Left = 0
'    RV_Top = 0
'    RV_TorKb = 5
'    RV_Call = ""
'
'    CKK0025.Show vbModal
'    Unload CKK0025
'    Set CKK0025 = Nothing
'
'    If RV_Rtn Then
'        iH1_KTorcd = RV_TorcdK
'        H1lb_KTorNM = RV_TorRName
'    End If
'End Sub
'-----------------------------------------------' 2009.10.28 del end

'+--------------------+
'+     H0Op_Tanku     +
'+--------------------+
Private Sub H0Op_Tanku_Click(Index As Integer)
    H0Op_Tanku(0).BackColor = mLostColor
    H0Op_Tanku(1).BackColor = mLostColor
    H0Op_Tanku(Index).BackColor = mGotColor
    
    If Index = 0 Then
        OpH1_HinKb(1).Caption = "品番マスタ"
        Back_Hinbn.Caption = "　品　　　番"
        Back_ZHinbn.Visible = False
        Back_ZTanka.Visible = False                 ' 2009.10.13 add
        H1Op_Hinku(0).Visible = False
        H1Op_Hinku(1).Visible = False
    Else
        OpH1_HinKb(1).Caption = "材料マスタ"
        Back_Hinbn.Caption = "　品　　　番           "
        Back_ZHinbn.Visible = True
        Back_ZTanka.Visible = True                  ' 2009.10.13 add
        H1Op_Hinku(0).Visible = True
        H1Op_Hinku(1).Visible = True
    End If
    
    H1Op_Hinku(0).Visible = H0Op_Tanku(1).Value
    b_HinNM(0).Visible = False
    H1lb_Hinbn(0).Visible = False
    
    For cnt = 1 To 3
        H1Op_Hinku(cnt).Visible = H0Op_Tanku(1).Value
        b_HinNM(cnt).Visible = H0Op_Tanku(1).Value
        H1lb_Hinbn(cnt).Visible = H0Op_Tanku(1).Value
        b_HinNM(cnt).Top = Hinku_top(cnt)
        H1lb_Hinbn(cnt).Top = Hinku_top(cnt)
    Next
End Sub

'+--------------------+
'+     H1Op_Hinku     +
'+--------------------+
Private Sub H1Op_Hinku_Click(Index As Integer)
'    H1Op_Hinku(0).BackColor = mLostColor
'    H1Op_Hinku(1).BackColor = mLostColor
'    H1Op_Hinku(INDEX).BackColor = mGotColor
'
'    If INDEX = 0 Then
'       Back_ZHinbn.Caption = "発注品番"
'    Else
'       Back_ZHinbn.Caption = "受給品番"
'    End If
    H1Op_Hinku(0).BackColor = mLostColor
    H1Op_Hinku(1).BackColor = mLostColor
    H1Op_Hinku(2).BackColor = mLostColor
    H1Op_Hinku(3).BackColor = mLostColor
    H1Op_Hinku(Index).BackColor = mGotColor
    
    If H0Op_Tanku(1).Value = True Then
       For cnt = 0 To 3
           b_HinNM(cnt).Visible = True
           H1lb_Hinbn(cnt).Visible = True
       Next
    End If
    
    b_HinNM(Index).Visible = False
    H1lb_Hinbn(Index).Visible = False

    gStr = RTrim(iH1_Hinbn.Text)
    iH1_Hinbn.Text = RTrim(H1lb_Hinbn(Index))
    H1lb_Hinbn(gMaeIndex) = gStr
   'Call HinbnSet(0)

    If H0Op_Tanku(1).Value Then
       Call HinbnToLabel
       Select Case Index
          Case 0:
            iH1_Hinbn.Text = H1lb_Hinbn(0).Caption
            b_HinNM(1).Top = Hinku_top(1)
            H1lb_Hinbn(1).Top = Hinku_top(1)
            b_HinNM(2).Top = Hinku_top(2)
            H1lb_Hinbn(2).Top = Hinku_top(2)
            b_HinNM(3).Top = Hinku_top(3)
            H1lb_Hinbn(3).Top = Hinku_top(3)
         
         Case 1:
            iH1_Hinbn.Text = H1lb_Hinbn(1).Caption
            b_HinNM(0).Top = Hinku_top(1)
            H1lb_Hinbn(0).Top = Hinku_top(1)
            b_HinNM(2).Top = Hinku_top(2)
            H1lb_Hinbn(2).Top = Hinku_top(2)
            b_HinNM(3).Top = Hinku_top(3)
            H1lb_Hinbn(3).Top = Hinku_top(3)
          
          Case 2:
            iH1_Hinbn.Text = H1lb_Hinbn(2).Caption
            b_HinNM(0).Top = Hinku_top(1)
            H1lb_Hinbn(0).Top = Hinku_top(1)
            b_HinNM(1).Top = Hinku_top(2)
            H1lb_Hinbn(1).Top = Hinku_top(2)
            b_HinNM(3).Top = Hinku_top(3)
            H1lb_Hinbn(3).Top = Hinku_top(3)

          Case 3:
            iH1_Hinbn.Text = H1lb_Hinbn(3).Caption
            b_HinNM(0).Top = Hinku_top(1)
            H1lb_Hinbn(0).Top = Hinku_top(1)
            b_HinNM(1).Top = Hinku_top(2)
            H1lb_Hinbn(1).Top = Hinku_top(2)
            b_HinNM(2).Top = Hinku_top(3)
            H1lb_Hinbn(2).Top = Hinku_top(3)
            Call H1_ZaiBox_Set
          Case Else:
       End Select
    End If
    gMaeIndex = Index
    
    If H1Op_Hinku(3).Value = True Then
       iH1_Hinbn.Visible = False
       H1_ZaiBox.Top = iH1_Hinbn.Top    ' 材料入力部：品番と同じ位置
       H1_ZaiBox.Visible = True
    Else
       H1_ZaiBox.Visible = False
       iH1_Hinbn.Visible = True
    End If
    
    If H1Op_Hinku(3).Value = True Then
       Call iH1_Zaist.SetFocus
    Else
       Call iH1_Hinbn.SetFocus
    End If
End Sub

'---------------------------------------' 2009.10.28 add start
'+-----------------------------+        '
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
End Sub
'---------------------------------------' 2009.10.28 add end

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
    If TorNmGet(iH1_Torcd, 9, 1) Then
       H1lb_Tornm = TRM_RName
    End If
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
    H1lb_KTorNM = ""
    If TorNmGet(iH1_KTorcd, 5, 1) Then
       H1lb_KTorNM = TRM_RName
    End If
End Sub

'+-----------------------------+
'+  iH1_SYmd
'+-----------------------------+
Private Sub iH1_SYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_SYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_EYmd
'+-----------------------------+
Private Sub iH1_EYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_EYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Tan
'+-----------------------------+
Private Sub iB1_Tan_GotFocus()
    Set ActObj = iB1_Tan
    B_TabDummy.Tag = "S"
End Sub
Private Sub iB1_Tan_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Tan_LostFocus()
    If CisFun.Val2(iB1_Tan) <> 0 Then
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = iB1_Tan
    Else
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = ""
    End If
End Sub

'+-----------------------------+
'+  B1Btn_Kari
'+-----------------------------+
Private Sub B1Btn_Kari_GotFocus()
    Set ActObj = B1Btn_Kari
End Sub
Private Sub B1Btn_Kari_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub B1Btn_Kari_LostFocus()
    If CisFun.Val2(iB1_Tan) <> 0 Then
        If B1Btn_Kari = "仮" Then
            VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 7) = "仮"
        Else
            VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 7) = "正"
        End If
    End If
End Sub
Private Sub B1Btn_Kari_Click()
    If B1Btn_Kari = "仮" Then
        B1Btn_Kari.Caption = "正"
    Else
        B1Btn_Kari.Caption = "仮"
    End If
End Sub

'-----------------------------------------------------------------------' 2009.10.13 del start
''+-----------------------------+
''+  B1Btn_Bum
''+-----------------------------+
'Private Sub B1Btn_Bum_GotFocus()
'    Set ActObj = B1Btn_Bum
'End Sub
'Private Sub B1Btn_Bum_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call Grid_RowChange(KeyCode)
'End Sub
'Private Sub B1Btn_Bum_LostFocus()
'    If CisFun.Val2(iB1_BTan) <> 0 Then
'        If B1Btn_Bum = "無" Then
'            VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 8) = "無"
'        Else
'            VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 8) = "有"
'        End If
'    End If
'End Sub
'Private Sub B1Btn_Bum_Click()
'    If B1Btn_Bum = "有" Then
'        B1Btn_Bum.Caption = "無"
'    Else
'        B1Btn_Bum.Caption = "有"
'    End If
'End Sub
'
''+-----------------------------+
''+  B1Btn_Bjs
''+-----------------------------+
'Private Sub B1Btn_Bjs_GotFocus()
'    Set ActObj = B1Btn_Bjs
'End Sub
'Private Sub B1Btn_Bjs_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call Grid_RowChange(KeyCode)
'End Sub
'Private Sub B1Btn_Bjs_LostFocus()
'    If CisFun.Val2(iB1_BTan) <> 0 Then
'        If B1Btn_Bjs = "支" Then
'            VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 9) = "支"
'        Else
'            VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 9) = "自"
'        End If
'    End If
'End Sub
'Private Sub B1Btn_Bjs_Click()
'    If B1Btn_Bjs = "自" Then
'        B1Btn_Bjs.Caption = "支"
'    Else
'        B1Btn_Bjs.Caption = "自"
'    End If
'End Sub
''+-----------------------------+
''+  iB1_BTan
''+-----------------------------+
'Private Sub iB1_BTan_GotFocus()
'    Set ActObj = iB1_BTan
'End Sub
'Private Sub iB1_BTan_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call Grid_RowChange(KeyCode)
'End Sub
'Private Sub iB1_BTan_LostFocus()
'    If CisFun.Val2(iB1_BTan) <> 0 Then
'        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 10) = iB1_BTan
'    Else
'        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 10) = ""
'    End If
'End Sub
''+-----------------------------+
''+  B1Btn_Zum
''+-----------------------------+
'Private Sub B1Btn_Zum_GotFocus()
'    Set ActObj = B1Btn_Zum
'End Sub
'Private Sub B1Btn_Zum_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call Grid_RowChange(KeyCode)
'End Sub
'Private Sub B1Btn_Zum_LostFocus()
'    If CisFun.Val2(iB1_BTan) <> 0 Then
'        If B1Btn_Zum = "無" Then
'            VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 11) = "無"
'        Else
'            VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 11) = "有"
'        End If
'    End If
'End Sub
'Private Sub B1Btn_Zum_Click()
'    If B1Btn_Zum = "有" Then
'        B1Btn_Zum.Caption = "無"
'    Else
'        B1Btn_Zum.Caption = "有"
'    End If
'End Sub
'
''+-----------------------------+
''+  B1Btn_Zjs
''+-----------------------------+
'Private Sub B1Btn_Zjs_GotFocus()
'    Set ActObj = B1Btn_Zjs
'End Sub
'Private Sub B1Btn_Zjs_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call Grid_RowChange(KeyCode)
'End Sub
'Private Sub B1Btn_Zjs_LostFocus()
'    If CisFun.Val2(iB1_BTan) <> 0 Then
'        If B1Btn_Zjs = "支" Then
'            VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 12) = "支"
'        Else
'            VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 12) = "自"
'        End If
'    End If
'End Sub
'Private Sub B1Btn_Zjs_Click()
'    If B1Btn_Zjs = "自" Then
'        B1Btn_Zjs.Caption = "支"
'    Else
'        B1Btn_Zjs.Caption = "自"
'    End If
'End Sub
'-----------------------------------------------------------------------' 2009.10.13 del end

'+-----------------------------+
'+  iB1_ZTan
'+-----------------------------+
Private Sub iB1_ZTan_GotFocus()
    Set ActObj = iB1_ZTan
End Sub
Private Sub iB1_ZTan_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_ZTan_LostFocus()
    If CisFun.Val2(iB1_ZTan) <> 0 Then
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 8) = iB1_ZTan    ' 2009.10.13 upd : 位置 13-->8
    Else
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 8) = ""          ' 2009.10.13 upd : 位置 13-->8
    End If
End Sub

'+-----------------------------+
'+  iB1_CTan
'+-----------------------------+
Private Sub iB1_CTan_GotFocus()
    Set ActObj = iB1_CTan
End Sub
Private Sub iB1_CTan_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_CTan_LostFocus()
    If CisFun.Val2(iB1_CTan) <> 0 Then
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 9) = iB1_CTan    ' 2009.10.13 upd : 位置 14-->9
    Else
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 9) = ""          ' 2009.10.13 upd : 位置 14-->9
    End If
End Sub

'+-----------------------------+
'+  iB1_STan
'+-----------------------------+
Private Sub iB1_STan_GotFocus()
    Set ActObj = iB1_STan
    B_TabDummy.Tag = "E"
End Sub
Private Sub iB1_STan_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_STan_LostFocus()
    If CisFun.Val2(iB1_STan) <> 0 Then
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 10) = iB1_STan   ' 2009.10.13 upd : 位置 15-->10
    Else
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 10) = ""         ' 2009.10.13 upd : 位置 15-->10
    End If
End Sub

'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If B_TabDummy.Tag = "S" Then
        Call TextToVsGrid(Val(Input_Area.Tag))
        Call Grid_RowChange(vbKeyUp)
    End If
    If B_TabDummy.Tag = "E" Then
        Call TextToVsGrid(Val(Input_Area.Tag))
        Call Grid_RowChange(vbKeyDown)
    End If
    SendKeys "{Tab}"
End Sub

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
'   Call CisFrm.TimeSet(1)                  ' 2009.10.09 del
   CisFrm.TimeSet CisFun.INI_FileID, 110    ' 2009.10.09 add
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
    Dim H0Mode      As Boolean
    Dim H0Color     As Long
    Dim H1Mode      As Boolean
    Dim H1Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    H0Mode = False
    H0Color = gPLostSel
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
       Case "H0"
          H0Mode = True
          H0Color = gPGotSel
       Case "H1"
          H1Mode = True
          H1Color = gPGotSel
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
    End Select
    
    H0_Area1.Enabled = H0Mode
    H0_Area1.BackColor = H0Color
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    PB_Look.Visible = H1Mode
    PB_CAN.Visible = Not H0Mode
    PB_END.Visible = H0Mode
'    PB_SET.Visible = B1Mode        ' 2009.10.13 del
    
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
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
                .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
        CisVsGrid.Clear
    End If
    If HeadBody = "B1" Then
        CisVsGrid.Clear
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H0" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
'-----------------------------------------------' 2009.10.09 upd start
'        H1Op_Hinku(0).Value = True
'        Call H1Op_Hinku_Click(0)
        If H0Op_Tanku(1).Value = True Then
           H1Op_Hinku(0).Value = True
           Call H1Op_Hinku_Click(0)
           gMaeIndex = 0
        Else
            For cnt = 0 To 3
                H1Op_Hinku(cnt).Visible = False
                H1lb_Hinbn(cnt).Visible = False
                b_HinNM(cnt).Visible = False
            Next
        End If
        Call iH1_SYmd.SetFocus
'-----------------------------------------------' 2009.10.09 upd end
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        VsGrid.Row = 1
        Call VsGridToText(VsGrid.Row)
        SaveRow = 1
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If H0Op_Tanku(0).Value Then
           If Not DBPut Then GoTo ReturnPress_Ed
        Else
           If Not DBPutZ Then GoTo ReturnPress_Ed
        End If
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        '---------------------------------------' 2009.10.09 add start
        If H0Op_Tanku(1).Value = True Then
           For cnt = 0 To 3: H1lb_Hinbn(cnt).Caption = "": Next
           H1lb_ZaiKan.Caption = ""
           H1_ZaiRyo.Caption = ""
        End If
        '---------------------------------------' 2009.10.09 add end
        iH1_Torcd.SetFocus
        GoTo ReturnPress_Ed
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
        
'---------------------------------------------------------------' 2009.10.09 add start
Private Sub HinbnToLabel()
    If H0Op_Tanku(1).Value = False Then Exit Sub
    
    If H1Op_Hinku(0).Value = True Then
       H1lb_Hinbn(0).Caption = Trim(iH1_Hinbn.Text)
    ElseIf H1Op_Hinku(1).Value = True Then
       H1lb_Hinbn(1).Caption = Trim(iH1_Hinbn.Text)
    ElseIf H1Op_Hinku(2).Value = True Then
       H1lb_Hinbn(2).Caption = Trim(iH1_Hinbn.Text)
    Else
       H1lb_Hinbn(3).Caption = "": H1_ZaiRyo.Caption = "": iH1_Hinbn.Text = ""
       If Trim(iH1_Zaist.Text) <> "" Then
          H1lb_Hinbn(3).Caption = Trim(iH1_Zaist.Text) & " " & _
                                  Format(CisFun.Val2(iH1_Itatu.Text), "###0.00") & "X" & _
                                  Format(CisFun.Val2(iH1_Width.Text), "###0.00") & "X" & _
                                  Format(CisFun.Val2(iH1_Long.Text), "###0.00")
          H1_ZaiRyo.Caption = Trim(iH1_Zaist.Text) & Space(25 - Len(Trim(iH1_Zaist.Text))) & "/" & _
                              Format(CisFun.Val2(iH1_Itatu.Text), "0000.00") & "x" & _
                              Format(CisFun.Val2(iH1_Width.Text), "0000.00") & "x" & _
                              Format(CisFun.Val2(iH1_Long.Text), "0000.00")
          iH1_Hinbn.Text = H1lb_Hinbn(3).Caption
       End If
    End If

    For cnt = 0 To 3
        If Trim(H1lb_Hinbn(cnt)) <> "" Then Exit For
    Next
    If cnt > 3 Then
        H1lb_ZaiKan.Caption = ""        ' 全ての品番が空白ならば材料管理番号を無しにする
    End If
End Sub
'---------------------------------------------------------------' 2009.10.09 add end

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
    
    '============= 対象期間(開始必須) ==============
    If RTrim(iH1_SYmd) = "" Then
        CisFun.MB_MSG(1) = "     　【  対象期間 必須！ 】     "
        CisFun.MB_MSG(3) = "     対象期間開始 を入力して下さい。      "
        CisFun.MBOX
        iH1_SYmd.SetFocus
        Exit Function
    End If
    '============= 対象期間(開始異常) ==============
    If Not iH1_SYmd.cChkResult Then
        CisFun.MB_MSG(1) = "     　【  対象期間 異常！ 】     "
        CisFun.MB_MSG(3) = "     対象期間開始 を再入力して下さい。      "
        CisFun.MBOX
        iH1_SYmd.SetFocus
        Exit Function
    End If
    '============= 対象期間(終了必須) ==============
    If RTrim(iH1_EYmd) = "" Then
        CisFun.MB_MSG(1) = "     　【  対象期間 必須！ 】     "
        CisFun.MB_MSG(3) = "     対象期間終了 を入力して下さい。      "
        CisFun.MBOX
        iH1_EYmd.SetFocus
        Exit Function
    End If
    '============= 対象期間(終了異常) ==============
    If Not iH1_EYmd.cChkResult Then
        CisFun.MB_MSG(1) = "     　【  対象期間 異常！ 】     "
        CisFun.MB_MSG(3) = "     対象期間終了 を再入力して下さい。      "
        CisFun.MBOX
        iH1_EYmd.SetFocus
        Exit Function
    End If
    
    If RTrim(iH1_SYmd) > RTrim(iH1_EYmd) Then
        CisFun.MB_MSG(1) = "     　【  対象期間 異常！ 】     "
        CisFun.MB_MSG(3) = "     対象期間 開始 ＜　終了にして下さい。      "
        CisFun.MBOX
        iH1_EYmd.SetFocus
        Exit Function
    End If
    
    '================ 仕入先(必須) =================
    If Trim(iH1_Torcd) = "" Then
        CisFun.MB_MSG(2) = "    仕入先を入力して下さい            "
        CisFun.MBOX
        iH1_Torcd.SetFocus
        Exit Function
    End If
    '=========== 仕入先(マスタチェック) ============
    If Not TorNmGet(iH1_Torcd, 9) Then
        CisFun.MB_MSG(2) = "    仕入先　未登録！（取引先マスタ）            "
        CisFun.MBOX
        iH1_Torcd.SetFocus
        Exit Function
    End If
    
' マスタ入力 SELECT文 ｾｯﾄ
    If Not DBInput("INV") Then
        If H0Op_Tanku(0).Value Then
           CisFun.MB_MSG(2) = "    購入単価マスタ　未登録！           "
        Else
           CisFun.MB_MSG(2) = "    材料購入単価マスタ　未登録！　　　　"
        End If
        CisFun.MBOX
        iH1_Torcd.SetFocus
        Exit Function
    End If
       
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
'
    With VsGrid
        For gCnt1 = 1 To .Rows - 1
            If RTrim(.TextMatrix(gCnt1, CisVsGrid.SelectCol)) <> "" Then
            '<< 修正の時 >>
                If CisFun.Val2(.TextMatrix(gCnt1, 21)) <> 0 Then                ' memo : 21は単価保持場所
                    If CisFun.Val2(.TextMatrix(gCnt1, 6)) = 0 Then
                        CisFun.MB_MSG(1) = "    　修正の時、単価＝０は不可！            "
                        CisFun.MB_MSG(3) = "    　　" & gCnt1 & "行目            "
                        CisFun.MBOX
                        Set ErrObj = iB1_Tan
                        GoTo Body1Chk_Err
                    End If
                End If
                
            '<< 単価チェック >>
                If CCur(CisFun.Val2(.TextMatrix(gCnt1, 6))) <> _
                   CCur(CisFun.Val2(.TextMatrix(gCnt1, 8)) + _
                        CisFun.Val2(.TextMatrix(gCnt1, 9)) + _
                        CisFun.Val2(.TextMatrix(gCnt1, 10))) Then   ' 2009.10.13 upd : 位置変更 および CisFun.Val2(.TextMatrix(gCnt1, 10)) 削除

                    If H0Op_Tanku(0).Value = True Then                                          ' 2009.10.13 add
                       CisFun.MB_MSG(1) = "単価 <> 部品単価 + 加工単価 + その他単価　です。"
                    Else                                                                        ' 2009.10.13 add
                       CisFun.MB_MSG(1) = "単価 <> 材料単価 + 加工単価 + その他単価　です。"    ' 2009.10.13 add
                    End If                                                                      ' 2009.10.13 add
                    
                    CisFun.MB_MSG(3) = "                    よろしいですか？  【" & gCnt1 & "行目】      "
                    CisFun.MB_Button = OK_CAN
                    If Not CisFun.MBOX Then
                       Set ErrObj = iB1_Tan
                       GoTo Body1Chk_Err
                    End If
                End If
            End If
        Next gCnt1
    End With
'
    Body1Chk = True
    Exit Function
Body1Chk_Err:
    VsGrid.Row = gCnt1
    Call VsGridToText(VsGrid.Row)
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
    
    If H0Op_Tanku(0).Value Then
       '=*=*=*=*=*=*=*=*=*=*=*=*=*=*
       '=*     単価区分(部品)     *=
       '=*=*=*=*=*=*=*=*=*=*=*=*=*=*
       If Not STWCreate Then GoTo DBInput_Ed
        
       gSL_Select = ""
       gSL_Select = gSL_Select & " SELECT KT.*,ISNULL(HN.表示品番,KT.品番) 表示品番,"
       gSL_Select = gSL_Select & "        CASE WHEN ISNULL(TR.略称, '') <> ''"
       gSL_Select = gSL_Select & "             THEN TR.取引先 + ':' + TR.略称"
       gSL_Select = gSL_Select & "             ELSE TR.取引先 END 取引先名,"
       gSL_Select = gSL_Select & "        '' 材料管理番号"
       gSL_Select = gSL_Select & "   FROM 購入単価ワーク KT"
       gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番マスタ HN"
       gSL_Select = gSL_Select & "             ON HN.品番 = KT.品番"
       gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR"
       gSL_Select = gSL_Select & "             ON TR.取引先CD = KT.取引先"
       gSL_Select = gSL_Select & "            AND TR.取引先区分 = 1"
       gSL_Select = gSL_Select & "        LEFT OUTER JOIN (SELECT 品番, 取引先"
       gSL_Select = gSL_Select & "                           FROM 品番取引先マスタ"
       gSL_Select = gSL_Select & "                          WHERE 品目 = 1"
       gSL_Select = gSL_Select & "                          GROUP BY 品番, 取引先) HT"
       gSL_Select = gSL_Select & "             ON HT.品番 = KT.品番"
       gSL_Select = gSL_Select & "            AND HT.取引先 = KT.取引先"
       gSL_Select = gSL_Select & "  WHERE 処理端末 = HOST_NAME()"
       gSL_Select = gSL_Select & "    AND ISNULL(TR.取引先CD,'') <> ''"
       gSL_Select = gSL_Select & "  ORDER BY KT.品番"
        
       If Not KTMRead(gSL_Select, 0, 1) Then
           Call KTMClose
       Else
           Call GridSet
           DBInput = True
       End If
    
    Else
       '=*=*=*=*=*=*=*=*=*=*=*=*=*=*
       '=*     単価区分(材料)     *=
       '=*=*=*=*=*=*=*=*=*=*=*=*=*=*
       If Not ZTWCreate Then GoTo DBInput_Ed
        
       gSL_Select = ""
       gSL_Select = gSL_Select & " SELECT ZW.*, ZA.表示寸法 表示品番,"                  ' 2009.10.13 rep : ZW.品番-->ZA.表示寸法
       gSL_Select = gSL_Select & "        CASE WHEN ISNULL(TR.略称, '') <> ''"
       gSL_Select = gSL_Select & "             THEN TR.取引先 + ':' + TR.略称"
       gSL_Select = gSL_Select & "             ELSE TR.取引先 END 取引先名"
       gSL_Select = gSL_Select & "      , ZM.発注品番, ZM.受給品番, ZM.自社品番"        ' 2009.10.13 add
       gSL_Select = gSL_Select & "      , ZA.材質, ZA.板厚, ZA.幅, ZA.長さ"             ' 2009.10.13 add
       gSL_Select = gSL_Select & "   FROM 材料購入単価ワーク ZW"
'       gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料マスタ ZM"
'       If H1Op_Hinku(0).Value Then
'          gSL_Select = gSL_Select & "             ON ZM.発注品番 = ZW.品番"
'       Else
'          gSL_Select = gSL_Select & "             ON ZM.受給品番 = ZW.品番"
'       End If
       gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料マスタ ZA"                ' 2009.10.13 add
       gSL_Select = gSL_Select & "             ON ZA.材料管理番号 = ZW.材料管理番号"    ' 2009.10.13 add
       gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料取引先マスタ ZM"
       gSL_Select = gSL_Select & "             ON ZM.材料管理番号 = ZW.材料管理番号"
       gSL_Select = gSL_Select & "            AND ZM.取引先 = ZW.取引先"
       gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR"
       gSL_Select = gSL_Select & "             ON TR.取引先CD = ZW.取引先"
       gSL_Select = gSL_Select & "            AND TR.取引先区分 = 1"
       gSL_Select = gSL_Select & "  WHERE 処理端末 = HOST_NAME()"
       gSL_Select = gSL_Select & "    AND ISNULL(TR.取引先CD,'') <> ''"
       gSL_Select = gSL_Select & "  ORDER BY ZW.取引先, ZW.品番"                        ' 2009.10.13 rep : ZW.取引先 を追加
        
       If Not KTMRead(gSL_Select, 0, 1) Then
           Call KTMClose
       Else
           Call GridSet
           DBInput = True
       End If
    End If
    
DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function

'+----------------------------+
'+     購入単価ワーク作成     +
'+----------------------------+
Private Function STWCreate() As Boolean
       '=*=*=*=*=*=*=*=*=*=*=*=*=*=*
       '=*     単価区分(部品)     *=
       '=*=*=*=*=*=*=*=*=*=*=*=*=*=*
    STWCreate = False

    gSL_Select = ""
    gSL_Select = gSL_Select & " DELETE FROM 購入単価ワーク WHERE 処理端末 = HOST_NAME()"
    CisDB.SQL = gSL_Select
    If Not CisDB.DBExec Then Exit Function
    
    If OpH1_HinKb(0).Value Then
        With CisFun
            .Where_CreateKB = NewDocument
            '=========== 取引先 ============
            .Where_Create AlphaMe, "KT.取引先", RTrim(iH1_Torcd)                ' 2009.10.29 upd : KT.追加
            '=========== 契約先 ============
            If Trim(iH1_KTorcd) <> "" Then                                      ' 2009.10.29 add
               .Where_Create AlphaMe, "KT.契約先", RTrim(iH1_KTorcd), , Create  ' 2009.10.29 upd : KT.追加
            End If
            '============ 品番 =============
            .Where_Create AlphaMe, "KT.品番", RTrim(iH1_Hinbn)                  ' 2009.10.29 upd : KT.追加
        End With
        
        '<< 購入単価マスタより品番抽出  >>
        gSL_Select = ""
        gSL_Select = gSL_Select & " INSERT INTO 購入単価ワーク"
        gSL_Select = gSL_Select & "             (処理端末, 品番, 取引先, 契約先, 前回履歴NO, 履歴NO)"
'-----------------------------------------------------------------------------------------------------------' 2009.10.29 upd start
        gSL_Select = gSL_Select & " SELECT    HOST_NAME(), KT.品番, KT.取引先, KT.契約先, 0         , 0"
        gSL_Select = gSL_Select & "   FROM 購入単価マスタ KT "
        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番取引先マスタ HT "                        ' add
        gSL_Select = gSL_Select & "               ON  HT.品番 = KT.品番 "                               ' add
        gSL_Select = gSL_Select & CisFun.Where_Phrase
        gSL_Select = gSL_Select & "    AND (HT.手配区分 = 0 OR HT.手配区分 = 1 OR HT.手配区分 = 3) "    ' add
        gSL_Select = gSL_Select & "  GROUP BY KT.品番, KT.取引先, KT.契約先"
'-----------------------------------------------------------------------------------------------------------' 2009.10.29 upd end
        CisDB.SQL = gSL_Select
        If Not CisDB.DBExec Then Exit Function
    Else
    
        With CisFun
            .Where_CreateKB = NewDocument
            .Where_Create Numeric, "品目", 1
            '=========== 取引先 ============
            .Where_Create AlphaMe, "取引先", RTrim(iH1_Torcd)
            '=========== 契約先 ============
            If Trim(iH1_KTorcd) <> "" Then                                      ' 2009.10.29 add
               .Where_Create AlphaMe, "契約先", RTrim(iH1_KTorcd), , Create     ' 2009.10.29 add
            End If
            '============ 品番 =============
            .Where_Create AlphaMe, "品番", RTrim(iH1_Hinbn)
        End With
        
        '<< 品番取引先マスタより品番抽出  >>
        gSL_Select = ""
        gSL_Select = gSL_Select & " INSERT INTO 購入単価ワーク"
        gSL_Select = gSL_Select & "             (処理端末, 品番, 取引先, 契約先, 前回履歴NO, 履歴NO)"
        gSL_Select = gSL_Select & " SELECT    HOST_NAME(), 品番, 取引先, '" & RTrim(iH1_KTorcd) & "', 0         , 0"
        gSL_Select = gSL_Select & "   FROM 品番取引先マスタ"
        gSL_Select = gSL_Select & CisFun.Where_Phrase
        'gSL_Select = gSL_Select & "   AND 品目 = 1"
        gSL_Select = gSL_Select & "    AND (手配区分 = 0 OR 手配区分 = 1 OR 手配区分 = 3) "                 ' 2009.10.29 add
        gSL_Select = gSL_Select & " GROUP BY 品番, 取引先"
        CisDB.SQL = gSL_Select
        If Not CisDB.DBExec Then Exit Function
    End If
    '============ ワーク作成 ===============
    With CisFun
        .Where_CreateKB = NewDocument
        '=========== 取引先 ============
        .Where_Create AlphaMe, "KT.取引先", RTrim(iH1_Torcd)
        '=========== 契約先 ============
        .Where_Create AlphaMe, "KT.契約先", RTrim(iH1_KTorcd), , Create
        '============ 品番 =============
        .Where_Create AlphaMe, "KT.品番", RTrim(iH1_Hinbn)
    End With

    gSL_Select = ""
    gSL_Select = gSL_Select & " UPDATE 購入単価ワーク SET"
    gSL_Select = gSL_Select & "        履歴NO = ISNULL(KT1.履歴NO, 0),"
    gSL_Select = gSL_Select & "        単価 = 0,"
    gSL_Select = gSL_Select & "        仮単価区分 = 0,"
    gSL_Select = gSL_Select & "        有無償区分 = 0,"
    gSL_Select = gSL_Select & "        適用開始日付 = ISNULL(KT1.適用開始日付, ''),"
    gSL_Select = gSL_Select & "        適用終了日付 = ISNULL(KT1.適用終了日付, ''),"
    gSL_Select = gSL_Select & "        前回履歴NO = ISNULL(KT1.履歴NO, 0),"
    gSL_Select = gSL_Select & "        前回単価 =  ISNULL(KT1.単価, 0),"
    gSL_Select = gSL_Select & "        前回仮単価区分 = ISNULL(KT1.仮単価区分, ''),"
    gSL_Select = gSL_Select & "        前回有無賞区分 = ISNULL(KT1.有無償区分, '')"
    gSL_Select = gSL_Select & "   FROM 購入単価ワーク KT"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT1"
    gSL_Select = gSL_Select & "             ON KT1.取引先 = KT.取引先"
    gSL_Select = gSL_Select & "            AND KT1.品番 = KT.品番"
    gSL_Select = gSL_Select & "            AND KT1.契約先 = KT.契約先"
    gSL_Select = gSL_Select & "            AND KT1.履歴NO = (SELECT MAX(KT1S.履歴NO)履歴NO"
    gSL_Select = gSL_Select & "                                FROM 購入単価マスタ KT1S"
    gSL_Select = gSL_Select & "                               WHERE KT1S.取引先 = KT.取引先"
    gSL_Select = gSL_Select & "                                 AND KT1S.品番 = KT.品番"
    gSL_Select = gSL_Select & "                                 AND KT1S.契約先 = KT.契約先"
    gSL_Select = gSL_Select & "                               GROUP BY KT1S.取引先,KT1S.品番,KT1S.契約先)"
    gSL_Select = gSL_Select & CisFun.Where_Phrase
    CisDB.SQL = gSL_Select
    If Not CisDB.DBExec Then Exit Function
    
    STWCreate = True

End Function

'+--------------------------------+
'+     材料購入単価ワーク作成     +
'+--------------------------------+
Private Function ZTWCreate() As Boolean
    ZTWCreate = False

    gSL_Select = ""
    gSL_Select = gSL_Select & " DELETE FROM 材料購入単価ワーク WHERE 処理端末 = HOST_NAME()"
    CisDB.SQL = gSL_Select
    If Not CisDB.DBExec Then Exit Function
        
    With CisFun
        .Where_CreateKB = NewDocument
        '=========== 取引先 ============
        .Where_Create AlphaMe, "ZT.取引先", RTrim(iH1_Torcd)
        '=========== 契約先 ============
        If Trim(iH1_KTorcd) <> "" Then                                      ' 2009.10.29 add
           .Where_Create AlphaMe, "ZT.契約先", RTrim(iH1_KTorcd), , Create
        End If                                                              ' 2009.10.29 add
        '============ 品番 =============
        If OpH1_HinKb(0).Value Then
            If H1Op_Hinku(0).Value Then
                .Where_Create AlphaMe, "ZM.発注品番", RTrim(iH1_Hinbn)
            ElseIf H1Op_Hinku(1).Value Then
                .Where_Create AlphaMe, "ZM.受給品番", RTrim(iH1_Hinbn)
            ElseIf H1Op_Hinku(2).Value Then
                .Where_Create AlphaMe, "ZM.自社品番", RTrim(iH1_Hinbn)
            End If
        Else
            If H1Op_Hinku(0).Value Then
                .Where_Create AlphaMe, "ZT.発注品番", RTrim(iH1_Hinbn)
            ElseIf H1Op_Hinku(1).Value Then
                .Where_Create AlphaMe, "ZT.受給品番", RTrim(iH1_Hinbn)
            ElseIf H1Op_Hinku(2).Value Then
                .Where_Create AlphaMe, "ZT.自社品番", RTrim(iH1_Hinbn)
            End If
        End If
        If H1Op_Hinku(3).Value = True Then
            If CisFun.Val2(iH1_Itatu) = 0 Then iH1_Itatu = ""
            If CisFun.Val2(iH1_Width) = 0 Then iH1_Width = ""
            If CisFun.Val2(iH1_Long) = 0 Then iH1_Long = ""
            If Not (Trim(iH1_Zaist) = "" And Trim(iH1_Itatu) = "" And Trim(iH1_Width) = "" And Trim(iH1_Long) = "") Then
               If Trim(iH1_Itatu) = "" And Trim(iH1_Width) = "" And Trim(iH1_Long) = "" Then
                  If .Where_Phrase = "" Then
                     .Where_Phrase = .Where_Phrase & "   WHERE ZA.材質 LIKE '" & Trim(iH1_Zaist) & "%'"
                  Else
                     .Where_Phrase = .Where_Phrase & "     AND ZA.材質 LIKE '" & Trim(iH1_Zaist) & "%'"
                  End If
               Else
                  .Where_Create AlphaMe, "ZA.材質", RTrim(iH1_Zaist)
                  If RTrim(iH1_Itatu) <> "" Then .Where_Create AlphaMe, "ZA.板厚", RTrim(iH1_Itatu)
                  If RTrim(iH1_Width) <> "" Then .Where_Create AlphaMe, "ZA.幅", RTrim(iH1_Width)
                  If RTrim(iH1_Long) <> "" Then .Where_Create AlphaMe, "ZA.長さ", RTrim(iH1_Long)
               End If
            End If
        End If
    End With
        
    If OpH1_HinKb(0).Value Then
        '<< 購入単価マスタより品番抽出  >>
        gSL_Select = ""
        gSL_Select = gSL_Select & " INSERT INTO 材料購入単価ワーク"
        gSL_Select = gSL_Select & "             (処理端末, 材料管理番号, 品番, 取引先, 契約先, 前回履歴NO, 履歴NO)"
'-----------------------------------------------------------------------------------------------------------------------' 2009.10.13 del start
'        If H1Op_Hinku(0).Value Then
'           gSL_Select = gSL_Select & " SELECT HOST_NAME(), ZT.材料管理番号, ISNULL(ZM.発注品番,''), ZT.取引先, ZT.契約先, 0, 0"
'        Else
'           gSL_Select = gSL_Select & " SELECT HOST_NAME(), ZT.材料管理番号, ISNULL(ZM.受給品番,''), ZT.取引先, ZT.契約先, 0, 0"
'        End If
'-----------------------------------------------------------------------------------------------------------------------' 2009.10.13 del end
        gSL_Select = gSL_Select & " SELECT HOST_NAME(), ZT.材料管理番号, ISNULL(ZA.材質,''), ZT.取引先, ZT.契約先, 0, 0" '2009.10.13 add : (上記品番別から材質に変更)
        gSL_Select = gSL_Select & "   FROM 材料購入単価マスタ ZT"
'        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料マスタ ZM"
'        gSL_Select = gSL_Select & "             ON ZM.材料管理番号 = ZT.材料管理番号"
        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料取引先マスタ ZM"
        gSL_Select = gSL_Select & "             ON ZM.材料管理番号 = ZT.材料管理番号"
        'gSL_Select = gSL_Select & "            AND ZM.品目 = 1"
        gSL_Select = gSL_Select & "            AND ZM.取引先 = ZT.取引先"
        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料マスタ ZA"                           ' 2009.10.13 add
        gSL_Select = gSL_Select & "             ON ZA.材料管理番号 = ZT.材料管理番号"               ' 2009.10.13 add
        gSL_Select = gSL_Select & CisFun.Where_Phrase
        
'-----------------------------------------------------------------------------------------------------------------------' 2009.10.13 add start
        If CisFun.Where_Phrase = "" Then
           gSL_Select = gSL_Select & "  WHERE ZT.材料管理番号 != '' AND ZA.材料管理番号 != '' "
        Else
           gSL_Select = gSL_Select & "    AND ZT.材料管理番号 != '' AND ZA.材料管理番号 != '' "
        End If
        gSL_Select = gSL_Select & "    AND (ZM.手配区分 = 0 OR ZM.手配区分 = 1 OR ZM.手配区分 = 3) "        ''''' 2009.10.29 add
        gSL_Select = gSL_Select & "  GROUP BY ISNULL(ZA.材質,''), ZT.取引先, ZT.契約先, ZT.材料管理番号"
'-----------------------------------------------------------------------------------------------------------------------' 2009.10.13 add end
'-----------------------------------------------------------------------------------------------------------------------' 2009.10.13 del start
'        If H1Op_Hinku(0).Value Then
'           gSL_Select = gSL_Select & "  GROUP BY ISNULL(ZM.発注品番,''), ZT.取引先, ZT.契約先, ZT.材料管理番号"
'        Else
'           gSL_Select = gSL_Select & "  GROUP BY ISNULL(ZM.受給品番,''), ZT.取引先, ZT.契約先, ZT.材料管理番号"
'        End If
'-----------------------------------------------------------------------------------------------------------------------' 2009.10.13 del end
        CisDB.SQL = gSL_Select
        
        If Not CisDB.DBExec Then Exit Function
    Else
        '<< 材料取引先マスタより品番抽出  >>
        gSL_Select = ""
        gSL_Select = gSL_Select & " INSERT INTO 材料購入単価ワーク"
        gSL_Select = gSL_Select & "             (処理端末, 材料管理番号, 品番, 取引先, 契約先, 前回履歴NO, 履歴NO)"
'-----------------------------------------------------------------------------------------------------------------------' 2009.10.13 del start
'        If H1Op_Hinku(0).Value Then
'           gSL_Select = gSL_Select & " SELECT HOST_NAME(), ZT.材料管理番号, ISNULL(ZT.発注品番,''), ZT.取引先, '" & RTrim(iH1_KTorcd) & "', 0, 0"
'        Else
'           gSL_Select = gSL_Select & " SELECT HOST_NAME(), ZT.材料管理番号, ISNULL(ZT.受給品番,''), ZT.取引先, '" & RTrim(iH1_KTorcd) & "', 0, 0"
'        End If
'-----------------------------------------------------------------------------------------------------------------------' 2009.10.13 del end
        gSL_Select = gSL_Select & " SELECT HOST_NAME(), ZT.材料管理番号, ISNULL(ZA.材質,''), ZT.取引先, ZT.契約先, 0, 0" '2009.10.13 add : (上記品番別から材質に変更)
        
        gSL_Select = gSL_Select & "   FROM 材料取引先マスタ ZT"
        gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料マスタ ZA"
        gSL_Select = gSL_Select & "             ON ZA.材料管理番号 = ZT.材料管理番号"
        gSL_Select = gSL_Select & CisFun.Where_Phrase
'-----------------------------------------------------------------------------------------------------------------------' 2009.10.13 add start
        If CisFun.Where_Phrase = "" Then                                                                    ' 2009.10.13 add start
           gSL_Select = gSL_Select & "  WHERE ZT.材料管理番号 != '' AND ZA.材料管理番号 != '' "
        Else
           gSL_Select = gSL_Select & "    AND ZT.材料管理番号 != '' AND ZA.材料管理番号 != '' "
        End If
        gSL_Select = gSL_Select & "    AND (ZT.手配区分 = 0 OR ZT.手配区分 = 1 OR ZT.手配区分 = 3) "        ''''' 2009.10.29 add
        gSL_Select = gSL_Select & "  GROUP BY ISNULL(ZA.材質,''), ZT.取引先, ZT.契約先, ZT.材料管理番号"    ' 2009.10.13 add end
'-----------------------------------------------------------------------------------------------------------------------' 2009.10.13 add end
'        gSL_Select = gSL_Select & "   AND (ZT.品目 = 0 OR (ZT.品目 = 1 AND ZT.手配区分 = '1'))"
        'gSL_Select = gSL_Select & "   AND ZT.品目 = 1 "
'-----------------------------------------------------------------------------------------------------------------------' 2009.10.13 del start
'        If H1Op_Hinku(0).Value Then
'           gSL_Select = gSL_Select & "  GROUP BY ISNULL(ZT.発注品番,''), ZT.取引先, ZT.材料管理番号"
'        Else
'           gSL_Select = gSL_Select & "  GROUP BY ISNULL(ZT.受給品番,''), ZT.取引先, ZT.材料管理番号"
'        End If
'-----------------------------------------------------------------------------------------------------------------------' 2009.10.13 del end
        CisDB.SQL = gSL_Select
        If Not CisDB.DBExec Then Exit Function
    End If
''''Call Test_LogOutput(gSL_Select, 1)  '''試験用

    '============ ワーク作成 ===============
    With CisFun
        .Where_CreateKB = NewDocument
        '=========== 取引先 ============
        .Where_Create AlphaMe, "ST.取引先", RTrim(iH1_Torcd)
        '=========== 契約先 ============
        If Trim(iH1_KTorcd) <> "" Then                                      ' 2009.10.29 add
           .Where_Create AlphaMe, "ST.契約先", RTrim(iH1_KTorcd), , Create
        End If                                                              ' 2009.10.29 add
        '============ 品番 =============
        .Where_Create AlphaMe, "ST.品番", RTrim(iH1_Hinbn)
    End With

    gSL_Select = ""
    gSL_Select = gSL_Select & " UPDATE 材料購入単価ワーク SET"
    gSL_Select = gSL_Select & "        履歴NO = ISNULL(ST1.履歴NO,0),"
    gSL_Select = gSL_Select & "        単価 = 0,"
    gSL_Select = gSL_Select & "        仮単価区分 = 0,"
    gSL_Select = gSL_Select & "        適用開始日付 = ISNULL(ST1.適用開始日付, ''),"
    gSL_Select = gSL_Select & "        適用終了日付 = ISNULL(ST1.適用終了日付, ''),"
    gSL_Select = gSL_Select & "        有無償区分 = 0,"
    gSL_Select = gSL_Select & "        前回履歴NO = ISNULL(ST1.履歴NO, 0),"
    gSL_Select = gSL_Select & "        前回単価 =  ISNULL(ST1.単価,0),"
    gSL_Select = gSL_Select & "        前回仮単価区分 = ISNULL(ST1.仮単価区分, ''),"
    gSL_Select = gSL_Select & "        前回有無賞区分 = ISNULL(ST1.有無償区分, '')"
    gSL_Select = gSL_Select & "   FROM 材料購入単価ワーク ST"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料購入単価マスタ ST1"
    gSL_Select = gSL_Select & "             ON ST1.取引先 = ST.取引先"
    gSL_Select = gSL_Select & "            AND ST1.契約先 = ST.契約先"
    gSL_Select = gSL_Select & "            AND ST1.材料管理番号 = ST.材料管理番号"
    gSL_Select = gSL_Select & "            AND ST1.履歴NO = (SELECT MAX(ST1S.履歴NO) 履歴NO"
    gSL_Select = gSL_Select & "                                FROM 材料購入単価マスタ ST1S"
    gSL_Select = gSL_Select & "                                     LEFT OUTER JOIN 材料マスタ ZM"
    gSL_Select = gSL_Select & "                                          ON ZM.材料管理番号 = ST1S.材料管理番号"
    gSL_Select = gSL_Select & "                               WHERE ST1S.取引先 = ST.取引先"
    gSL_Select = gSL_Select & "                                 AND ST1S.契約先 = ST.契約先"
    gSL_Select = gSL_Select & "                                 AND ST1S.材料管理番号 = ST.材料管理番号"
    gSL_Select = gSL_Select & "                               GROUP BY ST1S.取引先, ST1S.契約先, ST1S.材料管理番号)"
    gSL_Select = gSL_Select & CisFun.Where_Phrase
    CisDB.SQL = gSL_Select
    If Not CisDB.DBExec Then Exit Function
    
    ZTWCreate = True

End Function

'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------------------+
'+                                  +
'+          追加･更新 処理          +
'+                                  +
'+----------------------------------+
Private Function DBPut()
    DBPut = False
    Dim RirekiNo  As Integer

    K_Sykbnm.Caption = SyoriNM
'    K_Sykbnm.Visible = True        ' 2009.10.13 del
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     追加・更新処理を行います      "
        .MB_Title = "追加・更新処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Call VsGridToText(Val(Input_Area.Tag))
            Call GridTextEnabled(True)
            Exit Function
        End If
    End With
    On Error GoTo DBPut_Err
    Message.Visible = True
    DoEvents
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
'   マスタ項目移送
    For gIndex = 1 To VsGrid.Rows - 1
        If RTrim(VsGrid.TextMatrix(gIndex, CisVsGrid.SelectCol)) <> "" Then
           SyoriKB = "ADD"
           gSL_Select = "SELECT MAX(履歴NO)履歴NO FROM 購入単価マスタ"
           gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
           gSL_Select = gSL_Select & "   AND 契約先 = '" & RTrim(iH1_KTorcd) & "'"
           gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(RTrim(VsGrid.TextMatrix(gIndex, 16))) & "'"
           If KTMRead(gSL_Select, 1) Then
              RirekiNo = KTM.履歴NO + 1
           Else
              RirekiNo = 1
           End If
            
           Call ItemsClearKTM
           With KTM
               '===== 取引先 =====
               .取引先 = RTrim(iH1_Torcd)
               '===== 契約先 =====
               .契約先 = RTrim(iH1_KTorcd)
               '====== 品番 ======
               .品番 = RTrim(VsGrid.TextMatrix(gIndex, 16))
               '===== 履歴NO =====
               .履歴NO = RirekiNo
               If SyoriKB = "ADD" Then
                  '=== 適用期間 ===
                  .適用開始日付 = RTrim(iH1_SYmd)
                  .適用終了日付 = RTrim(iH1_EYmd)
               End If
               '====== 単価 ======
               .単価 = CisFun.Val2(VsGrid.TextMatrix(gIndex, 6))
               '===== 作成者 =====
               .作成者 = gTanto
               '=== 仮単価区分 ===
               If RTrim(VsGrid.TextMatrix(gIndex, 7)) = "正" Then
                  .仮単価区分 = ""
               Else
                  .仮単価区分 = "1"
               End If
'-------------------------------------------------------------------------------' 2009.10.13 del start
'               '=== 部品有無区分 ===
'               If RTrim(VsGrid.TextMatrix(gIndex, 8)) = "有" Then
'                  .部品有無区分 = 0
'               Else
'                  .部品有無区分 = 1
'               End If
'               '=== 部品自支給区分 ===
'               If RTrim(VsGrid.TextMatrix(gIndex, 9)) = "自" Then
'                  .部品自支給区分 = 0
'               Else
'                  .部品自支給区分 = 1
'               End If
'               '====== 部品単価 ======
'               .部品単価 = CisFun.Val2(VsGrid.TextMatrix(gIndex, 10))
'               '=== 材料有無区分 ===
'               If RTrim(VsGrid.TextMatrix(gIndex, 11)) = "有" Then
'                  .材料有無区分 = 0
'               Else
'                  .材料有無区分 = 1
'               End If
'               '=== 材料自支給区分 ===
'               If RTrim(VsGrid.TextMatrix(gIndex, 12)) = "自" Then
'                  .材料自支給区分 = 0
'               Else
'                  .材料自支給区分 = 1
'               End If
'-------------------------------------------------------------------------------' 2009.10.13 del end
               '====== 材料単価 ======
               .部品単価 = CisFun.Val2(VsGrid.TextMatrix(gIndex, 8))            ' 2009.10.13 upd : 材料単価-->部品単価 および 位置
               '====== 加工単価 ======
               .加工単価 = CisFun.Val2(VsGrid.TextMatrix(gIndex, 9))            ' 2009.10.13 upd : 位置
               '====== その他単価 ======
               .その他単価 = CisFun.Val2(VsGrid.TextMatrix(gIndex, 10))         ' 2009.10.13 add
           End With
           
       ' ( 追加処理 )
           Call KTMInsert
           iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
            
        End If
    Next gIndex
'
' ( 追加処理 )
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    Message.Visible = False
    DoEvents
    DBPut = True
    Exit Function
DBPut_Err:
    Message.Visible = False
    DoEvents
    CisFun.ErrorBox
    End
End Function

'+--------------------------------------+
'+                                      +
'+         追加･更新 処理(材料)         +
'+                                      +
'+--------------------------------------+
Private Function DBPutZ()
    DBPutZ = False
    Dim RirekiNo  As Integer

    K_Sykbnm.Caption = SyoriNM
'    K_Sykbnm.Visible = True        ' 2009.10.13 del
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     追加・更新処理を行います      "
        .MB_Title = "追加・更新処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Call VsGridToText(Val(Input_Area.Tag))
            Call GridTextEnabled(True)
            Exit Function
        End If
    End With
    On Error GoTo DBPut_Err
    Message.Visible = True
    DoEvents
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
'   マスタ項目移送
    For gIndex = 1 To VsGrid.Rows - 1
        If RTrim(VsGrid.TextMatrix(gIndex, CisVsGrid.SelectCol)) <> "" Then
           SyoriKB = "ADD"
           gSL_Select = "SELECT MAX(履歴NO)履歴NO FROM 材料購入単価マスタ"
           gSL_Select = gSL_Select & " WHERE 取引先 = '" & RTrim(iH1_Torcd) & "'"
           gSL_Select = gSL_Select & "   AND 契約先 = '" & RTrim(iH1_KTorcd) & "'"
           gSL_Select = gSL_Select & "   AND 材料管理番号 = '" & RTrim(VsGrid.TextMatrix(gIndex, 30)) & "'"
           If ZKTRead(gSL_Select, 1) Then
              RirekiNo = ZKT.履歴NO + 1
           Else
              RirekiNo = 1
           End If
            
           Call ItemsClearZKT
           With ZKT
               '===== 取引先 =====
               .取引先 = RTrim(iH1_Torcd)
               '===== 契約先 =====
               .契約先 = RTrim(iH1_KTorcd)
               '===== 履歴NO =====
               .履歴NO = RirekiNo
               '=== 適用期間 ===
               .適用開始日付 = RTrim(iH1_SYmd)
               .適用終了日付 = RTrim(iH1_EYmd)
               '====== 単価 ======
               .単価 = CisFun.Val2(VsGrid.TextMatrix(gIndex, 6))
               '===== 作成者 =====
               .作成者 = gTanto
               '=== 仮単価区分 ===
               If RTrim(VsGrid.TextMatrix(gIndex, 7)) = "正" Then
                  .仮単価区分 = ""
               Else
                  .仮単価区分 = "1"
               End If
'-----------------------------------------------------------------------' 2009.10.13 del start
'               '=== 部品有無区分 ===
'               If RTrim(VsGrid.TextMatrix(gIndex, 8)) = "有" Then
'                  .部品有無区分 = 0
'               Else
'                  .部品有無区分 = 1
'               End If
'               '=== 部品自支給区分 ===
'               If RTrim(VsGrid.TextMatrix(gIndex, 9)) = "自" Then
'                  .部品自支給区分 = 0
'               Else
'                  .部品自支給区分 = 1
'               End If
'               '====== 部品単価 ======
'               .部品単価 = CisFun.Val2(VsGrid.TextMatrix(gIndex, 10))
'               '=== 材料有無区分 ===
'               If RTrim(VsGrid.TextMatrix(gIndex, 11)) = "有" Then
'                  .材料有無区分 = 0
'               Else
'                  .材料有無区分 = 1
'               End If
'               '=== 材料自支給区分 ===
'               If RTrim(VsGrid.TextMatrix(gIndex, 12)) = "自" Then
'                  .材料自支給区分 = 0
'               Else
'                  .材料自支給区分 = 1
'               End If
'-----------------------------------------------------------------------' 2009.10.13 del end
               '====== 材料単価 ======
               .部品単価 = CisFun.Val2(VsGrid.TextMatrix(gIndex, 8))    ' 2009.10.13 upd : 材料単価-->部品単価 および 位置
               '====== 加工単価 ======
               .加工単価 = CisFun.Val2(VsGrid.TextMatrix(gIndex, 9))    ' 2009.10.13 upd : 位置
               '====== その他単価 ======
               .その他単価 = CisFun.Val2(VsGrid.TextMatrix(gIndex, 10)) ' 2009.10.13 add
               '==== 材料管理番号 ====
               .材料管理番号 = RTrim(VsGrid.TextMatrix(gIndex, 30))
           End With
           
       ' ( 追加処理 )
           Call ZKTInsert
           
            
       ' ( 影響データ自動更新 )
'           Call ZTanAutoUpdate(gIndex)     ' 2009.10.13 del
            
        End If
    Next gIndex
'
' ( 追加処理 )
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    Message.Visible = False
    DoEvents
    DBPutZ = True
    Exit Function
DBPut_Err:
    Message.Visible = False
    DoEvents
    CisFun.ErrorBox
    End
End Function
'-----------------------------------------------------------------------------------------------------------' 2009.10.13 del start
''+----------------------------------------------------+
''+     材料単価更新時の影響の出るデータを自動更新     +
''+----------------------------------------------------+
'Private Sub ZTanAutoUpdate(Row As Long)
'    Dim wTanka          As Currency ' 単価退避
'    Dim wTankaSa        As Currency ' 単価差退避
'
'    '**********************************
'    '*****     単価の変更確認     *****
'    '**********************************
'    gSL_Select = ""
'    gSL_Select = gSL_Select & " SELECT 単価 FROM 材料購入単価マスタ"
'    gSL_Select = gSL_Select & "  WHERE 材料管理番号 = " & CisFun.Val2(VsGrid.TextMatrix(Row, 30))
'    gSL_Select = gSL_Select & "    AND 取引先 = '" & RTrim(iH1_Torcd) & "'"
'    gSL_Select = gSL_Select & "    AND 契約先 = '" & RTrim(iH1_KTorcd) & "'"
'    gSL_Select = gSL_Select & "  ORDER BY 履歴NO DESC"
'    Call ZKTRead(gSL_Select)
'    wTanka = ZKT.単価
'    Call ZKTReadNext
'    If wTanka = ZKT.単価 Then Exit Sub
'
'    '****************************************************
'    '*****     影響のあるデータが存在するか確認     *****
'    '****************************************************
'
'    gSL_Select = ""
'    gSL_Select = gSL_Select & " SELECT KU.品番, KU.取引先, KU.契約先, KU.履歴NO" & vbCrLf
'    gSL_Select = gSL_Select & "   FROM 購入単価内訳マスタ KU" & vbCrLf
'    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT" & vbCrLf
'    gSL_Select = gSL_Select & "             ON KT.品番   = KU.品番" & vbCrLf
'    gSL_Select = gSL_Select & "            AND KT.取引先 = KU.取引先" & vbCrLf
'    gSL_Select = gSL_Select & "            AND KT.契約先 = KU.契約先" & vbCrLf
'    gSL_Select = gSL_Select & "            AND KT.履歴NO = (SELECT MAX(履歴NO) FROM 購入単価マスタ KT1" & vbCrLf
'    gSL_Select = gSL_Select & "                              WHERE KT1.品番   = KT.品番" & vbCrLf
'    gSL_Select = gSL_Select & "                                AND KT1.取引先 = KT.取引先" & vbCrLf
'    gSL_Select = gSL_Select & "                                AND KT1.契約先 = KT.契約先" & vbCrLf
'    If RTrim(iH1_SYmd) <> "" Then
'       gSL_Select = gSL_Select & "                             AND KT1.適用開始日付 <= '" & RTrim(iH1_SYmd) & "'" & vbCrLf
'    End If
'    If RTrim(iH1_EYmd) <> "" Then
'       gSL_Select = gSL_Select & "                             AND KT1.適用終了日付 >= '" & RTrim(iH1_EYmd) & "'" & vbCrLf
'    End If
'    gSL_Select = gSL_Select & "                                AND KT1.材料自支給区分 = 0)" & vbCrLf
'    gSL_Select = gSL_Select & "  WHERE KU.材料管理番号 = " & CisFun.Val2(VsGrid.TextMatrix(Row, 30)) & vbCrLf
'    gSL_Select = gSL_Select & "    AND KU.材料取引先 = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
'    gSL_Select = gSL_Select & "    AND KU.材料契約先 = '" & RTrim(iH1_KTorcd) & "'" & vbCrLf
'    gSL_Select = gSL_Select & "    AND KU.履歴NO = (SELECT MAX(履歴NO) FROM 購入単価マスタ KT2" & vbCrLf
'    gSL_Select = gSL_Select & "                      WHERE KT2.品番   = KU.品番" & vbCrLf
'    gSL_Select = gSL_Select & "                        AND KT2.取引先 = KU.取引先" & vbCrLf
'    gSL_Select = gSL_Select & "                        AND KT2.契約先 = KU.契約先" & vbCrLf
'    If RTrim(iH1_SYmd) Then
'       gSL_Select = gSL_Select & "                     AND KT2.適用開始日付 <= '" & RTrim(iH1_SYmd) & "'" & vbCrLf
'    End If
'    If RTrim(iH1_EYmd) <> "" Then
'       gSL_Select = gSL_Select & "                     AND KT2.適用終了日付 >= '" & RTrim(iH1_EYmd) & "'" & vbCrLf
'    End If
'    gSL_Select = gSL_Select & "                        AND KT2.材料自支給区分 = 0)" & vbCrLf
'    gSL_Select = gSL_Select & "  GROUP BY KU.品番, KU.取引先, KU.契約先, KU.履歴NO" & vbCrLf
'    If Not KTURead(gSL_Select) Then Exit Sub
'
'    '**********************************
'    '*****     確認メッセージ     *****
'    '**********************************
'    With CisFun
'        .MB_Lines = 9
'        .MB_MSG(2) = "品番 ＝ 【" & VsGrid.TextMatrix(Row, 1) & "】"
'        .MB_MSG(4) = "           単価が変更されました。"
'        .MB_MSG(6) = "影響のある部品単価が【 " & Format(CisDB.RecordCount, "#,###") & " 件 】存在します。"
'        .MB_MSG(8) = "     変更後の値に一括置換しますか？"
'        .MB_Title = ""
'        .MB_Button = OK_CAN
'        If Not .MBOX Then Exit Sub
'    End With
'
'    Call AutoUpdate(Row)
'
'End Sub
'
''+---------------------------------+
''+     購入単価内訳マスタ 更新     +
''+---------------------------------+
'Private Sub AutoUpdate(Row As Long)
'    Dim SetRirekiNo As Integer
'
'    Do Until Not KTU_RDSTS
'       '************************************
'       '**     セットする履歴NOの抽出     **
'       '************************************
'       gSL_Select = ""
'       gSL_Select = gSL_Select & " SELECT MAX(履歴NO) 履歴NO FROM 購入単価マスタ"
'       gSL_Select = gSL_Select & "  WHERE 品番 = '" & RTrim(KTU.品番) & "'"
'       gSL_Select = gSL_Select & "    AND 取引先 = '" & RTrim(KTU.取引先) & "'"
'       gSL_Select = gSL_Select & "    AND 契約先 = '" & RTrim(KTU.契約先) & "'"
'       Call KTMRead(gSL_Select, , 1)
'       SetRirekiNo = KTM.履歴NO + 1
'
'       '**************************************
'       '**     購入単価内訳マスタの更新     **
'       '**************************************
'       gSL_Select = ""
'       gSL_Select = gSL_Select & " INSERT INTO 購入単価内訳マスタ" & vbCrLf
'       gSL_Select = gSL_Select & "            (品番, 取引先, 契約先, 履歴NO, 単価内訳区分, 連番, 内訳コード, 内訳内容," & vbCrLf
'       gSL_Select = gSL_Select & "             材料管理番号, 材料取引先, 材料契約先, 単価, 作成日, 作成者, 作成端末)" & vbCrLf
'       gSL_Select = gSL_Select & " SELECT KU.品番, KU.取引先, KU.契約先, " & SetRirekiNo & "," & vbCrLf
'       gSL_Select = gSL_Select & "        KU.単価内訳区分, KU.連番, KU.内訳コード, KU.内訳内容," & vbCrLf
'       gSL_Select = gSL_Select & "        KU.材料管理番号, KU.材料取引先, KU.材料契約先," & vbCrLf
'       '************************ 更新該当データの場合 新単価を設定 ***************************
'       gSL_Select = gSL_Select & "        CASE WHEN KU.単価内訳区分 = 2" & vbCrLf
'       gSL_Select = gSL_Select & "              AND KU.材料管理番号 = " & CisFun.Val2(VsGrid.TextMatrix(Row, 30)) & vbCrLf
'       gSL_Select = gSL_Select & "              AND KU.材料取引先 = '" & RTrim(iH1_Torcd) & "'" & vbCrLf
'       gSL_Select = gSL_Select & "              AND KU.材料契約先 = '" & RTrim(iH1_KTorcd) & "'" & vbCrLf
'       gSL_Select = gSL_Select & "             THEN CEILING(" & CisFun.Val2(VsGrid.TextMatrix(Row, 6)) & _
'                                                            " * ZH.使用量 * 100) / 100" & vbCrLf
'       gSL_Select = gSL_Select & "             ELSE KU.単価 END,"
'       gSL_Select = gSL_Select & "        GETDATE(), '', HOST_NAME()" & vbCrLf
'       gSL_Select = gSL_Select & "   FROM 購入単価内訳マスタ KU" & vbCrLf
'       gSL_Select = gSL_Select & "        LEFT OUTER JOIN 材料品番マスタ ZH" & vbCrLf
'       gSL_Select = gSL_Select & "             ON ZH.材料管理番号 = KU.材料管理番号" & vbCrLf
'       gSL_Select = gSL_Select & "            AND ZH.使用品番 = KU.品番" & vbCrLf
'       gSL_Select = gSL_Select & "  WHERE KU.品番 = '" & RTrim(KTU.品番) & "'" & vbCrLf
'       gSL_Select = gSL_Select & "    AND KU.取引先 = '" & RTrim(KTU.取引先) & "'" & vbCrLf
'       gSL_Select = gSL_Select & "    AND KU.契約先 = '" & RTrim(KTU.契約先) & "'" & vbCrLf
'       gSL_Select = gSL_Select & "    AND KU.履歴NO = " & CisFun.Val2(KTU.履歴NO) & vbCrLf
'
'       CisDB.SQL = gSL_Select
'
'       Call CisDB.DBExec
'
'       '**********************************
'       '**     購入単価マスタの更新     **
'       '**********************************
'       gSL_Select = ""
'       gSL_Select = gSL_Select & " INSERT INTO 購入単価マスタ" & vbCrLf
'       gSL_Select = gSL_Select & "            (品番, 取引先, 契約先, 履歴NO, 仮単価区分, 単価, 有無償区分," & vbCrLf
'       gSL_Select = gSL_Select & "             適用開始日付, 適用終了日付, 部品単価, 部品自支給区分, 部品有無区分," & vbCrLf
'       gSL_Select = gSL_Select & "             材料単価, 材料自支給区分, 材料有無区分, 加工単価, その他単価," & vbCrLf
'       gSL_Select = gSL_Select & "             作成日,作成者,作成端末)" & vbCrLf
'       gSL_Select = gSL_Select & " SELECT KT.品番, KT.取引先, KT.契約先, " & SetRirekiNo & ", KT.仮単価区分,"
'       gSL_Select = gSL_Select & "        KT.部品単価 + KUS.単価 + KT.加工単価," & vbCrLf
'       gSL_Select = gSL_Select & "        KT.有無償区分, KT.適用開始日付, KT.適用終了日付," & vbCrLf
'       gSL_Select = gSL_Select & "        KT.部品単価, KT.部品自支給区分, KT.部品有無区分," & vbCrLf
'       gSL_Select = gSL_Select & "        KUS.単価, KT.材料自支給区分, KT.材料有無区分," & vbCrLf
'       gSL_Select = gSL_Select & "        KT.加工単価, KT.その他単価," & vbCrLf
'       gSL_Select = gSL_Select & "        GETDATE(), '', HOST_NAME()" & vbCrLf
'       gSL_Select = gSL_Select & "   FROM 購入単価マスタ Kt" & vbCrLf
'       gSL_Select = gSL_Select & "        LEFT OUTER JOIN (SELECT SUM(単価) 単価, 品番, 取引先, 契約先, 履歴NO" & vbCrLf
'       gSL_Select = gSL_Select & "                           FROM 購入単価内訳マスタ KU" & vbCrLf
'       gSL_Select = gSL_Select & "                          WHERE KU.品番   = '" & KTU.品番 & "'" & vbCrLf
'       gSL_Select = gSL_Select & "                            AND KU.取引先 = '" & KTU.取引先 & "'" & vbCrLf
'       gSL_Select = gSL_Select & "                            AND KU.契約先 = '" & KTU.契約先 & "'" & vbCrLf
'       gSL_Select = gSL_Select & "                            AND KU.履歴NO = " & SetRirekiNo & vbCrLf
'       gSL_Select = gSL_Select & "                            AND KU.単価内訳区分 = 2" & vbCrLf
'       gSL_Select = gSL_Select & "                          GROUP BY 品番, 取引先, 契約先, 履歴NO) KUS" & vbCrLf
'       gSL_Select = gSL_Select & "             ON KUS.品番 = KT.品番" & vbCrLf
'       gSL_Select = gSL_Select & "            AND KUS.取引先 = KT.取引先" & vbCrLf
'       gSL_Select = gSL_Select & "            AND KUS.契約先 = KT.契約先" & vbCrLf
'       gSL_Select = gSL_Select & "            AND KUS.履歴NO - 1 = KT.履歴NO" & vbCrLf
'       gSL_Select = gSL_Select & "  WHERE KT.品番 = '" & RTrim(KTU.品番) & "'" & vbCrLf
'       gSL_Select = gSL_Select & "    AND KT.取引先 = '" & RTrim(KTU.取引先) & "'" & vbCrLf
'       gSL_Select = gSL_Select & "    AND KT.契約先 = '" & RTrim(KTU.契約先) & "'" & vbCrLf
'       gSL_Select = gSL_Select & "    AND KT.履歴NO = " & CisFun.Val2(KTU.履歴NO) & vbCrLf
'
'       CisDB.SQL = gSL_Select
'
'       Call CisDB.DBExec
'
'       Call KTUReadNext(0)
'    Loop
'
'End Sub
'-----------------------------------------------------------------------------------------------------------' 2009.10.13 del end

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
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")                '-------' 2009.10.09 udp start
        Call .FixedSet(中中, 左中, あり, 49, 1, "品  番")   ' 20-->49
        Call .FixedSet(中中, 右中, あり, 8, 2, "前回開始")
        Call .FixedSet(中中, 右中, あり, 8, 3, "前回終了")
        Call .FixedSet(中中, 中中, あり, 2, 4, " 仮 ")
        Call .FixedSet(中中, 右中, あり, 12, 5, "前回単価")
        Call .FixedSet(中中, 右中, あり, 12, 6, "単価")
        Call .FixedSet(中中, 中中, あり, 3, 7, "仮")
'        Call .FixedSet(中中, 中中, あり, 3, 8, " 有 ")
'        Call .FixedSet(中中, 中中, あり, 3, 9, " 自 ")
'        Call .FixedSet(中中, 右中, あり, 12, 10, "部品単価")
'        Call .FixedSet(中中, 中中, あり, 3, 11, "有")
'        Call .FixedSet(中中, 中中, あり, 3, 12, "自")
        Call .FixedSet(中中, 右中, あり, 12, 8, "部品単価")
        Call .FixedSet(中中, 右中, あり, 12, 9, "加工単価")
        Call .FixedSet(中中, 中中, あり, 12, 10, "その他単価")
        Call .FixedSet(中中, 中中, なし, 0, 11, "")
        Call .FixedSet(中中, 中中, なし, 0, 12, "")
        Call .FixedSet(中中, 中中, なし, 0, 13, "")
        Call .FixedSet(中中, 中中, なし, 0, 14, "")
        Call .FixedSet(中中, 中中, なし, 0, 15, "")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, あり, 0, 16, "X品  番")
        Call .FixedSet(中中, 中中, あり, 0, 17, "X前回開始")
        Call .FixedSet(中中, 中中, あり, 0, 18, "X前回終了")
        Call .FixedSet(中中, 中中, あり, 0, 19, " X仮 ")
        Call .FixedSet(中中, 右中, あり, 0, 20, "X前回単価")
        Call .FixedSet(中中, 右中, あり, 0, 21, "X単価")
        Call .FixedSet(中中, 中中, あり, 0, 22, "X仮")
        Call .FixedSet(左中, 中中, あり, 0, 23, "X有無 ")
        Call .FixedSet(左中, 中中, あり, 0, 24, "X自支 ")
        Call .FixedSet(中中, 右中, あり, 0, 25, "X部品単価")
        Call .FixedSet(中中, 中中, あり, 0, 26, "X有無")
        Call .FixedSet(中中, 中中, あり, 0, 27, "X自支")
        Call .FixedSet(中中, 右中, あり, 0, 28, "X材料単価")
        Call .FixedSet(中中, 右中, あり, 0, 29, "X加工単価")
        Call .FixedSet(中中, 左中, なし, 0, 30, "材料管理番号")
        Call .FixedSet(中中, 左中, なし, 0, 31, "処理")
        Call .FixedSet(中中, 中中, あり, 0, 32, "履歴NO")
        Call .FixedSet(中中, 左中, なし, 0, 33, "Select")
        Call .FixedSet(中中, 左中, あり, 0, 34, "X発注品番")
        Call .FixedSet(中中, 左中, あり, 0, 35, "X受給品番")
        Call .FixedSet(中中, 左中, あり, 0, 36, "X自社品番")
        Call .FixedSet(中中, 左中, あり, 0, 37, "X材質")
        Call .FixedSet(中中, 左中, あり, 0, 38, "X板厚")
        Call .FixedSet(中中, 左中, あり, 0, 39, "X幅")
        Call .FixedSet(中中, 左中, あり, 0, 40, "X長さ")
        Call .FixedSet(中中, 左中, あり, 0, 41, "X材質寸法")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10")  ' 8,9,10,11,12, 削除
        .InitDisp                                                   '-------' 2009.10.09 udp end
    End With
    
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid
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
        With VsGrid
            .TextMatrix(ID, GridCol(0)) = ID
                        
            '============== 品番 ===============
            If H0Op_Tanku(0).Value = True Then                      ' 2009.10.13 add
               .TextMatrix(ID, GridCol(1)) = RTrim(KTM.表示品番)
               .TextMatrix(ID, GridCol(16)) = RTrim(KTM.品番)
            Else                                                    ' 2009.10.13 add
               .TextMatrix(ID, GridCol(1)) = Trim(KTM.材質) & _
                        " " & Format(KTM.板厚, "#,##0.00") & _
                        "x" & Format(KTM.幅, "#,##0.00") & _
                        "x" & Format(KTM.長さ, "#,##0.00")          ' 2009.10.13 add
            End If                                                  ' 2009.10.13 add
            '========== 前回適用期間 ===========
            If KTM.適用開始日付 <> "" Then
               gStr = ""
               gStr = gStr & CisFun.Mid2(KTM.適用開始日付, 1, 4) & "/"
               gStr = gStr & CisFun.Mid2(KTM.適用開始日付, 5, 2) & "/"
               gStr = gStr & CisFun.Mid2(KTM.適用開始日付, 7, 2)
               .TextMatrix(ID, GridCol(2)) = gStr
            End If
            
            If KTM.適用終了日付 <> "" Then
               gStr = ""
               gStr = gStr & CisFun.Mid2(KTM.適用終了日付, 1, 4) & "/"
               gStr = gStr & CisFun.Mid2(KTM.適用終了日付, 5, 2) & "/"
               gStr = gStr & CisFun.Mid2(KTM.適用終了日付, 7, 2)
               .TextMatrix(ID, GridCol(3)) = gStr
            End If
            
            '========= 前回仮単価区分 ==========
            If KTM.前回単価 <> 0 Then
                If KTM.前回仮単価区分 = "" Then
                    .TextMatrix(ID, GridCol(4)) = "正"
                Else
                    .TextMatrix(ID, GridCol(4)) = "仮"
                End If
            End If
            
            '============ 前回単価 =============
            If KTM.前回単価 <> 0 Then
                .TextMatrix(ID, GridCol(5)) = Format(KTM.前回単価, "#,##0.00")
            End If
            
            If OpH1_TanKb(1).Value Then
                '============ 単価 =============
                If KTM.前回単価 <> 0 Then
                    .TextMatrix(ID, GridCol(6)) = Format(KTM.前回単価, "#,##0.00")
                    '====== 仮単価区分 ======
                    If KTM.仮単価区分 = "" Then
                        .TextMatrix(ID, GridCol(7)) = "正"
                        .TextMatrix(ID, GridCol(22)) = "正"
                    Else
                        .TextMatrix(ID, GridCol(7)) = "仮"
                        .TextMatrix(ID, GridCol(22)) = "仮"
                    End If
'-----------------------------------------------------------------------------------' 2009.10.13 del end
'                    '====== 有無区分 ======
'                    If KTM.有無区分 = 0 Then
'                        .TextMatrix(ID, GridCol(8)) = "有"
'                        .TextMatrix(ID, GridCol(23)) = "有"
'                    Else
'                        .TextMatrix(ID, GridCol(8)) = "無"
'                        .TextMatrix(ID, GridCol(23)) = "無"
'                    End If
'                    '====== 自支給区分 ======
'                    If KTM.有無区分 = 0 Then
'                        .TextMatrix(ID, GridCol(9)) = "自"
'                        .TextMatrix(ID, GridCol(24)) = "自"
'                    Else
'                        .TextMatrix(ID, GridCol(9)) = "支"
'                        .TextMatrix(ID, GridCol(24)) = "支"
'                    End If
'-----------------------------------------------------------------------------------' 2009.10.13 del end
                End If
            Else
                .TextMatrix(ID, GridCol(6)) = Format(KTM.単価, "#,##0.00;;#")
            End If
            
            .TextMatrix(ID, GridCol(32)) = RTrim(KTM.履歴NO)
            
            '========== 材料管理番号 ===========
            .TextMatrix(ID, GridCol(30)) = RTrim(KTM.材料管理番号)
                        
'-----------------------------------------------------------------------------------' 2009.10.13 add start
            '========== 材料品番情報 ===========
            .TextMatrix(ID, GridCol(34)) = RTrim(KTM.発注品番)
            .TextMatrix(ID, GridCol(35)) = RTrim(KTM.受給品番)
            .TextMatrix(ID, GridCol(36)) = RTrim(KTM.自社品番)
            .TextMatrix(ID, GridCol(41)) = RTrim(KTM.材質) & Space(25 - Len(Trim(KTM.材質))) & "/" & _
                                           Format(KTM.板厚, "0000.00") & "x" & _
                                           Format(KTM.幅, "0000.00") & "x" & _
                                           Format(KTM.長さ, "000.00")
'-----------------------------------------------------------------------------------' 2009.10.13 add end
        End With
        Call KTMReadNext(1)
    Loop

    With VsGrid
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    Call KTMClose(1)
    
    Call VsGridToHnbn(1)    ' 2009.10.13 add
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
    gInt = VsGrid.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > VsGrid.Rows - 1 Then
            GoTo Grid_RowChange_End
        Else
            If RTrim(VsGrid.TextMatrix(gInt + 1, 1)) = "" Then
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
        If gInt + CisVsGrid.DispMax > VsGrid.Rows - 1 Then
            gInt = VsGrid.Rows - 1
        Else
            gInt = gInt + CisVsGrid.DispMax
        End If
    End If
    VsGrid.Row = gInt
'グリッドからテキストに表示
    Call VsGridToText(VsGrid.Row)
    Call VsGridToHnbn(VsGrid.Row)       ' 2009.10.13
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    With VsGrid
        '======== 単価 ========
        If CisFun.Val2(iB1_Tan) <> 0 Then
            .TextMatrix(tRow, 6) = Format(CisFun.Val2(iB1_Tan), "#,##0.00")
        Else
            .TextMatrix(tRow, 6) = ""
        End If
        
        '===== 仮単価区分 =====
        If CisFun.Val2(iB1_Tan) <> 0 Then
            If B1Btn_Kari = "正" Then
                .TextMatrix(tRow, 7) = "正"
            Else
                .TextMatrix(tRow, 7) = "仮"
            End If
        Else
                .TextMatrix(tRow, 7) = ""
        End If

'-----------------------------------------------------------------------------------' 2009.10.13 del start
'        '===== 部品有無区分 =====
'        If CisFun.Val2(iB1_BTan) <> 0 Then
'            If B1Btn_Bum = "有" Then
'                .TextMatrix(tRow, 8) = "有"
'            Else
'                .TextMatrix(tRow, 8) = "無"
'            End If
'        Else
'                .TextMatrix(tRow, 8) = ""
'        End If
'
'        '===== 部品自支給区分 =====
'        If CisFun.Val2(iB1_BTan) <> 0 Then
'            If B1Btn_Bjs = "自" Then
'                .TextMatrix(tRow, 9) = "自"
'            Else
'                .TextMatrix(tRow, 9) = "支"
'            End If
'        Else
'                .TextMatrix(tRow, 9) = ""
'        End If
'
'        '======== 部品単価 ========
'        If CisFun.Val2(iB1_BTan) <> 0 Then
'            .TextMatrix(tRow, 10) = Format(CisFun.Val2(iB1_BTan), "#,##0.00")
'        Else
'            .TextMatrix(tRow, 10) = ""
'        End If
'
'        '===== 材料有無区分 =====
'        If CisFun.Val2(iB1_ZTan) <> 0 Then
'            If B1Btn_Zum = "有" Then
'                .TextMatrix(tRow, 11) = "有"
'            Else
'                .TextMatrix(tRow, 11) = "無"
'            End If
'        Else
'                .TextMatrix(tRow, 11) = ""
'        End If
'
'        '===== 材料自支給区分 =====
'        If CisFun.Val2(iB1_ZTan) <> 0 Then
'            If B1Btn_Zjs = "自" Then
'                .TextMatrix(tRow, 12) = "自"
'            Else
'                .TextMatrix(tRow, 12) = "支"
'            End If
'        Else
'                .TextMatrix(tRow, 12) = ""
'        End If
'-----------------------------------------------------------------------------------' 2009.10.13 del end
        
        '======== 材料単価 ========
        If CisFun.Val2(iB1_ZTan) <> 0 Then
            .TextMatrix(tRow, 8) = Format(CisFun.Val2(iB1_ZTan), "#,##0.00")        ' 2009.10.13 upd : (tRow, 13) --> (tRow, 8)
        Else
            .TextMatrix(tRow, 8) = ""                                               ' 2009.10.13 upd : (tRow, 13) --> (tRow, 8)
        End If
        
        '======== 加工単価 ========
        If CisFun.Val2(iB1_CTan) <> 0 Then
            .TextMatrix(tRow, 9) = Format(CisFun.Val2(iB1_CTan), "#,##0.00")        ' 2009.10.13 upd : (tRow, 14) --> (tRow, 9)
        Else
            .TextMatrix(tRow, 9) = ""                                               ' 2009.10.13 upd : (tRow, 14) --> (tRow, 9)
        End If
        
        '======= その他単価 =======
        If CisFun.Val2(iB1_STan) <> 0 Then
            .TextMatrix(tRow, 10) = Format(CisFun.Val2(iB1_STan), "#,##0.00")       ' 2009.10.13 upd : (tRow, 15) --> (tRow, 10)
        Else
            .TextMatrix(tRow, 10) = ""                                              ' 2009.10.13 upd : (tRow, 15) --> (tRow, 10)
        End If
        
    End With
    SaveRow = VsGrid.Row
    Call UpdateCheck(tRow)
    VsGrid.Row = SaveRow
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With VsGrid
        Input_Area.Tag = tRow
        Input_Area.Move Input_Area.Left, VsGrid.CellTop, Input_Area.Width, VsGrid.CellHeight
        '単価
        If CisFun.Val2(.TextMatrix(tRow, 6)) <> 0 Then
            iB1_Tan = Format(CisFun.Val2(.TextMatrix(tRow, 6)), "#,##0.00")
        Else
            iB1_Tan = ""
        End If
        
'-----------------------------------------------------------------------------------' 2009.10.28 upd start : 常時"正"とする
'        '仮区分
'        If RTrim(.TextMatrix(tRow, 7)) = "正" Then
            B1Btn_Kari = "正"
'        Else
'            B1Btn_Kari = "仮"
'        End If
'-----------------------------------------------------------------------------------' 2009.10.28 upd end
        
'-----------------------------------------------------------------------------------' 2009.10.13 del start
'        ' 部品　有無区分
'        If RTrim(.TextMatrix(tRow, 8)) = "無" Then
'            B1Btn_Bum = "無"
'        Else
'            B1Btn_Bum = "有"
'        End If
'
'        ' 部品　自支給区分
'        If RTrim(.TextMatrix(tRow, 9)) = "支" Then
'            B1Btn_Bjs = "支"
'        Else
'            B1Btn_Bjs = "自"
'        End If
'
'        '部品単価
'        If CisFun.Val2(.TextMatrix(tRow, 10)) <> 0 Then
'            iB1_BTan = Format(CisFun.Val2(.TextMatrix(tRow, 10)), "#,##0.00")
'        Else
'            iB1_BTan = ""
'        End If
'
'        ' 材料　有無区分
'        If RTrim(.TextMatrix(tRow, 11)) = "無" Then
'            B1Btn_Zum = "無"
'        Else
'            B1Btn_Zum = "有"
'        End If
'
'        ' 材料　自支給区分
'        If RTrim(.TextMatrix(tRow, 12)) = "支" Then
'            B1Btn_Zjs = "支"
'        Else
'            B1Btn_Zjs = "自"
'        End If
'-----------------------------------------------------------------------------------' 2009.10.13 del end
        
        '材料単価
        If CisFun.Val2(.TextMatrix(tRow, 8)) <> 0 Then
            iB1_ZTan = Format(CisFun.Val2(.TextMatrix(tRow, 8)), "#,##0.00")
        Else
            iB1_ZTan = ""
        End If
        
        '加工単価
        If CisFun.Val2(.TextMatrix(tRow, 9)) <> 0 Then
            iB1_CTan = Format(CisFun.Val2(.TextMatrix(tRow, 9)), "#,##0.00")
        Else
            iB1_CTan = ""
        End If
        
        'その他単価
        If CisFun.Val2(.TextMatrix(tRow, 10)) <> 0 Then
            iB1_STan = Format(CisFun.Val2(.TextMatrix(tRow, 10)), "#,##0.00")
        Else
            iB1_STan = ""
        End If
    End With
End Sub
Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call VsGridToText(VsGrid.Row)
End Sub
Private Sub VSGrid_Click()
    Call TextToVsGrid(Val(Input_Area.Tag))
    If VsGrid.TextMatrix(VsGrid.Row, 0) = "" Then Exit Sub
    Call GridTextEnabled(True)
    Call VsGridToText(VsGrid.Row)
    Call VsGridToHnbn(VsGrid.Row)       ' 2009.10.13
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
B1Btn_Kari.Visible = True
iB1_ZTan.Visible = True
iB1_CTan.Visible = True
iB1_STan.Visible = True
End Sub
Private Sub VSGrid_DblClick()
    If RTrim(VsGrid.TextMatrix(VsGrid.Row, 1)) = "" Then Exit Sub
    If VsGrid.Col = 1 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     " & VsGrid.Row & " 行目をクリアしますか？      "
            .MB_Title = ""
            .MB_Button = OK_CAN
            If Not .MBOX Then Exit Sub
        End With
        iB1_Tan = ""
        B1Btn_Kari = "仮"
        Call TextToVsGrid(Val(Input_Area.Tag))
    End If
    If VsGrid.Col = 4 And CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 5)) <> 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     " & VsGrid.Row & " 行目 前回単価を単価にコピーしますか？      "
            .MB_Title = ""
            .MB_Button = OK_CAN
            If Not .MBOX Then Exit Sub
        End With
        iB1_Tan = VsGrid.TextMatrix(VsGrid.Row, 5)
        If VsGrid.TextMatrix(VsGrid.Row, 4) = "正" Then
            B1Btn_Kari = "正"
        Else
            B1Btn_Kari = "仮"
        End If
        Call TextToVsGrid(Val(Input_Area.Tag))
    End If

B1Btn_Kari.Visible = True
iB1_ZTan.Visible = True
iB1_CTan.Visible = True
iB1_STan.Visible = True

End Sub

'+---------------------------------------------+
'+      変更チェック
'+---------------------------------------------+
Private Sub UpdateCheck(sRow As Integer)
    VsGrid.Row = sRow
    If CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 21)) = 0 Then  ' 元単価が０の場合追加
        If CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 6)) <> 0 Or _
           CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 10)) <> 0 Or _
           CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 13)) <> 0 Or _
           CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 14)) <> 0 Then     ' 単価が０ではない場合
            If RTrim(VsGrid.TextMatrix(VsGrid.Row, CisVsGrid.SelectCol)) = "" Then      ' グリッド選択されていない場合
                VsGrid.TextMatrix(VsGrid.Row, 31) = "追加"
                iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
                Call CisVsGrid.vsColor(個別, &HFFFFC0)
            End If
        Else                                                            ' 入力単価が０の場合
            If RTrim(VsGrid.TextMatrix(VsGrid.Row, CisVsGrid.SelectCol)) <> "" Then     ' グリッド選択されている場合
                VsGrid.TextMatrix(VsGrid.Row, 31) = ""
                iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) - 1, "###")
                Call CisVsGrid.vsColor(消去個別)
            End If
        End If
    Else                                                        ' 元単価が０ではない場合修正
        If CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 6)) <> CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 21)) Or _
           RTrim(VsGrid.TextMatrix(VsGrid.Row, 7)) <> RTrim(VsGrid.TextMatrix(VsGrid.Row, 22)) Or _
           RTrim(VsGrid.TextMatrix(VsGrid.Row, 8)) <> RTrim(VsGrid.TextMatrix(VsGrid.Row, 23)) Or _
           RTrim(VsGrid.TextMatrix(VsGrid.Row, 9)) <> RTrim(VsGrid.TextMatrix(VsGrid.Row, 24)) Or _
           CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 10)) <> CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 25)) Or _
           RTrim(VsGrid.TextMatrix(VsGrid.Row, 11)) <> RTrim(VsGrid.TextMatrix(VsGrid.Row, 26)) Or _
           RTrim(VsGrid.TextMatrix(VsGrid.Row, 12)) <> RTrim(VsGrid.TextMatrix(VsGrid.Row, 27)) Or _
           CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 13)) <> CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 28)) Or _
           CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 14)) <> CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 29)) Then
            If RTrim(VsGrid.TextMatrix(VsGrid.Row, CisVsGrid.SelectCol)) = "" Then
                VsGrid.TextMatrix(VsGrid.Row, 31) = "修正"
                iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
                Call CisVsGrid.vsColor(個別, &HFFFFC0)
            End If
        Else
            If RTrim(VsGrid.TextMatrix(VsGrid.Row, CisVsGrid.SelectCol)) <> "" Then
                VsGrid.TextMatrix(VsGrid.Row, 31) = "*"
                iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) - 1, "###")
                Call CisVsGrid.vsColor(消去個別)
            End If
        End If
    End If
End Sub
                
'-----------------------------------------------------------------------------------------------' 2009.10.09 add start
Private Sub VsGridToHnbn(iRow As Integer)
    If H0Op_Tanku(1).Value = False Then Exit Sub
    'Dim c As Integer: For c = 0 To VsGrid.Cols - 1: H1lb_HinBn(1) = VsGrid.TextMatrix(iRow, c): Next
    
    H1lb_Hinbn(0) = VsGrid.TextMatrix(iRow, 34)
    H1lb_Hinbn(1) = VsGrid.TextMatrix(iRow, 35)
    H1lb_Hinbn(2) = VsGrid.TextMatrix(iRow, 36)
    H1_ZaiRyo.Caption = VsGrid.TextMatrix(iRow, 41)
    H1lb_Hinbn(3) = Trim(Mid(H1_ZaiRyo.Caption, 1, 25)) & " " & _
                    Format(Trim(Mid(H1_ZaiRyo.Caption, 27, 7)), "###0.00") & "x" & _
                    Format(Trim(Mid(H1_ZaiRyo.Caption, 35, 7)), "###0.00") & "x" & _
                    Format(Trim(Mid(H1_ZaiRyo.Caption, 43, 7)), "###0.00")

'    Call HinbnToLabel
    If H1Op_Hinku(0).Value = True Then
       iH1_Hinbn.Text = H1lb_Hinbn(0).Caption
       b_HinNM(1).Top = Hinku_top(1)
       H1lb_Hinbn(1).Top = Hinku_top(1)
       b_HinNM(2).Top = Hinku_top(2)
       H1lb_Hinbn(2).Top = Hinku_top(2)
       b_HinNM(3).Top = Hinku_top(3)
       H1lb_Hinbn(3).Top = Hinku_top(3)
         
    ElseIf H1Op_Hinku(1).Value = True Then
       iH1_Hinbn.Text = H1lb_Hinbn(1).Caption
       b_HinNM(0).Top = Hinku_top(1)
       H1lb_Hinbn(0).Top = Hinku_top(1)
       b_HinNM(2).Top = Hinku_top(2)
       H1lb_Hinbn(2).Top = Hinku_top(2)
       b_HinNM(3).Top = Hinku_top(3)
       H1lb_Hinbn(3).Top = Hinku_top(3)
          
    ElseIf H1Op_Hinku(2).Value = True Then
       iH1_Hinbn.Text = H1lb_Hinbn(2).Caption
       b_HinNM(0).Top = Hinku_top(1)
       H1lb_Hinbn(0).Top = Hinku_top(1)
       b_HinNM(1).Top = Hinku_top(2)
       H1lb_Hinbn(1).Top = Hinku_top(2)
       b_HinNM(3).Top = Hinku_top(3)
       H1lb_Hinbn(3).Top = Hinku_top(3)

    Else
       iH1_Hinbn.Text = H1lb_Hinbn(3).Caption
       b_HinNM(0).Top = Hinku_top(1)
       H1lb_Hinbn(0).Top = Hinku_top(1)
       b_HinNM(1).Top = Hinku_top(2)
       H1lb_Hinbn(1).Top = Hinku_top(2)
       b_HinNM(2).Top = Hinku_top(3)
       H1lb_Hinbn(2).Top = Hinku_top(3)
       Call H1_ZaiBox_Set
    End If
    
    If H1Op_Hinku(3).Value = True Then
       iH1_Hinbn.Visible = False
       H1_ZaiBox.Top = iH1_Hinbn.Top
       H1_ZaiBox.Visible = True
       iH1_Zaist.Text = Trim(Mid(H1_ZaiRyo.Caption, 1, 25))
       iH1_Itatu.Text = Trim(Mid(H1_ZaiRyo.Caption, 27, 7))
       iH1_Width.Text = Trim(Mid(H1_ZaiRyo.Caption, 35, 7))
       iH1_Long.Text = Trim(Mid(H1_ZaiRyo.Caption, 43, 7))
    Else
       H1_ZaiBox.Visible = False
       iH1_Hinbn.Visible = True
       For cnt = 0 To 2
           If H1Op_Hinku(cnt).Value = True Then
              iH1_Hinbn.Text = H1lb_Hinbn(cnt)
              Exit For
           End If
       Next
    End If

    H1lb_ZaiKan.Caption = VsGrid.TextMatrix(iRow, 30)

End Sub
'-----------------------------------------------------------------------------------------------' 2009.10.09 add end

''+-------------------------------------+
''+      Debug用 SQL記録                +
''+-------------------------------------+
Private Sub Test_LogOutput(str As String, Msg As Integer)
Exit Sub   '=========================================== 実行しないときはこの行を生かすこと =========================
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


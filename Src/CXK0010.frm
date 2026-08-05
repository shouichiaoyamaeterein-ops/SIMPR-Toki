VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{1BA4D453-0150-11CE-89B0-0000C037528B}#3.2#0"; "PDQCom32.OCX"
Begin VB.Form CXK0010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "検収入力"
   ClientHeight    =   11115
   ClientLeft      =   90
   ClientTop       =   1140
   ClientWidth     =   15240
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
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  '最大化
   Begin VB.PictureBox H3_Area1 
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
      Height          =   1275
      Left            =   3315
      ScaleHeight     =   1215
      ScaleWidth      =   7080
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1455
      Width           =   7140
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   1110
         Left            =   60
         Top             =   60
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "      出 荷 日    - 便"
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
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   345
            Index           =   1
            Left            =   90
            Top             =   345
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   609
            ForeColor       =   0
            Caption         =   "開始"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   345
            Index           =   2
            Left            =   90
            Top             =   705
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   609
            ForeColor       =   0
            Caption         =   "終了"
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
            cAlingnment     =   7
         End
         Begin CisText_V60.CisText iH3_SBin 
            Height          =   375
            Left            =   2475
            TabIndex        =   2
            Top             =   330
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            cFormat         =   "00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "00"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFaZero         =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
         Begin CisText_V60.CisText iH3_EBin 
            Height          =   375
            Left            =   2475
            TabIndex        =   4
            Top             =   690
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            cFormat         =   "00"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "00"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFaZero         =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
         Begin CisYMD_With_Btn.CisYMDwB iH3_EYmd 
            Height          =   360
            Left            =   615
            TabIndex        =   3
            Top             =   690
            Width           =   1860
            _ExtentX        =   3069
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
            Object.Width           =   1860
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CisYMD_With_Btn.CisYMDwB iH3_SYmd 
            Height          =   360
            Left            =   615
            TabIndex        =   1
            Top             =   330
            Width           =   1860
            _ExtentX        =   3069
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
            Object.Width           =   1860
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   1110
         Left            =   3015
         Top             =   60
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "納 入 先"
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
         Begin CisText_V60.CisText iH3_Torcd 
            Height          =   360
            Left            =   90
            TabIndex        =   5
            Top             =   330
            Width           =   1005
            _ExtentX        =   1773
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
         Begin Cis3D_v60.CIS3D H3lb_Tornm 
            Height          =   360
            Left            =   90
            Top             =   690
            Width           =   2460
            _ExtentX        =   4339
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
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   1110
         Left            =   5700
         Top             =   60
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1958
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
         cPositionY      =   30
         Begin VB.OptionButton H3Op_HKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "全 て"
            Height          =   285
            Index           =   2
            Left            =   150
            TabIndex        =   8
            Top             =   750
            Width           =   1035
         End
         Begin VB.OptionButton H3Op_HKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "検収済"
            Height          =   285
            Index           =   1
            Left            =   150
            TabIndex        =   7
            Top             =   450
            Width           =   1035
         End
         Begin VB.OptionButton H3Op_HKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "未検収"
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   6
            Top             =   150
            Value           =   -1  'True
            Width           =   1035
         End
      End
   End
   Begin PdqcommLib.PDQComm PDQComm1 
      Height          =   480
      Left            =   1260
      TabIndex        =   20
      Top             =   660
      Width           =   480
      _Version        =   196610
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.26
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FastScroll      =   0   'False
      Object.Height          =   0
      Object.Width           =   0
   End
   Begin Cis3D_v60.CIS3D PNL_Wait 
      Height          =   930
      Left            =   4830
      Top             =   4290
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
   Begin Cis3D_v60.CIS3D CIS3D1 
      Height          =   930
      Left            =   10635
      Top             =   1800
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1640
      BackColor       =   0
      ForeColor       =   65535
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
      cBoderColor1    =   8421504
      cBoderColor2    =   -2147483643
      cFont3DColor1   =   16777215
      cFont3DColor2   =   12632256
      cAlingnment     =   7
      cBoderStyle     =   1
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   420
         Left            =   45
         Top             =   45
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   741
         Caption         =   "表示件数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D B1lb_HKensu 
            Height          =   360
            Left            =   1155
            Top             =   30
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS UI Gothic"
               Size            =   15.75
               Charset         =   0
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
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   420
         Left            =   45
         Top             =   465
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   741
         Caption         =   "検 収 済"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D B1lb_KKensu 
            Height          =   390
            Left            =   1170
            Top             =   15
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   688
            BackColor       =   12648384
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS UI Gothic"
               Size            =   15.75
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
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   420
         Left            =   2160
         Top             =   465
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   741
         ForeColor       =   0
         Caption         =   "未 検 収"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D B1lb_MKensu 
            Height          =   345
            Left            =   1170
            Top             =   30
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            BackColor       =   16761087
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS UI Gothic"
               Size            =   15.75
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
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   420
         Left            =   2160
         Top             =   45
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   741
         ForeColor       =   0
         Caption         =   "選 択 数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D B1lb_SKensu 
            Height          =   345
            Left            =   1170
            Top             =   15
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            BackColor       =   16777215
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS UI Gothic"
               Size            =   15.75
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
            cPositionX      =   -30
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
      Height          =   750
      Left            =   9630
      ScaleHeight     =   690
      ScaleWidth      =   5265
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   525
      Width           =   5325
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   2670
         Top             =   45
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
         Left            =   3525
         Top             =   45
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
         Left            =   4380
         Top             =   45
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
         Left            =   60
         Top             =   45
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
         Left            =   960
         Top             =   45
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
      Begin CisBtn_60.CisBtn PB_UPD 
         Height          =   585
         Left            =   1815
         Top             =   45
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
         BtnCaption      =   "更新"
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
         PFForeColor     =   65535
         PFShadowColor   =   0
         PFBackColor     =   0
      End
      Begin CisBtn_60.CisBtn PB_Bun 
         Height          =   585
         Left            =   945
         Top             =   45
         Visible         =   0   'False
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
         BtnCaption      =   "個別"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   192
         BtnShadowColor  =   16777215
         PFCaption       =   "( F4 )"
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
         PFBackColor     =   192
      End
   End
   Begin Cis3D_v60.CIS3D H4_Area1 
      Height          =   1320
      Left            =   495
      Top             =   3750
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2328
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
      Begin VB.TextBox QRDummy 
         Height          =   330
         Left            =   3645
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   30
         Width           =   240
      End
      Begin Cis3D_v60.CIS3D QRMsg 
         Height          =   1170
         Left            =   90
         Top             =   75
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   2064
         BackColor       =   12648384
         ForeColor       =   16711680
         Caption         =   "QRリーダー読取"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   7
         cAlingnment     =   7
      End
   End
   Begin VB.PictureBox H2_Area1 
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
      Height          =   1275
      Left            =   4440
      ScaleHeight     =   1215
      ScaleWidth      =   4200
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   450
      Width           =   4260
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   1110
         Left            =   2250
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "納 番"
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
         Begin CisText_V60.CisText iH2_Nouno 
            Height          =   360
            Left            =   195
            TabIndex        =   10
            Top             =   480
            Width           =   780
            _ExtentX        =   1376
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
            cDataReplace    =   1
            cFaZero         =   0
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1110
         Left            =   60
         Top             =   60
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "納 入 日"
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
         Begin CisYMD_With_Btn.CisYMDwB iH2_NYmd 
            Height          =   360
            Left            =   165
            TabIndex        =   9
            Top             =   495
            Width           =   1860
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
            Object.Width           =   1860
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
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   1110
         Left            =   3465
         Top             =   60
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "頁"
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
         Begin CisText_V60.CisText iH2_NPage 
            Height          =   360
            Left            =   135
            TabIndex        =   11
            Top             =   480
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
            Text            =   "##"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
      End
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
      Height          =   1275
      Left            =   390
      ScaleHeight     =   1215
      ScaleWidth      =   2265
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1455
      Width           =   2325
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   1110
         Left            =   60
         Top             =   60
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "検 収 日"
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_NYmd 
            Height          =   360
            Left            =   120
            TabIndex        =   0
            Top             =   450
            Width           =   1860
            _ExtentX        =   3069
            _ExtentY        =   582
            cCalender       =   -1  'True
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
            Object.Width           =   1860
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   7725
      Left            =   390
      ScaleHeight     =   7665
      ScaleWidth      =   14505
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2850
      Width           =   14565
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   7560
         Left            =   60
         Top             =   45
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   13335
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
         Begin vsFlexLib.vsFlexArray vsGrid1 
            Height          =   7350
            Left            =   90
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   105
            Width           =   14265
            _Version        =   196608
            _ExtentX        =   25162
            _ExtentY        =   12965
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
            FormatString    =   "№|処理 |納　期 　　-便|出荷日  　　-便|納入先 |受 |納 番 |頁|受注区||品　番　　　　　　　|納品書管理NO|納品書頁|Select|"
            Rows            =   24
            Cols            =   17
            BackColor       =   12648447
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   300
            GridLines       =   2
            SelectionMode   =   1
         End
      End
   End
   Begin VB.TextBox Dummy 
      Height          =   330
      IMEMode         =   3  'ｵﾌ固定
      Left            =   90
      TabIndex        =   14
      Text            =   "Dummy"
      Top             =   420
      Width           =   240
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
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 検収入力 [一括)]  】"
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
      cPositionX      =   3000
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   661
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "XXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cBoderWidth     =   3
         cFont3D         =   2
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14340
         Top             =   90
         Width           =   945
         _ExtentX        =   1667
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
         Height          =   315
         Left            =   11850
         Top             =   45
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
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
   Begin Cis3D_v60.CIS3D ErrMsg 
      Height          =   465
      Left            =   3315
      Top             =   945
      Visible         =   0   'False
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   820
      BackColor       =   16761087
      Caption         =   "エラー"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   14.25
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
End
Attribute VB_Name = "CXK0010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   検収入力
'**       フォームID    :   CXK0010
'**       処理概要      :
'**
'**       作  成  日    :   2004/03/25  By CIS
'**       変  更  日    :   2004/06/28  By CIS - 納入先の検索は受注全てを対象にする
'**       変  更  日    :   2005/01/18  By CIS - 受注区分 = 7(代替)は検収の対象外とする
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim CisVsGrid               As New CisVsGrid3
    Dim SL_Select               As String
    Dim mSL_Select              As String
    Dim mSL_Group               As String
    Dim mSL_Order               As String
    
    Dim mHeadProc               As String
    
    Dim mKCount                 As Long
    Dim mMCount                 As Long
    Dim mSCount                 As Long
    
    Dim QRReadFlg               As Boolean
    
    Dim mURYM                   As String
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    If Rv_NCall = "" Then
    '   #-------------------#
    '   # 二 重 起 動 防 止  #
    '   #-------------------#
        If Not Dupli_Start Then End
    '   #-------------------#
    '   # 初 期 内 容 取 得  #
    '   #-------------------#
        If Not IniGet Then End
    End If
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    QRDummy.Left = -1000
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 600    '
    Timer2.Enabled = True    '

    H2_Area1.Move H3_Area1.Left, H3_Area1.Top
    H4_Area1.Move H3_Area1.Left, H3_Area1.Top

    Set QT10QSR_Comm = PDQComm1
    Call QT10QSRInfoGet
    
    If QT10QSR_UseFlg Then
       IRN_NKRead = 1
    Else
       IRN_NKRead = 0
    End If
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 2
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 16
        Set .SelectDispObj = B1lb_SKensu
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
    
    With CisFun
        .INI_FileID = gIniExe & gIniName
        .INI_Section = "検収入力"
        .INI_Key = "検収"
        .INI_Default = ""
        
        .GetIni
    
        If RTrim$(.INI_String) = "" Then
            .INI_String = "1"
            .PutIni
        End If
        
        gInt = .Val2(.INI_String)
    End With
    
    If gInt <> 0 Then gInt = 1
    

    iH1_NYmd = Format(Now(), "YYYYMMDD")

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    
    K_Sykbnm.Tag = "2"
    Call PB_CHG_Click

    H3Op_HKbn(0).Value = True

'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    If Rv_NCall = "" Then
       Me.Show
    End If
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    mSL_Select = "SELECT JY.受注管理NO,JY.納入日,JY.納入便,"
    mSL_Select = mSL_Select & "  ISNULL(SK.出荷日,'') 出荷日,ISNULL(SK.出荷便,'') 出荷便,"
    mSL_Select = mSL_Select & "  KN.実績年月,"
    mSL_Select = mSL_Select & "  JY.納入先,JY.受入,JY.納入番号,JY.納品書頁,JY.納品書行,"
    mSL_Select = mSL_Select & "  JY.受注区分,JY.出荷指示書NO,JY.出荷指示書行NO,"
    mSL_Select = mSL_Select & "  CASE WHEN ISNULL(検収日,'') = '' THEN 出荷数 ELSE 検収数 END 出荷数,"
    mSL_Select = mSL_Select & "  ISNULL(TR.略称,'') 納入先名,"
    mSL_Select = mSL_Select & "  ISNULL(SY1.値名称,'') 受注区分名,"
    mSL_Select = mSL_Select & "  HN.表示品番,HN.実数区分,"
    mSL_Select = mSL_Select & "  ISNULL(KN.検収日,'') 検収日,"
    mSL_Select = mSL_Select & "  ISNULL(KN.検収管理NO,'') 検収管理NO "
    mSL_Select = mSL_Select & " From 出荷テーブル SK"
    mSL_Select = mSL_Select & "  LEFT OUTER Join 受注テーブル JY"
    mSL_Select = mSL_Select & "     On SK.出荷指示書NO = JY.出荷指示書NO"
    mSL_Select = mSL_Select & "    AND SK.出荷指示書行NO = JY.出荷指示書行NO"
    mSL_Select = mSL_Select & "  LEFT OUTER Join 検収テーブル KN"
    mSL_Select = mSL_Select & "     On KN.指示書NO = JY.出荷指示書NO"
    mSL_Select = mSL_Select & "    AND KN.指示書行NO = JY.出荷指示書行NO"
    mSL_Select = mSL_Select & "  Left Outer Join 品番マスタ HN"
    mSL_Select = mSL_Select & "     ON HN.品番 = JY.品番"
    mSL_Select = mSL_Select & "  Left Outer Join 取引先マスタ TR"
    mSL_Select = mSL_Select & "     ON TR.取引先区分 = 0"
    mSL_Select = mSL_Select & "     And TR.取引先CD = JY.納入先"
    mSL_Select = mSL_Select & "  Left Outer Join 名称マスタ SY1"
    mSL_Select = mSL_Select & "     ON SY1.区分名称 = '受注区分'"
    mSL_Select = mSL_Select & "     And SY1.区分タイプ = ''"
    mSL_Select = mSL_Select & "     And SY1.値 = JY.受注区分"
    
    mSL_Order = " Order by JY.納入日,JY.納入便,JY.納入先,JY.受入,"
    mSL_Order = mSL_Order & "JY.納入番号,JY.納品書頁,JY.納品書行,JY.受注区分"

    
' * 売上実績更新年月
    Call ItemsClearSYM
    gSL_Select = "SELECT * FROM 名称マスタ "
    gSL_Select = gSL_Select & " WHERE 区分名称 = '処理内容' "
    gSL_Select = gSL_Select & "   AND 値 = 'PXU0110' "
    gSL_Select = gSL_Select & "   AND 区分タイプ = '' "
    Call SYMRead(gSL_Select, 1)
    mURYM = SYM.英数字1
    
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
        Case vbKeyF1:       If PB_UPD.Visible Then Call PB_UPD_Click        '【更新】
        Case vbKeyF4:       If PB_Bun.Visible Then Call PB_Bun_Click        '【個別】
        Case vbKeyF6:       If PB_CHG.Visible Then Call PB_CHG_Click        '【切替】
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
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    
    Call CisDB.DBDISConnect                         ' ﾃﾞｰﾀﾍﾞｰｽ切断
    Call CisFrm.UnLoadDisp                          ' ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
    End

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
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "B1" Then
        Call HeadBodyClear("B")
        ProcHB = mHeadProc
        Call DispChange(ProcHB)
        ' 表示件数カウント初期化
        mKCount = 0
        mMCount = 0
        mSCount = 0
        GoTo PB_CAN_Ed
    End If
    
    If ProcHB = "H2" Or ProcHB = "H3" Or ProcHB = "H4" Then
        ProcHB = "H1"
        Call HeadBodyClear("H2")
        Call HeadBodyClear("H3")
        Call HeadBodyClear("B")
        Call DispChange(ProcHB)
        GoTo PB_CAN_Ed
    End If
    
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+---------------------+
'+    一括画面切替     +
'+---------------------+
Private Sub PB_Bun_Click()
'    Rv_NCall = "K2"
'    Rv_NTtlSu = 0
'
'    Timer1.Enabled = False
'    Call QR_Comm_Close
'
'    CXK0020.Show vbModal
'    Unload CXK0020
'    Set CXK0020 = Nothing
'
'    Timer1.Enabled = True
'    Rv_NCall = ""
'
'    '   ( QR Reader 環境 )
'    Set QRr_Object = PDQComm1
'    Set QRr_MsgObject = QRMsg
'    Set QRr_ErrMsgObj = ErrMsg
'    Call QRInit
'    Call DispChange(ProcHB)
    
End Sub
'+------------------+
'+    切替   (F6)   +
'+------------------+
Private Sub PB_CHG_Click()
    gLong = CisFun.Val2(K_Sykbnm.Tag)
    
    gLong = gLong + 1
    If gLong > 4 Then gLong = 1
    
    gStr = ""
    gStr = "H3"
    If IRN_NKRead = 1 Then
       gStr = gStr & "H4"
    Else
       gStr = gStr & "  "
    End If
    gStr = gStr & "H2"
    
    mHeadProc = ""
    For gInt = 1 To 4
        mHeadProc = Mid(gStr, gLong + (gLong - 1), 2)
        If Trim(mHeadProc) <> "" Then
           K_Sykbnm.Tag = gLong
           Exit For
        End If
        gLong = gLong + 1
        If gLong > 4 Then gLong = 1
    Next gInt
    
    Select Case mHeadProc
        Case "H3"
            K_Sykbnm.Caption = "【 一覧 】"
            H2_Area1.Visible = False
            H3_Area1.Visible = True
            H4_Area1.Visible = False
        Case "H4"
            K_Sykbnm.Caption = "【 QRﾀｯﾁﾘｰﾀﾞ 】"
            H2_Area1.Visible = False
            H3_Area1.Visible = False
            H4_Area1.Visible = True
        Case Else
            K_Sykbnm.Caption = "【 納番 】"
            H2_Area1.Visible = True
            H3_Area1.Visible = False
            H4_Area1.Visible = False
    End Select
    If ProcHB <> "H1" Then
       ProcHB = mHeadProc
       Call DispChange(ProcHB)
       Select Case mHeadProc
            Case "H2": iH2_NYmd.SetFocus
            Case "H3"
                iH3_SYmd.SetFocus
                Call HeadBodyClear("H2")
                Call HeadBodyClear("H3")
                Call HeadBodyClear("B")
                ' 表示件数カウント初期化
                mSCount = 0
                mKCount = 0
                mMCount = 0
       End Select
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_NYmd" Then iH1_NYmd.ShowCalender
    If PB_Look.Tag = "iH2_NYmd" Then iH2_NYmd.ShowCalender
    If PB_Look.Tag = "iH3_SYmd" Then iH3_SYmd.ShowCalender
    If PB_Look.Tag = "iH3_EYmd" Then iH3_EYmd.ShowCalender
    If PB_Look.Tag = "iH3_Torcd" Then Call Look_TORCD(1)
End Sub
'+-------------------------+
'+    取　引　先　検　索    +
'+-------------------------+
Private Sub Look_TORCD(pKB As Byte)

    RV_Left = 0:    RV_Top = 0
'*---------------------------------------------<< 2004/06/28 Update Start >>
'   RV_TorKb = 1
    RV_TorKb = 3
'*---------------------------------------------<< 2004/06/28 Update End >>
    
    CKK0025.Show vbModal

    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH3_Torcd = RTrim$(RV_TorcdK)
        Call iH3_Torcd_LostFocus
    End If
    
    iH3_Torcd.SetFocus
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*--------------------------------------*
'*      入 荷 日    ( iH1_NYmd )        *
'*--------------------------------------*
Private Sub iH1_NYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_NYmd_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*----------------------------------------------*
'*      納 入 日  ( iH2_NYmd )       *
'*----------------------------------------------*
Private Sub iH2_NYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH2_NYmd_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*----------------------------------------------*
'*      ( 開 始 ) 納 入 日  ( iH3_SNYmd )       *
'*----------------------------------------------*
Private Sub iH3_SYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH3_SYmd_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*----------------------------------------------*
'*      ( 終 了 ) 納 入 日  ( iH3_ENYmd )       *
'*----------------------------------------------*
Private Sub iH3_EYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH3_EYmd_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'*----------------------------------*
'*      取引先  ( iH3_Torcd )       *
'*----------------------------------*
Private Sub iH3_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH3_Torcd_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    TRM.略称 = ""
    If RTrim$(iH3_Torcd) = "" Then
        H3lb_Tornm = ""
    Else
        Call TorNmGet(RTrim$(iH3_Torcd), 3)
        H3lb_Tornm = RTrim$(TRM.略称)
    End If
End Sub
'*----------------------------------*
'*      表示区分    (H3Op_HKbn)
'*----------------------------------*
Private Sub H3Op_HKbn_Click(Index As Integer)
    H3Op_HKbn(0).Tag = Index
End Sub

'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = mHeadProc
        Select Case mHeadProc
            Case "H3"
               Call HeadBodyClear("H3")
               iH3_EYmd = iH1_NYmd
               If Trim(iH3_EYmd) = "" Then iH3_EYmd = Format(Now(), "YYYYMMDD")
        End Select
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "H2" Then
       If Not Head2Chk Then GoTo ReturnPress_Ed
       GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "H3" Then
        If Not Head3Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        vsGrid1.SetFocus
        GoTo ReturnPress_Ed
    End If

    
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
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
    Dim H1Mode          As Boolean
    Dim H1Color         As Long
    Dim H2Mode          As Boolean
    Dim H2Color         As Long
    Dim H3Mode          As Boolean
    Dim H3Color         As Long
    Dim H4Mode          As Boolean
    Dim H4Color         As Long
    Dim B1Mode          As Boolean
    Dim B1Color         As Long
    
    H1Mode = False:     H1Color = gPLostSel
    H2Mode = False:     H2Color = gPLostSel
    H3Mode = False:     H3Color = gPLostSel
    H4Mode = False:     H4Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "H2":      H2Mode = True:  H2Color = gPGotSel
        Case "H3":      H3Mode = True:  H3Color = gPGotSel
        Case "H4":      H4Mode = True:  H4Color = gPGotSel
        Case "B1":      B1Mode = True:  B1Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    H2_Area1.Enabled = H2Mode
    H2_Area1.BackColor = H2Color
    H3_Area1.Enabled = H3Mode
    H3_Area1.BackColor = H3Color
    H4_Area1.Enabled = H4Mode
    H4_Area1.BackColor = H4Color
    
    B1_Area1.Enabled = B1Mode Or H2Mode Or H3Mode Or H4Mode
    B1_Area1.BackColor = B1Color
    
    PB_CAN.Visible = Not H1Mode
    PB_CHG.Visible = Not B1Mode
    PB_Look.Visible = H1Mode
    PB_Bun.Visible = H1Mode
    
    PB_UPD.Visible = B1Mode
    If H2Mode Or H4Mode Then
        If mSCount <> 0 Then
            PB_UPD.Visible = True
        End If
    End If
    
    PB_ENT.Visible = Not (H4Mode Or B1Mode)
    
    If ProcHB = "H1" Then
       B1lb_HKensu = ""
       B1lb_SKensu = ""
       B1lb_KKensu = ""
       B1lb_MKensu = ""
       PB_Look.Visible = False
    End If
'
    If ProcHB = "H4" Then
       Call QT10QSR_Open
       QRMsg.BackColor = &HC0FFC0
       QRMsg.Caption = "QRﾀｯﾁﾘｰﾀﾞ読取中"
    Else
       Call QT10QSR_Close
       QRMsg.BackColor = &HFFC0C0
       QRMsg.Caption = "QRﾀｯﾁﾘｰﾀﾞ読取"
    End If

' ﾀﾞﾐｰ項目(入力禁止)
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
'            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj
    If HeadBody = "B" Then
       CisVsGrid.Clear
    End If
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
'( 検収日 )
    If RTrim$(iH1_NYmd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    検収日を入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_NYmd.SetFocus
        Exit Function
    Else
        If RTrim$(iH1_NYmd) <> "" And Not iH1_NYmd.cChkResult Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    正しい日付を入力して下さい         "
                 .MB_Button = Error
                 .MBOX
             End With
             iH1_NYmd.SetFocus
             Exit Function
         End If
    End If
    
    ' 表示件数カウント初期化
    mKCount = 0
    mMCount = 0
    mSCount = 0
    
    Head1Chk = True
End Function
'*-----------------------------------------*
'*      HEAD2ﾁｪｯｸ処理 ( 納入日～頁 )       *
'*-----------------------------------------*
Private Function Head2Chk() As Boolean
    Head2Chk = False
    
'( 納入日 )
    If RTrim$(iH2_NYmd) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  納入日を入力して下さい  "
            .MB_Button = Error
            .MBOX
        End With
        iH2_NYmd.SetFocus
        Exit Function
    Else
        If Not iH2_NYmd.cChkResult Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    正しい日付を入力して下さい         "
                 .MB_Button = Error
                 .MBOX
             End With
             iH2_NYmd.SetFocus
             Exit Function
         End If
    End If
    
'( 納入番号 )
    If RTrim$(iH2_Nouno) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  納番を入力して下さい  "
            .MB_Button = Error
            .MBOX
        End With
        iH2_Nouno.SetFocus
        Exit Function
    End If
    
'( 納品書頁 )
    If CisFun.Val2(iH2_NPage) = 0 Then iH2_NPage = " 1"
    
'( データ確認 )
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where JY.納入番号 = '" & RTrim$(iH2_Nouno) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   JY.納品書頁 =  " & CisFun.Val2(iH2_NPage) & vbCrLf
    gSL_Select = gSL_Select & " And   JY.納入日   = '" & RTrim$(iH2_NYmd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " Order by JY.納品書行 Desc "
    
    If Not JYTRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  該当データがありません  "
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
    
    If Not GridSetKobetsu Then Exit Function

    Head2Chk = True
End Function
'*-----------------------------------------*
'*      HEAD3ﾁｪｯｸ処理 ( 納入日～実績 )     *
'*-----------------------------------------*
Private Function Head3Chk() As Boolean
    Dim DateStr As String
    Dim DateSt  As String
    Dim DateEd  As String
    Dim BinSt   As String
    Dim BinEd   As String
    
    Head3Chk = False
    
    gSL_Select = mSL_Select
    With CisFun
         .Where_CreateKB = NewDocument
'---<< 受注区分 = 7(代替)は検収の対象外とする　>>　　　　2005/01/18 追加 (START) -----'
         .Where_Create AlphaMe, "JY.受注区分", 7, 等しくない
'---------------------------------------------------　2005/01/18 追加 (END  ) -----'
         ' 日付
         DateStr = "出荷"
         DateSt = "": DateEd = "": BinSt = "": BinEd = ""
         ' 日付 月指定
         If Trim(iH3_SYmd) <> "" And CisFun.Val2(iH3_SYmd.cDay) = 0 Then
            .Where_Create AlphaMe, "Substring(" & DateStr & "日,1,6)", iH3_SYmd
            .Where_Create AlphaMe, DateStr & "便", Format(iH3_SBin, "00")
         End If
         ' 日付開始編集
         If Trim(iH3_SYmd) <> "" Then
            DateSt = iH3_SYmd
         End If
         If CisFun.Val2(iH3_SBin) <> 0 Then
            BinSt = Format(iH3_SBin, "00")
         End If
         ' 日付終了編集
         If Trim(iH3_EYmd) <> "" Then
            DateEd = iH3_EYmd
         End If
         If CisFun.Val2(iH3_EBin) <> 0 Then
            BinEd = Format(iH3_EBin, "00")
         End If
         ' 日付開始のみ
         If Trim(DateSt) <> "" And Trim(DateEd) = "" And CisFun.Val2(iH3_SYmd.cDay) <> 0 Then
            If Trim(BinSt) = "" Then
               .Where_Create AlphaMe, DateStr & "日", DateSt
            Else
               .Where_Create AlphaMe, DateStr & "日+" & DateStr & "便", DateSt & BinSt
            End If
         End If
         ' 日付終了のみ
         If Trim(DateSt) = "" And Trim(DateEd) <> "" Then
            If Trim(BinEd) = "" Then
               .Where_Create AlphaMe, DateStr & "日", DateEd, 以下
            Else
               .Where_Create AlphaMe, DateStr & "日+" & DateStr & "便", DateEd & BinEd, 以下
            End If
         End If
         ' 日付開始・終了
         If Trim(DateSt) <> "" And Trim(DateEd) <> "" Then
            If Trim(BinSt) = "" And Trim(BinEd) = "" Then
               .Where_Create AlphaMe, DateStr & "日", DateSt, Between, , DateEd
            Else
               If Trim(BinSt) = "" Then BinSt = "00"
               If Trim(BinEd) = "" Then BinEd = "99"
               .Where_Create AlphaMe, DateStr & "日+" & DateStr & "便", DateSt & BinSt, Between, , DateEd & BinEd
            End If
         End If
         ' 仕入先
         .Where_Create AlphaMe, "JY.納入先", iH3_Torcd
         ' 未検収のみ
         If CisFun.Val2(H3Op_HKbn(0).Tag) = 0 Then
            .Where_Create AlphaMe, "ISNULL(KN.検収日,'')", "", 等しい, Create
         End If
         ' 検収済のみ
         If CisFun.Val2(H3Op_HKbn(0).Tag) = 1 Then
            .Where_Create AlphaMe, "ISNULL(KN.検収日,'')", "", 等しくない, Create
         End If
         ' 未検収のみ
         If CisFun.Val2(H3Op_HKbn(0).Tag) = 0 Then
            If Trim(.Where_Phrase) = "" Then
               gSL_Select = gSL_Select & " Where ISNULL(SK.出荷数,0) <> 0 or ISNULL(KN.検収数,0) <> 0 "
            Else
               gSL_Select = gSL_Select & .Where_Phrase & " And ( ISNULL(SK.出荷数,0) <> 0 or ISNULL(KN.検収数,0) <> 0 )"
            End If
         Else
            gSL_Select = gSL_Select & .Where_Phrase
         End If
    End With
    
    gSL_Select = gSL_Select & mSL_Group & mSL_Order
    
    PNL_Wait.Visible = True
    Message = "【抽出中】しばらくお待ち下さい"
    DoEvents
    
    If Not JYTRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  該当データがありません  "
            .MB_Button = Error
            .MBOX
        End With
        PNL_Wait.Visible = False
        Exit Function
    End If
    
    Call GridSet
    
    PNL_Wait.Visible = False
    
    Head3Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
'    If iB1_???? = "" Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    ???を入力して下さい         "
'            .MB_Button = Error
'            .MBOX
'        End With
'        iB1_???.SetFocus
'        Exit Function
'    End If
    
'   < 印刷データ確認 >
'    gSL_Select = "SELECT * FROM ??テーブル"
'    gSL_Select = gSL_Select & " WHERE Key = '" & RTrim(iH1_NYmd) & "'"
    
'    If Not NKTRead(gSL_Select, 1) Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    該当データがありません            "
'            .MB_Button = Error
'            .MBOX
'        End With
''        iB1_???.SetFocus
'        Exit Function
'    End If
    
    Body1Chk = True
End Function
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'☆
'☆                         検   収　　更　　新　　(  F1  )
'☆
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
Private Sub PB_UPD_Click()
    Static ws_KNo   As Long
    
    gCnt1 = 0
    With vsGrid1
        For gCnt2 = 0 To .Rows - 1
            If Trim(.TextMatrix(gCnt2, 1)) <> "" Then
               gCnt1 = 1
               Exit For
            End If
        Next gCnt2
    End With
        
    If gCnt1 = 0 Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "      処理を選択して下さい               "
            .MB_Button = OK_CAN
            .MBOX
            Exit Sub
        End With
    End If
    
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "      検収実績更新を行います。               "
        .MB_MSG(4) = "          よろしいですか？                      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Sub
    End With

    On Error GoTo PB_UPD_Er:
    

    PNL_Wait.Visible = True
    Message = "【実績更新中】しばらくお待ち下さい"
    DoEvents
    '+---------------
    '+   完納処理
    '+---------------
    Call CisDB.DBTran(TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    With vsGrid1
        .Redraw = False
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, 1)) <> "" Then
               
                If CisFun.Val2(.TextMatrix(gLong, 15)) <> 0 Then
                ' 検収取消
                    With CisDB
                        .SQL = "DELETE FROM 検収テーブル "
                        .SQL = .SQL & " WHERE 検収管理NO = " & CisFun.Val2(vsGrid1.TextMatrix(gLong, 15))
                        .SQL = .SQL & "   AND 検収管理行NO = 1"
                        If Not .DBExec Then GoTo PB_UPD_Er
                    End With
                Else
                ' 検収処理
                    With CisDB
                    '   <><><> 検収管理NOの獲得 <><><>
                        .SQL = "番号獲得"
                        .StoadoCount = 3
                        .ParaNo = 0: .ParaIO = Return用
                        .ParaNo = 1: .ParaIO = Input用
                        .ParaNo = 2: .ParaIO = OutPut用
                        .ParaNo = 3: .ParaIO = Return用
                    '
                        .ParaNo = 1: .ParaValue = 13
                    '
                        If Not .DBStored Then
                            CisFun.MB_Lines = 4
                            CisFun.MB_MSG(2) = "    検収管理NOの獲得に失敗しました            "
                            CisFun.MB_Title = "採番処理"
                            CisFun.MB_Button = Error
                            CisFun.MBOX
                            Exit Sub
                        Else
                            .ParaNo = 2: ws_KNo = .ParaValue
                        End If
                    End With
               
               
                     With CisDB
                         .SQL = "Insert Into 検収テーブル  ( " & vbCrLf
                         .SQL = .SQL & "検収管理NO,検収管理行NO,指示書NO,指示書行NO,データ区分," & vbCrLf
                         .SQL = .SQL & "検収日,納入先,受入,機番,納入番号,納品書頁,納品書行," & vbCrLf
                         .SQL = .SQL & "品番,背番号,赤黒区分,枚数,収容数,検収数," & vbCrLf
                         .SQL = .SQL & "作成種別,照合年月,照合処理日,照合結果," & vbCrLf
                         .SQL = .SQL & "実績年月,実績処理日,実績結果," & vbCrLf
                         .SQL = .SQL & "作成日,作成端末,作成者 )" & vbCrLf
                         
                         .SQL = .SQL & "Select " & vbCrLf
                         .SQL = .SQL & ws_KNo & ",1,JY.出荷指示書NO,JY.出荷指示書行NO,0," & vbCrLf
                         If Trim(iH1_NYmd) <> "" Then
                            .SQL = .SQL & "'" & iH1_NYmd & "'," & vbCrLf
                         Else
                            .SQL = .SQL & "SK.出荷日," & vbCrLf
                         End If
                         .SQL = .SQL & "JY.納入先,JY.受入,JY.機番," & vbCrLf
                         .SQL = .SQL & "JY.納入番号,JY.納品書頁,JY.納品書行," & vbCrLf
                         .SQL = .SQL & "JY.品番,JY.背番号,0," & vbCrLf
                         .SQL = .SQL & "SK.出荷枚数,JY.収容数,SK.出荷数," & vbCrLf
                         .SQL = .SQL & "JY.作成種別,'','',''," & vbCrLf
                         .SQL = .SQL & "'','',''," & vbCrLf
                         .SQL = .SQL & "Getdate(),Host_Name(),'" & gTanto & "'" & vbCrLf
                         .SQL = .SQL & " From 受注テーブル JY"
                         .SQL = .SQL & " LEFT OUTER JOIN 出荷テーブル SK"
                         .SQL = .SQL & "   ON SK.出荷指示書NO = JY.出荷指示書NO "
                         .SQL = .SQL & "  AND SK.出荷指示書行NO = JY.出荷指示書行NO "
                         .SQL = .SQL & " Where JY.受注管理NO = " & CisFun.Val2(vsGrid1.TextMatrix(gLong, 12))
                         .SQL = .SQL & "   And JY.出荷指示書NO = " & CisFun.Val2(vsGrid1.TextMatrix(gLong, 13))
                         .SQL = .SQL & "   And JY.出荷指示書行NO = " & CisFun.Val2(vsGrid1.TextMatrix(gLong, 14))
                         
                         If Not .DBExec Then GoTo PB_UPD_Er
                    End With
               
                End If
                
                .Row = gLong
                Call CisVsGrid.vsColor(消去個別)
                DoEvents
            
            End If
        Next gLong
        .Row = 1: .Col = 1: .TopRow = 1
        .Redraw = True
    End With
    Call CisDB.DBTran(TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ終了
    
PB_UPD_Ed:
    PNL_Wait.Visible = False
    PB_UPD.Visible = False

    Call PB_CAN_Click
'    Timer1.Enabled = True
    Exit Sub
PB_UPD_Er:
    vsGrid1.Redraw = True
    PNL_Wait.Visible = False
    Call CisDB.DBTran(TransRollback) ' ﾄﾗﾝｻﾞｸｼｮﾝ破棄

End Sub
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 中中, あり, 5, 1, "処理")
        Call .FixedSet(中中, 中中, あり, 13, 2, "出荷日-便")
        Call .FixedSet(中中, 中中, あり, 13, 3, "検収日")
        Call .FixedSet(中中, 左中, あり, 21, 4, "納入先")
        Call .FixedSet(中中, 中中, あり, 2, 5, "受")
        Call .FixedSet(中中, 中中, あり, 6, 6, "納 番")
        Call .FixedSet(中中, 右中, あり, 2, 7, "頁")
        Call .FixedSet(中中, 中中, あり, 7, 8, "受注")
        Call .FixedSet(中中, 左中, あり, 29, 9, "品　番")
        Call .FixedSet(中中, 右中, あり, 7, 10, "数量")
        Call .FixedSet(中中, 左中, あり, 0, 11, "X")
        Call .FixedSet(中中, 左中, あり, 0, 12, "発注管理NO")
        Call .FixedSet(中中, 左中, あり, 0, 13, "出荷指示書NO")
        Call .FixedSet(中中, 左中, あり, 0, 14, "出荷指示書行NO")
        Call .FixedSet(中中, 左中, あり, 0, 15, "検収管理NO")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 16, "Select")

'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10")

        .InitDisp
    End With
End Sub
'******************************************'
'*****        行選択（クリック）       *****
'******************************************'
Private Sub VSGrid1_Click()
    
    With vsGrid1
        .Redraw = False
        For gLong = .Row To .RowSel
            If Trim(.TextMatrix(gLong, CisVsGrid.NoSpaceCol)) <> "" And _
               Trim(.TextMatrix(gLong, 11)) = "" Then
                gStr = Trim(.TextMatrix(gLong, CisVsGrid.SelectCol))
                .Row = gLong
                Select Case Trim(.TextMatrix(gLong, 1))
                    Case ""
                        If gStr = "X" Then
                           Call CisVsGrid.vsColor(消去個別)
                        End If
                        If Trim(.TextMatrix(gLong, 3)) = "" Then
                            .TextMatrix(gLong, 1) = "検収"
                            Call CisVsGrid.vsColor(個別, B1lb_KKensu.BackColor)
                        Else
                            .TextMatrix(gLong, 1) = "取消"
                            Call CisVsGrid.vsColor(個別, B1lb_MKensu.BackColor)
                        End If
                        mSCount = mSCount + 1
                    Case Else
                        If gStr = "X" Then
                           Call CisVsGrid.vsColor(消去個別)
                        End If
                        Call CisVsGrid.vsColor(消去個別, B1lb_HKensu.BackColor)
                        .TextMatrix(gLong, 1) = ""
                        mSCount = mSCount - 1
                End Select
            End If
        Next gLong
        .Redraw = True
    End With
    
    B1lb_SKensu = Format(mSCount, "#,###")
    
    If mSCount <> 0 Then
        PB_UPD.Visible = True
    Else
        PB_UPD.Visible = False
    End If

End Sub
'******************************************'
'*****        行選択（スペースｷｰ)      *****
'******************************************'
Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    Call VSGrid1_Click
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
    B1lb_HKensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
       .Redraw = False
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    mKCount = 0: mMCount = 0
    Do Until Not JYT_RDSTS
        With vsGrid1
            ID = ID + 1
       
            Call GridMove(ID)
        End With
        Call JYTReadNext
    Loop
    Call JYTClose
    
    Call CisVsGrid.ReNumber
    
    With vsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With

    B1lb_KKensu = Format(mKCount, "#,###")
    B1lb_MKensu = Format(mMCount, "#,###")
End Sub
'+------------------------------------+
'+      グリッドセット（個別)
'+------------------------------------+
Private Function GridSetKobetsu() As Boolean
    GridSetKobetsu = False
    
    If JYT.検収日 <> "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "      検収済みデータです              "
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
'---<< 受注区分 = 7(代替)は検収の対象外とする　　　　　　2005/01/18 追加 (START) -----'
    If JYT.受注区分 = 7 Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  検収対象外！ （代替）  "
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
'---------------------------------------------------　2005/01/18 追加 (END  ) -----'
    
    With vsGrid1
        .Redraw = False

        For gCnt1 = 1 To .Rows - 1
            gCnt2 = CisFun.Val2(.TextMatrix(gCnt1, 13))
            If gCnt2 <> 0 And gCnt2 = JYT.出荷指示書NO Then
               With CisFun
                   .MB_Lines = 5
                   .MB_MSG(2) = "      表示済みです。               "
                   .MB_Button = Error
                   .MBOX
               End With
               .Redraw = True
               Exit Function
            End If
        Next gCnt1
        
        gCnt1 = CisFun.Val2(B1lb_HKensu)
        
        Do Until Not JYT_RDSTS
        
           If Trim(.TextMatrix(.Rows - 1, 2)) <> "" Then
              .Rows = .Rows + 1
           End If
           
           Call GridMove(.Rows - 1)
           .RowPosition(.Rows - 1) = 1
    
            .TextMatrix(1, 1) = "検収"
            mSCount = mSCount + 1
            gCnt1 = gCnt1 + 1
            .Row = 1
            Call CisVsGrid.vsColor(個別, B1lb_KKensu.BackColor)
            
            Call JYTReadNext
        Loop
        Call JYTClose
        
        Call CisVsGrid.ReNumber
        
        .Row = 1: .Col = 1: .ColSel = .Cols - 1
        .Redraw = True
    End With

    B1lb_KKensu = Format(mKCount, "#,###")
    B1lb_MKensu = Format(mMCount, "#,###")
    B1lb_HKensu = Format(gCnt1, "#,###")
    B1lb_SKensu = Format(mSCount, "#,###")
    If mSCount <> 0 Then
        PB_UPD.Visible = True
    Else
        PB_UPD.Visible = False
    End If

    GridSetKobetsu = True
End Function
'+------------------------------------+
'+      グリッド項目転送
'+------------------------------------+
Private Sub GridMove(mRow As Long)

    With vsGrid1
        
        With CisFun
            .DateE_BefVal = JYT.出荷日
            .DateEdit
            gStr = CisFun.DateE_AfVal
            If Trim(JYT.出荷便) <> "" Then
                gStr = gStr & "-" & JYT.出荷便
            Else
                gStr = gStr & "   "
            End If
        End With
        .TextMatrix(mRow, 2) = gStr
        
        If Trim(JYT.検収日) <> "" Then
            With CisFun
                .DateE_BefVal = JYT.検収日
                .DateEdit
                gStr = CisFun.DateE_AfVal
            End With
            .TextMatrix(mRow, 3) = gStr
            
            mKCount = mKCount + 1
        Else
            mMCount = mMCount + 1
        End If
        
        If Trim(JYT.納入先名) = "" Then
           gStr = JYT.納入先
        Else
           gStr = RTrim(JYT.納入先) & Space(7 - Len(JYT.納入先)) & ":"
           gStr = gStr & JYT.納入先名
        End If
        .TextMatrix(mRow, 4) = gStr
        .TextMatrix(mRow, 5) = JYT.受入
        .TextMatrix(mRow, 6) = JYT.納入番号
        .TextMatrix(mRow, 7) = JYT.納品書頁
        .TextMatrix(mRow, 8) = JYT.受注区分名
        .TextMatrix(mRow, 9) = JYT.表示品番
        .TextMatrix(mRow, 10) = Format(JYT.出荷数, "#,##0")
        If JYT.実績年月 <> "" And JYT.実績年月 < mURYM Then: .TextMatrix(mRow, 11) = "X"
        
        .TextMatrix(mRow, 12) = JYT.受注管理NO
        .TextMatrix(mRow, 13) = JYT.出荷指示書NO
        .TextMatrix(mRow, 14) = JYT.出荷指示書行NO
        .TextMatrix(mRow, 15) = JYT.検収管理NO
    
    End With
End Sub
'*=======================================================================================================*
'*=======================================================================================================*
'*
'*                      Ｑ  Ｒ  リ  ー  ダ  (　Ｃｏｍｍ　)  処  理
'*
'*=======================================================================================================*
'*=======================================================================================================*
'+--------------------------+
'+      読  取  処  理      +
'+--------------------------+
Private Sub PDQComm1_OnComm()
    
    gStr = QT10QSR_Event
    If gStr = "ERR" Then Exit Sub
    If gStr <> "RCV" Then Exit Sub
    
    If Not QT10QSR_RCVProc Then Exit Sub    ' CR が存在しない時抜ける

'( データ確認 )
    If Trim(QT10QSR_ReadData) = "" Then Exit Sub
    
    Call QRConvert(QT10QSR_ReadData)
    
    QT10QSR_ReadData = ""
    
'( データ確認 )
    gSL_Select = mSL_Select
    
    gSL_Select = gSL_Select & " Where JY.納入番号 = '" & RTrim$(typQRAfter.納入番号) & "'" & vbCrLf
    gSL_Select = gSL_Select & " And   JY.納品書頁 =  " & CisFun.Val2(typQRAfter.ページNO) & vbCrLf
    gSL_Select = gSL_Select & " And   JY.納入日   = '" & typQRAfter.納入日 & "'" & vbCrLf
    gSL_Select = gSL_Select & " Order By JY.納品書行 Desc"
    If Not JYTRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  該当データがありません  "
            .MB_Button = Error
            .MBOX
        End With
        Exit Sub
    End If
    
    If Not GridSetKobetsu Then Exit Sub
    
End Sub



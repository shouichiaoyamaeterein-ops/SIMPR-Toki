VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXJ0010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "受注画面入力 [号口]"
   ClientHeight    =   10980
   ClientLeft      =   1545
   ClientTop       =   2595
   ClientWidth     =   15330
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
   ScaleHeight     =   10980
   ScaleWidth      =   15330
   WindowState     =   2  '最大化
   Begin VB.PictureBox C1_Area1 
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
      Height          =   480
      Left            =   2100
      ScaleHeight     =   420
      ScaleWidth      =   9270
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   10020
      Visible         =   0   'False
      Width           =   9330
      Begin Cis3D_v60.CIS3D B1lb_ProcMsg 
         Height          =   360
         Left            =   30
         Top             =   30
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   635
         ForeColor       =   255
         Caption         =   "メッセージ"
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
         cAlingnment     =   6
         cBoderStyle     =   1
         cPositionX      =   30
         cPositionY      =   45
      End
      Begin Cis3D_v60.CIS3D PB_PCAN 
         Height          =   375
         Left            =   7110
         Top             =   30
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   661
         Caption         =   "処 理 取 消（ Esc ）"
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
         cBoderWidth     =   4
         cAlingnment     =   7
         cButton         =   -1  'True
      End
   End
   Begin VB.PictureBox H2_Area2 
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
      Height          =   1050
      Left            =   9270
      ScaleHeight     =   990
      ScaleWidth      =   1410
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   2940
      Width           =   1470
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   840
         Index           =   1
         Left            =   90
         Top             =   90
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1482
         ForeColor       =   16711680
         Caption         =   "納入番号"
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
         cPositionX      =   30
         cPositionY      =   75
         Begin CisText_V60.CisText iH2_Nouban 
            Height          =   360
            Left            =   240
            TabIndex        =   12
            Top             =   420
            Width           =   825
            _ExtentX        =   1455
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
      Height          =   705
      Left            =   7965
      ScaleHeight     =   645
      ScaleWidth      =   3405
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   1500
      Width           =   3465
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   870
         Top             =   30
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
         Left            =   1710
         Top             =   30
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
         Left            =   2550
         Top             =   30
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
         Left            =   30
         Top             =   30
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
      Height          =   1050
      Left            =   2100
      ScaleHeight     =   990
      ScaleWidth      =   7110
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   2940
      Width           =   7170
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   840
         Index           =   3
         Left            =   6360
         Top             =   90
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   1482
         ForeColor       =   16711680
         Caption         =   "受入"
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
         cPositionX      =   30
         cPositionY      =   75
         Begin CisText_V60.CisText iH2_Ukeir 
            Height          =   360
            Left            =   120
            TabIndex        =   11
            Top             =   420
            Width           =   450
            _ExtentX        =   794
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
            Text            =   "XX"
            MaxLength       =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   840
         Index           =   1
         Left            =   90
         Top             =   90
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   1482
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
         cPositionX      =   120
         cPositionY      =   75
         Begin CisText_V60.CisText iH2_Torcd 
            Height          =   360
            Left            =   90
            TabIndex        =   8
            Top             =   420
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
         Begin Cis3D_v60.CIS3D H2lb_Tornm 
            Height          =   360
            Left            =   1170
            Top             =   60
            Width           =   2490
            _ExtentX        =   4392
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
            cPositionX      =   45
            cPositionY      =   15
         End
         Begin Cis3D_v60.CIS3D H2lb_Kojyonm 
            Height          =   360
            Left            =   1170
            Top             =   420
            Width           =   2490
            _ExtentX        =   4392
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
            cPositionX      =   45
            cPositionY      =   15
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   840
         Index           =   2
         Left            =   3840
         Top             =   90
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   1482
         ForeColor       =   16711680
         Caption         =   "－ 便  "
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
         cAlingnment     =   3
         cPositionX      =   120
         cPositionY      =   75
         Begin CisYMD_With_Btn.CisYMDwB iH2_Ymd 
            Height          =   345
            Left            =   90
            TabIndex        =   9
            Top             =   420
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
         Begin Cis3D_v60.CIS3D Back_Ymd 
            Height          =   300
            Left            =   510
            Top             =   75
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   529
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
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin CisText_V60.CisText iH2_Bin 
            Height          =   360
            Left            =   2055
            TabIndex        =   10
            Top             =   420
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   635
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
      Left            =   480
      TabIndex        =   64
      Text            =   "X"
      Top             =   450
      Width           =   345
   End
   Begin VB.Timer Timer1 
      Left            =   900
      Top             =   450
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   5640
      Left            =   2100
      ScaleHeight     =   5580
      ScaleWidth      =   9270
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   4380
      Width           =   9330
      Begin Cis3D_v60.CIS3D BLB_Back 
         Height          =   5475
         Index           =   0
         Left            =   60
         Top             =   60
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   9657
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
         Begin Cis3D_v60.CIS3D BLB_Back 
            Height          =   5415
            Index           =   1
            Left            =   30
            Top             =   30
            Width           =   9100
            _ExtentX        =   16060
            _ExtentY        =   9551
            ForeColor       =   0
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
            Begin Cis3D_v60.CIS3D BLB_Line 
               Height          =   480
               Index           =   1
               Left            =   60
               Top             =   540
               Width           =   8970
               _ExtentX        =   15822
               _ExtentY        =   847
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
               Begin CisText_V60.CisText iB1_Seban 
                  Height          =   360
                  Index           =   1
                  Left            =   3615
                  TabIndex        =   14
                  Top             =   60
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
                  IMEMode         =   2
               End
               Begin CisText_V60.CisText iB1_Hinbn 
                  Height          =   360
                  Index           =   1
                  Left            =   450
                  TabIndex        =   13
                  Top             =   60
                  Width           =   3210
                  _ExtentX        =   5662
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   25
               End
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   360
                  Index           =   1
                  Left            =   4440
                  Top             =   60
                  Width           =   1020
                  _ExtentX        =   1799
                  _ExtentY        =   635
                  BackColor       =   16777152
                  ForeColor       =   0
                  Caption         =   "###,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -50
               End
               Begin Cis3D_v60.CIS3D Body_GNo 
                  Height          =   360
                  Index           =   1
                  Left            =   60
                  Top             =   60
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   635
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   "１"
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
                  cAlingnment     =   6
                  cBoderStyle     =   1
                  cPositionY      =   30
               End
               Begin CisText_V60.CisText iB1_MaiSu 
                  Height          =   360
                  Index           =   1
                  Left            =   5490
                  TabIndex        =   15
                  Top             =   60
                  Width           =   900
                  _ExtentX        =   1588
                  _ExtentY        =   635
                  cFormat         =   "##,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "##,###"
                  MaxLength       =   6
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#####"
                  cILength        =   5
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_HaSu 
                  Height          =   360
                  Index           =   1
                  Left            =   6390
                  TabIndex        =   16
                  Top             =   60
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_JyuSu 
                  Height          =   360
                  Index           =   1
                  Left            =   7380
                  TabIndex        =   17
                  Top             =   60
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   635
                  cFormat         =   "###,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###,###"
                  MaxLength       =   11
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#########"
                  cILength        =   9
                  cAlignment      =   1
               End
            End
            Begin Cis3D_v60.CIS3D BLB_Title 
               Height          =   480
               Left            =   60
               Top             =   60
               Width           =   8970
               _ExtentX        =   15822
               _ExtentY        =   847
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
               Begin Cis3D_v60.CIS3D CIS3D6 
                  Height          =   420
                  Index           =   1
                  Left            =   60
                  Top             =   30
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   16711680
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
                  cBoderColor1    =   8421504
                  cBoderColor2    =   -2147483643
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   7
                  cBoderStyle     =   1
                  cPositionX      =   15
                  cPositionY      =   15
               End
               Begin Cis3D_v60.CIS3D CIS3D6 
                  Height          =   420
                  Index           =   2
                  Left            =   450
                  Top             =   30
                  Width           =   3150
                  _ExtentX        =   5556
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   16711680
                  Caption         =   "品　　　番"
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
                  cPositionX      =   15
                  cPositionY      =   15
               End
               Begin Cis3D_v60.CIS3D CIS3D6 
                  Height          =   420
                  Index           =   7
                  Left            =   7380
                  Top             =   30
                  Width           =   1530
                  _ExtentX        =   2699
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   16711680
                  Caption         =   "受 注 数"
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
                  cPositionX      =   15
                  cPositionY      =   15
               End
               Begin Cis3D_v60.CIS3D CIS3D6 
                  Height          =   420
                  Index           =   6
                  Left            =   6390
                  Top             =   30
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   16711680
                  Caption         =   "端 数"
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
                  cPositionX      =   15
                  cPositionY      =   15
               End
               Begin Cis3D_v60.CIS3D CIS3D6 
                  Height          =   420
                  Index           =   5
                  Left            =   5475
                  Top             =   30
                  Width           =   900
                  _ExtentX        =   1588
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   16711680
                  Caption         =   "枚 数"
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
                  cPositionX      =   15
                  cPositionY      =   15
               End
               Begin Cis3D_v60.CIS3D CIS3D6 
                  Height          =   420
                  Index           =   4
                  Left            =   4440
                  Top             =   30
                  Width           =   1020
                  _ExtentX        =   1799
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   16711680
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
                  cBoderColor1    =   8421504
                  cBoderColor2    =   -2147483643
                  cFont3DColor1   =   -2147483643
                  cFont3DColor2   =   8421504
                  cAlingnment     =   7
                  cBoderStyle     =   1
                  cPositionX      =   15
                  cPositionY      =   15
               End
               Begin Cis3D_v60.CIS3D CIS3D6 
                  Height          =   420
                  Index           =   3
                  Left            =   3615
                  Top             =   30
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   741
                  BackColor       =   14737632
                  ForeColor       =   16711680
                  Caption         =   "背番号"
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
                  cPositionX      =   15
                  cPositionY      =   15
               End
            End
            Begin Cis3D_v60.CIS3D BLB_Line 
               Height          =   480
               Index           =   10
               Left            =   60
               Top             =   4860
               Width           =   8970
               _ExtentX        =   15822
               _ExtentY        =   847
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
               Begin CisText_V60.CisText iB1_Seban 
                  Height          =   360
                  Index           =   10
                  Left            =   3615
                  TabIndex        =   59
                  Top             =   60
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
                  IMEMode         =   2
               End
               Begin CisText_V60.CisText iB1_Hinbn 
                  Height          =   360
                  Index           =   10
                  Left            =   450
                  TabIndex        =   58
                  Top             =   60
                  Width           =   3210
                  _ExtentX        =   5662
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   25
               End
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   360
                  Index           =   10
                  Left            =   4440
                  Top             =   60
                  Width           =   1020
                  _ExtentX        =   1799
                  _ExtentY        =   635
                  BackColor       =   16777152
                  ForeColor       =   0
                  Caption         =   "###,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -50
               End
               Begin Cis3D_v60.CIS3D Body_GNo 
                  Height          =   360
                  Index           =   10
                  Left            =   60
                  Top             =   60
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   635
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   "10"
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
                  cAlingnment     =   6
                  cBoderStyle     =   1
                  cPositionY      =   30
               End
               Begin CisText_V60.CisText iB1_MaiSu 
                  Height          =   360
                  Index           =   10
                  Left            =   5490
                  TabIndex        =   60
                  Top             =   60
                  Width           =   900
                  _ExtentX        =   1588
                  _ExtentY        =   635
                  cFormat         =   "##,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "##,###"
                  MaxLength       =   6
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#####"
                  cILength        =   5
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_HaSu 
                  Height          =   360
                  Index           =   10
                  Left            =   6390
                  TabIndex        =   61
                  Top             =   60
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_JyuSu 
                  Height          =   360
                  Index           =   10
                  Left            =   7380
                  TabIndex        =   62
                  Top             =   60
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   635
                  cFormat         =   "###,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###,###"
                  MaxLength       =   11
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#########"
                  cILength        =   9
                  cAlignment      =   1
               End
            End
            Begin Cis3D_v60.CIS3D BLB_Line 
               Height          =   480
               Index           =   9
               Left            =   60
               Top             =   4380
               Width           =   8970
               _ExtentX        =   15822
               _ExtentY        =   847
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
               Begin CisText_V60.CisText iB1_Seban 
                  Height          =   360
                  Index           =   9
                  Left            =   3615
                  TabIndex        =   54
                  Top             =   60
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
                  IMEMode         =   2
               End
               Begin CisText_V60.CisText iB1_Hinbn 
                  Height          =   360
                  Index           =   9
                  Left            =   450
                  TabIndex        =   53
                  Top             =   60
                  Width           =   3210
                  _ExtentX        =   5662
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   25
               End
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   360
                  Index           =   9
                  Left            =   4440
                  Top             =   60
                  Width           =   1020
                  _ExtentX        =   1799
                  _ExtentY        =   635
                  BackColor       =   16777152
                  ForeColor       =   0
                  Caption         =   "###,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -50
               End
               Begin Cis3D_v60.CIS3D Body_GNo 
                  Height          =   360
                  Index           =   9
                  Left            =   60
                  Top             =   60
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   635
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   "９"
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
                  cAlingnment     =   6
                  cBoderStyle     =   1
                  cPositionY      =   30
               End
               Begin CisText_V60.CisText iB1_MaiSu 
                  Height          =   360
                  Index           =   9
                  Left            =   5490
                  TabIndex        =   55
                  Top             =   60
                  Width           =   900
                  _ExtentX        =   1588
                  _ExtentY        =   635
                  cFormat         =   "##,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "##,###"
                  MaxLength       =   6
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#####"
                  cILength        =   5
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_HaSu 
                  Height          =   360
                  Index           =   9
                  Left            =   6390
                  TabIndex        =   56
                  Top             =   60
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_JyuSu 
                  Height          =   360
                  Index           =   9
                  Left            =   7380
                  TabIndex        =   57
                  Top             =   60
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   635
                  cFormat         =   "###,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###,###"
                  MaxLength       =   11
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#########"
                  cILength        =   9
                  cAlignment      =   1
               End
            End
            Begin Cis3D_v60.CIS3D BLB_Line 
               Height          =   480
               Index           =   8
               Left            =   60
               Top             =   3900
               Width           =   8970
               _ExtentX        =   15822
               _ExtentY        =   847
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
               Begin CisText_V60.CisText iB1_Seban 
                  Height          =   360
                  Index           =   8
                  Left            =   3615
                  TabIndex        =   49
                  Top             =   60
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
                  IMEMode         =   2
               End
               Begin CisText_V60.CisText iB1_Hinbn 
                  Height          =   360
                  Index           =   8
                  Left            =   450
                  TabIndex        =   48
                  Top             =   60
                  Width           =   3210
                  _ExtentX        =   5662
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   25
               End
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   360
                  Index           =   8
                  Left            =   4440
                  Top             =   60
                  Width           =   1020
                  _ExtentX        =   1799
                  _ExtentY        =   635
                  BackColor       =   16777152
                  ForeColor       =   0
                  Caption         =   "###,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -50
               End
               Begin Cis3D_v60.CIS3D Body_GNo 
                  Height          =   360
                  Index           =   8
                  Left            =   60
                  Top             =   60
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   635
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   "８"
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
                  cAlingnment     =   6
                  cBoderStyle     =   1
                  cPositionY      =   30
               End
               Begin CisText_V60.CisText iB1_MaiSu 
                  Height          =   360
                  Index           =   8
                  Left            =   5490
                  TabIndex        =   50
                  Top             =   60
                  Width           =   900
                  _ExtentX        =   1588
                  _ExtentY        =   635
                  cFormat         =   "##,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "##,###"
                  MaxLength       =   6
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#####"
                  cILength        =   5
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_HaSu 
                  Height          =   360
                  Index           =   8
                  Left            =   6390
                  TabIndex        =   51
                  Top             =   60
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_JyuSu 
                  Height          =   360
                  Index           =   8
                  Left            =   7380
                  TabIndex        =   52
                  Top             =   60
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   635
                  cFormat         =   "###,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###,###"
                  MaxLength       =   11
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#########"
                  cILength        =   9
                  cAlignment      =   1
               End
            End
            Begin Cis3D_v60.CIS3D BLB_Line 
               Height          =   480
               Index           =   7
               Left            =   60
               Top             =   3420
               Width           =   8970
               _ExtentX        =   15822
               _ExtentY        =   847
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
               Begin CisText_V60.CisText iB1_Seban 
                  Height          =   360
                  Index           =   7
                  Left            =   3615
                  TabIndex        =   44
                  Top             =   60
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
                  IMEMode         =   2
               End
               Begin CisText_V60.CisText iB1_Hinbn 
                  Height          =   360
                  Index           =   7
                  Left            =   450
                  TabIndex        =   43
                  Top             =   60
                  Width           =   3210
                  _ExtentX        =   5662
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   25
               End
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   360
                  Index           =   7
                  Left            =   4440
                  Top             =   60
                  Width           =   1020
                  _ExtentX        =   1799
                  _ExtentY        =   635
                  BackColor       =   16777152
                  ForeColor       =   0
                  Caption         =   "###,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -50
               End
               Begin Cis3D_v60.CIS3D Body_GNo 
                  Height          =   360
                  Index           =   7
                  Left            =   60
                  Top             =   60
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   635
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   "７"
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
                  cAlingnment     =   6
                  cBoderStyle     =   1
                  cPositionY      =   30
               End
               Begin CisText_V60.CisText iB1_MaiSu 
                  Height          =   360
                  Index           =   7
                  Left            =   5490
                  TabIndex        =   45
                  Top             =   60
                  Width           =   900
                  _ExtentX        =   1588
                  _ExtentY        =   635
                  cFormat         =   "##,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "##,###"
                  MaxLength       =   6
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#####"
                  cILength        =   5
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_HaSu 
                  Height          =   360
                  Index           =   7
                  Left            =   6390
                  TabIndex        =   46
                  Top             =   60
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_JyuSu 
                  Height          =   360
                  Index           =   7
                  Left            =   7380
                  TabIndex        =   47
                  Top             =   60
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   635
                  cFormat         =   "###,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###,###"
                  MaxLength       =   11
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#########"
                  cILength        =   9
                  cAlignment      =   1
               End
            End
            Begin Cis3D_v60.CIS3D BLB_Line 
               Height          =   480
               Index           =   6
               Left            =   60
               Top             =   2940
               Width           =   8970
               _ExtentX        =   15822
               _ExtentY        =   847
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
               Begin CisText_V60.CisText iB1_Seban 
                  Height          =   360
                  Index           =   6
                  Left            =   3615
                  TabIndex        =   39
                  Top             =   60
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
                  IMEMode         =   2
               End
               Begin CisText_V60.CisText iB1_Hinbn 
                  Height          =   360
                  Index           =   6
                  Left            =   450
                  TabIndex        =   38
                  Top             =   60
                  Width           =   3210
                  _ExtentX        =   5662
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   25
               End
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   360
                  Index           =   6
                  Left            =   4440
                  Top             =   60
                  Width           =   1020
                  _ExtentX        =   1799
                  _ExtentY        =   635
                  BackColor       =   16777152
                  ForeColor       =   0
                  Caption         =   "###,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -50
               End
               Begin Cis3D_v60.CIS3D Body_GNo 
                  Height          =   360
                  Index           =   6
                  Left            =   60
                  Top             =   60
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   635
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   "６"
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
                  cAlingnment     =   6
                  cBoderStyle     =   1
                  cPositionY      =   30
               End
               Begin CisText_V60.CisText iB1_MaiSu 
                  Height          =   360
                  Index           =   6
                  Left            =   5490
                  TabIndex        =   40
                  Top             =   60
                  Width           =   900
                  _ExtentX        =   1588
                  _ExtentY        =   635
                  cFormat         =   "##,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "##,###"
                  MaxLength       =   6
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#####"
                  cILength        =   5
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_HaSu 
                  Height          =   360
                  Index           =   6
                  Left            =   6390
                  TabIndex        =   41
                  Top             =   60
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_JyuSu 
                  Height          =   360
                  Index           =   6
                  Left            =   7380
                  TabIndex        =   42
                  Top             =   60
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   635
                  cFormat         =   "###,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###,###"
                  MaxLength       =   11
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#########"
                  cILength        =   9
                  cAlignment      =   1
               End
            End
            Begin Cis3D_v60.CIS3D BLB_Line 
               Height          =   480
               Index           =   5
               Left            =   60
               Top             =   2460
               Width           =   8970
               _ExtentX        =   15822
               _ExtentY        =   847
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
               Begin CisText_V60.CisText iB1_Seban 
                  Height          =   360
                  Index           =   5
                  Left            =   3615
                  TabIndex        =   34
                  Top             =   60
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
                  IMEMode         =   2
               End
               Begin CisText_V60.CisText iB1_Hinbn 
                  Height          =   360
                  Index           =   5
                  Left            =   450
                  TabIndex        =   33
                  Top             =   60
                  Width           =   3210
                  _ExtentX        =   5662
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   25
               End
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   360
                  Index           =   5
                  Left            =   4440
                  Top             =   60
                  Width           =   1020
                  _ExtentX        =   1799
                  _ExtentY        =   635
                  BackColor       =   16777152
                  ForeColor       =   0
                  Caption         =   "###,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -50
               End
               Begin Cis3D_v60.CIS3D Body_GNo 
                  Height          =   360
                  Index           =   5
                  Left            =   60
                  Top             =   60
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   635
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   "５"
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
                  cAlingnment     =   6
                  cBoderStyle     =   1
                  cPositionY      =   30
               End
               Begin CisText_V60.CisText iB1_MaiSu 
                  Height          =   360
                  Index           =   5
                  Left            =   5490
                  TabIndex        =   35
                  Top             =   60
                  Width           =   900
                  _ExtentX        =   1588
                  _ExtentY        =   635
                  cFormat         =   "##,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "##,###"
                  MaxLength       =   6
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#####"
                  cILength        =   5
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_HaSu 
                  Height          =   360
                  Index           =   5
                  Left            =   6390
                  TabIndex        =   36
                  Top             =   60
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_JyuSu 
                  Height          =   360
                  Index           =   5
                  Left            =   7380
                  TabIndex        =   37
                  Top             =   60
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   635
                  cFormat         =   "###,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###,###"
                  MaxLength       =   11
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#########"
                  cILength        =   9
                  cAlignment      =   1
               End
            End
            Begin Cis3D_v60.CIS3D BLB_Line 
               Height          =   480
               Index           =   4
               Left            =   60
               Top             =   1980
               Width           =   8970
               _ExtentX        =   15822
               _ExtentY        =   847
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
               Begin CisText_V60.CisText iB1_Seban 
                  Height          =   360
                  Index           =   4
                  Left            =   3615
                  TabIndex        =   29
                  Top             =   60
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
                  IMEMode         =   2
               End
               Begin CisText_V60.CisText iB1_Hinbn 
                  Height          =   360
                  Index           =   4
                  Left            =   450
                  TabIndex        =   28
                  Top             =   60
                  Width           =   3210
                  _ExtentX        =   5662
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   25
               End
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   360
                  Index           =   4
                  Left            =   4440
                  Top             =   60
                  Width           =   1020
                  _ExtentX        =   1799
                  _ExtentY        =   635
                  BackColor       =   16777152
                  ForeColor       =   0
                  Caption         =   "###,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -50
               End
               Begin Cis3D_v60.CIS3D Body_GNo 
                  Height          =   360
                  Index           =   4
                  Left            =   60
                  Top             =   60
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   635
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   "４"
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
                  cAlingnment     =   6
                  cBoderStyle     =   1
                  cPositionY      =   30
               End
               Begin CisText_V60.CisText iB1_MaiSu 
                  Height          =   360
                  Index           =   4
                  Left            =   5490
                  TabIndex        =   30
                  Top             =   60
                  Width           =   900
                  _ExtentX        =   1588
                  _ExtentY        =   635
                  cFormat         =   "##,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "##,###"
                  MaxLength       =   6
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#####"
                  cILength        =   5
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_HaSu 
                  Height          =   360
                  Index           =   4
                  Left            =   6390
                  TabIndex        =   31
                  Top             =   60
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_JyuSu 
                  Height          =   360
                  Index           =   4
                  Left            =   7380
                  TabIndex        =   32
                  Top             =   60
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   635
                  cFormat         =   "###,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###,###"
                  MaxLength       =   11
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#########"
                  cILength        =   9
                  cAlignment      =   1
               End
            End
            Begin Cis3D_v60.CIS3D BLB_Line 
               Height          =   480
               Index           =   2
               Left            =   60
               Top             =   1020
               Width           =   8970
               _ExtentX        =   15822
               _ExtentY        =   847
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
               Begin CisText_V60.CisText iB1_Seban 
                  Height          =   360
                  Index           =   2
                  Left            =   3615
                  TabIndex        =   19
                  Top             =   60
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
                  IMEMode         =   2
               End
               Begin CisText_V60.CisText iB1_Hinbn 
                  Height          =   360
                  Index           =   2
                  Left            =   450
                  TabIndex        =   18
                  Top             =   60
                  Width           =   3210
                  _ExtentX        =   5662
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   25
               End
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   360
                  Index           =   2
                  Left            =   4440
                  Top             =   60
                  Width           =   1020
                  _ExtentX        =   1799
                  _ExtentY        =   635
                  BackColor       =   16777152
                  ForeColor       =   0
                  Caption         =   "###,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -50
               End
               Begin Cis3D_v60.CIS3D Body_GNo 
                  Height          =   360
                  Index           =   2
                  Left            =   60
                  Top             =   60
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   635
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   "２"
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
                  cAlingnment     =   6
                  cBoderStyle     =   1
                  cPositionY      =   30
               End
               Begin CisText_V60.CisText iB1_MaiSu 
                  Height          =   360
                  Index           =   2
                  Left            =   5490
                  TabIndex        =   20
                  Top             =   60
                  Width           =   900
                  _ExtentX        =   1588
                  _ExtentY        =   635
                  cFormat         =   "##,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "##,###"
                  MaxLength       =   6
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#####"
                  cILength        =   5
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_HaSu 
                  Height          =   360
                  Index           =   2
                  Left            =   6390
                  TabIndex        =   21
                  Top             =   60
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_JyuSu 
                  Height          =   360
                  Index           =   2
                  Left            =   7380
                  TabIndex        =   22
                  Top             =   60
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   635
                  cFormat         =   "###,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###,###"
                  MaxLength       =   11
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#########"
                  cILength        =   9
                  cAlignment      =   1
               End
            End
            Begin Cis3D_v60.CIS3D BLB_Line 
               Height          =   480
               Index           =   3
               Left            =   60
               Top             =   1500
               Width           =   8970
               _ExtentX        =   15822
               _ExtentY        =   847
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
               Begin CisText_V60.CisText iB1_Seban 
                  Height          =   360
                  Index           =   3
                  Left            =   3615
                  TabIndex        =   24
                  Top             =   60
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
                  IMEMode         =   2
               End
               Begin CisText_V60.CisText iB1_Hinbn 
                  Height          =   360
                  Index           =   3
                  Left            =   450
                  TabIndex        =   23
                  Top             =   60
                  Width           =   3210
                  _ExtentX        =   5662
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   25
               End
               Begin Cis3D_v60.CIS3D B1lb_Syuyo 
                  Height          =   360
                  Index           =   3
                  Left            =   4440
                  Top             =   60
                  Width           =   1020
                  _ExtentX        =   1799
                  _ExtentY        =   635
                  BackColor       =   16777152
                  ForeColor       =   0
                  Caption         =   "###,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -50
               End
               Begin Cis3D_v60.CIS3D Body_GNo 
                  Height          =   360
                  Index           =   3
                  Left            =   60
                  Top             =   60
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   635
                  BackColor       =   14737632
                  ForeColor       =   0
                  Caption         =   "３"
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
                  cAlingnment     =   6
                  cBoderStyle     =   1
                  cPositionY      =   30
               End
               Begin CisText_V60.CisText iB1_MaiSu 
                  Height          =   360
                  Index           =   3
                  Left            =   5490
                  TabIndex        =   25
                  Top             =   60
                  Width           =   900
                  _ExtentX        =   1588
                  _ExtentY        =   635
                  cFormat         =   "##,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "##,###"
                  MaxLength       =   6
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#####"
                  cILength        =   5
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_HaSu 
                  Height          =   360
                  Index           =   3
                  Left            =   6390
                  TabIndex        =   26
                  Top             =   60
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  cFormat         =   "###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###"
                  MaxLength       =   7
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "######"
                  cILength        =   6
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_JyuSu 
                  Height          =   360
                  Index           =   3
                  Left            =   7380
                  TabIndex        =   27
                  Top             =   60
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   635
                  cFormat         =   "###,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "###,###,###"
                  MaxLength       =   11
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#########"
                  cILength        =   9
                  cAlignment      =   1
               End
            End
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
      Caption         =   "【 受注画面入力 [号口]  】"
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
         Left            =   14280
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
         Left            =   11700
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
   Begin Cis3D_v60.CIS3D H1_Area1 
      Height          =   1035
      Left            =   2100
      Top             =   1500
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1826
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
      Begin Cis3D_v60.CIS3D CIS3D17 
         Height          =   870
         Left            =   90
         Top             =   90
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   1535
         ForeColor       =   16711680
         Caption         =   "受注区分"
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
         cPositionX      =   90
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "代替"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   3030
            TabIndex        =   6
            Top             =   480
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "スポット"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   3960
            TabIndex        =   3
            Top             =   120
            Width           =   1320
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "その他"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   3960
            TabIndex        =   7
            Top             =   480
            Width           =   1320
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "号試"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   3030
            TabIndex        =   2
            Top             =   120
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H008080FF&
            Caption         =   "通常"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1170
            TabIndex        =   0
            Top             =   120
            Value           =   -1  'True
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "臨時"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   1170
            TabIndex        =   4
            Top             =   480
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "特配"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   2100
            TabIndex        =   5
            Top             =   480
            Width           =   840
         End
         Begin VB.OptionButton H1Op_JKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "指示"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   2100
            TabIndex        =   1
            Top             =   120
            Width           =   840
         End
      End
   End
End
Attribute VB_Name = "CXJ0010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   受注画面入力 [単品]
'**       フォームID    :   CXJ0010
'**       処理概要      :
'**
'**       作  成  日    :   2004/03/22  By CIS SIMPR-A
'**       変  更  日    :   2004/04/16  By CIS - 納品書発行区分のセット値を変更（受注入力ワーク）
'**       変  更  日    :   2004/04/19  By CIS - 画面上の受注数の算出：収容数×枚数＋端数
'**       　  　  　                             ・変更点 >> 端数が０以外は枚数を－１する
'**       変  更  日    :   2004/04/21  By CIS - 納入番号が未入力時、品番取引先マスタの伝票種類が空白はエラーとする。
'**       　  　  　                             納入番号をセット時、納品書頁＆行もセット（自社伝票は発行しない場合）
'**       変  更  日    :   2004/05/28  By CIS - 背番号カナ入力・背番号半角チェック
'**       変  更  日    :   2004/06/24  By CIS - 受領書・出荷時発行の対応
'**       変  更  日    :   2004/09/01  By CIS - 受注数を入力可に変更
'**       変  更  日    :   2005/05/09  By CIS - 受注数を入力可に変更( iB1_JyuSu Index Error [ 0->10 ] )
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替

    Dim wEndChk         As Boolean
    Dim wIndex          As Integer
    Dim wJyuSu          As Long
    Dim wInput          As Boolean      ' 入力ﾌﾗｸﾞ

    Dim w_JKbn          As Byte         '>受注区分
    Dim w_DenKnd(10)    As String       '>伝票種類
'*---------------------------------------------<< 2004/04/21 Insert Start >>
    Dim w_PageNo        As Byte         '>納品書頁
    Dim w_PageGyo       As Byte         '>納品書行
'*---------------------------------------------<< 2004/04/21 Insert Start >>
    Dim wProcF          As Integer
    Dim Sav_Index       As Integer      ' 行コピー元行№
    Dim Sto_Index       As Integer      ' 行コピー先行№
    Dim wSetColor       As Long         '>DragDropｾｯﾄｶﾗｰ
    Dim wLosColor       As Long         '>DragDrop終了ｶﾗｰ
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

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

'   >> 名称マスタ確認
    Call MeisyoInit
'   >> 日付表示
    iH2_Ymd = Format(Date, "yyyymmdd")

    wSetColor = &HC0FFC0    ' DragDropｾｯﾄｶﾗｰ
    wLosColor = &HE0E0E0    ' DragDrop終了ｶﾗｰ

    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

    If InitProc Then        ' 初期処理（データ削除）
        wEndChk = True
    Else
        wEndChk = False
        Call PB_END_Click
    End If

    FormAct = Not FormAct
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
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:
                If PB_PCAN.Visible Then
                    Call PB_PCAN_Click
                Else
                                If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
                End If
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                      '【入力】
                                '  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT_Click
                                '  Dummy.Enabled = False
                                End If
           Case Else
                If ProcHB = "B1" And Shift = 4 Then
                    Select Case Key_Code
                           Case vbKeyC: Call Input_Alt_C    '>明細入力時の行複写 ( Alt + C )
                           Case vbKeyD: Call Input_Alt_D    '>明細入力時の行削除 ( Alt + D )
                           Case vbKeyE: Call Input_Alt_E    '>明細入力時の行追加 ( Alt + E )
                           Case vbKeyF:
                           Case Else:   Exit Function
                    End Select
                Else
                    Exit Function
                End If
    End Select

    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
'   #### 終了処理 ####
    Call EndProc
'   #### ﾃﾞｰﾀﾍﾞｰｽ切断 ####
    Call CisDB.DBDISConnect
'   #### ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去 ####
    Call CisFrm.UnLoadDisp
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
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call CisFrm.MousePT(11)
    If ProcHB = "H2" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("H", "H2")
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B1" Then
        ProcHB = "H2"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        If C1_Area1.Visible = True Then
            Call PB_PCAN_Click
        End If
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
    If PB_Look.Tag = "iH2_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH2_Ukeir" Then Call Look_Ukeir: Exit Sub
    If PB_Look.Tag Like "iB1_Hinbn*" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag Like "iB1_Seban*" Then Call Look_Hinbn: Exit Sub
'
    If PB_Look.Tag = "iH2_Ymd" Then iH2_Ymd.ShowCalender: Exit Sub
End Sub
'+----------------------------+
'+       取引先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 3
    RV_Call = "C"

    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
       iH2_Torcd = RTrim(RV_TorcdK)
       Call iH2_Torcd_LostFocus
    End If
End Sub
'+----------------------------+
'+       受入検索            +
'+----------------------------+
Private Sub Look_Ukeir()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 3
    RV_Torcd = RTrim(iH2_Torcd)
    RV_Kojyo = ""

    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
       iH2_Torcd = RTrim(RV_TorcdK)
       iH2_Ukeir = RTrim(RV_Ukeir)
       Call iH2_Torcd_LostFocus
    End If
End Sub
'+----------------------------+
'+       品番検索            +
'+----------------------------+
Private Sub Look_Hinbn()
    wIndex = CisFun.Val2(Right(PB_Look.Tag, 2))

    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 0
    RVI_SK = 9
    RVI_Torcd = RTrim(iH2_Torcd)
    RVI_Ukeir = RTrim(iH2_Ukeir)
    RVI_Seban = ""
    RVI_Tehai = ""

    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
       iB1_Seban(wIndex) = RTrim(RV_Seban)
       iB1_Hinbn(wIndex) = RTrim(RV_Hinbn)
    End If
End Sub
'+----------------------------------------+
'+  H1Op_Jkbn ( Click )      受注区分
'+----------------------------------------+
Private Sub H1Op_JKbn_Click(Index As Integer)
    For wIndex = 1 To 8
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'      H1Op_JKbn(wIndex).BackColor = &H8000000F
       H1Op_JKbn(wIndex).BackColor = &HC0C0C0
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
    Next wIndex
    H1Op_JKbn(Index).BackColor = &H8080FF
    Select Case Index
        Case 1:     Back_Ymd.Caption = "納 入 日"
        Case Else:  Back_Ymd.Caption = "納 入 日"
    End Select
End Sub
'+-----------------------------+
'+  iH2_Torcd - 納入先
'+-----------------------------+
Private Sub iH2_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH2_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    If RTrim(iH2_Torcd) <> "" Then
        If SqlSelectTRM(iH2_Torcd) Then
            H2lb_Tornm.Caption = RTrim(TRM.略称)
            H2lb_Kojyonm.Caption = RTrim(TRM.工場名)
        Else
            H2lb_Tornm.Caption = "*** ﾐﾄｳﾛｸ ***"
            H2lb_Kojyonm.Caption = ""
        End If
    Else
        H2lb_Tornm.Caption = ""
        H2lb_Kojyonm.Caption = ""
    End If
End Sub
'+-----------------------------+
'+  iH2_Ymd - 受注日・納入日
'+-----------------------------+
Private Sub iH2_Ymd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH2_Ymd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH2_Ukeir - 受入
'+-----------------------------+
Private Sub iH2_Ukeir_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH2_Ukeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Hinbn - 品番
'+-----------------------------+
Private Sub iB1_Hinbn_GotFocus(Index As Integer)
    PB_Look.Tag = ActiveControl.Name & Format(Index, "00")
    PB_Look.Visible = True
End Sub
Private Sub iB1_Hinbn_LostFocus(Index As Integer)
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
'   >> 背番号ｾｯﾄ
    B1lb_Syuyo(Index) = ""
    If RTrim(iB1_Hinbn(Index)) <> "" Then
        If SqlSelectHTM(iB1_Hinbn(Index), "") Then
            iB1_Seban(Index) = RTrim(HTM.背番号)
            Call Syuyo_Set(Index)
        End If
    End If
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'   Call JyuSu_Set(Index)
'   受注数を元に、枚数＆端数を算出
    Call Suryo_Set(Index, "S")
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
End Sub
'+-----------------------------+
'+  iB1_Seban - 背番号
'+-----------------------------+
Private Sub iB1_Seban_GotFocus(Index As Integer)
    PB_Look.Tag = ActiveControl.Name & Format(Index, "00")
    PB_Look.Visible = True
End Sub
Private Sub iB1_Seban_LostFocus(Index As Integer)
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
'   >> 品番ｾｯﾄ
    If RTrim(iB1_Seban(Index)) <> "" Then
        If SqlSelectHTM("", iB1_Seban(Index)) Then
            iB1_Hinbn(Index) = RTrim(HTM.品番)
            Call Syuyo_Set(Index)
        End If
    End If
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'   Call JyuSu_Set(Index)
'   受注数を元に、枚数＆端数を算出
    Call Suryo_Set(Index, "S")
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
End Sub
'+-----------------------------+
'+  iB1_MaiSu - 枚数
'+-----------------------------+
Private Sub iB1_MaiSu_LostFocus(Index As Integer)
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'   Call JyuSu_Set(Index)
'   枚数＆端数を元に、受注数を算出
    Call Suryo_Set(Index, "M")
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
End Sub
'+-----------------------------+
'+  iB1_HaSu - 端数
'+-----------------------------+
Private Sub iB1_HaSu_LostFocus(Index As Integer)
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'   Call JyuSu_Set(Index)
'   枚数＆端数を元に、受注数を算出
    Call Suryo_Set(Index, "M")
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
End Sub
'*↓↓*-------------------------------------------<< 2004/09/01 >> Insert Start
'+-----------------------------+
'+  iB1_JyuSu - 受注数
'+-----------------------------+
Private Sub iB1_JyuSu_LostFocus(Index As Integer)
'   受注数を元に、枚数＆端数を算出
    Call Suryo_Set(Index, "S")
End Sub
'*↑↑*-------------------------------------------<< 2004/09/01 >> Insert End

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
    Dim H2Mode      As Boolean
    Dim H2Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long

    H1Mode = False: H1Color = gPLostSel
    H2Mode = False: H2Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1":   H1Mode = True:  H1Color = gPGotSel
       Case "H2":   H2Mode = True:  H2Color = gPGotSel
       Case "B1":   B1Mode = True:  B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode:  H1_Area1.BackColor = H1Color
    H2_Area1.Enabled = H2Mode:  H2_Area1.BackColor = H2Color
    H2_Area2.Enabled = H2Mode:  H2_Area2.BackColor = H2Color
    B1_Area1.Enabled = B1Mode:  B1_Area1.BackColor = B1Color

    PB_Look.Visible = False
    PB_CAN.Visible = Not H1Mode
   'PB_END.Visible = H1Mode

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
            If .Name Like "i" & wbc_Enm & "_*" Then .text = "": .Enabled = True
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Cmb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "H" Then
'       >> 日付表示
        iH2_Ymd = Format(Date, "yyyymmdd")
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "H2"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "H2" Then
        If RTrim(iH2_Bin) = "" Then iH2_Bin = "01"
        If Not Head2Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        iH2_Nouban = ""
        ProcHB = "H2"
        Call DispChange(ProcHB)
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

    For wIndex = 1 To 8
        If H1Op_JKbn(wIndex).Value Then
            w_JKbn = CisFun.Val2(H1Op_JKbn(wIndex).Tag)
            Exit For
        End If
    Next wIndex

    Head1Chk = True
End Function
'****************************
'*      HEAD2ﾁｪｯｸ処理
'****************************
Private Function Head2Chk() As Boolean
    Head2Chk = False

    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With

'   << 納入先 >>
    If Trim(iH2_Torcd) = "" Then
        CisFun.MB_MSG(2) = "    取引先を入力して下さい            "
        CisFun.MBOX
        iH2_Torcd.SetFocus
        Exit Function
    Else
        If Not SqlSelectTRM(iH2_Torcd) Then
            CisFun.MB_MSG(2) = "    取引先　未登録！（取引先マスタ）            "
            CisFun.MBOX
            iH2_Torcd.SetFocus
            Exit Function
        End If
    End If
'   << 受注日 or 納入日 >>
    If RTrim(iH2_Ymd) = "" Then
        CisFun.MB_MSG(2) = "    " & Replace(Back_Ymd.Caption, " ", "") & "を入力して下さい。      "
        CisFun.MBOX
        iH2_Ymd.SetFocus
        Exit Function
    Else
        If Not iH2_Ymd.cChkResult Then
            CisFun.MB_MSG(2) = "    " & Replace(Back_Ymd.Caption, " ", "") & "を正しく入力して下さい            "
            CisFun.MBOX
            iH2_Ymd.SetFocus
            Exit Function
        End If
    End If
'   << 便 >>
    If Trim(iH2_Bin) = "" Then
        CisFun.MB_MSG(2) = "    便を入力して下さい            "
        CisFun.MBOX
        iH2_Bin.SetFocus
        Exit Function
    End If
'   << 受入 >>
    If RTrim(iH2_Ukeir) = "" Then
        CisFun.MB_MSG(2) = "    受入を入力して下さい            "
        CisFun.MBOX
        iH2_Ukeir.SetFocus
        Exit Function
    Else
        TUM.取引先CD = RTrim(iH2_Torcd)
        TUM.受入 = RTrim(iH2_Ukeir)
        TUM.取引先区分 = 0
        If Not TUMRead("", 1) Then
            CisFun.MB_MSG(2) = "    受入　未登録！（取引先受入マスタ）            "
            CisFun.MBOX
            iH2_Ukeir.SetFocus
            Exit Function
        End If
    End If
'   << 納入番号 >>
    If RTrim(iH2_Nouban) = "" Then
        If RTrim(TRM.伝票種類) = "" Then
            CisFun.MB_MSG(2) = "    納入番号を入力して下さい            "
            CisFun.MBOX
            iH2_Nouban.SetFocus
            Exit Function
        End If
    End If

    Head2Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False

    Erase w_DenKnd

    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With

'   <<< 最初に品番・背番号を表示する >>>
    wInput = False
    For wIndex = 1 To 10
        If RTrim(iB1_Hinbn(wIndex)) <> "" Or _
           RTrim(iB1_Seban(wIndex)) <> "" Then
            wInput = True
'           >> 品番入力、背番号未入力時
            If RTrim(iB1_Hinbn(wIndex)) <> "" And RTrim(iB1_Seban(wIndex)) = "" Then
                If SqlSelectHTM(iB1_Hinbn(wIndex), "") Then
                    iB1_Seban(wIndex) = RTrim(HTM.背番号)
                    Call Syuyo_Set(wIndex)
                End If
            End If
'           >> 品番未入力、背番号入力時
            If RTrim(iB1_Hinbn(wIndex)) = "" And RTrim(iB1_Seban(wIndex)) <> "" Then
                If SqlSelectHTM("", iB1_Seban(wIndex)) Then
                    iB1_Hinbn(wIndex) = RTrim(HTM.品番)
                    Call Syuyo_Set(wIndex)
                End If
            End If
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'           受注数の計算
'           Call JyuSu_Set(wIndex)
'           受注数を元に、枚数＆端数を算出
            Call Suryo_Set(wIndex, "S")
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
        End If
    Next wIndex
    If Not wInput Then
        CisFun.MB_MSG(2) = "    データを入力して下さい            "
        CisFun.MBOX
        iB1_Hinbn(1).SetFocus
        Exit Function
    End If
'
    For wIndex = 1 To 10
'       >> 入力判定
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'       If RTrim(iB1_Hinbn(wIndex)) <> "" Or _
'          RTrim(iB1_Seban(wIndex)) <> "" Or _
'          CisFun.Val2(iB1_MaiSu(wIndex)) <> 0 Or _
'          CisFun.Val2(iB1_HaSu(wIndex)) <> 0 Then
'
        If RTrim(iB1_Hinbn(wIndex)) <> "" Or _
           RTrim(iB1_Seban(wIndex)) <> "" Or _
           CisFun.Val2(iB1_MaiSu(wIndex)) <> 0 Or _
           CisFun.Val2(iB1_HaSu(wIndex)) <> 0 Or _
           CisFun.Val2(iB1_JyuSu(wIndex)) <> 0 Then
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
'           ## 明細行入力 ##
'           << 重複チェック >>
            For gIndex = wIndex + 1 To 10
                If RTrim(iB1_Hinbn(wIndex)) = RTrim(iB1_Hinbn(gIndex)) Then
                    CisFun.MB_MSG(1) = "    品番が重複しています。            "
                    CisFun.MB_MSG(3) = "    " & wIndex & "行目と " & gIndex & "行目            "
                    CisFun.MBOX
                    iB1_Hinbn(wIndex).SetFocus
                    Exit Function
                End If
            Next gIndex
'           << 品番 >>
            If RTrim(iB1_Hinbn(wIndex)) = "" Then
                CisFun.MB_MSG(2) = "    品番を入力して下さい            "
                CisFun.MBOX
                iB1_Hinbn(wIndex).SetFocus
                Exit Function
            Else
                If Not SqlSelectHNM(iB1_Hinbn(wIndex)) Then
                    CisFun.MB_MSG(2) = "    品番　未登録！（品番マスタ）            "
                    CisFun.MBOX
                    iB1_Hinbn(wIndex).SetFocus
                    Exit Function
                End If
                If Not SqlSelectHTM(iB1_Hinbn(wIndex), "") Then
                    CisFun.MB_MSG(2) = "    品番　未登録！（品番取引先マスタ）            "
                    CisFun.MBOX
                    iB1_Hinbn(wIndex).SetFocus
                    Exit Function
                End If
            End If
'*---------------------------------------------<< 2004/04/21 Insert Start >>
            If RTrim(iH2_Nouban) = "" Then
                If RTrim(HTM.伝票種類) = "" Then
                    CisFun.MB_MSG(2) = "    納入番号が未入力の為、この品番は入力できません！            "
                    CisFun.MBOX
                    iB1_Hinbn(wIndex).SetFocus
                    Exit Function
                End If
            End If
'*---------------------------------------------<< 2004/04/21 Insert End >>
        ' 背番号半角チェック
            If RTrim(iB1_Seban(wIndex)) <> "" And Not HalfSizeChk(iB1_Seban(wIndex)) Then
                With CisFun
                     .MB_Lines = 4
                     .MB_MSG(2) = "    背番号は半角にて入力して下さい            "
                     .MB_Title = ""
                     .MB_Button = Error
                     .MBOX
                End With
                iB1_Seban(wIndex).SetFocus
                Exit Function
            End If
            w_DenKnd(wIndex) = RTrim(HTM.伝票種類)
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
''          << 枚数・端数 >>
'           If CisFun.Val2(iB1_MaiSu(wIndex)) = 0 And _
'              CisFun.Val2(iB1_HaSu(wIndex)) = 0 Then
'               CisFun.MB_MSG(2) = "    枚数・端数を入力して下さい            "
'               CisFun.MBOX
'               iB1_MaiSu(wIndex).SetFocus
'               Exit Function
'           End If
''          << 収容数 <= 端数 >>
'           If CisFun.Val2(B1lb_Syuyo(wIndex)) <= CisFun.Val2(iB1_HaSu(wIndex)) Then
'               CisFun.MB_MSG(2) = "    端数　エラー！（端数 ≧ 収容数）            "
'               CisFun.MBOX
'               iB1_HaSu(wIndex).SetFocus
'               Exit Function
'           End If
'
'           << 枚数・端数・受注数 >>
            If CisFun.Val2(iB1_MaiSu(wIndex)) = 0 And _
               CisFun.Val2(iB1_HaSu(wIndex)) = 0 And _
               CisFun.Val2(iB1_JyuSu(wIndex)) = 0 Then
                CisFun.MB_MSG(2) = "    枚数・端数・受注数を入力して下さい            "
                CisFun.MBOX
                iB1_MaiSu(wIndex).SetFocus
                Exit Function
            End If
'           << 収容数 <= 端数 >>
            If CisFun.Val2(B1lb_Syuyo(wIndex)) <> 0 Then
                If CisFun.Val2(B1lb_Syuyo(wIndex)) <= CisFun.Val2(iB1_HaSu(wIndex)) Then
                    CisFun.MB_MSG(2) = "    端数　エラー！（端数 ≧ 収容数）            "
                    CisFun.MBOX
                    iB1_HaSu(wIndex).SetFocus
                    Exit Function
                End If
            End If
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End

        End If

    Next wIndex
'
'   >> 明細情報重複チェック
    If Not Daburi_Check Then

        CXJ0012.Show vbModal
        Unload CXJ0012
        Set CXJ0012 = Nothing

        If Not RV_Rtn Then Exit Function
    End If

    Body1Chk = True
End Function
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False

' ( 確認メッセージ )
    With CisFun
        .MB_Lines = 7
        .MB_MSG(1) = "     【 追加処理確認 】            "
        .MB_MSG(3) = "    納入先 ＝ " & StrConv(iH2_Torcd, vbWide)
        .MB_MSG(4) = "    納入日－便 ＝ " & StrConv(iH2_Ymd, vbWide) & "-" & StrConv(iH2_Bin, vbWide)
        .MB_MSG(5) = "    受入 ＝ " & StrConv(iH2_Ukeir, vbWide)
        .MB_MSG(6) = "    納番 ＝ " & StrConv(iH2_Nouban, vbWide)
        .MB_Title = "追加処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err

    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

'   <><><> 受注管理NOの獲得 <><><>
    Call ItemsClearJNW
    With CisDB
        .SQL = "番号獲得"
        .StoadoCount = 3
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = OutPut用
        .ParaNo = 3: .ParaIO = Return用
    '
        .ParaNo = 1: .ParaValue = 11
    '
        If Not .DBStored Then
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(2) = "    受注管理NOの獲得に失敗しました            "
            CisFun.MB_Title = "採番処理"
            CisFun.MB_Button = Error
            CisFun.MBOX
            Exit Function
        Else
            .ParaNo = 2: JNW.受注管理NO = .ParaValue
        End If
    End With
    
    For wIndex = 1 To 10
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'       If RTrim(iB1_Hinbn(wIndex)) <> "" Or _
'          RTrim(iB1_Seban(wIndex)) <> "" Or _
'          CisFun.Val2(iB1_MaiSu(wIndex)) <> 0 Or _
'          CisFun.Val2(iB1_HaSu(wIndex)) <> 0 Then
'
        If RTrim(iB1_Hinbn(wIndex)) <> "" Or _
           RTrim(iB1_Seban(wIndex)) <> "" Or _
           CisFun.Val2(iB1_MaiSu(wIndex)) <> 0 Or _
           CisFun.Val2(iB1_HaSu(wIndex)) <> 0 Or _
           CisFun.Val2(iB1_JyuSu(wIndex)) <> 0 Then
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
        '   >> DBｾｯﾄ
            If Not DBSet_JNW(wIndex) Then GoTo DBPut_Err
        '   >> 出力
            Call JNWInsert
        End If
    Next wIndex

    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

    On Error GoTo 0

    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+------------------------------------+
'+      セット - 受注入力ワーク
'+------------------------------------+
Private Function DBSet_JNW(Index As Integer)
    DBSet_JNW = False

'''    Call ItemsClearJNW
'''    With CisDB
'''    '   <><><> 受注管理NOの獲得 <><><>
'''        .SQL = "番号獲得"
'''        .StoadoCount = 3
'''        .ParaNo = 0: .ParaIO = Return用
'''        .ParaNo = 1: .ParaIO = Input用
'''        .ParaNo = 2: .ParaIO = OutPut用
'''        .ParaNo = 3: .ParaIO = Return用
'''    '
'''        .ParaNo = 1: .ParaValue = 11
'''    '
'''        If Not .DBStored Then
'''            CisFun.MB_Lines = 4
'''            CisFun.MB_MSG(2) = "    受注管理NOの獲得に失敗しました            "
'''            CisFun.MB_Title = "採番処理"
'''            CisFun.MB_Button = Error
'''            CisFun.MBOX
'''            Exit Function
'''        Else
'''            .ParaNo = 2: JNW.受注管理NO = .ParaValue
'''        End If
'''    End With
    With JNW
        .受注管理行NO = Index
        .端末番号 = "' + substring(host_name(),1,20) + '"
        .受注区分 = w_JKbn
        .受注日 = Format(Date, "YYYYMMDD")
        .受注便 = ""
        .納入日 = iH2_Ymd
        .納入便 = iH2_Bin
        .納入先 = RTrim(iH2_Torcd)
        .出荷予定日 = .納入日
        .出荷予定便 = .納入便
        .受入 = RTrim(iH2_Ukeir)
        .機番 = ""
'*---------------------------------------------<< 2004/04/21 Update Start >>
'       Select Case w_DenKnd(Index)
'           Case "":    .納入番号 = RTrim(iH2_Nouban)
'           Case Else:  .納入番号 = ""
'       End Select
'       .納品書頁 = 0
'       .納品書行 = 0
'
'       <><><> 品番取引先マスタの伝票種類は？ <><><>
        Select Case w_DenKnd(Index)
            Case ""
'               >> 自社伝票：発行無
                Call SqlSelectJYT
                .納入番号 = RTrim(iH2_Nouban)
                .納品書頁 = w_PageNo
                .納品書行 = w_PageGyo
                .納品書発行区分 = 9         '>対象外
            Case Else
'               >> 自社伝票：発行有
                .納入番号 = ""
                .納品書頁 = 0
                .納品書行 = 0
                .納品書発行区分 = 0         '>未発行
        End Select
'*---------------------------------------------<< 2004/04/21 Update End >>
        .分納回数 = 0
        .品番 = RTrim(iB1_Hinbn(Index))
        .背番号 = RTrim(iB1_Seban(Index))
        .枚数 = CisFun.Val2(iB1_MaiSu(Index))
        .収容数 = CisFun.Val2(B1lb_Syuyo(Index))
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'       .納入数 = CisFun.Val2(B1lb_JyuSu(Index))
        .納入数 = CisFun.Val2(iB1_JyuSu(Index))
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
        .端数 = CisFun.Val2(iB1_HaSu(Index))
'*---------------------------------------------<< 2004/04/19 Delete Start >>
'       If .端数 <> 0 Then .枚数 = .枚数 + 1
'*---------------------------------------------<< 2004/04/19 Delete End >>
        .出荷指示書NO = 0
        .出荷指示書行NO = 0
        .抜取枚数 = 0
'*---------------------------------------------<< 2004/04/21 Delete Start >>
''*--------------------------------------------<< 2004/04/16 Update Start >>
''      .納品書発行区分 = 0
'       Select Case w_DenKnd(Index)
'           Case "":    .納品書発行区分 = 9     '>対象外
'           Case Else:  .納品書発行区分 = 0     '>未発行
'       End Select
''*--------------------------------------------<< 2004/04/16 Update End >>
'*---------------------------------------------<< 2004/04/21 Delete End >>
        Select Case .受注区分
            Case 2, 4, 5, 6, 7, 8
                .エフ発行枚数 = .枚数
                .エフ発行端数 = .端数
            Case Else
                .エフ発行枚数 = 0
                .エフ発行端数 = 0
        End Select
        .作成種別 = "G"
        .累積区分 = 0
        .納品書管理NO = 0
        .元出荷指示書NO = 0
        .元出荷指示書行NO = 0
    End With

    DBSet_JNW = True
End Function

'******************************************************************
'*      収容数の表示
'******************************************************************
Private Sub Syuyo_Set(Index As Integer)
    If Trim(HTM.変更日) = "" Then
        B1lb_Syuyo(Index) = Format(HTM.収容数1, "##,###")
    Else
        If (HTM.変更日 > iH2_Ymd) Or _
           (HTM.変更日 = iH2_Ymd And HTM.変更便 > iH2_Bin) Then
            B1lb_Syuyo(Index) = Format(HTM.収容数1, "##,###")
        Else
            B1lb_Syuyo(Index) = Format(HTM.収容数2, "##,###")
        End If
    End If
End Sub
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
''******************************************************************
''*      受注数の表示
''******************************************************************
'Private Sub JyuSu_Set(Index As Integer)
''*---------------------------------------------<< 2004/04/19 Update Start >>
''   wJyuSu = CisFun.Val2(iB1_MaiSu(Index)) * CisFun.Val2(B1lb_Syuyo(Index)) _
''          + CisFun.Val2(iB1_HaSu(Index))
'    If CisFun.Val2(iB1_MaiSu(Index)) = 0 Then
'        If CisFun.Val2(iB1_HaSu(Index)) <> 0 Then
'            iB1_MaiSu(Index) = 1
'        End If
'    End If
'    wJyuSu = CisFun.Val2(iB1_MaiSu(Index)) * CisFun.Val2(B1lb_Syuyo(Index)) _
'           + CisFun.Val2(iB1_HaSu(Index))
''   << 端数が入力時は、枚数を－１する！ >>
'    If CisFun.Val2(iB1_HaSu(Index)) <> 0 Then
'        wJyuSu = wJyuSu - CisFun.Val2(B1lb_Syuyo(Index))
'    End If
''*---------------------------------------------<< 2004/04/19 Update End >>
'    B1lb_JyuSu(Index) = Format(wJyuSu, "###,###,###")
'End Sub
'
'******************************************************************
'*      枚数・端数・受注数の再計算＆表示
'******************************************************************
Private Sub Suryo_Set(Index As Integer, Proc As String)
'   >> 収容数の判定
    If CisFun.Val2(B1lb_Syuyo(Index)) <> 0 Then
'       >> 収容数が入力済
        Select Case Proc
            Case "M"
            '   枚数＆端数を元に、受注数を算出
                If CisFun.Val2(iB1_HaSu(Index)) <> 0 Then
                    If CisFun.Val2(iB1_MaiSu(Index)) = 0 Then
                        iB1_MaiSu(Index) = 1
                    End If
                    iB1_JyuSu(Index) = (CisFun.Val2(iB1_MaiSu(Index)) - 1) * CisFun.Val2(B1lb_Syuyo(Index)) _
                                     + CisFun.Val2(iB1_HaSu(Index))
                Else
                    iB1_JyuSu(Index) = CisFun.Val2(iB1_MaiSu(Index)) * CisFun.Val2(B1lb_Syuyo(Index))
                End If
            Case "S"
            '   受注数を元に、枚数＆端数を算出
                iB1_MaiSu(Index) = CisFun.Val2(iB1_JyuSu(Index)) \ CisFun.Val2(B1lb_Syuyo(Index))
                iB1_HaSu(Index) = CisFun.Val2(iB1_JyuSu(Index)) Mod CisFun.Val2(B1lb_Syuyo(Index))
                If CisFun.Val2(iB1_HaSu(Index)) <> 0 Then
                    iB1_MaiSu(Index) = CisFun.Val2(iB1_MaiSu(Index)) + 1
                End If
        End Select
    Else
'       >> 収容数が未入力
        iB1_MaiSu(Index) = ""
        iB1_HaSu(Index) = ""
    End If
End Sub
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
'******************************************************************
'*      初期処理
'******************************************************************
Private Function InitProc() As Boolean
    InitProc = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select *"
    gSL_Select = gSL_Select & "  from 受注入力ワーク"
    gSL_Select = gSL_Select & " where 端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and 累積区分 = 0"
    If JNWRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 6
            .MB_MSG(1) = "    当端末処理分の受注未累積データが存在します。            "
            .MB_MSG(4) = "      『受注画面入力（号口）』処理を終了しますか？"
            .MB_MSG(5) = "      いいえの場合は、未累積データを削除し続行します"
            .MB_Title = "未累積データの確認"
            .MB_Button = Yes_No
            If .MBOX Then Exit Function
        End With
    End If

    With CisDB
        On Error GoTo InitProc_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
'       << 受注入力テーブル >>
        .SQL = ""
        .SQL = .SQL & "delete"
        .SQL = .SQL & "  from 受注入力ワーク"
        .SQL = .SQL & " where 端末番号 = substring(host_name(),1,20)"
        .DBExec
        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        On Error GoTo 0
    End With

    InitProc = True
    Exit Function
InitProc_Err:
    CisFun.ErrorBox
    End
End Function
'==========================================================='
'   終了処理（入力データがあれば後方処理を行う）
'==========================================================='
Private Sub EndProc()

'   << 入力件数確認 >>
    gSL_Select = ""
    gSL_Select = gSL_Select & "select count(*) 件数"
    gSL_Select = gSL_Select & "  from 受注入力ワーク"
    gSL_Select = gSL_Select & " where 端末番号 = substring(host_name(),1,20)"
    gSL_Select = gSL_Select & "   and"
    gSL_Select = gSL_Select & " (     出荷指示書NO = 0"
    gSL_Select = gSL_Select & "    or 納入番号 = ''"
    gSL_Select = gSL_Select & "    or 累積区分 = 0"
    gSL_Select = gSL_Select & " )"
    If JNWRead(gSL_Select, 1) Then
        If JNW.件数 > 0 Then
        '   出荷指示NO採番
'*--------------------------------------------------<< 2004/06/24 Insert Start >>
            gCHT_DelKbn53 = 0
'*--------------------------------------------------<< 2004/06/24 Insert End >>
            CXJ0110.Show vbModal
            Unload CXJ0110
            Set CXJ0110 = Nothing
            DoEvents
        '   出荷指示書発行
            Call CXJ0120LPrint(0)
        '   自社伝票発行
            Call CXJ0140LPrint(0)
        '   自社伝票ﾌｧｲﾙ出力
            Call CXJ0150TFile(0)
        '   ノーかんばんエフ発行
            Call CXJ0130LPrint(0)
        End If
    End If

End Sub
'+---------------------------------------+
'+      名称マスタ初期設定
'+---------------------------------------+
Private Sub MeisyoInit()
'   <><><><><> 受注区分 <><><><><>
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "受注区分"
        .区分桁数 = 1
        .値 = "1"
        .値名称 = "通常"
        .作成者 = gTanto
        .制御区分 = "X"
    End With
    If Not SYMInitCreate(True) Then
        With SYMCreate
            .値 = "2":  .値名称 = "特配":       Call SYMInitCreate
            .値 = "3":  .値名称 = "臨時":       Call SYMInitCreate
            .値 = "4":  .値名称 = "スポット":   Call SYMInitCreate
            .値 = "5":  .値名称 = "指示":       Call SYMInitCreate
            .値 = "6":  .値名称 = "号試":       Call SYMInitCreate
            .値 = "7":  .値名称 = "代替":       Call SYMInitCreate
            .値 = "8":  .値名称 = "その他":     Call SYMInitCreate
        End With
    End If
'   >> 受注区分の表示
    gSL_Select = ""
    gSL_Select = gSL_Select & "select 値"
    gSL_Select = gSL_Select & "     , 値名称"
    gSL_Select = gSL_Select & "  from 名称マスタ"
    gSL_Select = gSL_Select & " where 区分名称 = '受注区分'"
    gSL_Select = gSL_Select & "   and 区分タイプ = ''"
    gSL_Select = gSL_Select & " order by 値"
    If SYMRead(gSL_Select, 0) Then
        Do Until Not SYM_RDSTS
            Select Case SYM.値
                Case "1":   H1Op_JKbn(1).Caption = SYM.値名称:  H1Op_JKbn(1).Tag = SYM.値
                Case "5":   H1Op_JKbn(2).Caption = SYM.値名称:  H1Op_JKbn(2).Tag = SYM.値
                Case "6":   H1Op_JKbn(3).Caption = SYM.値名称:  H1Op_JKbn(3).Tag = SYM.値
                Case "4":   H1Op_JKbn(4).Caption = SYM.値名称:  H1Op_JKbn(4).Tag = SYM.値
                Case "3":   H1Op_JKbn(5).Caption = SYM.値名称:  H1Op_JKbn(5).Tag = SYM.値
                Case "2":   H1Op_JKbn(6).Caption = SYM.値名称:  H1Op_JKbn(6).Tag = SYM.値
                Case "7":   H1Op_JKbn(7).Caption = SYM.値名称:  H1Op_JKbn(7).Tag = SYM.値
                Case "8":   H1Op_JKbn(8).Caption = SYM.値名称:  H1Op_JKbn(8).Tag = SYM.値
            End Select
            Call SYMReadNext
        Loop
    End If
    Call SYMClose
End Sub
'+---------------------------------------+
'+      取引先マスタ索引（名称索引）
'+---------------------------------------+
Public Function SqlSelectTRM(Torcd As String, Optional DBNo As Integer = 0) As Boolean
    SqlSelectTRM = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select 略称"
    gSL_Select = gSL_Select & "     , 工場名"
    gSL_Select = gSL_Select & "     , 伝票種類"
    gSL_Select = gSL_Select & "  from 取引先マスタ "
    gSL_Select = gSL_Select & " where 取引先CD = '" & RTrim(Torcd) & "'"
    gSL_Select = gSL_Select & "   and 取引先区分 = 0 "

    If Not TRMRead(gSL_Select, 1, DBNo) Then Exit Function

    SqlSelectTRM = True
End Function
'+---------------------------------------+
'+      品番マスタ索引
'+---------------------------------------+
Public Function SqlSelectHNM(Hinbn As String, Optional DBNo As Integer = 0) As Boolean
    SqlSelectHNM = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select 品番"
    gSL_Select = gSL_Select & "  from 品番マスタ "
    gSL_Select = gSL_Select & " where 品番 = '" & RTrim(Hinbn) & "'"

    If Not HNMRead(gSL_Select, 1, DBNo) Then Exit Function

    SqlSelectHNM = True
End Function
'+---------------------------------------+
'+      品番取引先マスタ索引
'+---------------------------------------+
Public Function SqlSelectHTM(Hinbn As String, Seban As String, Optional DBNo As Integer = 0) As Boolean
    SqlSelectHTM = False

    gSL_Select = ""
    gSL_Select = gSL_Select & "select 品番"
    gSL_Select = gSL_Select & "     , 背番号"
    gSL_Select = gSL_Select & "     , 収容数1"
    gSL_Select = gSL_Select & "     , 収容数2"
    gSL_Select = gSL_Select & "     , 変更日"
    gSL_Select = gSL_Select & "     , 変更便"
    gSL_Select = gSL_Select & "     , 伝票種類"
    gSL_Select = gSL_Select & "  from 品番取引先マスタ"
'   >> 条件ｾｯﾄ開始
    If RTrim(Hinbn) <> "" Then
        gSL_Select = gSL_Select & " where 品番 = '" & RTrim(Hinbn) & "'"
    Else
        gSL_Select = gSL_Select & " where 背番号 = '" & RTrim(Seban) & "'"
    End If
'   >> 条件ｾｯﾄ終了
    gSL_Select = gSL_Select & "   and 取引先 = '" & RTrim(iH2_Torcd) & "'"
    gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(iH2_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and 品目 = 0"

    If HTMRead(gSL_Select, , DBNo) Then
        If CisDB.RecordCount = 1 Then SqlSelectHTM = True
    End If

    Call HTMClose(DBNo)
End Function
'*---------------------------------------------<< 2004/04/21 Insert Start >>
'+---------------------------------------+
'+      受注テーブル索引（納品書頁＆行）
'+---------------------------------------+
Public Function SqlSelectJYT(Optional DBNo As Integer = 0) As Boolean
    SqlSelectJYT = False

    w_PageNo = 1
    w_PageGyo = 1

    gSL_Select = ""
    gSL_Select = gSL_Select & "select jyt.納品書頁"
    gSL_Select = gSL_Select & "     , jyt.納品書行"
    gSL_Select = gSL_Select & "  from 受注テーブル jyt"
    gSL_Select = gSL_Select & " where jyt.納入番号 = '" & RTrim(iH2_Nouban) & "'"
    gSL_Select = gSL_Select & "   and jyt.受注区分 = " & w_JKbn
    gSL_Select = gSL_Select & "   and jyt.納入日 = '" & RTrim(iH2_Ymd) & "'"
    gSL_Select = gSL_Select & "   and jyt.納入便 = '" & RTrim(iH2_Bin) & "'"
    gSL_Select = gSL_Select & "   and jyt.納入先 = '" & RTrim(iH2_Torcd) & "'"
    gSL_Select = gSL_Select & "   and jyt.受入 = '" & RTrim(iH2_Ukeir) & "'"
    gSL_Select = gSL_Select & " union all "
    gSL_Select = gSL_Select & "select jnw.納品書頁"
    gSL_Select = gSL_Select & "     , jnw.納品書行"
    gSL_Select = gSL_Select & "  from 受注入力ワーク jnw"
    gSL_Select = gSL_Select & " where jnw.納入番号 = '" & RTrim(iH2_Nouban) & "'"
    gSL_Select = gSL_Select & "   and jnw.受注区分 = " & w_JKbn
    gSL_Select = gSL_Select & "   and jnw.納入日 = '" & RTrim(iH2_Ymd) & "'"
    gSL_Select = gSL_Select & "   and jnw.納入便 = '" & RTrim(iH2_Bin) & "'"
    gSL_Select = gSL_Select & "   and jnw.納入先 = '" & RTrim(iH2_Torcd) & "'"
    gSL_Select = gSL_Select & "   and jnw.受入 = '" & RTrim(iH2_Ukeir) & "'"
    gSL_Select = gSL_Select & " order by 1 desc"
    gSL_Select = gSL_Select & "        , 2 desc"
    If JYTRead(gSL_Select, 1, DBNo) Then
        w_PageNo = JYT.納品書頁
        w_PageGyo = JYT.納品書行 + 1
    Else
        gSL_Select = ""
        gSL_Select = gSL_Select & "select jyt.納品書頁"
        gSL_Select = gSL_Select & "  from 受注テーブル jyt"
        gSL_Select = gSL_Select & " where jyt.納入番号 = '" & RTrim(iH2_Nouban) & "'"
        gSL_Select = gSL_Select & " union all "
        gSL_Select = gSL_Select & "select jnw.納品書頁"
        gSL_Select = gSL_Select & "  from 受注入力ワーク jnw"
        gSL_Select = gSL_Select & " where jnw.納入番号 = '" & RTrim(iH2_Nouban) & "'"
        gSL_Select = gSL_Select & " order by 1 desc"
        If JYTRead(gSL_Select, 1, DBNo) Then
            w_PageNo = JYT.納品書頁 + 1
            w_PageGyo = 1
        End If
    End If

    SqlSelectJYT = True
End Function
'*---------------------------------------------<< 2004/04/21 Insert Start >>

'******************************************************************
'*
'*      受注入力ワーク＆受注テーブル　重複チェック
'*
'******************************************************************
Private Function Daburi_Check() As Boolean
    Daburi_Check = False

    gStr1 = ""
    gStr2 = ""
    For wIndex = 1 To 10
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'       If RTrim(iB1_Hinbn(wIndex)) <> "" Or _
'          RTrim(iB1_Seban(wIndex)) <> "" Or _
'          CisFun.Val2(iB1_MaiSu(wIndex)) <> 0 Or _
'          CisFun.Val2(iB1_HaSu(wIndex)) <> 0 Then
'
        If RTrim(iB1_Hinbn(wIndex)) <> "" Or _
           RTrim(iB1_Seban(wIndex)) <> "" Or _
           CisFun.Val2(iB1_MaiSu(wIndex)) <> 0 Or _
           CisFun.Val2(iB1_HaSu(wIndex)) <> 0 Or _
           CisFun.Val2(iB1_JyuSu(wIndex)) <> 0 Then
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
        '   >> 比較
            If w_DenKnd(wIndex) = "" Then
                gStr1 = gStr1 & gStr2 & "'" & RTrim(iB1_Hinbn(wIndex)) & "'"
                gStr2 = ", "
            End If
        End If
    Next wIndex
    If gStr1 <> "" Then
    '   >> 自社伝票以外の時！
        gSL_Select = ""
        gSL_Select = gSL_Select & "select"
        gSL_Select = gSL_Select & "  jnw.受注区分"
        gSL_Select = gSL_Select & ", jnw.受注日"
        gSL_Select = gSL_Select & ", jnw.品番"
        gSL_Select = gSL_Select & ", jnw.納入数"
        gSL_Select = gSL_Select & ", isnull(hnm.表示品番, jnw.品番) 表示品番"
        gSL_Select = gSL_Select & ", isnull(sy1.値名称, '') 受注区分名"
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  受注入力ワーク jnw"
        gSL_Select = gSL_Select & "  left outer join 品番マスタ hnm"
        gSL_Select = gSL_Select & "    on hnm.品番 = jnw.品番"
        gSL_Select = gSL_Select & "  left outer join 名称マスタ sy1"
        gSL_Select = gSL_Select & "    on sy1.区分名称 = '受注区分'"
        gSL_Select = gSL_Select & "   and sy1.区分タイプ = ''"
        gSL_Select = gSL_Select & "   and sy1.値 = jnw.受注区分"
        gSL_Select = gSL_Select & " where jnw.納入先 = '" & RTrim(iH2_Torcd) & "'"
        gSL_Select = gSL_Select & "   and jnw.受入 = '" & RTrim(iH2_Ukeir) & "'"
        gSL_Select = gSL_Select & "   and jnw.納入日 = '" & RTrim(iH2_Ymd) & "'"
        gSL_Select = gSL_Select & "   and jnw.納入便 = '" & RTrim(iH2_Bin) & "'"
        gSL_Select = gSL_Select & "   and jnw.納入番号 = '" & RTrim(iH2_Nouban) & "'"
        gSL_Select = gSL_Select & "   and jnw.品番 in (" & gStr1 & ")"
        gSL_Select = gSL_Select & " union all "
        gSL_Select = gSL_Select & "select"
        gSL_Select = gSL_Select & "  jyt.受注区分"
        gSL_Select = gSL_Select & ", jyt.受注日"
        gSL_Select = gSL_Select & ", jyt.品番"
        gSL_Select = gSL_Select & ", jyt.納入数"
        gSL_Select = gSL_Select & ", isnull(hnm.表示品番, jyt.品番) 表示品番"
        gSL_Select = gSL_Select & ", isnull(sy2.値名称, '') 受注区分名"
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  受注テーブル jyt"
        gSL_Select = gSL_Select & "  left outer join 品番マスタ hnm"
        gSL_Select = gSL_Select & "    on hnm.品番 = jyt.品番"
        gSL_Select = gSL_Select & "  left outer join 名称マスタ sy2"
        gSL_Select = gSL_Select & "    on sy2.区分名称 = '受注区分'"
        gSL_Select = gSL_Select & "   and sy2.区分タイプ = ''"
        gSL_Select = gSL_Select & "   and sy2.値 = jyt.受注区分"
        gSL_Select = gSL_Select & " where jyt.納入先 = '" & RTrim(iH2_Torcd) & "'"
        gSL_Select = gSL_Select & "   and jyt.受入 = '" & RTrim(iH2_Ukeir) & "'"
        gSL_Select = gSL_Select & "   and jyt.納入日 = '" & RTrim(iH2_Ymd) & "'"
        gSL_Select = gSL_Select & "   and jyt.納入便 = '" & RTrim(iH2_Bin) & "'"
        gSL_Select = gSL_Select & "   and jyt.納入番号 = '" & RTrim(iH2_Nouban) & "'"
        gSL_Select = gSL_Select & "   and jyt.品番 in (" & gStr1 & ")"
        If JYTRead(gSL_Select, 0, 1) Then
            Exit Function
        End If
        Call JYTClose(1)
    End If

    Daburi_Check = True
End Function

'**************************************************************
'*                                                            *
'*       明細入力時の行単位 削除＆複写＆移動＆挿入 処理       *
'*                                                            *
'**************************************************************
'+---------------------------------------+
'+      明細入力時の行削除 ( Alt + D )
'+---------------------------------------+
Private Function Input_Alt_D() As Boolean
    Input_Alt_D = False

    wIndex = ActiveControl.Index
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'   If RTrim(iB1_Hinbn(wIndex)) = "" And _
'      RTrim(iB1_Seban(wIndex)) = "" And _
'      CisFun.Val2(iB1_MaiSu(wIndex)) = 0 And _
'      CisFun.Val2(iB1_HaSu(wIndex)) = 0 Then Exit Function
'
    If RTrim(iB1_Hinbn(wIndex)) = "" And _
       RTrim(iB1_Seban(wIndex)) = "" And _
       CisFun.Val2(iB1_MaiSu(wIndex)) = 0 And _
       CisFun.Val2(iB1_HaSu(wIndex)) = 0 And _
       CisFun.Val2(iB1_JyuSu(wIndex)) = 0 Then Exit Function
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End

    Body_GNo(wIndex).BackColor = wSetColor
    wProcF = 3
    Call SetMode(wProcF, wIndex)

    Input_Alt_D = True
End Function
'+---------------------------------------+
'+      明細入力時の行追加 ( Alt + E )
'+---------------------------------------+
Private Function Input_Alt_E() As Boolean
    Input_Alt_E = False

    wIndex = ActiveControl.Index
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'   If RTrim(iB1_Hinbn(wIndex)) = "" And _
'      RTrim(iB1_Seban(wIndex)) = "" And _
'      CisFun.Val2(iB1_MaiSu(wIndex)) = 0 And _
'      CisFun.Val2(iB1_HaSu(wIndex)) = 0 Then Exit Function
'
    If RTrim(iB1_Hinbn(wIndex)) = "" And _
       RTrim(iB1_Seban(wIndex)) = "" And _
       CisFun.Val2(iB1_MaiSu(wIndex)) = 0 And _
       CisFun.Val2(iB1_HaSu(wIndex)) = 0 And _
       CisFun.Val2(iB1_JyuSu(wIndex)) = 0 Then Exit Function
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End

    wProcF = 1
    Call SetMode(wProcF, wIndex)

    Input_Alt_E = True
End Function
'+---------------------------------------+
'+      明細入力時の行複写 ( Alt + C )
'+---------------------------------------+
Private Function Input_Alt_C() As Boolean
    Input_Alt_C = False

    wIndex = ActiveControl.Index
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'   If RTrim(iB1_Hinbn(wIndex)) = "" And _
'      RTrim(iB1_Seban(wIndex)) = "" And _
'      CisFun.Val2(iB1_MaiSu(wIndex)) = 0 And _
'      CisFun.Val2(iB1_HaSu(wIndex)) = 0 Then Exit Function
'
    If RTrim(iB1_Hinbn(wIndex)) = "" And _
       RTrim(iB1_Seban(wIndex)) = "" And _
       CisFun.Val2(iB1_MaiSu(wIndex)) = 0 And _
       CisFun.Val2(iB1_HaSu(wIndex)) = 0 And _
       CisFun.Val2(iB1_JyuSu(wIndex)) = 0 Then Exit Function
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End

    Body_GNo(wIndex).BackColor = wSetColor
    wProcF = 0
    Call SetMode(wProcF, wIndex)

    Input_Alt_C = True
End Function
'+---------------------------------------+
'+      明細入力時の行移動 ( Alt + F )
'+---------------------------------------+
Private Function Input_Alt_F() As Boolean
    Input_Alt_F = False

    wIndex = ActiveControl.Index
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'   If RTrim(iB1_Hinbn(wIndex)) = "" And _
'      RTrim(iB1_Seban(wIndex)) = "" And _
'      CisFun.Val2(iB1_MaiSu(wIndex)) = 0 And _
'      CisFun.Val2(iB1_HaSu(wIndex)) = 0 Then Exit Function
'
    If RTrim(iB1_Hinbn(wIndex)) = "" And _
       RTrim(iB1_Seban(wIndex)) = "" And _
       CisFun.Val2(iB1_MaiSu(wIndex)) = 0 And _
       CisFun.Val2(iB1_HaSu(wIndex)) = 0 And _
       CisFun.Val2(iB1_JyuSu(wIndex)) = 0 Then Exit Function
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End

    Body_GNo(wIndex).BackColor = wSetColor
    wProcF = 2
    Call SetMode(wProcF, wIndex)

    Input_Alt_F = True
End Function
'+----------------------------------------------+
'+    明細入力時の行削除・行複写処理取消 ( F9 ) +
'+----------------------------------------------+
Private Sub PB_PCAN_Click()
    Call SetMode(9, Sav_Index)
End Sub
'+---------------------------------+
'+    行削除・行複写の設定＆解除     +
'+---------------------------------+
Private Sub SetMode(ProcF As Integer, Index As Integer)
    Select Case ProcF
        Case 0
            C1_Area1.Visible = True
            B1lb_ProcMsg = "複写元行№を複写先行№へ《DragDrop》 して下さい"
            Body_GNo(Index).DragMode = vbAutomatic
            Sav_Index = Index
            Sto_Index = -1
            PB_PCAN.Visible = True
'           Timer1.Enabled = False
        Case 1
            B1lb_ProcMsg = ""
            Call Gyo_Insert(Index)
            Body_GNo(Index).BackColor = wLosColor
        Case 2
            C1_Area1.Visible = True
            B1lb_ProcMsg = "移動元行№を移動先行№へ《DragDrop》 して下さい"
            Body_GNo(Index).DragMode = vbAutomatic
            Sav_Index = Index
            Sto_Index = -1
            PB_PCAN.Visible = True
'           Timer1.Enabled = False
        Case 3
            B1lb_ProcMsg = ""
            Call Gyo_Delete(Index)
            Body_GNo(Index).BackColor = wLosColor
        Case Else
            B1lb_ProcMsg = ""
            Body_GNo(Index).DragMode = vbManual
            If Sav_Index > 0 Then Body_GNo(Sav_Index).ForeColor = &H0&
            If Sto_Index > 0 Then Body_GNo(Sto_Index).ForeColor = &H0&
            PB_PCAN.Visible = False
            C1_Area1.Visible = False
            Body_GNo(Index).BackColor = wLosColor
'           Timer1.Enabled = True
    End Select
End Sub
'+---------------------------------------------------------------+
'+   Body_GNo( MouseDown ) 処理選択画面表示
'+---------------------------------------------------------------+
Private Sub Body_GNo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
'
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'   If RTrim(iB1_Hinbn(Index)) = "" And _
'      RTrim(iB1_Seban(Index)) = "" And _
'      CisFun.Val2(iB1_MaiSu(Index)) = 0 And _
'      CisFun.Val2(iB1_HaSu(Index)) = 0 Then Exit Sub
'
    If RTrim(iB1_Hinbn(Index)) = "" And _
       RTrim(iB1_Seban(Index)) = "" And _
       CisFun.Val2(iB1_MaiSu(Index)) = 0 And _
       CisFun.Val2(iB1_HaSu(Index)) = 0 And _
       CisFun.Val2(iB1_JyuSu(Index)) = 0 Then Exit Sub
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
'
    If Button = 2 Then
        gProcF_Left = Me.Left + B1_Area1.Left + 660
        gProcF_Top = Me.Top + B1_Area1.Top + 570 + ((Index) * 480)
        Body_GNo(Index).BackColor = wSetColor
'
        ProcForm.Show vbModal
        Unload ProcForm
        Set ProcForm = Nothing
'
        Sav_Index = -1
        Sto_Index = -1
        wProcF = gProcF_Rtn
        Call SetMode(wProcF, Index)
    End If
'
End Sub
'+---------------------------------------------------------+
'+   Body_GNo ( DragOver & DragDrop ) 行コピー  ( 複写元 )
'+---------------------------------------------------------+
Private Sub Body_GNo_DragOver(Index As Integer, Source As Control, x As Single, Y As Single, State As Integer)
    If Sto_Index = -1 Then
'       複写元の文字色を変更
        Body_GNo(Index).ForeColor = &H80FF&
    End If
'   >> ドラッグ行№を複写先として退避
    Sto_Index = Index
End Sub
'+---------------------------------------------------------+
'+   Body_GNo ( DragOver & DragDrop ) 行コピー  ( 複写先 )
'+---------------------------------------------------------+
Private Sub Body_GNo_DragDrop(Index As Integer, Source As Control, x As Single, Y As Single)

    If Sto_Index = -1 Then Exit Sub

    Body_GNo(Sto_Index).ForeColor = &H80FF&
    If Sav_Index = Sto_Index Then GoTo Body_GNo_DragDrop_ED2
'   挿入以外
    If wProcF <> 1 Then
        If RTrim(iB1_Hinbn(Index)) <> "" Or _
           RTrim(iB1_Seban(Index)) <> "" Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(0) = "【 データが存在します 】"
                .MB_MSG(2) = "   旧品番 = " & iB1_Hinbn(Sto_Index) & Space(25 - CisFun.Len2(iB1_Hinbn(Sto_Index)))
                .MB_MSG(3) = "   新品番 = " & iB1_Hinbn(Sav_Index) & Space(25 - CisFun.Len2(iB1_Hinbn(Sav_Index)))
                .MB_MSG(5) = "     上書きしますか？"
                Select Case wProcF
                    Case 0: .MB_Title = "《 行複写 》"
                    Case 2: .MB_Title = "《 行移動 》"
                End Select
                .MB_Button = OK_CAN
                If Not .MBOX Then GoTo Body_GNo_DragDrop_ED2
            End With
        End If
    End If
'
    iB1_Hinbn(Sto_Index) = iB1_Hinbn(Sav_Index)
    iB1_Seban(Sto_Index) = iB1_Seban(Sav_Index)
    iB1_MaiSu(Sto_Index) = iB1_MaiSu(Sav_Index)
    iB1_HaSu(Sto_Index) = iB1_HaSu(Sav_Index)
    B1lb_Syuyo(Sto_Index) = B1lb_Syuyo(Sav_Index)
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'   B1lb_JyuSu(Sto_Index) = B1lb_JyuSu(Sav_Index)
    iB1_JyuSu(Sto_Index) = iB1_JyuSu(Sav_Index)
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
    If wProcF = 2 Then
        iB1_Hinbn(Sav_Index) = ""
        iB1_Seban(Sav_Index) = ""
        iB1_MaiSu(Sav_Index) = ""
        iB1_HaSu(Sav_Index) = ""
        B1lb_Syuyo(Sav_Index) = ""
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'       B1lb_JyuSu(Sav_Index) = ""
        iB1_JyuSu(Sav_Index) = ""
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
    End If
'
Body_GNo_DragDrop_ED2:
    Call PB_PCAN_Click
End Sub
'+------------------+
'+    行№挿入      +
'+------------------+
Private Function Gyo_Insert(Index As Integer) As Boolean
    Gyo_Insert = False

    If RTrim(iB1_Hinbn(10)) <> "" Or _
       RTrim(iB1_Seban(10)) <> "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(0) = "【 最終データが存在します 】"
            .MB_MSG(2) = "   品番 = " & iB1_Hinbn(10) & Space(25 - CisFun.Len2(iB1_Hinbn(10)))
            .MB_MSG(4) = "挿入すると､この行は切れますがよろしいですか？"
            .MB_Title = "《 行挿入 》"
            .MB_Button = OK_CAN
            If Not .MBOX Then Exit Function
        End With
    End If

    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
'   >> 内容の移動
    For gIndex = 9 To Index Step -1
        iB1_Hinbn(gIndex + 1) = iB1_Hinbn(gIndex)
        iB1_Seban(gIndex + 1) = iB1_Seban(gIndex)
        iB1_MaiSu(gIndex + 1) = iB1_MaiSu(gIndex)
        iB1_HaSu(gIndex + 1) = iB1_HaSu(gIndex)
        B1lb_Syuyo(gIndex + 1) = B1lb_Syuyo(gIndex)
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'       B1lb_JyuSu(gIndex + 1) = B1lb_JyuSu(gIndex)
        iB1_JyuSu(gIndex + 1) = iB1_JyuSu(gIndex)
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
    Next gIndex
'   >> 内容ｸﾘｱ
    iB1_Hinbn(Index) = ""
    iB1_Seban(Index) = ""
    iB1_MaiSu(Index) = ""
    iB1_HaSu(Index) = ""
    B1lb_Syuyo(Index) = ""
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'   B1lb_JyuSu(Index) = ""
    iB1_JyuSu(Index) = ""
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
'   >> ﾌｫｰｶｽ･ｾｯﾄ
    Dummy.Enabled = False
    iB1_Hinbn(Index).SetFocus

    Gyo_Insert = True
End Function
'+------------------+
'+    行削除        +
'+------------------+
Private Function Gyo_Delete(Index As Integer) As Boolean
    Gyo_Delete = False

    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "  【 行№ = " & Body_GNo(Index) & " を削除します 】            "
        .MB_MSG(4) = "           削除しますか？"
        .MB_Title = "《 行削除 》"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With

'   >> 内容ｸﾘｱ
    iB1_Hinbn(Index) = ""
    iB1_Seban(Index) = ""
    iB1_MaiSu(Index) = ""
    iB1_HaSu(Index) = ""
    B1lb_Syuyo(Index) = ""
'*↓↓*-------------------------------------------<< 2004/09/01 >> Update Start
'   B1lb_JyuSu(Index) = ""
    iB1_JyuSu(Index) = ""
'*↑↑*-------------------------------------------<< 2004/09/01 >> Update End
'   >> ﾌｫｰｶｽ･ｾｯﾄ
    iB1_Hinbn(Index).SetFocus

    Gyo_Delete = True
End Function


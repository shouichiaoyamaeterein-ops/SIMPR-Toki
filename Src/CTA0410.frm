VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "Cis3D_6.0.ocx"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CisBtn_6.0.ocx"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CisText_6.0.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CisYMDwB_6.0.ocx"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CisYM_6.0.ocx"
Begin VB.Form CTA0410 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "棚卸状況照会"
   ClientHeight    =   11010
   ClientLeft      =   6000
   ClientTop       =   4425
   ClientWidth     =   15390
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   700
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
   ScaleHeight     =   11010
   ScaleWidth      =   15390
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D PNL_Wait 
      Height          =   930
      Left            =   5205
      Top             =   5985
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
   Begin VB.PictureBox B1_Area1 
      Height          =   7920
      Left            =   135
      ScaleHeight     =   7860
      ScaleWidth      =   15060
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2685
      Width           =   15120
      Begin TrueOleDBGrid60.TDBGrid DBGrid1 
         Bindings        =   "CTA0410.frx":0000
         Height          =   7740
         Left            =   45
         OleObjectBlob   =   "CTA0410.frx":0015
         TabIndex        =   1
         Top             =   45
         Width           =   14955
      End
   End
   Begin VB.PictureBox H1_Area1 
      Height          =   1290
      Left            =   270
      ScaleHeight     =   1230
      ScaleWidth      =   12765
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1260
      Width           =   12825
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   570
         Index           =   1
         Left            =   9180
         Top             =   45
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   1005
         ForeColor       =   16711680
         Caption         =   "加工度"
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
         cPositionX      =   75
         Begin VB.ComboBox H1Comb_KakoDo 
            BackColor       =   &H00C0E0FF&
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
            ItemData        =   "CTA0410.frx":16389
            Left            =   960
            List            =   "CTA0410.frx":1638B
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   6
            Top             =   120
            Width           =   1695
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   585
         Left            =   4740
         Top             =   615
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   1032
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
         cAlingnment     =   1
         cPositionX      =   75
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   870
            TabIndex        =   5
            Top             =   120
            Width           =   3210
            _ExtentX        =   5662
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
            Text            =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
            MaxLength       =   25
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   585
         Left            =   9180
         Top             =   615
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   1032
         ForeColor       =   16711680
         Caption         =   "担当者"
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
         cPositionX      =   75
         Begin CisText_V60.CisText iH1_Sagyo 
            Height          =   375
            Left            =   960
            TabIndex        =   7
            Top             =   120
            Width           =   2500
            _ExtentX        =   4419
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
            Text            =   "XXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   20
            IMEMode         =   4
            cFaZero         =   0
         End
         Begin Cis3D_v60.CIS3D H1lb_Sagyo 
            Height          =   360
            Left            =   1470
            Top             =   120
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXX"
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
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   570
         Left            =   4740
         Top             =   45
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   1005
         ForeColor       =   16711680
         Caption         =   "契約先"
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
         cPositionX      =   75
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   870
            TabIndex        =   4
            Top             =   120
            Width           =   990
            _ExtentX        =   1746
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
            Text            =   "XXXXXXX"
            MaxLength       =   7
         End
         Begin Cis3D_v60.CIS3D H1lb_Torcd 
            Height          =   360
            Left            =   1860
            Top             =   120
            Width           =   2460
            _ExtentX        =   4339
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
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   1155
         Index           =   0
         Left            =   2160
         Top             =   45
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   2037
         ForeColor       =   16711680
         Caption         =   "作業日"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         cPositionY      =   45
         Begin CisYMD_With_Btn.CisYMDwB iH1_SDate 
            Height          =   345
            Left            =   570
            TabIndex        =   2
            Top             =   360
            Width           =   1815
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
            Object.Width           =   1815
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
         Begin Cis3D_v60.CIS3D PB_SDate 
            Height          =   345
            Left            =   90
            Top             =   360
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   609
            Caption         =   "開始"
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_EDate 
            Height          =   345
            Left            =   90
            Top             =   720
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   609
            Caption         =   "終了"
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin CisYMD_With_Btn.CisYMDwB iH1_EDate 
            Height          =   345
            Left            =   570
            TabIndex        =   3
            Top             =   720
            Width           =   1815
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
            Object.Width           =   1815
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
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   1155
         Index           =   1
         Left            =   60
         Top             =   45
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   2037
         BackColor       =   16761087
         ForeColor       =   16711680
         Caption         =   "棚卸年月"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         cPositionY      =   45
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   330
            Left            =   400
            TabIndex        =   0
            Top             =   510
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   570
         Index           =   0
         Left            =   9180
         Top             =   45
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   1005
         ForeColor       =   16711680
         Caption         =   "納入先"
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
         cPositionX      =   75
         Begin CisText_V60.CisText iH1_Kikai 
            Height          =   375
            Left            =   960
            TabIndex        =   23
            Top             =   120
            Width           =   510
            _ExtentX        =   900
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
            Text            =   "XXX"
            MaxLength       =   3
         End
         Begin Cis3D_v60.CIS3D H1lb_Kikai 
            Height          =   360
            Left            =   1470
            Top             =   120
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXX"
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
   Begin Cis3D_v60.CIS3D UC_3D31 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "　　　【 棚卸状況問合せ 】"
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
      Begin VB.PictureBox Back_Over 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   4665
         ScaleHeight     =   255
         ScaleWidth      =   1935
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   45
         Visible         =   0   'False
         Width           =   1995
         Begin VB.OptionButton Op_Overlap 
            BackColor       =   &H00C0C0C0&
            Caption         =   "単一"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   1020
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Value           =   -1  'True
            Width           =   780
         End
         Begin VB.OptionButton Op_Overlap 
            BackColor       =   &H00C0C0C0&
            Caption         =   "重複"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   780
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D32 
         Height          =   255
         Left            =   7500
         Top             =   45
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   450
         BackColor       =   14737632
         Caption         =   "抽出件数[              ]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         cBoderStyle     =   2
         cPositionX      =   50
         Begin Cis3D_v60.CIS3D lb_Kensu 
            Height          =   285
            Left            =   1095
            Top             =   15
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "##,### / ##,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   14737632
            cBoderColor2    =   14737632
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
         End
      End
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14415
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
         Left            =   11925
         Top             =   30
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
   Begin VB.PictureBox Back_Dsp 
      BackColor       =   &H00404040&
      Height          =   720
      Left            =   270
      ScaleHeight     =   660
      ScaleWidth      =   5295
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   450
      Width           =   5355
      Begin Cis3D_v60.CIS3D UC_3D5 
         Height          =   645
         Left            =   15
         Top             =   0
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   1138
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         Begin Cis3D_v60.CIS3D UC_3D9 
            Height          =   570
            Index           =   0
            Left            =   60
            Top             =   30
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   1005
            BackColor       =   16761024
            Caption         =   "表示切替"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D UC_3D10 
            Height          =   570
            Left            =   735
            Top             =   30
            Width           =   3510
            _ExtentX        =   6191
            _ExtentY        =   1005
            Caption         =   ""
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderStyle     =   1
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "⑤"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   4
               Left            =   2730
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   150
               Width           =   615
            End
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "④"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   3
               Left            =   2070
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   150
               Width           =   615
            End
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "③"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   2
               Left            =   1395
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   150
               Width           =   615
            End
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "②"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   1
               Left            =   720
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   150
               Width           =   615
            End
            Begin VB.OptionButton B1Op_Patn 
               BackColor       =   &H008080FF&
               Caption         =   "①"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   14.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   0
               Left            =   60
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   150
               Value           =   -1  'True
               Width           =   615
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D7 
            Height          =   570
            Left            =   4260
            Top             =   30
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   1005
            ForeColor       =   16711935
            Caption         =   ""
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
            Begin VB.OptionButton Op_Sort 
               BackColor       =   &H00C0FFC0&
               Caption         =   "昇順"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Index           =   0
               Left            =   30
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   60
               Value           =   -1  'True
               Width           =   825
            End
            Begin VB.OptionButton Op_Sort 
               BackColor       =   &H00C0C0C0&
               Caption         =   "降順"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Index           =   1
               Left            =   30
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   315
               Width           =   825
            End
         End
      End
   End
   Begin VB.TextBox Dummy 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   345
      IMEMode         =   3  'ｵﾌ固定
      Left            =   240
      TabIndex        =   17
      Text            =   "D"
      Top             =   210
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   510
      Top             =   210
   End
   Begin VB.Timer Timer2 
      Left            =   1230
      Top             =   210
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   705
      Left            =   8835
      ScaleHeight     =   645
      ScaleWidth      =   5970
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   450
      Width           =   6030
      Begin CisBtn_60.CisBtn PB_Set 
         Height          =   585
         Left            =   2580
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
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
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   5100
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
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
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   4260
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
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
         Left            =   3420
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
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
      Begin CisBtn_60.CisBtn PB_Csv 
         Height          =   585
         Left            =   1740
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   14
         BtnCaption      =   "出力"
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
         PFCaption       =   "( F10 )"
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
      Begin CisBtn_60.CisBtn PB_Prt 
         Height          =   585
         Left            =   900
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   13
         BtnCaption      =   "印刷"
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
         PFCaption       =   "( F8 )"
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   45
         Top             =   30
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
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
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   405
      Left            =   0
      Top             =   765
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
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
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "CTA0410"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   棚卸状況問合せ
'**       フォームID    :   CTA0410 (Based CSJ0410 (Based CXJ0410))
'**       処理概要      :
'**
'**       作  成  日    :   2019/07/08  By CIS
'**       変  更  日    :
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean          ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2       ' 入力域切替
    Dim MaxCnt                  As Long
    Dim NowCnt                  As Long
    Dim mSL_SelectW             As String           ' SQL SELECT(Where)条件文

    Dim CisDBGrid               As New CisDBGrid    ' TDBGridPro60ｸﾗｽ定義
    
    Dim mGotColor               As Long
    Dim mLostColor              As Long
    
    Dim CisComb1                As CisCombo         ' 2019/07/12 ADD

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()

    If Rv_Call3 = "" Then
    '   #-------------------#
    '   # 二 重 起 動 防 止  #
    '   #-------------------#
        If Not Dupli_Start Then End
    '   #-------------------#
    '   # 初 期 内 容 取 得  #
    '   #-------------------#
        If Not IniGet Then End
    Else
        FormAct = False
    End If
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
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
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)

    mGotColor = Op_Sort(0).BackColor
    mLostColor = Op_Sort(1).BackColor
    
'+---------------------+
'+ TrueGrid60 初期設定  +
'+---------------------+
    With CisDBGrid
        Set .DBGridObj = DBGrid1
        .FrmCaption = Me.Caption
        .MaxSort = 7                        ' ﾚｲｱｳﾄｿｰﾄ最大数
        .LayOut = True                      ' ﾚｲｱｳﾄ変更処理 True:有り,False:無し
        .LayOutMax = 5                      ' ﾚｲｱｳﾄ情報最大数
        .LayOutColMax = 30                  ' ﾚｲｱｳﾄ表示可能項目最大数
        .LayOutPath = gIniExe               ' ﾚｲｱｳﾄ情報格納先、ﾌｧｲﾙ場所
        .LayOutName = CisFun.WSID & "TA0410.Lay"       ' ﾚｲｱｳﾄ情報格納先、ﾌｧｲﾙ名    ' 2019/07/08 UPD："SJ0410.Lay"
        .PrintHeadLine = あり
        .PrintBodyLine = あり
        .PrintBodyGyo = 10
        .CsvName = Trim(Me.Caption) & ".csv"
        .CsvRegKey = gRegKey
        .CsvValue = "CTA0410"               ' 2019/07/08 UPD："CSJ0410"
        .GridInitGet = False                 ' ｸﾞﾘｯﾄﾞ内容初期設定 ｸﾘｯﾌﾟﾎﾞｰﾄﾞ作成ﾌﾗｸﾞ
        .GridProcGet = False
        .Init
    End With
    Call GridInitSet
    
    '+---------------------+
    '+ ｺﾝﾎﾞﾎﾞｯｸｽ 初期設定  +
    '+---------------------+
    Set CisComb1 = New CisCombo
    With CisComb1
        Set .Connect = CisDB
        Set .CombObj = H1Comb_KakoDo
        .Control = "名称マスタ"
        .ComboName = "品番区分"
        .Ess = False
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
        .ComboSet
    End With

'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
'   PNL_ExeMsg.Move 4600, 400               ' 2019/07/08 DEL
    Call Str_Monitor("Start")               ' 2019/07/08 ADD
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
' Head部初期化
    B1Op_Patn(0).Value = True
    Op_Sort(0).Value = True
    Op_Sort(0).BackColor = mGotColor
    
    Op_Overlap(0).Value = True
    Op_Overlap(0).BackColor = mGotColor
    
    Call HeadBodyClear("H")
    
' ADOｺﾝﾄﾛｰﾙ　接続情報　セット
    Adodc1.ConnectionString = CisDB.ConnectString
    Adodc2.ConnectionString = CisDB.ConnectString
    '-----2014/02/04 EDIT START-----
    gStr = CisDB.ConnectString
    gConnect = "server=" & CisDB.Server & ";database=" & CisDB.DBName & ";User ID=" & CisDB.User & ";password=" & CisDB.PassWord(GetSetting(gRegKey, "Environment", "DrowsSap", ""))
    
    gStr = Replace(gStr, "Use Procedure", "Extended Properties=" & gConnect & ";Use Procedure")
    
    Adodc1.ConnectionString = gStr
    Adodc2.ConnectionString = gStr
    '-----2014/02/04 EDIT END-------
    
    iH1_YM.Text = Format(Now(), "yyyyMM")   ' 2019/07/09 ADD
    iH1_SDate.SetFocus                      ' 2019/08/09 ADD
        
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
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF8:        If PB_Prt.Visible Then Call PB_Prt_Click    '【印刷】
           Case vbKeyF10:       If PB_Csv.Visible Then Call PB_Csv_Click    '【出力】
           Case vbKeyF11:       If PB_Set.Visible Then Call PB_Set_Click    '【設定】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
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
    If Rv_Call3 = "" Then
    '+---------------------+
    '+   ﾃﾞｰﾀﾍﾞｰｽ切断
    '+---------------------+
        Call CisDB.DBDISConnect
    '+---------------------+
    '+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
    '+---------------------+
        Call CisFrm.UnLoadDisp
        End
    Else
        Unload CTA0410                      ' 2019/07/08 UPD：CSJ0410
        If Not Rv_CallObj Is Nothing Then
           Rv_CallObj.Show
        End If
    End If
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
    
    If ProcHB = "H1" Then
       Call HeadBodyClear("H")
        iH1_SDate.SetFocus
        GoTo PB_CAN_Ed
    End If
    
    If ProcHB = "B1" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call DBGridClr
        iH1_SDate.SetFocus
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B2" Then
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    CSV出力           　     +
'+----------------------------+
Private Sub PB_Csv_Click()
    Call GridInitSetCsv

    With CisDBGrid
        Set .Recordset = Adodc1.Recordset
        .PutCsv
    End With

    Call GridInitSet
End Sub
'+----------------------------+
'+    印刷             　     +
'+----------------------------+
Private Sub PB_Prt_Click()
    With CisDBGrid
        Set .Recordset = Adodc1.Recordset
        .PutPrt
    End With
End Sub
'+----------------------------+
'+    設定              　     +
'+----------------------------+
Private Sub PB_Set_Click()
    With CisDBGrid
        .LayOutPatan
        If .LayOutSetFlg Then
           B1Op_Patn(.LayOutSetNo).Value = True
           If ProcHB = "B1" Then
              Call DBInput("R")
           End If
        End If
    End With
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_SDate" Then iH1_SDate.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_EDate" Then iH1_EDate.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Kikai" Then Call Look_Kikai: Exit Sub
    If PB_Look.Tag = "iH1_Sagyo" Then Call Look_Sagyo: Exit Sub
End Sub

'+-----------------------------+
'+       品番検索              +
'+-----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
        
    RVI_Torcd = Trim(iH1_Torcd)
    RVI_Hinmoku = 1
    RVI_Tehai = 2
    RVI_SK = 9      '
    RV_Call = ""
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
       iH1_Hinbn = RV_Hinbn
    End If
End Sub

'+-----------------------------+
'+       取引先(社内)検索      +
'+-----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 3
    RV_TorKb2 = 1

    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
       iH1_Torcd = RV_Torcd
       Call iH1_Torcd_LostFocus
    End If
End Sub
'+----------------------------+
'+       機械検索             +
'+----------------------------+
Private Sub Look_Kikai()
    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "機械NO"
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
       iH1_Kikai = RV_Code          ' 値
       H1lb_Kikai = RV_Kubun        ' 値名称
    End If
End Sub
'+----------------------------+
'+       作業者検索           +
'+----------------------------+
Private Sub Look_Sagyo()
    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "担当管理"
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
       iH1_Sagyo = RV_Code          ' 値
       H1lb_Sagyo = RV_Kubun        ' 値名称
    End If
End Sub

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'+-----------------------------+
'+  iH1_SDate
'+-----------------------------+
Private Sub iH1_SDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_SDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_EDate
'+-----------------------------+
Private Sub iH1_EDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_EDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
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
    
    H1lb_Torcd = ""
    If Trim(iH1_Torcd) <> "" Then
        gSL_Select = "SELECT * FROM 取引先マスタ"
        gSL_Select = gSL_Select & " WHERE 取引先CD = '" & Trim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND 社内区分 = 1"
        If TRMRead(gSL_Select) Then
            H1lb_Torcd = TRM.正式名称
        Else
            H1lb_Torcd = ""
        End If
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
End Sub
'+-----------------------------+
'+  機械
'+-----------------------------+
Private Sub iH1_Kikai_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Kikai_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    
    H1lb_Kikai = ""
    If Trim(iH1_Kikai) <> "" Then
        gSL_Select = "SELECT * FROM 名称マスタ"
        gSL_Select = gSL_Select & " WHERE 区分名称 = '機械NO'"
        gSL_Select = gSL_Select & "   AND 値       = '" & Trim(iH1_Kikai) & "'"
        If SYMRead(gSL_Select) Then
            H1lb_Kikai = SYM.値名称
        Else
            H1lb_Kikai = ""
        End If
    End If
End Sub
'+-----------------------------+
'+  作業者
'+-----------------------------+
Private Sub iH1_Sagyo_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Sagyo_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    
    H1lb_Sagyo = ""
    If Trim(iH1_Sagyo) <> "" Then
        gSL_Select = "SELECT * FROM 名称マスタ"
        gSL_Select = gSL_Select & " WHERE 区分名称 = '担当管理'"
        gSL_Select = gSL_Select & "   AND 値       = '" & Trim(iH1_Sagyo) & "'"
        If SYMRead(gSL_Select) Then
            H1lb_Sagyo = SYM.値名称
        Else
            H1lb_Sagyo = ""
        End If
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
    
    H1Mode = False:     H1Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel: PB_ENT.BtnCaption = "入力"
        Case "B1":      B1Mode = True:  B1Color = gPGotSel: PB_ENT.BtnCaption = "詳細"
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    
    PB_Look.Visible = H1Mode
    PB_END.Visible = H1Mode
'==========================================' 2019/07/08 DEL START
'    If Not B1Mode Then
'       If PNL_ExeMsg.Visible Then
'          PNL_ExeMsg.Visible = False
'       End If
'    End If
'==========================================' 2019/07/08 DEL END
    PB_ENT.Visible = H1Mode             ' 2009.12.11 upd : PB_ENT.Visible = H1Mode Or B1Mode
    PB_Set.Visible = Not B1Mode
    PB_Prt.Visible = B1Mode
    PB_Csv.Visible = B1Mode
    
    If Rv_Call3 = "K" Then
       PB_END.BtnCaption = "戻る"
       PB_END.Visible = True
    End If
    
    lb_Kensu.Visible = Not H1Mode           ' 2019/07/12 ADD
    
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
            If .Name Like "Op" & wbc_Enm & "*" Then
               .Value = False
               .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "*Comb*" Then .ListIndex = 0
        End With
    Next gObj
    If HeadBody = "H" Then
        iH1_SDate = Format(Date, "YYYYMMDD")
        iH1_YM.Text = Format(Now(), "yyyyMM")   ' 2019/07/09 ADD
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
       If Not Head1Chk Then
          GoTo ReturnPress_Ed
       End If
       ProcHB = "B1"
       Call DispChange(ProcHB)
       DBGrid1.SetFocus
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
        
    '===============================================================================' 2019/07/09 ADD START
    ' 棚卸年月（必須項目）
    If Trim(iH1_YM.Text) = "" Then
        With CisFun
            .MB_Lines = 3
            .MB_MSG(2) = "    棚卸年月を入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If
    If Not iH1_YM.cChkResult Then
        With CisFun
            .MB_Lines = 3
            .MB_MSG(2) = "    棚卸年月を正しく入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If
    '===============================================================================' 2019/07/09 ADD END
    
    ' 日付
    If Trim(iH1_SDate) <> "" Then
       If CisFun.Val2(iH1_SDate.cDay) = 0 Then
          iH1_EDate = ""
       End If
    End If
    If Trim(iH1_SDate) <> "" And Trim(iH1_EDate) <> "" Then
       If iH1_SDate > iH1_EDate Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    作業日を正しく入力して下さい         "
                .MB_MSG(4) = "        開始　＞   終了             "
                .MB_Button = Error
                .MBOX
            End With
            iH1_SDate.SetFocus
            Exit Function
       End If
    End If
    
    If Not DBInput Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    該当データが存在しません｡            "
            .MB_MSG(3) = "    再度抽出条件を入力してください       "
            .MB_Title = "抽出処理"
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
    
    Head1Chk = True
End Function
'+----------------------------+
'+    表示切替ｵﾌﾟｼｮﾝﾎﾞﾀﾝ　     +
'+----------------------------+
Private Sub B1Op_Patn_Click(Index As Integer)
    Call CisFrm.MousePT(11)

    For Each gObj In B1Op_Patn
        If Index = gObj.Index Then
           gObj.BackColor = &H8080FF
        Else
           gObj.BackColor = &HC0C0C0
        End If
    Next gObj

    CisDBGrid.LayOutNow = Index + 1
    CisDBGrid.LayOutDisp

    If ProcHB = "B1" Then
        Call DBInput("R")
    End If
    Call CisFrm.MousePT(1)

End Sub
'+---------------------------+
'+    重複表示　ﾁｪｯｸﾎﾞｯｸｽ     +
'+---------------------------+
Private Sub Op_Overlap_GotFocus(Index As Integer)
    For Each gObj In Op_Overlap
        If gObj.Index = Index Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
Private Sub Op_Overlap_Click(Index As Integer)

    If Not FormAct Then Exit Sub
    
    Call CisFrm.MousePT(11)

    Call CisDBGrid.MargeDisp(Op_Overlap(1).Value)

    If ProcHB = "H1" Then
        iH1_SDate.SetFocus
    Else
        DBGrid1.SetFocus
    End If

    Call CisFrm.MousePT(1)
End Sub
'+---------------------------+
'+    分類表示　ﾁｪｯｸﾎﾞｯｸｽ     +
'+---------------------------+
Private Sub Op_Sort_GotFocus(Index As Integer)
    For Each gObj In Op_Sort
        If gObj.Index = Index Then
           gObj.BackColor = mGotColor
        Else
           gObj.BackColor = mLostColor
        End If
    Next gObj
End Sub
Private Sub Op_Sort_Click(Index As Integer)
    If ProcHB = "H1" Then Exit Sub
    Call DBInput
End Sub
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示     *
'*                                    *
'**************************************
Private Function DBInput(Optional InType As String = "") As Boolean
    Dim DateStr As String
    Dim DateSt  As String
    Dim DateEd  As String
    Dim BinSt   As String
    Dim BinEd   As String
    
    DBInput = True
    
    Call CisFrm.MousePT(11)
    
    On Error GoTo DBInput_Err:
' 抽出中メッセージ表示
    PNL_Wait.Visible = True
    PNL_PF.Enabled = False
    DoEvents

'===============================================================================================================================' 2019/07/11 DEL START
'' 初期表示時、条件文作成(Where)
'    If ProcHB = "H1" And InType = "" Then
''       表示条件 セット
'       CisDBGrid.Select_Where = ""
'       With CisFun
'            .Where_CreateKB = NewDocument
'            ' 日付 月指定
'            If Trim(iH1_SDate) <> "" And CisFun.Val2(iH1_SDate.cDay) = 0 Then
'               .Where_Create AlphaMe, "Substring(" & DateStr & "作業日,1,6)", iH1_SDate
'            End If
'            ' 日付開始編集
'            If Trim(iH1_SDate) <> "" Then
'               DateSt = iH1_SDate
'            End If
'            ' 日付終了編集
'            If Trim(iH1_EDate) <> "" Then
'               DateEd = iH1_EDate
'            End If
'            ' 日付開始のみ
'            If Trim(DateSt) <> "" And Trim(DateEd) = "" And CisFun.Val2(iH1_SDate.cDay) <> 0 Then
'               .Where_Create AlphaMe, DateStr & "作業日", DateSt
'            End If
'            ' 日付終了のみ
'            If Trim(DateSt) = "" And Trim(DateEd) <> "" Then
'               .Where_Create AlphaMe, DateStr & "作業日", DateEd, 以下
'            End If
'            ' 日付開始・終了
'            If Trim(DateSt) <> "" And Trim(DateEd) <> "" Then
'                  .Where_Create AlphaMe, DateStr & "作業日", DateSt, Between, , DateEd
'            End If
'
'            ' 部署
'            .Where_Create AlphaMe, "SN.部門", iH1_Torcd
'            ' 品番
'        ''''.Where_Create AlphaMe, "SN.品番", iH1_Hinbn                                     ' 2009.12.28 del
'            .Where_Create AlphaMe, "REPLACE(SN.品番, '-', '')", Replace(iH1_Hinbn, "-", "") ' 2009.12.28 add
'            ' 機械№
'            .Where_Create AlphaMe, "SN.機械NO", iH1_Kikai
'            ' 作業者
'            .Where_Create AlphaMe, "SN.作業者", iH1_Sagyo
'            CisDBGrid.Select_Where = .Where_Phrase
'       End With
'    End If
'
'''''  " WHERE 作業日 BetWeen '20190709' And '20190731' AND   SN.部門 = '1010' AND   REPLACE(SN.品番, '-', '') = '789' AND   SN.作業者 = 'なべ'"
'''''  " WHERE 作業日 = '20190709' AND   SN.部門 = '1010'"
'''''  " WHERE 作業日 <= '20190731' AND   SN.作業者 = '2100'"
''Call Str_Monitor(CisDBGrid.Select_Where)
''Dim aaa As String
''aaa = CisDBGrid.Select_Where
''MsgBox (aaa)
'
'
'    '-- データ存在確認 -------
'    gSL_Select = ""
'    gSL_Select = gSL_Select & "    SELECT SN.作業日報管理NO, SN.作業日報管理行," ' SN.作業日,"
'    gSL_Select = gSL_Select & "       SUBSTRING(SN.作業日,1,4) + '/' + SUBSTRING(SN.作業日,5,2) + '/' + SUBSTRING(SN.作業日,7,2) AS 作業日,"
'    gSL_Select = gSL_Select & "       ISNULL(SN.部門,  '') + ':' + ISNULL(TM.正式名称,'') AS 部門  ,"
'    gSL_Select = gSL_Select & "       ISNULL(SN.作業者,'') + ':' + ISNULL(M1.値名称  ,'') AS 作業者,"
'    gSL_Select = gSL_Select & "       ISNULL(SN.機械NO,'') + ':' + ISNULL(M2.値名称  ,'') AS 機械NO,"
'    gSL_Select = gSL_Select & "       ISNULL(H1.表示品番,'') AS 品番,"
'    gSL_Select = gSL_Select & "       SN.工順 , SN.加工数,"             ' 2009.12.28 upd : 工程→工順
'   'gSL_Select = gSL_Select & "       SUBSTRING(SN.開始,1,2)+ ':'+ SUBSTRING(SN.開始,3,2) AS 開始  ,"
'   'gSL_Select = gSL_Select & "       SUBSTRING(SN.終了,1,2)+ ':'+ SUBSTRING(SN.終了,3,2) AS 終了  ,"
'    gSL_Select = gSL_Select & "       CASE WHEN ISNULL(SN.開始,'') = '' THEN ''"
'    gSL_Select = gSL_Select & "       ELSE substring('0000'+SN.開始,len('0000'+SN.開始)-4+1,2) + ':'"
'    gSL_Select = gSL_Select & "          + SUBSTRING('0000'+SN.開始,len('0000'+sn.開始)-4+3,2)"
'    gSL_Select = gSL_Select & "       END 開始,"
'    gSL_Select = gSL_Select & "       CASE WHEN ISNULL(SN.終了,'') = '' THEN ''"
'    gSL_Select = gSL_Select & "       ELSE substring('0000'+SN.終了,len('0000'+SN.終了)-4+1,2) + ':'"
'    gSL_Select = gSL_Select & "          + SUBSTRING('0000'+SN.終了,len('0000'+sn.終了)-4+3,2)"
'    gSL_Select = gSL_Select & "       END 終了,"
'    gSL_Select = gSL_Select & "       SN.不良数, SN.段取"
'   'gSL_Select = gSL_Select & "      ,M1.値名称 AS 作業者名, M2.値名称 AS 機械NO名"
'   'gSL_Select = gSL_Select & "      ,TM.正式名称 AS 部門名"
'    gSL_Select = gSL_Select & "       ,HT.基準工数1 基準工数,HT.基準工数2 目標工数"             '2013/09/25
'    gSL_Select = gSL_Select & " FROM  作業日報テーブル SN"
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ M1"
'    gSL_Select = gSL_Select & "            ON  M1.区分名称 = '担当管理'"
'    gSL_Select = gSL_Select & "            AND M1.値       = RTRIM(SN.作業者)"
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ M2"
'    gSL_Select = gSL_Select & "            ON  M2.区分名称 = '機械NO'"
'    gSL_Select = gSL_Select & "            AND M2.値       = RTRIM(SN.機械NO)"
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TM"
'    gSL_Select = gSL_Select & "            ON  TM.取引先CD = RTRIM(SN.部門)"
'    gSL_Select = gSL_Select & "            AND TM.社内区分 = 1"
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 品番マスタ H1"
'  ''gSL_Select = gSL_Select & "            ON  H1.品番     = SN.品番"
'    gSL_Select = gSL_Select & "            ON  REPLACE(H1.品番, '-', '') = REPLACE(SN.品番, '-', '')"
''==============- 2013/09/25 START
'    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 品番工順マスタ HT"
'    gSL_Select = gSL_Select & "            ON  REPLACE(HT.品番, '-', '') = REPLACE(SN.品番, '-', '')"
'    gSL_Select = gSL_Select & "            AND HT.取引先 = RTRIM(SN.部門)"
'    gSL_Select = gSL_Select & "            AND HT.受入 = RTRIM(SN.受入)"
'    gSL_Select = gSL_Select & "            AND HT.品目 = 1"
'    gSL_Select = gSL_Select & "            AND HT.工順 = 1"
''==============- 2013/09/25 END
''-----------------------------------------------------------------------------------------------' 2009.12.22 add start
'' 分類内容(Order By句) ｾｯﾄ
'    With CisDBGrid
'        .OrderSet (Op_Sort(0).Value)
'        If .Select_OrderBy = "" Then
'            DBInput = False
'            GoTo DBInput_Ed
'        End If
'    End With
'    gSL_Select = gSL_Select & CisDBGrid.Select_Where & CisDBGrid.Select_OrderBy
''-----------------------------------------------------------------------------------------------' 2009.12.22 add end
''
''-----------------------------------------------------------------------------------------------' 2009.12.22 del start
''    gSL_Select = gSL_Select & CisFun.Where_Phrase       ' Where 結合
''    If Op_Sort(0).Value Then                            ' Order 結合
''       gSL_Select = gSL_Select & " ORDER BY SN.作業日, SN.作業日報管理NO, SN.作業日報管理行"
''    Else
''       gSL_Select = gSL_Select & " ORDER BY SN.作業日 DESC, SN.作業日報管理NO DESC, SN.作業日報管理行 DESC"
''    End If
''-----------------------------------------------------------------------------------------------' 2009.12.22 del end
'===============================================================================================================================' 2019/07/11 DEL END
    
    
    '===================================================================================================================' 2019/07/10 ADD START
'    ''＜＜　検索条件作成　＞＞
'    Dim wSQLW As String
'    Dim wSYMD As String
'    Dim wEYMD As String
'    Dim wCYMD As String
'    Dim wKako As String
'
'    wCYMD = Trim(iH1_SDate.Text)
'    If wCYMD = "" Then
'        wSYMD = "0000/00/00"
'    Else
'        wSYMD = Mid(wCYMD, 1, 4) & "/" & Mid(wCYMD, 5, 2) & "/" & Mid(wCYMD, 7, 2)
'    End If
'    wCYMD = Trim(iH1_EDate.Text)
'    If wCYMD = "" Then
'        wEYMD = "9999/99/99"
'    Else
'        wEYMD = Mid(wCYMD, 1, 4) & "/" & Mid(wCYMD, 5, 2) & "/" & Mid(wCYMD, 7, 2)
'    End If
'
'    wSQLW = ""
'        wSQLW = wSQLW & vbCrLf & " WHERE YK.対象年月 = '" + iH1_YM.Text + "'"
'    If Trim(iH1_SDate.Text) <> "" Or Trim(iH1_EDate.Text) <> "" Then
'        wSQLW = wSQLW & vbCrLf & "   AND YK.処理日   BETWEEN '" + wSYMD + "' AND '" + wEYMD + "'"
'    End If
'    If Trim(iH1_Torcd.Text) <> "" Then
'        wSQLW = wSQLW & vbCrLf & "   AND HM.得意先   = '" + Trim(iH1_Torcd.Text) + "'"
'    End If
'    If Trim(iH1_Hinbn.Text) <> "" Then
'        wSQLW = wSQLW & vbCrLf & "   AND REPLACE(HM.品番, '-', '') LIKE '" + Trim(Replace(iH1_Hinbn.Text, "-", "")) + "%'"
'    End If
'    If Trim(iH1_Sagyo.Text) <> "" Then
'        wSQLW = wSQLW & vbCrLf & "   AND YK.担当者   LIKE '" + Trim(iH1_Sagyo.Text) + "%'"
'    End If
'    Call CisComb1.CodeGet
'    wKako = CisComb1.Code
'    If wKako <> "" Then
'        wSQLW = wSQLW & vbCrLf & "   AND HM.品番区分 = " & wKako
'    End If

'' 初期表示時、条件文作成(Where)
    If ProcHB = "H1" And InType = "" Then
        ''＜＜　検索条件作成　＞＞
        CisDBGrid.Select_Where = ""
        With CisFun
            .Where_CreateKB = NewDocument
            ' 棚卸年月
            If Trim(iH1_YM) <> "" Then
                .Where_Create AlphaMe, "YK.対象年月", iH1_YM
            End If
            ' 作業日：日付開始編集
            If Trim(iH1_SDate) <> "" Then
                DateSt = Mid(iH1_SDate, 1, 4) & "/" & Mid(iH1_SDate, 5, 2) & "/" & Mid(iH1_SDate, 7, 2)
            End If
            ' 作業日：日付終了編集
            If Trim(iH1_EDate) <> "" Then
                DateEd = Mid(iH1_EDate, 1, 4) & "/" & Mid(iH1_EDate, 5, 2) & "/" & Mid(iH1_EDate, 7, 2)
            End If
            ' 日付開始のみ
            If Trim(DateSt) <> "" And Trim(DateEd) = "" And CisFun.Val2(iH1_SDate.cDay) <> 0 Then
                .Where_Create AlphaMe, "YK.処理日", DateSt
            End If
            ' 日付終了のみ
            If Trim(DateSt) = "" And Trim(DateEd) <> "" Then
                .Where_Create AlphaMe, "YK.処理日", DateEd, 以下
            End If
            ' 日付開始・終了
            If Trim(DateSt) <> "" And Trim(DateEd) <> "" Then
                .Where_Create AlphaMe, "YK.処理日", DateSt, Between, , DateEd
            End If
            ' 契約先
            .Where_Create AlphaMe, "HM.得意先", iH1_Torcd
            ' 品番
            .Where_Create AlphaMe, "HM.品番", iH1_Hinbn
            ' 担当者
            .Where_Create AlphaMe, "YK.担当者", iH1_Sagyo
            ' 加工度
            Call CisComb1.CodeGet
            If Trim(CisComb1.Code) <> "" Then
               .Where_Create Numeric, "HM.品番区分", CisFun.Val2(CisComb1.Code), , Create
            End If
            CisDBGrid.Select_Where = .Where_Phrase
        End With
    End If
    Call Str_Monitor(CisDBGrid.Select_Where)
    
    '-- データ存在確認 -------
    gSL_Select = ""
    gSL_Select = gSL_Select & vbCrLf & "SELECT YK.対象年月,"
    gSL_Select = gSL_Select & vbCrLf & "       YK.処理日 作業日,"
    gSL_Select = gSL_Select & vbCrLf & "       CASE WHEN ISNULL(HM.得意先, '') = ''"
    gSL_Select = gSL_Select & vbCrLf & "            THEN ISNULL(TR.略称  , '')"
    gSL_Select = gSL_Select & vbCrLf & "            ELSE HM.得意先 + ':' + ISNULL(TR.略称  , '')"
    gSL_Select = gSL_Select & vbCrLf & "       END  契約先,"
    gSL_Select = gSL_Select & vbCrLf & "       ISNULL(HM.表示品番, '') 品番,"
    gSL_Select = gSL_Select & vbCrLf & "       HT.背番号,"
    gSL_Select = gSL_Select & vbCrLf & "       SUM(YK.実棚数) 実棚数,"
    gSL_Select = gSL_Select & vbCrLf & "       ISNULL(ZM.在庫数,  0) 理論在庫数,"
    gSL_Select = gSL_Select & vbCrLf & "       ISNULL(MM.値名称, '') 加工度,        /* HM.品番区分 */"
    gSL_Select = gSL_Select & vbCrLf & "       ISNULL("
    gSL_Select = gSL_Select & vbCrLf & "           CASE HM.品番区分"
    gSL_Select = gSL_Select & vbCrLf & "               WHEN 1 /* 製品     */"
    gSL_Select = gSL_Select & vbCrLf & "                      THEN  (SELECT TOP(1) 単価 FROM 受注単価マスタ"
    gSL_Select = gSL_Select & vbCrLf & "                             WHERE 品番   = YK.品番"
    gSL_Select = gSL_Select & vbCrLf & "                               AND 取引先 = HT.納入先"
    gSL_Select = gSL_Select & vbCrLf & "                               AND (適用開始日付 <= YK.処理日"
    gSL_Select = gSL_Select & vbCrLf & "                                    And (適用終了日付  = '' Or (適用終了日付 <> ''"
    gSL_Select = gSL_Select & vbCrLf & "                                                           And  適用終了日付 >= YK.処理日)))"
    gSL_Select = gSL_Select & vbCrLf & "                             ORDER BY 適用開始日付 DESC, 適用終了日付, 単価 DESC, 履歴NO DESC)"
    gSL_Select = gSL_Select & vbCrLf & "               WHEN 2 /* 仕掛品   */"
    gSL_Select = gSL_Select & vbCrLf & "                      THEN /*HT.基準単価*/"
    gSL_Select = gSL_Select & vbCrLf & "                           (SELECT MAX(基準単価) 基準単価"
    gSL_Select = gSL_Select & vbCrLf & "                              FROM 品番取引先マスタ"
    gSL_Select = gSL_Select & vbCrLf & "                             WHERE 品番   = YK.品番"
    gSL_Select = gSL_Select & vbCrLf & "                             /*AND 取引先 = HM.得意先*/)"
    gSL_Select = gSL_Select & vbCrLf & "               WHEN 3 /* 購入品   */"
    gSL_Select = gSL_Select & vbCrLf & "                      THEN (SELECT TOP(1) 単価 FROM 購入単価マスタ"
    gSL_Select = gSL_Select & vbCrLf & "                             WHERE 品番 = YK.品番"
    gSL_Select = gSL_Select & vbCrLf & "                               AND (適用開始日付 <= YK.処理日"
    gSL_Select = gSL_Select & vbCrLf & "                                    And (適用終了日付  = '' Or (適用終了日付 <> ''"
    gSL_Select = gSL_Select & vbCrLf & "                                                           And  適用終了日付 >= YK.処理日)))"
    gSL_Select = gSL_Select & vbCrLf & "                             ORDER BY 適用開始日付 DESC, 適用終了日付, 単価 DESC, 履歴NO DESC)"
    gSL_Select = gSL_Select & vbCrLf & "               WHEN 5 /* 長期在庫 */"
    gSL_Select = gSL_Select & vbCrLf & "                      THEN 1"
    gSL_Select = gSL_Select & vbCrLf & "               ELSE 0 /* 4 無償品　または該当なしの場合 */"
    gSL_Select = gSL_Select & vbCrLf & "           END"
    gSL_Select = gSL_Select & vbCrLf & "       , 0) AS 単価,"
    gSL_Select = gSL_Select & vbCrLf & "       ROUND(ISNULL("
    gSL_Select = gSL_Select & vbCrLf & "           CASE HM.品番区分"
    gSL_Select = gSL_Select & vbCrLf & "               WHEN 1 /* 製品     */"
    gSL_Select = gSL_Select & vbCrLf & "                      THEN  (SELECT TOP(1) 単価 FROM 受注単価マスタ"
    gSL_Select = gSL_Select & vbCrLf & "                             WHERE 品番   = YK.品番"
    gSL_Select = gSL_Select & vbCrLf & "                               AND 取引先 = HT.納入先"
    gSL_Select = gSL_Select & vbCrLf & "                               AND (適用開始日付 <= YK.処理日"
    gSL_Select = gSL_Select & vbCrLf & "                                    And (適用終了日付  = '' Or (適用終了日付 <> ''"
    gSL_Select = gSL_Select & vbCrLf & "                                                           And  適用終了日付 >= YK.処理日)))"
    gSL_Select = gSL_Select & vbCrLf & "                             ORDER BY 適用開始日付 DESC, 適用終了日付, 単価 DESC, 履歴NO DESC)"
    gSL_Select = gSL_Select & vbCrLf & "               WHEN 2 /* 仕掛品   */"
    gSL_Select = gSL_Select & vbCrLf & "                      THEN /*HT.基準単価*/"
    gSL_Select = gSL_Select & vbCrLf & "                           (SELECT MAX(基準単価) 基準単価"
    gSL_Select = gSL_Select & vbCrLf & "                              FROM 品番取引先マスタ"
    gSL_Select = gSL_Select & vbCrLf & "                             WHERE 品番   = YK.品番"
    gSL_Select = gSL_Select & vbCrLf & "                             /*AND 取引先 = HM.得意先*/)"
    gSL_Select = gSL_Select & vbCrLf & "               WHEN 3 /* 購入品   */"
    gSL_Select = gSL_Select & vbCrLf & "                      THEN (SELECT TOP(1) 単価 FROM 購入単価マスタ"
    gSL_Select = gSL_Select & vbCrLf & "                             WHERE 品番 = YK.品番"
    gSL_Select = gSL_Select & vbCrLf & "                               AND (適用開始日付 <= YK.処理日"
    gSL_Select = gSL_Select & vbCrLf & "                                    And (適用終了日付  = '' Or (適用終了日付 <> ''"
    gSL_Select = gSL_Select & vbCrLf & "                                                           And  適用終了日付 >= YK.処理日)))"
    gSL_Select = gSL_Select & vbCrLf & "                             ORDER BY 適用開始日付 DESC, 適用終了日付, 単価 DESC, 履歴NO DESC)"
    gSL_Select = gSL_Select & vbCrLf & "               WHEN 5 /* 長期在庫 */"
    gSL_Select = gSL_Select & vbCrLf & "                      THEN 1"
    gSL_Select = gSL_Select & vbCrLf & "               ELSE 0 /* 4 無償品　または該当なしの場合 */"
    gSL_Select = gSL_Select & vbCrLf & "           END"
    gSL_Select = gSL_Select & vbCrLf & "       , 0) * SUM(YK.実棚数), 0) 金額,"
    gSL_Select = gSL_Select & vbCrLf & "       ISNULL(YK.置場,   '') 置場,"
    gSL_Select = gSL_Select & vbCrLf & "       ISNULL(YK.担当者, '') 担当者,"
    gSL_Select = gSL_Select & vbCrLf & "       ISNULL(YK.読取連番,0) 読取連番"
    gSL_Select = gSL_Select & vbCrLf & "  FROM 棚卸読取結果テーブル YK"
    gSL_Select = gSL_Select & vbCrLf & "  LEFT OUTER JOIN 品番マスタ       HM"
    gSL_Select = gSL_Select & vbCrLf & "    ON HM.品番 = YK.品番"
    gSL_Select = gSL_Select & vbCrLf & "  LEFT OUTER JOIN"
    gSL_Select = gSL_Select & vbCrLf & "      (SELECT YK.品番,"
    gSL_Select = gSL_Select & vbCrLf & "              ISNULL(MAX(HT.背番号), '') 背番号,"
    gSL_Select = gSL_Select & vbCrLf & "              ISNULL(    HT.納入先 , '') 納入先"
    gSL_Select = gSL_Select & vbCrLf & "         FROM 棚卸読取結果テーブル YK"
    gSL_Select = gSL_Select & vbCrLf & "         LEFT OUTER join 品番取引先マスタ HT"
    gSL_Select = gSL_Select & vbCrLf & "           ON HT.品番 = YK.品番"
    gSL_Select = gSL_Select & vbCrLf & "          AND HT.品目 = 1"
    gSL_Select = gSL_Select & vbCrLf & "        GROUP BY YK.品番, 納入先"
    gSL_Select = gSL_Select & vbCrLf & "      )                            HT"
    gSL_Select = gSL_Select & vbCrLf & "    ON HT.品番 = YK.品番"
    gSL_Select = gSL_Select & vbCrLf & "  LEFT OUTER JOIN 取引先マスタ     TR"
    gSL_Select = gSL_Select & vbCrLf & "    ON TR.取引先CD   = HM.得意先"
    gSL_Select = gSL_Select & vbCrLf & "   AND TR.取引先区分 = 0"
    gSL_Select = gSL_Select & vbCrLf & "  LEFT OUTER JOIN 名称マスタ       MM"
    gSL_Select = gSL_Select & vbCrLf & "    ON MM.区分名称 = '品番区分'"
    gSL_Select = gSL_Select & vbCrLf & "   AND MM.値       = HM.品番区分"
    gSL_Select = gSL_Select & vbCrLf & "  LEFT OUTER JOIN 在庫マスタ       ZM"
    gSL_Select = gSL_Select & vbCrLf & "    ON ZM.品番     = YK.品番"
    gSL_Select = gSL_Select & vbCrLf & "   AND ZM.年月度   = YK.対象年月"
    
  ''gSL_Select = gSL_Select & vbCrLf & wSQLW
    gSL_Select = gSL_Select & vbCrLf & CisDBGrid.Select_Where
    
    gSL_Select = gSL_Select & vbCrLf & " GROUP BY YK.対象年月, YK.処理日,"
    gSL_Select = gSL_Select & vbCrLf & "          HT.納入先  , HM.得意先, TR.略称,"
    gSL_Select = gSL_Select & vbCrLf & "          HM.表示品番, HT.背番号, YK.品番,"
    gSL_Select = gSL_Select & vbCrLf & "          ZM.在庫数,"
    gSL_Select = gSL_Select & vbCrLf & "          HM.品番区分, MM.値名称,"
    gSL_Select = gSL_Select & vbCrLf & "          YK.置場    , YK.担当者, YK.読取連番" + vbCrLf
  ''gSL_Select = gSL_Select & vbCrLf & " ORDER BY 品番, 契約先, 作業日"
    
    ' 分類内容(Order By句) ｾｯﾄ
    With CisDBGrid
        .OrderSet (Op_Sort(0).Value)
        If .Select_OrderBy = "" Then
            DBInput = False
            GoTo DBInput_Ed
        End If
    End With
    gSL_Select = vbCrLf & gSL_Select & CisDBGrid.Select_OrderBy
    
    Call Str_Monitor(gSL_Select)            ' デバッグ用（外さなくても問題は無い）
    '===================================================================================================================' 2019/07/10 ADD END

' グリッド内表示
    If Not DBGridDisp Then
        DBInput = False
        GoTo DBInput_Ed
    End If
    On Error Resume Next
' 表示データなし
    On Error GoTo 0
    If Adodc1.Recordset.EOF Then
        DBInput = False
        GoTo DBInput_Ed
    End If
    
' 抽出件数表示
    lb_Kensu.Visible = True
    Call KensuDisp
    
    On Error GoTo 0
DBInput_Ed:
    Call CisFrm.MousePT(0)
    
    PNL_Wait.Visible = False
    PNL_PF.Enabled = True
    
    '=======================' 2019/07/09 DEL START：エラーになるので外す（Win7から？原因不明）  ' 2019/08/09 UPD：復帰
    SendKeys "{END}"
    SendKeys "{RIGHT}"
    SendKeys "{LEFT}"
    SendKeys "{HOME}"
    '=======================' 2019/07/09 DEL END
    Exit Function
DBInput_Err:
    Call CisFrm.MousePT(0)
    CisFun.ErrorBox
    PNL_Wait.Visible = False
    PNL_PF.Enabled = True
End Function
'+----------------------------------------------------------+
'+                                                          +
'+                                                          +
'+      グリッド処理(TrueDBGrid)                             +
'+                                                          +
'+                                                          +
'+----------------------------------------------------------+
'+----------------------------------------------------------+
'+   グリッド内容初期設定
'+----------------------------------------------------------+
Private Sub GridInitSet()
' Call .GridInitSet(ﾏｰｼﾞ処理,分類処理,ﾃﾞｰﾀ連結名,印刷見出し,印刷桁数,合計印刷)
    With CisDBGrid
        Call .GridInitSet(, , "作業日", "作業日", 10)
        Call .GridInitSet(, , "契約先", "契約先", 20)
        Call .GridInitSet(, , "品番", "品　　番", 30)
        Call .GridInitSet(, , "背番号", "背番号", 5)
        Call .GridInitSet(, , "実棚数", "実棚数", 7)
        Call .GridInitSet(, , "理論在庫数", "理論在庫数", 20)
        Call .GridInitSet(, , "加工度", "加工度", 8)
        Call .GridInitSet(, , "単価", "単価", 7)
        Call .GridInitSet(, , "金額", "金額", 10)
        Call .GridInitSet(, , "置場", "置場", 10)
        Call .GridInitSet(, , "担当者", "担当者", 20)
        Call .GridInitSet(, , "読取連番", "連番", 6)
    
'         Call .GridInitSet(True, , "作業日", "作業日", 10)
'         Call .GridInitSet(True, , "部門", "部　門", 14)
'         Call .GridInitSet(True, , "作業者", "作業者", 14)
'         Call .GridInitSet(True, , "機械NO", "機械№", 10)
'         Call .GridInitSet(True, , "品番", "品　　番", 21)
'         Call .GridInitSet(True, , "工順", "工順", 4)
'         Call .GridInitSet(, , "加工数", "加工数", 10, True)
'         Call .GridInitSet(, , "開始", "開始", 5)
'         Call .GridInitSet(, , "終了", "終了", 5)
'         Call .GridInitSet(, , "不良数", "不良数", 8, True)
'         Call .GridInitSet(, , "段取", "段取", 3)
'         Call .GridInitSet(, , "作業日報管理NO", "手配NO", 10)
'         Call .GridInitSet(, , "基準工数", "基準工数", 8)
'         Call .GridInitSet(, , "目標工数", "目標工数", 8)
    End With
End Sub
Private Sub GridInitSetCsv()
' Call .GridInitSet(ﾏｰｼﾞ処理,分類処理,ﾃﾞｰﾀ連結名,印刷見出し,印刷桁数,合計印刷)
    With CisDBGrid
        Call .GridInitSet(, , "作業日", "作業日", 10)
        Call .GridInitSet(, , "契約先", "契約先", 20)
        Call .GridInitSet(, , "品番", "品　　番", 30)
        Call .GridInitSet(, , "背番号", "背番号", 5)
        Call .GridInitSet(, , "実棚数", "実棚数", 7)
        Call .GridInitSet(, , "理論在庫数", "理論在庫数", 20)
        Call .GridInitSet(, , "加工度", "加工度", 8)
        Call .GridInitSet(, , "単価", "単価", 7)
        Call .GridInitSet(, , "金額", "金額", 10)
        Call .GridInitSet(, , "置場", "置場", 10)
        Call .GridInitSet(, , "担当者", "担当者", 20)
        Call .GridInitSet(, , "読取連番", "連番", 6)
         
'         Call .GridInitSet(True, , "作業日", "作業日", 10)
'         Call .GridInitSet(True, , "部門", "部　門", 30)
'         Call .GridInitSet(True, , "作業者", "作業者", 20)
'         Call .GridInitSet(True, , "機械NO", "機械№", 16)
'         Call .GridInitSet(True, , "品番", "品　　番", 26)
'         Call .GridInitSet(True, , "工順", "工順", 4)
'         Call .GridInitSet(, , "加工数", "加工数", 10, True)
'         Call .GridInitSet(, , "開始", "開始", 5)
'         Call .GridInitSet(, , "終了", "終了", 5)
'         Call .GridInitSet(, , "不良数", "不良数", 8, True)
'         Call .GridInitSet(, , "段取", "段取", 3)
'         Call .GridInitSet(, , "作業日報管理NO", "手配NO", 10)
'         Call .GridInitSet(, , "基準工数", "基準工数", 8)
'         Call .GridInitSet(, , "目標工数", "目標工数", 8)
    End With
End Sub

'+----------------------------------------------------------+
'+      ADODC1(件数表示)                                     +
'+----------------------------------------------------------+
Private Sub KensuDisp()
    Static NowCnt       As Long
    Static MaxCnt       As Long
    NowCnt = Adodc1.Recordset.AbsolutePosition
    MaxCnt = Adodc1.Recordset.RecordCount
    If NowCnt < 0 Then Exit Sub
    lb_Kensu.Caption = Format(NowCnt, "###,##0") & " / " & _
                       Format(MaxCnt, "###,##0")

End Sub
'+---------------------+
'+  グリッド内容表示    +
'+---------------------+
Private Function DBGridDisp()
    DBGridDisp = False
    
    On Error GoTo DBGridDisp_ER:
    
    On Error Resume Next

    Adodc1.RecordSource = gSL_Select
    DBGrid1.ReBind
    Adodc1.Refresh
    DBGrid1.Refresh
    
' 重複表示
    Call CisDBGrid.MargeDisp(Op_Overlap(1).Value)
    Call KensuDisp

    With DBGrid1
         .SetFocus
    End With
    SendKeys "{HOME}"
    
    DBGridDisp = True
    
    On Error GoTo 0
    Exit Function
DBGridDisp_ER:
    If Err <> 40002 Then                        ' RDO タイムアウト
       CisFun.ErrorBox
    End If
End Function
'+-----------------------+
'+  グリッド内容初期化    +
'+-----------------------+
Private Sub DBGridClr()

    Call CisDBGrid.SplitsClr("", 0)             ' 分割ﾍﾟｲﾝ解除
    
    Adodc1.Recordset.Close
    DBGrid1.Close True
    
End Sub
'+----------------------------------------------------------+
'+ キー入力にて行移動 《 KeyUp イベント   》                  +
'+----------------------------------------------------------+
Private Sub DBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or _
        KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then
        Call KensuDisp
    Else
        Exit Sub
    End If
End Sub
'+----------------------------------------------------------+
'+ 表示項目移動後 分類表示 《 ColMove イベント   》            +
'+----------------------------------------------------------+
Private Sub DBGrid1_ColMove(ByVal Position As Integer, Cancel As Integer)
    CisDBGrid.ColMoveFlg = True
End Sub
'+----------------------------------------------------------+
'+ 表示項目移動後 分類表示 《 Paint イベント   》              +
'+----------------------------------------------------------+
Private Sub DBGrid1_Paint()
    If Not FormAct Then Exit Sub
    If CisDBGrid.ColMoveFlg Then                ' 再表示(分類)
        CisDBGrid.ColMoveFlg = False
        Call DBInput
    End If
End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseDown  イベント   》               +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If DBGrid1.RowContaining(Y) = -1 And _
        DBGrid1.ColContaining(x) = -1 Then
        Call CisDBGrid.SplitsClr("C", 0)        ' 分割ﾍﾟｲﾝのみ解除
        Call DBInput
        Exit Sub
    End If
    If DBGrid1.RowContaining(Y) < 0 Then        ' 分割ﾍﾟｲﾝ表示
        If Button = 2 Then
            Call CisDBGrid.SplitsAdd(DBGrid1.ColContaining(x))
            Call DBInput
        End If
        Exit Sub
    End If
    If DBGrid1.RowContaining(Y) < 0 Then Exit Sub
    If Button = 2 Then
        DBGrid1.Row = DBGrid1.RowContaining(Y)
      ''Call TotalDsp(DBGrid1.ColContaining(x)) ' 2019/07/08 DEL
    Else
        DBGrid1.Row = DBGrid1.RowContaining(Y)
    End If
    If DBGrid1.RowContaining(x) < 0 Then Exit Sub
    
    If Button = 2 Then
        DBGrid1.Row = DBGrid1.RowContaining(Y)
      ''Call TotalDsp(DBGrid1.ColContaining(x)) ' 2019/07/08 DEL
    Else
      ''If PNL_ExeMsg.Visible Then              ' 2019/07/08 DEL
      ''   PNL_ExeMsg.Visible = False           ' 2019/07/08 DEL
      ''End If                                  ' 2019/07/08 DEL
        DBGrid1.Row = DBGrid1.RowContaining(Y)
    End If
End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseUp  イベント   》                +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call KensuDisp
'===================================' 2019/07/08 DEL START
'    If Button = 2 Then
'       PNL_ExeMsgX = ""
'       lb_TKSu = "": lb_TFSu = ""
'       PNL_ExeMsg.Visible = False
'    End If
'===================================' 2019/07/08 DEL END
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++    '++ デバッグ用です +++
'+                                                 +
'+  文字列をイミディエイトウィンドウに表示する     +
'+                                                 +
'+++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub Str_Monitor(ByVal wStr As String)
    Debug.Print vbCrLf
    Debug.Print Format(Now, "yyyy/MM/dd HH:mm:ss") + _
                " .................................................." + _
                vbCrLf + wStr
End Sub


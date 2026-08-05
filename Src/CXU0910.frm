VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form CXU0910 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "売上入力"
   ClientHeight    =   11025
   ClientLeft      =   420
   ClientTop       =   795
   ClientWidth     =   15630
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11025
   ScaleWidth      =   15630
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   405
      Index           =   0
      Left            =   12480
      ScaleHeight     =   345
      ScaleWidth      =   2325
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3525
      Width           =   2385
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   345
         Index           =   1
         Left            =   15
         Top             =   0
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   609
         Caption         =   " 登録済件数"
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
            Left            =   1290
            Top             =   30
            Width           =   915
            _ExtentX        =   1614
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
   Begin VB.Timer Timer4 
      Left            =   210
      Top             =   810
   End
   Begin VB.PictureBox B2_Area1 
      Height          =   7170
      Left            =   435
      ScaleHeight     =   7110
      ScaleWidth      =   14415
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3960
      Width           =   14475
      Begin vsFlexLib.vsFlexArray VsGrid1 
         Height          =   6960
         Left            =   90
         TabIndex        =   2
         Top             =   90
         Width           =   14265
         _Version        =   196608
         _ExtentX        =   25162
         _ExtentY        =   12277
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "No |検 収 日  |納入番号 |品　番|赤黒|数量|仮区|単価|金額|背番号|車型|納入先|受入|検収管理NO|検収日|赤黒|納入先||||"
         Rows            =   20
         Cols            =   22
         BackColor       =   12648447
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   330
         SelectionMode   =   1
         AllowUserResizing=   1
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
      Height          =   1905
      Left            =   435
      ScaleHeight     =   1845
      ScaleWidth      =   11865
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2025
      Width           =   11925
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   885
         Left            =   60
         Top             =   45
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   1561
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
         Begin CisYMD_With_Btn.CisYMDwB iB1_KYmd 
            Height          =   345
            Left            =   120
            TabIndex        =   3
            Top             =   390
            Width           =   1905
            _ExtentX        =   2831
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
            Object.Width           =   1905
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   885
         Left            =   2205
         Top             =   45
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   1561
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
         cPositionY      =   30
         Begin CisText_V60.CisText iB1_NouNo 
            Height          =   360
            Left            =   150
            TabIndex        =   4
            Top             =   375
            Width           =   1230
            _ExtentX        =   2170
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
            Text            =   "XXXXXXXXX"
            MaxLength       =   9
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   885
         Left            =   3735
         Top             =   930
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "数 量"
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
         Begin CisText_V60.CisText iB1_Suryo 
            Height          =   360
            Left            =   165
            TabIndex        =   11
            Top             =   360
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   635
            cFormat         =   "#,###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "#,###,###"
            MaxLength       =   9
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "#######"
            cILength        =   7
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D19 
         Height          =   885
         Left            =   6315
         Top             =   930
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "単 価"
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
         Begin CisText_V60.CisText iB1_Tanka 
            Height          =   360
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   635
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
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   885
         Left            =   8145
         Top             =   930
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "金　額"
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
         Begin CisText_V60.CisText iB1_Kingk 
            Height          =   360
            Left            =   120
            TabIndex        =   14
            Top             =   345
            Width           =   1530
            _ExtentX        =   2699
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   885
         Index           =   0
         Left            =   2775
         Top             =   930
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "赤黒"
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
         Begin VB.CheckBox B1Chk_Aka 
            BackColor       =   &H00C0C0C0&
            Caption         =   "黒"
            Height          =   345
            Left            =   195
            TabIndex        =   10
            Top             =   360
            Width           =   555
         End
      End
      Begin Cis3D_v60.CIS3D BLB_Zaist 
         Height          =   885
         Left            =   3735
         Top             =   45
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1561
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
         cPositionX      =   90
         cPositionY      =   30
         Begin CisText_V60.CisText iB1_Hinbn 
            Height          =   360
            Left            =   105
            TabIndex        =   5
            Top             =   360
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
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   885
         Index           =   0
         Left            =   60
         Top             =   930
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1561
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisText_V60.CisText iB1_Seban 
            Height          =   360
            Left            =   285
            TabIndex        =   8
            Top             =   375
            Width           =   1155
            _ExtentX        =   2037
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
            Text            =   "XXXXXXXX"
            MaxLength       =   8
            IMEMode         =   2
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   885
         Left            =   11010
         Top             =   45
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1561
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
         cPositionY      =   30
         Begin CisText_V60.CisText iB1_Ukeir 
            Height          =   360
            Left            =   195
            TabIndex        =   7
            Top             =   375
            Width           =   390
            _ExtentX        =   688
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
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   885
         Left            =   1785
         Top             =   930
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "車型"
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
         Begin CisText_V60.CisText iB1_Syakei 
            Height          =   360
            Left            =   120
            TabIndex        =   9
            Top             =   375
            Width           =   765
            _ExtentX        =   1349
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   885
         Left            =   7110
         Top             =   45
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   1561
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
         Begin Cis3D_v60.CIS3D B1lb_Nounm 
            Height          =   360
            Left            =   1170
            Top             =   360
            Width           =   2505
            _ExtentX        =   4419
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin CisText_V60.CisText iB1_Noucd 
            Height          =   360
            Left            =   180
            TabIndex        =   6
            Top             =   360
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
      End
      Begin Cis3D_v60.CIS3D CIS3D10 
         Height          =   885
         Left            =   5265
         Top             =   930
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "仮単区"
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
         Begin VB.CheckBox B1Chk_Kari 
            BackColor       =   &H00C0C0C0&
            Caption         =   "正式"
            Height          =   345
            Left            =   90
            TabIndex        =   12
            Top             =   345
            Width           =   810
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   885
         Left            =   9885
         Top             =   930
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "売上照合"
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
         Begin VB.CheckBox B1Chk_Syogo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "対象外"
            Height          =   345
            Left            =   435
            TabIndex        =   15
            Top             =   345
            Width           =   1065
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
      Height          =   930
      Left            =   450
      ScaleHeight     =   870
      ScaleWidth      =   5610
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   870
      Width           =   5670
      Begin Cis3D_v60.CIS3D HLB_UYm 
         Height          =   795
         Left            =   60
         Top             =   45
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1402
         ForeColor       =   16711680
         Caption         =   "売上年月"
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
         Begin CisYM_V60.CisYM iH1_UYm 
            Height          =   390
            Left            =   150
            TabIndex        =   0
            Top             =   345
            Width           =   1485
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
      Begin Cis3D_v60.CIS3D HLB_Hinnm 
         Height          =   795
         Left            =   1815
         Top             =   45
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   1402
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
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1125
            Top             =   360
            Width           =   2505
            _ExtentX        =   4419
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   135
            TabIndex        =   1
            Top             =   360
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
      End
   End
   Begin VB.Timer Timer3 
      Left            =   1980
      Top             =   810
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
      Height          =   1365
      Left            =   10470
      ScaleHeight     =   1305
      ScaleWidth      =   4365
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   480
      Width           =   4425
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1755
         Top             =   660
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
         Left            =   2625
         Top             =   660
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
         Left            =   3480
         Top             =   660
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
         Left            =   2610
         Top             =   30
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
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   900
         Top             =   660
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
         Left            =   45
         Top             =   660
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1740
         Top             =   30
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
         Left            =   30
         Top             =   30
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
         Left            =   885
         Top             =   30
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
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   990
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   540
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Left            =   930
      Top             =   660
   End
   Begin VB.Timer Timer2 
      Left            =   1440
      Top             =   720
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
      Caption         =   "【　売上入力  】"
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
         Left            =   10620
         ScaleHeight     =   345
         ScaleWidth      =   3675
         TabIndex        =   18
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
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   8070
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
End
Attribute VB_Name = "CXU0910"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************************'
'**
'**       フォーム名    :   売上入力
'**       フォームID    :   CXU0910
'**       処理概要      :
'**
'**       作  成  日    :   2004/03/29  By CIS SIMPR-A
'**       変  更  日    :   2004/05/28  By CIS 背番号カナ入力・背番号半角チェック
'**       変  更  日    :   2004/11/26  By CIS 照合区分追加、売上明細テーブルへ出力追加
'**       変  更  日    :   2006/12/13  By CIS 修正時に売上明細テーブルが更新されない点を修正
'**       変  更  日    :   2008/04/11  By CIS 背番号8桁,車型5桁
'**
'********************************************************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)
    Dim CisVsGrid1      As New CisVsGrid3

    Dim SaveUYm         As String       ' 売上処理年月
    Dim SaveRow         As Integer
    Dim SaveKeisan      As Byte
    Dim RowColFlg       As Boolean

    Dim mGotColor       As Long
    Dim mLostColor      As Long
    
    Dim KDouble         As Double
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
'
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
'
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
'
    Dummy.Left = -1000
    Dummy.Enabled = False
'    Dummy2.Left = -1000
'    Dummy3.Left = -1000
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔(基本)
    Timer2.Enabled = False   ' 処理区分点滅禁止(基本)
'    Timer3.Interval = 500    ' 処理区分点滅間隔(取引先)
'    Timer3.Enabled = False   ' 処理区分点滅禁止(取引先)
'
    '+--------------------------------+
    '+  グリッド初期設定(材料⇒品番)
    '+--------------------------------+
    With CisVsGrid1
        Set .GridObj = VsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .SelectCol = 21
        .InitGet = False
        .Init
    End With
    Call GridInit1
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
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
'
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    If MeisyoGet("処理内容", "PXU0110") Then
        SaveUYm = RTrim(SYM.英数字1)
    End If
'
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
           Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click    '【新規】
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click    '【修正】
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
'           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
'           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_Can_Click    '【取消】
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
        Call CisVsGrid1.vsColor(消去個別)
        Call HeadBodyClear("B1")
        If Val(Dsp_Kensu) <> 0 Then
            ProcHB = "B2"
            SyoriKB = "MNT"
            Call VSGrid1_Click
        Else
            ProcHB = "H1"
            SyoriKB = ""
        End If
        Call SyoriName(SyoriKB)
        Call DispChange(ProcHB)
        GoTo PB_CAN_Ed
    End If
            
    If ProcHB = "B2" Then
        SyoriKB = ""
        Call SyoriName(SyoriKB)
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        GoTo PB_CAN_Ed
    End If

PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新　規  キ ー(F1)        +
'+----------------------------+
Private Sub PB_ADD_Click()
    
    If SyoriKB <> "ADD" Then
        ProcHB = "B1"
        Call DispChange(ProcHB)
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
        Call HeadBodyClear("B1")
        RowColFlg = True
        Call CisVsGrid1.vsColor(消去個別)
        RowColFlg = False
        iB1_KYmd.SetFocus
    End If
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)        +
'+----------------------------+
Private Sub PB_MNT_Click()
    
    If SyoriKB <> "MNT" Then
        ProcHB = "B2"
        Call DispChange(ProcHB)
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
        RowColFlg = True
        Call CisVsGrid1.vsColor(消去個別)
        RowColFlg = False

            
        VsGrid1.SetFocus
        Call VSGrid1_Click
    End If
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)        +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        ProcHB = "B2"
        Call DispChange(ProcHB)
        SyoriKB = "DEL"
        Call SyoriName(SyoriKB)
        RowColFlg = True
        Call CisVsGrid1.vsColor(消去個別)
        RowColFlg = False
        
        VsGrid1.SetFocus
        Call VSGrid1_Click
    End If

End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    
'''    If wMNTKb = "HIN" Then
'''        gSL_Select = " SELECT ZH.使用品番,HN.社内背番 背番号 "
'''        gSL_Select = gSL_Select & " FROM 材料品番マスタ ZH "
'''        gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
'''        gSL_Select = gSL_Select & "   ON HN.品番 = ZH.使用品番 "
'''        gSL_Select = gSL_Select & " WHERE ZH.使用品番 < '" & RTrim(iH1_Hinbn) & "'"
'''        gSL_Select = gSL_Select & " GROUP BY ZH.使用品番,HN.社内背番"
'''        gSL_Select = gSL_Select & " ORDER BY ZH.使用品番 Desc "
'''    Else
'''        gSL_Select = " SELECT ZR.材質"
'''        gSL_Select = gSL_Select & " FROM 材料マスタ ZR"
'''        gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料品番マスタ ZH"
'''        gSL_Select = gSL_Select & "   ON ZR.材料管理番号 = ZH.材料管理番号"
'''        gSL_Select = gSL_Select & " Where ZH.使用品番 Is Not Null"
'''        gSL_Select = gSL_Select & "   AND ZR.材質 < '" & RTrim(iH1_Hinbn) & "'"
'''        gSL_Select = gSL_Select & " GROUP BY ZR.材質"
'''        gSL_Select = gSL_Select & " ORDER BY ZR.材質 DESC"
'''    End If
'''
'''    Call DBInput("Back")
'''    Dummy.Enabled = True: Dummy.SetFocus
'''    iH1_Hinbn.SetFocus
'''    Dummy.Enabled = False
    
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    
'''    If wMNTKb = "HIN" Then
'''        gSL_Select = " SELECT ZH.使用品番,HN.社内背番 背番号"
'''        gSL_Select = gSL_Select & " FROM 材料品番マスタ ZH "
'''        gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
'''        gSL_Select = gSL_Select & "   ON HN.品番 = ZH.使用品番 "
'''        gSL_Select = gSL_Select & " WHERE ZH.使用品番 > '" & RTrim(iH1_Hinbn) & "'"
'''        gSL_Select = gSL_Select & " GROUP BY ZH.使用品番,HN.社内背番"
'''        gSL_Select = gSL_Select & " ORDER BY ZH.使用品番 "
'''    Else
'''        gSL_Select = " SELECT ZR.材質"
'''        gSL_Select = gSL_Select & " FROM 材料マスタ ZR"
'''        gSL_Select = gSL_Select & " LEFT OUTER JOIN 材料品番マスタ ZH"
'''        gSL_Select = gSL_Select & "   ON ZR.材料管理番号 = ZH.材料管理番号"
'''        gSL_Select = gSL_Select & " Where ZH.使用品番 Is Not Null"
'''        gSL_Select = gSL_Select & "   AND ZR.材質 > '" & RTrim(iH1_Hinbn) & "'"
'''        gSL_Select = gSL_Select & " GROUP BY ZR.材質 "
'''        gSL_Select = gSL_Select & " ORDER BY ZR.材質 "
'''    End If
'''
'''
'''    Call DBInput("Next")
'''    Dummy.Enabled = True: Dummy.SetFocus
'''    iH1_Hinbn.SetFocus
'''    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
'
    If PB_Look.Tag = "iB1_KYmd" Then
        iB1_KYmd.ShowCalender
        Exit Sub
    End If
    If PB_Look.Tag = "iB1_Noucd" Then Call Look_Noucd: Exit Sub
    If PB_Look.Tag = "iB1_Ukeir" Then Call Look_Ukeir: Exit Sub
    If PB_Look.Tag = "iB1_Hinbn" Then Call Look_Hinbn: Exit Sub

End Sub
'+----------------------------+
'+       契約先検索            +
'+----------------------------+
Private Sub Look_Torcd()

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 0
    RV_TorKb1 = 1
    
    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RTrim(RV_TorcdK)
        Call iH1_Torcd_LostFocus
    End If
End Sub
'+----------------------------+
'+       品番検索              +
'+----------------------------+
Private Sub Look_Hinbn()

    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
    RVI_Hinmoku = 0
    RVI_SK = 9
    RVI_Torcd = RTrim(iB1_Noucd)
    RVI_Ukeir = RTrim(iB1_Ukeir)
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iB1_Hinbn = RTrim(RV_Hinbn)
        iB1_Seban = RTrim(RV_Seban)
    End If
End Sub
'+----------------------------+
'+       納入先検索            +
'+----------------------------+
Private Sub Look_Noucd()

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 0
    RV_TorKb1 = 0
    
    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
        iB1_Noucd = RTrim(RV_TorcdK)
        Call iB1_Noucd_LostFocus
    End If
End Sub
'+----------------------------+
'+     納入先受入検索          +
'+----------------------------+
Private Sub Look_Ukeir()

    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 3
    RV_Torcd = RTrim(iB1_Noucd)
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iB1_Noucd = RTrim(RV_TorcdK)
        Call iB1_Noucd_LostFocus
        iB1_Ukeir = RTrim(RV_Ukeir)
    End If
End Sub
'+---( H1 )----------------------------------------------------------------------------
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
    If TorNmGet(iH1_Torcd, 2) Then
        H1lb_Tornm = TRM.略称
    End If
End Sub
'+---( B1 )----------------------------------------------------------------------------
'+-----------------------------+
'+  iB1_Hinbn
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
'+  iB1_KYmd
'+-----------------------------+
Private Sub iB1_KYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_KYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_Noucd
'+-----------------------------+
Private Sub iB1_Noucd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Noucd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    B1lb_Nounm = ""
    If TorNmGet(iB1_Noucd, 3) Then
        B1lb_Nounm = TRM.略称
    End If
End Sub
'+-----------------------------+
'+  iB1_Ukeir
'+-----------------------------+
Private Sub iB1_Ukeir_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Ukeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  B1Chk_Aka
'+-----------------------------+
Private Sub B1Chk_Aka_Click()
    If B1Chk_Aka.Value = 0 Then
        B1Chk_Aka.Caption = "黒"
    Else
        B1Chk_Aka.Caption = "赤"
    End If
End Sub
'+-----------------------------+
'+  iB1_Suryo
'+-----------------------------+
Private Sub iB1_Suryo_LostFocus()
    If CisFun.Val2(iB1_Suryo) <> 0 And CisFun.Val2(iB1_Tanka) <> 0 Then
        KDouble = CisFun.Val2(iB1_Tanka) * CisFun.Val2(iB1_Suryo)
        Select Case SaveKeisan
            Case 1: KDouble = CisFun.HasuProc(KDouble, 切上げ, 0)
            Case 2: KDouble = CisFun.HasuProc(KDouble, 四捨五入, 0)
            Case Else: KDouble = CisFun.HasuProc(KDouble, 切捨て, 0)
        End Select
        iB1_Kingk = KDouble
    End If
End Sub
'+-----------------------------+
'+  B1Chk_Kari
'+-----------------------------+
Private Sub B1Chk_Kari_Click()
    If B1Chk_Kari.Value = 0 Then
        B1Chk_Kari.Caption = "正式"
    Else
        B1Chk_Kari.Caption = "仮"
    End If
End Sub
'+-----------------------------+
'+  iB1_Tanka
'+-----------------------------+
Private Sub iB1_Tanka_LostFocus()
    If CisFun.Val2(iB1_Suryo) <> 0 And CisFun.Val2(iB1_Tanka) <> 0 Then
        KDouble = CisFun.Val2(iB1_Tanka) * CisFun.Val2(iB1_Suryo)
        Select Case SaveKeisan
            Case 1: KDouble = CisFun.HasuProc(KDouble, 切上げ, 0)
            Case 2: KDouble = CisFun.HasuProc(KDouble, 四捨五入, 0)
            Case Else: KDouble = CisFun.HasuProc(KDouble, 切捨て, 0)
        End Select
        iB1_Kingk = KDouble
    End If
End Sub
'+-----------------------------+
'+  B1Chk_Syogo
'+-----------------------------+
Private Sub B1Chk_Syogo_Click()
    If B1Chk_Syogo.Value = 0 Then
        B1Chk_Syogo.Caption = "対象外"
    Else
        B1Chk_Syogo.Caption = "対象"
    End If
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
'+------------------------------------------------------------------+
'+      基本情報
'+------------------------------------------------------------------+
Private Sub SyoriName(SYKB As String)
    Dim w_Color     As Long
    w_Color = &H80000008 '初期色ｾｯﾄ(黒)
   
'   H2LB_Size_B.Enabled = False
   
    Select Case SyoriKB
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
        If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
        If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
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
    Dim H1Mode      As Boolean
    Dim H1Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    Dim B2Mode      As Boolean
    Dim B2Color     As Long
    
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
    B2Mode = False
    B2Color = gPLostSel

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
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B2_Area1.Enabled = B2Mode
    B2_Area1.BackColor = B2Color
    

    PB_NEXT.Visible = False
    PB_BACK.Visible = False
    PB_CAN.Visible = Not H1Mode
    PB_ADD.Visible = Not H1Mode
    PB_MNT.Visible = Not H1Mode
    PB_DEL.Visible = Not H1Mode
    
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
              '  .Enabled = True
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
    
    
    If HeadBody = "B1" Or HeadBody = "B" Or HeadBody = "ALL" Then
        B1Chk_Aka.Value = 0
        B1Chk_Aka.Caption = "黒"
        B1Chk_Kari.Value = 0
        B1Chk_Kari.Caption = "正式"
    End If
    If HeadBody = "B2" Or HeadBody = "B" Or HeadBody = "ALL" Then
        CisVsGrid1.Clear
        Call CisVsGrid1.Clear
        Dsp_Kensu = ""
    End If
    Exit Sub
HeadBodyClear_ER:
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        If ProcHB = "B2" Then
            Call VSGrid1_Click
        End If
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut(0) Then
            GoTo ReturnPress_Ed
        End If
        If SyoriKB = "ADD" Then
            ProcHB = "B1"
            Call DispChange(ProcHB)
            iB1_KYmd.SetFocus
        Else
            ProcHB = "B2"
            Call DispChange(ProcHB)
            Call VSGrid1_Click
            VsGrid1.SetFocus
        End If
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B2" Then
        If Not Body2Chk Then GoTo ReturnPress_Ed
' << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then
                RowColFlg = True
                Call CisVsGrid1.vsColor(消去個別)
                RowColFlg = False
                GoTo ReturnPress_Ed
            End If
            Call HeadBodyClear("B1")
            If CisFun.Val2(Dsp_Kensu) <> 0 Then
                ProcHB = "B2"
                Call DispChange(ProcHB)
                Call SyoriName(SyoriKB)
                Call VSGrid1_Click
                Call VsGrid1.SetFocus
            Else
                ProcHB = "B1"
                Call DispChange(ProcHB)
                SyoriKB = "ADD"
                Call SyoriName(SyoriKB)
            End If
        Else
            ProcHB = "B1"
            Call DispChange(ProcHB)
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
'
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
' 売上年月
    If Trim(iH1_UYm) = "" Then
        CisFun.MB_MSG(2) = "    売上年月を入力して下さい            "
        CisFun.MBOX
       iH1_UYm.SetFocus
       Exit Function
    End If
    If Not iH1_UYm.cChkResult Then
        CisFun.MB_MSG(2) = "    正しい売上年月を入力して下さい            "
        CisFun.MBOX
       iH1_UYm.SetFocus
       Exit Function
    End If
    If Trim(iH1_UYm) < RTrim(SaveUYm) Then
        CisFun.MB_MSG(2) = "    売上更新年月より前の年月の入力は不可！            "
        CisFun.MBOX
       iH1_UYm.SetFocus
       Exit Function
    End If
' 契約先
    If Trim(iH1_Torcd) = "" Then
        CisFun.MB_MSG(2) = "    契約先を入力して下さい            "
        CisFun.MBOX
       iH1_Torcd.SetFocus
       Exit Function
    End If
    If Not TorNmGet(iH1_Torcd, 2) Then
        CisFun.MB_MSG(2) = "     契約先 未登録！            "
        CisFun.MB_MSG(4) = "    （取引先マスタ）            "
        CisFun.MBOX
       iH1_Torcd.SetFocus
       Exit Function
    End If
    SaveKeisan = TRM.計算単位
'売上実績テーブル
    If Not DBInput("INV") Then
        SyoriKB = "ADD"
        ProcHB = "B1"
    Else
        SyoriKB = "MNT"
        ProcHB = "B2"
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
        .MB_Title = ""
        .MB_Button = Error
    End With
    ' 検収日
    If Trim(iB1_KYmd) = "" Then
        CisFun.MB_MSG(2) = "    検収日を入力して下さい            "
        CisFun.MBOX
       iB1_KYmd.SetFocus
       Exit Function
    End If
    If Not iB1_KYmd.cChkResult Then
        CisFun.MB_MSG(2) = "    正しい検収日を入力して下さい            "
        CisFun.MBOX
       iB1_KYmd.SetFocus
       Exit Function
    End If
    ' 背番号半角チェック
    If RTrim(RTrim(iB1_Seban)) <> "" And Not HalfSizeChk(RTrim(iB1_Seban)) Then
        With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "    背番号は半角にて入力して下さい            "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
        End With
        iB1_Seban.SetFocus
        Exit Function
    End If
    ' 納入先
    If RTrim(iB1_Noucd) <> "" Then
        If Not TorNmGet(iB1_Noucd, 3) Then
            CisFun.MB_MSG(2) = "     納入先 未登録！            "
            CisFun.MB_MSG(4) = "    （取引先マスタ）            "
            CisFun.MBOX
           iB1_Noucd.SetFocus
           Exit Function
        End If
    End If
    ' 金額
    If CisFun.Val2(iB1_Suryo) <> 0 And CisFun.Val2(iB1_Tanka) <> 0 Then
        KDouble = CisFun.Val2(iB1_Tanka) * CisFun.Val2(iB1_Suryo)
        Select Case SaveKeisan
            Case 1: KDouble = CisFun.HasuProc(KDouble, 切上げ, 0)
            Case 2: KDouble = CisFun.HasuProc(KDouble, 四捨五入, 0)
            Case Else: KDouble = CisFun.HasuProc(KDouble, 切捨て, 0)
        End Select

        If CisFun.Val2(iB1_Kingk) <> KDouble Then
            CisFun.MB_MSG(2) = "    金額 <> 数量 × 単価            "
            CisFun.MBOX
            iB1_Kingk.SetFocus
            Exit Function
        End If
    End If
    ' 金額
    If CisFun.Val2(iB1_Kingk) = 0 Then
        CisFun.MB_MSG(2) = "    金額を入力して下さい            "
        CisFun.MBOX
        iB1_Kingk.SetFocus
        Exit Function
    End If
    
    
    Body1Chk = True
End Function
'****************************
'*      Body2ﾁｪｯｸ処理
'****************************
Private Function Body2Chk() As Boolean
    
    Body2Chk = False
        
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
    ' 選択
    If CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(14))) = 0 Then Exit Function
    
    RowColFlg = True
    If SyoriKB = "MNT" Then
        Call CisVsGrid1.vsColor(個別, &HC0FFC0)
    Else
        Call CisVsGrid1.vsColor(個別, &HFFC0FF)
    End If
    RowColFlg = False
    
    
    Body2Chk = True
End Function
'***********************************************************************************
'*
'*       データベース入力 ＆ 表示
'*
'***********************************************************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    Call HeadBodyClear("B")
    
    gSL_Select = "SELECT UR.*,ISNULL(TR.略称,'')納入先名, "
    gSL_Select = gSL_Select & " CASE WHEN UM.検収管理NO IS NULL THEN '対象外' ELSE '対象' END 照合対象 "
    gSL_Select = gSL_Select & " FROM 売上実績テーブル UR"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 売上明細テーブル UM "
    gSL_Select = gSL_Select & "   ON UM.検収管理NO = UR.検収管理NO "
    gSL_Select = gSL_Select & "   AND UM.入力区分 = '1' "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR "
    gSL_Select = gSL_Select & "   ON TR.取引先区分 = 0 "
    gSL_Select = gSL_Select & "   AND TR.取引先CD = UR.納入先 "
    gSL_Select = gSL_Select & " WHERE UR.実績年月 = '" & RTrim(iH1_UYm) & "'"
    gSL_Select = gSL_Select & "   AND UR.契約先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND UR.入力区分 = '1'"
    gSL_Select = gSL_Select & " ORDER BY UR.検収年月日,UR.品番,UR.納入番号"
    
    If Not UJTRead(gSL_Select, 0) Then
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
        Call Grid1Set
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
    
    With CisDB
        On Error GoTo DBDelete_Err
        .DBTran (TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 削除処理 )
        .SQL = "DELETE FROM 売上実績テーブル "
        .SQL = .SQL & " WHERE 検収管理NO = " & CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(14))) & ""
        .DBExec

        If B1Chk_Syogo.Value = 1 Then
            .SQL = "DELETE FROM 売上明細テーブル "
            .SQL = .SQL & " WHERE 検収管理NO = " & CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(14))) & ""
            .DBExec
        End If

        .DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
        
        iT1_Delsu = Format(CisFun.Val2(iT1_Delsu) + 1, "###")
        
        On Error GoTo 0
    End With
    
'Grid再表示
    Call vsGrid1_ReSetDisp
    
    DBDelete = True
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut(KB As Byte)
    DBPut = False
    Dim AddFlg  As Boolean
    Dim SaveKanriNo    As Long

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
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err

'
    If SyoriKB = "ADD" Then
        With CisDB
            .SQL = "番号獲得"
            .StoadoCount = 3
            .ParaNo = 0: .ParaIO = Return用
            .ParaNo = 1: .ParaIO = Input用
            .ParaNo = 2: .ParaIO = OutPut用
            .ParaNo = 3: .ParaIO = Return用

            .ParaNo = 1: .ParaValue = 14

            If Not .DBStored Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "     検収管理NOの採番に失敗しました    "
                    .MB_Title = "採番処理"
                    .MB_Button = Error
                    If Not .MBOX Then GoTo DBPut_Err
                End With
            Else
                .ParaNo = 2: SaveKanriNo = .ParaValue
            End If
        End With
    Else
        SaveKanriNo = CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(14)))
    End If
'
    Call ItemsClearUJT
    With UJT
        .検収管理NO = SaveKanriNo
        .検収管理行NO = 1
        .実績年月 = RTrim(iH1_UYm)
        .品番 = RTrim(iB1_Hinbn)
        .背番号 = RTrim(iB1_Seban)
        .納入先 = RTrim(iB1_Noucd)
        .契約先 = RTrim(iH1_Torcd)
        .受入 = RTrim(iB1_Ukeir)
        .車型 = RTrim(iB1_Syakei)
        .納入番号 = RTrim(iB1_NouNo)
        .検収年月日 = RTrim(iB1_KYmd)
        .実績処理日 = Format(Now(), "yyyymmdd")
        .赤黒区分 = B1Chk_Aka.Value
        .単価 = CisFun.Val2(iB1_Tanka)
        If B1Chk_Kari.Value = 0 Then
            .仮単価区分 = ""
        Else
            .仮単価区分 = "1"
        End If
        If B1Chk_Aka.Value = 0 Then
            .納入数 = CisFun.Val2(iB1_Suryo)
        Else
            .納入数 = CisFun.Val2(iB1_Suryo) * -1
        End If
        If B1Chk_Aka.Value = 0 Then
            .金額 = CisFun.Val2(iB1_Kingk)
        Else
            .金額 = CisFun.Val2(iB1_Kingk) * -1
        End If
        .入力区分 = "1"
    End With

    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
' ( 追加処理 )
    If SyoriKB = "ADD" Then
'        UJT.作成者 = gTanto
        Call UJTInsert_Proc
        iT1_Addsu = Format(CisFun.Val2(iT1_Addsu) + 1, "###")
    Else
' ( 更新処理 )
'        UJT.更新者 = gTanto
        Call UJTUpdate_Proc
        iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
    End If
'
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

'Grid再表示
    Call vsGrid1_ReSetDisp
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
Public Function UJTInsert_Proc() As Boolean
    UJTInsert_Proc = False
'売上実績テーブル出力
    Call UJTInsert
'売上明細テーブル出力
    If B1Chk_Syogo.Value = 1 Then
        Call ItemsClearUMT
        With UMT
            .検収管理NO = UJT.検収管理NO
            .検収管理行NO = UJT.検収管理行NO
            .実績年月 = UJT.実績年月
            .品番 = UJT.品番
            .背番号 = UJT.背番号
            .納入先 = UJT.納入先
            .契約先 = UJT.契約先
            .受入 = UJT.受入
            .車型 = UJT.車型
            .実績処理日 = UJT.実績処理日
            .納入番号 = UJT.納入番号
            .検収年月日 = UJT.検収年月日
            .赤黒区分 = UJT.赤黒区分
            .単価 = UJT.単価
            .仮単価区分 = UJT.仮単価区分
            .納入数 = UJT.納入数
            .金額 = UJT.金額
            .入力区分 = "1"
        End With
        Call UMTInsert
    End If
    UJTInsert_Proc = True
End Function
Public Function UJTUpdate_Proc() As Boolean
    UJTUpdate_Proc = False
'売上実績テーブル更新
    With UJT
        gSL_Select = "Update 売上実績テーブル  Set "
        gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
        gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
        gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
        gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
        gSL_Select = gSL_Select & "車型 = '" & RTrim(.車型) & "',"
        gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
        gSL_Select = gSL_Select & "検収年月日 = '" & RTrim(.検収年月日) & "',"
        gSL_Select = gSL_Select & "実績処理日 = '" & RTrim(.実績処理日) & "',"
        gSL_Select = gSL_Select & "赤黒区分 = " & .赤黒区分 & ","
        gSL_Select = gSL_Select & "単価 = " & .単価 & ","
        gSL_Select = gSL_Select & "仮単価区分 = '" & RTrim(.仮単価区分) & "',"
        gSL_Select = gSL_Select & "納入数 = " & .納入数 & ","
        gSL_Select = gSL_Select & "金額 = " & .金額 & ","
        gSL_Select = gSL_Select & "入力区分 = '" & RTrim(.入力区分) & "'"
        gSL_Select = gSL_Select & " Where 検収管理NO = " & .検収管理NO & ""
    End With
    
    With CisDB
        .ConnectNo = 0
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
'売上明細テーブル更新
    gSL_Select = "SELECT * FROM 売上明細テーブル "
    gSL_Select = gSL_Select & " WHERE 検収管理NO = " & UJT.検収管理NO & ""
    If UMTRead(gSL_Select, 1) Then
        If B1Chk_Syogo.Value = 1 Then
        '変更
            'With UMT-------------------------------------------------------------- 2006/12/13 Delete
            With UJT '------------------------------------------------------------- 2006/12/13 Insert
                gSL_Select = "Update 売上明細テーブル  Set "
                gSL_Select = gSL_Select & "品番 = '" & RTrim(.品番) & "',"
                gSL_Select = gSL_Select & "背番号 = '" & RTrim(.背番号) & "',"
                gSL_Select = gSL_Select & "納入先 = '" & RTrim(.納入先) & "',"
                gSL_Select = gSL_Select & "受入 = '" & RTrim(.受入) & "',"
                gSL_Select = gSL_Select & "車型 = '" & RTrim(.車型) & "',"
                gSL_Select = gSL_Select & "納入番号 = '" & RTrim(.納入番号) & "',"
                gSL_Select = gSL_Select & "検収年月日 = '" & RTrim(.検収年月日) & "',"
                gSL_Select = gSL_Select & "実績処理日 = '" & RTrim(.実績処理日) & "',"
                gSL_Select = gSL_Select & "赤黒区分 = " & .赤黒区分 & ","
                gSL_Select = gSL_Select & "単価 = " & .単価 & ","
                gSL_Select = gSL_Select & "仮単価区分 = '" & RTrim(.仮単価区分) & "',"
                gSL_Select = gSL_Select & "納入数 = " & .納入数 & ","
                gSL_Select = gSL_Select & "金額 = " & .金額 & ","
                gSL_Select = gSL_Select & "入力区分 = '" & RTrim(.入力区分) & "'"
                gSL_Select = gSL_Select & " Where 検収管理NO = " & UJT.検収管理NO & ""
            End With
            
            With CisDB
                .ConnectNo = 0
                .SQL = gSL_Select
                If Not .DBExec Then Exit Function
            End With
        Else
        '削除
            gSL_Select = "DELETE FROM 売上明細テーブル "
            gSL_Select = gSL_Select & " WHERE 検収管理NO = " & UJT.検収管理NO & ""
            With CisDB
                .ConnectNo = 0
                .SQL = gSL_Select
                If Not .DBExec Then Exit Function
            End With
        End If
    Else
        If B1Chk_Syogo.Value = 1 Then
        '追加
            Call ItemsClearUMT
            With UMT
                .検収管理NO = UJT.検収管理NO
                .検収管理行NO = UJT.検収管理行NO
                .実績年月 = UJT.実績年月
                .品番 = UJT.品番
                .背番号 = UJT.背番号
                .納入先 = UJT.納入先
                .契約先 = UJT.契約先
                .受入 = UJT.受入
                .車型 = UJT.車型
                .実績処理日 = UJT.実績処理日
                .納入番号 = UJT.納入番号
                .検収年月日 = UJT.検収年月日
                .赤黒区分 = UJT.赤黒区分
                .単価 = UJT.単価
                .仮単価区分 = UJT.仮単価区分
                .納入数 = UJT.納入数
                .金額 = UJT.金額
                .入力区分 = "1"
            End With
            Call UMTInsert
        End If
    End If
    
    UJTUpdate_Proc = True
End Function
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        ReDim SColZ(21)

        Call .FixedSet(中中, 右中, あり, 3, 0, "No")
        Call .FixedSet(中中, 中中, あり, 10, 1, "検 収 日")
        Call .FixedSet(中中, 中中, あり, 9, 2, "納入番号")
        Call .FixedSet(中中, 左中, あり, 25, 3, "品番")
        Call .FixedSet(中中, 中中, あり, 4, 4, "赤黒")
        Call .FixedSet(中中, 右中, あり, 7, 5, "数量")
        Call .FixedSet(中中, 中中, あり, 4, 6, "仮区")
        Call .FixedSet(中中, 右中, あり, 9, 7, "単 価")
        Call .FixedSet(中中, 右中, あり, 11, 8, "金 額")
        
        '*-*-*-*-* 2008/04/11 Update Start
        'Call .FixedSet(中中, 中中, あり, 6, 9, "背番号")
        Call .FixedSet(中中, 中中, あり, 8, 9, "背番号")
        'Call .FixedSet(中中, 中中, あり, 4, 10, "車型")
        Call .FixedSet(中中, 中中, あり, 5, 10, "車型")
        '*-*-*-*-* 2008/04/11 Update End
        
        Call .FixedSet(中中, 左中, あり, 28, 11, "納入先")
        Call .FixedSet(中中, 中中, あり, 4, 12, "受入")
        Call .FixedSet(中中, 中中, あり, 6, 13, "照合")
        
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 中中, あり, 0, 14, "検収管理NO")
        Call .FixedSet(中中, 中中, あり, 0, 15, "X検収日")
        Call .FixedSet(中中, 中中, あり, 0, 16, "X納入先")
        Call .FixedSet(中中, 中中, あり, 0, 17, "X仮単価区分")
        Call .FixedSet(中中, 左中, あり, 0, 18, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp

    End With

End Sub
'****************************************************************************************'
'*      グリッド内容セット
'****************************************************************************************'
Private Sub Grid1Set()
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
    Dsp_Kensu = Format(Ttl_Kensu, "#,###")
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid1
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid1.FixedGet(ID)
        Next ID
        If Ttl_Kensu + 1 < CisVsGrid1.DispMax Then
            .Rows = CisVsGrid1.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    
    ID = 0
    
    Do Until Not UJT_RDSTS
        With VsGrid1
            ID = ID + 1
            '№
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            ' 検収日
            If Trim(UJT.検収年月日) <> "" Then
               With CisFun
                    .DateE_BefVal = UJT.検収年月日
                    .DateEdit
                    gStr = .DateE_AfVal
               End With
               .TextMatrix(ID, GridCol(1)) = gStr
            End If
            '納入番号
            .TextMatrix(ID, GridCol(2)) = RTrim(UJT.納入番号)
            '品番
            .TextMatrix(ID, GridCol(3)) = RTrim(UJT.品番)
            '赤黒
            If UJT.金額 < 0 Then
                .TextMatrix(ID, GridCol(4)) = "赤"
            Else
                .TextMatrix(ID, GridCol(4)) = "黒"
            End If
            ' 数量
            If UJT.納入数 <> 0 Then
                If UJT.納入数 < 0 Then
                    .TextMatrix(ID, GridCol(5)) = Format(UJT.納入数 * -1, "#,###")
                Else
                    .TextMatrix(ID, GridCol(5)) = Format(UJT.納入数, "#,###")
                End If
            End If
            '仮区分
            If RTrim(UJT.仮単価区分) = "" Then
                .TextMatrix(ID, GridCol(6)) = "正式"
            Else
                .TextMatrix(ID, GridCol(6)) = "仮"
            End If
            ' 単価
            If UJT.単価 <> 0 Then
                .TextMatrix(ID, GridCol(7)) = Format(UJT.単価, "#,##0.00")
            End If
            ' 金額
            If UJT.金額 <> 0 Then
                If UJT.金額 < 0 Then
                    .TextMatrix(ID, GridCol(8)) = Format(UJT.金額 * -1, "#,###")
                Else
                    .TextMatrix(ID, GridCol(8)) = Format(UJT.金額, "#,###")
                End If
            End If
            '背番号
            .TextMatrix(ID, GridCol(9)) = RTrim(UJT.背番号)
            '車型
            .TextMatrix(ID, GridCol(10)) = RTrim(UJT.車型)
            '納入先・名
            If RTrim(UJT.納入先) <> "" Then
                .TextMatrix(ID, GridCol(11)) = RTrim(UJT.納入先) & ":" & RTrim(UJT.納入先名)
            End If
            '受入
            .TextMatrix(ID, GridCol(12)) = RTrim(UJT.受入)
            '照合対象
            .TextMatrix(ID, GridCol(13)) = RTrim(UJT.照合対象)
            '検収管理NO
            .TextMatrix(ID, GridCol(14)) = UJT.検収管理NO
            '検収日
            .TextMatrix(ID, GridCol(15)) = RTrim(UJT.検収年月日)
            '納入先
            .TextMatrix(ID, GridCol(16)) = RTrim(UJT.納入先)
            '仮単価区分
            .TextMatrix(ID, GridCol(17)) = RTrim(UJT.仮単価区分)
            
        End With
        
        Call UJTReadNext
    Loop
    
    Call UJTClose
    
    With VsGrid1
        .Row = 1: .Col = 1: .ColSel = .Cols - 1
        .Redraw = True
    End With
End Sub
'+---------------------------------------------------------------------------------------+
'+  グリッドより入力域へセット
'+---------------------------------------------------------------------------------------+
Private Sub GridToText()
    
    With VsGrid1
        If Trim(.TextMatrix(.Row, CisVsGrid1.FixedGet(14))) = "" Then Exit Sub
        
        iB1_KYmd = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(15)))
        iB1_NouNo = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(2)))
        iB1_Hinbn = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(3)))
        iB1_Seban = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(9)))
        iB1_Syakei = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(10)))
        iB1_Noucd = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(16)))
        Call iB1_Noucd_LostFocus
        iB1_Ukeir = RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(12)))
        If RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(4))) = "黒" Then
            B1Chk_Aka.Value = 0
        Else
            B1Chk_Aka.Value = 1
        End If
        iB1_Suryo = CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(5)))
        
        If RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(17))) = "" Then
            B1Chk_Kari.Value = 0
        Else
            B1Chk_Kari.Value = 1
        End If
        iB1_Tanka = CisFun.RSetFld(.TextMatrix(.Row, CisVsGrid1.FixedGet(7)), 12, "#,##0.00")
        iB1_Kingk = CisFun.Val2(.TextMatrix(.Row, CisVsGrid1.FixedGet(8)))
        If RTrim(.TextMatrix(.Row, CisVsGrid1.FixedGet(13))) = "対象外" Then
            B1Chk_Syogo.Value = 0
        Else
            B1Chk_Syogo.Value = 1
        End If
    End With

End Sub
'+---------------------------------------+
'+  グリッド (Click) サイズ情報
'+---------------------------------------+
Private Sub VSGrid1_Click()
    With VsGrid1
        If ProcHB = "B2" Then
            Call HeadBodyClear("B1")

            Call GridToText
        End If
    End With
End Sub
'+---------------------------------------+
'+  グリッド (DblClick)
'+---------------------------------------+
Private Sub vsGrid1_DblClick()
    
    With VsGrid1
        If ProcHB = "B2" Then
            If Trim(.TextMatrix(.Row, CisVsGrid1.FixedGet(14))) <> "" Then
               Call ReturnPress
            End If
        End If
    End With
End Sub
'+---------------------------------------+
'+  グリッド (GotFocus)
'+---------------------------------------+
Private Sub vsGrid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace And Trim(VsGrid1.TextMatrix(VsGrid1.Row, CisVsGrid1.FixedGet(14))) <> "" Then
       Call ReturnPress
    End If
End Sub
'+---------------------------------------+
'+  グリッド (RowColChange)
'+---------------------------------------+
Private Sub vsGrid1_RowColChange()
    If Not FormAct Then Exit Sub
    If ProcHB <> "B2" Then Exit Sub
    If RowColFlg Then Exit Sub
    Call VSGrid1_Click
End Sub
'+---------------------------------------+
'+  グリッド 再表示
'+---------------------------------------+
Private Sub vsGrid1_ReSetDisp()
    
    SaveRow = VsGrid1.Row
    RowColFlg = True
    Call DBInput("INV")
    RowColFlg = False
    If SaveRow > VsGrid1.Rows - 1 Then
        SaveRow = SaveRow - 1
    End If
    If SyoriKB <> "ADD" Then
        VsGrid1.Row = SaveRow
    End If
    If VsGrid1.RowIsVisible(SaveRow) = False Then
        VsGrid1.TopRow = SaveRow
    End If

End Sub





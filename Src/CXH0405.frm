VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CXH0405 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "自社伝票　ファイル再出力"
   ClientHeight    =   10980
   ClientLeft      =   1320
   ClientTop       =   2325
   ClientWidth     =   15315
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   10980
   ScaleWidth      =   15315
   WindowState     =   2  '最大化
   Begin VB.PictureBox H1_Area2 
      Height          =   780
      Left            =   750
      ScaleHeight     =   720
      ScaleWidth      =   5415
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   690
      Width           =   5475
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   645
         Left            =   60
         Top             =   45
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   1138
         ForeColor       =   16711680
         Caption         =   "自給チケットファイル出力先"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
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
         Begin Cis3D_v60.CIS3D PB_JDenFDir 
            Height          =   315
            Left            =   4875
            Top             =   255
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
         Begin Cis3D_v60.CIS3D H1lb_JDenFDir 
            Height          =   315
            Left            =   75
            Top             =   240
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   556
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
            cPositionX      =   40
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      Height          =   1305
      Left            =   750
      ScaleHeight     =   1245
      ScaleWidth      =   13575
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1560
      Width           =   13635
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   1155
         Left            =   6555
         Top             =   60
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   2037
         ForeColor       =   16711680
         Caption         =   "仕 入 先"
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
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1125
            Top             =   480
            Width           =   2475
            _ExtentX        =   4366
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
            cPositionX      =   40
         End
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   150
            TabIndex        =   12
            Top             =   480
            Width           =   975
            _ExtentX        =   1720
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
      Begin Cis3D_v60.CIS3D UC_3D9 
         Height          =   1155
         Left            =   4695
         Top             =   60
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   2037
         ForeColor       =   16711680
         Caption         =   "     納 番   頁"
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
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_SNouNo 
            Height          =   360
            Left            =   570
            TabIndex        =   8
            Top             =   330
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   635
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
            Text            =   "XXXXX"
            MaxLength       =   5
         End
         Begin CisText_V60.CisText iH1_ENouNo 
            Height          =   360
            Left            =   570
            TabIndex        =   10
            Top             =   690
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
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   360
            Left            =   135
            Top             =   330
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   635
            Caption         =   "開始"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   360
            Left            =   135
            Top             =   690
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   635
            Caption         =   "終了"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
            cPositionX      =   30
         End
         Begin CisText_V60.CisText iH1_SPage 
            Height          =   360
            Left            =   1335
            TabIndex        =   9
            Top             =   330
            Width           =   375
            _ExtentX        =   661
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
         Begin CisText_V60.CisText iH1_EPage 
            Height          =   360
            Left            =   1335
            TabIndex        =   11
            Top             =   690
            Width           =   375
            _ExtentX        =   661
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
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   1155
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   2037
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin VB.OptionButton H1Op_SKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "再発行既存"
            ForeColor       =   &H00000000&
            Height          =   360
            Index           =   2
            Left            =   60
            MaskColor       =   &H00000000&
            TabIndex        =   2
            Top             =   690
            Width           =   1545
         End
         Begin VB.OptionButton H1Op_SKbn 
            BackColor       =   &H00C0FFC0&
            Caption         =   "再発行最新"
            ForeColor       =   &H00000000&
            Height          =   360
            Index           =   1
            Left            =   60
            MaskColor       =   &H00000000&
            TabIndex        =   1
            Top             =   330
            Value           =   -1  'True
            Width           =   1545
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   1155
         Index           =   1
         Left            =   10335
         Top             =   60
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2037
         ForeColor       =   16711680
         Caption         =   "発注区分"
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
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   30
         Begin VB.ComboBox H1Cmb_HKbn 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   225
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   13
            Top             =   465
            Width           =   1170
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   1155
         Index           =   2
         Left            =   11910
         Top             =   60
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   2037
         ForeColor       =   16711680
         Caption         =   "手配区分"
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
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   30
         Begin VB.ComboBox H1Cmb_TKbn 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   195
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   14
            Top             =   450
            Width           =   1185
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   1155
         Left            =   1755
         Top             =   60
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   2037
         ForeColor       =   16711680
         Caption         =   "        納 入 日       - 便"
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
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   30
         Begin VB.CheckBox H1Chk_Ymd 
            BackColor       =   &H00C0C0C0&
            Caption         =   "納　入　日　　便"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   480
            TabIndex        =   3
            Top             =   15
            Value           =   1  'ﾁｪｯｸ
            Width           =   2250
         End
         Begin Cis3D_v60.CIS3D PB_SYmd 
            Height          =   360
            Left            =   90
            Top             =   330
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   635
            Caption         =   "開始"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D PB_EYmd 
            Height          =   360
            Left            =   90
            Top             =   690
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   635
            Caption         =   "終了"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
            cPositionX      =   30
         End
         Begin CisText_V60.CisText iH1_SBin 
            Height          =   360
            Left            =   2385
            TabIndex        =   5
            Top             =   330
            Width           =   375
            _ExtentX        =   661
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
         Begin CisText_V60.CisText iH1_EBin 
            Height          =   360
            Left            =   2385
            TabIndex        =   7
            Top             =   690
            Width           =   375
            _ExtentX        =   661
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_SYmd 
            Height          =   360
            Left            =   525
            TabIndex        =   4
            Top             =   330
            Width           =   1845
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
            Object.Width           =   1845
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_EYmd 
            Height          =   360
            Left            =   525
            TabIndex        =   6
            Top             =   690
            Width           =   1845
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
            Object.Width           =   1845
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
   End
   Begin Cis3D_v60.CIS3D CIS3D1 
      Height          =   375
      Left            =   9585
      Top             =   0
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   661
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
         Height          =   345
         Left            =   1920
         Top             =   15
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   609
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B1lb_Kensu 
            Height          =   270
            Left            =   975
            Top             =   30
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   476
            BackColor       =   16777152
            Caption         =   "#,###"
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
      End
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   345
         Left            =   60
         Top             =   15
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   609
         Caption         =   "選択件数"
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
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B1lb_SKensu 
            Height          =   270
            Left            =   975
            Top             =   30
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   476
            BackColor       =   12648384
            Caption         =   "#,###"
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
      Height          =   720
      Left            =   9570
      ScaleHeight     =   660
      ScaleWidth      =   4260
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   675
      Width           =   4320
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1725
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
         Left            =   2565
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
         Left            =   3405
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
         Left            =   45
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
      Begin CisBtn_60.CisBtn PB_Prt 
         Height          =   585
         Left            =   885
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
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   7065
      Left            =   750
      ScaleHeight     =   7005
      ScaleWidth      =   13590
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3150
      Width           =   13650
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   6900
         Left            =   60
         Top             =   60
         Width           =   13485
         _ExtentX        =   23786
         _ExtentY        =   12171
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
            Height          =   6690
            Left            =   90
            TabIndex        =   17
            Top             =   105
            Width           =   13290
            _Version        =   196608
            _ExtentX        =   23442
            _ExtentY        =   11800
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
            FormatString    =   "№  |作成ファイル名 　          |納 入 日 　-便 |仕入先 |   仕  入  先  名    |受入|納 番 |頁|発注区|手配区|SK|"
            Rows            =   21
            Cols            =   17
            BackColor       =   12648447
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   320
            GridLines       =   2
            SelectionMode   =   1
            BackColorAlternate=   12648447
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
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   -1320
      Width           =   585
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
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 自給チケット　ファイル再出力  】"
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
      cPositionX      =   1600
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   16711935
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
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
      End
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
         Left            =   7065
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
End
Attribute VB_Name = "CXH0405"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   自給チケット　ファイル再出力
'**       フォームID    :   CXH0405
'**       処理概要      :
'**
'**       作  成  日    :   2004/02/25  By CIS
'**       変  更  日    :   2005/12/13   By CIS - INIファイル-クライアント対応（TEXT出力先）
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl

    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim CisVsGrid1              As New CisVsGrid3
    Dim CisCombo1               As CisCombo
    Dim CisCombo2               As CisCombo
'   出力区分の背景色
    Dim mColor_On               As Long
    Dim mColor_Off              As Long

    Dim wPrintKB                As Integer      ' 帳票区分
    Dim wYmdName                As String       ' 日付名
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
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
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

'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 320
        .PatanMax = 1
        .SelectCol = 14
        Set .SelectDispObj = B1lb_SKensu
        .InitGet = False
        .Init
    End With
    Call GridInit1

    '+---------------------+
    '+ ｺﾝﾎﾞﾎﾞｯｸｽ 初期設定  +
    '+---------------------+
'   発注区分
    Set CisCombo1 = New CisCombo
    With CisCombo1
        Set .Connect = CisDB
        Set .CombObj = H1Cmb_HKbn       'ｺﾝﾎﾞﾎﾞｯｸｽｵﾌﾞｼﾞｪｸﾄ名 ｾｯﾄ
        .Control = "名称マスタ"
        .ComboName = "発注区分"
        .SetSpace = 6           '空白文字数
        .Widht = 0              'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ幅(省略時、変更なし)
        .Ess = False            '空白行(True:必須入力用,False:任意入力用)
        .IniPath = gIniExe
        If .ComboInit Then
            Call .ComboSet
        End If
    End With
'   手配区分
    Set CisCombo2 = New CisCombo
    With CisCombo2
        Set .Connect = CisDB
        Set .CombObj = H1Cmb_TKbn       'ｺﾝﾎﾞﾎﾞｯｸｽｵﾌﾞｼﾞｪｸﾄ名 ｾｯﾄ
        .Control = "名称マスタ"
        .ComboName = "手配区分"
        .SetSpace = 6           '空白文字数
        .Widht = 0              'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ幅(省略時、変更なし)
        .Ess = False            '空白行(True:必須入力用,False:任意入力用)
        .IniPath = gIniExe
        If .ComboInit Then
            Call .ComboSet
        End If
    End With

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

'   背景色の設定（出力区分）
    mColor_On = H1Op_SKbn(1).BackColor
    mColor_Off = H1Op_SKbn(2).BackColor
    H1Op_SKbn(1).Value = True

    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

'--( TEXT出力先　獲得 )-- [クライアント情報]
    With CisFun
'-----------------------------------------------< 2005/12/13 ADD START >
        .INI_FileID = gIniExe & gwsIniName
'-----------------------------------------------< 2005/12/13 ADD START >
        .INI_Section = "Option"
        .INI_Key = "JTICKETFILEOUT"
        .INI_Default = gIniExe
        .GetIni
        H1lb_JDenFDir = RTrim(.INI_String)
'-----------------------------------------------< 2005/12/13 ADD START >
        .INI_FileID = gIniExe & gIniName
'-----------------------------------------------< 2005/12/13 ADD START >
    End With

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
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click      '【検索】
           Case vbKeyF8:        If PB_Prt.Visible Then Call PB_Prt_Click        '【出力】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click        '【取消】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click        '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                          '【入力】
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
    Call CisDB.DBDISConnect
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
Private Sub PB_CAN_Click()
    Call CisFrm.MousePT(11)
    If ProcHB = "B1" Then
       ProcHB = "H1"
       Call DispChange(ProcHB)
       Call HeadBodyClear("B")
       GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    出　力  キ ー(F8)       +
'+----------------------------+
Private Sub PB_Prt_Click()
    If ProcHB = "B1" Then
        If Body1Chk Then
            Call DBPut

        '--( TEXT出力先　更新 )-- [クライアント情報]
            With CisFun
'-----------------------------------------------< 2005/12/13 ADD START >
                .INI_FileID = gIniExe & gwsIniName
'-----------------------------------------------< 2005/12/13 ADD START >
                .INI_Section = "Option"
                .INI_Key = "JTICKETFILEOUT"
                .INI_String = H1lb_JDenFDir
                .PutIni
'-----------------------------------------------< 2005/12/13 ADD START >
                .INI_FileID = gIniExe & gIniName
'-----------------------------------------------< 2005/12/13 ADD START >
            End With

            wPrintKB = 1                ' 帳票区分：0=18 Else 1=118
            Call CXH0400TFile(wPrintKB)

            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call HeadBodyClear("B")
        End If
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
'   日付
    If PB_Look.Tag = "iH1_SYmd" Then iH1_SYmd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_EYmd" Then iH1_EYmd.ShowCalender: Exit Sub
End Sub
'+----------------------------+
'+      取引先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 9            ' [発注]手配先/契約先/社内検索

    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+-----------------------------+
'+  PB_JDenFDir
'+-----------------------------+
Private Sub PB_JDenFDir_Click()
    With CisFun
        .Folder_Title = "自給チケット　ファイル出力先参照"
        .Folder_DefDir = H1lb_JDenFDir
        .FolderSel

        If RTrim$(.Folder_Result) <> "" Then H1lb_JDenFDir = RTrim$(.Folder_Result)
    End With
End Sub
'+-----------------------------+
'+  H1Op_SKbn
'+-----------------------------+
Private Sub H1Op_SKbn_Click(Index As Integer)
    If Index = 1 Then
        H1Op_SKbn(2).BackColor = mColor_Off
        H1Op_SKbn(1).BackColor = mColor_On
    Else
        H1Op_SKbn(1).BackColor = mColor_Off
        H1Op_SKbn(2).BackColor = mColor_On
    End If
End Sub
'+-----------------------------+
'+  H1Chk_Ymd
'+-----------------------------+
Private Sub H1Chk_Ymd_Click()
    If H1Chk_Ymd = 1 Then
        H1Chk_Ymd.Caption = "作　成　日"
        iH1_SBin = ""
        iH1_EBin = ""
        iH1_SBin.Enabled = False
        iH1_EBin.Enabled = False
    Else
        H1Chk_Ymd.Caption = "納　入　日　　便"
        iH1_SBin.Enabled = True
        iH1_EBin.Enabled = True
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
    If TorNmGet(iH1_Torcd, 9) Then
       H1lb_Tornm = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  H1Cmb_HKbn
'+-----------------------------+
Private Sub H1Cmb_HKbn_GotFocus()
    H1Cmb_HKbn.BackColor = gIGotSel
End Sub
Private Sub H1Cmb_HKbn_LostFocus()
    H1Cmb_HKbn.BackColor = gILostSel
End Sub
'+-----------------------------+
'+  H1Cmb_TKbn
'+-----------------------------+
Private Sub H1Cmb_TKbn_GotFocus()
    H1Cmb_TKbn.BackColor = gIGotSel
End Sub
Private Sub H1Cmb_TKbn_LostFocus()
    H1Cmb_TKbn.BackColor = gILostSel
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

    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1": H1Mode = True: H1Color = gPGotSel
       Case "B1": B1Mode = True: B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode: H1_Area1.BackColor = H1Color
    H1_Area2.Enabled = H1Mode: H1_Area2.BackColor = H1Color
    B1_Area1.Enabled = B1Mode: B1_Area1.BackColor = B1Color

    PB_Look.Visible = False
    PB_Prt.Visible = Not H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    PB_ENT.Visible = H1Mode

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
            If .Name Like wbc_Enm & "Cmb_*" Then .ListIndex = 0
        End With
    Next gObj
    If HeadBody = "B" Then
        CisVsGrid1.SelectCount = 0
        Call CisVsGrid1.Clear
    End If
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
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False

    If H1Chk_Ymd = 1 Then
        wYmdName = "作成日"
    Else
        wYmdName = "納入日"
    End If
'   納入日＆便／作成日
    If Trim(iH1_SYmd) <> "" Then
        If CisFun.Val2(iH1_SYmd.cDay) = 0 Then
            iH1_EYmd = ""
            iH1_EBin = ""
        Else
            If Not iH1_SYmd.cChkResult Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    " & wYmdName & "（開始）を正しく入力して下さい            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                iH1_SYmd.SetFocus
                Exit Function
            End If
        End If
    Else
        If Trim(iH1_SBin) <> "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    納入日（開始）を入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_SYmd.SetFocus
            Exit Function
        End If
    End If
    If Trim(iH1_EYmd) <> "" Then
        If Not iH1_EYmd.cChkResult Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    " & wYmdName & "（終了）を正しく入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_EYmd.SetFocus
            Exit Function
        End If
    Else
        If Trim(iH1_EBin) <> "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    納入日（終了）を入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_EYmd.SetFocus
            Exit Function
        End If
    End If
    If Trim(iH1_SYmd) <> "" And Trim(iH1_EYmd) <> "" Then
       If iH1_SYmd > iH1_EYmd Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    " & wYmdName & "を正しく入力して下さい（開始＞終了）            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iH1_SYmd.SetFocus
          Exit Function
       End If
       If iH1_SYmd = iH1_EYmd And _
          iH1_SBin > iH1_EBin Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    納入日・便を正しく入力して下さい（開始＞終了）            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
          iH1_SYmd.SetFocus
          Exit Function
       End If
    End If
'   納番＆頁
    If InStr(1, iH1_SNouNo, "*") > 0 Then
        iH1_ENouNo = ""
    End If
    If Trim(iH1_SNouNo) = "" Then
        If Trim(iH1_SPage) <> "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    納番（開始）を入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_SNouNo.SetFocus
            Exit Function
        End If
    End If
    If Trim(iH1_ENouNo) = "" Then
        If Trim(iH1_EPage) <> "" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    納番（終了）を入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_ENouNo.SetFocus
            Exit Function
        End If
    End If
    If Trim(iH1_SNouNo) <> "" And Trim(iH1_ENouNo) <> "" Then
        If iH1_SNouNo > iH1_ENouNo Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    納番を正しく入力して下さい（開始＞終了）            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_SNouNo.SetFocus
          Exit Function
        End If
        If Trim(iH1_SPage) <> "" And Trim(iH1_EPage) <> "" Then
            If iH1_SNouNo = iH1_ENouNo And _
               CisFun.Val2(iH1_SPage) > CisFun.Val2(iH1_EPage) Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    納番・頁を正しく入力して下さい（開始＞終了）            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                iH1_SNouNo.SetFocus
                Exit Function
            End If
        End If
    End If
'   仕入先
    If Trim(iH1_Torcd) <> "" Then
        If Not TorNmGet(iH1_Torcd, 9) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    仕入先　未登録！（取引先マスタ）            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Torcd.SetFocus
           Exit Function
        End If
    End If

    Call SqlSelectSet
    If Not DBInput("INV") Then
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
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False

    If CisFun.Val2(B1lb_SKensu) = 0 Then
        With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    出力データを選択して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
        End With
        Exit Function
    End If

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

    If GridSet1 Then
        DBInput = True
    End If

DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'****************************
'*      SQL文ｾｯﾄ処理
'****************************
Private Sub SqlSelectSet(Optional InputKey As String = "")
    Dim YmdSta      As String
    Dim YmdEnd      As String
    Dim BinSta      As String
    Dim BinEnd      As String
'
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  hch.*"
    gSL_Select = gSL_Select & ", isnull(trm.略称, '') 仕入先名"
    gSL_Select = gSL_Select & ", isnull(sy1.値名称, '') 発注区分名"
    gSL_Select = gSL_Select & ", isnull(sy2.値名称, '') 手配区分名"
    If H1Op_SKbn(1) Then
'   ### 再発行最新 ###
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  帳票出力テーブル cht"
        gSL_Select = gSL_Select & ", 発注帳票発行テーブル hch"
    Else
'   ### 再発行既存 ###
        gSL_Select = gSL_Select & " from"
        gSL_Select = gSL_Select & "  発注帳票発行テーブル hch"
    End If
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ trm"
    gSL_Select = gSL_Select & "    on trm.取引先CD = hch.仕入先"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 1"
    gSL_Select = gSL_Select & "  left outer join 名称マスタ sy1"
    gSL_Select = gSL_Select & "    on sy1.区分名称 = '発注区分'"
    gSL_Select = gSL_Select & "   and sy1.区分タイプ = ''"
    gSL_Select = gSL_Select & "   and sy1.値 = hch.発注区分"
    gSL_Select = gSL_Select & "  left outer join 名称マスタ sy2"
    gSL_Select = gSL_Select & "    on sy2.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "   and sy2.区分タイプ = ''"
    gSL_Select = gSL_Select & "   and sy2.値 = hch.手配区分"
    If H1Op_SKbn(1) Then
'   ### 再発行最新 ###
        gSL_Select = gSL_Select & " where cht.端末番号 = substring(host_name(),1,20)"
        gSL_Select = gSL_Select & "   and cht.帳票区分 = 18"
        gSL_Select = gSL_Select & "   and hch.納品書管理NO = cht.管理NO"
        gSL_Select = gSL_Select & "   and hch.納入番号 = cht.管理文字"
        gSL_Select = gSL_Select & "   and hch.頁NO = cht.補助番号"
        gSL_Select = gSL_Select & "   and hch.カード = cht.検索情報"
        gSL_Select = gSL_Select & "   and hch.発行種類 = 3"
        gSL_Select = gSL_Select & "   and hch.発行形態 = 2"
        gSL_Select = gSL_Select & "   and hch.TEXTNAME <> ''"
    Else
'   ### 再発行既存 ###
        gSL_Select = gSL_Select & " where hch.発行種類 = 3"
        gSL_Select = gSL_Select & "   and hch.発行形態 = 2"
        gSL_Select = gSL_Select & "   and hch.TEXTNAME <> ''"
    End If
'   >> 納入番号
    If InStr(1, iH1_SNouNo, "*") > 0 Then
        gSL_Select = gSL_Select & "   and hch.納入番号 like '" & RTrim(Replace(iH1_SNouNo, "*", "%")) & "'"
    Else
        If Trim(iH1_SNouNo) <> "" Then
            If Trim(iH1_ENouNo) = "" Then
                '> 開始のみ
                If CisFun.Val2(iH1_SPage) = 0 Then
                    gSL_Select = gSL_Select & "   and hch.納入番号 = '" & RTrim(iH1_SNouNo) & "'"
                Else
                    gSL_Select = gSL_Select & "   and"
                    gSL_Select = gSL_Select & "   hch.納入番号 + space(5 - len(hch.納入番号))"
                    gSL_Select = gSL_Select & " + right('0' + cast(cast(hch.頁NO as decimal(2)) as varchar(2)), 2)"
                    gSL_Select = gSL_Select & " = "
                    gSL_Select = gSL_Select & "'"
                    gSL_Select = gSL_Select & iH1_SNouNo & Space(5 - CisFun.Len2(iH1_SNouNo))
                    gSL_Select = gSL_Select & Format(iH1_SPage, "00")
                    gSL_Select = gSL_Select & "'"
                End If
            Else
                '> 開始
                If CisFun.Val2(iH1_SPage) = 0 Then
                    gSL_Select = gSL_Select & "   and hch.納入番号 >= '" & RTrim(iH1_SNouNo) & "'"
                Else
                    gSL_Select = gSL_Select & "   and"
                    gSL_Select = gSL_Select & "   hch.納入番号 + space(5 - len(hch.納入番号))"
                    gSL_Select = gSL_Select & " + right('0' + cast(cast(hch.頁NO as decimal(2)) as varchar(2)), 2)"
                    gSL_Select = gSL_Select & " >= "
                    gSL_Select = gSL_Select & "'"
                    gSL_Select = gSL_Select & iH1_SNouNo & Space(5 - CisFun.Len2(iH1_SNouNo))
                    gSL_Select = gSL_Select & Format(iH1_SPage, "00")
                    gSL_Select = gSL_Select & "'"
                End If
            End If
        End If
        '> 終了
        If Trim(iH1_ENouNo) <> "" Then
            If CisFun.Val2(iH1_EPage) = 0 Then
                gSL_Select = gSL_Select & "   and hch.納入番号 <= '" & RTrim(iH1_ENouNo) & "'"
            Else
                gSL_Select = gSL_Select & "   and"
                gSL_Select = gSL_Select & "   hch.納入番号 + space(5 - len(hch.納入番号))"
                gSL_Select = gSL_Select & " + right('0' + cast(cast(hch.頁NO as decimal(2)) as varchar(2)), 2)"
                gSL_Select = gSL_Select & " <= "
                gSL_Select = gSL_Select & "'"
                gSL_Select = gSL_Select & iH1_ENouNo & Space(5 - CisFun.Len2(iH1_ENouNo))
                gSL_Select = gSL_Select & Format(iH1_EPage, "00")
                gSL_Select = gSL_Select & "'"
            End If
        End If
    End If
'   >> 仕入先
    If RTrim(iH1_Torcd) <> "" Then
        gSL_Select = gSL_Select & "   and hch.仕入先 = '" & RTrim(iH1_Torcd) & "'"
    End If
'   >> 納入日／作成日
    If Trim(iH1_SYmd) <> "" And CisFun.Val2(iH1_SYmd.cDay) = 0 Then
    '   > 年月指定
        gSL_Select = gSL_Select & "   and substring(hch." & wYmdName & ", 1, 6) = '" & iH1_SYmd & "'"
    Else
        YmdSta = "": BinSta = ""
        YmdEnd = "": BinEnd = ""
        '> 日付＆便　編集
        If Trim(iH1_SYmd) <> "" Then YmdSta = iH1_SYmd
        If Trim(iH1_EYmd) <> "" Then YmdEnd = iH1_EYmd
        If CisFun.Val2(iH1_SBin) <> 0 Then BinSta = Format(iH1_SBin, "00")
        If CisFun.Val2(iH1_EBin) <> 0 Then BinEnd = Format(iH1_EBin, "00")
        '> 開始日のみ
        If Trim(YmdSta) <> "" And Trim(YmdEnd) = "" Then
            If Trim(BinSta) = "" Then
                gSL_Select = gSL_Select & "   and hch." & wYmdName & " = '" & YmdSta & "'"
            Else
                gSL_Select = gSL_Select & "   and hch.納入日 + hch.納入便 = '" & YmdSta & BinSta & "'"
            End If
        End If
        '> 終了日のみ
        If Trim(YmdSta) = "" And Trim(YmdEnd) <> "" Then
            If Trim(BinEnd) = "" Then
                gSL_Select = gSL_Select & "   and hch." & wYmdName & " <= '" & YmdEnd & "'"
            Else
                gSL_Select = gSL_Select & "   and hch.納入日 + hch.納入便 <= '" & YmdEnd & BinEnd & "'"
            End If
        End If
        '> 開始日＆終了日
        If Trim(YmdSta) <> "" And Trim(YmdEnd) <> "" Then
            If Trim(BinSta) = "" And Trim(BinEnd) = "" Then
                gSL_Select = gSL_Select & "   and hch." & wYmdName & " between '" & YmdSta & "'" _
                                                                       & " and '" & YmdEnd & "'"
            Else
                If Trim(BinSta) = "" Then BinSta = "00"
                If Trim(BinEnd) = "" Then BinEnd = "99"
                gSL_Select = gSL_Select & "   and hch.納入日 + hch.納入便 between '" & YmdSta & BinSta & "'" _
                                                                         & " and '" & YmdEnd & BinEnd & "'"
            End If
        End If
    End If
'   >> 発注区分
    If RTrim(H1Cmb_HKbn) <> "" Then
        Call CisCombo1.CodeGet
        gSL_Select = gSL_Select & "   and hch.発注区分 = " & CisCombo1.Code & ""
    End If
'   >> 手配区分
    If RTrim(H1Cmb_TKbn) <> "" Then
        Call CisCombo2.CodeGet
        gSL_Select = gSL_Select & "   and hch.手配区分 = " & CisCombo2.Code & ""
    End If
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  hch.TEXTNAME"
    gSL_Select = gSL_Select & ", hch.カード"
    gSL_Select = gSL_Select & ", hch.納入日"
    gSL_Select = gSL_Select & ", hch.納入便"
    gSL_Select = gSL_Select & ", hch.仕入先"
    gSL_Select = gSL_Select & ", hch.受入"
    gSL_Select = gSL_Select & ", hch.納入番号"
    gSL_Select = gSL_Select & ", hch.頁NO"
'
End Sub
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

    On Error GoTo DBPut_Err

    Call CisDB.DBTran(TransBegin)   ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 削除処理 )
    Call DBDelete_CHT
' ( 作成処理 )
    Dim ID          As Long
    With vsGrid1
        For ID = 1 To .Rows - 1
            If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(0))) = "" Then Exit For
            If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(14))) = "X" Then
                Call ItemsClearCHT
                CHT.端末番号 = "' + substring(host_name(),1,20) + '"
                CHT.帳票区分 = 118
                CHT.管理文字 = RTrim(.TextMatrix(ID, CisVsGrid1.FixedGet(6)))
                CHT.管理NO = CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(12)))
                CHT.補助番号 = CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(7)))
                Call CHTInsert
            End If
        Next ID
    End With

    CisDB.DBTran (TransCommit)      ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'*--------------------*'
'*     削除処理
'*--------------------*'
Public Function DBDelete_CHT(Optional DBNo As Integer) As Boolean
    DBDelete_CHT = False
    With CHT
        gSL_Select = ""
        gSL_Select = gSL_Select & "delete 帳票出力テーブル"
        gSL_Select = gSL_Select & " where 端末番号 = substring(host_name(),1,20)"
        gSL_Select = gSL_Select & "   and 帳票区分 = 118"
    End With
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBDelete_CHT = True
End Function



'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ - 流動日付
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 4, 0, "№")
        Call .FixedSet(中中, 左中, あり, 25, 1, "作成ファイル名")
        Call .FixedSet(中中, 左中, あり, 13, 2, "納 入 日  -便")
        Call .FixedSet(中中, 左中, あり, 7, 3, "仕入先")
        Call .FixedSet(中中, 左中, あり, 20, 4, "仕　入　先　名")
        Call .FixedSet(中中, 中中, あり, 4, 5, "受入")
        Call .FixedSet(中中, 中中, あり, 6, 6, "納 番")
        Call .FixedSet(中中, 右中, あり, 2, 7, "頁")
        Call .FixedSet(中中, 中中, あり, 6, 8, "発注区")
        Call .FixedSet(中中, 中中, あり, 6, 9, "手配区")
        Call .FixedSet(中中, 中中, あり, 2, 10, "SK")
        Call .FixedSet(中中, 左中, あり, 25, 11, "X品番")
        Call .FixedSet(中中, 左中, あり, 0, 12, "納品書管理NO")
        Call .FixedSet(中中, 左中, あり, 0, 13, "カード")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 14, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10")
        .InitDisp
    End With
End Sub
'***********************************************'
'*****      グリッド内容セット - 流動日付
'***********************************************'
Private Function GridSet1() As Boolean
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

    GridSet1 = False

    If HCHRead(gSL_Select, , 1) Then

'       抽出処理 初期ｾｯﾄ
        B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
        Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
'       ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
        With vsGrid1
           .Redraw = False
            If Ttl_Kensu + 1 < CisVsGrid1.DispMax Then
                .Rows = CisVsGrid1.DispMax
            Else
                .Rows = Ttl_Kensu + 1
            End If
        End With
        ID = 0
        Do Until Not HCH_RDSTS
            ID = ID + 1

            Call HchInfoToGrid1(ID)

            Call HCHReadNext(1)
        Loop
        With vsGrid1
             .Row = 1: .Col = 1: .ColSel = .Cols - 1
             .Redraw = True
        End With

        GridSet1 = True
    End If
    Call HCHClose
End Function
'+----------------------------------------------+
'+      発注テーブルよりグリッドに転送
'+----------------------------------------------+
Private Sub HchInfoToGrid1(SetRow As Long)
    With vsGrid1
        'No
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(0)) = SetRow
        '作成ﾌｧｲﾙ名
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(1)) = HCH.TEXTNAME
        '納入日＆便
        If Trim(HCH.納入日) <> "" Then
            If RTrim(HCH.納入便) = "" Then
                .TextMatrix(SetRow, CisVsGrid1.FixedGet(2)) = Format(RTrim(HCH.納入日), "0000/00/00")
            Else
                .TextMatrix(SetRow, CisVsGrid1.FixedGet(2)) = Format(RTrim(HCH.納入日), "0000/00/00") & "-" & RTrim(HCH.納入便)
            End If
        End If
        '仕入先
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(3)) = RTrim(HCH.仕入先)
        '仕入先名
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(4)) = RTrim(HCH.仕入先名)
        '受入
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(5)) = RTrim(HCH.受入)
        '納番
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(6)) = RTrim(HCH.納入番号)
        '頁
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(7)) = Format(HCH.頁NO, "##")
        '発注
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(8)) = RTrim(HCH.発注区分名)
        '手配
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(9)) = RTrim(HCH.手配区分名)
        'SK
        Select Case HCH.SK区分
            Case 0: .TextMatrix(SetRow, CisVsGrid1.FixedGet(10)) = "K"
            Case 1: .TextMatrix(SetRow, CisVsGrid1.FixedGet(10)) = "S"
        End Select
        '納品書管理NO
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(12)) = HCH.納品書管理NO
        'カード
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(13)) = HCH.カード
    End With
End Sub
'******************************************'
'*****      行選択（クリック）
'******************************************'
Private Sub vsGrid1_Click()
    If ProcHB = "B1" Then
        Call CisVsGrid1.vsColor
    End If
End Sub
'******************************************'
'*****      行選択（スペース）
'******************************************'
Private Sub vsGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If ProcHB = "B1" Then
        If KeyCode = vbKeySpace Then
            Call CisVsGrid1.vsColor
        End If
    End If
End Sub




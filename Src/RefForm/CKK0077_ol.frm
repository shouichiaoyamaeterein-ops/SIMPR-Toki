VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form CKK0077 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "材料取引先マスタ検索"
   ClientHeight    =   9525
   ClientLeft      =   420
   ClientTop       =   1605
   ClientWidth     =   14520
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
   ScaleHeight     =   9525
   ScaleWidth      =   14520
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   1035
      Left            =   105
      ScaleHeight     =   975
      ScaleWidth      =   13380
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   780
      Width           =   13440
      Begin Cis3D_v60.CIS3D UC_3D 
         Height          =   855
         Index           =   0
         Left            =   45
         Top             =   60
         Width           =   13290
         _ExtentX        =   23442
         _ExtentY        =   1508
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
         cFont3DColor2   =   8421504
         cBoderStyle     =   1
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   795
            Left            =   8280
            Top             =   30
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "受"
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
            Begin CisText_V60.CisText iH1_Ukeir 
               Height          =   375
               Left            =   60
               TabIndex        =   3
               Top             =   330
               Width           =   375
               _ExtentX        =   661
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
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   795
            Index           =   3
            Left            =   10110
            Top             =   30
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   1402
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
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            Begin VB.CheckBox H1Ck_SK 
               BackColor       =   &H00C0C0C0&
               Caption         =   "かんばん"
               Height          =   285
               Index           =   0
               Left            =   360
               TabIndex        =   6
               Top             =   90
               Width           =   1305
            End
            Begin VB.CheckBox H1Ck_SK 
               BackColor       =   &H00C0C0C0&
               Caption         =   "指示"
               Height          =   285
               Index           =   1
               Left            =   360
               TabIndex        =   7
               Top             =   390
               Width           =   825
            End
            Begin Cis3D_v60.CIS3D CIS3D3 
               Height          =   585
               Left            =   90
               Top             =   120
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   1032
               ForeColor       =   16711680
               Caption         =   "ＳＫ"
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
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   795
            Index           =   5
            Left            =   1320
            Top             =   30
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "材　　　質"
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
            cPositionY      =   30
            Begin CisText_V60.CisText iH1_Zai 
               Height          =   375
               Left            =   105
               TabIndex        =   1
               Top             =   330
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
               Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   25
               IMEMode         =   2
               cDataReplace    =   1
               cFaZero         =   0
               cFbComma        =   0
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D2 
            Height          =   795
            Left            =   4695
            Top             =   30
            Width           =   3585
            _ExtentX        =   6324
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "取 引 先"
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
               Height          =   360
               Left            =   60
               TabIndex        =   2
               Top             =   330
               Width           =   990
               _ExtentX        =   1746
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
               Text            =   "XXXXXXX"
               MaxLength       =   7
            End
            Begin Cis3D_v60.CIS3D H1lb_Tornm 
               Height          =   360
               Left            =   1035
               Top             =   315
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
               cPositionX      =   30
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   795
            Index           =   2
            Left            =   30
            Top             =   30
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   1402
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
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   30
            Begin CisText_V60.CisText iH1_Seban 
               Height          =   375
               Left            =   90
               TabIndex        =   0
               Top             =   330
               Width           =   1095
               _ExtentX        =   1931
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
               Text            =   "XXXXXXXX"
               MaxLength       =   8
               IMEMode         =   2
               cDataReplace    =   1
               cFaZero         =   0
               cFbComma        =   0
            End
         End
         Begin Cis3D_v60.CIS3D Back_Hachu 
            Height          =   795
            Index           =   3
            Left            =   11820
            Top             =   30
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "手配区分"
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
            Begin VB.ComboBox H1Comb_Tehai 
               BackColor       =   &H00C0E0FF&
               Height          =   360
               Left            =   120
               Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
               TabIndex        =   8
               Top             =   360
               Width           =   1245
            End
            Begin Cis3D_v60.CIS3D B3lb_Tehai 
               Height          =   330
               Left            =   120
               Top             =   360
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   582
               BackColor       =   16777152
               Caption         =   "社内"
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
               cAlingnment     =   7
               cBoderStyle     =   1
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   795
            Index           =   1
            Left            =   8790
            Top             =   30
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   1402
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
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            Begin VB.CheckBox H1Ck_Hinmoku 
               BackColor       =   &H00C0C0C0&
               Caption         =   "受注"
               Height          =   285
               Index           =   0
               Left            =   360
               TabIndex        =   4
               Top             =   90
               Width           =   825
            End
            Begin VB.CheckBox H1Ck_Hinmoku 
               BackColor       =   &H00C0C0C0&
               Caption         =   "発注"
               Height          =   285
               Index           =   1
               Left            =   360
               TabIndex        =   5
               Top             =   450
               Width           =   825
            End
            Begin Cis3D_v60.CIS3D CIS3D1 
               Height          =   585
               Left            =   90
               Top             =   150
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   1032
               ForeColor       =   16711680
               Caption         =   "品目"
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
      End
   End
   Begin Cis3D_v60.CIS3D HLB_Title 
      Height          =   675
      Left            =   120
      Top             =   30
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   1191
      BackColor       =   12648384
      Caption         =   "材料取引先マスタ検索"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      cAlingnment     =   7
   End
   Begin VB.TextBox Dummy 
      Height          =   270
      IMEMode         =   3  'ｵﾌ固定
      Left            =   180
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   150
      Width           =   855
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   10800
      ScaleHeight     =   615
      ScaleWidth      =   3510
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   30
      Width           =   3570
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   1785
         Top             =   0
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
         Left            =   2640
         Top             =   0
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
         BtnCaption      =   "戻る"
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   930
         Top             =   0
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   60
         Top             =   0
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
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7515
      Left            =   120
      ScaleHeight     =   7455
      ScaleWidth      =   14265
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1860
      Width           =   14325
      Begin Cis3D_v60.CIS3D Disp_Msg 
         Height          =   1005
         Left            =   3540
         Top             =   1110
         Visible         =   0   'False
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   1773
         BackColor       =   4210752
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
         Begin Cis3D_v60.CIS3D Msg 
            Height          =   885
            Left            =   90
            Top             =   60
            Width           =   7755
            _ExtentX        =   13679
            _ExtentY        =   1561
            BackColor       =   12648384
            ForeColor       =   8388608
            Caption         =   "抽出中　しばらくお待ち下さい"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   21.75
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
      Begin TrueOleDBGrid60.TDBGrid DBGrid1 
         Bindings        =   "CKK0077.frx":0000
         Height          =   7320
         Left            =   60
         OleObjectBlob   =   "CKK0077.frx":0015
         TabIndex        =   9
         Top             =   60
         Width           =   14130
      End
   End
   Begin VB.PictureBox Back_Sort 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   6900
      ScaleHeight     =   570
      ScaleWidth      =   1710
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   60
      Width           =   1770
      Begin VB.OptionButton Op_Sort 
         BackColor       =   &H00808000&
         Caption         =   "降順"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   870
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   180
         Width           =   675
      End
      Begin VB.OptionButton Op_Sort 
         BackColor       =   &H00808000&
         Caption         =   "昇順"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   0
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   105
         Value           =   -1  'True
         Width           =   705
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   360
      Left            =   105
      Top             =   150
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "CKK0077"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   材料取引先マスタ検索
'**       フォームID    :   CKK0077
'**       処理概要      :
'**
'**       作  成  日    :   2004/02/18  By CIS
'**       変  更  日    :   2008/03/01  Bt CIS 材料取引先対応
'**       変  更  日    :   2009/12/08  Bt CIS 置場表示変更
'**
'**********************************************************'

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim mSL_Order       As String       ' ソート順(Order BY)
    Dim mSL_OrderA      As String       ' ソート順(Order BY)昇順
    Dim mSL_OrderD      As String       ' ソート順(Order BY)降順
    Dim mSL_SelectW     As String       ' SQL SELECT(Where)条件文
    Dim FormCap         As String
    
    Dim CisFrm          As CISFormContorl
    Dim CisComboTehai   As CisCombo
    Dim DBSpCols        As Columns      ' Columnsｵﾌﾞｼﾞｪｸﾄﾜｰｸ1
    Dim DBSpCol         As Column       ' Columnｵﾌﾞｼﾞｪｸﾄﾜｰｸ1
    Dim mColMoveFlg     As Boolean
    Dim SortCnt         As Long
    Dim mOrder()        As String
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub form_load()
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Set CisFrm = New CISFormContorl
    
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
'
    Dummy.Left = -1000
    Dummy.Enabled = False
'    Timer1.Interval = 1000   ' 時刻表示間隔
'    Timer1.Enabled = True    ' 時刻表示開始
    '+---------------------+
    '+ ｺﾝﾎﾞﾎﾞｯｸｽ 初期設定  +
    '+---------------------+
    With CisFun
        .INI_Section = "発注処理手配"
        .INI_FileID = gIniExe & "Combo.ini"
        .INI_Default = ""
        .INI_Key = "TYPE"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_Key = "TYPE"
           .INI_String = "SQL"
           .PutIni
           .INI_Key = "CNT"
           .INI_String = "4"
           .PutIni
           .INI_Key = "DATA1"
           .INI_String = "SELECT 値,値名称 FROM 名称マスタ"
           .PutIni
           .INI_Key = "DATA2"
           .INI_String = "WHERE 区分名称 = '手配区分'   AND 区分タイプ = ''"
           .PutIni
           .INI_Key = "DATA3"
           .INI_String = "And 値 IN('0','1','3','4')"
           .PutIni
           .INI_Key = "DATA4"
           .INI_String = "ORDER BY 値"
           .PutIni
        End If
    End With
    ' 手配区分
    Set CisComboTehai = New CisCombo
    With CisComboTehai
        Set .Connect = CisDB
        Set .CombObj = H1Comb_Tehai
        .Control = "発注処理手配"
        .ComboName = ""
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .Ess = False
        .ComboInit
        .ComboSet
    End With
'    ' 手配区分
'    Set CisComboTehai = New CisCombo
'    With CisComboTehai
'        Set .Connect = CisDB
'        Set .CombObj = H1Comb_Tehai
'        .Control = ""
'        .ComboName = "手配区分"
'        .Ess = True
'        .IniPath = gIniExe
'        .SetSpace = 50
'        .Ess = False
'        .ComboInit
'
'        .CombowkSet (2)
'        .Code = 0: .CombText = "購入"
'        .CombowkSet
'        .Code = 1: .CombText = "直送"
'        .CombowkSet
'        .ComboSet
'    End With
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    If RV_Left <> 0 Or RV_Top <> 0 Then
        Me.Move RV_Left, RV_Top
    Else
        Call CisFrm.Frm_Center(False, x1024y768)
    End If
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    FormCap = Me.Caption
    Op_Sort(0).Value = True
' ADOｺﾝﾄﾛｰﾙ　接続情報　セット
    Adodc1.ConnectionString = CisDB.ConnectString
    Call HeadBodyClear("H")
    RV_Rtn = False
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/03/01 INSERT START
    If RVI_Hinmoku = 9 Then
    Else
       H1Ck_Hinmoku(RVI_Hinmoku).Value = 1
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/03/01 INSERT END
    
    If RVI_SK = 9 Then
    Else
       H1Ck_SK(RVI_SK).Value = 1
    End If
    If Trim(RVI_Torcd) <> "" Then
       iH1_Torcd = Trim(RVI_Torcd)
        Call iH1_Torcd_LostFocus
    End If
    If Trim(RVI_Ukeir) <> "" Then
       iH1_Ukeir = Trim(RVI_Ukeir)
    End If
    If Trim(RVI_Seban) <> "" Then
       iH1_Seban = Trim(RVI_Seban)
    End If
    If Trim(RVI_Tehai) <> "" Then
       CisComboTehai.Code = Trim(RVI_Tehai)
       CisComboTehai.TextGet
    End If
    
    If RV_Call <> "" Then
       Call ReturnPress
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
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    Select Case Key_Code
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click   '【検索】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_Ent.Visible Then                      '【入力】
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
    Call PB_END_Click
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    Unload CKK0077
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
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B1" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        PB_Ent.BtnCaption = "開始"
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_UTorcd: Exit Sub
End Sub
'+----------------------------+
'+       仕入先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 9
    RV_Call = ""
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+-----------------------------+
'+       取引先受入検索         +
'+-----------------------------+
Private Sub Look_UTorcd()
    RV_Left = 0
    RV_Top = 0
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/03/01 UPDATE START
'''    RV_TorKb = 9
    If H1Ck_Hinmoku(0).Value <> 0 Then
        RV_TorKb = 3
    Else
        RV_TorKb = 9
    End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/03/01 UPDATE END
    RV_Torcd = iH1_Torcd
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
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
'+  iH1_Ukeir
'+-----------------------------+
Private Sub iH1_Ukeir_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Ukeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  H1Comb_Tehai
'+-----------------------------+
Private Sub H1Comb_Tehai_GotFocus()
    H1Comb_Tehai.BackColor = gIGotSel
End Sub
Private Sub H1Comb_Tehai_LostFocus()
    H1Comb_Tehai.BackColor = gILostSel
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
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
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
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj
    If HeadBody = "B" Then
       DBGrid1.Close True
       Adodc1.Recordset.Close
       Adodc1.RecordSource = ""
       Me.Caption = FormCap
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
       If Not BodySet("") Then GoTo ReturnPress_Ed
       ProcHB = "B1"
       PB_Ent.BtnCaption = "選択"
       Call DispChange(ProcHB)
       GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then Call RtnSet
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
'    If iH1_???? = "" Then
'       ReDim gMBMsg(5)
'       gMBMsg(2) = "  ???を入力して下さい  "
'       gMBMsg(3) = "   "
'       gMBTitle = " "
'       gMBButton = "E"
'       gRtn = MBox
'       iH1_???.SetFocus
'       Exit Function
'    End If
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
'    If iB1_???? = "" Then
'       ReDim gMBMsg(5)
'       gMBMsg(2) = "  ???を入力して下さい  "
'       gMBMsg(3) = "   "
'       gMBTitle = " "
'       gMBButton = "E"
'       gRtn = MBox
'       iB1_???.SetFocus
'       Exit Function
'    End If
    Body1Chk = True
End Function
'****************************
'*      分類切替
'****************************
Private Sub Op_Sort_Click(Index As Integer)
    If ProcHB = "H1" Then Exit Sub
    If Adodc1.Recordset.RecordCount <> 0 Then Call BodySet("S")
    SendKeys "{TAB}"
End Sub
'****************************
'*      ボディ内容セット
'****************************
Private Function BodySet(SetType As String) As Boolean
    
    BodySet = False
    
    Disp_Msg.Visible = True
    If SetType = "" Then
        Msg = "抽出中　しばらくお待ち下さい"
    Else
        Msg = "分類中　しばらくお待ち下さい"
    End If
    If FormAct Then
       DoEvents
    End If
    
    gSL_Select = "Select ZT.背番号,ZT.取引先,ZM.材質,ZM.板厚,ZM.幅,ZM.長さ,ZM.寸法区分,"
    gSL_Select = gSL_Select & "    CASE WHEN ZM.寸法区分 = 0 Then Cast(ZM.板厚 as varchar(7)) Else '' End 表示板厚,"
    gSL_Select = gSL_Select & "    CASE WHEN ZM.寸法区分 = 0 Then Cast(ZM.幅 as varchar(7)) Else '' End 表示幅,"
    gSL_Select = gSL_Select & "    CASE WHEN ZM.寸法区分 = 0 Then Cast(ZM.長さ as varchar(7)) Else '' End 表示長さ,"
    gSL_Select = gSL_Select & "    ISNULL(ZT.取引先,'') 取引先,ISNULL(ZT.受入,'') 受入,"
    gSL_Select = gSL_Select & "    CASE WHEN ISNULL(TR.略称,'') <> '' Then ISNULL(ZT.取引先,'') + '：' + TR.略称 Else ISNULL(ZT.取引先,'') End 表示取引先,"
    gSL_Select = gSL_Select & "    CASE WHEN ISNULL(TR2.略称,'') <> '' Then ISNULL(ZT.契約先,'') + '：' + TR2.略称 Else ISNULL(ZT.契約先,'') End 表示契約先,"
    gSL_Select = gSL_Select & "    CASE WHEN ISNULL(ZT.手配区分,'') = '1' THEN"
    gSL_Select = gSL_Select & "    CASE WHEN ISNULL(TR3.略称,'') <> '' Then ISNULL(ZT.ﾒｰｶｰ,'') + '：' + TR3.略称 Else ISNULL(ZT.ﾒｰｶｰ,'') End"
    gSL_Select = gSL_Select & "    Else '' END  表示納入先,"
    gSL_Select = gSL_Select & "    ISNULL(ZT.背番号,'') 背番号,' ' + ISNULL(ZT.背番号,'') 表示背番号,"
    gSL_Select = gSL_Select & "    ISNULL(ZT.収容数1,0) 収容数1,ISNULL(ZT.SK区分,0) SK区分,"
    gSL_Select = gSL_Select & "    CASE WHEN ZT.SK区分 IS NUll THEN ''"
    gSL_Select = gSL_Select & "    Else"
    gSL_Select = gSL_Select & "    Case When  ISNULL(ZT.SK区分,0) = 0 Then 'K'"
    gSL_Select = gSL_Select & "    Else 'S' End End SK,"
    gSL_Select = gSL_Select & "    CASE When ISNULL(ZT.サイクル,'') <> '' Then"
    gSL_Select = gSL_Select & "    Substring(ZT.サイクル,1,1) + '-' + Substring(ZT.サイクル,2,2) + '-' + Substring(ZT.サイクル,4,2)"
    gSL_Select = gSL_Select & "    Else '' End 表示サイクル,"
    gSL_Select = gSL_Select & "    ISNull(ZT.手配区分,'') 手配区分,RTRIM(ISNULL(SY.値名称,'')) 手配区分名,"
    gSL_Select = gSL_Select & "    ISNULL(ZT.収容器1,'') 収容器1,"
    gSL_Select = gSL_Select & "    CASE WHEN ISNULL(ZT.変更日,'') <> '' THEN"
    gSL_Select = gSL_Select & "        SubString(ZT.変更日,1,4) + '/' + SubString(ZT.変更日,5,2) + '/' + SubString(ZT.変更日,7,2) + '-' + Cast(ZT.変更便 as char(2))"
    gSL_Select = gSL_Select & "        Else '' END 変更日便,"
    gSL_Select = gSL_Select & "    ISNULL(ZT.収容器2,'') 収容器2,ISNULL(ZT.収容数2,0) 収容数2,"
    gSL_Select = gSL_Select & "    ISNULL(ZT.ﾒｰｶｰ受入,'') ﾒｰｶｰ受入,"
    gSL_Select = gSL_Select & "    CASE WHEN ISNULL(ZT.置場,'') <> '' THEN "                       '2009/12/08
    gSL_Select = gSL_Select & "         SUBSTRING(ISNULL(ZT.置場,''),1,2) + '-' + SUBSTRING(ISNULL(ZT.置場,''),3,2) + '-' + SUBSTRING(ISNULL(ZT.置場,''),5,2) + '-' + SUBSTRING(ISNULL(ZT.置場,''),7,2)  "      '2009/12/08
    gSL_Select = gSL_Select & "    ELSE '' END 置場 ,"
    gSL_Select = gSL_Select & "    ISNULL(ZT.アドレス,'') アドレス,ISNULL(ZM.材料名,'') 材料名"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/03/01 INSERT START
    gSL_Select = gSL_Select & "   ,ISNULL(ZT.受給品番, '') 受給品番, ISNULL(ZT.発注品番, '') 発注品番, ZT.材料管理番号"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/03/01 INSERT END

    gSL_Select = gSL_Select & "     From 材料取引先マスタ ZT"
    gSL_Select = gSL_Select & "     Left Outer join 材料マスタ ZM"
    gSL_Select = gSL_Select & "        On ZM.材料管理番号 = ZT.材料管理番号"
    gSL_Select = gSL_Select & "     Left Outer Join 取引先マスタ TR"
    gSL_Select = gSL_Select & "        ON  TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "        AND TR.取引先CD = ISNULL(ZT.取引先,'')"
    gSL_Select = gSL_Select & "    Left Outer Join 取引先マスタ TR2"
    gSL_Select = gSL_Select & "    ON  TR2.取引先区分 = 1"
    gSL_Select = gSL_Select & "    AND TR2.取引先CD = ISNULL(ZT.契約先,'')"
    gSL_Select = gSL_Select & "    Left Outer Join 取引先マスタ TR3"
    gSL_Select = gSL_Select & "    ON  TR3.取引先区分 = 1"
    gSL_Select = gSL_Select & "    AND TR3.取引先CD = ISNULL(ZT.ﾒｰｶｰ,'')"
    gSL_Select = gSL_Select & "    Left Outer Join 名称マスタ SY"
    gSL_Select = gSL_Select & "    ON SY.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "    And Sy.区分タイプ = ''"
    gSL_Select = gSL_Select & "    And SY.値 = ISNULL(ZT.手配区分,'')"
    
    
    mSL_OrderA = " Order By ISNULL(ZT.取引先,''),ZM.材質,ZM.板厚,ZM.幅,ZM.長さ "
    mSL_OrderD = " Order By ISNULL(ZT.取引先,'') Desc,ZM.材質 Desc,ZM.板厚 Desc,ZM.幅 Desc,ZM.長さ Desc "
' 選択条件セット
    mSL_SelectW = ""
    With CisFun
        .Where_CreateKB = NewDocument
        .Where_Create AlphaMe, "ISNULL(ZT.背番号,'')", iH1_Seban
        .Where_Create AlphaMe, "ISNULL(ZM.材質,'')", iH1_Zai
        .Where_Create AlphaMe, "ISNULL(ZT.取引先,'')", iH1_Torcd
        .Where_Create AlphaMe, "ISNULL(ZT.受入,'')", iH1_Ukeir
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/03/01 INSERT START
'        If Not (H1Ck_Hinmoku(0).Value = H1Ck_Hinmoku(1).Value) Then
'           If H1Ck_Hinmoku(0).Value = 1 Then
'              .Where_Create Numeric, "ISNULL(ZT.品目,9)", "0", 等しい, Create
'           End If
'           If H1Ck_Hinmoku(1).Value = 1 Then
'              .Where_Create Numeric, "ISNULL(ZT.品目,9)", "1"
'           End If
'        End If
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/03/01 INSERT END
        If Not (H1Ck_SK(0).Value = H1Ck_SK(1).Value) Then
           If H1Ck_SK(0).Value = 1 Then
              .Where_Create Numeric, "ISNULL(ZT.SK区分,9)", "0", 等しい, Create
           End If
           If H1Ck_SK(1).Value = 1 Then
              .Where_Create Numeric, "ISNULL(ZT.SK区分,9)", "1"
           End If
        End If
        CisComboTehai.CodeGet
        If Trim(CisComboTehai.Code) <> "" Then
           .Where_Create Numeric, "ISNULL(ZT.手配区分,9)", Trim(CisComboTehai.Code), 等しい, Create
        End If
                
        mSL_SelectW = .Where_Phrase
    End With
    
    Adodc1.RecordSource = gSL_Select & mSL_SelectW & SortGet
    
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount = 0 Then
        Disp_Msg.Visible = False
        DoEvents
        With CisFun
            .MB_Lines = 3
            .MB_MSG(1) = "    該当するデータはありません  "
            .MB_Button = OK
            .MBOX
        End With
        Exit Function
    End If
    Call TotalDisp
    Disp_Msg.Visible = False
    BodySet = True
End Function
'+---------------------------+
'+  分類順を求める
'+---------------------------+
Private Function SortGet() As String
    mSL_Order = ""
    gStr = "": SortCnt = 0
    ReDim mOrder(5)
    Set DBSpCols = DBGrid1.Splits(0).Columns
    For Each DBSpCol In DBSpCols
         With DBSpCol
             If .Visible Then
                If .Order < 5 Then
                   mOrder(.Order) = .DataField
                   SortCnt = SortCnt + 1
                End If
                If SortCnt = 5 Then Exit For
             End If
         End With
    Next DBSpCol
    
    For gInt = 0 To 4
        mSL_Order = mSL_Order & gStr & mOrder(gInt)
        If Op_Sort(1).Value Then
           mSL_Order = mSL_Order & " Desc"
        End If
        gStr = ","
    Next
    If mSL_Order <> "" Then
       SortGet = " Order By " & Replace(mSL_Order, "材質", "ZM.材質")
    Else
       SortGet = ""
    End If
End Function
'+----------------------------------------------------------+
'+                                                          +
'+                                                          +
'+      グリッド処理(TDBGridPro)                             +
'+                                                          +
'+                                                          +
'+----------------------------------------------------------+
'+----------------------------------------------------------+
'+      ADODC
'+----------------------------------------------------------+
Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Adodc1.RecordSource <> "" Then Call TotalDisp
End Sub
'+----------------------------------------------------------+
'+      ADODC1(件数表示)                                     +
'+----------------------------------------------------------+
Private Sub TotalDisp()
    Static NowCnt       As Long
    Static MaxCnt       As Long
    NowCnt = Adodc1.Recordset.AbsolutePosition
    MaxCnt = Adodc1.Recordset.RecordCount
    Me.Caption = FormCap & Space(5)
    Me.Caption = Me.Caption & "抽出件数 = " & Format(NowCnt, "##,##0")
    Me.Caption = Me.Caption & "/" & Format(MaxCnt, "##,##0") & " 件"
End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseUp  イベント   》                +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DBGrid1.RowContaining(Y) < 0 Then Exit Sub
    If DBGrid1.ColContaining(X) < 0 Then Exit Sub
    SendKeys "{RIGHT}"
    DBGrid1.Row = DBGrid1.RowContaining(Y)
    DoEvents
    Call RtnSet
End Sub
'+----------------------------------------------------------+
'+ 表示項目移動後 分類表示 《 ColMove イベント   》            +
'+----------------------------------------------------------+
Private Sub DBGrid1_ColMove(ByVal Position As Integer, Cancel As Integer)
    mColMoveFlg = True
End Sub
'+----------------------------------------------------------+
'+ 表示項目移動後 分類表示 《 Paint イベント   》              +
'+----------------------------------------------------------+
Private Sub DBGrid1_Paint()
    If Not FormAct Then Exit Sub
    If mColMoveFlg Then  ' 再表示(分類)
        mColMoveFlg = False
        Call BodySet("S")
        SendKeys "{RIGHT}"
    End If
End Sub
'+-------------------------------------+
'+      選択内容ｾｯﾄ & 復帰ｽﾃｰﾀｽｾｯﾄ      +
'+-------------------------------------+
Private Sub RtnSet()
    With Adodc1
        RV_Zaist = .Recordset("材質")
        RV_Seban = .Recordset("背番号")
        RV_TorcdK = .Recordset("取引先")
        RV_Ukeir = .Recordset("受入")
        RV_Itatu = .Recordset("板厚")
        RV_Width = .Recordset("幅")
        RV_Long = .Recordset("長さ")
        RV_Size = .Recordset("寸法区分")
        RV_Tehai = .Recordset("手配区分")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/03/01 INSERT START
        RV_JHinbn = .Recordset("受給品番")
        RV_HHinbn = .Recordset("発注品番")
        Rv_KanriNo = .Recordset("材料管理番号")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/03/01 INSERT END
    End With
    Adodc1.Recordset.Close
    RV_Rtn = True
    Call PB_END_Click
End Sub

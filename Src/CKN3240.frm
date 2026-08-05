VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form CKN3240 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "品番取引先マスタ検索"
   ClientHeight    =   9525
   ClientLeft      =   465
   ClientTop       =   1260
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   9525
   ScaleWidth      =   14520
   Begin VB.PictureBox Back_Dsp 
      BackColor       =   &H00404040&
      Height          =   720
      Left            =   9870
      ScaleHeight     =   660
      ScaleWidth      =   4305
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1740
      Width           =   4365
      Begin Cis3D_v60.CIS3D UC_3D5 
         Height          =   645
         Left            =   15
         Top             =   0
         Width           =   4275
         _ExtentX        =   7541
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
               TabIndex        =   24
               Top             =   150
               Value           =   -1  'True
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
               TabIndex        =   23
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
               TabIndex        =   22
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
               TabIndex        =   21
               Top             =   150
               Width           =   615
            End
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
               TabIndex        =   20
               Top             =   150
               Width           =   615
            End
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   1830
      Left            =   105
      ScaleHeight     =   1770
      ScaleWidth      =   14220
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   780
      Width           =   14280
      Begin Cis3D_v60.CIS3D UC_3D 
         Height          =   1650
         Index           =   0
         Left            =   45
         Top             =   60
         Width           =   14100
         _ExtentX        =   24871
         _ExtentY        =   2910
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
         Begin Cis3D_v60.CIS3D CIS3D7 
            Height          =   795
            Left            =   3630
            Top             =   825
            Width           =   10445
            _ExtentX        =   18415
            _ExtentY        =   1402
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
         End
         Begin Cis3D_v60.CIS3D CIS3D5 
            Height          =   795
            Left            =   8220
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
            Left            =   10980
            Top             =   30
            Width           =   1650
            _ExtentX        =   2910
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
               TabIndex        =   7
               Top             =   90
               Width           =   1275
            End
            Begin VB.CheckBox H1Ck_SK 
               BackColor       =   &H00C0C0C0&
               Caption         =   "指示"
               Height          =   285
               Index           =   1
               Left            =   360
               TabIndex        =   8
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
            Index           =   1
            Left            =   9780
            Top             =   30
            Width           =   1200
            _ExtentX        =   2117
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
            Begin VB.CheckBox H1Ck_Hinmoku 
               BackColor       =   &H00C0C0C0&
               Caption         =   "発注"
               Height          =   285
               Index           =   1
               Left            =   360
               TabIndex        =   6
               Top             =   450
               Width           =   795
            End
            Begin VB.CheckBox H1Ck_Hinmoku 
               BackColor       =   &H00C0C0C0&
               Caption         =   "受注"
               Height          =   285
               Index           =   0
               Left            =   360
               TabIndex        =   5
               Top             =   90
               Width           =   795
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   795
            Index           =   0
            Left            =   8730
            Top             =   30
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "社内背番"
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
            Begin CisText_V60.CisText iH1_SSeban 
               Height          =   375
               Left            =   180
               TabIndex        =   4
               Top             =   330
               Width           =   705
               _ExtentX        =   1244
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
               Text            =   "XXXXX"
               MaxLength       =   5
               cDataReplace    =   1
               cFaZero         =   0
               cFbComma        =   0
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   795
            Index           =   5
            Left            =   1260
            Top             =   30
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "品        番"
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
            Begin CisText_V60.CisText iH1_Hinbn 
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
            Left            =   4635
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
            Width           =   1230
            _ExtentX        =   2170
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
            Begin CisText_V60.CisText iH1_Seban 
               Height          =   375
               Left            =   90
               TabIndex        =   0
               Top             =   330
               Width           =   1065
               _ExtentX        =   1879
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
            Left            =   12630
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
               TabIndex        =   9
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
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   795
            Left            =   30
            Top             =   825
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   1402
            ForeColor       =   16711680
            Caption         =   "代 表 得 意 先"
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
            Begin CisText_V60.CisText iH1_Tokui 
               Height          =   360
               Left            =   60
               TabIndex        =   10
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
            Begin Cis3D_v60.CIS3D H1lb_Tokui 
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
      End
   End
   Begin Cis3D_v60.CIS3D HLB_Title 
      Height          =   675
      Left            =   120
      Top             =   30
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   1191
      BackColor       =   12640511
      Caption         =   "品番マスタ検索(かんばん情報)"
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
      TabIndex        =   17
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
      Left            =   8310
      ScaleHeight     =   615
      ScaleWidth      =   6015
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   30
      Width           =   6075
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   4290
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
         Left            =   5145
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
         Left            =   3435
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
         Left            =   45
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
      Begin CisBtn_60.CisBtn PB_Csv 
         Height          =   585
         Left            =   1725
         Top             =   0
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
         Left            =   885
         Top             =   0
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
      Begin CisBtn_60.CisBtn PB_Set 
         Height          =   585
         Left            =   2580
         Top             =   0
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
      Height          =   6720
      Left            =   75
      ScaleHeight     =   6660
      ScaleWidth      =   14265
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2655
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
         Bindings        =   "CKN3240.frx":0000
         Height          =   6525
         Left            =   90
         OleObjectBlob   =   "CKN3240.frx":0015
         TabIndex        =   11
         Top             =   60
         Width           =   14115
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
      ScaleWidth      =   855
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   60
      Width           =   915
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
         Left            =   60
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   300
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
         Left            =   60
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   -60
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
Attribute VB_Name = "CKN3240"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   工程かんばん情報検索
'**       フォームID    :   CKN3240 (Based CKN3235)
'**       処理概要      :
'**
'**       作  成  日    :   2010/07/23  By CIS
'**       変  更  日    :   2014/02/04  By CIS WIN7対応
'**       変  更  日    :
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
    Dim CisDBGrid       As New CisDBGrid ' TDBGridPro60ｸﾗｽ定義
    Dim mColMoveFlg     As Boolean
    Dim SortCnt         As Long
    Dim mOrder()        As String
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
    ' 手配区分
    Set CisComboTehai = New CisCombo
    With CisComboTehai
        Set .Connect = CisDB
        Set .CombObj = H1Comb_Tehai
        .Control = "名称マスタ"
        .ComboName = "手配区分"
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .Ess = False
        .ComboInit
        .ComboSet
    End With
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    
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
        .LayOutName = CisFun.WSID & "K3240.Lay"       ' ﾚｲｱｳﾄ情報格納先、ﾌｧｲﾙ名
        .PrintHeadLine = あり
        .PrintBodyLine = あり
        .PrintBodyGyo = 10
        .CsvName = Trim(Me.Caption) & ".csv"
        .CsvRegKey = gRegKey
        .CsvValue = "CKN3240"
        .GridInitGet = False                 ' ｸﾞﾘｯﾄﾞ内容初期設定 ｸﾘｯﾌﾟﾎﾞｰﾄﾞ作成ﾌﾗｸﾞ
        .GridProcGet = False
        .Init
    End With
    Call GridInitSet

    RVI_Hinmoku = 1 ' 発注
    RVI_SK = 0      ' かんばん
    RVI_Tehai = 2   ' 社内

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
' Head部初期化
    B1Op_Patn(0).Value = True
    Op_Sort(0).Value = True
'    Op_Sort(0).BackColor = mGotColor
' ADOｺﾝﾄﾛｰﾙ　接続情報　セット
    Adodc1.ConnectionString = CisDB.ConnectString
    '-----2014/02/04 EDIT START-----
    gStr = CisDB.ConnectString
    gConnect = "server=" & CisDB.Server & ";database=" & CisDB.DBName & ";User ID=" & CisDB.User & ";password=" & CisDB.PassWord(GetSetting(gRegKey, "Environment", "DrowsSap", ""))
    
    gStr = Replace(gStr, "Use Procedure", "Extended Properties=" & gConnect & ";Use Procedure")
    
    Adodc1.ConnectionString = gStr
    '-----2014/02/04 EDIT END-------
    Call HeadBodyClear("H")
    RV_Rtn = False
    
    If RVI_Hinmoku = 9 Then
    Else
       H1Ck_Hinmoku(RVI_Hinmoku).Value = 1
    End If
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
'
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
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF8:        If PB_Prt.Visible Then Call PB_Prt_Click   '【印刷】
           Case vbKeyF10:       If PB_Csv.Visible Then Call PB_Csv_Click   '【出力】
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
    Call PB_END_Click
End Sub
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
        Unload CKN3240
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
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B1" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        PB_ENT.BtnCaption = "開始"
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    CSV出力           　     +
'+----------------------------+
Private Sub PB_Csv_Click()
    With CisDBGrid
        Set .Recordset = Adodc1.Recordset
        .PutCsv
    End With
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
              Call BodySet("R")
           End If
        End If
    End With
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_UTorcd: Exit Sub
    If PB_Look.Tag = "iH1_Tokui" Then Call Look_Tokui: Exit Sub     ' 2009.12.24 add
End Sub
'+----------------------------+
'+       仕入先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 0
    RV_Call = ""
    
    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

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
    If H1Ck_Hinmoku(0).Value <> 0 Then
        RV_TorKb = 3
    Else
        RV_TorKb = 9
    End If
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
'+----------------------------+
'+       代表得意先検索       +
'+----------------------------+
Private Sub Look_Tokui()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 0
    RV_Call = ""
    
    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
        iH1_Tokui = RV_TorcdK
        H1lb_Tokui = RV_TorRName
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
    If TorNmGet(iH1_Torcd, 0) Then
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
'+-----------------------------+
'+  iH1_Tokui
'+-----------------------------+
Private Sub iH1_Tokui_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Tokui_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H1lb_Tokui = ""
    If TorNmGet(iH1_Tokui, 0) Then
       H1lb_Tokui = TRM_RName
    End If
End Sub
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
'        Call DBInput("R")
        Call BodySet("R")
        DBGrid1.SetFocus
    End If
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
    PB_Prt.Visible = B1Mode
    PB_Csv.Visible = B1Mode
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
       PB_ENT.BtnCaption = "選択"
       Call DispChange(ProcHB)
       DBGrid1.SetFocus
       GoTo ReturnPress_Ed
    End If
'    If ProcHB = "B1" Then Call RtnSet
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
    
'    gSL_Select = "SELECT HN.品番,CASE WHEN ISNULL(HN.表示品番,'') = '' Then HT.品番 Else HN.表示品番 End 表示品番,"
'    gSL_Select = gSL_Select & "    HN.社内背番,' ' + HN.社内背番 表示社内背番,HN.品名,HN.車型,"
'    gSL_Select = gSL_Select & "      ISNULL(HT.取引先,'') 取引先,ISNULL(HT.受入,'') 受入,"
'    gSL_Select = gSL_Select & "    CASE WHEN ISNULL(TR.略称,'') <> '' Then ISNULL(HT.取引先,'') + '：' + TR.略称 Else ISNULL(HT.取引先,'') End 表示取引先,"
'    gSL_Select = gSL_Select & "    CASE WHEN ISNULL(TR2.略称,'') <> '' Then ISNULL(HT.契約先,'') + '：' + TR2.略称 Else ISNULL(HT.契約先,'') End 表示契約先,"
'    gSL_Select = gSL_Select & "    CASE WHEN ISNULL(HT.手配区分,'') = '1' THEN"
'    gSL_Select = gSL_Select & "             CASE WHEN ISNULL(TR3.略称,'') <> '' Then ISNULL(HT.納入先,'') + '：' + TR3.略称 Else ISNULL(HT.納入先,'') End"
'    gSL_Select = gSL_Select & "             Else '' END  表示納入先,"
'    gSL_Select = gSL_Select & "    ISNULL(HT.背番号,'') 背番号,' ' + ISNULL(HT.背番号,'') 表示背番号,"
'    gSL_Select = gSL_Select & "    ISNULL(HT.収容数1,0) 収容数1,ISNULL(HT.SK区分,0) SK区分,"
'    gSL_Select = gSL_Select & "    CASE WHEN HT.SK区分 IS NUll THEN ''"
'    gSL_Select = gSL_Select & "             Else"
'    gSL_Select = gSL_Select & "            Case When  ISNULL(HT.SK区分,0) = 0 Then 'K'"
'    gSL_Select = gSL_Select & "            Else 'S' End End SK,"
'    gSL_Select = gSL_Select & "    CASE When ISNULL(HT.サイクル,'') <> '' Then"
'    gSL_Select = gSL_Select & "        Substring(HT.サイクル,1,1) + '-' + Substring(HT.サイクル,2,2) + '-' + Substring(HT.サイクル,4,2)"
'    gSL_Select = gSL_Select & "             Else '' End 表示サイクル,"
'    gSL_Select = gSL_Select & "    ISNull(HT.手配区分,'') 手配区分,RTRIM(ISNULL(SY.値名称,'')) 手配区分名,"
'    gSL_Select = gSL_Select & "    ISNULL(HT.品目,0) 品目,"
'    gSL_Select = gSL_Select & "    CASE WHEN HT.品目 IS Null Then ''"
'    gSL_Select = gSL_Select & "        Else"
'    gSL_Select = gSL_Select & "            Case When ISnull(HT.品目,0) = 0 Then '受' Else '発' End"
'    gSL_Select = gSL_Select & "        End 品目名,"
'    gSL_Select = gSL_Select & "    ISNULL(HT.収容器1,'') 収容器1,"
'    gSL_Select = gSL_Select & "    CASE WHEN ISNULL(HT.変更日,'') <> '' THEN"
'    gSL_Select = gSL_Select & "        SubString(HT.変更日,1,4) + '/' + SubString(HT.変更日,5,2) + '/' + SubString(HT.変更日,7,2) + '-' + HT.変更便"
'    gSL_Select = gSL_Select & "        Else '' END 変更日便,"
'    gSL_Select = gSL_Select & "    ISNULL(HT.収容器2,'') 収容器2,ISNULL(HT.収容数2,0) 収容数2,"
'    gSL_Select = gSL_Select & "    ISNULL(HT.納入先受入,'') 納入先受入,ISNULL(HT.置場,'') 置場,"
'    gSL_Select = gSL_Select & "    ISNULL(HT.アドレス,'') アドレス"
'    gSL_Select = gSL_Select & "   ,CASE WHEN ISNULL(TR4.略称,'') <> '' Then ISNULL(HN.得意先,'') + '：' + TR4.略称 Else ISNULL(HN.得意先,'') End 取引先名"
'    gSL_Select = gSL_Select & "   ,ISNULL(HT.最終発行連番,'') 最終発行連番"
'    gSL_Select = gSL_Select & "   ,CASE WHEN ISNULL(HN.部品区分,0) = 0 THEN '部品' "
'    gSL_Select = gSL_Select & "         WHEN ISNULL(HN.部品区分,0) = 1 THEN '外注加工' "
'    gSL_Select = gSL_Select & "         WHEN ISNULL(HN.部品区分,0) = 2 THEN '補助材' "
'    gSL_Select = gSL_Select & "    ELSE '' END 部品区分名 "
'    gSL_Select = gSL_Select & "    From 品番取引先マスタ HT"
'    gSL_Select = gSL_Select & "    Left Outer Join 品番マスタ HN"
'    gSL_Select = gSL_Select & "    ON HT.品番 = HN.品番"
'    gSL_Select = gSL_Select & "    Left Outer Join 取引先マスタ TR"
'    gSL_Select = gSL_Select & "    ON  TR.取引先区分 = ISNULL(HT.品目,'')"
'    gSL_Select = gSL_Select & "    AND TR.取引先CD = ISNULL(HT.取引先,'')"
'    gSL_Select = gSL_Select & "    Left Outer Join 取引先マスタ TR2"
'    gSL_Select = gSL_Select & "    ON  TR2.取引先区分 = ISNULL(HT.品目,'')"
'    gSL_Select = gSL_Select & "    AND TR2.取引先CD = ISNULL(HT.契約先,'')"
'    gSL_Select = gSL_Select & "    Left Outer Join 取引先マスタ TR3"
'    gSL_Select = gSL_Select & "    ON  TR3.取引先区分 = ISNULL(HT.品目,'')"
'    gSL_Select = gSL_Select & "    AND TR3.取引先CD = ISNULL(HT.納入先,'')"
'    gSL_Select = gSL_Select & "    Left Outer Join 取引先マスタ TR4"
'    gSL_Select = gSL_Select & "     ON  TR4.取引先CD   = ISNULL(HN.得意先,'')"
'    gSL_Select = gSL_Select & "     AND TR4.取引先区分 = 0"
'    gSL_Select = gSL_Select & "    Left Outer Join 名称マスタ SY"
'    gSL_Select = gSL_Select & "    ON SY.区分名称 = '手配区分'"
'    gSL_Select = gSL_Select & "    And Sy.区分タイプ = ''"
'    gSL_Select = gSL_Select & "    And SY.値 = ISNULL(HT.手配区分,'')"
    
    
' ﾚｲｱｳﾄ表示内容よりﾃﾞｰﾀ連結名を求める
    CisDBGrid.LayOutGet
    
    gSL_Select = "select " & CisDBGrid.LayOutfields
    gSL_Select = gSL_Select & " from 工程かんばん情報検索ビュー "
' 分類内容(Order By句) ｾｯﾄ
    With CisDBGrid
        .OrderSet (Op_Sort(0).Value)
        If .Select_OrderBy = "" Then
            BodySet = False
            'GoTo DBInput_Ed
            Exit Function
        End If
    End With
    
    
'    mSL_OrderA = " Order By ISNULL(HT.取引先,''),HN.品番 "
'    mSL_OrderD = " Order By ISNULL(HT.取引先,'') DESC,HN.品番 Desc"
'

' 選択条件セット
    mSL_SelectW = ""
    With CisFun
        .Where_CreateKB = NewDocument
        .Where_Create AlphaMe, "ISNULL(背番号,'')", iH1_Seban
        .Where_Create AlphaMe, "品番", iH1_Hinbn
        .Where_Create AlphaMe, "社内背番", iH1_SSeban
        .Where_Create AlphaMe, "ISNULL(取引先,'')", iH1_Torcd
        .Where_Create AlphaMe, "ISNULL(受入,'')", iH1_Ukeir
        .Where_Create AlphaMe, "ISNULL(得意先,'')", iH1_Tokui            ' 2009.12.24 add
        If Not (H1Ck_Hinmoku(0).Value = H1Ck_Hinmoku(1).Value) Then
           If H1Ck_Hinmoku(0).Value = 1 Then
              .Where_Create Numeric, "ISNULL(品目,9)", "0", 等しい, Create
           End If
           If H1Ck_Hinmoku(1).Value = 1 Then
              .Where_Create Numeric, "ISNULL(品目,9)", "1"
           End If
        End If
        If Not (H1Ck_SK(0).Value = H1Ck_SK(1).Value) Then
           If H1Ck_SK(0).Value = 1 Then
              .Where_Create Numeric, "ISNULL(SK区分,9)", "0", 等しい, Create
           End If
           If H1Ck_SK(1).Value = 1 Then
              .Where_Create Numeric, "ISNULL(SK区分,9)", "1"
           End If
        End If
        CisComboTehai.CodeGet
        If Trim(CisComboTehai.Code) <> "" Then
           .Where_Create Numeric, "ISNULL(手配区分,9)", Trim(CisComboTehai.Code), 等しい, Create
        End If
                
        mSL_SelectW = .Where_Phrase
    End With
    
'    Adodc1.RecordSource = gSL_Select & mSL_SelectW & SortGet
    gSL_Select = gSL_Select & mSL_SelectW & CisDBGrid.Select_OrderBy
    Adodc1.RecordSource = gSL_Select
    
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
       SortGet = " Order By " & mSL_Order
    Else
       SortGet = ""
    End If
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
         Call .GridInitSet(True, , "表示取引先", "取引先", 25)
         Call .GridInitSet(True, , "受入", "受", 2)
         Call .GridInitSet(True, , "表示品番", "品     番", 30)
         Call .GridInitSet(True, , "車型", "車型", 5)
         Call .GridInitSet(True, , "表示社内背番", "社内背番", 10)
         Call .GridInitSet(True, , "表示背番号", "背番号", 10)
         Call .GridInitSet(, , "収容数1", "収容数1", 6)
         Call .GridInitSet(True, , "SK", "SK", 4)
         Call .GridInitSet(True, , "表示サイクル", "サイクル", 8)
         Call .GridInitSet(True, , "手配区分名", "手配", 4)
         Call .GridInitSet(True, , "品名", "品            名", 35)
         Call .GridInitSet(True, , "品目名", "品目", 4)
         Call .GridInitSet(True, , "収容器1", "収容器1", 10)
         Call .GridInitSet(True, , "変更日便", "変更日-便", 12)
         Call .GridInitSet(, , "収容数2", "収容数2", 6)
         Call .GridInitSet(True, , "収容器2", "収容器2", 10)
         Call .GridInitSet(True, , "表示契約先", "契約先", 25)
         Call .GridInitSet(True, , "表示納入先", "納入先", 25)
         Call .GridInitSet(True, , "納入先受入", "納受", 4)
'''         Call .GridInitSet(True, , "アドレス", "アドレス", 9)
         Call .GridInitSet(True, , "置場", "置場", 11)
         Call .GridInitSet(True, , "部品区分名", "部品区分", 8)
         Call .GridInitSet(True, , "取引先名", "代表得意先", 25)
         Call .GridInitSet(, , "最終発行連番", "か枚数", 6)
         Call .GridInitSet(True, , "品目区分", "品目区分", 8)
         Call .GridInitSet(, , "ロット1", "ロット1", 8)
         Call .GridInitSet(, , "ロット2", "ロット2", 8)
         Call .GridInitSet(, , "返却場所", "返却場所", 8)
         Call .GridInitSet(True, , "後工程", "後工程", 20)
         Call .GridInitSet(True, , "返却場所(後)", "後工程返却場所", 12)
         Call .GridInitSet(True, , "生区", "生産区分", 4)
    End With
End Sub
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
Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If DBGrid1.RowContaining(Y) < 0 Then Exit Sub
    If DBGrid1.ColContaining(x) < 0 Then Exit Sub
    SendKeys "{RIGHT}"
    DBGrid1.Row = DBGrid1.RowContaining(Y)
    DoEvents
'    Call RtnSet
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
''+-------------------------------------+
''+      選択内容ｾｯﾄ & 復帰ｽﾃｰﾀｽｾｯﾄ      +
''+-------------------------------------+
'Private Sub RtnSet()
'    With Adodc1
'        RV_Hinbn = .Recordset("品番")
'        RV_Seban = .Recordset("背番号")
'        RV_SSeban = .Recordset("社内背番")
'        RV_Hinnm = .Recordset("品名")
'        RV_TorcdK = .Recordset("取引先")
'        RV_Ukeir = .Recordset("受入")
'        RV_Hinmoku = .Recordset("品目")
'        RV_Tehai = .Recordset("手配区分")
'    End With
'    Adodc1.Recordset.Close
'    RV_Rtn = True
'    Call PB_END_Click
'End Sub

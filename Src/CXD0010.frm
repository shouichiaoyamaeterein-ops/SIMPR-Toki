VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form CXD0010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "データ削除・退避"
   ClientHeight    =   11100
   ClientLeft      =   75
   ClientTop       =   1500
   ClientWidth     =   15105
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
   ForeColor       =   &H00004000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11100
   ScaleWidth      =   15105
   WindowState     =   2  '最大化
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   60
      Top             =   1620
   End
   Begin VB.PictureBox B2_Area3 
      BackColor       =   &H00FF8080&
      Height          =   3915
      Left            =   11205
      ScaleHeight     =   3855
      ScaleWidth      =   3390
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3750
      Visible         =   0   'False
      Width           =   3450
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000000FF&
         Height          =   3405
         Index           =   8
         Left            =   60
         ScaleHeight     =   3345
         ScaleWidth      =   3225
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   390
         Width           =   3285
         Begin Cis3D_v60.CIS3D Blb_cCnt 
            Height          =   285
            Index           =   4
            Left            =   2325
            Top             =   0
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
            BackColor       =   14737632
            ForeColor       =   12583104
            Caption         =   "99,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   285
            Index           =   7
            Left            =   0
            Top             =   0
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   503
            ForeColor       =   12582912
            Caption         =   "受注ﾃﾞｰﾀ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D Blb_cCnt 
            Height          =   285
            Index           =   5
            Left            =   2325
            Top             =   300
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
            BackColor       =   14737632
            ForeColor       =   12583104
            Caption         =   "99,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   285
            Index           =   8
            Left            =   -15
            Top             =   300
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   503
            ForeColor       =   12582912
            Caption         =   "出荷ﾃﾞｰﾀ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   345
         Index           =   15
         Left            =   60
         Top             =   15
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   609
         BackColor       =   0
         ForeColor       =   65280
         Caption         =   "削 除 件 数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   7
         cPositionX      =   30
         cPositionY      =   10
      End
   End
   Begin VB.PictureBox H1_Area2 
      BackColor       =   &H00808080&
      Height          =   4425
      Left            =   3765
      ScaleHeight     =   4365
      ScaleWidth      =   7305
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3225
      Width           =   7365
      Begin VB.PictureBox Picture1 
         Height          =   2790
         Index           =   0
         Left            =   165
         ScaleHeight     =   2730
         ScaleWidth      =   6930
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   615
         Width           =   6990
         Begin Cis3D_v60.CIS3D UC_3D2 
            Height          =   1065
            Left            =   2280
            Top             =   90
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   1879
            ForeColor       =   16512
            Caption         =   "ﾃｷｽﾄ出力"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   60
            Begin VB.OptionButton Op_TextOut 
               BackColor       =   &H00C0C0C0&
               Caption         =   "無し"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   1230
               TabIndex        =   2
               Top             =   540
               Width           =   795
            End
            Begin VB.OptionButton Op_TextOut 
               BackColor       =   &H00C0C0C0&
               Caption         =   "有り"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   210
               TabIndex        =   1
               Top             =   540
               Value           =   -1  'True
               Width           =   795
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   1065
            Index           =   17
            Left            =   60
            Top             =   90
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   1879
            ForeColor       =   128
            Caption         =   "削除年月"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   60
            Begin CisYM_V60.CisYM iH1_YM 
               Height          =   405
               Left            =   255
               TabIndex        =   0
               Top             =   465
               Width           =   1635
               _ExtentX        =   2646
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
               cContent        =   1
            End
         End
         Begin VB.PictureBox Bar 
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   1
            Left            =   450
            ScaleHeight     =   165
            ScaleWidth      =   6390
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   2235
            Visible         =   0   'False
            Width           =   6450
         End
         Begin VB.PictureBox Bar 
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   0
            Left            =   450
            ScaleHeight     =   165
            ScaleWidth      =   6390
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   2445
            Visible         =   0   'False
            Width           =   6450
         End
         Begin Cis3D_v60.CIS3D UC_3D1 
            Height          =   1080
            Index           =   12
            Left            =   30
            Top             =   1155
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   1905
            BackColor       =   0
            ForeColor       =   65535
            Caption         =   "前回内容"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   6
            cPositionX      =   10
            cPositionY      =   80
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFC0&
            Height          =   1080
            Index           =   2
            Left            =   480
            ScaleHeight     =   1020
            ScaleWidth      =   6330
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1155
            Width           =   6390
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "(   処 理 日   )"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   240
               Index           =   0
               Left            =   540
               TabIndex        =   17
               Top             =   630
               Width           =   1665
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "(削除対象年月)"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004080&
               Height          =   240
               Index           =   1
               Left            =   525
               TabIndex        =   16
               Top             =   150
               Width           =   1710
            End
            Begin VB.Label H1lb_YM 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "Label3"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2460
               TabIndex        =   15
               Top             =   165
               Width           =   810
            End
            Begin VB.Label H1lb_PDate 
               AutoSize        =   -1  'True
               BackStyle       =   0  '透明
               Caption         =   "9999年99月99日 99:99"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2445
               TabIndex        =   14
               Top             =   630
               Width           =   2655
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D1 
            Height          =   1065
            Left            =   4650
            Top             =   90
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   1879
            ForeColor       =   16512
            Caption         =   "ﾃｷｽﾄ出力"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   6
            cPositionY      =   60
            Begin VB.OptionButton H1op_PType 
               BackColor       =   &H00C0C0C0&
               Caption         =   "有り"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   3
               Top             =   555
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton H1op_PType 
               BackColor       =   &H00C0C0C0&
               Caption         =   "無し"
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   1
               Left            =   1140
               TabIndex        =   4
               Top             =   525
               Width           =   975
            End
            Begin Cis3D_v60.CIS3D UC_3D1 
               Height          =   390
               Index           =   3
               Left            =   0
               Top             =   0
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   688
               BackColor       =   16744703
               Caption         =   "ﾃﾞｰﾀﾍﾞｰｽ退避"
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
         End
      End
      Begin Cis3D_v60.CIS3D H1lb_msg 
         Height          =   465
         Left            =   180
         Top             =   135
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   820
         BackColor       =   16761087
         Caption         =   "削除年月 以前のデータを削除します"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
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
      End
      Begin Cis3D_v60.CIS3D Back_Kensu 
         Height          =   2715
         Left            =   -5760
         Top             =   5370
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   4789
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "【削除件数】"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cFont3D         =   2
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   405
         Index           =   29
         Left            =   165
         Top             =   3885
         Width           =   6960
         _ExtentX        =   12277
         _ExtentY        =   714
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
         Begin Cis3D_v60.CIS3D PB_Open 
            Height          =   315
            Left            =   6015
            Top             =   30
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            Caption         =   "開く"
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
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D Dir_DelDB 
            Height          =   315
            Left            =   5550
            Top             =   30
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   556
            Caption         =   "…"
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
            cBoderWidth     =   3
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D H1lb_SaveDB 
            Height          =   315
            Left            =   30
            Top             =   30
            Width           =   5490
            _ExtentX        =   9684
            _ExtentY        =   556
            BackColor       =   12648384
            Caption         =   "C:\Program Files\Prpss\BackUp\"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
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
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   465
         Index           =   30
         Left            =   165
         Top             =   3420
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   820
         BackColor       =   16761024
         ForeColor       =   16711680
         Caption         =   "テ キ ス ト 退 避 先"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   14.25
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
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Index           =   9
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   0
   End
   Begin Cis3D_v60.CIS3D PB_ENT 
      Height          =   735
      Left            =   8055
      Top             =   8025
      Visible         =   0   'False
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   1296
      ForeColor       =   16576
      Caption         =   "処理開始"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
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
      cButton         =   -1  'True
   End
   Begin Cis3D_v60.CIS3D PB_END 
      Height          =   735
      Left            =   5175
      Top             =   8025
      Visible         =   0   'False
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   1296
      ForeColor       =   16711680
      Caption         =   "終  了"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   18
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
      cButton         =   -1  'True
   End
   Begin Cis3D_v60.CIS3D PB_OK 
      Height          =   735
      Left            =   6615
      Top             =   8025
      Visible         =   0   'False
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   1296
      ForeColor       =   16711680
      Caption         =   "確  認"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   18
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
      cButton         =   -1  'True
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
      Left            =   60
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1620
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   150
      Top             =   690
   End
   Begin VB.Timer Timer2 
      Left            =   90
      Top             =   1140
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   661
      BackColor       =   16761024
      ForeColor       =   0
      Caption         =   "【 データ削除・退避 】"
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
         Left            =   14445
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BackColor       =   16761024
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
         cBoderColor1    =   16761024
         cBoderColor2    =   16761024
         cFont3DColor1   =   8454143
         cFont3DColor2   =   16777215
         cFont3D         =   1
         cAlingnment     =   7
         cBoderStyle     =   2
      End
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   285
         Left            =   11940
         Top             =   30
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
   Begin Cis3D_v60.CIS3D lb_Err 
      Height          =   735
      Left            =   1830
      Top             =   1350
      Visible         =   0   'False
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   1296
      BackColor       =   255
      ForeColor       =   16777215
      Caption         =   "Err"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cFont3D         =   2
      cAlingnment     =   7
   End
   Begin VB.Label PNL_END 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5220
      TabIndex        =   12
      Top             =   8805
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label PNL_ENT 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( Enter )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   8160
      TabIndex        =   11
      Top             =   8835
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label PNL_OK 
      Alignment       =   2  '中央揃え
      BackStyle       =   0  '透明
      Caption         =   "( F12 )"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6660
      TabIndex        =   10
      Top             =   8805
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label BG_Logo 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  '透明
      Caption         =   "SIMPR-A"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   285
      Index           =   0
      Left            =   135
      TabIndex        =   7
      Top             =   420
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "CXD0010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  データ削除・退避
'**       フォームID    :  CXD0010
'**       処理概要      :
'**
'**       作  成  日    :  2004/04/25  By CIS
'**       変  更  日    :  2004/12/27  By CIS 生産指示テーブル・生産実績テーブル　削除追加
'**       変  更  日    :  2005/01/11  By CIS テーブル追加
'**       変  更  日    :  2005/03/04  By CIS 1.重複起動ﾁｪｯｸで起動済の時、『実行時エラー'91'』を表示する対応（Timerﾌﾟﾛﾊﾟﾃｨの Interval=0 に修正）
'**       　  　  　                          2.[処理開始]＆[確認]＆[終了]のﾌﾟﾛﾊﾟﾃｨを修正（cButton=True）
'**
'**********************************************************'
    Dim CISFrm          As CISFormContorl
    
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim SL_Select       As String       ' SQL SELECT
    Dim mMsgType        As String
 
    Dim wObj            As Object       ' ｵﾌﾞｼﾞｪｸﾄﾜｰｸｴﾘｱ
    Dim wOpObj          As OptionButton
    
    Dim W_Str           As String
    Dim w_Byte          As Byte
    Dim w_Test          As Boolean

    Dim w_ChkMon        As Byte         ' 入力チェック月数
    Dim w_ProcTime(4)   As String
    Dim m_SaveDB        As String

    Dim OpGotColor      As Long         ' ｵﾌﾟｼｮﾝﾎﾞﾀﾝ　GotFocus色
    Dim OpLostColor     As Long         ' ｵﾌﾟｼｮﾝﾎﾞﾀﾝ　LostFocus色

    Dim mZYm            As String       ' 前回指定年月
    
'    Dim C_Mater         As New Mater    'ｸﾗｽのｲﾝｽﾀﾝｽを宣言
'    Dim C_Mater2        As New Mater    'ｸﾗｽのｲﾝｽﾀﾝｽを宣言
    
    Dim mSaveDir        As String       ' 退避格納先
    Dim SaveDBName      As String       ' 削除分データベース名
    Dim mSaveMsg        As String
    
    Dim mMax_File       As Long
    Dim mNow_File       As Long
    Dim mTblName()      As String
    Dim mTblCnt()       As Long
    Dim mItemRep()      As String
    Dim mCrtTypee()     As Integer
    
    Dim Bar_Max(1)      As Long
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
    gConnectCount = 3
    If Not IniGet Then End
    
    Set CISFrm = New CISFormContorl
    Set CISFrm.ThisForm = Me
    
' ｵﾌﾟｼｮﾝﾎﾞﾀﾝ　色設定
    OpGotColor = &HC0FFC0
    OpLostColor = &HC0C0C0
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    mSaveMsg = H1lb_msg.Caption
    
    Call CISFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
'   < シングルプロセスチェック >
    If Not SingleProcess_Check Then PB_END_Click
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    With CISFrm
        Set .MSGObject = H1lb_msg
        Set .MSGTimer = Timer2
        .MSGStopCount = 10
        .MsgType = Left_To_Right
        .MSGText = mSaveMsg
        .MSGInterval = 100
        .MSGRelate
    End With

    Dummy.Left = -1000
    Dummy.Enabled = False
'+---------------------+
'+  ﾌｧｲﾙ処理初期設定    +
'+---------------------+
    With CisFun
        .Log_PutFlg = True
        .Log_FileInitial = "DLog"
        .Log_FilePath = gIniExe & "Log"
        .Log_DataHDTime = あり
        .Log_DelDay = 0
    End With
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   #     ロゴ表示       #
'   #-------------------#
    If gLogoDisp Then
       BG_Logo(0).Caption = gLogo
       CISFrm.BGLDisp BG_Logo, 10
    End If
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CISFrm.Frm_Center(True)

    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    
    If Command = "TEST" Then w_Test = True
    
    Call SetZenkai           ' 前回処理内容表示
    
    ' テキスト出力
    Op_TextOut(0).Value = True
    Op_TextOut(0).BackColor = OpGotColor
    ' ﾃﾞｰﾀﾍﾞｰｽ退避
    H1op_PType(0).Value = True
    H1op_PType(0).BackColor = OpGotColor
    
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ｐａｉｎｔ イベント》                   +
'+----------------------------------------------------------+
Private Sub Form_Paint()
'   ( グラデーション )
    With CISFrm
        .Gradation_Direction = Landscape
        .Gradation_Shading = Lighten
        .Gradation 80, 80, 200
    End With
'    Call GradationFormX(Me, 80, 80, 200, 4, 4, 4)
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》       +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    If Not FormAct Then Exit Function
    
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    
    Select Case Key_Code
           Case vbKeyF7:        If Dir_DelDB.Visible Then Call Dir_DelDB_Click  '【検索】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    '【終了】
                                If PB_OK.Visible Then Call PB_OK_Click      '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                      '【入力】
                                    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT_Click
                                    Dummy.Enabled = False
                                End If
          Case Else:            Exit Function
    End Select
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    
    Call SingleProcess_Release
    
    Call CisDB.DBDISConnect
    Call CISFrm.UnLoadDisp
    End

End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
   Call ReturnPress
End Sub
'+---------------------+
'+    確  認  キ ー     +
'+---------------------+
Private Sub PB_OK_Click()
    Call PB_END_Click
End Sub
'+-----------------------------------+
'+    削除データベース退避先(F7)       +
'+-----------------------------------+
Private Sub Dir_DelDB_Click()
    With CisFun
        .Folder_Flags = ffDirOnly
        .Folder_DefDir = H1lb_SaveDB
        .FolderSel
        gStr = RTrim(.Folder_Result)
    End With
    If Trim(gStr) <> "" Then
        H1lb_SaveDB = gStr
        H1lb_SaveDB = RTrim(H1lb_SaveDB)
        If Right(H1lb_SaveDB, 1) <> "\" Then H1lb_SaveDB = H1lb_SaveDB & "\"
    
    End If
    
End Sub
'+-----------------------------------+
'+    削除データベース退避先(開く)   +
'+-----------------------------------+
Private Sub PB_Open_Click()
    Dim resL As Long
    resL = ShellExecute(Me.hWnd, "explore", H1lb_SaveDB, vbNullString, vbNullString, 1)
    If resL = 2 Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "      削除退避先が存在しません。         "
            .MB_Title = "ﾌｫﾙﾀﾞ確認"
            .MB_Button = OK
            .MBOX
        End With
    End If
End Sub
'+--------------------------------------+
'+     TEXT出力 (Op_TextOut)            +
'+--------------------------------------+
Private Sub Op_TextOut_GotFocus(Index As Integer)
    For Each wOpObj In Op_TextOut
        wOpObj.BackColor = OpLostColor
    Next wOpObj
    
    Op_TextOut(Index).BackColor = OpGotColor
End Sub
'+--------------------------------------+
'+     ﾃﾞｰﾀﾍﾞｰｽ退避 (H1op_PType)        +
'+--------------------------------------+
Private Sub H1op_PType_GotFocus(Index As Integer)
    For Each wOpObj In H1op_PType
        wOpObj.BackColor = OpLostColor
    Next wOpObj
    
    H1op_PType(Index).BackColor = OpGotColor
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
    Dim B1Mode      As Boolean
    Dim B2Mode      As Boolean
    H1Mode = False
    B1Mode = False
    B2Mode = False
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
        Case "H1":      H1Mode = True
        Case "B1":      B1Mode = True
        Case "B2":      B2Mode = True
    End Select
    
 '   B2_Area1.Visible = B1Mode Or B2Mode

    PB_END.Visible = H1Mode
    PNL_END.Visible = H1Mode
    PB_ENT.Visible = H1Mode
    PNL_ENT.Visible = H1Mode
    
    PB_OK.Visible = B2Mode
    PNL_OK.Visible = B2Mode
    
' ﾀﾞﾐｰ項目(入力禁止)
    If H1Mode Then Dummy.Enabled = False
End Function
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CISFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        
        gStr = iH1_YM.cYear & " 年 " & iH1_YM.cMonth & " 月以前"
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = StrConv(gStr, vbWide) & "のデータ削除処理を行います      "
            .MB_MSG(4) = "                よろしいですか？                     "
            .MB_Title = "削除確認"
            .MB_Button = OK_CAN
            If Not .MBOX Then GoTo ReturnPress_Ed
        End With
        ProcHB = "B1"
        Call DispChange(ProcHB)

        Call DeleteProc
        
        ProcHB = "B2"
        Call DispChange(ProcHB)
        H1_Area2.Enabled = False
        GoTo ReturnPress_Ed
    End If
ReturnPress_Ed:
    Call CISFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
 
    If iH1_YM = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "   削除対象年月を入力して下さい  "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
            iH1_YM.SetFocus
        End With
        Exit Function
    End If
    
    Static sw_CYM       As String
'
    sw_CYM = Format(DateAdd("m", w_ChkMon * (-1), Now()), "yyyymm")

'
    If Not iH1_YM.cChkResult Then
       With CisFun
           .MB_Lines = 5
           .MB_MSG(2) = "   正しい年月を入力して下さい  "
           .MB_Title = ""
           .MB_Button = Error
           .MBOX
           iH1_YM.SetFocus
       End With
       Exit Function
    End If
'
    If Trim(mZYm) <> "" Then
        If Trim(iH1_YM) < Trim(mZYm) Then
           With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "   前回年月より前の年月は処理不可  "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
                iH1_YM.SetFocus
           End With
           Exit Function
        End If
    
        If Trim(iH1_YM) = Trim(mZYm) Then
           With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "  処理済み、再処理を行いますか？   "
                .MB_MSG(4) = "  OK = 続行、ｷｬﾝｾﾙ = 中止 "
                .MB_Title = ""
                .MB_Button = OK_CAN
                If Not .MBOX Then
                   iH1_YM.SetFocus
                   Exit Function
                End If
           End With
        End If
    End If
    
    If Not w_Test Then
        If iH1_YM >= sw_CYM Then
           With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "   " & Mid(sw_CYM, 1, 4) & "年" & _
                             Mid(sw_CYM, 5, 2) & "月以降の年月は指定出来ません"
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
                iH1_YM.SetFocus
           End With
           Exit Function
        End If
    End If
    
    If Op_TextOut(0).Value Then
        With CisFun
            .FE_Path = H1lb_SaveDB
            
            If Not .DirCheck(NoMsg_NoCrt) Then
               .MB_Lines = 5
               .MB_MSG(2) = "      削除退避先が存在しません。         "
               .MB_MSG(4) = "           作成しますか？                "
               .MB_Title = "ﾌｫﾙﾀﾞ確認"
               .MB_Button = OK_CAN
               If Not .MBOX Then Exit Function
               Call .DirCheck(NoMsg_Crt)
            End If
        End With
    End If
    
    Head1Chk = True
End Function

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*******************************
'*      メッセージ表示        *
'*******************************
Private Sub Timer2_Timer()
    CISFrm.MSGRelate
End Sub
'*******************************
'*     エラーメッセージ表示      *
'*******************************
Private Sub Timer3_Timer()
    If Not lb_Err.Visible Then
       lb_Err.Visible = True
       For gLong = 0 To 10000
       Next gLong
    Else
       lb_Err.Visible = False
    End If
End Sub
'*******************************
'*      前回処理内容表示        *
'*******************************
Private Sub SetZenkai()
' 名称マスタ検索　確認
    
    mZYm = ""
    
    Call ItemsClearSYM
    With SYM
        .区分名称 = "入力値管理"
        .区分桁数 = 2
        .値 = "09"
        .値名称 = "削除管理"
        .数字1 = 1
        .制御区分 = "X"
    End With
    
    If SYMReadChk Then
       If Trim(SYM.英数字1) <> "" Then
          H1lb_PDate = Mid(Trim(SYM.英数字1), 1, 4) & "年" & _
                       Mid(Trim(SYM.英数字1), 5, 2) & "月" & _
                       Mid(Trim(SYM.英数字1), 7, 2) & "日 " & _
                       Mid(Trim(SYM.英数字1), 9, 5)

          H1lb_YM = Mid(SYM.英数字2, 1, 4) & " 年 " & Mid(SYM.英数字2, 5, 2) & " 月"
          mZYm = Mid(SYM.英数字2, 1, 6)
       End If
    End If
'   * チェック月数 *
    w_ChkMon = SYM.数字1
    
    If w_ChkMon = 0 Then w_ChkMon = 3
'   * 削除退避ディレクトリ *
    With CisFun
        .INI_Section = "DataBase"
        .INI_FileID = gIniExe & gIniName
        .INI_Default = gIniExe & "DBDel\"
        .INI_Key = "DBDeleteDir"
        .GetIni
        W_Str = RTrim(.INI_String)
    End With
'    Call GetIni("削除退避", "DataBese", gIniExe & "BackUp\", W_Str, gIniExe, gIniName)
    If Right(RTrim(W_Str), 1) <> "\" Then W_Str = W_Str & "\"
    H1lb_SaveDB = W_Str

End Sub
'*******************************
'*      今回指定内容更新        *
'*******************************
Private Sub PutKonkai()
' 今回処理日時　更新
    gSL_Select = "UPDATE 名称マスタ SET "
    gSL_Select = gSL_Select & " 英数字1 = '" & Format(Now(), "YYYYMMDDhh:nn") & "',"
    gSL_Select = gSL_Select & " 英数字2 = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & " WHERE 区分名称 = '入力値管理'"
    gSL_Select = gSL_Select & "   AND 値 = '09'"
    
    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With

'   * 削除退避ディレクトリ（ＤＢ） *
    With CisFun
        .INI_Section = "削除処理"
        .INI_FileID = gIniExe & gIniName
        .INI_Default = gIniExe & "DBDel\"
        .INI_Key = "DBDeleteDir"
        .INI_String = RTrim(H1lb_SaveDB)
        .PutIni
    End With
End Sub
'*---------------------*
'*  データ削除＆退避    *
'*---------------------*
Private Function DeleteProc()
    
    DeleteProc = False
    
    Static sw_Cnt       As Byte
    
    On Error GoTo DeleteProc_Err
' 削除対象テーブル名セット
    Call CrtSql(9, 0, "")
    
' 退避先ﾌｫﾙﾀﾞ作成
    mSaveDir = H1lb_SaveDB & iH1_YM
    
    gLong = 0
    Do
        If gLong <> 0 Then
           gStr = mSaveDir & "_" & Format(gLong, "00")
        Else
           gStr = mSaveDir
        End If
        With CisFun
            .FE_Path = gStr
            If Not .DirCheck(NoMsg_NoCrt) Then
               mSaveDir = gStr
               .DirCheck (NoMsg_Crt)
               Exit Do
            End If
        End With
        gLong = gLong + 1
    Loop
'
    Bar(1).Visible = True
    
    Timer2.Enabled = False
    H1lb_msg = "【　　処　理　中　　】": DoEvents

' ﾃﾞｰﾀﾍﾞｰｽﾊﾞｯｸｱｯﾌﾟ
    If H1op_PType(0).Value Then
        Call CisFun.LogPut("データベース退避　『開始』")
        Call db_Init
        db_Proc = 0
        db_ProcType = "D"
        If Not db_BackUp Then
           Call CisFun.LogPut("データベース退避　処理異常終了")
           GoTo DeleteProc_Err
        End If
        Call CisFun.LogPut("データベース退避　『終了』")
    End If
'
    Call CisDB.DBTran(TransBegin)
    
' テキスト出力 & データ削除
'
    H1lb_msg = "データ 削除 【処理中】": DoEvents
   
    Call CisFun.LogPut("データ 削除処理【開始】")
    Call CisFun.LogPut("　　　 削除年月 = " & iH1_YM.cYear & "/" & iH1_YM.cMonth)
    
    mNow_File = 0
    Bar_Max(1) = mMax_File
    Bar_Max(0) = mMax_File * 100
    With CISFrm
        Set .MeterPic = Bar(1)
        .MeterMax = Bar_Max(1)
        .MeterReset
    End With
    
    If Op_TextOut(0).Value Then
        With CISFrm
            Set .MeterPic = Bar(0)
            .MeterMax = Bar_Max(0)
            .MeterReset
        End With
        Bar(0).Visible = True
    Else
        Bar(0).Visible = False
    End If
    
'    For gInt = 30 To mMax_File - 1         ' ﾃｽﾄ用
    For gInt = 0 To mMax_File - 1
        
        If Right(mSaveDir, 1) <> "\" Then mSaveDir = mSaveDir & "\"
        
        Call CisFun.LogPut("         【 " & mTblName(gInt) & " 】")
        
        mNow_File = mNow_File + 1
        
        gSL_Select = "select * from dbo.sysobjects where id = object_id(N'[" & mTblName(gInt) & "]')"
        gSL_Select = gSL_Select & " and OBJECTPROPERTY(id, N'IsUserTable') = 1"
        CisDB.SQL = gSL_Select
        If CisDB.DBExec Then
            If CisDB.RecordCount <> 0 Then
                If Op_TextOut(0).Value Then
                    ' ﾃｷｽﾄ出力
                    H1lb_msg = "【" & mTblName(gInt) & "】ﾃｷｽﾄ出力 【処理中】": DoEvents
                        
                    Call CrtSql(0, gInt, mTblName(gInt))
                    If Not SaveTextProc(mTblName(gInt)) Then
                      Call CisFun.LogPut("             ﾃｷｽﾄ出力処理エラー")
                      GoTo DeleteProc_Err
                    End If
                End If
                ' データ削除
                H1lb_msg = "【" & mTblName(gInt) & "】ﾃﾞｰﾀ削除 【処理中】": DoEvents
                Call CrtSql(1, gInt, mTblName(gInt))
                If Trim(gSL_Select) <> "" Then
                   CisDB.SQL = gSL_Select
                   If Not CisDB.DBExec Then
                      Call CisFun.LogPut("             削除処理エラー")
                      GoTo DeleteProc_Err
                   End If
                End If
            Else
                  Call CisFun.LogPut("             削除処理エラー " & mTblName(gInt) & "は存在しません")
            End If
        Else
              Call CisFun.LogPut("             削除処理エラー")
        End If
        
        With CISFrm
            Set .MeterPic = Bar(1)
            .MeterMax = Bar_Max(1)
            .MeterUpdate mNow_File
        End With

        Call CisFun.LogPut("                    ★★★ 削除件数     = " & CisDB.RecordCount & " ★★★")
    Next gInt

    With CISFrm
        Set .MeterPic = Bar(0)
        .MeterMax = Bar_Max(0)
        .MeterUpdate Bar_Max(0)
    End With
    
    Call CisFun.LogPut("データ 削除処理【終了】")
'
    If Not w_Test Then Call CisDB.DBTran(TransCommit) Else Call CisDB.DBTran(TransRollback)

' 今回処理内容退避
    Call PutKonkai
    
    On Error GoTo 0
    DeleteProc = True
    H1lb_msg = ""
    
    Timer2.Enabled = True
    With CISFrm
        .MsgType = FSize_Changes
        .MSGText = "処　理　終　了"
    End With
    
    Call CisFun.LogPut("データ 削除処理【正常終了】")
    
    Exit Function
DeleteProc_Err:
    H1lb_msg = ""
    Timer2.Enabled = True
    With CISFrm
        .MsgType = FSize_Changes
        .MSGText = "※　異　常　終　了　※"
    End With
    If Err <> 0 Then
        Call CisFun.ErrorBox
    End If
    
    Call CisFun.LogPut("データ 削除処理【異常終了】")

End Function
'*---------------------*
'*  テキスト出力       *
'*---------------------*
Private Function SaveTextProc(TBLNm As String) As Boolean
    
    SaveTextProc = False

    If Trim(gSL_Select) = "" Then
       SaveTextProc = True
       Exit Function
    End If
    
    With CisDB
        .SQL = gSL_Select
        If Not .DBRead Then
           Call CisFun.LogPut("                    ★★★ ﾃｷｽﾄ出力件数 = 0 ★★★")
           GoTo SaveTextProc_ED
        End If
    End With
    
    With CISFrm
        Bar_Max(0) = CisDB.RecordCount
        Set .MeterPic = Bar(0)
        .MeterMax = Bar_Max(0)
        .MeterReset
    End With
    
    Call CisFun.LogPut("                    ★★★ ﾃｷｽﾄ出力件数 = " & CisDB.RecordCount & " ★★★")

    Static sw_FNo       As Byte
    Static sw_Txt       As String
    Static sw_Str       As String
    Static sw_Str1      As String

    Static sw_ColMax    As Integer
    Static sw_ColCnt    As Integer
    Static sw_RowMax    As Long
    Static sw_RowCnt    As Long
    Static sw_Per       As Byte

    sw_ColMax = CisDB.Recordset.Fields.Count - 1
    sw_RowMax = CisDB.RecordCount
    sw_FNo = FreeFile

    sw_Txt = mSaveDir & RTrim(TBLNm) & ".csv"

    Open sw_Txt For Output Access Write As #sw_FNo

'   * 見出し部 *
    sw_Str = ""
    With CisDB.Recordset.Fields
        For sw_ColCnt = 0 To sw_ColMax - 1
            sw_Str = sw_Str & .Item(sw_ColCnt).Name & ","
        Next sw_ColCnt
        sw_Str = sw_Str & .Item(sw_ColMax).Name
    End With

    Print #sw_FNo, sw_Str

'   * データ部 *
    With CisDB.Recordset
        sw_RowCnt = 0
        Do Until .EOF

            sw_Str = ""
            With .Fields
                For sw_ColCnt = 0 To sw_ColMax - 1
                    sw_Str1 = ""
                    If Not IsNull(.Item(sw_ColCnt)) Then
                        sw_Str1 = Replace(RTrim(.Item(sw_ColCnt)), ",", ".")
                    End If
                    sw_Str = sw_Str & sw_Str1 & ","
                Next sw_ColCnt
                If Not IsNull(.Item(sw_ColCnt)) Then
                    sw_Str = sw_Str & .Item(sw_ColMax)
                End If
            End With

            Print #sw_FNo, sw_Str

            sw_RowCnt = sw_RowCnt + 1
'            sw_Per = sw_RowCnt / sw_RowMax * 100
'            If sw_Per > 100 Then sw_Per = 100
            
            With CISFrm
                Set .MeterPic = Bar(0)
                .MeterMax = Bar_Max(0)
                .MeterUpdate sw_RowCnt
            End With
            
            .MoveNext
        Loop
    End With

    Close sw_FNo
'
SaveTextProc_ED:
    On Error Resume Next
    CisDB.Recordset.Close
    On Error GoTo 0
    
    With CISFrm
        Set .MeterPic = Bar(0)
        .MeterMax = Bar_Max(0)
        .MeterUpdate Bar_Max(0)
    End With
    
    SaveTextProc = True
End Function
'*---------------------------------------------------------*
'*          シ ン グ ル プ ロ セ ス チ ェ ッ ク             *
'*---------------------------------------------------------*
Private Function SingleProcess_Check()
    SingleProcess_Check = False
    
    Call ItemsClearSYMCreate
    
    With SYMCreate
        .区分名称 = "シングルプロセス"
        .区分桁数 = 10
        .値 = App.EXEName
        .制御区分 = "X"
        .値名称 = "データ削除"
        .作成者 = gTanto
    End With
    
    Call SYMInitCreate(False)
    
    Call ItemsClearSYM
    gSL_Select = "Select 英数字1 From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称   = 'シングルプロセス'"
    gSL_Select = gSL_Select & " And   区分タイプ = ''"
    gSL_Select = gSL_Select & " And   値        = '" & App.EXEName & "'"
    Call SYMRead(gSL_Select, 1)
    
    With CisFun
        If RTrim$(SYM.英数字1) <> "" And RTrim$(SYM.英数字1) <> .WSID Then
            .MB_Lines = 3
            .MB_MSG(1) = "    他のパソコンで起動中です         "
            .MB_MSG(2) = "        ( " & RTrim$(SYM.英数字1) & " )"
            .MB_Button = Error
            .MBOX
            Exit Function
        End If
    End With

'   < プロセスのロック >
    gSL_Select = "UpDate 名称マスタ Set"
    gSL_Select = gSL_Select & " 英数字1 = SubString( Host_Name(), 1, 20 ),"
    gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
    gSL_Select = gSL_Select & " 更新者 = '" & RTrim(gTanto) & "',"
    gSL_Select = gSL_Select & " 更新端末 = SubString( Host_Name(), 1, 20 )"
    gSL_Select = gSL_Select & " Where 区分名称 = 'シングルプロセス'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 = '" & App.EXEName & "'"

    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With

    SingleProcess_Check = True
End Function
'*-------------------------------------------------*
'*          シ ン グ ル プ ロ セ ス 解 除           *
'*-------------------------------------------------*
Private Function SingleProcess_Release()
    SingleProcess_Release = False
    
    gSL_Select = "UpDate 名称マスタ Set"
    gSL_Select = gSL_Select & " 英数字1 = '',"
    gSL_Select = gSL_Select & " 更新日 = GETDATE(),"
    gSL_Select = gSL_Select & " 更新者 = '" & RTrim(gTanto) & "',"
    gSL_Select = gSL_Select & " 更新端末 = SubString( Host_Name(), 1, 20 )"
    gSL_Select = gSL_Select & " Where 区分名称 = 'シングルプロセス'"
    gSL_Select = gSL_Select & "   And 区分タイプ = ''"
    gSL_Select = gSL_Select & "   And 値 = '" & App.EXEName & "'"

    With CisDB
        .SQL = gSL_Select
        .DBExec
    End With
    
    SingleProcess_Release = True
End Function
'+-------------------------------------------------------------------------+
'+      SQL文作成
'+   SqlType : 9/削除テーブル情報セット,0/Select文 作成, 1/Delete文 作成
'+   CrtType : 配列No
'+-------------------------------------------------------------------------+
Private Sub CrtSql(SqlType As Integer, CrtType As Integer, TBLNm As String)
    Dim wWhere      As String
    gSL_Select = ""
    
    If SqlType = "9" Then
    ' 削除対象テーブル名セット
        mMax_File = 39
        ReDim mTblName(mMax_File)
        ReDim mTblCnt(mMax_File)
        ReDim mItemRep(mMax_File)
    ' 部品
        gInt = 0: mTblName(gInt) = "発注テーブル"
        gInt = 1: mTblName(gInt) = "入荷テーブル"
        gInt = 2: mTblName(gInt) = "発注帳票発行テーブル"
        gInt = 3: mTblName(gInt) = "内示テーブル": mItemRep(gInt) = "内示年月"
        gInt = 4: mTblName(gInt) = "生産計画テーブル": mItemRep(gInt) = "内示年月"
        gInt = 5: mTblName(gInt) = "発注内示テーブル": mItemRep(gInt) = "内示年月"
        gInt = 6: mTblName(gInt) = "かんばん回転枚数テーブル": mItemRep(gInt) = "年月"
        gInt = 7: mTblName(gInt) = "かんばん増減テーブル": mItemRep(gInt) = "Substring(終了日,1,6)"
        gInt = 8: mTblName(gInt) = "かんばん平準化テーブル": mItemRep(gInt) = "Substring(終了日,1,6)"
        gInt = 9: mTblName(gInt) = "実績管理テーブル": mItemRep(gInt) = "Substring(実績日,1,6)"
        gInt = 10: mTblName(gInt) = "在庫マスタ": mItemRep(gInt) = "年月度"
        gInt = 11: mTblName(gInt) = "在庫履歴マスタ": mItemRep(gInt) = "年月度"
        gInt = 12: mTblName(gInt) = "支給実績テーブル": mItemRep(gInt) = "実績年月"
    ' 材料
        gInt = 13: mTblName(gInt) = "材料在庫マスタ": mItemRep(gInt) = "年月度"
        gInt = 14: mTblName(gInt) = "材料在庫履歴マスタ": mItemRep(gInt) = "年月度"
        gInt = 15: mTblName(gInt) = "材料発注テーブル"
        gInt = 16: mTblName(gInt) = "材料入荷テーブル"
        gInt = 17: mTblName(gInt) = "材料発注帳票発行テーブル"
        gInt = 18: mTblName(gInt) = "材料実績管理テーブル": mItemRep(gInt) = "Substring(実績日,1,6)"
    '+
        gInt = 19: mTblName(gInt) = "生産計画必要数テーブル": mItemRep(gInt) = "内示年月"
        gInt = 20: mTblName(gInt) = "発注内示必要数テーブル": mItemRep(gInt) = "内示年月"
        gInt = 21: mTblName(gInt) = "材料発注内示テーブル": mItemRep(gInt) = "内示年月"
        gInt = 22: mTblName(gInt) = "材料発注内示必要数テーブル": mItemRep(gInt) = "内示年月"
        gInt = 23: mTblName(gInt) = "内示TEXTテーブル": mItemRep(gInt) = "内示年月"
        
        gInt = 24: mTblName(gInt) = "受注テーブル"
        gInt = 25: mTblName(gInt) = "出荷テーブル"
        gInt = 26: mTblName(gInt) = "検収テーブル": mItemRep(gInt) = "Substring(検収日,1,6)"
        gInt = 27: mTblName(gInt) = "売上実績テーブル": mItemRep(gInt) = "実績年月"
        gInt = 28: mTblName(gInt) = "売上明細テーブル": mItemRep(gInt) = "実績年月"
        
        gInt = 29: mTblName(gInt) = "平準化チェックテーブル": mItemRep(gInt) = "年月"
        '+-------------------------------------------------- 2004/12/27 Insert
        gInt = 30: mTblName(gInt) = "生産指示テーブル"
        gInt = 31: mTblName(gInt) = "生産実績テーブル"
        '+-------------------------------------------------- 2004/12/27 Insert End
        '+-------------------------------------------------- 2005/01/11 Insert
        gInt = 32: mTblName(gInt) = "購入実績テーブル": mItemRep(gInt) = "実績年月"
        gInt = 33: mTblName(gInt) = "購入明細テーブル": mItemRep(gInt) = "実績年月"
        gInt = 34: mTblName(gInt) = "支給明細テーブル": mItemRep(gInt) = "実績年月"
        gInt = 35: mTblName(gInt) = "買掛支払テーブル": mItemRep(gInt) = "実績年月"
        gInt = 36: mTblName(gInt) = "売上差異明細テーブル": mItemRep(gInt) = "照合年月"
        gInt = 37: mTblName(gInt) = "売上予想テーブル": mItemRep(gInt) = "内示年月"
        gInt = 38: mTblName(gInt) = "売上予想型費テーブル": mItemRep(gInt) = "予想年月"
        '+-------------------------------------------------- 2005/01/11 Insert End
        
        Exit Sub
    End If
    ' 条件が年月のみの時
    If Trim(mItemRep(CrtType)) <> "" Then
        wWhere = vbCr & " Where ?REPSTRING? <= '" & iH1_YM & "'" & vbCr
        wWhere = wWhere & " And RTrim(?REPSTRING?) <> ''" & vbCr
        
        If SqlType = 0 Then
           gSL_Select = "Select * From " & TBLNm & wWhere
        Else
           gSL_Select = "DELETE " & TBLNm
           gSL_Select = gSL_Select & wWhere
        End If
        gSL_Select = Replace(gSL_Select, "?REPSTRING?", mItemRep(CrtType))
    
        GoTo CrtSql_ed
    End If
    
    Select Case mTblName(CrtType)
        Case "発注テーブル"
'            wWhere = vbCr & "    Left Outer Join 入荷テーブル NY" & vbCr
'            wWhere = wWhere & "    ON NY.発注管理NO = HC.発注管理NO" & vbCr
'            wWhere = wWhere & "    Where ( Isnull(NY.入荷日,'') = '' And Substring(HC.納入日,1,6) <= '" & iH1_YM & "')" & vbCr
'            wWhere = wWhere & "         Or ( Isnull(NY.入荷日,'') <> '' And Substring(NY.入荷日,1,6) <= '" & iH1_YM & "')" & vbCr
            wWhere = vbCr & "    Left Outer Join 入荷テーブル NY" & vbCr
            wWhere = wWhere & "    ON NY.発注管理NO = HC.発注管理NO" & vbCr
            wWhere = wWhere & "    Where Isnull(NY.入荷日,'') <> '' And Substring(NY.入荷日,1,6) <= '" & iH1_YM & "' " & vbCr
            If SqlType = 0 Then
               gSL_Select = "Select HC.* From " & TBLNm & " HC " & wWhere
            Else
               gSL_Select = "DELETE " & TBLNm
               gSL_Select = gSL_Select & " From " & TBLNm & " HC "
               gSL_Select = gSL_Select & wWhere
            End If
        Case "入荷テーブル"
            wWhere = vbCr & "    Left Outer Join 発注テーブル HC" & vbCr
            wWhere = wWhere & "    ON HC.発注管理NO = HN.発注管理NO" & vbCr
            wWhere = wWhere & "    Where Isnull(HC.発注管理NO,0) = 0 "
            If SqlType = 0 Then
               gSL_Select = "Select HN.* From " & TBLNm & " HN " & wWhere
            Else
               gSL_Select = "DELETE " & TBLNm
               gSL_Select = gSL_Select & " From " & TBLNm & " HN "
               gSL_Select = gSL_Select & wWhere
            End If
        Case "発注帳票発行テーブル"
            wWhere = vbCr & "    Left Outer Join 発注テーブル HC" & vbCr
            wWhere = wWhere & "    ON HC.納品書管理NO = HH.納品書管理NO" & vbCr
            wWhere = wWhere & "    Where Isnull(HC.納品書管理NO,0) = 0 "
            If SqlType = 0 Then
               gSL_Select = "Select HH.* From " & TBLNm & " HH " & wWhere
            Else
               gSL_Select = "DELETE " & TBLNm
               gSL_Select = gSL_Select & " From " & TBLNm & " HH "
               gSL_Select = gSL_Select & wWhere
            End If
        Case "材料発注テーブル"
            wWhere = vbCr & "    Left Outer Join 材料入荷テーブル NY" & vbCr
            wWhere = wWhere & "    ON NY.発注管理NO = HC.発注管理NO" & vbCr
            wWhere = wWhere & "    Where Isnull(NY.入荷日,'') <> '' And Substring(NY.入荷日,1,6) <= '" & iH1_YM & "' " & vbCr
            If SqlType = 0 Then
               gSL_Select = "Select HC.* From " & TBLNm & " HC " & wWhere
            Else
               gSL_Select = "DELETE " & TBLNm
               gSL_Select = gSL_Select & " From " & TBLNm & " HC "
               gSL_Select = gSL_Select & wWhere
            End If
        Case "材料入荷テーブル"
            wWhere = vbCr & "    Left Outer Join 材料発注テーブル HC" & vbCr
            wWhere = wWhere & "    ON HC.発注管理NO = HN.発注管理NO" & vbCr
            wWhere = wWhere & "    Where Isnull(HC.発注管理NO,0) = 0 "
            If SqlType = 0 Then
               gSL_Select = "Select HN.* From " & TBLNm & " HN " & wWhere
            Else
               gSL_Select = "DELETE " & TBLNm
               gSL_Select = gSL_Select & " From " & TBLNm & " HN "
               gSL_Select = gSL_Select & wWhere
            End If
        Case "材料発注帳票発行テーブル"
            wWhere = vbCr & "    Left Outer Join 材料発注テーブル HC" & vbCr
            wWhere = wWhere & "    ON HC.納品書管理NO = HH.納品書管理NO" & vbCr
            wWhere = wWhere & "    Where Isnull(HC.納品書管理NO,0) = 0 "
            If SqlType = 0 Then
               gSL_Select = "Select HH.* From " & TBLNm & " HH " & wWhere
            Else
               gSL_Select = "DELETE " & TBLNm
               gSL_Select = gSL_Select & " From " & TBLNm & " HH "
               gSL_Select = gSL_Select & wWhere
            End If
        Case "受注テーブル"
            wWhere = vbCr & "    Left Outer Join 出荷テーブル NY" & vbCr
            wWhere = wWhere & "    ON NY.出荷指示書NO = HC.出荷指示書NO" & vbCr
            wWhere = wWhere & "    And NY.出荷指示書行NO = HC.出荷指示書行NO" & vbCr
            wWhere = wWhere & "    Where Isnull(NY.出荷日,'') <> '' And Substring(NY.出荷日,1,6) <= '" & iH1_YM & "' " & vbCr
            If SqlType = 0 Then
               gSL_Select = "Select HC.* From " & TBLNm & " HC " & wWhere
            Else
               gSL_Select = "DELETE " & TBLNm
               gSL_Select = gSL_Select & " From " & TBLNm & " HC "
               gSL_Select = gSL_Select & wWhere
            End If
        Case "出荷テーブル"
            wWhere = vbCr & "    Left Outer Join 受注テーブル HC" & vbCr
            wWhere = wWhere & "    ON HC.出荷指示書NO = HN.出荷指示書NO" & vbCr
            wWhere = wWhere & "    And HC.出荷指示書行NO = HN.出荷指示書行NO" & vbCr
            wWhere = wWhere & "    Where Isnull(HC.出荷指示書NO,0) = 0 "
            If SqlType = 0 Then
               gSL_Select = "Select HN.* From " & TBLNm & " HN " & wWhere
            Else
               gSL_Select = "DELETE " & TBLNm
               gSL_Select = gSL_Select & " From " & TBLNm & " HN "
               gSL_Select = gSL_Select & wWhere
            End If
        '+-------------------------------------------------- 2004/12/27 Insert
        Case "生産指示テーブル"
            wWhere = vbCr & "    Left Outer Join ( " & vbCr
            wWhere = wWhere & " select 生産指示書NO,生産指示書行,MAX(生産日) 生産日 From 生産実績テーブル" & vbCr
            wWhere = wWhere & "             Group by 生産指示書NO,生産指示書行 ) SJ " & vbCr
            wWhere = wWhere & "    ON SJ.生産指示書NO = SS.生産指示書NO" & vbCr
            wWhere = wWhere & "    And SJ.生産指示書行 = SS.生産指示書行" & vbCr
            wWhere = wWhere & "    Where Isnull(SJ.生産日,'') <> '' And Substring(SJ.生産日,1,6) <= '" & iH1_YM & "' " & vbCr
            If SqlType = 0 Then
               gSL_Select = "Select SS.* From " & TBLNm & " SS " & wWhere
            Else
               gSL_Select = "DELETE " & TBLNm
               gSL_Select = gSL_Select & " From " & TBLNm & " SS "
               gSL_Select = gSL_Select & wWhere
            End If
        Case "生産実績テーブル"
            wWhere = vbCr & "    Left Outer Join 生産指示テーブル SS" & vbCr
            wWhere = wWhere & "    ON SS.生産指示書NO = SJ.生産指示書NO" & vbCr
            wWhere = wWhere & "    And SS.生産指示書行 = SJ.生産指示書行" & vbCr
            wWhere = wWhere & "    Where Isnull(SS.生産指示書NO,0) = 0 "
            If SqlType = 0 Then
               gSL_Select = "Select SJ.* From " & TBLNm & " SJ " & wWhere
            Else
               gSL_Select = "DELETE " & TBLNm
               gSL_Select = gSL_Select & " From " & TBLNm & " SJ "
               gSL_Select = gSL_Select & wWhere
            End If
        '+-------------------------------------------------- 2004/12/27 Insert
    End Select

CrtSql_ed:
    If w_Test Then
        Call CisFun.LogPut(Replace(gSL_Select, vbCr, ""))
    End If
End Sub
